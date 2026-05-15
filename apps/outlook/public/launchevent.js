/* eslint-disable */
// Outlook on-send handler for the Halo plugin.
//
// Registered as `onMessageSendHandler` in the manifest's runtimes.actions and
// hooked to autoRunEvents[].messageSending. Lives in its own short-lived
// runtime served from the same origin (tools.iusehalo.com) as the task pane.
//
// Contract:
// - Outlook calls onMessageSendHandler(event) the moment the user clicks Send.
// - We have ~5 minutes to call event.completed() or the send is cancelled.
// - We always allow the send to proceed; Halo append failure is non-fatal.
//
// Diagnostics: every stage appends to a localStorage log ("halo.diagLog.v1")
// which the task pane's Settings → Diagnostics panel reads and polls. We use
// localStorage (not roamingSettings) because roamingSettings is loaded once
// per runtime and never refreshed — so the task pane couldn't see entries
// the launch-event runtime wrote after the task pane opened. localStorage
// is synchronous and shared across all same-origin runtimes.
// Stays in lock-step with apps/outlook/src/lib/diagnostics.ts (same key,
// same entry shape).

(function () {
  var TOKENS_KEY = "halo.tokens.v1";
  var CONFIG_KEY = "halo.tenantConfig.v1";
  var TICKET_PROP = "haloLogTicketId";
  var DIAG_LOG_KEY = "halo.diagLog.v1";
  var MAX_ENTRIES = 200;
  var MAX_MESSAGE_LEN = 500;
  var MAX_BYTES = 64000;

  // Outer safety: must fire BEFORE Outlook's own "taking longer than expected"
  // prompt — which the office-js issue tracker confirms appears at 5 seconds
  // on Outlook on the web (issue #3180). Anything ≥5s means the user sees
  // Outlook's generic dialog instead of our diagnostic errorMessage.
  var SAFETY_MS = 4000;
  // Per-fetch budget — token refresh and Action POST. AbortController turns a
  // hung CORS preflight into a real catchable error. Kept under SAFETY_MS so
  // a network failure surfaces through our .catch path, not the outer safety.
  var FETCH_TIMEOUT_MS = 3000;

  function getRoaming() {
    try { return Office.context.roamingSettings; } catch (e) { return null; }
  }

  function getTokens() {
    var rs = getRoaming();
    return rs ? rs.get(TOKENS_KEY) : undefined;
  }

  function getConfig() {
    var rs = getRoaming();
    return rs ? rs.get(CONFIG_KEY) : undefined;
  }

  // Append a diagnostic entry to the shared cross-runtime log. Synchronous —
  // no fire-and-forget races, so even an entry written immediately before
  // event.completed() is durable.
  function logEvent(level, message, data) {
    try {
      var raw = window.localStorage.getItem(DIAG_LOG_KEY);
      var arr = raw ? JSON.parse(raw) : [];
      if (!Array.isArray(arr)) arr = [];
      var entry = {
        ts: new Date().toISOString(),
        level: level,
        source: "on-send",
        message: message && message.length > MAX_MESSAGE_LEN
          ? message.slice(0, MAX_MESSAGE_LEN) + "…"
          : message,
      };
      if (data) entry.data = data;
      arr.push(entry);
      if (arr.length > MAX_ENTRIES) arr = arr.slice(-MAX_ENTRIES);
      while (arr.length > 1 && JSON.stringify(arr).length > MAX_BYTES) {
        arr = arr.slice(1);
      }
      window.localStorage.setItem(DIAG_LOG_KEY, JSON.stringify(arr));
    } catch (e) {
      // Logging must never throw into the handler path.
    }
    // Mirror to console for the dev-tools window if it happens to be open.
    try { console.log("[halo-on-send]", level, message, data || ""); } catch (e) {}
  }

  // Module-load breadcrumb. If the diagnostics panel never shows this entry
  // after a send attempt, the launch-event runtime itself didn't load — that
  // narrows the problem to manifest registration / Outlook activation rules
  // rather than anything inside the handler.
  logEvent("info", "launchevent.js module loaded");

  // Wrap fetch with an AbortController-driven timeout. Rejects with a tagged
  // error if the request runs longer than ms — covers hung CORS preflights.
  function fetchWithTimeout(url, init, ms) {
    var ctrl = (typeof AbortController !== "undefined") ? new AbortController() : null;
    var initWithSignal = Object.assign({}, init || {});
    if (ctrl) initWithSignal.signal = ctrl.signal;
    var timer = setTimeout(function () { if (ctrl) ctrl.abort(); }, ms);
    return fetch(url, initWithSignal).then(
      function (r) { clearTimeout(timer); return r; },
      function (e) {
        clearTimeout(timer);
        var msg = (e && e.name === "AbortError")
          ? ("Halo request timed out after " + ms + "ms (network/CORS)")
          : ("fetch failed: " + (e && e.message ? e.message : String(e)));
        throw new Error(msg);
      },
    );
  }

  function getAccessToken() {
    var tokens = getTokens();
    var cfg = getConfig();
    if (!tokens || !cfg) return Promise.reject(new Error("Not authenticated (no tokens/config in roamingSettings)"));
    if (Date.now() < tokens.expiresAt - 60000) return Promise.resolve(tokens.accessToken);

    logEvent("info", "refreshing access token");
    var body = new URLSearchParams({
      grant_type: "refresh_token",
      refresh_token: tokens.refreshToken,
      client_id: cfg.clientId,
    });
    return fetchWithTimeout(cfg.haloBaseUrl + "/auth/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body,
    }, FETCH_TIMEOUT_MS)
      .then(function (r) {
        if (!r.ok) throw new Error("Token refresh failed " + r.status);
        return r.json();
      })
      .then(function (json) {
        var fresh = {
          accessToken: json.access_token,
          refreshToken: json.refresh_token || tokens.refreshToken,
          expiresAt: Date.now() + json.expires_in * 1000,
          scope: json.scope,
        };
        var rs = getRoaming();
        if (rs) {
          rs.set(TOKENS_KEY, fresh);
          rs.saveAsync(function () {});
        }
        return fresh.accessToken;
      });
  }

  function loadCustomProps() {
    return new Promise(function (resolve, reject) {
      Office.context.mailbox.item.loadCustomPropertiesAsync(function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("custom props load failed: " + (r.error && r.error.message)));
          return;
        }
        resolve(r.value);
      });
    });
  }

  function readBody() {
    return new Promise(function (resolve, reject) {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("body read failed: " + (r.error && r.error.message)));
          return;
        }
        resolve(r.value);
      });
    });
  }

  function readSubject() {
    return new Promise(function (resolve) {
      Office.context.mailbox.item.subject.getAsync(function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) { resolve(""); return; }
        resolve(r.value || "");
      });
    });
  }

  function readRecipientsField(field) {
    return new Promise(function (resolve) {
      try {
        field.getAsync(function (r) {
          if (r.status !== Office.AsyncResultStatus.Succeeded) { resolve([]); return; }
          resolve((r.value || []).map(function (e) { return e.emailAddress; }).filter(Boolean));
        });
      } catch (e) {
        resolve([]);
      }
    });
  }

  function appendToHalo(ticketId, data) {
    return getAccessToken().then(function (token) {
      var cfg = getConfig();
      var senderEmail = "";
      var senderName = "";
      try {
        senderEmail = Office.context.mailbox.userProfile.emailAddress || "";
        senderName = Office.context.mailbox.userProfile.displayName || "";
      } catch (e) {
        // fall through with blanks
      }
      var payload = [{
        ticket_id: Number(ticketId),
        outcome: "Outgoing Email",
        note: data.body,
        emailfrom: senderEmail,
        emailfromname: senderName,
        emailto: (data.to || []).join("; "),
        emailcc: (data.cc || []).join("; "),
        emailsubject: data.subject,
        agent_id: 0,
      }];
      logEvent("info", "POST /api/Actions", {
        ticketId: ticketId,
        bodyLen: (data.body || "").length,
        toCount: (data.to || []).length,
      });
      return fetchWithTimeout(cfg.haloBaseUrl + "/api/Actions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": "Bearer " + token,
        },
        body: JSON.stringify(payload),
      }, FETCH_TIMEOUT_MS).then(function (r) {
        if (!r.ok) {
          return r.text().then(function (t) { throw new Error("Halo " + r.status + ": " + t.slice(0, 200)); });
        }
        return r.json();
      });
    });
  }

  function onMessageSendHandler(event) {
    var startedAt = Date.now();
    var completed = false;
    var stage = "init";

    logEvent("info", "handler invoked");

    var finish = function (errorMessage) {
      if (completed) return;
      completed = true;
      var opts = { allowEvent: true };
      if (errorMessage) opts.errorMessage = errorMessage;
      logEvent(errorMessage ? "error" : "info",
        errorMessage ? ("finished with error in '" + stage + "': " + errorMessage)
                     : ("finished ok in '" + stage + "'"),
        { stage: stage, durationMs: Date.now() - startedAt });
      // logEvent writes synchronously to localStorage, so the entry above is
      // durable before event.completed() terminates the runtime.
      try { event.completed(opts); } catch (e) {}
    };

    // Outer safety. If we get here, something below didn't propagate.
    var safety = setTimeout(function () {
      finish("HaloPSA log-on-send hung in stage \"" + stage + "\" — email sent anyway.");
    }, SAFETY_MS);

    stage = "loadCustomProps";
    logEvent("info", "stage → loadCustomProps");
    loadCustomProps().then(function (cp) {
      var ticketId = cp.get(TICKET_PROP);
      logEvent("info", "ticketId resolved", { ticketId: ticketId || null });
      if (!ticketId) {
        clearTimeout(safety);
        finish();
        return;
      }
      stage = "readItem";
      logEvent("info", "stage → readItem");
      return Promise.all([
        readBody(),
        readSubject(),
        readRecipientsField(Office.context.mailbox.item.to),
        readRecipientsField(Office.context.mailbox.item.cc),
      ]).then(function (parts) {
        stage = "appendToHalo";
        logEvent("info", "stage → appendToHalo");
        return appendToHalo(ticketId, {
          body: parts[0],
          subject: parts[1],
          to: parts[2],
          cc: parts[3],
        });
      }).then(function () {
        stage = "clearMarker";
        logEvent("info", "stage → clearMarker");
        // Clear the marker so re-sending the same draft doesn't double-log.
        cp.set(TICKET_PROP, null);
        cp.saveAsync(function () {
          clearTimeout(safety);
          finish();
        });
      });
    }).catch(function (err) {
      clearTimeout(safety);
      var msg = err && err.message ? err.message : String(err);
      finish("HaloPSA log-on-send failed in \"" + stage + "\": " + msg + ". Email sent anyway.");
    });
  }

  // Registration must happen at module load — Outlook dispatches by name.
  Office.onReady(function () {
    if (!Office.actions || typeof Office.actions.associate !== "function") {
      logEvent("error", "Office.actions.associate unavailable — handler cannot register");
      return;
    }
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    logEvent("info", "handler registered");
  });
})();
