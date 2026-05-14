/* eslint-disable */
// Outlook on-send handler for the Halo plugin.
//
// Registered as `onMessageSendHandler` in the manifest's runtimes.actions and
// hooked to autoRunEvents[].messageSending. Lives in its own short-lived
// runtime, so it cannot share JS state with the compose task pane — only the
// per-item CustomProperties bag and the per-mailbox roamingSettings bag are
// shared.
//
// Contract:
// - Outlook calls onMessageSendHandler(event) the moment the user clicks Send.
// - We have ~5 minutes to call event.completed() or the send is cancelled.
// - We always allow the send to proceed; Halo append failure is non-fatal.

(function () {
  var TOKENS_KEY = "halo.tokens.v1";
  var CONFIG_KEY = "halo.tenantConfig.v1";
  var TICKET_PROP = "haloLogTicketId";

  function getRoaming() {
    try {
      return Office.context.roamingSettings;
    } catch (e) {
      return null;
    }
  }

  function getTokens() {
    var rs = getRoaming();
    return rs ? rs.get(TOKENS_KEY) : undefined;
  }

  function getConfig() {
    var rs = getRoaming();
    return rs ? rs.get(CONFIG_KEY) : undefined;
  }

  // Refresh the access token if it's within 60s of expiry. Returns a Promise
  // resolving to a usable access token string.
  function getAccessToken() {
    var tokens = getTokens();
    var cfg = getConfig();
    if (!tokens || !cfg) return Promise.reject(new Error("Not authenticated"));
    if (Date.now() < tokens.expiresAt - 60000) return Promise.resolve(tokens.accessToken);

    var body = new URLSearchParams({
      grant_type: "refresh_token",
      refresh_token: tokens.refreshToken,
      client_id: cfg.clientId,
    });
    return fetch(cfg.haloBaseUrl + "/auth/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body,
    })
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
          reject(new Error("custom props load failed"));
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
          reject(new Error("body read failed"));
          return;
        }
        resolve(r.value);
      });
    });
  }

  function readSubject() {
    return new Promise(function (resolve, reject) {
      Office.context.mailbox.item.subject.getAsync(function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          resolve("");
          return;
        }
        resolve(r.value || "");
      });
    });
  }

  function readRecipientsField(field) {
    return new Promise(function (resolve) {
      try {
        field.getAsync(function (r) {
          if (r.status !== Office.AsyncResultStatus.Succeeded) {
            resolve([]);
            return;
          }
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
      return fetch(cfg.haloBaseUrl + "/api/Actions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": "Bearer " + token,
        },
        body: JSON.stringify(payload),
      }).then(function (r) {
        if (!r.ok) {
          return r.text().then(function (t) { throw new Error("Halo " + r.status + ": " + t.slice(0, 200)); });
        }
        return r.json();
      });
    });
  }

  function onMessageSendHandler(event) {
    var completed = false;
    var finish = function (errorMessage) {
      if (completed) return;
      completed = true;
      var opts = { allowEvent: true };
      if (errorMessage) opts.errorMessage = errorMessage;
      event.completed(opts);
    };

    // Safety net: if anything hangs, allow the send after 25s rather than
    // making the user wait for Microsoft's 5-minute timeout.
    var safety = setTimeout(function () { finish(); }, 25000);

    loadCustomProps().then(function (cp) {
      var ticketId = cp.get(TICKET_PROP);
      if (!ticketId) {
        clearTimeout(safety);
        finish();
        return;
      }
      return Promise.all([
        readBody(),
        readSubject(),
        readRecipientsField(Office.context.mailbox.item.to),
        readRecipientsField(Office.context.mailbox.item.cc),
      ]).then(function (parts) {
        return appendToHalo(ticketId, {
          body: parts[0],
          subject: parts[1],
          to: parts[2],
          cc: parts[3],
        });
      }).then(function () {
        // Clear the marker so re-sending the same draft doesn't double-log.
        cp.set(TICKET_PROP, null);
        cp.saveAsync(function () {
          clearTimeout(safety);
          finish();
        });
      });
    }).catch(function (err) {
      clearTimeout(safety);
      finish("HaloPSA log-on-send failed: " + (err && err.message ? err.message : err) + ". Email sent anyway.");
    });
  }

  // Registration must happen at module load — Outlook dispatches by name.
  Office.onReady(function () {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  });
})();
