// Office Dialog implementation of the @iusehalo/halo-api DialogOpener interface.
//
// Outlook's task pane opens the OAuth authorize URL via Office.context.ui.displayDialogAsync.
// The callback page (auth/callback.html) parses ?code=... / ?error=... and relays the payload
// back to us via Office.context.ui.messageParent, which we surface as a DialogResultMessage.
//
// displayDialogAsync requires the initial URL to be at the add-in's own origin — not a
// subdomain, not a different domain even if listed in validDomains/AppDomains. The
// wrapper page at /outlook/auth/start.html bounces to the third-party authorize URL.

import type { DialogOpener, DialogResultMessage } from "@iusehalo/halo-api";

const REDIRECT_PATH = "/outlook/auth/callback.html";
const START_PATH = "/outlook/auth/start.html";
const ADDIN_ORIGIN = "https://tools.iusehalo.com";

export function redirectUri(): string {
  return ADDIN_ORIGIN + REDIRECT_PATH;
}

/** Wrap the third-party authorize URL in our same-origin start page. */
export function wrapAuthorizeUrl(authorizeUrl: string): string {
  return `${ADDIN_ORIGIN}${START_PATH}?to=${encodeURIComponent(authorizeUrl)}`;
}

/** Singleton Office-dialog opener — pass into signIn(). */
export const officeDialogOpener: DialogOpener = {
  open(url: string): Promise<DialogResultMessage> {
    return new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(
        url,
        { height: 60, width: 30, promptBeforeOpen: false },
        (asyncResult) => {
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error(asyncResult.error?.message ?? "Failed to open auth dialog"));
            return;
          }
          const dialog = asyncResult.value;

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            dialog.close();
            try {
              // messageProperty is the JSON string posted from the callback page
              const data: DialogResultMessage = JSON.parse(
                (arg as { message: string }).message,
              );
              resolve(data);
            } catch (e) {
              reject(new Error(`Bad message from auth dialog: ${(e as Error).message}`));
            }
          });

          dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
            // 12006 = user closed the dialog
            const ev = arg as { error: number };
            dialog.close();
            if (ev.error === 12006) {
              reject(new Error("Sign-in cancelled."));
            } else {
              reject(new Error(`Auth dialog error: ${ev.error}`));
            }
          });
        },
      );
    });
  },
};
