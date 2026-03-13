/**
 * Kalexius – Wrike Task ID Validator
 * commands.js  v2.6.3
 *
 * On block: opens the taskpane automatically so user sees
 * exactly what to fix — works on Mailbox 1.10.
 */

/* global Office */

var WATCHED_EMAILS = [
  "panama@odyssey.limited",
  "gravity@odyssey.limited"
];

Office.initialize = function () {};
Office.onReady(function () {});

function validateSubject(event) {
  try {
    var item = Office.context.mailbox.item;

    item.from.getAsync(function (fromResult) {
      if (fromResult.status !== Office.AsyncResultStatus.Succeeded) {
        event.completed({ allowEvent: true });
        return;
      }

      var fromEmail = (fromResult.value.emailAddress || "").toLowerCase().trim();
      var isWatched = WATCHED_EMAILS.some(function (a) {
        return fromEmail === a.toLowerCase().trim();
      });

      if (!isWatched) {
        event.completed({ allowEvent: true });
        return;
      }

      item.subject.getAsync(function (subResult) {
        if (subResult.status !== Office.AsyncResultStatus.Succeeded) {
          event.completed({ allowEvent: true });
          return;
        }

        var subject   = subResult.value || "";
        var hasTaskId = /\[\d+\]/.test(subject);

        if (hasTaskId) {
          event.completed({ allowEvent: true });
        } else {
          // Open the taskpane so user sees the clear error message
          Office.context.ui.displayDialogAsync(
            "https://klxit.github.io/KLX-Outlook-Plug-in/taskpane.html",
            { height: 60, width: 30, displayInIframe: true },
            function () {}
          );
          event.completed({ allowEvent: false });
        }
      });
    });

  } catch (err) {
    console.error("[WrikeValidator] Error:", err);
    event.completed({ allowEvent: true });
  }
}

if (Office.actions) {
  Office.actions.associate("validateSubject", validateSubject);
}
