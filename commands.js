/**
 * Kalexius – Wrike Task ID Validator
 * commands.js  v2.6.1
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
      var isWatched = WATCHED_EMAILS.some(function (address) {
        return fromEmail === address.toLowerCase().trim();
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
