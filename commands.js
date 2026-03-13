/**
 * Kalexius – Wrike Task ID Validator
 * commands.js  v2.7.0
 *
 * Uses notificationMessages API (Mailbox 1.10 compatible) to show
 * a visible warning bar at the top of the compose window.
 *
 * Flow:
 *  - From address not in watched list     → send freely
 *  - Subject has [1234]                   → send immediately ✅
 *  - Subject missing Task ID, 1st Send   → show warning bar, block
 *  - Subject missing Task ID, 2nd Send   → allow through ✅
 */

/* global Office */

// ─── WATCHED EMAIL ADDRESSES ─────────────────────────────────────────────────
var WATCHED_EMAILS = [
  "panama@odyssey.limited",
  "gravity@odyssey.limited"
];
// ────────────────────────────────────────────────────────────────────────────

var NOTIF_ID = "wrikeWarning";

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

      // ── Watched address — check subject ──────────────────────────────────

      item.loadCustomPropertiesAsync(function (cpResult) {
        var props        = cpResult.value;
        var alreadyWarned = props.get("wrikeWarned") === true;

        item.subject.getAsync(function (subResult) {

          if (subResult.status !== Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true });
            return;
          }

          var subject   = subResult.value || "";
          var hasTaskId = /\[\d+\]/.test(subject);

          if (hasTaskId) {
            // ✅ Task ID present — clear notification and allow
            item.notificationMessages.removeAsync(NOTIF_ID);
            props.remove("wrikeWarned");
            props.saveAsync(function () {
              event.completed({ allowEvent: true });
            });

          } else if (alreadyWarned) {
            // ✅ 2nd Send — allow through and clear
            item.notificationMessages.removeAsync(NOTIF_ID);
            props.remove("wrikeWarned");
            props.saveAsync(function () {
              event.completed({ allowEvent: true });
            });

          } else {
            // ⚠️ 1st Send — show warning bar at top and block
            props.set("wrikeWarned", true);
            props.saveAsync(function () {

              // Show persistent warning bar at top of compose window
              item.notificationMessages.replaceWithItemNotificationAsync(
                NOTIF_ID,
                {
                  type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
                  message: "⚠️ Please add the Wrike Task ID to the subject using the format [1234] — e.g. 'Meeting notes [9876]'. Click Send again to send without one."
                },
                function () {
                  event.completed({ allowEvent: false });
                }
              );
            });
          }
        });
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
