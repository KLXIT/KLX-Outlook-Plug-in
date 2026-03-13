/**
 * Kalexius – Wrike Task ID Validator
 * commands.js  v2.6.0
 *
 * Triggers only when sending FROM specific email addresses.
 * If the From field matches a watched address → enforce Wrike Task ID.
 * All other senders → pass through freely.
 *
 * User flow (watched addresses only):
 *  - Subject has [12345]  → sends immediately ✅
 *  - Subject missing ID   → blocked with message, user adds ID and sends ✅
 */

/* global Office */

// ─── ADD YOUR WATCHED EMAIL ADDRESSES HERE ───────────────────────────────────
var WATCHED_EMAILS = [
  "panama@odyssey.limited",
  "gravity@odyssey.limited"
  // add more addresses as needed, one per line
];
// ────────────────────────────────────────────────────────────────────────────

Office.initialize = function () {};
Office.onReady(function () {});

function validateSubject(event) {
  try {
    var item = Office.context.mailbox.item;

    item.from.getAsync(function (fromResult) {

      // Can't read From — fail open
      if (fromResult.status !== Office.AsyncResultStatus.Succeeded) {
        event.completed({ allowEvent: true });
        return;
      }

      var fromEmail = (fromResult.value.emailAddress || "").toLowerCase().trim();

      var isWatched = WATCHED_EMAILS.some(function (address) {
        return fromEmail === address.toLowerCase().trim();
      });

      if (!isWatched) {
        // Not a watched address — send freely
        event.completed({ allowEvent: true });
        return;
      }

      // ── Watched address — check for Wrike Task ID in subject ─────────────

      item.subject.getAsync(function (subResult) {

        if (subResult.status !== Office.AsyncResultStatus.Succeeded) {
          event.completed({ allowEvent: true });
          return;
        }

        var subject   = subResult.value || "";
        var hasTaskId = /\[\d+\]/.test(subject);

        if (hasTaskId) {
          // ✅ Task ID found — allow
          event.completed({ allowEvent: true });

        } else {
          // ❌ No Task ID — block with clear instructions
          event.completed({
            allowEvent: false,
            errorMessage:
              "A Wrike Task ID is required when sending from " + fromEmail + ". " +
              "Please add it at the end of the subject using the format [12345], then click Send."
          });
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
