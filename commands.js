/**
 * Kalexius – Wrike Task ID Validator
 * commands.js  v2.4.0
 *
 * Shared mailbox aware:
 *  - Checks the "From" field shown in the compose window
 *    (works correctly even when sending from a shared mailbox)
 *  - Only validates emails being sent FROM watched domains
 *  - All other senders pass through freely
 *
 * Flow (watched domains only):
 *  - Subject HAS [12345]   → allow immediately
 *  - 1st Send, no Task ID  → block with warning, set flag
 *  - 2nd Send, flag set    → allow through
 */

/* global Office */

// ─── CONFIGURE YOUR DOMAINS HERE ────────────────────────────────────────────
var WATCHED_DOMAINS = [
  "odyssey.limited",
  "kalexius.com",
  "flagshiplegal.com"
];
// ────────────────────────────────────────────────────────────────────────────

Office.initialize = function () {};
Office.onReady(function () {});

function validateSubject(event) {
  try {
    var item = Office.context.mailbox.item;

    // item.from reflects the actual From field selected in compose
    // including shared mailboxes — this is the correct field to use
    item.from.getAsync(function (fromResult) {

      var senderDomain = "";

      if (fromResult.status === Office.AsyncResultStatus.Succeeded) {
        var email = (fromResult.value.emailAddress || "").toLowerCase();
        senderDomain = email.split("@")[1] || "";
      }

      // If item.from didn't give us a watched domain, also check
      // the logged-in user's profile as a fallback (edge case)
      var profileEmail = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();
      var profileDomain = profileEmail.split("@")[1] || "";

      var shouldValidate =
        WATCHED_DOMAINS.some(function (d) { return senderDomain === d.toLowerCase(); }) ||
        WATCHED_DOMAINS.some(function (d) { return profileDomain === d.toLowerCase(); });

      if (!shouldValidate) {
        // ✅ Not a watched domain — skip validation, send freely
        event.completed({ allowEvent: true });
        return;
      }

      // ── Watched domain matched — run Task ID check ───────────────────────

      item.loadCustomPropertiesAsync(function (cpResult) {
        var props = cpResult.value;
        var alreadyWarned = props.get("wrikeWarned") === true;

        item.subject.getAsync(function (subResult) {
          if (subResult.status !== Office.AsyncResultStatus.Succeeded) {
            props.remove("wrikeWarned");
            props.saveAsync(function () {
              event.completed({ allowEvent: true });
            });
            return;
          }

          var subject = subResult.value || "";
          var hasTaskId = /\[\d+\]/.test(subject);

          if (hasTaskId) {
            // ✅ Task ID present — allow
            props.remove("wrikeWarned");
            props.saveAsync(function () {
              event.completed({ allowEvent: true });
            });

          } else if (alreadyWarned) {
            // ✅ 2nd Send — user confirmed, allow through
            props.remove("wrikeWarned");
            props.saveAsync(function () {
              event.completed({ allowEvent: true });
            });

          } else {
            // ⚠️ 1st Send, no Task ID — warn and block
            props.set("wrikeWarned", true);
            props.saveAsync(function () {
              event.completed({
                allowEvent: false,
                errorMessage:
                  "No Wrike Task ID found in the subject. " +
                  "Add one using the format [12345] at the end of your subject. " +
                  "Click Send again to send without a Task ID."
              });
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
