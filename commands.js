/**
 * Kalexius – Wrike Task ID Validator
 * commands.js  v2.1.0
 *
 * Compatible with Mailbox 1.10 (no SendModeOverride needed).
 *
 * Flow:
 *  - Subject HAS [12345]       → allow immediately
 *  - Subject missing Task ID, 1st Send → block with warning message
 *  - Subject missing Task ID, 2nd Send → allow through (user confirmed)
 *
 * This gives users the "ignore and send anyway" behaviour
 * without requiring Mailbox 1.13.
 */

/* global Office */

// Per-compose-window flag. Resets when compose window closes.
var _userHasBeenWarned = false;

Office.initialize = function () {};
Office.onReady(function () {});

function validateSubject(event) {
  try {
    Office.context.mailbox.item.subject.getAsync(function (result) {

      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        // Can't read subject — fail open
        _userHasBeenWarned = false;
        event.completed({ allowEvent: true });
        return;
      }

      var subject = result.value || "";
      var hasTaskId = /\[\d+\]/.test(subject);

      if (hasTaskId) {
        // ✅ Task ID found — always allow
        _userHasBeenWarned = false;
        event.completed({ allowEvent: true });

      } else if (_userHasBeenWarned) {
        // ✅ Already warned — user is clicking Send a 2nd time, let it through
        _userHasBeenWarned = false;
        event.completed({ allowEvent: true });

      } else {
        // ⚠️ First send attempt, no Task ID — warn and block
        _userHasBeenWarned = true;
        event.completed({
          allowEvent: false,
          errorMessage:
            "No Wrike Task ID found in the subject. " +
            "Add one using the format [12345] at the end of your subject. " +
            "Click Send again to send without a Task ID."
        });
      }
    });

  } catch (err) {
    console.error("[WrikeValidator] Error:", err);
    _userHasBeenWarned = false;
    event.completed({ allowEvent: true });
  }
}

if (Office.actions) {
  Office.actions.associate("validateSubject", validateSubject);
}
