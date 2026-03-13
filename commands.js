/**
 * Kalexius – Wrike Task ID Validator
 * commands.js  v2.0.0
 *
 * Logic:
 *  - 1st Send with no Task ID  → block + show warning banner
 *  - 2nd Send with no Task ID  → allow through (user confirmed intentional)
 *  - Any Send WITH Task ID     → always allow immediately
 *
 * Task ID format: [12345]  (digits inside square brackets anywhere in subject)
 *
 * The "warned" flag lives in memory for the current compose window session.
 * It resets automatically when the compose window is closed/sent.
 */

/* global Office */

// In-memory flag — one instance per compose window (each compose window
// gets its own JS runtime in Outlook's add-in sandbox).
var _userHasBeenWarned = false;

// ─── Office initialisation ────────────────────────────────────────────────────

Office.initialize = function () {
  // Nothing needed here for an event-only add-in.
};

Office.onReady(function () {
  // Expose the handler globally so Outlook can call it by name
  // (required when using the manifest FunctionName attribute).
});

// ─── Main event handler ───────────────────────────────────────────────────────

/**
 * validateSubject
 *
 * Called synchronously by Outlook on every Send attempt.
 * Must call event.completed() to unblock Outlook.
 *
 * @param {Office.AddinCommands.Event} event
 */
function validateSubject(event) {
  try {
    Office.context.mailbox.item.subject.getAsync(function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        // Can't read subject — fail open so we never permanently block sending.
        console.warn("[WrikeValidator] Could not read subject:", result.error);
        _resetWarning();
        event.completed({ allowEvent: true });
        return;
      }

      var subject = result.value || "";
      var hasTaskId = /\[\d+\]/.test(subject);

      if (hasTaskId) {
        // ✅ Task ID present — always allow, reset state.
        _resetWarning();
        event.completed({ allowEvent: true });

      } else if (_userHasBeenWarned) {
        // ✅ User was already warned — this is their confirmed second attempt.
        // Let it through and reset for next compose.
        _resetWarning();
        event.completed({ allowEvent: true });

      } else {
        // ⚠️ No Task ID and first attempt — warn and block.
        _userHasBeenWarned = true;
        event.completed({
          allowEvent: false,
          errorMessage:
            "No Wrike Task ID detected in the subject. " +
            "Add a Task ID using the format [12345] at the end of the subject, " +
            "or click Send again to send without one."
        });
      }
    });

  } catch (err) {
    // Safety net — never leave Outlook hanging.
    console.error("[WrikeValidator] Unexpected error:", err);
    _resetWarning();
    event.completed({ allowEvent: true });
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function _resetWarning() {
  _userHasBeenWarned = false;
}

// ─── Required: register handler with Office runtime ──────────────────────────
// This line is mandatory for SmartAlerts / ItemSend event handlers.

if (Office.actions) {
  Office.actions.associate("validateSubject", validateSubject);
}
