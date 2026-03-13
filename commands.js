/**
 * Kalexius – Wrike Task ID Validator
 * commands.js  v2.1.0
 *
 * Logic:
 *  - Send WITH Task ID [12345]  → always allow immediately
 *  - Send WITHOUT Task ID       → soft warning with "Send Anyway" button
 *
 * Uses SmartAlerts "sendModeOverride: PromptUser" so Outlook shows
 * a native "Send Anyway" option — user only ever needs max 2 clicks.
 */

/* global Office */

Office.initialize = function () {};

Office.onReady(function () {});

/**
 * validateSubject
 * Called synchronously by Outlook on every Send attempt.
 *
 * @param {Office.AddinCommands.Event} event
 */
function validateSubject(event) {
  try {
    Office.context.mailbox.item.subject.getAsync(function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        // Can't read subject — fail open, never permanently block sending
        console.warn("[WrikeValidator] Could not read subject:", result.error);
        event.completed({ allowEvent: true });
        return;
      }

      var subject = result.value || "";
      var hasTaskId = /\[\d+\]/.test(subject);

      if (hasTaskId) {
        // ✅ Task ID present — allow immediately
        event.completed({ allowEvent: true });
      } else {
        // ⚠️ No Task ID — soft warning with "Send Anyway" option
        // sendModeOverride: "PromptUser" tells Outlook to show
        // the warning banner WITH a "Send Anyway" button.
        // Clicking "Send Anyway" bypasses the add-in and sends.
        event.completed({
          allowEvent: false,
          errorMessage:
            "No Wrike Task ID detected in the subject. " +
            "Add a Task ID using the format [12345] at the end of the subject, " +
            "or click 'Send Anyway' to send without one.",
          sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser
        });
      }
    });

  } catch (err) {
    console.error("[WrikeValidator] Unexpected error:", err);
    event.completed({ allowEvent: true });
  }
}

// Register handler with Office runtime
if (Office.actions) {
  Office.actions.associate("validateSubject", validateSubject);
}
