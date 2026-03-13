/**
 * Kalexius – Wrike Task ID Validator
 * commands.js  v3.1.0
 *
 * Shared mailbox aware:
 * - Uses item.from first
 * - Falls back to getSharedPropertiesAsync to get the real shared mailbox address
 * - Checks both against WATCHED_EMAILS
 */

/* global Office */

var WATCHED_EMAILS = [
  "panama@odyssey.limited",
  "gravity@odyssey.limited",
  "elektra@odyssey.limited",
  "eversana@odyssey.limited",
  "filings@odyssey.limited",
  "genesyscosec@odyssey.limited",
  "herschelcosec@odyssey.limited",
  "ifit@odyssey.limited",
  "quartz@odyssey.limited",
  "nexus@odyssey.limited",
  "expertise.innovation@kalexius.com",
  "bearcomcosec@odyssey.limited"
];

Office.initialize = function () {};
Office.onReady(function () {});

function isWatchedEmail(email) {
  return WATCHED_EMAILS.some(function (a) {
    return (email || "").toLowerCase().trim() === a.toLowerCase().trim();
  });
}

function checkSubjectAndComplete(event) {
  var item = Office.context.mailbox.item;
  item.subject.getAsync(function (subResult) {
    if (subResult.status !== Office.AsyncResultStatus.Succeeded) {
      event.completed({ allowEvent: true });
      return;
    }
    var hasTaskId = /\[\d+\]/.test(subResult.value || "");
    event.completed({ allowEvent: hasTaskId });
  });
}

function validateSubject(event) {
  try {
    var item = Office.context.mailbox.item;

    item.from.getAsync(function (fromResult) {
      var fromEmail = "";
      if (fromResult.status === Office.AsyncResultStatus.Succeeded) {
        fromEmail = (fromResult.value.emailAddress || "").toLowerCase().trim();
      }

      if (isWatchedEmail(fromEmail)) {
        // from.getAsync returned a watched address — check subject
        checkSubjectAndComplete(event);
        return;
      }

      // from.getAsync may have returned the delegate's address instead
      // of the shared mailbox — try getSharedPropertiesAsync as fallback
      if (item.getSharedPropertiesAsync) {
        item.getSharedPropertiesAsync(function (sharedResult) {
          if (sharedResult.status === Office.AsyncResultStatus.Succeeded) {
            var sharedEmail = (sharedResult.value.targetMailbox || "").toLowerCase().trim();
            if (isWatchedEmail(sharedEmail)) {
              checkSubjectAndComplete(event);
              return;
            }
          }
          // Neither from nor shared mailbox is watched — send freely
          event.completed({ allowEvent: true });
        });
      } else {
        // getSharedPropertiesAsync not available — send freely
        event.completed({ allowEvent: true });
      }
    });

  } catch (err) {
    console.error("[WrikeValidator] Error:", err);
    event.completed({ allowEvent: true });
  }
}

if (Office.actions) {
  Office.actions.associate("validateSubject", validateSubject);
}
