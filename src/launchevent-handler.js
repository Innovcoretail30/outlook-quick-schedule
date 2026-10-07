/* global Office, QuickScheduleOffice */

function onMessageSendHandler(event) {
  try {
    QuickScheduleOffice.applyDelayOnSend(event);
  } catch (e) {
    if (event && !event.__qsCompleted) {
      try {
        event.completed({ allowEvent: true });
      } catch (e2) {
        /* ignore */
      }
    }
  }
}

function registerLaunchActions() {
  if (typeof Office === "undefined" || !Office.actions || !Office.actions.associate) {
    return false;
  }
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  return true;
}

registerLaunchActions();
if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(function () {
    registerLaunchActions();
  });
}

if (typeof globalThis !== "undefined") {
  globalThis.onMessageSendHandler = onMessageSendHandler;
}
