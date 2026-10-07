/* global Office, QuickScheduleOffice */

function onMessageSendHandler(event) {
  QuickScheduleOffice.applyDelayOnSend(event);
}

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
