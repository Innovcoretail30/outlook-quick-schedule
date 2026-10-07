/* global Office, QuickScheduleOffice */

function applyRibbon1h(event) {
  QuickScheduleOffice.applyRibbonSlot(0, event);
}

function applyRibbon2h(event) {
  QuickScheduleOffice.applyRibbonSlot(1, event);
}

function applyRibbon3h(event) {
  QuickScheduleOffice.applyRibbonSlot(2, event);
}

function applyRibbon4h(event) {
  QuickScheduleOffice.applyRibbonSlot(3, event);
}

function applyRibbonTomorrow(event) {
  QuickScheduleOffice.applyRibbonSlot(4, event);
}

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("applyRibbon1h", applyRibbon1h);
  Office.actions.associate("applyRibbon2h", applyRibbon2h);
  Office.actions.associate("applyRibbon3h", applyRibbon3h);
  Office.actions.associate("applyRibbon4h", applyRibbon4h);
  Office.actions.associate("applyRibbonTomorrow", applyRibbonTomorrow);
}
