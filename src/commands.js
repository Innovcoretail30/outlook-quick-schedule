/* global Office, QuickScheduleOffice */

function applyRibbon1h(event) {
  try {
    QuickScheduleOffice.applyRibbonSlot(0, event);
  } catch (e) {
    if (event && !event.__qsCompleted) {
      try {
        event.completed();
      } catch (e2) {
        /* ignore */
      }
    }
  }
}

function applyRibbon2h(event) {
  try {
    QuickScheduleOffice.applyRibbonSlot(1, event);
  } catch (e) {
    if (event && !event.__qsCompleted) {
      try {
        event.completed();
      } catch (e2) {
        /* ignore */
      }
    }
  }
}

function applyRibbon3h(event) {
  try {
    QuickScheduleOffice.applyRibbonSlot(2, event);
  } catch (e) {
    if (event && !event.__qsCompleted) {
      try {
        event.completed();
      } catch (e2) {
        /* ignore */
      }
    }
  }
}

function applyRibbon4h(event) {
  try {
    QuickScheduleOffice.applyRibbonSlot(3, event);
  } catch (e) {
    if (event && !event.__qsCompleted) {
      try {
        event.completed();
      } catch (e2) {
        /* ignore */
      }
    }
  }
}

function applyRibbonTomorrow(event) {
  try {
    QuickScheduleOffice.applyRibbonSlot(4, event);
  } catch (e) {
    if (event && !event.__qsCompleted) {
      try {
        event.completed();
      } catch (e2) {
        /* ignore */
      }
    }
  }
}

function applyClearSchedule(event) {
  try {
    QuickScheduleOffice.applyClearSchedule(event);
  } catch (e) {
    if (event && !event.__qsCompleted) {
      try {
        event.completed();
      } catch (e2) {
        /* ignore */
      }
    }
  }
}

function registerCommandActions() {
  if (typeof Office === "undefined" || !Office.actions || !Office.actions.associate) {
    return false;
  }
  Office.actions.associate("applyRibbon1h", applyRibbon1h);
  Office.actions.associate("applyRibbon2h", applyRibbon2h);
  Office.actions.associate("applyRibbon3h", applyRibbon3h);
  Office.actions.associate("applyRibbon4h", applyRibbon4h);
  Office.actions.associate("applyRibbonTomorrow", applyRibbonTomorrow);
  Office.actions.associate("applyClearSchedule", applyClearSchedule);
  return true;
}

registerCommandActions();
if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(function () {
    registerCommandActions();
  });
}

if (typeof globalThis !== "undefined") {
  globalThis.applyRibbon1h = applyRibbon1h;
  globalThis.applyRibbon2h = applyRibbon2h;
  globalThis.applyRibbon3h = applyRibbon3h;
  globalThis.applyRibbon4h = applyRibbon4h;
  globalThis.applyRibbonTomorrow = applyRibbonTomorrow;
  globalThis.applyClearSchedule = applyClearSchedule;
}
