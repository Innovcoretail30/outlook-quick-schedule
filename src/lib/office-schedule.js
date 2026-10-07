/**
 * Office.js helpers for choosing a delay and applying it on Send.
 * Concatenated into launchevent.js after delay.js. Uses global QuickSchedule.
 */
(function (root) {
  "use strict";

  var QS = root.QuickSchedule;
  if (!QS) {
    throw new Error("QuickSchedule delay library must load first.");
  }

  var NOTIFY_ICON = "Icon.16x16";
  var CLEARED_MESSAGE = QS.CLEARED_MESSAGE;

  function later(fn, ms) {
    var t = typeof setTimeout === "function" ? setTimeout : null;
    if (t) {
      t(fn, ms);
    } else {
      fn();
    }
  }

  function mailboxSet(version) {
    try {
      return !!(
        Office &&
        Office.context &&
        Office.context.requirements &&
        Office.context.requirements.isSetSupported("Mailbox", version)
      );
    } catch (e) {
      return false;
    }
  }

  function officeTimeZone() {
    try {
      return QS.resolveTimeZone(Office.context.mailbox.userProfile.timeZone);
    } catch (e) {
      return QS.resolveTimeZone();
    }
  }

  function withItem(cb, onMissing) {
    try {
      var item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
      if (!item) {
        if (onMissing) {
          onMissing(new Error("No compose item."));
        }
        return;
      }
      cb(item);
    } catch (e) {
      if (onMissing) {
        onMissing(e);
      }
    }
  }

  function once(fn) {
    var called = false;
    return function () {
      if (called) {
        return;
      }
      called = true;
      if (fn) {
        fn.apply(null, arguments);
      }
    };
  }

  function getCustomProperties(item, callback) {
    if (!item.loadCustomPropertiesAsync) {
      callback(null);
      return;
    }
    item.loadCustomPropertiesAsync(function (result) {
      if (!result || result.status === Office.AsyncResultStatus.Failed) {
        callback(null);
        return;
      }
      callback(result.value);
    });
  }

  function canUseSessionData(item) {
    return (
      mailboxSet("1.11") &&
      item &&
      item.sessionData &&
      typeof item.sessionData.getAsync === "function" &&
      typeof item.sessionData.setAsync === "function"
    );
  }

  function canNotify(item) {
    return mailboxSet("1.3") && item && item.notificationMessages;
  }

  function readSchedule(callback) {
    var done = once(callback);
    withItem(
      function (item) {
        function fromCustom() {
          getCustomProperties(item, function (props) {
            if (!props) {
              done(null);
              return;
            }
            try {
              done(QS.parseSchedule(props.get(QS.CUSTOM_PROP_KEY)));
            } catch (e) {
              done(null);
            }
          });
        }

        if (!canUseSessionData(item)) {
          fromCustom();
          return;
        }
        item.sessionData.getAsync(QS.SESSION_KEY, function (result) {
          if (result && result.status !== Office.AsyncResultStatus.Failed) {
            var parsed = QS.parseSchedule(result.value);
            if (parsed) {
              done(parsed);
              return;
            }
          }
          fromCustom();
        });
      },
      function () {
        done(null);
      }
    );
  }

  function writeCustom(item, value, done) {
    getCustomProperties(item, function (props) {
      if (!props) {
        done();
        return;
      }
      try {
        if (value) {
          props.set(QS.CUSTOM_PROP_KEY, value);
        } else if (props.remove) {
          props.remove(QS.CUSTOM_PROP_KEY);
        } else {
          props.set(QS.CUSTOM_PROP_KEY, "");
        }
        props.saveAsync(function () {
          done();
        });
      } catch (e) {
        done();
      }
    });
  }

  function writeSchedule(preset, callback) {
    var payload = preset ? QS.serializeSchedule(preset) : "";
    var done = once(callback);
    later(done, 2500);
    withItem(
      function (item) {
        var pending = 2;
        function finish() {
          pending -= 1;
          if (pending <= 0) {
            done();
          }
        }

        if (canUseSessionData(item)) {
          try {
            if (payload) {
              item.sessionData.setAsync(QS.SESSION_KEY, payload, finish);
            } else if (typeof item.sessionData.removeAsync === "function") {
              item.sessionData.removeAsync(QS.SESSION_KEY, finish);
            } else {
              item.sessionData.setAsync(QS.SESSION_KEY, "", finish);
            }
          } catch (e) {
            finish();
          }
        } else {
          finish();
        }
        writeCustom(item, payload, finish);
      },
      function () {
        done();
      }
    );
  }

  function replaceInfoNotification(item, message, persistent, done) {
    try {
      item.notificationMessages.replaceAsync(
        QS.NOTIFICATION_KEY,
        {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: String(message).slice(0, QS.NOTIFY_MAX),
          icon: NOTIFY_ICON,
          persistent: persistent !== false
        },
        function () {
          if (done) {
            done();
          }
        }
      );
    } catch (e) {
      if (done) {
        done();
      }
    }
  }

  function showChosenNotification(preset, callback) {
    var done = once(callback);
    later(done, 2500);
    withItem(
      function (item) {
        if (!canNotify(item)) {
          done();
          return;
        }
        var copy = QS.describeScheduledBar(preset, new Date());
        replaceInfoNotification(item, copy.message, true, done);
      },
      function () {
        done();
      }
    );
  }

  function showClearedNotification(callback) {
    var done = once(callback);
    later(done, 2500);
    withItem(
      function (item) {
        if (!canNotify(item)) {
          done();
          return;
        }
        replaceInfoNotification(item, CLEARED_MESSAGE, false, done);
      },
      function () {
        done();
      }
    );
  }

  function restoreScheduledNotification(callback) {
    readSchedule(function (preset) {
      if (!preset) {
        if (callback) {
          callback(null);
        }
        return;
      }
      showChosenNotification(preset, function () {
        if (callback) {
          callback(preset);
        }
      });
    });
  }

  function clearDelayDeliveryTime(callback) {
    var done = once(callback);
    later(done, 2500);
    withItem(
      function (item) {
        if (!mailboxSet("1.13") || !item.delayDeliveryTime || !item.delayDeliveryTime.setAsync) {
          done();
          return;
        }
        try {
          item.delayDeliveryTime.setAsync(new Date(0), function () {
            done();
          });
        } catch (e) {
          done();
        }
      },
      function () {
        done();
      }
    );
  }

  function readRoamingPresets() {
    try {
      if (!Office.context || !Office.context.roamingSettings) {
        return QS.normalizePresets(QS.DEFAULT_PRESETS);
      }
      var raw = Office.context.roamingSettings.get(QS.ROAMING_KEY);
      if (!raw) {
        return QS.normalizePresets(QS.DEFAULT_PRESETS);
      }
      var parsed = typeof raw === "string" ? JSON.parse(raw) : raw;
      var list = QS.normalizePresets(parsed);
      return list.length ? list : QS.normalizePresets(QS.DEFAULT_PRESETS);
    } catch (e) {
      return QS.normalizePresets(QS.DEFAULT_PRESETS);
    }
  }

  function saveRoamingPresets(presets, callback) {
    var list = QS.normalizePresets(presets);
    try {
      Office.context.roamingSettings.set(QS.ROAMING_KEY, JSON.stringify(list));
      Office.context.roamingSettings.saveAsync(function (result) {
        if (callback) {
          callback(result);
        }
      });
    } catch (e) {
      if (callback) {
        callback({ status: "failed", error: e });
      }
    }
  }

  function choosePreset(preset, onDone) {
    var clean = QS.normalizePreset(preset);
    if (!clean) {
      if (onDone) {
        onDone(new Error("Invalid preset."));
      }
      return;
    }
    writeSchedule(clean, function () {
      showChosenNotification(clean, function () {
        if (onDone) {
          onDone(null, clean);
        }
      });
    });
  }

  function clearSchedule(onDone) {
    writeSchedule(null, function () {
      clearDelayDeliveryTime(function () {
        showClearedNotification(function () {
          if (onDone) {
            onDone();
          }
        });
      });
    });
  }

  function completeCommand(event) {
    if (!event || event.__qsCompleted) {
      return;
    }
    event.__qsCompleted = true;
    try {
      event.completed();
    } catch (e) {
      /* already completed */
    }
  }

  function completeSend(event, options) {
    if (!event || event.__qsCompleted) {
      return;
    }
    event.__qsCompleted = true;
    try {
      event.completed(options || { allowEvent: true });
    } catch (e) {
      try {
        event.completed({ allowEvent: true });
      } catch (e2) {
        /* already completed */
      }
    }
  }

  function applyDelayOnSend(event) {
    var finish = function (options) {
      completeSend(event, options || { allowEvent: true });
    };
    later(function () {
      finish({ allowEvent: true });
    }, 5000);
    try {
      readSchedule(function (preset) {
        if (!preset) {
          finish({ allowEvent: true });
          return;
        }
        var when;
        try {
          when = QS.applyPresetToDate(preset, new Date(), officeTimeZone());
        } catch (e) {
          finish({
            allowEvent: true,
            errorMessage: "Quick Schedule Send could not calculate the delay. The message will send now."
          });
          return;
        }

        var item;
        try {
          item = Office.context.mailbox.item;
        } catch (e) {
          finish({ allowEvent: true });
          return;
        }
        if (!mailboxSet("1.13") || !item.delayDeliveryTime || !item.delayDeliveryTime.setAsync) {
          finish({
            allowEvent: true,
            errorMessage: "This Outlook build cannot set delayDeliveryTime (Mailbox 1.13). The message will send now."
          });
          return;
        }

        item.delayDeliveryTime.setAsync(when, function (result) {
          if (result && result.status === Office.AsyncResultStatus.Failed) {
            finish({
              allowEvent: true,
              errorMessage:
                "Quick Schedule Send could not set delayed delivery. The message will send now. " +
                (result.error && result.error.message ? result.error.message : "")
            });
            return;
          }
          finish({ allowEvent: true });
        });
      });
    } catch (e) {
      finish({ allowEvent: true });
    }
  }

  function applyRibbonSlot(slotIndex, event) {
    var finish = function () {
      completeCommand(event);
    };
    later(finish, 4000);
    try {
      var presets = readRoamingPresets();
      var preset = presets[slotIndex] || QS.DEFAULT_PRESETS[slotIndex];
      choosePreset(preset, function () {
        finish();
      });
    } catch (e) {
      finish();
    }
  }

  function applyClearSchedule(event) {
    var finish = function () {
      completeCommand(event);
    };
    later(finish, 4000);
    try {
      clearSchedule(function () {
        finish();
      });
    } catch (e) {
      finish();
    }
  }

  root.QuickScheduleOffice = {
    officeTimeZone: officeTimeZone,
    readSchedule: readSchedule,
    writeSchedule: writeSchedule,
    choosePreset: choosePreset,
    clearSchedule: clearSchedule,
    showChosenNotification: showChosenNotification,
    showClearedNotification: showClearedNotification,
    restoreScheduledNotification: restoreScheduledNotification,
    readRoamingPresets: readRoamingPresets,
    saveRoamingPresets: saveRoamingPresets,
    applyDelayOnSend: applyDelayOnSend,
    applyRibbonSlot: applyRibbonSlot,
    applyClearSchedule: applyClearSchedule
  };
})(typeof globalThis !== "undefined" ? globalThis : this);
