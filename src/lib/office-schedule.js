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

  function officeTimeZone() {
    try {
      return QS.resolveTimeZone(Office.context.mailbox.userProfile.timeZone);
    } catch (e) {
      return QS.resolveTimeZone();
    }
  }

  function withItem(cb, onMissing) {
    try {
      var item = Office.context.mailbox.item;
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

  function getCustomProperties(item, callback) {
    if (!item.loadCustomPropertiesAsync) {
      callback(null);
      return;
    }
    item.loadCustomPropertiesAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        callback(null);
        return;
      }
      callback(result.value);
    });
  }

  function readSchedule(callback) {
    withItem(
      function (item) {
        function fromCustom() {
          getCustomProperties(item, function (props) {
            if (!props) {
              callback(null);
              return;
            }
            callback(QS.parseSchedule(props.get(QS.CUSTOM_PROP_KEY)));
          });
        }

        if (!item.sessionData || !item.sessionData.getAsync) {
          fromCustom();
          return;
        }
        item.sessionData.getAsync(QS.SESSION_KEY, function (result) {
          if (result.status !== Office.AsyncResultStatus.Failed) {
            var parsed = QS.parseSchedule(result.value);
            if (parsed) {
              callback(parsed);
              return;
            }
          }
          fromCustom();
        });
      },
      function () {
        callback(null);
      }
    );
  }

  function writeCustom(item, value, done) {
    getCustomProperties(item, function (props) {
      if (!props) {
        done();
        return;
      }
      if (value) {
        props.set(QS.CUSTOM_PROP_KEY, value);
      } else {
        props.remove(QS.CUSTOM_PROP_KEY);
      }
      props.saveAsync(function () {
        done();
      });
    });
  }

  function writeSchedule(preset, callback) {
    var payload = preset ? QS.serializeSchedule(preset) : "";
    withItem(
      function (item) {
        var pending = 2;
        function finish() {
          pending -= 1;
          if (pending <= 0 && callback) {
            callback();
          }
        }

        if (item.sessionData && item.sessionData.setAsync) {
          if (payload) {
            item.sessionData.setAsync(QS.SESSION_KEY, payload, finish);
          } else if (item.sessionData.removeAsync) {
            item.sessionData.removeAsync(QS.SESSION_KEY, finish);
          } else {
            item.sessionData.setAsync(QS.SESSION_KEY, "", finish);
          }
        } else {
          finish();
        }
        writeCustom(item, payload, finish);
      },
      function () {
        if (callback) {
          callback();
        }
      }
    );
  }

  function clearNotification(done) {
    withItem(
      function (item) {
        if (!item.notificationMessages) {
          if (done) {
            done();
          }
          return;
        }
        item.notificationMessages.removeAsync(QS.NOTIFICATION_KEY, function () {
          if (done) {
            done();
          }
        });
      },
      function () {
        if (done) {
          done();
        }
      }
    );
  }

  function showChosenNotification(preset, done) {
    withItem(
      function (item) {
        if (!item.notificationMessages) {
          if (done) {
            done();
          }
          return;
        }
        var copy = QS.describeSchedule(preset);
        var insightSupported =
          Office.context.requirements &&
          Office.context.requirements.isSetSupported("Mailbox", "1.10");

        function addInfo() {
          item.notificationMessages.replaceAsync(
            QS.NOTIFICATION_KEY,
            {
              type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
              message: copy.message,
              icon: "Icon.16x16",
              persistent: true
            },
            function () {
              if (done) {
                done();
              }
            }
          );
        }

        if (!insightSupported) {
          addInfo();
          return;
        }

        item.notificationMessages.replaceAsync(
          QS.NOTIFICATION_KEY,
          {
            type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
            message: copy.message,
            icon: "Icon.16x16",
            actions: [
              {
                actionText: "Clear schedule",
                actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
                commandId: "btnEditTimes",
                contextData: { action: "clear" }
              }
            ]
          },
          function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
              addInfo();
              return;
            }
            if (done) {
              done();
            }
          }
        );
      },
      function () {
        if (done) {
          done();
        }
      }
    );
  }

  function readRoamingPresets() {
    try {
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
    Office.context.roamingSettings.set(QS.ROAMING_KEY, JSON.stringify(list));
    Office.context.roamingSettings.saveAsync(function (result) {
      if (callback) {
        callback(result);
      }
    });
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
      clearNotification(function () {
        if (onDone) {
          onDone();
        }
      });
    });
  }

  function applyDelayOnSend(event) {
    readSchedule(function (preset) {
      if (!preset) {
        event.completed({ allowEvent: true });
        return;
      }
      var when;
      try {
        when = QS.applyPresetToDate(preset, new Date(), officeTimeZone());
      } catch (e) {
        event.completed({
          allowEvent: true,
          errorMessage: "Quick Schedule Send could not calculate the delay. The message will send now."
        });
        return;
      }

      var item = Office.context.mailbox.item;
      if (!item.delayDeliveryTime || !item.delayDeliveryTime.setAsync) {
        event.completed({
          allowEvent: true,
          errorMessage: "This Outlook build cannot set delayDeliveryTime (Mailbox 1.13). The message will send now."
        });
        return;
      }

      item.delayDeliveryTime.setAsync(when, function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          event.completed({
            allowEvent: true,
            errorMessage:
              "Quick Schedule Send could not set delayed delivery. The message will send now. " +
              (result.error && result.error.message ? result.error.message : "")
          });
          return;
        }
        event.completed({ allowEvent: true });
      });
    });
  }

  function applyRibbonSlot(slotIndex, event) {
    var presets = readRoamingPresets();
    var preset = presets[slotIndex] || QS.DEFAULT_PRESETS[slotIndex];
    choosePreset(preset, function () {
      if (event) {
        event.completed();
      }
    });
  }

  root.QuickScheduleOffice = {
    officeTimeZone: officeTimeZone,
    readSchedule: readSchedule,
    writeSchedule: writeSchedule,
    choosePreset: choosePreset,
    clearSchedule: clearSchedule,
    showChosenNotification: showChosenNotification,
    readRoamingPresets: readRoamingPresets,
    saveRoamingPresets: saveRoamingPresets,
    applyDelayOnSend: applyDelayOnSend,
    applyRibbonSlot: applyRibbonSlot
  };
})(typeof globalThis !== "undefined" ? globalThis : this);
