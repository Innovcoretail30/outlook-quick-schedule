/**
 * Office.js helpers for choosing a delay and applying it on Send.
 * Concatenated into launchevent.js after delay.js. Uses global QuickSchedule.
 *
 * Writes are allowed only after an explicit compose button click, or in
 * OnMessageSend for that same compose item. Select/read must never write.
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

  function currentItem() {
    try {
      return Office.context && Office.context.mailbox && Office.context.mailbox.item;
    } catch (e) {
      return null;
    }
  }

  function isComposeItem(item) {
    if (!item) {
      return false;
    }
    if (!item.sessionData) {
      return false;
    }
    return true;
  }

  function withComposeItem(cb, onMissing) {
    var item = currentItem();
    if (!isComposeItem(item)) {
      if (onMissing) {
        onMissing(new Error("Not a compose item."));
      }
      return;
    }
    cb(item);
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

  function stillIntended(intended) {
    var item = currentItem();
    return isComposeItem(item) && QS.fingerprintsMatch(intended, item);
  }

  function canUseSessionData(item) {
    return (
      mailboxSet("1.11") &&
      isComposeItem(item) &&
      typeof item.sessionData.getAsync === "function" &&
      typeof item.sessionData.setAsync === "function"
    );
  }

  function canNotify(item) {
    return mailboxSet("1.3") && item && item.notificationMessages;
  }

  function readSchedule(callback) {
    var done = once(callback);
    withComposeItem(
      function (item) {
        if (!canUseSessionData(item)) {
          done(null);
          return;
        }
        item.sessionData.getAsync(QS.SESSION_KEY, function (result) {
          if (!result || result.status === Office.AsyncResultStatus.Failed) {
            done(null);
            return;
          }
          done(QS.parseSchedule(result.value, item));
        });
      },
      function () {
        done(null);
      }
    );
  }

  function writeSession(item, payload, done) {
    if (!canUseSessionData(item)) {
      done();
      return;
    }
    try {
      if (payload) {
        item.sessionData.setAsync(QS.SESSION_KEY, payload, function () {
          done();
        });
      } else if (typeof item.sessionData.removeAsync === "function") {
        item.sessionData.removeAsync(QS.SESSION_KEY, function () {
          done();
        });
      } else {
        item.sessionData.setAsync(QS.SESSION_KEY, "", function () {
          done();
        });
      }
    } catch (e) {
      done();
    }
  }

  function replaceInfoNotification(item, message, persistent, done) {
    if (!canNotify(item)) {
      if (done) {
        done();
      }
      return;
    }
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

  function writeSchedule(preset, callback) {
    var done = once(callback);
    later(done, 2500);
    withComposeItem(
      function (item) {
        var plan = QS.planWrites({
          type: "composeClick",
          action: preset ? "choose" : "clear",
          item: { surface: "compose", itemId: item.itemId, conversationId: item.conversationId, sessionData: true },
          preset: preset
        });
        if (!plan.writes.length) {
          done();
          return;
        }
        var intended = QS.itemFingerprint(item);
        var payload = preset
          ? QS.serializeSchedule(preset, {
              itemId: intended.itemId,
              conversationId: intended.conversationId
            })
          : "";
        writeSession(item, payload, function () {
          if (!stillIntended(intended)) {
            done();
            return;
          }
          var live = currentItem();
          if (preset) {
            replaceInfoNotification(live, QS.describeScheduledBar(preset, new Date()).message, true, done);
          } else {
            replaceInfoNotification(live, QS.CLEARED_MESSAGE, false, done);
          }
        });
      },
      function () {
        done();
      }
    );
  }

  function showChosenNotification(preset, callback) {
    var done = once(callback);
    later(done, 2500);
    withComposeItem(
      function (item) {
        replaceInfoNotification(item, QS.describeScheduledBar(preset, new Date()).message, true, done);
      },
      function () {
        done();
      }
    );
  }

  function showClearedNotification(callback) {
    var done = once(callback);
    later(done, 2500);
    withComposeItem(
      function (item) {
        replaceInfoNotification(item, CLEARED_MESSAGE, false, done);
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
      if (onDone) {
        onDone(null, clean);
      }
    });
  }

  function clearSchedule(onDone) {
    writeSchedule(null, function () {
      if (onDone) {
        onDone();
      }
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
      var item = currentItem();
      if (!isComposeItem(item) || !canUseSessionData(item)) {
        finish({ allowEvent: true });
        return;
      }
      item.sessionData.getAsync(QS.SESSION_KEY, function (result) {
        var raw = result && result.status !== Office.AsyncResultStatus.Failed ? result.value : "";
        var live = currentItem();
        var plan = QS.planWrites({
          type: "send",
          item: {
            surface: "compose",
            itemId: live && live.itemId,
            conversationId: live && live.conversationId,
            sessionData: true
          },
          sessionData: raw,
          customProperties: null
        });
        if (!plan.applyDelay) {
          finish({ allowEvent: true });
          return;
        }
        if (!stillIntended(live)) {
          finish({ allowEvent: true });
          return;
        }
        var when;
        try {
          when = QS.applyPresetToDate(plan.preset, new Date(), officeTimeZone());
        } catch (e) {
          finish({ allowEvent: true });
          return;
        }
        if (!mailboxSet("1.13") || !live.delayDeliveryTime || !live.delayDeliveryTime.setAsync) {
          finish({ allowEvent: true });
          return;
        }
        live.delayDeliveryTime.setAsync(when, function () {
          var after = currentItem();
          if (isComposeItem(after) && canUseSessionData(after) && QS.fingerprintsMatch(live, after)) {
            writeSession(after, "", function () {
              finish({ allowEvent: true });
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
      if (!isComposeItem(currentItem())) {
        finish();
        return;
      }
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
      if (!isComposeItem(currentItem())) {
        finish();
        return;
      }
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
    readRoamingPresets: readRoamingPresets,
    saveRoamingPresets: saveRoamingPresets,
    applyDelayOnSend: applyDelayOnSend,
    applyRibbonSlot: applyRibbonSlot,
    applyClearSchedule: applyClearSchedule
  };
})(typeof globalThis !== "undefined" ? globalThis : this);
