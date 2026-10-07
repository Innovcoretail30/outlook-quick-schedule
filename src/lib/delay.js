/**
 * Delay maths for Quick Schedule Send.
 * Works as a script tag (global QuickSchedule) and as a Node module.
 * Keep this file import-free so it can be concatenated into launchevent.js.
 */
(function (root, factory) {
  var api = factory();
  if (typeof module === "object" && module.exports) {
    module.exports = api;
  }
  root.QuickSchedule = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var DEFAULT_TIMEZONE = "Europe/London";
  var SESSION_KEY = "qsSchedule";
  var CUSTOM_PROP_KEY = "qsSchedule";
  var ROAMING_KEY = "qsPresets";
  var NOTIFICATION_KEY = "qsNotify";

  var DEFAULT_PRESETS = [
    { id: "ribbon-1h", label: "1h", kind: "offset", hours: 1, minutes: 0 },
    { id: "ribbon-2h", label: "2h", kind: "offset", hours: 2, minutes: 0 },
    { id: "ribbon-3h", label: "3h", kind: "offset", hours: 3, minutes: 0 },
    { id: "ribbon-4h", label: "4h", kind: "offset", hours: 4, minutes: 0 },
    {
      id: "ribbon-tomorrow",
      label: "Tomorrow 8am",
      kind: "tomorrow",
      hour: 8,
      minute: 0
    }
  ];

  var WINDOWS_TO_IANA = {
    "GMT Standard Time": "Europe/London",
    "Greenwich Standard Time": "Atlantic/Reykjavik",
    UTC: "UTC",
    "UTC Standard Time": "UTC"
  };

  function pad2(n) {
    return String(n).padStart(2, "0");
  }

  function isValidKind(kind) {
    return kind === "offset" || kind === "tomorrow";
  }

  function asInt(value, fallback) {
    var n = Number(value);
    if (!Number.isFinite(n)) {
      return fallback;
    }
    return Math.trunc(n);
  }

  function newId() {
    return "p-" + Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 8);
  }

  function normalizePreset(raw, index) {
    if (!raw || typeof raw !== "object") {
      return null;
    }
    var kind = raw.kind === "tomorrow" ? "tomorrow" : raw.kind === "offset" ? "offset" : null;
    if (!kind) {
      return null;
    }
    var label = String(raw.label || "").trim();
    if (!label) {
      label = kind === "tomorrow" ? "Tomorrow" : "+" + asInt(raw.hours, 0) + "h";
    }
    var preset = {
      id: String(raw.id || newId()),
      label: label.slice(0, 32),
      kind: kind
    };
    if (kind === "offset") {
      var hours = asInt(raw.hours, 0);
      var minutes = asInt(raw.minutes, 0);
      if (hours < 0 || minutes < 0 || (hours === 0 && minutes === 0)) {
        return null;
      }
      if (minutes >= 60) {
        hours += Math.floor(minutes / 60);
        minutes = minutes % 60;
      }
      if (hours > 24 * 14) {
        return null;
      }
      preset.hours = hours;
      preset.minutes = minutes;
    } else {
      var hour = asInt(raw.hour, 8);
      var minute = asInt(raw.minute, 0);
      if (hour < 0 || hour > 23 || minute < 0 || minute > 59) {
        return null;
      }
      preset.hour = hour;
      preset.minute = minute;
    }
    if (typeof index === "number") {
      preset.slot = index;
    }
    return preset;
  }

  function normalizePresets(list) {
    if (!Array.isArray(list)) {
      return DEFAULT_PRESETS.map(function (p) {
        return normalizePreset(p);
      });
    }
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var preset = normalizePreset(list[i], i);
      if (preset) {
        out.push(preset);
      }
    }
    return out;
  }

  function zonedParts(date, timeZone) {
    var fmt = new Intl.DateTimeFormat("en-GB", {
      timeZone: timeZone,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hourCycle: "h23"
    });
    var map = {};
    fmt.formatToParts(date).forEach(function (part) {
      if (part.type !== "literal") {
        map[part.type] = part.value;
      }
    });
    return {
      year: Number(map.year),
      month: Number(map.month),
      day: Number(map.day),
      hour: Number(map.hour),
      minute: Number(map.minute),
      second: Number(map.second)
    };
  }

  function getOffsetMs(instant, timeZone) {
    var parts = zonedParts(new Date(instant), timeZone);
    var asUtc = Date.UTC(
      parts.year,
      parts.month - 1,
      parts.day,
      parts.hour,
      parts.minute,
      parts.second
    );
    return asUtc - instant;
  }

  /**
   * Convert a civil wall time in `timeZone` to a Date (UTC instant).
   * Iterates once so DST spring-forward / fall-back is resolved.
   */
  function zonedLocalToDate(year, month, day, hour, minute, timeZone) {
    var utcGuess = Date.UTC(year, month - 1, day, hour, minute, 0);
    var instant = utcGuess - getOffsetMs(utcGuess, timeZone);
    instant = utcGuess - getOffsetMs(instant, timeZone);
    var check = zonedParts(new Date(instant), timeZone);
    if (
      check.year !== year ||
      check.month !== month ||
      check.day !== day ||
      check.hour !== hour ||
      check.minute !== minute
    ) {
      instant += Date.UTC(year, month - 1, day, hour, minute, 0) -
        Date.UTC(check.year, check.month - 1, check.day, check.hour, check.minute, 0);
    }
    return new Date(instant);
  }

  function addCalendarDays(parts, days) {
    var utc = Date.UTC(parts.year, parts.month - 1, parts.day + days);
    var d = new Date(utc);
    return {
      year: d.getUTCFullYear(),
      month: d.getUTCMonth() + 1,
      day: d.getUTCDate()
    };
  }

  function addOffset(now, hours, minutes) {
    var ms = (asInt(hours, 0) * 60 + asInt(minutes, 0)) * 60 * 1000;
    if (ms <= 0) {
      throw new Error("Offset must be greater than zero.");
    }
    return new Date(now.getTime() + ms);
  }

  function tomorrowAt(now, hour, minute, timeZone) {
    var tz = timeZone || DEFAULT_TIMEZONE;
    var today = zonedParts(now, tz);
    var next = addCalendarDays(today, 1);
    return zonedLocalToDate(next.year, next.month, next.day, hour, minute, tz);
  }

  function applyPresetToDate(preset, now, timeZone) {
    var clean = normalizePreset(preset);
    if (!clean) {
      throw new Error("Invalid schedule preset.");
    }
    var when = now instanceof Date ? now : new Date(now);
    if (Number.isNaN(when.getTime())) {
      throw new Error("Invalid start time.");
    }
    if (clean.kind === "offset") {
      return addOffset(when, clean.hours, clean.minutes);
    }
    return tomorrowAt(when, clean.hour, clean.minute, timeZone || DEFAULT_TIMEZONE);
  }

  function humanOffset(hours, minutes) {
    var h = asInt(hours, 0);
    var m = asInt(minutes, 0);
    var parts = [];
    if (h > 0) {
      parts.push(h === 1 ? "1 hour" : h + " hours");
    }
    if (m > 0) {
      parts.push(m === 1 ? "1 minute" : m + " minutes");
    }
    return parts.join(" ");
  }

  var NOTIFY_MAX = 150;
  var DRAFT_HINT = "To change after sending: open it from Drafts.";
  var CLEARED_MESSAGE = "Schedule cleared, will send immediately";

  function formatLocalHm(date) {
    return pad2(date.getHours()) + ":" + pad2(date.getMinutes());
  }

  function formatLocalTomorrow(now) {
    var d = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
    var weekday = d.toLocaleDateString("en-GB", { weekday: "short" });
    var month = d.toLocaleDateString("en-GB", { month: "short" });
    return weekday + " " + d.getDate() + " " + month;
  }

  function packNotification(line1, line2) {
    var options = [
      line1 + "\n" + line2,
      line1 + " " + line2,
      line1 + "\nChange later: open from Drafts.",
      line1 + " Change later: open from Drafts.",
      line1
    ];
    for (var i = 0; i < options.length; i++) {
      if (options[i].length <= NOTIFY_MAX) {
        return options[i];
      }
    }
    return line1.slice(0, NOTIFY_MAX);
  }

  function describeScheduledBar(preset, now) {
    var clean = normalizePreset(preset);
    if (!clean) {
      return {
        message: "No send delay selected.",
        line1: "No send delay selected.",
        summary: "None",
        persistent: false
      };
    }
    var when = now instanceof Date ? now : new Date(now || Date.now());
    var line1;
    var summary;
    if (clean.kind === "offset") {
      var human = humanOffset(clean.hours, clean.minutes);
      var about = formatLocalHm(addOffset(when, clean.hours, clean.minutes));
      line1 =
        "Scheduled: sends " +
        human +
        " after you press Send (about " +
        about +
        " if you send now)";
      summary = human + " after Send";
    } else {
      var clock = pad2(clean.hour) + ":" + pad2(clean.minute);
      line1 =
        "Scheduled: sends tomorrow (" +
        formatLocalTomorrow(when) +
        ") at " +
        clock;
      summary = "Tomorrow " + clock + " after Send";
    }
    return {
      message: packNotification(line1, DRAFT_HINT),
      line1: line1,
      summary: summary,
      persistent: true
    };
  }

  function describeSchedule(preset, now) {
    return describeScheduledBar(preset, now);
  }

  function resolveTimeZone(officeTimeZone) {
    try {
      var iana = Intl.DateTimeFormat().resolvedOptions().timeZone;
      if (iana) {
        return iana;
      }
    } catch (e) {
      /* ignore */
    }
    if (officeTimeZone && WINDOWS_TO_IANA[officeTimeZone]) {
      return WINDOWS_TO_IANA[officeTimeZone];
    }
    return DEFAULT_TIMEZONE;
  }

  function serializeSchedule(preset) {
    var clean = normalizePreset(preset);
    if (!clean) {
      return "";
    }
    return JSON.stringify({ v: 1, preset: clean });
  }

  function parseSchedule(raw) {
    if (!raw) {
      return null;
    }
    try {
      var data = typeof raw === "string" ? JSON.parse(raw) : raw;
      if (!data || data.cleared) {
        return null;
      }
      return normalizePreset(data.preset || data);
    } catch (e) {
      return null;
    }
  }

  return {
    DEFAULT_TIMEZONE: DEFAULT_TIMEZONE,
    DEFAULT_PRESETS: DEFAULT_PRESETS,
    SESSION_KEY: SESSION_KEY,
    CUSTOM_PROP_KEY: CUSTOM_PROP_KEY,
    ROAMING_KEY: ROAMING_KEY,
    NOTIFICATION_KEY: NOTIFICATION_KEY,
    isValidKind: isValidKind,
    newId: newId,
    normalizePreset: normalizePreset,
    normalizePresets: normalizePresets,
    zonedParts: zonedParts,
    zonedLocalToDate: zonedLocalToDate,
    addCalendarDays: addCalendarDays,
    addOffset: addOffset,
    tomorrowAt: tomorrowAt,
    applyPresetToDate: applyPresetToDate,
    describeSchedule: describeSchedule,
    describeScheduledBar: describeScheduledBar,
    packNotification: packNotification,
    formatLocalHm: formatLocalHm,
    formatLocalTomorrow: formatLocalTomorrow,
    NOTIFY_MAX: NOTIFY_MAX,
    DRAFT_HINT: DRAFT_HINT,
    CLEARED_MESSAGE: CLEARED_MESSAGE,
    humanOffset: humanOffset,
    resolveTimeZone: resolveTimeZone,
    serializeSchedule: serializeSchedule,
    parseSchedule: parseSchedule,
    pad2: pad2
  };
});
