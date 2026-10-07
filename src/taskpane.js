/* global Office, QuickSchedule, QuickScheduleOffice */

(function () {
  "use strict";

  var QS = QuickSchedule;
  var inOffice = false;
  var previewPresetsKey = "qs-preview-presets";
  var previewScheduleKey = "qs-preview-schedule";
  var presets = [];
  var editingId = null;
  var currentPreset = null;

  var els = {
    banner: document.getElementById("previewBanner"),
    status: document.getElementById("currentStatus"),
    quickButtons: document.getElementById("quickButtons"),
    clearBtn: document.getElementById("clearBtn"),
    actionToast: document.getElementById("actionToast"),
    editorState: document.getElementById("editorState"),
    presetList: document.getElementById("presetList"),
    presetForm: document.getElementById("presetForm"),
    fieldLabel: document.getElementById("fieldLabel"),
    fieldKind: document.getElementById("fieldKind"),
    offsetFields: document.getElementById("offsetFields"),
    tomorrowFields: document.getElementById("tomorrowFields"),
    fieldHours: document.getElementById("fieldHours"),
    fieldMinutes: document.getElementById("fieldMinutes"),
    fieldHour: document.getElementById("fieldHour"),
    fieldMinute: document.getElementById("fieldMinute"),
    saveItemBtn: document.getElementById("saveItemBtn"),
    cancelEditBtn: document.getElementById("cancelEditBtn"),
    addNewBtn: document.getElementById("addNewBtn"),
    saveAllBtn: document.getElementById("saveAllBtn"),
    restoreBtn: document.getElementById("restoreBtn"),
    saveToast: document.getElementById("saveToast")
  };

  function showToast(el, text) {
    el.textContent = text;
    el.dataset.show = "true";
    window.setTimeout(function () {
      el.dataset.show = "false";
    }, 3200);
  }

  function setStatus(state, text) {
    els.status.dataset.state = state;
    els.status.textContent = text;
  }

  function presetSummary(preset) {
    if (preset.kind === "offset") {
      return "+" + QS.humanOffset(preset.hours, preset.minutes) + " from Send";
    }
    return "Tomorrow " + QS.pad2(preset.hour) + ":" + QS.pad2(preset.minute) + " local";
  }

  function loadPreviewPresets() {
    try {
      var raw = window.localStorage.getItem(previewPresetsKey);
      var list = QS.normalizePresets(raw ? JSON.parse(raw) : QS.DEFAULT_PRESETS);
      return list.length ? list : QS.normalizePresets(QS.DEFAULT_PRESETS);
    } catch (e) {
      return QS.normalizePresets(QS.DEFAULT_PRESETS);
    }
  }

  function savePreviewPresets(list) {
    window.localStorage.setItem(previewPresetsKey, JSON.stringify(list));
  }

  function readPresets() {
    if (inOffice) {
      return QuickScheduleOffice.readRoamingPresets();
    }
    return loadPreviewPresets();
  }

  function persistPresets(list, callback) {
    presets = QS.normalizePresets(list);
    if (inOffice) {
      QuickScheduleOffice.saveRoamingPresets(presets, function (result) {
        if (result && result.status === Office.AsyncResultStatus.Failed) {
          els.editorState.className = "error";
          els.editorState.textContent = "Could not save to your mailbox: " + result.error.message;
          if (callback) {
            callback(new Error(result.error.message));
          }
          return;
        }
        els.editorState.className = "empty";
        els.editorState.textContent = "";
        if (callback) {
          callback(null);
        }
      });
      return;
    }
    savePreviewPresets(presets);
    els.editorState.className = "empty";
    els.editorState.textContent = "";
    if (callback) {
      callback(null);
    }
  }

  function refreshCurrent() {
    function paint(preset) {
      currentPreset = preset;
      if (!preset) {
        setStatus("empty", "No delay chosen. This message will send immediately.");
        els.clearBtn.disabled = true;
        return;
      }
      setStatus("ok", QS.describeScheduledBar(preset, new Date()).message);
      els.clearBtn.disabled = false;
    }

    if (inOffice) {
      QuickScheduleOffice.readSchedule(paint);
      return;
    }
    try {
      paint(QS.parseSchedule(window.localStorage.getItem(previewScheduleKey)));
    } catch (e) {
      paint(null);
    }
  }

  function choose(preset) {
    if (inOffice) {
      QuickScheduleOffice.choosePreset(preset, function (err, clean) {
        if (err) {
          setStatus("error", err.message);
          return;
        }
        currentPreset = clean;
        setStatus("ok", QS.describeScheduledBar(clean, new Date()).message);
        els.clearBtn.disabled = false;
        showToast(els.actionToast, "Delay chosen. Press Send when the message is ready.");
      });
      return;
    }
    var clean = QS.normalizePreset(preset);
    window.localStorage.setItem(previewScheduleKey, QS.serializeSchedule(clean));
    refreshCurrent();
    showToast(els.actionToast, "Preview only — in Outlook this is stored on the draft.");
  }

  function clearChosen() {
    if (inOffice) {
      QuickScheduleOffice.clearSchedule(function () {
        currentPreset = null;
        setStatus("empty", QS.CLEARED_MESSAGE);
        els.clearBtn.disabled = true;
        showToast(els.actionToast, "Schedule cleared.");
      });
      return;
    }
    window.localStorage.removeItem(previewScheduleKey);
    refreshCurrent();
    showToast(els.actionToast, "Schedule cleared.");
  }

  function renderQuickButtons() {
    els.quickButtons.replaceChildren();
    if (!presets.length) {
      var empty = document.createElement("p");
      empty.className = "empty";
      empty.textContent = "No buttons yet. Add one below and press Save.";
      els.quickButtons.appendChild(empty);
      return;
    }
    presets.forEach(function (preset) {
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "preset-btn";
      btn.textContent = preset.label;
      btn.title = presetSummary(preset);
      btn.addEventListener("click", function () {
        choose(preset);
      });
      els.quickButtons.appendChild(btn);
    });
  }

  function renderEditor() {
    els.presetList.replaceChildren();
    if (!presets.length) {
      els.editorState.className = "empty";
      els.editorState.textContent = "No presets yet. Add one below.";
      return;
    }
    els.editorState.className = "empty";
    els.editorState.textContent = "";
    presets.forEach(function (preset, index) {
      var li = document.createElement("li");
      li.className = "preset-row";
      var info = document.createElement("div");
      var title = document.createElement("strong");
      title.textContent = preset.label;
      var meta = document.createElement("span");
      var ribbon = index < 5 ? " · ribbon slot “" + QS.DEFAULT_PRESETS[index].label + "”" : "";
      meta.textContent = presetSummary(preset) + ribbon;
      info.appendChild(title);
      info.appendChild(meta);
      var tools = document.createElement("div");
      tools.className = "tools";
      var edit = document.createElement("button");
      edit.type = "button";
      edit.className = "btn secondary";
      edit.textContent = "Edit";
      edit.addEventListener("click", function () {
        startEdit(preset);
      });
      var remove = document.createElement("button");
      remove.type = "button";
      remove.className = "btn danger";
      remove.textContent = "Remove";
      remove.addEventListener("click", function () {
        presets = presets.filter(function (p) {
          return p.id !== preset.id;
        });
        renderQuickButtons();
        renderEditor();
      });
      tools.appendChild(edit);
      tools.appendChild(remove);
      li.appendChild(info);
      li.appendChild(tools);
      els.presetList.appendChild(li);
    });
  }

  function toggleKindFields() {
    var tomorrow = els.fieldKind.value === "tomorrow";
    els.offsetFields.hidden = tomorrow;
    els.tomorrowFields.hidden = !tomorrow;
  }

  function resetForm() {
    editingId = null;
    els.presetForm.hidden = true;
    els.fieldLabel.value = "";
    els.fieldKind.value = "offset";
    els.fieldHours.value = "1";
    els.fieldMinutes.value = "0";
    els.fieldHour.value = "8";
    els.fieldMinute.value = "0";
    els.saveItemBtn.textContent = "Add button";
    els.cancelEditBtn.hidden = true;
    toggleKindFields();
  }

  function startEdit(preset) {
    editingId = preset.id;
    els.presetForm.hidden = false;
    els.fieldLabel.value = preset.label;
    els.fieldKind.value = preset.kind;
    if (preset.kind === "offset") {
      els.fieldHours.value = String(preset.hours);
      els.fieldMinutes.value = String(preset.minutes);
    } else {
      els.fieldHour.value = String(preset.hour);
      els.fieldMinute.value = String(preset.minute);
    }
    els.saveItemBtn.textContent = "Update button";
    els.cancelEditBtn.hidden = false;
    toggleKindFields();
    els.fieldLabel.focus();
  }

  function readForm() {
    var raw = {
      id: editingId || QS.newId(),
      label: els.fieldLabel.value,
      kind: els.fieldKind.value,
      hours: els.fieldHours.value,
      minutes: els.fieldMinutes.value,
      hour: els.fieldHour.value,
      minute: els.fieldMinute.value
    };
    return QS.normalizePreset(raw);
  }

  function wire() {
    els.clearBtn.addEventListener("click", clearChosen);
    els.addNewBtn.addEventListener("click", function () {
      resetForm();
      els.presetForm.hidden = false;
      els.fieldLabel.focus();
    });
    els.cancelEditBtn.addEventListener("click", resetForm);
    els.fieldKind.addEventListener("change", toggleKindFields);
    els.presetForm.addEventListener("submit", function (event) {
      event.preventDefault();
      var preset = readForm();
      if (!preset) {
        showToast(els.saveToast, "Enter a valid label and delay.");
        return;
      }
      if (editingId) {
        presets = presets.map(function (p) {
          return p.id === editingId ? preset : p;
        });
      } else {
        presets.push(preset);
      }
      resetForm();
      renderQuickButtons();
      renderEditor();
    });
    els.saveAllBtn.addEventListener("click", function () {
      persistPresets(presets, function (err) {
        if (!err) {
          showToast(
            els.saveToast,
            inOffice ? "Saved to your mailbox settings." : "Saved in this browser."
          );
          renderQuickButtons();
          renderEditor();
        }
      });
    });
    els.restoreBtn.addEventListener("click", function () {
      presets = QS.normalizePresets(QS.DEFAULT_PRESETS);
      resetForm();
      renderQuickButtons();
      renderEditor();
    });
  }

  var booted = false;

  function startPreview() {
    if (booted) {
      return;
    }
    booted = true;
    inOffice = false;
    els.banner.dataset.show = "true";
    presets = readPresets();
    els.editorState.textContent = "";
    renderQuickButtons();
    renderEditor();
    refreshCurrent();
    resetForm();
  }

  function startOffice() {
    if (booted && inOffice) {
      return;
    }
    booted = true;
    inOffice = true;
    els.banner.dataset.show = "false";
    try {
      presets = readPresets();
      els.editorState.textContent = "";
      renderQuickButtons();
      renderEditor();
      refreshCurrent();
      resetForm();
    } catch (e) {
      presets = QS.normalizePresets(QS.DEFAULT_PRESETS);
      renderQuickButtons();
      renderEditor();
      resetForm();
      els.editorState.className = "error";
      els.editorState.textContent = "Could not load mailbox settings. Showing defaults. " + e.message;
      setStatus("error", "Could not read this draft’s schedule. You can still choose a delay.");
    }

  }

  wire();

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(function (info) {
      if (info && info.host === Office.HostType.Outlook) {
        startOffice();
      } else {
        startPreview();
      }
    });
    window.setTimeout(function () {
      if (!booted) {
        startPreview();
      }
    }, 4000);
  } else {
    startPreview();
  }
})();
