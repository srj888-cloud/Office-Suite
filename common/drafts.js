/* Office Suite - drafts saved in this browser.

   Drafts are kept in the browser's IndexedDB database "officeSuite"
   (store "drafts"), so no export/import is needed day to day. A draft is
   the same data the app's own JSON export holds, plus a title and times:
     { id, app: "letter" | "notice" | "proceedings", title,
       created, updated, data }

   Storage belongs to this browser on this computer (and to the address the
   app is opened from). Clearing browsing data removes it, so the Drafts
   window offers "Back up all drafts" (one JSON file) and "Restore backup".

   Used by the three apps (auto-save + Drafts window) and by index.html
   (Recent drafts). An app takes part by defining, in its APP settings:
     draftApp             "letter" / "notice" / "proceedings"
     getDraftData()       -> object (what the JSON export holds)
     applyDraftData(data)    fills the form from that object
     draftTitle(data)     -> short title for the list
*/
var OfficeDrafts = (function () {
  var DB_NAME = "officeSuite";
  var STORE = "drafts";
  var APP_FILES = {
    letter: "Letter Generator.html",
    notice: "Notice.html",
    proceedings: "Proceedings.html",
  };
  var APP_NAMES = {
    letter: "Letter",
    notice: "Circular / Notice",
    proceedings: "Proceedings",
  };
  var APP_ICONS = { letter: "✉️", notice: "📢", proceedings: "⚖️" };

  /* ── storage ── */
  var dbPromise = null;
  function openDb() {
    if (!dbPromise) {
      dbPromise = new Promise(function (resolve, reject) {
        if (!window.indexedDB) {
          reject(new Error("This browser cannot save drafts (no IndexedDB)."));
          return;
        }
        var req = indexedDB.open(DB_NAME, 1);
        req.onupgradeneeded = function () {
          var store = req.result.createObjectStore(STORE, { keyPath: "id" });
          store.createIndex("updated", "updated");
          store.createIndex("app", "app");
        };
        req.onsuccess = function () {
          resolve(req.result);
        };
        req.onerror = function () {
          reject(req.error);
        };
      });
      dbPromise.catch(function () {
        dbPromise = null; // allow a retry later
      });
    }
    return dbPromise;
  }

  function run(mode, work) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(STORE, mode);
        var req = work(tx.objectStore(STORE));
        tx.oncomplete = function () {
          resolve(req ? req.result : undefined);
        };
        tx.onerror = tx.onabort = function () {
          reject(tx.error || new Error("Draft storage error"));
        };
      });
    });
  }

  function get(id) {
    return run("readonly", function (s) {
      return s.get(id);
    });
  }
  function put(draft) {
    return run("readwrite", function (s) {
      return s.put(draft);
    });
  }
  function remove(id) {
    return run("readwrite", function (s) {
      return s.delete(id);
    });
  }
  // newest first; app is optional
  function list(app) {
    return run("readonly", function (s) {
      return s.getAll();
    }).then(function (all) {
      return (all || [])
        .filter(function (d) {
          return !app || d.app === app;
        })
        .sort(function (a, b) {
          return (b.updated || 0) - (a.updated || 0);
        });
    });
  }

  function newId() {
    return (
      Date.now().toString(36) + Math.random().toString(36).slice(2, 8)
    );
  }

  // Ask the browser not to clear our storage when space runs low
  var persistAsked = false;
  function askPersist() {
    if (persistAsked) return;
    persistAsked = true;
    try {
      if (navigator.storage && navigator.storage.persist) {
        navigator.storage.persist();
      }
    } catch (e) {}
  }

  /* ── backup ── */
  function exportAll() {
    // resolves to { drafts: n, phrases: n } (zeros when nothing was saved)
    return list().then(function (drafts) {
      var hasPhrases = window.OfficePhrases && OfficePhrases.isCustom();
      if (!drafts.length && !hasPhrases) return { drafts: 0, phrases: 0 };
      var out = {
        type: "office-suite-drafts",
        version: 1,
        exported: new Date().toISOString(),
        drafts: drafts,
      };
      // the phrase list travels with the drafts (only when it was changed)
      if (window.OfficePhrases && OfficePhrases.isCustom()) {
        out.phrases = OfficePhrases.getList();
      }
      var blob = new Blob([JSON.stringify(out, null, 2)], {
        type: "application/json",
      });
      var url = URL.createObjectURL(blob);
      var a = document.createElement("a");
      a.href = url;
      a.download =
        "office_suite_drafts_" + new Date().toISOString().slice(0, 10) + ".json";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      return { drafts: drafts.length, phrases: out.phrases ? out.phrases.length : 0 };
    });
  }

  // Adds drafts from a backup file. A draft already here is replaced only
  // when the backup copy is newer. Resolves to {added, updated, skipped}.
  function importFile(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () {
        var json;
        try {
          json = JSON.parse(reader.result);
        } catch (e) {
          reject(new Error("Could not read the file. Choose a drafts backup file."));
          return;
        }
        if (!json || json.type !== "office-suite-drafts" || !Array.isArray(json.drafts)) {
          reject(new Error("This is not an Office Suite drafts backup file."));
          return;
        }
        var counts = { added: 0, updated: 0, skipped: 0, phrases: 0 };
        if (Array.isArray(json.phrases) && window.OfficePhrases) {
          try {
            counts.phrases = OfficePhrases.mergeList(json.phrases);
          } catch (e) {}
        }
        var valid = json.drafts.filter(function (d) {
          return d && d.id && APP_FILES[d.app] && d.data;
        });
        Promise.all(
          valid.map(function (d) {
            return get(d.id).then(function (have) {
              if (have && (have.updated || 0) >= (d.updated || 0)) {
                counts.skipped++;
                return;
              }
              counts[have ? "updated" : "added"]++;
              return put(d);
            });
          }),
        ).then(function () {
          askPersist();
          resolve(counts);
        }, reject);
      };
      reader.onerror = function () {
        reject(new Error("Could not read the file."));
      };
      reader.readAsText(file);
    });
  }

  /* ── helpers for lists ── */
  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }
  function when(t) {
    if (!t) return "";
    var d = new Date(t);
    var now = new Date();
    var time = d.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" });
    if (d.toDateString() === now.toDateString()) return "Today " + time;
    var y = new Date(now);
    y.setDate(now.getDate() - 1);
    if (d.toDateString() === y.toDateString()) return "Yesterday " + time;
    return (
      d.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" }) +
      " " +
      time
    );
  }

  /* ── inside an app: auto-save and the Drafts window ── */
  var current = null; // the draft being edited
  var lastJson = null; // what was last saved (or loaded)
  var timer = null;
  var saving = null;

  function setStatus(text, kind) {
    var el = document.getElementById("draftStatus");
    if (!el) return;
    el.textContent = text;
    el.className = "draft-status" + (kind ? " " + kind : "");
  }

  function urlWithDraft(id) {
    var params = new URLSearchParams(location.search);
    if (id) params.set("draft", id);
    else params.delete("draft");
    var q = params.toString();
    return location.pathname + (q ? "?" + q : "");
  }

  // Save now if anything changed. Resolves when stored.
  function saveNow() {
    clearTimeout(timer);
    timer = null;
    if (saving) return saving.then(saveNow);
    var data;
    try {
      data = APP.getDraftData();
    } catch (e) {
      return Promise.resolve();
    }
    var json = JSON.stringify(data);
    if (json === lastJson) return Promise.resolve();
    var now = Date.now();
    var draft = current
      ? Object.assign({}, current)
      : { id: newId(), app: APP.draftApp, created: now };
    draft.title = (APP.draftTitle(data) || "").trim() || "Untitled";
    draft.updated = now;
    draft.data = data;
    setStatus("Saving…");
    saving = put(draft).then(
      function () {
        var isNew = !current;
        current = draft;
        lastJson = json;
        saving = null;
        setStatus("Saved " + new Date(now).toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" }), "ok");
        if (isNew) {
          askPersist();
          // a reload now reopens this draft
          try {
            history.replaceState(history.state, "", urlWithDraft(draft.id));
          } catch (e) {}
        }
      },
      function (err) {
        saving = null;
        setStatus("Not saved", "error");
        showToast("Could not save the draft in this browser. " + (err && err.message ? err.message : ""), "error");
      },
    );
    return saving;
  }

  // Any edit or click (form fields, add/remove buttons, address picker,
  // table tools) schedules a check; saveNow() only writes when the form
  // data actually changed. The Drafts bar and window are ignored.
  function changed(e) {
    var t = e.target;
    if (!t || !t.closest) return;
    if (t.closest(".draft-bar, .draft-actions, .draft-list, .draft-backup")) return;
    if (e.type === "input" && t.closest("#formPanel")) setStatus("Editing…");
    clearTimeout(timer);
    timer = setTimeout(saveNow, 1500);
  }

  // After "Clear form": keep the saved draft as it was; further typing
  // starts a new draft instead of overwriting it with an empty form.
  function detach() {
    if (typeof APP === "undefined" || !APP.draftApp) return;
    clearTimeout(timer);
    current = null;
    try {
      lastJson = JSON.stringify(APP.getDraftData());
    } catch (e) {}
    setStatus("New draft");
    try {
      history.replaceState(history.state, "", urlWithDraft(null));
    } catch (e) {}
  }

  function goTo(url) {
    saveNow().then(function () {
      location.href = url;
    });
  }

  function openDraft(id) {
    goTo(urlWithDraft(id));
  }

  function newDraft() {
    goTo(urlWithDraft(null));
  }

  function deleteDraft(id) {
    get(id).then(function (d) {
      if (!d) return;
      if (!confirm('Delete the draft "' + d.title + '"? This cannot be undone.')) return;
      remove(id).then(function () {
        if (current && current.id === id) {
          current = null;
          lastJson = JSON.stringify(APP.getDraftData()); // don't re-save it straight away
          setStatus("Draft deleted");
          try {
            history.replaceState(history.state, "", urlWithDraft(null));
          } catch (e) {}
        }
        showToast("Draft deleted");
        showWindow();
      });
    });
  }

  function showWindow() {
    saveNow().then(function () {
      return list(APP.draftApp);
    }).then(
      function (drafts) {
        var html =
          '<div class="draft-actions">' +
          '<button type="button" class="btn btn-small" onclick="OfficeDrafts.newDraft()">+ New draft</button>' +
          '<span class="draft-hint">Drafts are saved automatically in this browser as you type.</span>' +
          "</div>";
        if (!drafts.length) {
          html += '<p class="draft-empty">No saved drafts yet. Start typing in the form and it is saved here.</p>';
        } else {
          html += '<ul class="draft-list">';
          drafts.forEach(function (d) {
            var isOpen = current && current.id === d.id;
            html +=
              '<li class="draft-item' + (isOpen ? " current" : "") + '">' +
              '<div class="draft-info"><strong>' + esc(d.title) + "</strong>" +
              "<small>" + esc(when(d.updated)) + (isOpen ? " · open now" : "") + "</small></div>" +
              (isOpen
                ? ""
                : '<button type="button" class="btn btn-primary btn-small" onclick="OfficeDrafts.openDraft(\'' + esc(d.id) + "')\">Open</button>") +
              '<button type="button" class="btn btn-danger btn-small" onclick="OfficeDrafts.deleteDraft(\'' + esc(d.id) + "')\">Delete</button>" +
              "</li>";
          });
          html += "</ul>";
        }
        html +=
          '<div class="draft-backup">' +
          '<button type="button" class="btn btn-small" onclick="OfficeDrafts.backup()">Back up all drafts</button>' +
          '<button type="button" class="btn btn-small" onclick="document.getElementById(\'draftRestoreInput\').click()">Restore backup</button>' +
          '<input type="file" id="draftRestoreInput" accept=".json,application/json" hidden onchange="OfficeDrafts.restore(this)">' +
          '<span class="draft-hint">One file with the drafts of all three apps. Keep a copy before clearing browser data or changing computers.</span>' +
          "</div>";
        showModal("My Drafts · " + (APP_NAMES[APP.draftApp] || ""), html);
      },
      function (err) {
        showToast("Could not open saved drafts. " + (err && err.message ? err.message : ""), "error");
      },
    );
  }

  function backup() {
    saveNow()
      .then(exportAll)
      .then(
        function (c) {
          if (!c.drafts && !c.phrases) {
            showToast("There are no drafts to back up yet.", "error");
            return;
          }
          showToast("Backed up " + c.drafts + " draft(s)" + (c.phrases ? " and " + c.phrases + " phrase(s)" : "") + ".");
        },
        function (err) {
          showToast("Could not back up drafts. " + err.message, "error");
        },
      );
  }

  function restore(input) {
    var file = input.files && input.files[0];
    input.value = "";
    if (!file) return;
    importFile(file).then(
      function (c) {
        showToast(
          "Restored: " + c.added + " new, " + c.updated + " updated" +
            (c.skipped ? ", " + c.skipped + " already up to date" : "") +
            (c.phrases ? ", " + c.phrases + " phrase(s) added" : "") + ".",
        );
        if (document.getElementById("draftRestoreInput")) showWindow();
      },
      function (err) {
        showToast(err.message, "error");
      },
    );
  }

  function addBar() {
    var header = document.querySelector("#formPanel .panel-header");
    if (!header || document.getElementById("draftStatus")) return;
    var bar = document.createElement("span");
    bar.className = "draft-bar";
    bar.innerHTML =
      '<span id="draftStatus" class="draft-status"></span>' +
      '<button type="button" class="draft-btn" onclick="OfficeDrafts.showWindow()" title="Saved drafts in this browser">📂 Drafts</button>';
    var fsBtn = document.getElementById("formFsBtn");
    header.insertBefore(bar, fsBtn || null);
  }

  function initApp() {
    if (typeof APP === "undefined" || !APP.draftApp) return;
    addBar();
    var id = new URLSearchParams(location.search).get("draft");
    var ready = id
      ? get(id).then(
          function (d) {
            if (!d || d.app !== APP.draftApp) {
              showToast("That draft was not found in this browser.", "error");
              try {
                history.replaceState(history.state, "", urlWithDraft(null));
              } catch (e) {}
              return;
            }
            APP.applyDraftData(d.data);
            current = d;
            setStatus("Opened " + when(d.updated), "ok");
          },
          function (err) {
            setStatus("Drafts unavailable", "error");
            showToast("Saved drafts are not available in this browser. " + (err && err.message ? err.message : ""), "error");
          },
        )
      : Promise.resolve();
    ready.then(function () {
      try {
        lastJson = JSON.stringify(APP.getDraftData()); // starting point: no save until something changes
      } catch (e) {}
      ["input", "change", "click"].forEach(function (type) {
        document.addEventListener(type, changed, true);
      });
      document.addEventListener("visibilitychange", function () {
        if (document.visibilityState === "hidden") saveNow();
      });
      window.addEventListener("pagehide", saveNow);
    });
  }

  // run after each app's own start-up code (window.onload / DOMContentLoaded)
  if (typeof window !== "undefined") {
    window.addEventListener("load", function () {
      setTimeout(initApp, 0);
    });
  }

  return {
    APP_FILES: APP_FILES,
    APP_NAMES: APP_NAMES,
    APP_ICONS: APP_ICONS,
    list: list,
    get: get,
    remove: remove,
    exportAll: exportAll,
    importFile: importFile,
    esc: esc,
    when: when,
    saveNow: saveNow,
    detach: detach,
    showWindow: showWindow,
    openDraft: openDraft,
    newDraft: newDraft,
    deleteDraft: deleteDraft,
    backup: backup,
    restore: restore,
  };
})();
