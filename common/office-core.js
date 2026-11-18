/* Office Suite — shared code used by Letter Generator, Notice and Proceedings.
   Loaded with a plain <script src> (not a module) so the apps still work when
   opened straight from disk (file://). Everything here is global, exactly as it
   was when each app held its own copy. Variables such as lastFocusedField,
   showToast and renderMergeTable are still defined by each app; they are only
   looked up when a function runs, so load order is not a problem.
*/

/* ══════════════════════════════════════
   Dates
   ══════════════════════════════════════ */

function todayFormatted() {
  var d = new Date();
  return (
    String(d.getDate()).padStart(2, "0") +
    "/" +
    String(d.getMonth() + 1).padStart(2, "0") +
    "/" +
    d.getFullYear()
  );
}


/* ══════════════════════════════════════
   Office profile (read from URL parameter or browser storage)
   ══════════════════════════════════════ */

function getOfficeProfile() {
  // 1. Try URL parameter (passed by index.html — works reliably with file:// protocol)
  try {
    var params = new URLSearchParams(window.location.search);
    var encoded = params.get("profile");
    if (encoded) {
      return JSON.parse(decodeURIComponent(escape(atob(encoded))));
    }
  } catch (e) {}
  // 2. Fall back to localStorage
  try {
    var raw = localStorage.getItem("officeProfile");
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}


/* ══════════════════════════════════════
   Preview zoom and fullscreen
   ══════════════════════════════════════ */

var previewZoom = 0.75;

function applyZoom() {
  document.querySelectorAll(".page-container").forEach(function (el) {
    el.style.zoom = previewZoom;
  });
  var el = document.getElementById("zoomLevel");
  if (el) el.textContent = Math.round(previewZoom * 100) + "%";
}

function zoomIn() {
  if (previewZoom < 2.0) {
    previewZoom = Math.round((previewZoom + 0.1) * 10) / 10;
    applyZoom();
  }
}

function zoomOut() {
  if (previewZoom > 0.3) {
    previewZoom = Math.round((previewZoom - 0.1) * 10) / 10;
    applyZoom();
  }
}

function zoomReset() {
  previewZoom = 0.75;
  applyZoom();
}

function toggleFullscreen() {
  var panel = document.getElementById("previewPanel");
  if (!document.fullscreenElement) {
    panel.requestFullscreen && panel.requestFullscreen();
  } else {
    document.exitFullscreen && document.exitFullscreen();
  }
}

function toggleFormFullscreen() {
  var panel = document.getElementById("formPanel");
  if (!document.fullscreenElement) {
    panel.requestFullscreen && panel.requestFullscreen();
  } else {
    document.exitFullscreen && document.exitFullscreen();
  }
}


/* ══════════════════════════════════════
   Form helpers
   ══════════════════════════════════════ */

function autoExpand(el) {
  el.style.height = "auto";
  el.style.height = el.scrollHeight + "px";
}

function closeModal() {
  document.getElementById("modal").style.display = "none";
}

function insertDataSheetValue(text) {
  insertValueAtCursor(lastFocusedField, text);
}


/* ══════════════════════════════════════
   Mail merge (shared state and row navigation)
   ══════════════════════════════════════ */

var mergeState = { headers: [], rows: [], currentIndex: -1 };

function mergeSubstitute(text) {
  if (mergeState.currentIndex < 0 || !mergeState.rows.length) return text;
  var row = mergeState.rows[mergeState.currentIndex];
  mergeState.headers.forEach(function (h) {
    var re = new RegExp(
      "\\{" + h.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\}",
      "gi",
    );
    text = text.replace(re, row[h] || "");
  });
  return text;
}

function escMerge(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function updateMergeIndicator() {
  var el = document.getElementById("mergeRowIndicator");
  if (!el) return;
  el.textContent =
    mergeState.currentIndex < 0
      ? "Showing template"
      : "Row " +
        (mergeState.currentIndex + 1) +
        " of " +
        mergeState.rows.length;
}

function nextMergeRow() {
  if (!mergeState.rows.length) return;
  var next = mergeState.currentIndex + 1;
  if (next >= mergeState.rows.length) next = 0;
  applyMergeRow(next);
}

function prevMergeRow() {
  if (!mergeState.rows.length) return;
  var prev = mergeState.currentIndex - 1;
  if (prev < 0) prev = mergeState.rows.length - 1;
  applyMergeRow(prev);
}


/* ══════════════════════════════════════
   Saved-address picker
   ══════════════════════════════════════ */

/* ── Saved-address picker (tick boxes) ──
   Same code in Letter, Notice and Proceedings. Lists the saved
   addresses (shared browser storage: ml_addresses for To,
   ml_copy_addresses for Copy To) with tick boxes. Addresses already
   in the form start ticked. On Apply, newly ticked ones fill empty
   rows or new rows, and unticked ones are removed (the first row is
   only emptied, since it holds the "+" button). Each app passes its
   own container and add-row function; see pickAddresses(). */
var AddressPicker = (function () {
  var root = null;
  var state = null;

  function readList(key) {
    try {
      var list = JSON.parse(localStorage.getItem(key));
      return Array.isArray(list) ? list : [];
    } catch (e) {
      return [];
    }
  }

  function norm(s) {
    return String(s == null ? "" : s)
      .replace(/\r\n?/g, "\n")
      .trim();
  }

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function rowsOf(container) {
    return Array.prototype.filter.call(container.children, function (r) {
      return !!r.querySelector("textarea");
    });
  }

  function textOf(row) {
    return norm(row.querySelector("textarea").value);
  }

  function build() {
    root = document.createElement("div");
    root.className = "ap-overlay";
    root.innerHTML =
      '<div class="ap-dialog" role="dialog" aria-modal="true" aria-labelledby="apTitle">' +
      '<div class="ap-head"><h3 id="apTitle"></h3>' +
      '<button type="button" class="ap-x" title="Close" aria-label="Close">&times;</button></div>' +
      '<input type="search" class="ap-search" placeholder="Search saved addresses..." aria-label="Search saved addresses" />' +
      '<div class="ap-list" role="group" aria-labelledby="apTitle"></div>' +
      '<div class="ap-foot"><span class="ap-count" role="status"></span>' +
      '<button type="button" class="ap-btn ap-cancel">Cancel</button>' +
      '<button type="button" class="ap-btn ap-apply">Apply</button></div>' +
      "</div>";
    document.body.appendChild(root);
    root.addEventListener("click", function (e) {
      if (e.target === root) close();
    });
    root.querySelector(".ap-x").onclick = close;
    root.querySelector(".ap-cancel").onclick = close;
    root.querySelector(".ap-apply").onclick = apply;
    root.querySelector(".ap-search").oninput = filter;
    root.querySelector(".ap-list").onchange = updateCount;
    root.addEventListener("keydown", function (e) {
      if (e.key === "Escape") close();
    });
  }

  function boxes() {
    return Array.prototype.slice.call(
      root.querySelectorAll('.ap-list input[type="checkbox"]'),
    );
  }

  function updateCount() {
    var n = boxes().filter(function (b) {
      return b.checked;
    }).length;
    root.querySelector(".ap-count").textContent = state.items.length
      ? n + " selected"
      : "";
  }

  function filter() {
    var q = norm(root.querySelector(".ap-search").value).toLowerCase();
    root.querySelectorAll(".ap-item").forEach(function (el) {
      el.hidden = !!q && el.textContent.toLowerCase().indexOf(q) === -1;
    });
  }

  function open(cfg) {
    if (!root) build();
    state = cfg;
    state.items = readList(cfg.storageKey);
    var present = rowsOf(cfg.container).map(textOf);
    state.initial = {};
    var html = "";
    state.items.forEach(function (item, i) {
      var on = present.indexOf(norm(item.address)) !== -1;
      if (on) state.initial[i] = true;
      html +=
        '<label class="ap-item"><input type="checkbox" value="' +
        i +
        '"' +
        (on ? " checked" : "") +
        " /><span><strong>" +
        esc(item.title) +
        '</strong><span class="ap-addr">' +
        esc(item.address) +
        "</span></span></label>";
    });
    if (!html) {
      html =
        '<p class="ap-empty">No saved addresses yet. Use the ⚙️ button next to the heading to save some.</p>';
    }
    root.querySelector("#apTitle").textContent = cfg.title;
    root.querySelector(".ap-list").innerHTML = html;
    var search = root.querySelector(".ap-search");
    search.value = "";
    search.hidden = state.items.length < 6;
    root.querySelector(".ap-apply").disabled = !state.items.length;
    updateCount();
    root.classList.add("open");
    (search.hidden
      ? root.querySelector(".ap-list input") || root.querySelector(".ap-x")
      : search
    ).focus();
  }

  function close() {
    if (root) root.classList.remove("open");
  }

  function clearOrRemove(row, container) {
    if (rowsOf(container)[0] === row) {
      row.querySelector("textarea").value = "";
      var sel = row.querySelector("select");
      if (sel) sel.value = "";
      if (typeof autoExpand === "function")
        autoExpand(row.querySelector("textarea"));
    } else {
      row.remove();
    }
  }

  function fill(row, item) {
    var text = row.querySelector("textarea");
    text.value = item.address;
    var sel = row.querySelector("select");
    if (sel) {
      var match = Array.prototype.find.call(sel.options, function (o) {
        return o.value !== "" && o.textContent === item.title;
      });
      sel.value = match ? match.value : "";
    }
    if (typeof autoExpand === "function") autoExpand(text);
  }

  function apply() {
    var cfg = state;
    var c = cfg.container;
    var checked = {};
    boxes().forEach(function (b) {
      if (b.checked) checked[b.value] = true;
    });
    var added = 0;
    var removed = 0;

    // Unticked: take out the rows that hold exactly that address.
    Object.keys(cfg.initial).forEach(function (i) {
      if (checked[i]) return;
      var addr = norm(cfg.items[i].address);
      rowsOf(c).forEach(function (row) {
        if (textOf(row) === addr) {
          clearOrRemove(row, c);
          removed++;
        }
      });
    });

    // Ticked: fill an empty row, or add one.
    Object.keys(checked).forEach(function (i) {
      var item = cfg.items[i];
      var addr = norm(item.address);
      var rows = rowsOf(c);
      if (rows.some(function (r) { return textOf(r) === addr; })) return;
      var target = rows.filter(function (r) { return !textOf(r); })[0];
      if (!target) {
        cfg.addRow();
        rows = rowsOf(c);
        target = rows[rows.length - 1];
      }
      if (!target) return;
      fill(target, item);
      added++;
    });

    close();
    cfg.refresh();
    if (cfg.toast && (added || removed)) {
      var parts = [];
      if (added) parts.push(added + " added");
      if (removed) parts.push(removed + " removed");
      cfg.toast("Addresses: " + parts.join(", "));
    }
  }

  return { open: open, close: close };
})();

