/* Office Suite - phrase suggestions.

   While typing in a form field (subject, body, references, Copy To ...),
   the first few letters of a saved phrase bring up a short list under the
   cursor:
     Tab            insert the highlighted phrase
     Down / Up      choose another, then Enter (or Tab) to insert
     Esc            close the list (plain typing and Enter work as usual)
     Ctrl+Space     browse all phrases
   A phrase may contain ___ blanks: after inserting, the first blank is
   selected; Tab moves to the next one.
   Select text in a field and click "Save as phrase" to add your own.

   Matching works on Malayalam typed with InScript or any other method:
   chillu letters (ൽ / ല്‍ etc.) and the invisible joiners are treated the
   same, so old and new spellings match each other.

   Phrases live in browser storage ("ml_phrases"; the built-in starter list
   is used until the list is first changed). Usage counts ("ml_phrase_usage")
   push often-used phrases up. The home page Settings has a Phrases tab to
   manage the list (renderManager).
*/
var OfficePhrases = (function () {
  var KEY = "ml_phrases";
  var USAGE_KEY = "ml_phrase_usage";
  var ENABLED_KEY = "ml_phrases_enabled";
  var MIN_CHARS = 3;
  var MAX_SHOWN = 6;
  var APP_NAMES = { letter: "Letter", notice: "Notice", proceedings: "Proceedings" };
  var APP_CHOICES = ["All", "Letter", "Notice", "Proceedings"];
  // fields that already have their own helpers (address book) or are numbers
  var EXCLUDE =
    "#toContainer *, #toAddresses *, .address-textarea, .to-input, #refNumber, #orderNumber, " +
    "#principalName, #fromAddress, #appellateAuthorityText, #dataSheetSearch, #dataSheetGrid *";

  // [category, use in, Malayalam, English]
  var STARTER = [
    ["Subject", "All", "___ - അനുമതി നൽകുന്നത് - സംബന്ധിച്ച്.", "___ - Sanction accorded - reg."],
    ["Subject", "All", "___ - വിവരങ്ങൾ ലഭ്യമാക്കുന്നത് - സംബന്ധിച്ച്.", "___ - Furnishing of details - reg."],
    ["Subject", "Letter", "___ - അയച്ചുതരുന്നത് - സംബന്ധിച്ച്.", "___ - Forwarding of ___ - reg."],
    ["Subject", "Letter", "___ - ഓർമ്മക്കുറിപ്പ്.", "___ - Reminder."],
    ["Reference", "All", "സർക്കാർ ഉത്തരവ് (സാധാ) നം. ___/___/ഉ.വി.വ. തീയതി ___.", "G.O. (Rt) No. ___/___/H.Edn. dated ___."],
    ["Reference", "All", "സർക്കാർ ഉത്തരവ് (കൈ) നം. ___/___/ഉ.വി.വ. തീയതി ___.", "G.O. (MS) No. ___/___/H.Edn. dated ___."],
    ["Reference", "All", "___ ന്റെ ___ തീയതിയിലെ ___ നമ്പർ കത്ത്.", "Letter No. ___ dated ___ from ___."],
    ["Reference", "All", "കോളേജ് വിദ്യാഭ്യാസ ഡയറക്ടറുടെ ___ തീയതിയിലെ ___ നമ്പർ സർക്കുലർ.", "Circular No. ___ dated ___ of the Director of Collegiate Education."],
    ["Reference", "All", "___ ന്റെ ___ തീയതിയിലെ അപേക്ഷ.", "Application dated ___ from ___."],
    ["Reference", "All", "ഈ ഓഫീസിലെ ___ തീയതിയിലെ ___ നമ്പർ കത്ത്.", "This office letter No. ___ dated ___."],
    ["Opening", "All", "മേൽ പരാമർശ പ്രകാരം,", "With reference to the above,"],
    ["Opening", "All", "പരാമർശം (1) പ്രകാരം ___", "As per reference (1) cited, ___"],
    ["Opening", "Letter", "മേൽ വിഷയത്തിലേക്ക് ശ്രദ്ധ ക്ഷണിക്കുന്നു.", "Kind attention is invited to the subject cited above."],
    ["Opening", "Letter", "പരാമർശിത കത്ത് പ്രകാരം ആവശ്യപ്പെട്ട വിവരങ്ങൾ ചുവടെ ചേർക്കുന്നു.", "The details called for in the letter cited are furnished below."],
    ["Opening", "All", "പരാമർശിത സർക്കുലറിലെ നിർദ്ദേശ പ്രകാരം", "As directed in the circular cited,"],
    ["Opening", "Proceedings", "മേൽ സാഹചര്യത്തിൽ", "In the above circumstances,"],
    ["Opening", "Proceedings", "വിഷയം വിശദമായി പരിശോധിച്ചതിന്റെ അടിസ്ഥാനത്തിൽ", "On a detailed examination of the matter,"],
    ["Action", "Letter", "___ ഇതോടൊപ്പം അയക്കുന്നു.", "___ is forwarded herewith."],
    ["Action", "Letter", "തുടർനടപടികൾ സ്വീകരിക്കുന്നതിനായി ഇതോടൊപ്പം അയക്കുന്നു.", "Forwarded herewith for further necessary action."],
    ["Action", "All", "ആവശ്യമായ നടപടികൾ സ്വീകരിക്കണമെന്ന് അഭ്യർത്ഥിക്കുന്നു.", "It is requested to take necessary action."],
    ["Action", "All", "ഇക്കാര്യത്തിൽ അടിയന്തര നടപടി സ്വീകരിക്കണമെന്ന് അഭ്യർത്ഥിക്കുന്നു.", "It is requested to take urgent action in this matter."],
    ["Action", "Letter", "ആയതിന് അനുമതി നൽകണമെന്ന് അഭ്യർത്ഥിക്കുന്നു.", "It is requested that sanction may kindly be accorded for the same."],
    ["Action", "Letter", "ആവശ്യമായ തുക അനുവദിക്കണമെന്ന് അഭ്യർത്ഥിക്കുന്നു.", "It is requested to allot the required funds."],
    ["Action", "All", "അറിവിലേക്കും തുടർനടപടികൾക്കുമായി", "For information and necessary action."],
    ["Action", "All", "വിശദമായ റിപ്പോർട്ട് ___ നകം ലഭ്യമാക്കേണ്ടതാണ്.", "A detailed report should be furnished on or before ___."],
    ["Action", "All", "വിവരങ്ങൾ ചുവടെ ചേർക്കുന്നു.", "The details are given below."],
    ["Action", "All", "വിശദാംശങ്ങൾ അനുബന്ധമായി ചേർക്കുന്നു.", "The details are appended."],
    ["Action", "Letter", "ഇക്കാര്യത്തിൽ ഈ ഓഫീസിന് യാതൊരു വിരോധവുമില്ല.", "This office has no objection in this regard."],
    ["Action", "All", "മേൽ വിവരങ്ങൾ ശരിയാണെന്ന് സാക്ഷ്യപ്പെടുത്തുന്നു.", "Certified that the above details are correct."],
    ["Action", "Letter", "ഇക്കാര്യം സംബന്ധിച്ച മറുപടി ഇതുവരെയും ലഭ്യമായിട്ടില്ല.", "No reply has been received in this regard so far."],
    ["Action", "Letter", "ആയതിനാൽ മറുപടി അടിയന്തരമായി ലഭ്യമാക്കണമെന്ന് വീണ്ടും അഭ്യർത്ഥിക്കുന്നു.", "Hence it is requested once again to furnish the reply urgently."],
    ["Closing", "Letter", "സർ/മാഡം,", "Sir/Madam,"],
    ["Closing", "Letter", "വിശ്വസ്തതയോടെ,", "Yours faithfully,"],
    ["Closing", "Letter", "ഉള്ളടക്കം: ___", "Encl: ___"],
    ["Copy To", "All", "ഓഫീസ് ഫയലിലേക്ക്", "Office copy"],
    ["Copy To", "All", "കരുതൽ ഫയൽ", "Stock file"],
    ["Copy To", "All", "നോട്ടീസ് ബോർഡ്", "Notice Board"],
    ["Copy To", "All", "എല്ലാ വകുപ്പ് മേധാവികൾക്കും", "All Heads of Departments"],
    ["Copy To", "All", "കോളേജ് വെബ്സൈറ്റിൽ പ്രസിദ്ധീകരിക്കുന്നതിന്", "For publishing on the college website"],
    ["Notice", "Notice", "എല്ലാ അദ്ധ്യാപക/അനദ്ധ്യാപക ജീവനക്കാരുടെയും ശ്രദ്ധയ്ക്ക്", "Attention of all teaching and non-teaching staff"],
    ["Notice", "Notice", "എല്ലാ വിദ്യാർത്ഥികളുടെയും ശ്രദ്ധയ്ക്ക്", "Attention of all students"],
    ["Notice", "Notice", "___ തീയതി ___ മണിക്ക് ___ ൽ വെച്ച് ___ നടത്തുന്നതാണ്.", "___ will be held at ___ on ___ at ___."],
    ["Notice", "Notice", "എല്ലാവരും കൃത്യസമയത്ത് പങ്കെടുക്കണമെന്ന് അറിയിക്കുന്നു.", "All are requested to attend on time."],
    ["Notice", "Notice", "___ തീയതി കോളേജിന് അവധിയായിരിക്കും.", "The college will remain closed on ___."],
    ["Notice", "Notice", "___ തീയതി ___ മണിക്ക് സ്റ്റാഫ് കൗൺസിൽ യോഗം ചേരുന്നതാണ്.", "A Staff Council meeting will be held on ___ at ___."],
    ["Notice", "Notice", "ഈ നിർദ്ദേശം കർശനമായി പാലിക്കേണ്ടതാണ്.", "This direction shall be strictly followed."],
    ["Notice", "Notice", "ഈ സർക്കുലർ എല്ലാ ക്ലാസുകളിലും വായിച്ചു കേൾപ്പിക്കേണ്ടതാണ്.", "This circular shall be read out in all classes."],
    ["Proceedings", "Proceedings", "ഇതിനാൽ ഉത്തരവാകുന്നു.", "Orders are issued accordingly."],
    ["Proceedings", "Proceedings", "___ ന് ___ രൂപ അനുവദിച്ച് ഇതിനാൽ ഉത്തരവാകുന്നു.", "Sanction is hereby accorded for ___ for an amount of Rs. ___."],
    ["Proceedings", "Proceedings", "ഇതിനുള്ള ചെലവ് ___ ശീർഷകത്തിൽ നിന്നും വഹിക്കേണ്ടതാണ്.", "The expenditure shall be met from the head of account ___."],
    ["Proceedings", "Proceedings", "___ ന് ___ മുതൽ ___ വരെ ___ ദിവസത്തെ ___ അവധി അനുവദിച്ച് ഉത്തരവാകുന്നു.", "___ is granted ___ days of ___ leave from ___ to ___."],
    ["Proceedings", "Proceedings", "___ നെ ___ തീയതി ഉച്ചയ്ക്ക് ശേഷം ഈ ഓഫീസിൽ നിന്നും വിടുതൽ ചെയ്തിരിക്കുന്നു.", "___ is relieved from this office on the afternoon of ___."],
    ["Proceedings", "Proceedings", "___ നെ ___ ചുമതല ഏൽപ്പിച്ച് ഉത്തരവാകുന്നു.", "___ is entrusted with the charge of ___."],
    ["Proceedings", "Proceedings", "താഴെപ്പറയുന്നവരെ ഉൾപ്പെടുത്തി ___ കമ്മിറ്റി രൂപീകരിച്ച് ഉത്തരവാകുന്നു.", "A ___ committee is constituted with the following members."],
    ["RTI", "Letter", "വിവരാവകാശ നിയമം 2005 പ്രകാരം ___ തീയതിയിൽ ലഭിച്ച അപേക്ഷയിലെ വിവരങ്ങൾ ചുവടെ ചേർക്കുന്നു.", "The information sought in your application dated ___ under the RTI Act, 2005 is furnished below."],
    ["RTI", "Letter", "ആവശ്യപ്പെട്ട വിവരം ഈ ഓഫീസിൽ ലഭ്യമല്ല.", "The information sought is not available in this office."],
    ["RTI", "Letter", "അപേക്ഷ വിവരാവകാശ നിയമം 2005 വകുപ്പ് 6(3) പ്രകാരം ___ ലേക്ക് കൈമാറിയിരിക്കുന്നു.", "The application has been transferred to ___ under Section 6(3) of the RTI Act, 2005."],
    ["RTI", "Letter", "ഈ മറുപടിയിൽ തൃപ്തനല്ലെങ്കിൽ 30 ദിവസത്തിനകം താഴെ പറയുന്ന അപ്പീൽ അധികാരിക്ക് അപ്പീൽ സമർപ്പിക്കാവുന്നതാണ്.", "If not satisfied with this reply, you may file an appeal within 30 days before the Appellate Authority given below."],
  ];

  /* ── text normalising (for matching only) ── */
  var CHILLU = {
    "ൺ": "ണ്", // ൺ
    "ൻ": "ന്", // ൻ
    "ർ": "ര്", // ർ
    "ൽ": "ല്", // ൽ
    "ൾ": "ള്", // ൾ
    "ൿ": "ക്", // ൿ
    "ൔ": "മ്", // ൔ
    "ൕ": "യ്", // ൕ
    "ൖ": "ഴ്", // ൖ
  };
  var CHILLU_RE = /[ൺ-ൿൔ-ൖ]/g;

  function norm(s) {
    s = String(s || "");
    try {
      s = s.normalize("NFC");
    } catch (e) {}
    return s
      .replace(CHILLU_RE, function (c) {
        return CHILLU[c];
      })
      .replace(/[​-‍⁠﻿]/g, "")
      .replace(/\s+/g, " ")
      .toLowerCase();
  }

  // atomic chillu -> consonant + virama + ZWJ (the older form many keyboards type)
  function toOldChillu(s) {
    return s.replace(CHILLU_RE, function (c) {
      return CHILLU[c] + "‍";
    });
  }

  function hasMalayalam(s) {
    return /[ഀ-ൿ]/.test(s);
  }

  /* ── storage ── */
  var cache = null;
  function starterList() {
    return STARTER.map(function (r, i) {
      return { id: "s" + (i + 1), category: r[0], apps: r[1], ml: r[2], en: r[3] };
    });
  }
  function getList() {
    if (cache) return cache;
    var list = null;
    try {
      list = JSON.parse(localStorage.getItem(KEY));
    } catch (e) {}
    cache = Array.isArray(list) ? list : starterList();
    return cache;
  }
  function isCustom() {
    try {
      return !!localStorage.getItem(KEY);
    } catch (e) {
      return false;
    }
  }
  function saveList(list) {
    localStorage.setItem(KEY, JSON.stringify(list));
    cache = null;
  }
  function resetList() {
    localStorage.removeItem(KEY);
    cache = null;
  }
  function getUsage() {
    try {
      return JSON.parse(localStorage.getItem(USAGE_KEY)) || {};
    } catch (e) {
      return {};
    }
  }
  function countUse(id) {
    var u = getUsage();
    u[id] = (u[id] || 0) + 1;
    try {
      localStorage.setItem(USAGE_KEY, JSON.stringify(u));
    } catch (e) {}
  }
  function isEnabled() {
    try {
      return localStorage.getItem(ENABLED_KEY) !== "0";
    } catch (e) {
      return true;
    }
  }
  function setEnabled(on) {
    try {
      localStorage.setItem(ENABLED_KEY, on ? "1" : "0");
    } catch (e) {}
  }
  function newId() {
    return "u" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  }
  function sameText(a, b) {
    return !!a && !!b && norm(a) === norm(b);
  }
  // Adds phrases not already in the list (same Malayalam or same English).
  function mergeList(incoming) {
    var list = getList().slice();
    var added = 0;
    (incoming || []).forEach(function (p) {
      if (!p || (!p.ml && !p.en)) return;
      var dup = list.some(function (q) {
        return sameText(q.ml, p.ml) || (!p.ml && sameText(q.en, p.en));
      });
      if (dup) return;
      list.push({
        id: p.id && !list.some(function (q) { return q.id === p.id; }) ? p.id : newId(),
        category: p.category || "My phrases",
        apps: p.apps || "All",
        ml: p.ml || "",
        en: p.en || "",
      });
      added++;
    });
    if (added) saveList(list);
    return added;
  }

  if (typeof window !== "undefined") {
    window.addEventListener("storage", function (e) {
      if (e.key === KEY || e.key === null) cache = null;
    });
  }

  /* ── matching ── */
  function appName() {
    return typeof APP !== "undefined" && APP.draftApp ? APP_NAMES[APP.draftApp] : null;
  }
  function forThisApp(p) {
    var name = appName();
    if (!name || !p.apps || p.apps === "All") return true;
    return String(p.apps).split(/[,/]/).some(function (a) {
      return a.trim().toLowerCase() === name.toLowerCase();
    });
  }
  function wordStarts(n) {
    var out = [0];
    for (var i = 1; i < n.length; i++) {
      if (/[\s(\-–—"'“‘/]/.test(n[i - 1]) && !/\s/.test(n[i])) out.push(i);
    }
    return out;
  }

  // sentence: text typed since the last . ! ? ; : or new line
  // word: text typed since the last space
  function findMatches(sentence, word) {
    var ns = norm(sentence).replace(/^ /, "");
    var nw = norm(word);
    var useSentence = ns.length >= MIN_CHARS && ns.length <= 120;
    var useWord = nw.length >= MIN_CHARS;
    if (!useSentence && !useWord) return [];
    var usage = getUsage();
    var out = [];
    getList().forEach(function (p) {
      if (!forThisApp(p)) return;
      ["ml", "en"].forEach(function (lang) {
        var text = p[lang];
        if (!text) return;
        var n = norm(text);
        var best = null;
        if (useSentence && n.indexOf(ns) === 0 && n !== ns && ns.length > nw.length) {
          best = { score: 3, replace: "sentence" };
        } else if (useWord) {
          var starts = wordStarts(n);
          for (var k = 0; k < starts.length; k++) {
            if (n.substr(starts[k], nw.length) === nw && n !== nw) {
              best = { score: starts[k] === 0 ? 2.5 : 2, replace: "word" };
              break;
            }
          }
        }
        if (best) {
          out.push({ p: p, lang: lang, text: text, score: best.score, replace: best.replace, uses: usage[p.id] || 0 });
        }
      });
    });
    // where the match is (start of phrase > start of a later word) plus how
    // often the phrase was used: a phrase used a few times rises to the top
    function rank(m) {
      return m.score + 0.4 * Math.log2(1 + m.uses);
    }
    out.sort(function (a, b) {
      return rank(b) - rank(a) || a.text.length - b.text.length;
    });
    // one entry per phrase
    var seen = {};
    return out.filter(function (m) {
      if (seen[m.p.id]) return false;
      seen[m.p.id] = true;
      return true;
    }).slice(0, MAX_SHOWN);
  }

  /* ── styles (injected once, used by apps and the home page) ── */
  var CSS =
    ".ph-pop{position:fixed;z-index:10001;min-width:260px;max-width:min(560px,92vw);background:#fff;border:1px solid #cfd8dc;border-radius:10px;box-shadow:0 8px 24px rgba(0,0,0,.18);padding:4px;font-size:14px;line-height:1.45;color:#2c3e50}" +
    ".ph-item{display:flex;gap:10px;align-items:baseline;padding:7px 10px;border-radius:7px;cursor:pointer}" +
    ".ph-item.active{background:#e8f1fb}" +
    ".ph-text{flex:1;white-space:normal}" +
    ".ph-cat{font-size:11px;color:#8a99a6;white-space:nowrap}" +
    ".ph-hint{font-size:11px;color:#8a99a6;padding:4px 10px 3px;border-top:1px solid #eef2f4;margin-top:2px}" +
    ".ph-chip{position:fixed;z-index:10001;background:#2c3e50;color:#fff;border:none;border-radius:16px;padding:5px 12px;font-size:12px;font-weight:600;cursor:pointer;box-shadow:0 4px 12px rgba(0,0,0,.2)}" +
    ".ph-chip:hover{background:#3498db}" +
    ".ph-browse input{width:100%;padding:8px 10px;border:1px solid #cfd8dc;border-radius:8px;font-size:14px;margin-bottom:10px}" +
    ".ph-browse h4{margin:12px 0 4px;font-size:12px;color:#8a99a6;text-transform:uppercase;letter-spacing:.04em}" +
    ".ph-browse .ph-item{border:1px solid #eef2f4;margin-bottom:4px}" +
    ".ph-browse .ph-item:hover{background:#e8f1fb}" +
    ".ph-browse-list{max-height:55vh;overflow-y:auto}" +
    ".ph-mgr-top{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:12px 0}" +
    ".ph-mgr-top input[type=search]{flex:1;min-width:160px;padding:7px 10px;border:1px solid #cfd8dc;border-radius:8px;font-size:14px}" +
    ".ph-row{border:1px solid #e6ecef;border-radius:10px;padding:9px 12px;margin-bottom:8px;background:#fbfcfd}" +
    ".ph-row-head{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:3px}" +
    ".ph-tag{font-size:11px;font-weight:600;color:#2c3e50;background:#e8eef2;border-radius:10px;padding:1px 8px}" +
    ".ph-tag.app{background:#eaf4fb;color:#2878b5}" +
    ".ph-uses{font-size:11px;color:#8a99a6;margin-left:auto}" +
    ".ph-ml{font-size:14px;color:#2c3e50}" +
    ".ph-en{font-size:12.5px;color:#6b7c88}" +
    ".ph-row-btns{display:flex;gap:6px;margin-top:6px}" +
    ".ph-row-btns button,.ph-mgr button{font-size:12px}" +
    ".ph-edit label{display:block;font-size:12px;font-weight:600;color:#5d6d7e;margin:8px 0 3px}" +
    ".ph-edit input,.ph-edit select,.ph-edit textarea{width:100%;padding:7px 9px;border:1px solid #cfd8dc;border-radius:7px;font-size:14px;font-family:inherit}" +
    ".ph-edit textarea{min-height:60px;resize:vertical}" +
    ".ph-mgr-foot{display:flex;flex-wrap:wrap;gap:8px;margin-top:12px;padding-top:12px;border-top:1px solid #ecf0f1}" +
    ".ph-status{font-size:13px;font-weight:600;min-height:1.2em;margin-top:8px}" +
    ".ph-status.ok{color:#27ae60}.ph-status.error{color:#e74c3c}" +
    ".ph-switch{display:flex;gap:8px;align-items:center;font-size:14px;color:#2c3e50;margin-top:12px}";
  function addStyles() {
    if (document.getElementById("phStyles")) return;
    var st = document.createElement("style");
    st.id = "phStyles";
    st.textContent = CSS;
    document.head.appendChild(st);
  }

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  /* ── caret position (for placing the list) ── */
  var MIRROR_PROPS = [
    "boxSizing", "width", "borderTopWidth", "borderRightWidth", "borderBottomWidth", "borderLeftWidth",
    "paddingTop", "paddingRight", "paddingBottom", "paddingLeft", "fontStyle", "fontVariant", "fontWeight",
    "fontStretch", "fontSize", "lineHeight", "fontFamily", "textAlign", "textTransform", "textIndent",
    "letterSpacing", "wordSpacing", "tabSize",
  ];
  function caretPoint(el, pos) {
    var cs = getComputedStyle(el);
    var div = document.createElement("div");
    MIRROR_PROPS.forEach(function (p) {
      div.style[p] = cs[p];
    });
    div.style.position = "absolute";
    div.style.visibility = "hidden";
    div.style.top = "0";
    div.style.left = "-9999px";
    div.style.overflow = "hidden";
    div.style.whiteSpace = el.tagName === "TEXTAREA" ? "pre-wrap" : "pre";
    div.style.wordWrap = "break-word";
    div.style.width = el.offsetWidth + "px";
    div.textContent = el.value.substring(0, pos);
    var span = document.createElement("span");
    span.textContent = el.value.substring(pos) || ".";
    div.appendChild(span);
    document.body.appendChild(div);
    var lh = parseFloat(cs.lineHeight);
    if (isNaN(lh)) lh = parseFloat(cs.fontSize) * 1.3;
    var r = el.getBoundingClientRect();
    var pt = {
      x: r.left + span.offsetLeft - el.scrollLeft,
      y: r.top + span.offsetTop - el.scrollTop + lh,
    };
    document.body.removeChild(div);
    // keep inside the field
    pt.y = Math.min(Math.max(pt.y, r.top + lh), r.bottom + 2);
    pt.x = Math.min(Math.max(pt.x, r.left), r.right - 20);
    return pt;
  }

  function layer() {
    return document.fullscreenElement || document.body;
  }
  function place(box, pt) {
    if (box.parentNode !== layer()) layer().appendChild(box);
    box.style.left = "0px";
    box.style.top = "0px";
    box.style.display = "block";
    var w = box.offsetWidth;
    var h = box.offsetHeight;
    var x = Math.min(pt.x, window.innerWidth - w - 8);
    var y = pt.y + 4;
    if (y + h > window.innerHeight - 8) y = Math.max(8, pt.y - h - 28); // open above the line
    box.style.left = Math.max(8, x) + "px";
    box.style.top = y + "px";
  }

  /* ── the suggestion list ── */
  var pop = null; // {el, items, active, navigated, sentenceStart, wordStart, caret}
  var box = null;
  var suppress = false;
  var blankEl = null;

  function eligible(el) {
    if (!el || !el.matches || !isEnabled()) return false;
    if (!el.matches("textarea, input[type=text], input:not([type])")) return false;
    if (el.readOnly || el.disabled) return false;
    if (!el.closest("#formPanel")) return false;
    if (el.matches(EXCLUDE)) return false;
    return true;
  }

  function hide() {
    pop = null;
    if (box) box.style.display = "none";
  }

  function render() {
    if (!box) {
      box = document.createElement("div");
      box.className = "ph-pop";
      box.setAttribute("role", "listbox");
      box.addEventListener("mousedown", function (e) {
        e.preventDefault(); // keep focus in the field
        var item = e.target.closest(".ph-item");
        if (item && pop) accept(+item.getAttribute("data-i"));
      });
    }
    box.innerHTML =
      pop.items
        .map(function (m, i) {
          return (
            '<div class="ph-item' + (i === pop.active ? " active" : "") + '" role="option" data-i="' + i + '">' +
            '<span class="ph-text">' + esc(m.text) + "</span>" +
            '<span class="ph-cat">' + esc(m.p.category || "") + "</span></div>"
          );
        })
        .join("") +
      '<div class="ph-hint">Tab insert · ↑ ↓ choose · Esc close</div>';
    place(box, caretPoint(pop.el, pop.caret));
  }

  function update(el) {
    if (suppress) return;
    if (!eligible(el) || el.selectionStart !== el.selectionEnd) {
      hide();
      return;
    }
    var caret = el.selectionStart;
    var before = el.value.slice(0, caret);
    var sm = before.match(/[^.!?;:\n]*$/);
    var sentence = sm ? sm[0] : "";
    var lead = sentence.match(/^\s*/)[0].length;
    sentence = sentence.slice(lead);
    var word = (before.match(/\S+$/) || [""])[0];
    var items = findMatches(sentence, word);
    if (!items.length) {
      hide();
      return;
    }
    pop = {
      el: el,
      items: items,
      active: 0,
      navigated: false,
      caret: caret,
      sentenceStart: caret - sentence.length,
      wordStart: caret - word.length,
    };
    render();
  }

  function insertText(el, start, end, text) {
    if (/്‍/.test(el.value)) text = toOldChillu(text); // match the field's spelling style
    el.focus();
    el.setRangeText(text, start, end, "end");
    var blank = el.value.indexOf("___", start);
    if (blank !== -1 && blank < start + text.length) {
      el.setSelectionRange(blank, blank + 3);
      blankEl = el;
    } else {
      blankEl = null;
    }
    suppress = true;
    try {
      el.dispatchEvent(new Event("input", { bubbles: true })); // preview, auto-grow, draft save
    } finally {
      suppress = false;
    }
  }

  function accept(i) {
    if (!pop) return;
    var m = pop.items[i];
    var el = pop.el;
    var start = m.replace === "sentence" ? pop.sentenceStart : pop.wordStart;
    var end = pop.caret;
    hide();
    insertText(el, start, end, m.text);
    countUse(m.p.id);
  }

  function onKeyDown(e) {
    var el = e.target;
    // Ctrl+Space: browse all phrases
    if (e.ctrlKey && (e.code === "Space" || e.key === " ") && eligible(el)) {
      e.preventDefault();
      hide();
      browse(el);
      return;
    }
    if (pop && pop.el === el) {
      var n = pop.items.length;
      if (e.key === "ArrowDown" || e.key === "ArrowUp") {
        e.preventDefault();
        pop.active = (pop.active + (e.key === "ArrowDown" ? 1 : n - 1)) % n;
        pop.navigated = true;
        render();
        return;
      }
      if (e.key === "Tab" && !e.shiftKey) {
        e.preventDefault();
        accept(pop.active);
        return;
      }
      if (e.key === "Enter" && pop.navigated) {
        e.preventDefault();
        accept(pop.active);
        return;
      }
      if (e.key === "Escape") {
        e.preventDefault();
        e.stopPropagation(); // do not close a surrounding dialog
        hide();
        return;
      }
      if (e.key === "Enter") hide(); // plain Enter: new line as usual
      return;
    }
    // Tab jumps to the next ___ blank after an inserted phrase
    if (e.key === "Tab" && !e.shiftKey && blankEl === el) {
      var from = el.selectionStart === el.selectionEnd ? el.selectionEnd : el.selectionEnd;
      var next = el.value.indexOf("___", from);
      if (next !== -1) {
        e.preventDefault();
        el.setSelectionRange(next, next + 3);
        return;
      }
      blankEl = null;
    }
  }

  /* ── browse all (Ctrl+Space) ── */
  var browseTarget = null;
  function browse(el) {
    if (typeof showModal !== "function") return;
    browseTarget = { el: el, start: el.selectionStart, end: el.selectionEnd };
    showModal(
      "Phrases",
      '<div class="ph-browse"><input type="search" id="phBrowseSearch" placeholder="Type to search (Malayalam or English)…" autocomplete="off">' +
        '<div class="ph-browse-list" id="phBrowseList"></div></div>',
    );
    var input = document.getElementById("phBrowseSearch");
    var list = document.getElementById("phBrowseList");
    function draw() {
      var q = norm(input.value).trim();
      var groups = {};
      var order = [];
      getList().forEach(function (p) {
        if (!forThisApp(p)) return;
        ["ml", "en"].forEach(function (lang) {
          var t = p[lang];
          if (!t || (q && norm(t).indexOf(q) === -1)) return;
          var g = p.category || "Other";
          if (!groups[g]) {
            groups[g] = [];
            order.push(g);
          }
          groups[g].push({ p: p, text: t });
        });
      });
      list.innerHTML = order.length
        ? order
            .map(function (g) {
              return (
                "<h4>" + esc(g) + "</h4>" +
                groups[g]
                  .map(function (m) {
                    return '<div class="ph-item" data-id="' + esc(m.p.id) + '" data-text="' + esc(m.text) + '"><span class="ph-text">' + esc(m.text) + "</span></div>";
                  })
                  .join("")
              );
            })
            .join("")
        : '<p style="color:#888;font-size:13px">No phrase matches.</p>';
    }
    input.addEventListener("input", draw);
    list.addEventListener("click", function (e) {
      var item = e.target.closest(".ph-item");
      if (!item || !browseTarget) return;
      var t = browseTarget;
      browseTarget = null;
      closeModal();
      insertText(t.el, t.start, t.end, item.getAttribute("data-text"));
      countUse(item.getAttribute("data-id"));
    });
    draw();
    input.focus();
  }

  /* ── "Save as phrase" for selected text ── */
  var chip = null;
  var chipFor = null;
  function hideChip() {
    if (chip) chip.style.display = "none";
    chipFor = null;
  }
  function checkSelection(e) {
    var el = e.target;
    if (!eligible(el)) {
      hideChip();
      return;
    }
    var s = el.selectionStart;
    var t = el.value.slice(s, el.selectionEnd).trim();
    if (t.length < 4 || t.length > 600) {
      hideChip();
      return;
    }
    if (!chip) {
      chip = document.createElement("button");
      chip.type = "button";
      chip.className = "ph-chip";
      chip.textContent = "★ Save as phrase";
      chip.addEventListener("mousedown", function (ev) {
        ev.preventDefault();
        if (chipFor) savePhrase(chipFor);
        hideChip();
      });
    }
    chipFor = t;
    place(chip, caretPoint(el, el.selectionEnd));
  }

  function savePhrase(text) {
    var ml = hasMalayalam(text);
    var dup = getList().some(function (p) {
      return sameText(ml ? p.ml : p.en, text);
    });
    if (dup) {
      if (typeof showToast === "function") showToast("This phrase is already saved.");
      return;
    }
    try {
      var list = getList().slice();
      list.push({ id: newId(), category: "My phrases", apps: "All", ml: ml ? text : "", en: ml ? "" : text });
      saveList(list);
      if (typeof showToast === "function") showToast("Saved as a phrase. Type its first letters to use it.");
    } catch (err) {
      if (typeof showToast === "function") showToast("Could not save the phrase.", "error");
    }
  }

  function initApp() {
    if (typeof APP === "undefined" || !document.getElementById("formPanel")) return;
    addStyles();
    document.addEventListener("input", function (e) {
      if (e.isComposing) return; // wait for the word to be finished (phonetic keyboards)
      update(e.target);
    }, true);
    document.addEventListener("compositionend", function (e) {
      setTimeout(function () {
        update(e.target);
      }, 0);
    }, true);
    document.addEventListener("keydown", onKeyDown, true);
    document.addEventListener("focusout", function (e) {
      if (pop && e.target === pop.el) hide();
      if (e.target === blankEl) blankEl = null;
      hideChip();
    }, true);
    document.addEventListener("mousedown", function (e) {
      if (pop && box && !box.contains(e.target) && e.target !== pop.el) hide();
    }, true);
    document.addEventListener("mouseup", checkSelection, true);
    document.addEventListener("keyup", function (e) {
      if (e.shiftKey || e.key === "Shift" || (e.ctrlKey && e.key === "a")) checkSelection(e);
      else if (chip && chip.style.display !== "none") hideChip();
    }, true);
    window.addEventListener("resize", hide);
    // Typing near the bottom of a long field makes the browser scroll to keep
    // the cursor in view: move the list with the cursor instead of closing it,
    // and close it only when the cursor has scrolled out of sight.
    document.addEventListener("scroll", function (e) {
      if (!pop || !box || box.style.display === "none") return;
      if (e.target !== pop.el && !(e.target.contains && e.target.contains(pop.el))) return;
      var r = pop.el.getBoundingClientRect();
      var pt = caretPoint(pop.el, pop.caret);
      if (r.bottom < 0 || r.top > window.innerHeight || pt.y < 0 || pt.y > window.innerHeight) hide();
      else place(box, pt);
    }, true);
  }

  /* ── CSV import / export ── */
  function parseCSV(text) {
    var rows = [], row = [], field = "", q = false;
    text = text.replace(/^﻿/, "");
    for (var i = 0; i < text.length; i++) {
      var c = text[i];
      if (q) {
        if (c === '"' && text[i + 1] === '"') {
          field += '"';
          i++;
        } else if (c === '"') q = false;
        else field += c;
      } else if (c === '"') q = true;
      else if (c === ",") {
        row.push(field);
        field = "";
      } else if (c === "\n" || c === "\r") {
        if (c === "\r" && text[i + 1] === "\n") i++;
        row.push(field);
        rows.push(row);
        row = [];
        field = "";
      } else field += c;
    }
    if (field || row.length) {
      row.push(field);
      rows.push(row);
    }
    return rows.filter(function (r) {
      return r.some(function (v) {
        return v.trim();
      });
    });
  }

  // Reads a phrase sheet (the starter list layout, or any sheet with
  // Malayalam / English columns). Rows with Keep = N are skipped and
  // "Your correction" replaces the Malayalam text when filled in.
  function phrasesFromCSV(text) {
    var rows = parseCSV(text);
    if (rows.length < 2) throw new Error("The file has no phrases.");
    var head = rows[0].map(function (h) {
      return h.trim().toLowerCase();
    });
    function col(re) {
      for (var i = 0; i < head.length; i++) if (re.test(head[i])) return i;
      return -1;
    }
    var cMl = col(/malayalam/), cEn = col(/english/), cCat = col(/categ/), cApps = col(/use in|^apps?$/);
    var cKeep = col(/keep/), cFix = col(/correction/);
    if (cMl === -1 && cEn === -1) throw new Error("No Malayalam or English column found in the file.");
    var out = [];
    rows.slice(1).forEach(function (r) {
      var get = function (i) {
        return i === -1 ? "" : (r[i] || "").trim();
      };
      if (/^n/i.test(get(cKeep))) return;
      var ml = get(cFix) || get(cMl);
      var en = get(cEn);
      if (!ml && !en) return;
      var apps = get(cApps) || "All";
      out.push({ category: get(cCat) || "My phrases", apps: apps, ml: ml, en: en });
    });
    return out;
  }

  function csvCell(v) {
    v = String(v == null ? "" : v);
    return /[",\r\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
  }
  function exportCSV() {
    var lines = [["Category", "Use in", "Malayalam", "English", "Keep (Y/N)", "Your correction"].join(",")];
    getList().forEach(function (p) {
      lines.push([p.category, p.apps, p.ml, p.en, "", ""].map(csvCell).join(","));
    });
    var blob = new Blob(["﻿" + lines.join("\r\n") + "\r\n"], { type: "text/csv;charset=utf-8" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = "office_suite_phrases.csv";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    return getList().length;
  }

  /* ── manager (home page, Settings > Phrases) ── */
  function renderManager(root) {
    addStyles();
    var editing = null; // id being edited, or "new"
    var query = "";
    var status = "";
    var statusKind = "";

    function setStatus(t, k) {
      status = t;
      statusKind = k || "";
      var el = root.querySelector(".ph-status");
      if (el) {
        el.textContent = t;
        el.className = "ph-status" + (k ? " " + k : "");
      }
    }

    function editForm(p) {
      var cats = [];
      getList().forEach(function (x) {
        if (x.category && cats.indexOf(x.category) === -1) cats.push(x.category);
      });
      return (
        '<div class="ph-row ph-edit" data-edit="' + esc(p.id || "new") + '">' +
        '<label>Malayalam</label><textarea data-f="ml">' + esc(p.ml) + "</textarea>" +
        '<label>English</label><textarea data-f="en">' + esc(p.en) + "</textarea>" +
        '<div style="display:flex;gap:10px;flex-wrap:wrap"><div style="flex:1;min-width:150px"><label>Category</label>' +
        '<input data-f="category" list="phCats" value="' + esc(p.category || "My phrases") + '"><datalist id="phCats">' +
        cats.map(function (c) { return '<option value="' + esc(c) + '">'; }).join("") +
        '</datalist></div><div style="flex:1;min-width:150px"><label>Use in</label><select data-f="apps">' +
        APP_CHOICES.map(function (a) {
          return "<option" + ((p.apps || "All") === a ? " selected" : "") + ">" + a + "</option>";
        }).join("") +
        "</select></div></div>" +
        '<p style="font-size:12px;color:#8a99a6;margin-top:6px">Use ___ (three underscores) for blanks to fill in.</p>' +
        '<div class="ph-row-btns"><button type="button" class="btn-sm btn-sample" data-act="save">Save</button>' +
        '<button type="button" class="btn-sm btn-outline" data-act="cancel">Cancel</button></div></div>'
      );
    }

    function draw() {
      var list = getList();
      var usage = getUsage();
      var q = norm(query).trim();
      var shown = list.filter(function (p) {
        return !q || norm(p.ml).indexOf(q) !== -1 || norm(p.en).indexOf(q) !== -1 || norm(p.category).indexOf(q) !== -1;
      });
      root.innerHTML =
        '<div class="ph-mgr">' +
        '<label class="ph-switch"><input type="checkbox" data-act="toggle"' + (isEnabled() ? " checked" : "") +
        "> Suggest phrases while typing in the apps</label>" +
        '<div class="ph-mgr-top"><input type="search" placeholder="Search phrases…" value="' + esc(query) + '" data-act="search">' +
        '<button type="button" class="btn-sm btn-upload" data-act="add">+ Add phrase</button></div>' +
        (editing === "new" ? editForm({ id: "new", category: "My phrases", apps: "All", ml: "", en: "" }) : "") +
        (shown.length
          ? shown
              .map(function (p) {
                if (editing === p.id) return editForm(p);
                return (
                  '<div class="ph-row"><div class="ph-row-head"><span class="ph-tag">' + esc(p.category || "—") + "</span>" +
                  (p.apps && p.apps !== "All" ? '<span class="ph-tag app">' + esc(p.apps) + "</span>" : "") +
                  '<span class="ph-uses">' + (usage[p.id] ? "used " + usage[p.id] + "×" : "") + "</span></div>" +
                  (p.ml ? '<div class="ph-ml">' + esc(p.ml) + "</div>" : "") +
                  (p.en ? '<div class="ph-en">' + esc(p.en) + "</div>" : "") +
                  '<div class="ph-row-btns"><button type="button" class="btn-sm btn-outline" data-act="edit" data-id="' + esc(p.id) + '">Edit</button>' +
                  '<button type="button" class="btn-sm btn-outline" data-act="del" data-id="' + esc(p.id) + '">Delete</button></div></div>'
                );
              })
              .join("")
          : '<p style="color:#888;font-size:13px">No phrases match.</p>') +
        '<div class="ph-mgr-foot">' +
        '<button type="button" class="btn-sm btn-outline" data-act="import">Import CSV</button>' +
        '<button type="button" class="btn-sm btn-outline" data-act="export">Export CSV</button>' +
        '<button type="button" class="btn-sm btn-outline" data-act="reset"' + (isCustom() ? "" : " disabled") + ">Reset to starter list</button>" +
        '<input type="file" accept=".csv,text/csv" hidden data-act="file"></div>' +
        '<p style="font-size:12px;color:#8a99a6;margin-top:8px">' + list.length +
        " phrases. Import adds phrases from an Excel sheet saved as <b>CSV UTF-8</b> (columns Category, Use in, Malayalam, English). Phrases are included in “Back up all drafts”.</p>" +
        '<div class="ph-status' + (statusKind ? " " + statusKind : "") + '" role="status">' + esc(status) + "</div></div>";
      var s = root.querySelector('[data-act="search"]');
      if (s && document.activeElement === document.body && query) s.focus();
    }

    root.oninput = function (e) {
      if (e.target.getAttribute("data-act") === "search") {
        query = e.target.value;
        var pos = e.target.selectionStart;
        draw();
        var s = root.querySelector('[data-act="search"]');
        s.focus();
        s.setSelectionRange(pos, pos);
      }
    };
    root.onchange = function (e) {
      var act = e.target.getAttribute("data-act");
      if (act === "toggle") {
        setEnabled(e.target.checked);
        setStatus(e.target.checked ? "Suggestions are on." : "Suggestions are off.", "ok");
      }
      if (act === "file") {
        var file = e.target.files[0];
        e.target.value = "";
        if (!file) return;
        var r = new FileReader();
        r.onload = function () {
          try {
            var incoming = phrasesFromCSV(r.result);
            var n = mergeList(incoming);
            draw();
            setStatus(n ? "Added " + n + " phrase(s)." : "No new phrases: all of them are already in your list.", "ok");
          } catch (err) {
            setStatus(err.message, "error");
          }
        };
        r.readAsText(file, "utf-8");
      }
    };
    root.onclick = function (e) {
      var b = e.target.closest("[data-act]");
      if (!b || b.tagName === "INPUT") return;
      var act = b.getAttribute("data-act");
      var id = b.getAttribute("data-id");
      if (act === "add") {
        editing = "new";
        draw();
        root.querySelector('[data-f="ml"]').focus();
      } else if (act === "edit") {
        editing = id;
        draw();
      } else if (act === "cancel") {
        editing = null;
        draw();
      } else if (act === "save") {
        var form = b.closest("[data-edit]");
        var v = {};
        form.querySelectorAll("[data-f]").forEach(function (f) {
          v[f.getAttribute("data-f")] = f.value.trim();
        });
        if (!v.ml && !v.en) {
          setStatus("Enter the phrase in Malayalam or English.", "error");
          return;
        }
        var list = getList().slice();
        if (editing === "new") {
          list.push({ id: newId(), category: v.category || "My phrases", apps: v.apps || "All", ml: v.ml, en: v.en });
        } else {
          list = list.map(function (p) {
            return p.id === editing ? { id: p.id, category: v.category || "My phrases", apps: v.apps || "All", ml: v.ml, en: v.en } : p;
          });
        }
        try {
          saveList(list);
          editing = null;
          draw();
          setStatus("Phrase saved.", "ok");
        } catch (err) {
          setStatus("Could not save. Browser storage may be blocked.", "error");
        }
      } else if (act === "del") {
        var p = getList().filter(function (x) { return x.id === id; })[0];
        if (!p || !confirm('Delete this phrase?\n\n' + (p.ml || p.en))) return;
        saveList(getList().filter(function (x) { return x.id !== id; }));
        draw();
        setStatus("Phrase deleted.", "ok");
      } else if (act === "import") {
        root.querySelector('[data-act="file"]').click();
      } else if (act === "export") {
        var n = exportCSV();
        setStatus("Exported " + n + " phrases to a CSV file.", "ok");
      } else if (act === "reset") {
        if (!confirm("Replace your phrase list with the starter list? Phrases you added will be removed.")) return;
        resetList();
        draw();
        setStatus("Starter list restored.", "ok");
      }
    };
    draw();
  }

  if (typeof window !== "undefined") {
    window.addEventListener("load", function () {
      setTimeout(initApp, 0);
    });
  }

  return {
    norm: norm,
    findMatches: findMatches,
    getList: getList,
    isCustom: isCustom,
    mergeList: mergeList,
    phrasesFromCSV: phrasesFromCSV,
    exportCSV: exportCSV,
    renderManager: renderManager,
    isEnabled: isEnabled,
    setEnabled: setEnabled,
  };
})();
