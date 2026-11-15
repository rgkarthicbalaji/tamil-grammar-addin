/* global Office, Word, TamilChecker */
/*
 * Tamil Grammar Checker — Word task pane.
 *
 * Flow: read every paragraph → TamilChecker.checkParagraph() → show cards in
 * the pane and (when the Word build supports it) underline each issue with
 * Word's annotation API (WordApi 1.7; pop-up suggestions need 1.8).
 * Fixes are applied by searching the issue text inside its paragraph and
 * replacing that occurrence.
 */
(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const SEV_LABEL = { error: "பிழை · Error", warning: "எச்சரிக்கை · Warning", suggestion: "பரிந்துரை · Suggestion" };
  const COLOR = { error: "Red", warning: "Lavender", suggestion: "Blue" };

  const state = {
    api: { annotations: false, popup: false, ids: false },
    paras: new Map(), // key -> { key, index, text, issues }
    annotations: new Map(), // annotation id -> { key, issue }
    ignored: new Set(load("ignored", [])),
    busy: false,
    live: null, // event handler results
    pending: new Set(),
    timer: null,
  };

  // ---------------------------------------------------------------------------
  // settings and storage
  // ---------------------------------------------------------------------------
  function load(k, d) {
    try { const v = localStorage.getItem("tg." + k); return v ? JSON.parse(v) : d; } catch (e) { return d; }
  }
  function save(k, v) {
    try { localStorage.setItem("tg." + k, JSON.stringify(v)); } catch (e) { /* storage unavailable */ }
  }
  const SETTINGS = ["opt-formal", "opt-suggestions", "opt-live", "opt-underline"];
  function restoreSettings() {
    const s = load("settings", {});
    SETTINGS.forEach((id) => { if (id in s) $(id).checked = !!s[id]; });
  }
  function storeSettings() {
    const s = {};
    SETTINGS.forEach((id) => (s[id] = $(id).checked));
    save("settings", s);
  }
  const options = () => ({
    formal: $("opt-formal").checked,
    style: true,
    minSeverity: $("opt-suggestions").checked ? "suggestion" : "warning",
  });
  const underlineOn = () => state.api.annotations && $("opt-underline").checked;

  const ignoreKey = (text, it) =>
    it.rule + "|" + it.text + "|" + text.slice(Math.max(0, it.start - 10), it.start + it.length + 10);
  const notIgnored = (text, issues) => issues.filter((it) => !state.ignored.has(ignoreKey(text, it)));
  const runChecker = (text) => notIgnored(text, TamilChecker.checkParagraph(text, options()));

  function setStatus(msg, ok) {
    $("status").textContent = msg;
    $("status").className = "status" + (ok ? " ok" : "");
  }
  function showError(e) {
    console.error(e);
    setStatus("பிழை ஏற்பட்டது / Something went wrong: " + (e && e.message ? e.message : e));
  }

  // ---------------------------------------------------------------------------
  // Word helpers
  // ---------------------------------------------------------------------------
  function critique(it) {
    const c = { colorScheme: COLOR[it.severity], start: it.start, length: Math.max(1, it.length) };
    if (state.api.popup && it.suggestions.length) {
      c.popupOptions = {
        brandingTextResourceId: "TG.Brand",
        titleResourceId: "TG.PopupTitle",
        subtitleResourceId: "TG.PopupSubtitle",
        suggestions: it.suggestions.slice(0, 3),
      };
    }
    return c;
  }

  async function getParagraph(ctx, entry) {
    if (state.api.ids) return ctx.document.getParagraphByUniqueLocalId(entry.key);
    const paras = ctx.document.body.paragraphs;
    paras.load("items");
    await ctx.sync();
    return paras.items[entry.index];
  }

  /** Find the Word range for an issue: the n-th occurrence of its text in the paragraph. */
  async function findRange(ctx, para, entry, it) {
    let n = 0;
    let pos = entry.text.indexOf(it.text);
    while (pos !== -1 && pos < it.start) {
      n++;
      pos = entry.text.indexOf(it.text, pos + Math.max(1, it.text.length));
    }
    const res = para.search(it.text, { matchCase: true });
    res.load("items");
    await ctx.sync();
    return res.items[n] || res.items[0] || null;
  }

  async function clearAnnotations(ctx, para) {
    const anns = para.getAnnotations();
    anns.load("items/id");
    await ctx.sync();
    anns.items.forEach((a) => { state.annotations.delete(a.id); a.delete(); });
    await ctx.sync();
  }

  async function annotate(ctx, para, entry) {
    if (!entry.issues.length) return;
    const ids = para.insertAnnotations({ critiques: entry.issues.map(critique) });
    await ctx.sync();
    ids.value.forEach((id, k) => state.annotations.set(id, { key: entry.key, issue: entry.issues[k] }));
  }

  function annotationFailed(e) {
    console.warn("Annotations unavailable:", e);
    state.api.annotations = false;
    $("api-note").textContent =
      "இந்த Word பதிப்பில் அடிக்கோடு இயலாது; பிழைகள் இப்பலகையில் மட்டும் காட்டப்படும். " +
      "(Underlines need Word with a Microsoft 365 subscription; issues are listed here instead.)";
  }

  // ---------------------------------------------------------------------------
  // Checking
  // ---------------------------------------------------------------------------
  async function checkDocument() {
    if (state.busy) return;
    state.busy = true;
    $("check").disabled = true;
    setStatus("சரிபார்க்கப்படுகிறது… / Checking…");
    try {
      await Word.run(async (ctx) => {
        const paras = ctx.document.body.paragraphs;
        paras.load(state.api.ids ? "items/text,items/uniqueLocalId" : "items/text");
        await ctx.sync();
        state.paras.clear();
        const byIndex = paras.items.map((p, i) => {
          const key = state.api.ids ? p.uniqueLocalId : "p" + i;
          const entry = { key, index: i, text: p.text, issues: runChecker(p.text) };
          state.paras.set(key, entry);
          return entry;
        });
        if (state.api.annotations) {
          try {
            state.annotations.clear();
            const colls = paras.items.map((p) => { const a = p.getAnnotations(); a.load("items/id"); return a; });
            await ctx.sync();
            colls.forEach((c) => c.items.forEach((a) => a.delete()));
            await ctx.sync();
            if (underlineOn()) {
              const pendingIds = [];
              paras.items.forEach((p, i) => {
                const entry = byIndex[i];
                if (entry.issues.length)
                  pendingIds.push({ entry, ids: p.insertAnnotations({ critiques: entry.issues.map(critique) }) });
              });
              await ctx.sync();
              pendingIds.forEach(({ entry, ids }) =>
                ids.value.forEach((id, k) => state.annotations.set(id, { key: entry.key, issue: entry.issues[k] })));
            }
          } catch (e) {
            annotationFailed(e);
          }
        }
      });
      render();
    } catch (e) {
      showError(e);
    } finally {
      state.busy = false;
      $("check").disabled = false;
    }
  }

  /** Re-check one paragraph (after a fix, an ignore, or a live edit). */
  async function recheck(ctx, para, key, index) {
    para.load("text");
    await ctx.sync();
    const entry = { key, index, text: para.text, issues: runChecker(para.text) };
    state.paras.set(key, entry);
    if (state.api.annotations) {
      try {
        await clearAnnotations(ctx, para);
        if (underlineOn()) await annotate(ctx, para, entry);
      } catch (e) {
        annotationFailed(e);
      }
    }
    return entry;
  }

  async function applyFix(key, it, suggestion) {
    try {
      await Word.run(async (ctx) => {
        const entry = state.paras.get(key);
        const para = await getParagraph(ctx, entry);
        para.load("text");
        await ctx.sync();
        if (para.text !== entry.text) {
          // the paragraph was edited since the check: refresh instead of guessing
          await recheck(ctx, para, key, entry.index);
          setStatus("பத்தி மாறியுள்ளது; மீண்டும் சரிபார்க்கப்பட்டது. / The paragraph changed; re-checked it.");
          return;
        }
        const range = await findRange(ctx, para, entry, it);
        if (!range) {
          setStatus("இடம் கண்டுபிடிக்க முடியவில்லை. / Couldn't locate the text in Word.");
          return;
        }
        range.insertText(suggestion, "Replace");
        await ctx.sync();
        await recheck(ctx, para, key, entry.index);
      });
      render();
    } catch (e) {
      showError(e);
    }
  }

  async function goTo(key, it) {
    try {
      await Word.run(async (ctx) => {
        const entry = state.paras.get(key);
        const para = await getParagraph(ctx, entry);
        const range = await findRange(ctx, para, entry, it);
        (range || para).select();
        await ctx.sync();
      });
    } catch (e) {
      showError(e);
    }
  }

  async function ignore(key, it) {
    const entry = state.paras.get(key);
    state.ignored.add(ignoreKey(entry.text, it));
    save("ignored", [...state.ignored]);
    entry.issues = entry.issues.filter((x) => x !== it);
    if (state.api.annotations) {
      const ids = [...state.annotations.entries()].filter(([, v]) => v.issue === it).map(([id]) => id);
      try {
        await Word.run(async (ctx) => {
          ids.forEach((id) => { ctx.document.getAnnotationById(id).delete(); state.annotations.delete(id); });
          await ctx.sync();
        });
      } catch (e) { console.warn(e); }
    }
    render();
  }

  // ---------------------------------------------------------------------------
  // Events: pop-up accept/reject and live checking
  // ---------------------------------------------------------------------------
  async function onPopupAction(args) {
    const hit = state.annotations.get(args.id);
    if (!hit) return;
    if (args.action !== "Accept") {
      await ignore(hit.key, hit.issue);
      return;
    }
    // Word may already have replaced the text; if not, apply the chosen suggestion.
    try {
      await Word.run(async (ctx) => {
        const entry = state.paras.get(hit.key);
        const para = await getParagraph(ctx, entry);
        para.load("text");
        await ctx.sync();
        const unchanged = para.text.substr(hit.issue.start, hit.issue.length) === hit.issue.text && para.text === entry.text;
        if (unchanged && args.critiqueSuggestion) {
          const range = await findRange(ctx, para, entry, hit.issue);
          if (range) { range.insertText(args.critiqueSuggestion, "Replace"); await ctx.sync(); }
        }
        await recheck(ctx, para, hit.key, entry.index);
      });
      render();
    } catch (e) {
      showError(e);
    }
  }

  function onParagraphsChanged(args) {
    (args.uniqueLocalIds || []).forEach((id) => state.pending.add(id));
    clearTimeout(state.timer);
    state.timer = setTimeout(flushPending, 900);
  }

  async function flushPending() {
    const ids = [...state.pending];
    state.pending.clear();
    if (!ids.length) return;
    try {
      await Word.run(async (ctx) => {
        for (const id of ids) {
          const para = ctx.document.getParagraphByUniqueLocalIdOrNullObject
            ? ctx.document.getParagraphByUniqueLocalIdOrNullObject(id)
            : ctx.document.getParagraphByUniqueLocalId(id);
          para.load("text");
          await ctx.sync();
          if (para.isNullObject) { state.paras.delete(id); continue; }
          const old = state.paras.get(id);
          await recheck(ctx, para, id, old ? old.index : Number.MAX_SAFE_INTEGER);
        }
      });
      render();
    } catch (e) {
      console.warn(e);
    }
  }

  async function setLive(on) {
    if (!state.api.ids) return;
    try {
      if (on && !state.live) {
        await Word.run(async (ctx) => {
          state.live = [
            ctx.document.onParagraphChanged.add(onParagraphsChanged),
            ctx.document.onParagraphAdded.add(onParagraphsChanged),
          ];
          await ctx.sync();
        });
      } else if (!on && state.live) {
        const handlers = state.live;
        state.live = null;
        await Word.run(handlers[0].context, async (ctx) => {
          handlers.forEach((h) => h.remove());
          await ctx.sync();
        });
      }
    } catch (e) {
      console.warn("Live checking unavailable:", e);
      $("opt-live").checked = false;
      $("opt-live").disabled = true;
    }
  }

  async function registerPopupEvents() {
    if (!state.api.popup) return;
    try {
      await Word.run(async (ctx) => {
        ctx.document.onAnnotationPopupAction.add(onPopupAction);
        await ctx.sync();
      });
    } catch (e) {
      console.warn("Pop-up events unavailable:", e);
    }
  }

  // ---------------------------------------------------------------------------
  // Rendering
  // ---------------------------------------------------------------------------
  function el(tag, cls, text) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined) e.textContent = text;
    return e;
  }

  function card(entry, it) {
    const c = el("div", "card " + it.severity);
    const meta = el("div", "meta");
    meta.append(el("span", "rule", it.rule + " · " + SEV_LABEL[it.severity]));
    const go = el("button", "goto", "காட்டு · Show");
    go.onclick = () => goTo(entry.key, it);
    meta.append(go);
    c.append(meta);
    c.append(el("div", "msg", it.message_ta));
    c.append(el("div", "en", it.message_en));
    const fix = el("div", "fix");
    fix.append(el("span", "wrong", it.text.trim() || "␣"));
    if (it.suggestions.length) {
      fix.append(el("span", "arrow", "→"));
      it.suggestions.slice(0, 3).forEach((s) => {
        const b = el("button", "sugg", s.trim() || "␣");
        b.title = "இந்தத் திருத்தத்தைச் செய் · Apply this fix";
        b.onclick = () => applyFix(entry.key, it, s);
        fix.append(b);
      });
    }
    const ig = el("button", "ignore", "புறக்கணி · Ignore");
    ig.onclick = () => ignore(entry.key, it);
    fix.append(ig);
    c.append(fix);
    return c;
  }

  function render() {
    const list = $("list");
    list.textContent = "";
    const counts = { error: 0, warning: 0, suggestion: 0 };
    const entries = [...state.paras.values()].filter((e) => e.issues.length).sort((a, b) => a.index - b.index);
    for (const entry of entries) {
      const snippet = entry.text.trim().slice(0, 40);
      list.append(el("div", "para-label", "¶ " + snippet + (entry.text.trim().length > 40 ? "…" : "")));
      for (const it of entry.issues) {
        counts[it.severity]++;
        list.append(card(entry, it));
      }
    }
    $("n-error").textContent = counts.error;
    $("n-warning").textContent = counts.warning;
    $("n-suggestion").textContent = counts.suggestion;
    $("summary").hidden = false;
    const total = counts.error + counts.warning + counts.suggestion;
    if (!state.paras.size) setStatus("ஆவணத்தில் உரை இல்லை. / The document is empty.");
    else if (!total) setStatus("பிழைகள் எதுவும் கண்டறியப்படவில்லை ✓ / No issues found.", true);
    else setStatus(total + " குறிப்புகள் · " + total + " issues");
  }

  // ---------------------------------------------------------------------------
  // Start-up
  // ---------------------------------------------------------------------------
  Office.onReady((info) => {
    if (info.host !== Office.HostType.Word) {
      setStatus("இந்தச் செருகுநிரல் Word-இல் மட்டும் இயங்கும். / This add-in runs in Word only.");
      return;
    }
    const req = Office.context.requirements;
    state.api.annotations = req.isSetSupported("WordApi", "1.7");
    state.api.popup = req.isSetSupported("WordApi", "1.8");
    state.api.ids = req.isSetSupported("WordApi", "1.6");
    if (!state.api.annotations) annotationFailed("WordApi 1.7 not supported");
    if (!state.api.ids) {
      $("opt-live").checked = false;
      $("opt-live").disabled = true;
    }

    restoreSettings();
    SETTINGS.forEach((id) => ($(id).onchange = () => {
      storeSettings();
      if (id === "opt-live") setLive($(id).checked);
      else if (state.paras.size) checkDocument();
    }));
    $("reset-ignored").onclick = () => {
      state.ignored.clear();
      save("ignored", []);
      if (state.paras.size) checkDocument();
    };
    $("check").onclick = checkDocument;
    $("check").disabled = false;
    setStatus("ஆவணத்தைச் சரிபார்க்க \"சரிபார்\" அழுத்துக. / Press Check to scan the document.");

    registerPopupEvents();
    if ($("opt-live").checked) setLive(true);
  });
})();
