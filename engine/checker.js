/*
 * Tamil grammar checker — rule engine.
 *
 * checkParagraph(text, options) -> Issue[]
 *   Issue = { rule, family, severity: "error"|"warning"|"suggestion",
 *             start, length, text, suggestions: string[],
 *             message_ta, message_en }
 *
 * Offsets are UTF-16 indices into `text`, which is what Word's
 * paragraph-relative critiques (insertAnnotations) expect.
 *
 * Rule IDs match "Tamil Grammar Checker — Rule Reference". This file
 * implements the rules that need no morphological analyser: letter-level
 * rules (L1), word lists (L2), and suffix heuristics standing in for
 * morphology (L3/L4). Heuristic rules are marked "warning" so they underline
 * gently when they might be wrong.
 */
(function (root) {
  "use strict";

  const LEX =
    typeof module !== "undefined" && module.exports
      ? require("./lexicon.js")
      : root.TamilLexicon;

  // ---------------------------------------------------------------------------
  // Letters
  // ---------------------------------------------------------------------------
  const VIRAMA = "்"; // ்
  const AYTHAM = "ஃ"; // ஃ
  const VOWEL_SIGNS = "ாிீுூெேைொோௌ";
  const IND_VOWELS = "அஆஇஈஉஊஎஏஐஒஓஔ";
  const SHORT_VOWELS = "அஇஉஎஒ";
  const CONSONANTS = "கஙசஞடணதநபமயரலவழளறனஜஷஸஹ";
  const GRANTHA = "ஜஷஸஹ";
  const HARD = "கசதப"; // the four that double at word boundaries
  const VOWEL_TO_SIGN = {
    "அ": "", "ஆ": "ா", "இ": "ி", "ஈ": "ீ", "உ": "ு", "ஊ": "ூ",
    "எ": "ெ", "ஏ": "ே", "ஐ": "ை", "ஒ": "ொ", "ஓ": "ோ", "ஔ": "ௌ",
  };
  const isCons = (c) => c !== undefined && CONSONANTS.includes(c);
  const isSign = (c) => c !== undefined && (VOWEL_SIGNS.includes(c) || c === VIRAMA || c === "ௗ");
  const isIndVowel = (c) => c !== undefined && IND_VOWELS.includes(c);
  const hasGrantha = (w) => /[ஜஷஸஹ]|ஸ்ரீ|க்ஷ/.test(w);

  /** First hard consonant of a word if it starts with க/ச/த/ப + a vowel. */
  function hardInitial(word) {
    const c = word[0];
    if (HARD.includes(c) && word[1] !== VIRAMA) return c;
    return null;
  }

  /** If the word ends in a doubling consonant (…க் …ச் …த் …ப்) return {base, cons}. */
  function trailingDoubling(word) {
    if (word.length < 3 || !word.endsWith(VIRAMA)) return null;
    const cons = word[word.length - 2];
    if (!HARD.includes(cons)) return null;
    // a consonant cluster before it (…ர்க், …ஸ்ட்) means a loanword or verb stem, not sandhi
    if (word[word.length - 3] === VIRAMA) return null;
    return { base: word.slice(0, -2), cons };
  }

  // ---------------------------------------------------------------------------
  // Tokenising
  // ---------------------------------------------------------------------------
  const TOKEN_RE = /[஀-௿‌‍]+|[A-Za-z0-9]+|[^\s]/g;
  const TAMIL_WORD_RE = /^[஀-௿‌‍]+$/;
  const TAMIL_DIGIT_RE = /^[௦-௯]+$/;

  function tokenize(text) {
    const tokens = [];
    let m;
    TOKEN_RE.lastIndex = 0;
    while ((m = TOKEN_RE.exec(text)) !== null) {
      const t = m[0];
      const isTamil = TAMIL_WORD_RE.test(t) && !TAMIL_DIGIT_RE.test(t);
      tokens.push({
        text: t,
        start: m.index,
        end: m.index + t.length,
        tamil: isTamil,
        punct: !isTamil && /^[^\sA-Za-z0-9஀-௿]$/.test(t),
      });
    }
    return tokens;
  }

  /** Split tokens into sentences on . ? ! (the terminator stays with its sentence). */
  function sentences(tokens) {
    const out = [];
    let cur = [];
    for (const t of tokens) {
      cur.push(t);
      if (t.punct && /[.?!।]/.test(t.text)) {
        out.push(cur);
        cur = [];
      }
    }
    if (cur.length) out.push(cur);
    return out;
  }

  // ---------------------------------------------------------------------------
  // Issue helpers
  // ---------------------------------------------------------------------------
  const FAMILY = (id) => id.split("-")[1];
  function issue(rule, severity, start, length, text, suggestions, ta, en) {
    return {
      rule, family: FAMILY(rule), severity, start, length,
      text: text.substr(start, length),
      suggestions: (suggestions || []).filter((s) => s !== undefined && s !== null),
      message_ta: ta, message_en: en,
    };
  }
  const tokIssue = (rule, sev, tok, text, sugg, ta, en) =>
    issue(rule, sev, tok.start, tok.end - tok.start, text, sugg, ta, en);

  // ---------------------------------------------------------------------------
  // Word-level rules (TA-ORTH, TA-WORD, TA-STYLE)
  // ---------------------------------------------------------------------------
  function checkWord(tok, text, opts, out) {
    const w = tok.text;
    const add = (rule, sev, sugg, ta, en, start, len) =>
      out.push(
        start === undefined
          ? tokIssue(rule, sev, tok, text, sugg, ta, en)
          : issue(rule, sev, tok.start + start, len, text, sugg, ta, en)
      );

    // Stray zero-width characters (pre-processing check)
    if (/[‌‍]/.test(w)) {
      add("TA-ORTH-000", "warning", [w.replace(/[‌‍]/g, "")],
        "சொல்லில் கண்ணுக்குத் தெரியாத எழுத்து (ZWJ/ZWNJ) உள்ளது.",
        "The word contains an invisible joiner character (ZWJ/ZWNJ), common in pasted or OCR'd text.");
    }
    const clean = w.replace(/[‌‍]/g, "");
    if (clean !== clean.normalize("NFC")) {
      add("TA-ORTH-000", "warning", [clean.normalize("NFC")],
        "உயிர்க்குறிகள் பிரிந்து தட்டச்சு செய்யப்பட்டுள்ளன.",
        "Vowel signs are stored decomposed; normalise to NFC.");
    }

    const first = clean[0];
    const loan = hasGrantha(clean);

    // TA-ORTH-001: word starting with a vowel sign or virama
    if (isSign(first)) {
      add("TA-ORTH-001", "error", [],
        "சொல் உயிர்க்குறியோடு அல்லது புள்ளியோடு தொடங்க முடியாது.",
        "A word cannot start with a vowel sign or a bare virama — likely a typing or encoding error.");
      return; // nothing else is reliable for this token
    }

    // TA-ORTH-002: letters that never start a native word
    if ("ஙணழளறன".includes(first) || (first === "ட" && !LEX.LOAN_INITIAL_OK.has(clean) && !loan)) {
      const fix = first === "ன" || first === "ண" ? ["ந" + clean.slice(1)] : [];
      add("TA-ORTH-002", first === "ட" ? "warning" : "error", fix,
        `"${first}" தமிழ்ச் சொல்லின் முதலில் வராது.`,
        `"${first}" never starts a native Tamil word.`, 0, 1);
    }

    // TA-ORTH-003: ர / ல at the start (helping vowel)
    if ((first === "ர" || first === "ல") && opts.formal) {
      add("TA-ORTH-003", "suggestion", ["இ" + clean],
        `செந்தமிழில் "${first}" சொல்லின் முதலில் வராது; முன் "இ" சேர்க்கலாம்.`,
        `In formal Tamil "${first}" does not start a word; add a helping vowel (e.g. இராமன், இலண்டன்).`, 0, 1);
    }

    // TA-SANDHI-041: an independent vowel inside a word
    for (let i = 1; i < clean.length; i++) {
      if (isIndVowel(clean[i]) && clean[i - 1] !== AYTHAM) {
        const L = clean.slice(0, i), V = clean[i], R = clean.slice(i + 1);
        const sugg = [L + " " + V + R];
        const last = L[L.length - 1];
        const sign = VOWEL_TO_SIGN[V];
        if (last === VIRAMA) {
          // மெய் முன் உயிர்: the vowel joins the consonant
          sugg.push(L.slice(0, -1) + sign + R);
        } else if (last === "ு" && "கசடதபற".includes(L[L.length - 2]) && L.length > 3) {
          // குற்றியலுகரம் drops before a vowel (காசு + இல்லை → காசில்லை)
          sugg.push(L.slice(0, -1) + sign + R);
        } else {
          const glide = "ிீை".includes(last) || isIndVowel(last) && "இஈஐ".includes(last) ? "ய" : last === "ே" ? "ய" : "வ";
          sugg.push(L + glide + sign + R);
        }
        add("TA-SANDHI-041", "error", sugg,
          "உயிர் எழுத்து சொல்லின் நடுவில் வராது. பிரித்து எழுதுக, அல்லது உடம்படுமெய் சேர்த்துப் புணர்த்துக.",
          "An independent vowel cannot appear inside a word. Split the words, or join them with the glide (ய்/வ்).");
        break;
      }
    }

    // TA-ORTH-005: nasal + stop must be a matching pair (இன எழுத்து)
    const nasalPairs = [
      [/ன்ட/g, "ண்ட"], [/ந்ட/g, "ண்ட"], [/ண்ற/g, "ன்ற"], [/ந்ற/g, "ன்ற"],
      [/ன்த/g, "ந்த"], [/ண்த/g, "ந்த"],
    ];
    for (const [re, rep] of nasalPairs) {
      re.lastIndex = 0;
      let m;
      while ((m = re.exec(clean)) !== null) {
        add("TA-ORTH-005", "error", [clean.slice(0, m.index) + rep + clean.slice(m.index + 3)],
          `இன எழுத்து: "${m[0]}" தவறு; "${rep}" என்று எழுதுக (ங்க, ஞ்ச, ண்ட, ந்த, ம்ப, ன்ற).`,
          `Nasal–stop pairs must match: write "${rep}", not "${m[0]}" (ங்க ஞ்ச ண்ட ந்த ம்ப ன்ற).`);
      }
    }

    // TA-ORTH-006: ங் only before க/ங; ஞ் only before ச/ஞ/ய; ந் only before த/ந
    if (!loan) {
      for (let i = 0; i < clean.length - 1; i++) {
        if (clean[i + 1] !== VIRAMA) continue;
        const c = clean[i], nx = clean[i + 2];
        let bad = false, fix = [];
        if (c === "ங" && nx !== undefined && !"கங".includes(nx)) bad = true;
        if (c === "ஞ" && nx !== undefined && !"சஞய".includes(nx)) bad = true;
        if (c === "ந" && nx !== undefined && !"தந".includes(nx)) bad = true;
        if (c === "ந" && nx === undefined) {
          bad = true;
          fix = [clean.slice(0, -2) + "ன்"];
        }
        if (bad) {
          add("TA-ORTH-006", "error", fix,
            `"${c}்" இந்த இடத்தில் வராது${nx === undefined && c === "ந" ? "; சொல் இறுதியில் \"ன்\" வரும்" : ""}.`,
            `"${c}்" cannot appear here${nx === undefined && c === "ந" ? "; a word ends in ன், not ந்" : ""}.`);
          break;
        }
      }
    }

    // TA-ORTH-007: ட்ற / ற்ட never occur
    if (/ட்ற|ற்ட/.test(clean)) {
      add("TA-ORTH-007", "error", [clean.replace(/ட்ற/g, "ற்ற").replace(/ற்ட/g, "ட்ட")],
        "ட்ற, ற்ட என்ற கூட்டெழுத்துகள் தமிழில் இல்லை.",
        "The clusters ட்ற and ற்ட do not occur in Tamil.");
    }

    // TA-ORTH-008: three consonants in a row (except ய்/ர்/ழ் + a doubled stop)
    if (!loan) {
      const re = /([கஙசஞடணதநபமயரலவழளறன])்([கஙசஞடணதநபமயரலவழளறன])்([கஙசஞடணதநபமயரலவழளறன])/g;
      let m;
      while ((m = re.exec(clean)) !== null) {
        // ய்/ர்/ழ் + a doubled stop (வாழ்க்கை) or + a nasal–stop pair (அமர்ந்தது, சேர்ந்து)
        const pair = { "ங": "க", "ஞ": "ச", "ண": "ட", "ந": "த", "ம": "ப", "ன": "ற" };
        const ok = "யரழ".includes(m[1]) && (m[2] === m[3] || pair[m[2]] === m[3]);
        if (!ok) {
          // the usual mistake is a doubled stop after ற்/ட்/ர் etc.: முயற்ச்சி → முயற்சி
          const fix = m[2] === m[3] ? [clean.slice(0, m.index + 2) + clean.slice(m.index + 4)] : [];
          add("TA-ORTH-008", "error", fix,
            "மூன்று மெய்யெழுத்துகள் தொடர்ந்து வர முடியாது (ய், ர், ழ் பின் இரட்டும் வல்லினம் தவிர).",
            "Three consonants cannot occur in a row (except ய்/ர்/ழ் + a doubled stop, as in வாழ்க்கை).");
          break;
        }
        re.lastIndex = m.index + 2;
      }
    }

    // TA-ORTH-009: dental ந inside a word
    if (!loan) {
      for (let i = 1; i < clean.length; i++) {
        if (clean[i] === "ந" && clean[i + 1] !== VIRAMA && isCons(clean[i - 1])) {
          add("TA-ORTH-009", "warning", ["ன"],
            "சொல்லின் நடுவில் பொதுவாக \"ன\" வரும்; \"ந\" சொல் முதலிலும் ந்த, ந்ந ஆகியவற்றிலும் வரும்.",
            "Inside a word Tamil normally uses ன; ந appears word-initially and in ந்த/ந்ந (compounds excepted).",
            i, 1);
          break;
        }
      }
    }

    // TA-ORTH-010: ஆய்தம் only between a short vowel and a hard consonant
    const ai = clean.indexOf(AYTHAM);
    if (ai >= 0) {
      const before = clean[ai - 1], after = clean[ai + 1];
      const okBefore = ai === 0 ? false : (isCons(before) || SHORT_VOWELS.includes(before) || "ிுெொ".includes(before));
      const okAfter = after !== undefined && "கசடதபற".includes(after);
      if (!(okBefore && okAfter) && !(after === "ப")) {
        add("TA-ORTH-010", "error", [],
          "ஆய்த எழுத்து (ஃ) குறிலுக்கும் வல்லினத்துக்கும் இடையில் மட்டுமே வரும் (அஃது, எஃகு).",
          "ஃ stands only between a short vowel and a hard consonant (அஃது, எஃகு).");
      }
    }

    // TA-ORTH-011 / SANDHI-045 / WORD-010: known misspellings
    if (LEX.MISSPELLINGS[clean]) {
      add("TA-ORTH-011", "error", [LEX.MISSPELLINGS[clean]],
        `எழுத்துப் பிழை: "${LEX.MISSPELLINGS[clean]}" என்று எழுதுக.`,
        `Spelling: write "${LEX.MISSPELLINGS[clean]}".`);
    }

    // TA-WORD-001: oblique stem (வீடுக்கு → வீட்டுக்கு, மரமுக்கு → மரத்துக்கு)
    for (const base in LEX.OBLIQUE_DOUBLING) {
      const obl = LEX.OBLIQUE_DOUBLING[base];
      const stem = base.slice(0, -1); // drop ு
      let fix = null;
      if (clean === base + "க்கு") fix = obl + "க்கு";
      else if (clean.startsWith(stem) && !clean.startsWith(obl)) {
        const rest = clean.slice(stem.length);
        const m = /^(ை|ில்|ின்|ால்|ோடு|ிலிருந்து|ுடைய)(.*)$/.exec(rest);
        // காடை, ஆடை, கூடை … are real words, so -ஐ is only checked for safe stems
        if (m && m[1] === "ை" && !LEX.OBLIQUE_ACC_SAFE.has(base)) continue;
        if (m) fix = obl.slice(0, -1) + m[1] + m[2];
      }
      if (fix) {
        add("TA-WORD-001", "error", [fix],
          `வேற்றுமை உருபு ஏற்கும்போது "${base}" என்பது "${obl}" என மாறும்.`,
          `With a case ending, "${base}" takes the oblique stem "${obl}".`);
        break;
      }
    }
    {
      const m = /^(.{2,})ம(ுக்கு|ை|ில்|ின்|ால்|ோடு)$/.exec(clean);
      if (m && /[^்]$/.test(m[1]) && !clean.endsWith("மை")) {
        add("TA-WORD-001", "warning", [m[1] + "த்த" + m[2]],
          "-அம் ஈற்றுப் பெயர் வேற்றுமை ஏற்கும்போது \"அத்து\" சாரியை பெறும் (மரத்தை, மரத்தில்).",
          "Nouns in -அம் take the அத்து stem before a case ending (மரத்தை, மரத்தில்).");
      }
    }

    // TA-WORD-002: locative of persons uses இடம்
    if (LEX.LOCATIVE_PERSON[clean]) {
      add("TA-WORD-002", "warning", [LEX.LOCATIVE_PERSON[clean]],
        "உயர்திணைப் பெயருக்கு ஏழாம் வேற்றுமையில் \"இடம்\" வரும்.",
        "With people, the locative is இடம், not இல்.");
    }

    // TA-WORD-004: double plural
    if (/கள்கள்$/.test(clean)) {
      add("TA-WORD-004", "error", [clean.replace(/கள்கள்$/, "கள்")],
        "பன்மை விகுதி இருமுறை வந்துள்ளது.", "The plural suffix -கள் appears twice.");
    }

    // TA-STYLE rules (lists)
    const style = (rule, map, ta, en, sev) => {
      if (map[clean]) add(rule, sev || "suggestion", map[clean], ta, en);
    };
    if (opts.style !== false) {
      style("TA-STYLE-001", LEX.STYLE_WORDS,
        "இலக்கண நூல்கள் பரிந்துரைக்கும் சொல்லைப் பயன்படுத்துக.",
        "The grammar books recommend a standard word here.");
      if (opts.formal) {
        style("TA-STYLE-002", LEX.SPOKEN,
          "பேச்சு வழக்குச் சொல்; எழுத்து வழக்கில் மாற்றுக.",
          "Spoken (colloquial) form; use the written form in formal text.");
        // productive spoken patterns
        let m;
        if ((m = /^(.+(?:ந்த|த்த|ட்ட|ய்த|ன்ற|ண்ட|க்க|ச்ச|ப்ப|ற்ற))ுட்ட(ான்|ாள்|ார்|ேன்|ோம்|ாய்|து|ாங்க)$/.exec(clean))) {
          add("TA-STYLE-002", "suggestion", [m[1] + "ுவிட்ட" + (m[2] === "ாங்க" ? "ார்கள்" : m[2])],
            "பேச்சு வழக்கு: \"-விட்ட-\" என்று எழுதுக.", "Spoken form: write -விட்ட- in formal text.");
        } else if ((m = /^(.+(?:ந்த|ன|ட்ட|த்த|ன்ன|ன்ற|ண்ட))ாங்க$/.exec(clean))) {
          add("TA-STYLE-002", "suggestion", [m[1] + "ார்கள்"],
            "பேச்சு வழக்கு: \"-ஆர்கள்\" என்று எழுதுக.", "Spoken form: write -ஆர்கள் in formal text.");
        }
        if (LEX.CLIPPED_NAMES.has(clean)) {
          add("TA-STYLE-003", "suggestion", [],
            "முறைசார் எழுத்தில் முழுப் பெயரை எழுதுக.", "Use the full name in formal writing.");
        }
        style("TA-STYLE-004", LEX.LOANWORDS,
          "பிறமொழிச் சொல்லுக்கு இணையான தமிழ்ச் சொல் உள்ளது.",
          "A Tamil equivalent exists for this loanword.");
      }
      style("TA-STYLE-005", LEX.MARABU,
        "மரபுச் சொல்லைப் பயன்படுத்துக.", "Use the conventional (மரபு) word.");
    }
  }

  // ---------------------------------------------------------------------------
  // Pair rules (TA-SANDHI, TA-ORTH-004)
  // ---------------------------------------------------------------------------
  // infinitives can end in ட்ட (கட்ட) or ிய (அறிய), so those stay warnings
  const REL_PART_STRONG = /(ான|கிற|கின்ற|ந்த|த்த|ற்ற|ன்ற|ண்ட|ன்ன)$/;
  const REL_PART_WEAK = /(ட்ட|ிய)$/;
  const VERB_LIKE = /(ான்|ாள்|ார்|ார்கள்|ேன்|ோம்|ாய்|ீர்கள்|ீர்|னர்|ந்தது|த்தது|ட்டது|கிறது|கின்றது|ந்தன|த்தன|கின்றன|ும்|ு|ி)$/;

  /** Is doubling REQUIRED after word w (before a hard-initial word)? */
  function requiredDoubling(w, n) {
    if (LEX.DEMONSTRATIVES.has(w)) return ["TA-SANDHI-001", "error", "சுட்டுச் சொல்லின் பின் வல்லினம் மிகும் (அந்தக் காகம்).", "Doubling follows a demonstrative (அந்தக் காகம்)."];
    if (LEX.DEMONSTRATIVES_OPTIONAL.has(w)) return ["TA-SANDHI-001", "suggestion", "இவ்விடத்தில் வல்லினம் மிகலாம் (அப்படிச் செய்தான்).", "Doubling is usual here (அப்படிச் செய்தான்)."];
    if (LEX.PRONOUN_ACC.has(w) || (w.length > 3 && w.endsWith("களை")))
      return ["TA-SANDHI-002", "error", "இரண்டாம் வேற்றுமை உருபு \"ஐ\"யின் பின் வல்லினம் மிகும் (அவனைக் கண்டேன்).", "Doubling follows the object ending -ஐ (அவனைக் கண்டேன்)."];
    if (!LEX.ACC_EXCEPTIONS.has(w)) {
      if (w.length > 2 && w.endsWith("யை"))
        return ["TA-SANDHI-002", "error", "இரண்டாம் வேற்றுமை உருபு \"ஐ\"யின் பின் வல்லினம் மிகும் (மணியைப் பார்).", "Doubling follows the object ending -ஐ (மணியைப் பார்)."];
      if (w.length > 3 && w.endsWith("த்தை"))
        return ["TA-SANDHI-002", "warning", "இது இரண்டாம் வேற்றுமை என்றால் வல்லினம் மிகும் (மரத்தைப் பார்).", "If this is an object (-ஐ), doubling follows (மரத்தைப் பார்)."];
      if (w.length > 3 && /[ாூோ]வை$/.test(w))
        return ["TA-SANDHI-002", "warning", "இது இரண்டாம் வேற்றுமை என்றால் வல்லினம் மிகும் (அம்மாவைக் கேள்).", "If this is an object (-ஐ), doubling follows (அம்மாவைக் கேள்)."];
    }
    if (/(ுக்கு|ற்கு|க்குக்கு)$/.test(w) && w !== "இருக்கு")
      return ["TA-SANDHI-003", "error", "நான்காம் வேற்றுமை உருபு \"கு\"வின் பின் வல்லினம் மிகும் (அவனுக்குச் சோறு).", "Doubling follows the dative ending -கு (அவனுக்குச் சோறு)."];
    if (LEX.DOUBLING_AFTER.has(w) || (w.length > 2 && /(ாக|ாகி|வகை|வித|பற்றி)$/.test(w)))
      return ["TA-SANDHI-004", "error", "இச்சொல்லின் பின் வல்லினம் மிகும் (போய்ப் பார், மெதுவாகப் பேசு, எல்லாப் பைகளும்).", "Doubling follows this word (போய்ப் பார், மெதுவாகப் பேசு, எல்லாப் பைகளும்)."];
    if (w.length > 2 && /(க்க|ய்ய|ட்க|ற்க)$/.test(w) && !LEX.INFINITIVE_EXCEPTIONS.has(w))
      return ["TA-SANDHI-005", "error", "செய என் எச்சத்தின் பின் வல்லினம் மிகும் (பார்க்கப் போனான்).", "Doubling follows an infinitive (பார்க்கப் போனான்)."];
    // skipped before a verb (object + verb, light verbs like தட்டச்சு செய்), where modern usage does not double
    if (w.length > 2 && /(க்கு|ட்டு|த்து|ப்பு|ற்று|ச்சு)$/.test(w) && w !== "எட்டு" && w !== "இருக்கு" && !(n && (VERB_LIKE.test(n) || n.startsWith("செய்"))))
      return ["TA-SANDHI-008", "warning", "வன்றொடர்க் குற்றியலுகரத்தின் பின் வல்லினம் மிகும் (எழுத்துப் பிழை).", "Doubling usually follows a word ending in a hard cluster + உ (எழுத்துப் பிழை)."];
    return null;
  }

  /** Is doubling FORBIDDEN after base word b (before next word n)? */
  function forbiddenDoubling(b, n) {
    if (LEX.NO_DOUBLING_AFTER.has(b)) {
      const id = /^(ஒரு|இரு|ஈர்|ஒவ்வொரு|இரண்டு|மூன்று|நான்கு|ஐந்து|ஆறு|ஏழு|எட்டு|ஒன்பது)$/.test(b) ? "TA-SANDHI-025"
        : /^(பல|சில)$/.test(b) ? "TA-SANDHI-026"
        : /^(அத்தனை|இத்தனை|எத்தனை|அவ்வளவு|இவ்வளவு|எவ்வளவு)$/.test(b) ? "TA-SANDHI-027"
        : /^(என்று|என்ன)$/.test(b) ? "TA-SANDHI-028" : "TA-SANDHI-021";
      return [id, "error", "இச்சொல்லின் பின் வல்லினம் மிகாது.", "No doubling after this word."];
    }
    if (LEX.ADJ_NO_DOUBLING.has(b) || (REL_PART_STRONG.test(b) && b.length > 2 && !LEX.REL_PART_EXCEPTIONS.has(b) && !LEX.DEMONSTRATIVES.has(b)))
      return ["TA-SANDHI-022", "error", "பெயரெச்சத்தின் பின் வல்லினம் மிகாது (படித்த பையன், அழகான பெண்).", "No doubling after a relative participle or -ஆன adjective (படித்த பையன், அழகான பெண்)."];
    if (LEX.OPTATIVES.has(b))
      return ["TA-SANDHI-023", "error", "வியங்கோள் வினைமுற்றின் பின் வல்லினம் மிகாது (வாழ்க தமிழ்).", "No doubling after a finite verb (வாழ்க தமிழ்)."];
    if (LEX.SOFT_PARTICIPLES.has(b) || (/(ந்து|ய்து)$/.test(b) && VERB_LIKE.test(n)))
      return ["TA-SANDHI-028", "error", "மென்றொடர் வினையெச்சத்தின் பின் வல்லினம் மிகாது (வந்து போனான்).", "No doubling after a verbal participle like வந்து, செய்து (வந்து போனான்)."];
    if (b.endsWith("ும்படி"))
      return ["TA-SANDHI-029", "warning", "வினையை அடுத்த \"படி\"யின் பின் வல்லினம் மிகாது (வரும்படி சொன்னான்).", "No doubling after verb + படி (வரும்படி சொன்னான்)."];
    if (REL_PART_WEAK.test(b) && b.length > 2 && !LEX.REL_PART_EXCEPTIONS.has(b) && !LEX.DEMONSTRATIVES.has(b))
      return ["TA-SANDHI-022", "warning", "இது பெயரெச்சம் என்றால் வல்லினம் மிகாது (வந்த குதிரை).", "If this is a relative participle, no doubling follows (வந்த குதிரை)."];
    return null;
  }

  function checkPairs(tokens, text, out) {
    for (let i = 0; i < tokens.length; i++) {
      const a = tokens[i];
      if (!a.tamil) continue;
      const aw = a.text.replace(/[‌‍]/g, "");
      const b = tokens[i + 1];
      const joined = b && b.tamil && /^[ \t ]+$/.test(text.slice(a.end, b.start));
      const d = trailingDoubling(aw);

      if (d && !hasGrantha(aw)) {
        const bw = joined ? b.text : "";
        const nh = joined ? hardInitial(bw) : null;
        if (!nh) {
          // TA-ORTH-004 / TA-SANDHI-010: a doubling consonant with nothing to double
          out.push(tokIssue(joined ? "TA-SANDHI-010" : "TA-ORTH-004", "error", a, text, [d.base],
            "அடுத்த சொல் வல்லினத்தில் தொடங்காததால் இந்த ஒற்று தேவையில்லை.",
            "This doubling consonant has no matching hard consonant after it; remove it."));
        } else if (nh !== d.cons) {
          out.push(tokIssue("TA-SANDHI-010", "error", a, text, [d.base + nh + VIRAMA],
            `மிகும் ஒற்று அடுத்த சொல்லின் முதல் எழுத்தோடு ஒத்திருக்க வேண்டும்: "${nh}்".`,
            `The doubling consonant must match the next word's first consonant: "${nh}்".`));
        } else {
          const f = forbiddenDoubling(d.base, bw);
          if (f) out.push(tokIssue(f[0], f[1], a, text, [d.base], f[2], f[3]));
        }
        continue;
      }

      if (!joined) continue;
      const bw = b.text;
      const nh = hardInitial(bw);

      // TA-SANDHI-040: ஒரு / ஓர்
      if (aw === "ஒரு" && isIndVowel(bw[0])) {
        out.push(tokIssue("TA-SANDHI-040", "warning", a, text, ["ஓர்"],
          "உயிர் முதல் சொல்லுக்கு முன் \"ஓர்\" வரும் (ஓர் ஊர்).", "Use ஓர் before a word starting with a vowel (ஓர் ஊர்)."));
      } else if (aw === "ஓர்" && isCons(bw[0])) {
        out.push(tokIssue("TA-SANDHI-040", "error", a, text, ["ஒரு"],
          "மெய் முதல் சொல்லுக்கு முன் \"ஒரு\" வரும் (ஒரு மரம்).", "Use ஒரு before a word starting with a consonant (ஒரு மரம்)."));
      }

      if (nh) {
        const r = requiredDoubling(aw, bw);
        if (r) out.push(tokIssue(r[0], r[1], a, text, [aw + nh + VIRAMA], r[2], r[3]));
      }
    }
  }

  // ---------------------------------------------------------------------------
  // Sentence rules (TA-AGR, TA-SENT, TA-PUNCT)
  // ---------------------------------------------------------------------------
  // verb endings, longest first, with their person/gender/number class
  const VERB_ENDINGS = [
    ["ார்கள்", "3hp"], ["ீர்கள்", "2p"], ["னர்", "3hp"], ["ேன்", "1s"], ["ோம்", "1p"],
    ["ாய்", "2s"], ["ீர்", "2p"], ["ான்", "3m"], ["ாள்", "3f"], ["ார்", "3h"],
    ["து", "3n"], ["ன", "3np"],
  ];
  const TENSE_BEFORE = /(கிற|கின்ற|ந்த|த்த|ட்ட|ற்ற|ன்ற|ண்ட|ன்ன|ின|ய்த|ன|ப்ப|வ|ப)$/;
  const PAST_MARK = /(ந்த|த்த|ட்ட|ற்ற|ன்ற|ண்ட|ன்ன|ின|ய்த)$/;
  const FUT_MARK = /(ப்ப|வ)$/;
  const PRES_MARK = /(கிற|கின்ற)$/;
  const TARGET_ENDING = { "1s": "ேன்", "1p": "ோம்", "2s": "ாய்", "2p": "ீர்கள்", "3m": "ான்", "3f": "ாள்", "3h": "ார்", "3hp": "ார்கள்", "3n": "து", "3np": "ன" };
  const COMPATIBLE = {
    "1s": ["1s"], "1p": ["1p"], "2s": ["2s"], "2p": ["2p"],
    "3m": ["3m"], "3f": ["3f"], "3h": ["3h", "3hp"], "3hp": ["3hp", "3h"],
    "3n": ["3n"], "3np": ["3np"], "PL": ["3hp", "3h", "3np"],
  };
  const CLASS_TA = { "1s": "தன்மை ஒருமை", "1p": "தன்மைப் பன்மை", "2s": "முன்னிலை ஒருமை", "2p": "முன்னிலைப் பன்மை", "3m": "ஆண்பால்", "3f": "பெண்பால்", "3h": "பலர்பால் (மரியாதை)", "3hp": "பலர்பால்", "3n": "ஒன்றன்பால்", "3np": "பலவின்பால்", "PL": "பன்மை" };

  function analyseVerb(w) {
    for (const [end, cls] of VERB_ENDINGS) {
      if (w.endsWith(end) && w.length > end.length + 1) {
        const stem = w.slice(0, -end.length);
        if (TENSE_BEFORE.test(stem)) return { stem, end, cls };
      }
    }
    return null;
  }
  function inflect(stem, cls) {
    let s = stem;
    if (cls === "3np" && s.endsWith("கிற")) s = s.slice(0, -3) + "கின்ற";
    return s + TARGET_ENDING[cls];
  }

  function checkSentence(sent, text, opts, out) {
    const words = sent.filter((t) => t.tamil);
    if (!words.length) return;
    const last = sent[sent.length - 1];
    const terminator = last.punct && /[.?!]/.test(last.text) ? last : null;
    const verbTok = words[words.length - 1];
    const vw = verbTok.text.replace(/[‌‍]/g, "");
    const verb = analyseVerb(vw);

    // TA-AGR-001..005: subject–verb agreement
    if (verb && words.length >= 2) {
      const subjTok = words[0];
      const sw = subjTok.text;
      let subj = LEX.SUBJECTS[sw] || null;
      let rule = null;
      if (!subj && sw.length > 3 && sw.endsWith("கள்") && !LEX.POSSESSIVE_PLURALS.has(sw)) subj = "PL";
      // coordinated subjects: X-உம் Y-உம் → plural
      if (!subj && words.length >= 3 && words[0].text.endsWith("ும்") && words[1].text.endsWith("ும்") && !VERB_LIKE.test(words[0].text.slice(0, -3))) {
        subj = "PL"; rule = "TA-AGR-005";
      }
      if (subj && !COMPATIBLE[subj].includes(verb.cls)) {
        // pick the rule by the kind of mismatch
        if (!rule) {
          const human = (c) => ["3m", "3f", "3h", "3hp"].includes(c);
          const person = (c) => c[0];
          if (subj === "PL") rule = "TA-AGR-003";
          else if (person(subj) !== person(verb.cls) || person(subj) !== "3") rule = "TA-AGR-004";
          else if (human(subj) !== human(verb.cls)) rule = "TA-AGR-001";
          else if (subj === "3np" || verb.cls === "3np") rule = "TA-AGR-003";
          else rule = "TA-AGR-002";
        }
        let targets = subj === "PL" ? (rule === "TA-AGR-005" ? ["3hp"] : ["3np", "3hp"]) : [subj];
        const sugg = targets.map((c) => inflect(verb.stem, c));
        const sev = rule === "TA-AGR-003" && subj === "3np" ? "warning" : (subj === "PL" ? "warning" : "error");
        out.push(tokIssue(rule, sev, verbTok, text, sugg,
          `எழுவாய் "${sw}" (${CLASS_TA[subj]}) — பயனிலை "${vw}" (${CLASS_TA[verb.cls]}) பொருந்தவில்லை.`,
          `Subject "${sw}" and verb "${vw}" do not agree in person, gender or number.`));
      }

      // TA-AGR-006: time word vs tense
      const timePast = words.some((t) => LEX.PAST_TIME.has(t.text));
      const timeFut = words.some((t) => LEX.FUTURE_TIME.has(t.text));
      if (timeFut && PAST_MARK.test(verb.stem)) {
        out.push(tokIssue("TA-AGR-006", "warning", verbTok, text, [],
          "கால வழு: எதிர்காலச் சொல்லுடன் இறந்தகால வினை வந்துள்ளது.", "Time word points to the future but the verb is past tense."));
      } else if (timePast && (FUT_MARK.test(verb.stem) || PRES_MARK.test(verb.stem))) {
        out.push(tokIssue("TA-AGR-006", "warning", verbTok, text, [],
          "கால வழு: இறந்தகாலச் சொல்லுடன் நிகழ்/எதிர்கால வினை வந்துள்ளது.", "Time word points to the past but the verb is not past tense."));
      }
    }

    // TA-SENT-001: sentence ending in a verbal participle
    if (terminator && terminator.text === "." && (LEX.SOFT_PARTICIPLES.has(vw) || /(ித்து|ய்து)$/.test(vw)) && words.length >= 2) {
      out.push(tokIssue("TA-SENT-001", "warning", verbTok, text, [],
        "வாக்கியம் வினையெச்சத்தில் முடிவது போல் உள்ளது; வினைமுற்றில் முடிக்கவும்.",
        "The sentence seems to end in a verbal participle (…வந்து.) — finish it with a finite verb."));
    }

    // TA-SENT-003: questions end with ?
    if (terminator && terminator.text === ".") {
      const qParticle = /(ானா|ாளா|ாரா|ார்களா|ேனா|ோமா|ாயா|ீர்களா|தா|னவா|ில்லையா|ாதா|ுமா)$/.test(vw);
      const qWord = LEX.QUESTION_WORDS.has(words[0].text);
      if (qParticle || qWord) {
        out.push(tokIssue("TA-SENT-003", "warning", terminator, text, ["?"],
          "இது வினா வாக்கியம்; இறுதியில் கேள்விக்குறி இடுக.", "This is a question; end it with a question mark."));
      }
    }

    // TA-SENT-006: very long sentences
    if (opts.style !== false && words.length > 30) {
      out.push(issue("TA-SENT-006", "suggestion", words[0].start, verbTok.end - words[0].start, text, [],
        "நீண்ட வாக்கியம்; தனி வாக்கியங்களாகப் பிரிக்கலாம்.", "Long sentence; consider splitting it into shorter sentences."));
    }
  }

  function checkPunctuation(text, tokens, out) {
    // TA-PUNCT-001: no space before punctuation; one space after
    let m;
    const before = /([^\s])([ \t ]+)([,.;:?!])/g;
    while ((m = before.exec(text)) !== null) {
      const s = m.index + m[1].length;
      out.push(issue("TA-PUNCT-001", "warning", s, m[2].length + 1, text, [m[3]],
        "நிறுத்தக் குறிக்கு முன் இடைவெளி வேண்டாம்.", "No space before a punctuation mark."));
    }
    const after = /([,;:?!])([஀-௿])/g;
    while ((m = after.exec(text)) !== null) {
      out.push(issue("TA-PUNCT-001", "warning", m.index, 2, text, [m[1] + " " + m[2]],
        "நிறுத்தக் குறிக்குப் பின் இடைவெளி விடுக.", "Add a space after the punctuation mark."));
    }
    const dot = /([஀-௿]{2,})\.([஀-௿]{2,})/g;
    while ((m = dot.exec(text)) !== null) {
      const s = m.index + m[1].length;
      out.push(issue("TA-PUNCT-001", "warning", s, 1, text, [". "],
        "முற்றுப்புள்ளிக்குப் பின் இடைவெளி விடுக.", "Add a space after the full stop."));
    }
    // TA-PUNCT-006: double spaces
    const dbl = /([^\s]) {2,}(?=[^\s])/g;
    while ((m = dbl.exec(text)) !== null) {
      const s = m.index + 1;
      out.push(issue("TA-PUNCT-006", "suggestion", s, m[0].length - 1, text, [" "],
        "இரட்டை இடைவெளி.", "Double space."));
    }
    // TA-PUNCT-002: paragraph should end with . ? ! (skip headings / short lines)
    const words = tokens.filter((t) => t.tamil);
    const lastTok = tokens[tokens.length - 1];
    if (words.length >= 6 && lastTok && !/[.?!:"'”’)\]]/.test(lastTok.text)) {
      out.push(issue("TA-PUNCT-002", "warning", lastTok.start, lastTok.end - lastTok.start, text, [lastTok.text + "."],
        "வாக்கியத்தின் இறுதியில் முற்றுப்புள்ளி இடுக.", "End the sentence with a full stop."));
    }
    // TA-PUNCT-004: balanced quotes
    const straight = (text.match(/"/g) || []).length;
    const open = (text.match(/[“]/g) || []).length, close = (text.match(/[”]/g) || []).length;
    if (straight % 2 === 1 || open !== close) {
      const idx = Math.max(text.lastIndexOf('"'), text.lastIndexOf("“"), text.lastIndexOf("”"));
      out.push(issue("TA-PUNCT-004", "warning", idx, 1, text, [],
        "மேற்கோள் குறிகள் சரியாக மூடப்படவில்லை.", "Quotation marks are not balanced."));
    }
    // TA-PUNCT-005: mixed numerals
    if (/[௦-௯]/.test(text) && /[0-9]/.test(text)) {
      const i = text.search(/[௦-௯]/);
      out.push(issue("TA-PUNCT-005", "suggestion", i, 1, text, [],
        "தமிழ் எண்களையும் அரபு எண்களையும் கலக்க வேண்டாம்.", "Don't mix Tamil and Arabic numerals."));
    }
  }

  // ---------------------------------------------------------------------------
  // Public API
  // ---------------------------------------------------------------------------
  const SEV_RANK = { error: 3, warning: 2, suggestion: 1 };

  function dedupe(issues) {
    issues.sort((a, b) => a.start - b.start || SEV_RANK[b.severity] - SEV_RANK[a.severity]);
    const out = [];
    for (const it of issues) {
      const clash = out.find((o) => o.start === it.start && o.length === it.length);
      if (clash) {
        if (SEV_RANK[it.severity] > SEV_RANK[clash.severity]) out[out.indexOf(clash)] = it;
        continue;
      }
      out.push(it);
    }
    return out;
  }

  /**
   * @param {string} text  one paragraph
   * @param {{formal?: boolean, style?: boolean, minSeverity?: string, disabled?: string[]}} options
   */
  function checkParagraph(text, options) {
    const opts = Object.assign({ formal: true, style: true, minSeverity: "suggestion", disabled: ["TA-ORTH-003"] }, options || {});
    if (!text || !/[஀-௿]/.test(text)) return [];
    const tokens = tokenize(text);
    const out = [];
    for (const t of tokens) if (t.tamil) checkWord(t, text, opts, out);
    checkPairs(tokens, text, out);
    for (const s of sentences(tokens)) checkSentence(s, text, opts, out);
    checkPunctuation(text, tokens, out);
    const min = SEV_RANK[opts.minSeverity] || 1;
    return dedupe(out.filter((i) => SEV_RANK[i.severity] >= min && !opts.disabled.includes(i.rule) && i.start >= 0));
  }

  /** Apply one suggestion to a paragraph string (used by tests and the task pane fallback). */
  function applyFix(text, issue, suggestion) {
    return text.slice(0, issue.start) + suggestion + text.slice(issue.start + issue.length);
  }

  const api = { checkParagraph, applyFix, tokenize, _internal: { hardInitial, trailingDoubling, analyseVerb, inflect } };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.TamilChecker = api;
})(typeof self !== "undefined" ? self : this);
