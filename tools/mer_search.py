"""mer_search.py - headless FATF MER keyword search (same matching engine as MER Search Tool v8).

Usage:
  python mer_search.py --r4 <round-4 folder or CSVs> --r5 <round-5 folder or CSVs> \
      --round both --terms terms.txt --out results.xlsx [--mode strict|exact|fuzzy] \
      [--caps-only] [--jurisdictions "Bahrain,Qatar"] [--tables "Effectiveness,Risk and Context"] \
      [--quotes paragraph|sentence]
Each round needs its two CSVs: one with RC_IO/KFRA in the name, one with TECH_COMP/TCOMP.
terms.txt: one term per line; ~ joins equivalent names; 'X (Y)' brackets handled as in the tool.
"""
import os, re, sys, glob, argparse, hashlib, unicodedata
from collections import Counter, defaultdict
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
try:
    from rapidfuzz import fuzz as _rf
    def _ratio(a, b): return _rf.ratio(a, b)
except ImportError:
    import difflib
    def _ratio(a, b): return difflib.SequenceMatcher(None, a, b).ratio() * 100

CONTEXT_CHARS = 250
ROUND_LABEL = {"4": "4th round", "5": "5th round"}
TABLES = ["Effectiveness", "Risk and Context", "Technical Compliance"]
ILLEGAL_XLSX = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]')

def strip_html(text):
    if pd.isna(text):
        return ""
    text = str(text)
    text = re.sub(r'<[^>]+>', ' ', text)
    text = text.replace('&nbsp;', ' ').replace('&amp;', '&')
    text = text.replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"')
    text = re.sub(r'&#\d+;', '', text)
    text = re.sub(r'&\w+;', '', text)
    text = ILLEGAL_XLSX.sub('', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def combine_rating(code, label):
    """Return the SHORT code only: 'LC' + 'Largely compliant' -> 'LC'; '_Z' -> 'NA'."""
    code = "" if pd.isna(code) else str(code).strip()
    label = "" if pd.isna(label) else str(label).strip()
    if code in ("", "_Z", "_X"):
        return "NA"
    return code or "NA"


def sentence_of(text, start, end):
    """The sentence containing the match (capped at ~500 chars)."""
    s = start
    while s > 0 and text[s-1] not in ".!?\u2022;\n" and start - s < 400:
        s -= 1
    e = end
    while e < len(text) and text[e] not in ".!?\u2022;\n" and e - end < 400:
        e += 1
    if e < len(text):
        e += 1
    return text[s:e].strip()


# --- accent-tolerant tokenising -------------------------------------------
_ACCENTS = {
    'a': 'a\u00e0\u00e1\u00e2\u00e3\u00e4\u00e5\u0101\u0103\u0105',
    'c': 'c\u00e7\u0107\u0109\u010d',
    'd': 'd\u010f\u0111',
    'e': 'e\u00e8\u00e9\u00ea\u00eb\u0113\u0115\u0117\u0119\u011b',
    'g': 'g\u011f\u0121\u0123',
    'i': 'i\u00ec\u00ed\u00ee\u00ef\u0129\u012b\u012d\u012f\u0131',
    'l': 'l\u013a\u013e\u0142',
    'n': 'n\u00f1\u0144\u0146\u0148',
    'o': 'o\u00f2\u00f3\u00f4\u00f5\u00f6\u00f8\u014d\u014f\u0151',
    'r': 'r\u0155\u0159',
    's': 's\u015b\u015d\u015f\u0161',
    't': 't\u0163\u0165\u0167',
    'u': 'u\u00f9\u00fa\u00fb\u00fc\u0169\u016b\u016d\u016f\u0171\u0173',
    'y': 'y\u00fd\u00ff',
    'z': 'z\u017a\u017c\u017e',
}

def deaccent(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(c))

def tokenize(term):
    """Unicode-aware, accent-stripped word tokens. 'Comite' == 'Comit\u00e9'."""
    return [deaccent(t).lower() for t in re.findall(r'[^\W_]+', str(term), re.UNICODE)]

def _char_class(ch):
    """Match a letter and all its accented variants, so 'comite' finds 'comit\u00e9'."""
    if ch in _ACCENTS:
        return '[' + _ACCENTS[ch] + ']'
    return re.escape(ch)

def core_of(term):
    """Case / space / punctuation / accent insensitive core regex for a phrase.
       'ship-based' == 'ship based' == 'ship.based' == 'shipbased'"""
    tokens = tokenize(term)
    if not tokens:
        return None
    parts = ["".join(_char_class(c) for c in t) for t in tokens]
    return r'[\W_]*'.join(parts)

# --- how a single alias should be matched ---------------------------------
BOUND_BEFORE = r'(?<![^\W_])'      # start of a word (unicode-safe)
BOUND_AFTER  = r'(?![^\W_])'       # end of a word

# single-token names that are also ordinary English words -> flagged, not blocked
COMMON_WORDS = {
    "frontier", "finest", "cosmic", "protect", "shadow", "guardian", "athena",
    "chameleon", "odin", "anton", "next", "generation", "forum", "platform",
    "exchange", "alliance", "partnership", "committee", "taskforce", "network",
    "summit", "roundtable", "section", "cooperation", "compliance",
}

def alias_rule(alias):
    """Decide how one alias is matched. Returns a dict describing the rule."""
    a = alias.strip()
    plus = bool(re.search(r'[+]\s*$', a))          # 'JMLIT+' / 'FRONTIER +'
    base = re.sub(r'\s*[+]\s*$', '', a)
    toks = tokenize(base)
    raw_single = re.findall(r'[^\W_]+', base, re.UNICODE)
    is_single = len(raw_single) == 1
    # ALL-CAPS single token (letters/digits) e.g. PPP, ASC, JMLIT, I4C
    acronym = bool(is_single and raw_single
                   and raw_single[0].isupper() and 2 <= len(raw_single[0]) <= 8
                   and any(c.isalpha() for c in raw_single[0]))
    ends_digit = bool(toks and toks[-1][-1].isdigit())
    if acronym:
        kind = "acronym"
    elif plus:
        kind = "plus"
    elif len(toks) > 1:
        kind = "phrase"
    elif ends_digit:
        kind = "number"
    else:
        kind = "word"
    return {"alias": a, "base": base, "plus": plus, "kind": kind,
            "acronym": acronym, "ends_digit": ends_digit,
            "common": bool(is_single and toks and toks[0] in COMMON_WORDS)}

def alias_regex(rule, mode, case_sensitive_acronyms):
    """Regex fragment for one alias, honouring its matching rule."""
    if rule["acronym"] and case_sensitive_acronyms:
        # exact capitals only: build the core from the ORIGINAL text
        toks = re.findall(r'[^\W_]+', rule["base"], re.UNICODE)
        core = r'[\W_]*'.join(re.escape(t) for t in toks)
        core = r'(?-i:' + core + r')'
    else:
        core = core_of(rule["base"])
    if not core:
        return None
    frag = core
    if rule["plus"]:                                # must be followed by + / plus
        frag += r'\s*(?:\+|plus' + BOUND_AFTER + r')'
    elif rule["acronym"]:
        frag += r"(?:['\u2019]?s)?(?:\s*\+)?"       # PPP -> PPPs, JMLIT -> JMLIT+
        frag += BOUND_AFTER
    elif rule["ends_digit"] or mode == "exact":
        frag += BOUND_AFTER
    return frag

def is_acronym_of(short, long_form):
    """True if `short` plausibly abbreviates `long_form`
       (its letters appear in order in the long form)."""
    s = [c for c in deaccent(short).lower() if c.isalpha()]
    L = [c for c in deaccent(long_form).lower() if c.isalpha()]
    if len(s) < 2 or len(s) > 12 or len(re.findall(r'[^\W_]+', long_form)) < 2:
        return False
    if len(s) >= len(L):
        return False
    i = 0
    for c in L:
        if i < len(s) and c == s[i]:
            i += 1
    return i == len(s)

def strip_wrappers(s):
    """Remove stray quotes / trailing punctuation a pasted list often carries."""
    s = str(s).strip()
    s = re.sub(r'^["\u201c\u201d\u2018\u2019\'`]+', '', s)
    s = re.sub(r'["\u201c\u201d\u2018\u2019\'`]+$', '', s)
    s = re.sub(r'[.,;]+$', '', s)
    return s.strip()

def split_group(group_str):
    """Split one input line into equivalent aliases on '~' or a standalone 'OR'."""
    s = strip_wrappers(group_str)
    parts = re.split(r'\s*~\s*|\s+OR\s+', s)
    return [strip_wrappers(p) for p in parts if strip_wrappers(p)]

# jurisdiction shorthands users type in brackets
JUR_ALIASES = {
    "uae": "United Arab Emirates", "uk": "United Kingdom",
    "us": "United States", "usa": "United States",
    "hk": "Hong Kong (China)", "nl": "Netherlands", "nz": "New Zealand",
    "rsa": "South Africa", "sa": "South Africa",
}

def match_jurisdiction(text, known):
    """Return the dataset's jurisdiction name if `text` names one."""
    t = strip_wrappers(text).lower()
    if not known:
        return None
    lookup = {k.lower(): k for k in known}
    if t in lookup:
        return lookup[t]
    if t in JUR_ALIASES and JUR_ALIASES[t] in known:
        return JUR_ALIASES[t]
    for k in known:                       # 'Hong Kong' -> 'Hong Kong (China)'
        base = re.sub(r'\s*\(.*\)\s*$', '', k).lower()
        if t == base:
            return k
    return None

def prepare_term(group_str, mode, case_sensitive_acronyms=False, known_jurisdictions=None):
    """Work out exactly how one input line will be searched.
       Returns the aliases, compiled pattern, any jurisdiction scope, and
       plain-English notes explaining every decision."""
    raw_aliases = split_group(group_str)
    notes, scope = [], None
    aliases, seen = [], set()

    def add(a):
        a = strip_wrappers(a)
        key = " ".join(tokenize(a))
        if key and key not in seen:
            seen.add(key)
            aliases.append(a)

    for a in raw_aliases:
        m = re.match(r'^(.+?)\s*[\(\[]\s*([^\(\)\[\]]+?)\s*[\)\]]\s*$', a)
        if not m:
            add(a)
            continue
        head, inner = m.group(1).strip(), m.group(2).strip()
        jur = match_jurisdiction(inner, known_jurisdictions)
        if is_acronym_of(head, inner) or is_acronym_of(inner, head):
            add(a); add(head); add(inner)      # JIG (Joint Intelligence Group)
            notes.append(f'"{inner}" treated as the same thing as "{head}" '
                         f'(one abbreviates the other), so either on its own also counts.')
        elif jur:
            add(head)                          # SAMLIT (South Africa)
            scope = jur
            notes.append(f'"{inner}" read as a country rather than part of the name: '
                         f'searching "{head}" and showing only {jur}.')
        else:
            add(a)                             # keep the whole thing literal
            notes.append(f'"{inner}" is not an abbreviation of "{head}" and is not a country, '
                         f'so the whole phrase is searched exactly as written. '
                         f'Use ~ if you also want "{head}" on its own.')

    if not aliases:
        return None

    rules = [alias_rule(a) for a in aliases]
    frags = [f for f in (alias_regex(r, mode, case_sensitive_acronyms) for r in rules) if f]
    pairs = []
    for r1 in rules:                            # "JIG (Joint Intelligence Group)" as one unit
        for r2 in rules:
            if r1["alias"] == r2["alias"]:
                continue
            c1, c2 = core_of(r1["base"]), core_of(r2["base"])
            if c1 and c2:
                pairs.append(c1 + r'\s*[\(\[]\s*' + c2 + r'\s*[\)\]]')
    alts = sorted(pairs + frags, key=len, reverse=True)
    pattern = re.compile(BOUND_BEFORE + r'(?:' + '|'.join(alts) + r')', re.IGNORECASE)

    # ---------- plain-English explanation + warnings ----------
    how, warn = [], []
    for r in rules:
        if r["plus"]:
            how.append(f'"{r["alias"]}" must be followed by "+" or "plus", so it stays '
                       f'separate from plain "{r["base"]}".')
        elif r["acronym"]:
            cs = "capitals only" if case_sensitive_acronyms else "any capitalisation"
            how.append(f'"{r["alias"]}" - whole word only, {cs}; also finds {r["base"]}s '
                       f'and {r["base"]}+. It will NOT match longer words that merely '
                       f'begin with those letters.')
        elif r["kind"] == "phrase":
            how.append(f'"{r["alias"]}" - these words must appear together in this order; '
                       f'spacing, punctuation and accents are ignored.')
        elif r["ends_digit"]:
            how.append(f'"{r["alias"]}" - exact number, will not match longer numbers.')
        else:
            how.append(f'"{r["alias"]}" - matches words starting with this, so plurals and '
                       f'longer endings are included.')
        if r["common"] and not (r["acronym"] and case_sensitive_acronyms):
            warn.append(f'"{r["alias"]}" is also an everyday word, so some hits may be '
                        f'ordinary prose rather than the body you mean. Check the '
                        f'"matched as" breakdown, or tick capitals-only.')
        if r["kind"] == "word" and len(r["base"]) <= 4 and not r["acronym"]:
            warn.append(f'"{r["alias"]}" is short and not in capitals, so it matches any word '
                        f'starting with those letters. Typing it in CAPITALS makes it exact.')
    if len(tokenize(aliases[0])) > 6:
        warn.append("Long phrase: it only matches if this exact wording appears in this "
                    "order. If the report words it differently you will get no hits - "
                    "consider searching a distinctive part instead.")
    if re.search(r'[\u00e8\u2192\u21d2]', str(group_str)):
        warn.append("This line contains unusual characters and may be two names run "
                    "together - check it and split with ~ if so.")

    return {"label": aliases[0], "aliases": aliases, "pattern": pattern,
            "tokens": [tokenize(a) for a in aliases],
            # folded tokens for the fast pre-filter (see norm_text)
            "ptokens": [[norm_text(t) for t in tokenize(a)] for a in aliases],
            "rules": rules,
            "scope": scope, "notes": notes, "how": how, "warn": warn,
            "input": str(group_str).strip()}

def fuzzy_find(text, tokens, threshold):
    """Word-window fuzzy match. Returns (start, end, matched_text) or None."""
    # unicode-aware words, compared accent-free: "Comité" is one word, ~ "comite"
    words = [(norm_text(m.group()), m.start(), m.end())
             for m in re.finditer(r'[^\W_]+', text)]
    if not words:
        return None
    tokens = [norm_text(t) for t in tokens]
    n = len(tokens)
    target = " ".join(tokens)
    best = None  # (score, start, end)
    sizes = sorted({max(1, n - 1), n, n + 1})
    for size in sizes:
        if size > len(words):
            continue
        for i in range(len(words) - size + 1):
            window = words[i:i + size]
            cand = " ".join(w[0] for w in window)
            if abs(len(cand) - len(target)) > max(4, len(target) * (1 - threshold / 100) * 2 + 3):
                continue
            score = _ratio(cand, target)
            if score >= threshold and (best is None or score > best[0]):
                best = (score, window[0][1], window[-1][2])
    if best is None:
        return None
    return best[1], best[2], text[best[1]:best[2]]

def find_match(text, tp, mode, threshold):
    """Return (start, end, [distinct matched strings]) or None.
       tp is a prepare_term() dict, so any alias in the group counts."""
    if mode in ("strict", "exact"):
        if tp["pattern"] is None:
            return None
        hits = list(tp["pattern"].finditer(text))
        if not hits:
            return None
        variants, seen = [], set()
        for m in hits:
            g = m.group()
            if g.lower() not in seen:
                seen.add(g.lower())
                variants.append(g)
        return hits[0].start(), hits[0].end(), variants
    # fuzzy: try every alias, keep the earliest match
    best = None
    for toks in tp["tokens"]:
        if not toks:
            continue
        r = fuzzy_find(text, toks, threshold)
        if r and (best is None or r[0] < best[0]):
            best = r
    if best is None:
        return None
    return best[0], best[1], [best[2]]

def make_context(text, start, end):
    s = max(0, start - CONTEXT_CHARS)
    e = min(len(text), end + CONTEXT_CHARS)
    snippet = text[s:e]
    if s > 0:
        snippet = "..." + snippet
    if e < len(text):
        snippet += "..."
    return snippet

def apply_filters(df, filters):
    sub = df
    f = filters or {}
    if f.get("jurisdictions"):
        sub = sub[sub["Jurisdiction"].isin(f["jurisdictions"])]
    if f.get("tables"):
        sub = sub[sub["Table"].isin(f["tables"])]
    if f.get("indicators"):
        sub = sub[sub["Indicator"].isin(f["indicators"])]
    if f.get("ratings"):
        sub = sub[sub["Rating"].isin(f["ratings"])]
    if f.get("crit_ratings"):
        sub = sub[sub["Criteria rating"].isin(f["crit_ratings"])]
    return sub


# --- fast pre-filter --------------------------------------------------------
# Every letter the search pattern treats as equivalent (é = e, ø = o ...) is
# folded to one form, so a cheap substring test can rule paragraphs out before
# the full pattern runs. It never rules out a paragraph the pattern would match.
_NORM_TABLE = {}
for _base, _variants in _ACCENTS.items():
    for _v in _variants:
        if _v != _base:
            _NORM_TABLE[ord(_v)] = _base
_NORM_TABLE.update({ord("ø"): "o", ord("đ"): "d", ord("ł"): "l",
                    ord("ı"): "i", ord("ſ"): "s", ord("ħ"): "h"})
for _cp in range(0x300, 0x370):          # combining accents
    _NORM_TABLE[_cp] = None

def norm_text(s):
    s = str(s).lower()
    return s if s.isascii() else s.translate(_NORM_TABLE)

def _measure_key(m):
    """Natural sort: IO2 before IO10, R3 before R22."""
    return [int(t) if t.isdigit() else t for t in re.split(r'(\d+)', str(m))]

_TABLE_ORDER = {t: i for i, t in enumerate(TABLES)}

def citation(r):
    """'Austria, 5th round MER, IO.4 (rated ME)'"""
    ref = f'{r["Jurisdiction"]}, {r["Round"]} MER, {pretty_measure(r["MEASURE"])}'
    if r.get("Rating") and r["Rating"] != "NA":
        ref += f' (rated {r["Rating"]})'
    if r.get("Criteria rating") and r["Criteria rating"] != "NA":
        ref += f', criterion {r["Criteria rating"]}'
    return ref

def row_id(r):
    """Stable id for one hit, so the user can exclude it before exporting."""
    key = "|".join(str(r.get(k, "")) for k in
                   ("Round", "File", "Jurisdiction", "Table", "MEASURE", "Indicator",
                    "Location"))
    key += "|" + str(r.get("OBS_VALUE") or r.get("Full text") or "")
    return hashlib.md5(key.encode("utf-8", "replace")).hexdigest()[:14]


def pretty_measure(m):
    """IO4CI5 -> IO.4 ; R22_C1 -> R.22 ; EXECSUM4 -> Executive summary"""
    m = str(m or "")
    mm = re.match(r'IO(\d+)', m)
    if mm: return "IO.%s" % mm.group(1)
    mm = re.match(r'R(\d+)_', m)
    if mm: return "R.%s" % mm.group(1)
    if m.startswith("EXECSUM"): return "Executive summary"
    if m.startswith("RC") or m.startswith("MLTF"): return "Risk and context"
    return m

# ----------------------------------------------------------------- loading
UNIFORM = ["Round", "Jurisdiction", "Table", "MEASURE", "Indicator", "Rating", "Criteria rating", "OBS_VALUE"]

def _pick(src, keys):
    files = [src] if os.path.isfile(src) else glob.glob(os.path.join(src, "*.csv"))
    files += glob.glob(os.path.join(os.path.dirname(src), "*.csv")) if os.path.isfile(src) else []
    for f in sorted(set(files)):
        if any(k in os.path.basename(f).lower() for k in keys):
            return f
    raise SystemExit(f"No CSV containing {keys} found in {src}")

def load_round(rn, src):
    kf = _pick(src, ["rc_io", "kfra"]); tf = _pick(src, ["tech_comp", "tcomp"])
    k = pd.read_csv(kf, encoding="utf-8-sig", dtype=str, usecols=["Jurisdiction", "MEASURE", "Indicator", "EFF_CAT", "Effectiveness rating", "OBS_VALUE"])
    k["OBS_VALUE"] = k["OBS_VALUE"].apply(strip_html); k = k[k["OBS_VALUE"] != ""].copy()
    k["Rating"] = [combine_rating(c, l) for c, l in zip(k["EFF_CAT"], k["Effectiveness rating"])]
    k["Criteria rating"] = "NA"
    io = k["MEASURE"].fillna("").str.startswith("IO")
    k["Table"] = ["Effectiveness" if x else "Risk and Context" for x in io]
    k.loc[~io, "Rating"] = "NA"
    t = pd.read_csv(tf, encoding="utf-8-sig", dtype=str, usecols=["Jurisdiction", "MEASURE", "Indicator", "COMP_RATING", "Technical compliance rating", "CRIT_RATING", "Criteria rating", "OBS_VALUE"])
    t = t.rename(columns={"Criteria rating": "_cl"})
    t["OBS_VALUE"] = t["OBS_VALUE"].apply(strip_html); t = t[t["OBS_VALUE"] != ""].copy()
    t["Rating"] = [combine_rating(c, l) for c, l in zip(t["COMP_RATING"], t["Technical compliance rating"])]
    t["Criteria rating"] = [combine_rating(c, l) for c, l in zip(t["CRIT_RATING"], t["_cl"])]
    t["Table"] = "Technical Compliance"
    k["Round"] = t["Round"] = ROUND_LABEL[rn]
    df = pd.concat([k[UNIFORM], t[UNIFORM]], ignore_index=True).fillna("")
    df["_norm"] = [norm_text(x) for x in df["OBS_VALUE"].tolist()]
    return df

# ----------------------------------------------------------------- search
def search(df, tp, mode, threshold, filters):
    sub = apply_filters(df, filters)
    cols = {c: sub[c].tolist() for c in UNIFORM}
    texts = cols["OBS_VALUE"]
    if mode in ("strict", "exact"):
        groups = [g for g in tp["ptokens"] if g]
        idx = [i for i, n in enumerate(sub["_norm"].tolist()) if any(all(x in n for x in g) for g in groups)]
    else:
        idx = range(len(texts))
    out = []
    for i in idx:
        hit = find_match(texts[i], tp, mode, threshold)
        if not hit: continue
        s, e, v = hit
        r = {c: cols[c][i] for c in UNIFORM}
        r.update(context=make_context(texts[i], s, e), sentence=sentence_of(texts[i], s, e), matched=v[0], variants=v)
        r["ref"] = citation(r)
        out.append(r)
    out.sort(key=lambda r: (_TABLE_ORDER.get(r["Table"], 9), r["Jurisdiction"], r["Round"], _measure_key(r["MEASURE"])))
    return out

def coverage(df_round, filters, res):
    scope = apply_filters(df_round, filters); total = scope["Jurisdiction"].nunique()
    c = Counter(r["Jurisdiction"] for r in res)
    return {"mentioned": len(c), "total": int(total), "pct": round(100 * len(c) / total, 1) if total else 0.0,
            "list": ", ".join(f"{k} ({v})" for k, v in c.most_common())}

# ----------------------------------------------------------------- excel
BAD = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]")
def _v(c, v):
    if isinstance(v, str):
        v = BAD.sub("", v)[:32000]; c.value = v
        if v.startswith("="): c.data_type = "s"
    else: c.value = v

def _bold(text, variants):
    text = BAD.sub("", str(text))[:32000]
    vs = sorted({x for x in variants if x}, key=len, reverse=True)
    if not vs: return text
    pat = re.compile("|".join(re.escape(x) for x in vs), re.I); parts, last = [], 0
    for m in pat.finditer(text):
        if m.start() > last: parts.append(text[last:m.start()])
        parts.append(TextBlock(InlineFont(b=True), m.group())); last = m.end()
    if last < len(text): parts.append(text[last:])
    return CellRichText(*parts) if len(parts) > 1 or not isinstance(parts[0], str) else text

def _table(ws, name, hr, last, n):
    if last > hr:
        t = Table(displayName=name, ref=f"A{hr}:{get_column_letter(n)}{last}")
        t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True); ws.add_table(t)

def _head(ws, r, heads, widths):
    for j, h in enumerate(heads, 1):
        c = ws.cell(row=r, column=j, value=h); c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1E2F4D"); c.alignment = Alignment(wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = widths[j - 1]

def export(blocks, labels, path, quotes):
    wb = Workbook(); ws = wb.active; ws.title = "Summary"
    ws["A1"] = "FATF MER keyword search - rounds are reported separately, never added together"; ws["A1"].font = Font(bold=True, size=12)
    heads = ["Search term", "Round", "Total hits", "Effectiveness", "Risk and Context", "Technical Compliance",
             "Jurisdictions mentioned", "Jurisdictions in scope", "Coverage %", "Jurisdictions list", "Equivalent terms", "Matched in text as", "Warnings"]
    _head(ws, 3, heads, [28, 11, 10, 12, 12, 13, 13, 12, 10, 70, 34, 40, 60]); r = 4
    for b in blocks:
        for lab in labels:
            rr = [x for x in b["results"] if x["Round"] == lab]; cv = b["cov"][lab]; n = Counter(x["Table"] for x in rr)
            forms = Counter(v for x in rr for v in x["variants"])
            vals = [b["term"], lab, len(rr), n["Effectiveness"], n["Risk and Context"], n["Technical Compliance"],
                    cv["mentioned"], cv["total"], cv["pct"] / 100, cv["list"], " ~ ".join(b["aliases"]),
                    ", ".join(f"{k} ({v})" for k, v in forms.most_common(15)), " | ".join(b["warn"])]
            for j, v in enumerate(vals, 1):
                _v(ws.cell(row=r, column=j), v)
            ws.cell(row=r, column=9).number_format = "0.0%"; r += 1
    _table(ws, "tSummary", 3, r - 1, len(heads))
    used = {"summary", "all hits"}
    th = ["Round", "Table", "Reference (for citing)", "Jurisdiction", "MEASURE", "Indicator", "Rating", "Criteria rating",
          "Quote (sentence with the term)" if quotes == "sentence" else "Quote (full paragraph)", "Term in context"]
    for i, b in enumerate(blocks, 1):
        name = re.sub(r"[\\/*?\[\]:]", "_", b["term"]).strip("'")[:31] or "term"; base, k = name, 2
        while name.lower() in used: name = base[:28] + f"_{k}"; k += 1
        used.add(name.lower()); w = wb.create_sheet(name)
        w["A1"] = f'"{b["term"]}"  -  ' + "   |   ".join(f'{lab}: {b["cov"][lab]["mentioned"]} of {b["cov"][lab]["total"]} jurisdictions ({b["cov"][lab]["pct"]}%)' for lab in labels)
        w["A1"].font = Font(bold=True)
        if not b["results"]: w["A3"] = "No mentions found."; continue
        _head(w, 3, th, [11, 19, 36, 18, 12, 32, 8, 10, 95, 70]); r = 4
        for x in b["results"]:
            q = x["sentence"] if quotes == "sentence" else x["OBS_VALUE"]
            for j, v in enumerate([x["Round"], x["Table"], x["ref"], x["Jurisdiction"], x["MEASURE"], x["Indicator"], x["Rating"],
                                   x["Criteria rating"], _bold(q, x["variants"]), _bold(x["context"], x["variants"])], 1):
                c = w.cell(row=r, column=j); _v(c, v); c.alignment = Alignment(wrap_text=True, vertical="top")
            r += 1
        _table(w, f"tTerm{i}", 3, r - 1, len(th)); w.freeze_panes = "A4"
    w = wb.create_sheet("All hits")
    ah = ["Search term", "Round", "Table", "Reference (for citing)", "Jurisdiction", "MEASURE", "Rating", "Matched as", "Sentence with term"]
    _head(w, 1, ah, [26, 11, 19, 36, 18, 12, 8, 18, 90]); r = 2
    for b in blocks:
        for x in b["results"]:
            for j, v in enumerate([b["term"], x["Round"], x["Table"], x["ref"], x["Jurisdiction"], x["MEASURE"], x["Rating"], ", ".join(x["variants"]), x["sentence"]], 1):
                _v(w.cell(row=r, column=j), v)
            r += 1
    _table(w, "tAll", 1, r - 1, len(ah)); w.freeze_panes = "A2"
    wb.save(path)

# ----------------------------------------------------------------- CLI
def main():
    a = argparse.ArgumentParser()
    a.add_argument("--r4"); a.add_argument("--r5"); a.add_argument("--round", default="both", choices=["4", "5", "both"])
    a.add_argument("--terms", required=True); a.add_argument("--out", default="mer_search_results.xlsx")
    a.add_argument("--mode", default="strict", choices=["strict", "exact", "fuzzy"]); a.add_argument("--threshold", type=float, default=85)
    a.add_argument("--caps-only", action="store_true"); a.add_argument("--jurisdictions", default=""); a.add_argument("--tables", default="")
    a.add_argument("--quotes", default="paragraph", choices=["paragraph", "sentence"])
    o = a.parse_args()
    rounds = ["4", "5"] if o.round == "both" else [o.round]
    dfs = []
    for rn in rounds:
        src = o.r4 if rn == "4" else o.r5
        if not src: raise SystemExit(f"--r{rn} is required for round {rn}")
        dfs.append(load_round(rn, src)); print(f"loaded {ROUND_LABEL[rn]}: {len(dfs[-1]):,} paragraphs, {dfs[-1]['Jurisdiction'].nunique()} jurisdictions")
    df = pd.concat(dfs, ignore_index=True); labels = [ROUND_LABEL[r] for r in rounds]
    known = sorted(df["Jurisdiction"].unique())
    base = {}
    if o.jurisdictions: base["jurisdictions"] = [j.strip() for j in o.jurisdictions.split(",") if j.strip()]
    if o.tables: base["tables"] = [t.strip() for t in o.tables.split(",") if t.strip()]
    terms = [l.strip() for l in open(o.terms, encoding="utf-8") if l.strip()]
    blocks = []
    for t in terms:
        tp = prepare_term(t, o.mode, o.caps_only, known)
        if tp is None: print(f"skipped (nothing searchable): {t}"); continue
        f = dict(base)
        if tp["scope"]: f["jurisdictions"] = [tp["scope"]]
        res = search(df, tp, o.mode, o.threshold, f)
        cov = {lab: coverage(df[df["Round"] == lab], f, [r for r in res if r["Round"] == lab]) for lab in labels}
        blocks.append({"term": tp["label"], "aliases": tp["aliases"], "warn": tp["warn"], "results": res, "cov": cov})
        print(f'{tp["label"][:40]:42}' + "  ".join(f'{lab}: {sum(1 for r in res if r["Round"] == lab):5} hits, {cov[lab]["mentioned"]}/{cov[lab]["total"]} juris' for lab in labels))
        for w in tp["warn"]: print("    note:", w[:150])
    export(blocks, labels, o.out, o.quotes); print("saved", o.out)

if __name__ == "__main__":
    main()
