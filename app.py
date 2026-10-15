"""
FATF Mutual Evaluation Search Tool  (v8)
========================================
Local web app. Run:  python app.py   (or double-click run_search_tool.bat)
Opens automatically in your browser at http://127.0.0.1:5177

Folder layout expected (the tool looks next to this file first):
  Search_tool\\
      app.py  (this file)
      run_search_tool.bat
      exports\\                      (created automatically)
      source files\\MER\\round_4\\   (2 CSVs: *RC_IO* and *TECH_COMP*)
      source files\\MER\\round_5\\   (2 CSVs: *RC_IO* and *TECH_COMP*)
      source files\\other\\          (docx, pdf, xlsx, csv, txt)
"""

import os
import re
import json
import uuid
import pickle
import tempfile
import unicodedata
import hashlib
import threading
import webbrowser
import traceback
from collections import Counter, defaultdict, OrderedDict
from copy import copy
from datetime import datetime
from html import escape as html_escape

import pandas as pd
from flask import Flask, request, jsonify, Response

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter

# Optional fuzzy library (falls back to difflib if missing)
try:
    from rapidfuzz import fuzz as _rf_fuzz
    def _ratio(a, b):
        return _rf_fuzz.ratio(a, b)
except ImportError:
    import difflib
    def _ratio(a, b):
        return difflib.SequenceMatcher(None, a, b).ratio() * 100

# =============================================================================
# CONFIGURATION
# =============================================================================
# The tool uses the folder it sits in, as long as that folder has a
# "source files" sub-folder. Otherwise it falls back to the path below.
_HERE = os.path.dirname(os.path.abspath(__file__))
FALLBACK_BASE_DIR = r"C:\Users\LUO_D\Documents\Search_tool"
BASE_DIR = _HERE if os.path.isdir(os.path.join(_HERE, "source files")) else FALLBACK_BASE_DIR

MER_DIR     = os.path.join(BASE_DIR, "source files", "MER")
ROUND4_DIR  = os.path.join(MER_DIR, "round_4")
ROUND5_DIR  = os.path.join(MER_DIR, "round_5")
OTHER_DIR   = os.path.join(BASE_DIR, "source files", "other")
DEFAULT_EXPORT_DIR = os.path.join(BASE_DIR, "exports")

VERSION = "v8  (30 Sep 2026)"

CONTEXT_CHARS = 250          # characters before/after detected term
PAGE_SIZE     = 25           # rows fetched per "show more" click in the browser
PORT          = 5177

ROUND_LABEL = {"4": "4th round", "5": "5th round"}
TABLES = ["Effectiveness", "Risk and Context", "Technical Compliance"]

def rounds_of(sel):
    """'4' -> ['4'], '5' -> ['5'], 'both' -> ['4', '5']"""
    sel = str(sel or "5").strip().lower()
    if sel in ("both", "all", "4+5", "45"):
        return ["4", "5"]
    return ["4"] if sel == "4" else ["5"]

def round_labels(sel):
    return [ROUND_LABEL[r] for r in rounds_of(sel)]

def source_label(sel):
    rs = rounds_of(sel)
    if len(rs) == 2:
        return "MER 4th and 5th round"
    return "MER " + ROUND_LABEL[rs[0]]

# folders can be overridden from the interface
PATHS = {"round_4": ROUND4_DIR, "round_5": ROUND5_DIR, "other": OTHER_DIR}

def set_paths(p):
    """Accept folder overrides typed by the user; blank means the default."""
    defaults = {"round_4": ROUND4_DIR, "round_5": ROUND5_DIR, "other": OTHER_DIR}
    for k in ("round_4", "round_5", "other"):
        v = (p or {}).get(k)
        newv = str(v).strip().strip('"').strip("'") if v and str(v).strip() else defaults[k]
        if newv != PATHS[k]:
            PATHS[k] = newv
            _CACHE.pop({"round_4": "mer4", "round_5": "mer5", "other": "other"}[k], None)
            _CACHE.pop("merboth", None)

app = Flask(__name__)

@app.after_request
def _no_cache(resp):
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    return resp

# =============================================================================
# CLEANING
# =============================================================================
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

# code -> human label, used only to caption the filter dropdowns
RATING_LABELS = {}

def register_labels(codes, labels):
    for c, l in zip(codes, labels):
        c = "" if pd.isna(c) else str(c).strip()
        l = "" if pd.isna(l) else str(l).strip()
        if c and c not in ("_Z", "_X") and l and c not in RATING_LABELS:
            RATING_LABELS[c] = l

def rating_option(code):
    """'SE' -> 'SE – Substantial effectiveness' (dropdown caption only)."""
    lbl = RATING_LABELS.get(code)
    return f"{code} – {lbl}" if lbl else code

# =============================================================================
# DATA LOADING (cached in memory, and on disk so restarts are fast)
# =============================================================================
_CACHE = {}   # keys: "mer4", "mer5", "merboth", "other"
LOADER_VERSION = "8.0"

def _cache_dir():
    """Local (not synced) folder for the fast-restart cache."""
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    d = os.path.join(base, "MER_Search_cache")
    try:
        os.makedirs(d, exist_ok=True)
        return d
    except Exception:
        return None

def _file_sig(path):
    st = os.stat(path)
    return "%s|%d|%d" % (os.path.abspath(path), st.st_size, int(st.st_mtime))

def _find_csv(folder, keywords, what):
    """Find the first CSV in folder whose name contains any keyword."""
    if not os.path.isdir(folder):
        raise FileNotFoundError(f"Folder not found: {folder}")
    for f in sorted(os.listdir(folder)):
        low = f.lower()
        if low.endswith(".csv") and any(k in low for k in keywords):
            return os.path.join(folder, f)
    raise FileNotFoundError(
        f"No {what} CSV found in {folder} (looking for a file name containing "
        f"{' or '.join(keywords)}). Files present: {', '.join(os.listdir(folder)) or 'none'}")

UNIFORM_COLS = ["Round", "Jurisdiction", "Table", "MEASURE", "Indicator",
                "Rating", "Criteria rating", "OBS_VALUE"]

def _build_round(rn):
    """Read and clean one round's two CSVs into one uniform table."""
    folder = PATHS["round_" + rn]
    kfra_path  = _find_csv(folder, ["rc_io", "kfra"], "effectiveness / risk and context")
    tcomp_path = _find_csv(folder, ["tech_comp", "tcomp"], "technical compliance")

    # ---- KFRA file: Effectiveness + Risk & Context combined ----
    kfra = pd.read_csv(kfra_path, encoding="utf-8-sig", dtype=str,
                       usecols=["Jurisdiction", "MEASURE", "Indicator",
                                "EFF_CAT", "Effectiveness rating", "OBS_VALUE"])
    kfra["OBS_VALUE"] = kfra["OBS_VALUE"].apply(strip_html)
    kfra = kfra[kfra["OBS_VALUE"] != ""].copy()
    register_labels(kfra["EFF_CAT"], kfra["Effectiveness rating"])
    kfra["Rating"] = [combine_rating(c, l) for c, l in
                      zip(kfra["EFF_CAT"], kfra["Effectiveness rating"])]
    kfra["Criteria rating"] = "NA"
    is_io = kfra["MEASURE"].fillna("").str.startswith("IO")
    kfra["Table"] = ["Effectiveness" if x else "Risk and Context" for x in is_io]
    kfra.loc[~is_io, "Rating"] = "NA"       # Risk & Context rows carry no rating

    # ---- TCOMP file: Technical Compliance ----
    tc = pd.read_csv(tcomp_path, encoding="utf-8-sig", dtype=str,
                     usecols=["Jurisdiction", "MEASURE", "Indicator",
                              "COMP_RATING", "Technical compliance rating",
                              "CRIT_RATING", "Criteria rating", "OBS_VALUE"])
    tc = tc.rename(columns={"Criteria rating": "_crit_label"})
    tc["OBS_VALUE"] = tc["OBS_VALUE"].apply(strip_html)
    tc = tc[tc["OBS_VALUE"] != ""].copy()
    register_labels(tc["COMP_RATING"], tc["Technical compliance rating"])
    register_labels(tc["CRIT_RATING"], tc["_crit_label"])
    tc["Rating"] = [combine_rating(c, l) for c, l in
                    zip(tc["COMP_RATING"], tc["Technical compliance rating"])]
    tc["Criteria rating"] = [combine_rating(c, l) for c, l in
                             zip(tc["CRIT_RATING"], tc["_crit_label"])]
    tc["Table"] = "Technical Compliance"

    kfra["Round"] = ROUND_LABEL[rn]
    tc["Round"] = ROUND_LABEL[rn]
    df = pd.concat([kfra[UNIFORM_COLS], tc[UNIFORM_COLS]], ignore_index=True)
    df = df.fillna("")
    # lower-case, accent-free copy of every paragraph: lets the search skip
    # paragraphs that cannot possibly match before running the full pattern
    df["_norm"] = [norm_text(x) for x in df["OBS_VALUE"].tolist()]
    return df, [kfra_path, tcomp_path]

def load_round(rn):
    key = "mer" + rn
    if key in _CACHE:
        return _CACHE[key]
    folder = PATHS["round_" + rn]
    kfra_path  = _find_csv(folder, ["rc_io", "kfra"], "effectiveness / risk and context")
    tcomp_path = _find_csv(folder, ["tech_comp", "tcomp"], "technical compliance")
    sig = hashlib.md5("|".join([LOADER_VERSION, pd.__version__, _file_sig(kfra_path),
                                _file_sig(tcomp_path)]).encode()).hexdigest()[:16]
    cdir = _cache_dir()
    cpath = os.path.join(cdir, f"round{rn}_{sig}.pkl") if cdir else None
    df = None
    if cpath and os.path.exists(cpath):
        try:
            with open(cpath, "rb") as fh:
                obj = pickle.load(fh)
            df = obj["df"]
            RATING_LABELS.update(obj.get("labels", {}))
        except Exception:
            df = None
    if df is None:
        df, _ = _build_round(rn)
        if cpath:
            try:
                for old in os.listdir(cdir):          # drop stale caches
                    if old.startswith(f"round{rn}_") and old.endswith(".pkl"):
                        try:
                            os.remove(os.path.join(cdir, old))
                        except Exception:
                            pass
                with open(cpath, "wb") as fh:
                    pickle.dump({"df": df, "labels": dict(RATING_LABELS)}, fh,
                                protocol=pickle.HIGHEST_PROTOCOL)
            except Exception:
                pass
    _CACHE[key] = df
    return df

def load_mer(sel):
    """One round, or both rounds stacked (each row keeps its Round label)."""
    rs = rounds_of(sel)
    if len(rs) == 1:
        return load_round(rs[0])
    parts = [load_round(r) for r in rs]
    ids = tuple(id(x) for x in parts)
    hit = _CACHE.get("merboth")
    if hit is not None and hit[0] == ids:
        return hit[1]
    df = pd.concat(parts, ignore_index=True)
    _CACHE["merboth"] = (ids, df)
    return df

MAX_RECORDS_PER_FILE = 40000

def _xl_col(n):
    """0-based column index -> Excel letters (0->A, 27->AB)."""
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def _clean(t):
    if t is None or (isinstance(t, float) and pd.isna(t)):
        return ""
    return strip_html(str(t))

def _records_from_tabular(path, is_csv):
    """Every non-empty cell of every sheet becomes a record with an exact
       cell address, the column heading, and the whole row as details."""
    recs = []
    if is_csv:
        sheets = {"": pd.read_csv(path, encoding="utf-8-sig", header=None, dtype=str)}
    else:
        sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=str)
    for sname, d in sheets.items():
        d = d.fillna("")
        if d.empty:
            continue
        header = [_clean(v) for v in d.iloc[0].tolist()]   # row 1, used as labels
        for ri in range(len(d)):
            rowvals = [_clean(v) for v in d.iloc[ri].tolist()]
            details = " | ".join(f"{header[ci] or _xl_col(ci)}: {v}"
                                 for ci, v in enumerate(rowvals) if v)[:1200]
            for ci, val in enumerate(rowvals):
                if not val:
                    continue
                where = (f"Sheet '{sname}' \u00b7 cell {_xl_col(ci)}{ri+1}" if sname
                         else f"Row {ri+1}, column {ci+1}")
                colname = header[ci] if ri > 0 and header[ci] else ""
                if colname:
                    where += f" (column: {colname})"
                recs.append({"Location": where, "Text": val, "Details": details})
                if len(recs) >= MAX_RECORDS_PER_FILE:
                    return recs, "file truncated at record cap"
    return recs, None

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def _records_from_docx(path):
    """Walk the Word document in order. Track the heading trail so every
       paragraph and table cell is located as Section > paragraph N or
       Section > Table N, row R, column 'heading'."""
    import zipfile
    from xml.etree import ElementTree as ET
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    body = root.find(_W + "body")

    def ptext(p):
        return re.sub(r"\s+", " ",
                      "".join(t.text or "" for t in p.iter(_W + "t"))).strip()

    def pstyle(p):
        st = p.find(_W + "pPr/" + _W + "pStyle")
        return (st.get(_W + "val") or "") if st is not None else ""

    recs, trail = [], {}          # trail: heading level -> text
    para_no, tbl_no = 0, 0

    def section():
        return " > ".join(trail[k] for k in sorted(trail)) or "(start of document)"

    def walk(elems):
        nonlocal para_no, tbl_no
        for el in elems:
            tag = el.tag
            if tag == _W + "p":
                txt = ptext(el)
                if not txt:
                    continue
                m = re.match(r"[Hh]eading\s*(\d)", pstyle(el))
                if m:
                    lvl = int(m.group(1))
                    trail[lvl] = txt
                    for k in [k for k in trail if k > lvl]:
                        del trail[k]
                    para_no = 0
                    recs.append({"Location": f"Heading: {section()}",
                                 "Text": txt, "Details": "heading"})
                else:
                    para_no += 1
                    recs.append({"Location": f"{section()} \u00b7 paragraph {para_no}",
                                 "Text": strip_html(txt), "Details": ""})
            elif tag == _W + "tbl":
                tbl_no += 1
                rows = el.findall(_W + "tr")
                headcells = []
                if rows:
                    headcells = [re.sub(r"\s+", " ",
                                 " ".join(ptext(p) for p in tc.iter(_W + "p"))).strip()
                                 for tc in rows[0].findall(_W + "tc")]
                for ri, row in enumerate(rows):
                    for ci, tc in enumerate(row.findall(_W + "tc")):
                        cell = re.sub(r"\s+", " ",
                               " ".join(ptext(p) for p in tc.iter(_W + "p"))).strip()
                        if not cell:
                            continue
                        col = (headcells[ci] if ri > 0 and ci < len(headcells)
                               and headcells[ci] else f"col {ci+1}")
                        recs.append({
                            "Location": f"{section()} \u00b7 Table {tbl_no}, "
                                        f"row {ri+1}, column '{col[:40]}'",
                            "Text": strip_html(cell), "Details": ""})
            if len(recs) >= MAX_RECORDS_PER_FILE:
                return
    walk(list(body))
    return recs, None

def _records_from_pdf(path):
    try:
        from pypdf import PdfReader
    except ImportError:
        try:
            from PyPDF2 import PdfReader
        except ImportError:
            return None, ("PDF support needs one small package: run "
                          "pip install pypdf   then restart the tool")
    recs = []
    reader = PdfReader(path)
    for pi, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        # split into paragraphs on blank lines, else on single newlines
        paras = [p for p in re.split(r"\n\s*\n", text) if p.strip()]
        if len(paras) <= 1:
            paras = [p for p in text.split("\n") if p.strip()]
        for qi, para in enumerate(paras, start=1):
            clean = strip_html(re.sub(r"\s+", " ", para).strip())
            if len(clean) < 3:
                continue
            recs.append({"Location": f"Page {pi} \u00b7 paragraph {qi}",
                         "Text": clean, "Details": ""})
            if len(recs) >= MAX_RECORDS_PER_FILE:
                return recs, "file truncated at record cap"
    return recs, None

def load_other():
    """Read every file in the 'other' folder into location-tagged records.
       Supported: .docx, .pdf, .xlsx/.xlsm/.xls, .csv, .txt"""
    if "other" in _CACHE:
        return _CACHE["other"]
    other_dir = PATHS["other"]
    if not os.path.isdir(other_dir):
        raise FileNotFoundError(f"Folder not found: {other_dir}")
    datasets = []
    for f in sorted(os.listdir(other_dir)):
        path = os.path.join(other_dir, f)
        low = f.lower()
        if low.endswith(".txt") and ("put " in low or "saved here" in low):
            continue                       # our placeholder files
        try:
            if low.endswith(".csv"):
                recs, note = _records_from_tabular(path, True)
            elif low.endswith((".xlsx", ".xlsm", ".xls")):
                recs, note = _records_from_tabular(path, False)
            elif low.endswith(".docx"):
                recs, note = _records_from_docx(path)
            elif low.endswith(".pdf"):
                recs, note = _records_from_pdf(path)
            elif low.endswith(".txt"):
                lines = open(path, encoding="utf-8", errors="replace").read()
                paras = [p for p in re.split(r"\n\s*\n", lines) if p.strip()]
                recs = [{"Location": f"paragraph {i+1}",
                         "Text": strip_html(re.sub(r"\s+", " ", p)), "Details": ""}
                        for i, p in enumerate(paras)]
                note = None
            elif low.endswith(".doc"):
                datasets.append({"file": f, "error":
                    "old .doc format is not supported - open it in Word and "
                    "save as .docx", "records": None})
                continue
            else:
                continue
            if recs is None:
                datasets.append({"file": f, "error": note, "records": None})
            else:
                datasets.append({"file": f, "error": None, "records": recs,
                                 "note": note})
        except Exception as e:
            datasets.append({"file": f, "error": str(e), "records": None})
    if not any(x.get("records") for x in datasets):
        raise FileNotFoundError(
            f"No readable files in {other_dir} "
            f"(supported: docx, pdf, xlsx, csv, txt)")
    _CACHE["other"] = datasets
    return datasets

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

# =============================================================================
# SEARCH ENGINE
# =============================================================================
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

def search_mer(df, tp, mode, threshold, filters):
    """Every paragraph in df matching the term group, sorted by round, table,
       jurisdiction and measure."""
    sub = apply_filters(df, filters)
    if sub.empty or (tp["pattern"] is None and not any(tp["tokens"])):
        return []
    cols = {c: sub[c].tolist() for c in UNIFORM_COLS}
    texts = cols["OBS_VALUE"]
    if mode in ("strict", "exact"):
        norms = sub["_norm"].tolist()
        groups = [g for g in tp["ptokens"] if g]
        idx = [i for i, n in enumerate(norms)
               if any(all(t in n for t in g) for g in groups)]
    else:
        idx = range(len(texts))
    results = []
    for i in idx:
        text = texts[i]
        hit = find_match(text, tp, mode, threshold)
        if hit is None:
            continue
        start, end, variants = hit
        r = {c: cols[c][i] for c in UNIFORM_COLS}
        r["context"] = make_context(text, start, end)
        r["sentence"] = sentence_of(text, start, end)
        r["matched"] = variants[0]
        r["variants"] = variants
        r["ref"] = citation(r)
        r["_id"] = row_id(r)
        results.append(r)
    results.sort(key=lambda r: (_TABLE_ORDER.get(r["Table"], 9), r["Jurisdiction"],
                                r["Round"], _measure_key(r["MEASURE"])))
    return results

def coverage_stats(df, filters, results):
    """% of jurisdictions (within the current filter scope) that mention the term.
       Always called for ONE round at a time."""
    scope = apply_filters(df, filters)
    total = int(scope["Jurisdiction"].nunique())
    counts = Counter(r["Jurisdiction"] for r in results)
    mentioned = len(counts)
    pct = round(100.0 * mentioned / total, 1) if total else 0.0
    per_table = {tbl: len({r["Jurisdiction"] for r in results if r["Table"] == tbl})
                 for tbl in TABLES}
    return {"mentioned": mentioned, "total": total, "pct": pct,
            "per_table": per_table,
            "jurisdictions": [{"name": k, "hits": v} for k, v in counts.most_common()]}

def round_compare(scope4, scope5, results):
    """For jurisdictions assessed in both rounds: hits in each round."""
    both = sorted(set(scope4["Jurisdiction"].unique()) & set(scope5["Jurisdiction"].unique()))
    c4 = Counter(r["Jurisdiction"] for r in results if r["Round"] == "4th round")
    c5 = Counter(r["Jurisdiction"] for r in results if r["Round"] == "5th round")
    out = []
    for j in both:
        a, b = c4.get(j, 0), c5.get(j, 0)
        pattern = ("Both rounds" if a and b else "4th round only" if a
                   else "5th round only" if b else "Neither round")
        out.append({"name": j, "r4": a, "r5": b, "pattern": pattern})
    return out

OTHER_COLUMNS = ["Location", "Sentence with term", "Details", "Full text"]

def search_other(datasets, tp, mode, threshold):
    """Search the Text of every record; report exactly where it was found
       and the sentence containing the term."""
    out = []
    pgroups = [g for g in tp["ptokens"] if g]
    for ds in datasets:
        recs = ds.get("records")
        if recs is None:
            out.append({"file": ds["file"], "columns": OTHER_COLUMNS,
                        "rows": [], "error": ds.get("error")})
            continue
        rows = []
        for rec in recs:
            text = rec["Text"]
            if mode in ("strict", "exact"):
                n = rec.get("_norm")
                if n is None:
                    n = rec["_norm"] = norm_text(text)
                if not any(all(t in n for t in g) for g in pgroups):
                    continue
            hit = find_match(text, tp, mode, threshold)
            if hit is None:
                continue
            start, end, variants = hit
            r = {
                "File": ds["file"],
                "Location": rec["Location"],
                "Sentence with term": sentence_of(text, start, end),
                "Details": rec.get("Details", ""),
                "Full text": text,
                "context": make_context(text, start, end),
                "matched": variants[0],
                "variants": variants,
                "ref": f'{ds["file"]}, {rec["Location"]}',
            }
            r["sentence"] = r["Sentence with term"]
            r["_id"] = row_id(r)
            rows.append(r)
        out.append({"file": ds["file"], "columns": OTHER_COLUMNS,
                    "rows": rows, "error": None, "note": ds.get("note")})
    return out

def _form_counts(rows):
    """Exactly which spellings/capitalisations were matched, and how often."""
    c = Counter()
    for r in rows:
        for v in (r.get("variants") or [r.get("matched")]):
            if v:
                c[v] += 1
    return [{"form": k, "n": v} for k, v in c.most_common(30)]

# =============================================================================
# RUNNING A SEARCH  (raw results are kept in memory so paging and exports are
# instant and always match what was on screen)
# =============================================================================
_STORE = OrderedDict()     # sid -> {"key": ..., "raw": [...]}
_STORE_MAX = 6
_PROGRESS = {}             # sid -> {"done", "total", "term", "finished"}
_LOCK = threading.Lock()
_EXPORT_LOCK = threading.Lock()

def _search_key(p):
    keep = ("data_mode", "round", "terms", "search_mode", "threshold",
            "caps_only", "filters", "paths")
    return json.dumps({k: p.get(k) for k in keep}, sort_keys=True, default=str)

def _set_progress(sid, done, total, term, finished=False):
    if sid:
        _PROGRESS[sid] = {"done": done, "total": total, "term": term,
                          "finished": finished}

def run_search(p, sid=None):
    """Search every term. Returns raw blocks (no exclusions applied)."""
    set_paths(p.get("paths"))
    terms = [t.strip() for t in p.get("terms", []) if t.strip()]   # no limit
    mode = p.get("search_mode", "strict")
    threshold = float(p.get("threshold", 85))
    csa = bool(p.get("caps_only", False))
    raw = []
    if p["data_mode"] == "MER":
        df = load_mer(p["round"])
        known = sorted(df["Jurisdiction"].unique().tolist())
        for i, group in enumerate(terms):
            _set_progress(sid, i, len(terms), group)
            tp = prepare_term(group, mode, csa, known)
            if tp is None:
                continue
            filters = dict(p.get("filters", {}) or {})
            if tp["scope"]:                      # "SAMLIT (South Africa)"
                filters["jurisdictions"] = [tp["scope"]]
            res = search_mer(df, tp, mode, threshold, filters)
            raw.append({"term": tp["label"], "aliases": tp["aliases"],
                        "scope": tp["scope"], "notes": tp["notes"],
                        "how": tp["how"], "warn": tp["warn"], "input": tp["input"],
                        "filters": filters, "results": res})
    else:
        ds = load_other()
        for i, group in enumerate(terms):
            _set_progress(sid, i, len(terms), group)
            tp = prepare_term(group, mode, csa, None)
            if tp is None:
                continue
            res = search_other(ds, tp, mode, threshold)
            raw.append({"term": tp["label"], "aliases": tp["aliases"],
                        "scope": None, "notes": tp["notes"], "how": tp["how"],
                        "warn": tp["warn"], "input": tp["input"],
                        "filters": {}, "results": res})
    _set_progress(sid, len(terms), len(terms), "", finished=True)
    return raw

def get_raw(p, sid):
    """Reuse the stored search if the settings are unchanged, else run it."""
    key = _search_key(p)
    with _LOCK:
        ent = _STORE.get(sid) if sid else None
    if ent and ent["key"] == key:
        return ent["raw"]
    raw = run_search(p, sid)
    if sid:
        with _LOCK:
            _STORE[sid] = {"key": key, "raw": raw, "data_mode": p["data_mode"]}
            _STORE.move_to_end(sid)
            while len(_STORE) > _STORE_MAX:
                _STORE.popitem(last=False)
    return raw

def finalize(p, raw, excluded=None):
    """Apply removed hits, then compute statistics - always per round."""
    excluded = set(excluded or [])
    blocks = []
    if p["data_mode"] == "MER":
        df = load_mer(p["round"])
        labels = round_labels(p["round"])
        per_round = {lab: df[df["Round"] == lab] for lab in labels}
        for b in raw:
            res = [r for r in b["results"] if r["_id"] not in excluded]
            nb = dict(b)
            nb["results"] = res
            nb["rounds"] = labels
            nb["coverage"] = {lab: coverage_stats(per_round[lab], b["filters"],
                                                  [r for r in res if r["Round"] == lab])
                              for lab in labels}
            if len(labels) == 2:
                nb["round_compare"] = round_compare(
                    apply_filters(per_round["4th round"], b["filters"]),
                    apply_filters(per_round["5th round"], b["filters"]), res)
            nb["forms"] = _form_counts(res)
            nb["source_label"] = source_label(p["round"])
            blocks.append(nb)
    else:
        for b in raw:
            files = []
            for fb in b["results"]:
                nfb = dict(fb)
                nfb["rows"] = [r for r in fb["rows"] if r["_id"] not in excluded]
                files.append(nfb)
            nb = dict(b)
            nb["results"] = files
            nb["rounds"] = []
            nb["forms"] = _form_counts([r for fb in files for r in fb["rows"]])
            nb["source_label"] = "Other files"
            blocks.append(nb)
    return blocks

def _slim_row(r):
    """What the browser needs for one hit."""
    keep = ("_id", "Round", "Jurisdiction", "Table", "MEASURE", "Indicator", "Rating",
            "Criteria rating", "OBS_VALUE", "context", "sentence", "matched",
            "variants", "ref", "File", "Location", "Sentence with term", "Details",
            "Full text")
    return {k: r[k] for k in keep if k in r}


# =============================================================================
# EXCEL HELPERS
# =============================================================================
def sanitize_sheet(name, used):
    s = re.sub(r'[\\/*?\[\]:\x00-\x1f]', '_', str(name)).strip().strip("'")[:31].strip("'")
    s = s or "term"
    base, i = s, 2
    while s.lower() in used:
        suffix = f"_{i}"
        s = base[:31 - len(suffix)] + suffix
        i += 1
    used.add(s.lower())
    return s

def bold_cell(text, matched):
    """Build a rich-text cell with every occurrence of the matched term(s)
       bolded. `matched` may be a string or a list of variant strings."""
    if not matched or not text:
        return text
    variants = [matched] if isinstance(matched, str) else list(matched)
    variants = sorted({v for v in variants if v}, key=len, reverse=True)
    if not variants:
        return text
    pat = re.compile("|".join(re.escape(v) for v in variants), re.IGNORECASE)
    if not pat.search(text):
        return text
    bold = InlineFont(b=True)
    parts, last = [], 0
    for m in pat.finditer(text):
        if m.start() > last:
            parts.append(text[last:m.start()])
        parts.append(TextBlock(bold, m.group()))
        last = m.end()
    if last < len(text):
        parts.append(text[last:])
    return CellRichText(*parts)


# =============================================================================
# EXCEL EXPORT
# Every data block is a real Excel table, so every column - including Round -
# has its own filter arrow and the data can go straight into a pivot table.
# =============================================================================
TABLE_FILL = {
    "Effectiveness":        PatternFill("solid", fgColor="DCE6F1"),  # light blue
    "Risk and Context":     PatternFill("solid", fgColor="E2EFDA"),  # light green
    "Technical Compliance": PatternFill("solid", fgColor="FCE4D6"),  # light orange
}
ROUND_FILL = {"4th round": PatternFill("solid", fgColor="E4DDF4"),   # lavender
              "5th round": PatternFill("solid", fgColor="FFEFC1")}   # amber
ROUND_INK  = {"4th round": "4B3B8F", "5th round": "8A5A00"}
HEAD_FILL  = PatternFill("solid", fgColor="1E2F4D")
BAND_FILL  = PatternFill("solid", fgColor="1E2F4D")
NOTE_FILL  = PatternFill("solid", fgColor="F4F8FC")
WARN_FILL  = PatternFill("solid", fgColor="FBF1DA")
WRAP = Alignment(wrap_text=True, vertical="top")
RESERVED_SHEETS = {"summary", "country matrix", "round comparison", "all hits"}

_STYLE_CACHE = {}

def _style(cell, key, **kw):
    """Style a cell. The first cell of each style is styled normally; later
       cells copy its style record, which is many times faster on big exports."""
    sa = _STYLE_CACHE.get(key)
    if sa is None:
        for k, v in kw.items():
            setattr(cell, k, v)
        _STYLE_CACHE[key] = copy(cell._style)
    else:
        cell._style = copy(sa)

def _setv(cell, v):
    """Write a value safely: strip characters Excel rejects, never let text
       that starts with '=' turn into a formula."""
    if isinstance(v, CellRichText):
        cell.value = v
        return
    if isinstance(v, str):
        v = ILLEGAL_XLSX.sub("", v)[:32000]
        cell.value = v
        if v.startswith("="):
            cell.data_type = "s"
    else:
        cell.value = v

def _rich(v, variants):
    v = ILLEGAL_XLSX.sub("", "" if v is None else str(v))[:32000]
    out = bold_cell(v, variants)
    return out

def _sheet_link(sname):
    """Internal hyperlink to cell A1 of a sheet (apostrophes doubled)."""
    from openpyxl.worksheet.hyperlink import Hyperlink
    return Hyperlink(ref="A1", location="'%s'!A1" % sname.replace("'", "''"),
                     display=sname)

def _unique(headers):
    seen, out = {}, []
    for h in headers:
        h = re.sub(r"\s+", " ", str(h)).strip() or "Column"
        base, n = h, 2
        while h.lower() in seen:
            h = f"{base} ({n})"
            n += 1
        seen[h.lower()] = 1
        out.append(h)
    return out

_TABLE_NAMES = set()
def _add_table(ws, name, header_row, last_row, ncols):
    """Turn a block into an Excel table (filter arrows on every column)."""
    if last_row <= header_row:
        return
    safe = re.sub(r"[^A-Za-z0-9_]", "_", name)
    if not re.match(r"[A-Za-z_]", safe):
        safe = "t_" + safe
    base, n = safe, 2
    while safe.lower() in _TABLE_NAMES:
        safe = f"{base}_{n}"
        n += 1
    _TABLE_NAMES.add(safe.lower())
    t = Table(displayName=safe,
              ref=f"A{header_row}:{get_column_letter(ncols)}{last_row}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True,
                                      showColumnStripes=False, showFirstColumn=False,
                                      showLastColumn=False)
    ws.add_table(t)

def _write_header(ws, row, headers, widths=None, height=None):
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=j)
        _setv(c, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
        if widths and j <= len(widths):
            ws.column_dimensions[get_column_letter(j)].width = widths[j - 1]
    if height:
        ws.row_dimensions[row].height = height

def _band(ws, row, text, ncols, size=13):
    c = ws.cell(row=row, column=1)
    _setv(c, text)
    c.font = Font(bold=True, size=size, color="FFFFFF")
    for j in range(1, max(ncols, 1) + 1):
        ws.cell(row=row, column=j).fill = BAND_FILL
    ws.row_dimensions[row].height = 22

def _note(ws, row, text, ncols, fill=NOTE_FILL, color="3A4356", italic=False, bold=False):
    c = ws.cell(row=row, column=1)
    _setv(c, text)
    c.font = Font(size=10, color=color, italic=italic, bold=bold)
    for j in range(1, max(ncols, 1) + 1):
        ws.cell(row=row, column=j).fill = fill

def _round_cell(cell, label):
    _setv(cell, label)
    if label in ROUND_FILL:
        _style(cell, "round:" + label, fill=ROUND_FILL[label],
               font=Font(bold=True, color=ROUND_INK[label]), alignment=WRAP)
    else:
        _style(cell, "wrap", alignment=WRAP)

def _stat_line(b, lab):
    rows = [x for x in b["results"] if x["Round"] == lab]
    cov = b["coverage"][lab]
    nby = Counter(x["Table"] for x in rows)
    meas = Counter(pretty_measure(x["MEASURE"]) for x in rows)
    rat = Counter(x["Rating"] for x in rows if x.get("Rating") and x["Rating"] != "NA")
    s = (f"{lab.upper()}:  {len(rows)} mention{'s' if len(rows) != 1 else ''} in "
         f"{cov['mentioned']} of {cov['total']} jurisdictions ({cov['pct']}%)   |   "
         f"Effectiveness {nby.get('Effectiveness', 0)}, Risk and context "
         f"{nby.get('Risk and Context', 0)}, Technical compliance "
         f"{nby.get('Technical Compliance', 0)}")
    if meas:
        s += "   |   Most under: " + ", ".join(f"{k} ({v})" for k, v in meas.most_common(5))
    if rat:
        s += "   |   Ratings: " + ", ".join(f"{k} ({v})" for k, v in rat.most_common())
    return s

def _write_summary(ws, p, blocks, sheet_names):
    is_mer = p["data_mode"] == "MER"
    labels = blocks[0]["rounds"] if (blocks and is_mer) else []
    ncols = 14 if is_mer else 5
    _band(ws, 1, "FATF Search Tool  -  Summary", ncols)
    _note(ws, 2, f"Source: {blocks[0]['source_label'] if blocks else ''}   |   "
                 f"Search mode: {p.get('search_mode')}   |   {len(blocks)} search term"
                 f"{'s' if len(blocks) != 1 else ''}   |   Exported "
                 f"{datetime.now().strftime('%d %B %Y %H:%M')}", ncols)
    r = 3
    if len(labels) == 2:
        _note(ws, r, "Each term has one row per round. The rounds are never added "
                     "together: they cover different sets of countries, so a combined "
                     "count or percentage would mislead. Use the filter arrow on the "
                     "Round column to see one round.", ncols, fill=WARN_FILL,
              color="7A5A12", bold=True)
        r += 1
    _note(ws, r, "Click a search term to jump to its sheet. 'Country matrix' shows which "
                 "jurisdictions mention which terms; 'All hits' lists every hit in one "
                 "filterable table.", ncols, italic=True)
    r += 2
    hr = r
    if is_mer:
        headers = ["Search term", "Round", "Total hits", "Effectiveness",
                   "Risk and Context", "Technical Compliance", "Jurisdictions mentioned",
                   "Jurisdictions in scope", "Coverage %", "Jurisdictions list",
                   "Equivalent terms searched", "Restricted to",
                   "Matched in text as (count)", "How it was searched / warnings"]
        widths = [30, 12, 10, 13, 13, 14, 14, 13, 11, 70, 34, 16, 40, 80]
        _write_header(ws, hr, headers, widths, height=32)
        r = hr + 1
        for b, sname in zip(blocks, sheet_names):
            expl = " | ".join(list(b.get("notes", [])) + list(b.get("how", []))
                              + ["WARNING: " + w for w in b.get("warn", [])])
            for lab in labels:
                rows = [x for x in b["results"] if x["Round"] == lab]
                cov = b["coverage"][lab]
                nby = Counter(x["Table"] for x in rows)
                forms = _form_counts(rows)
                vals = [b["term"], lab, len(rows),
                        nby.get("Effectiveness", 0), nby.get("Risk and Context", 0),
                        nby.get("Technical Compliance", 0), cov["mentioned"],
                        cov["total"], (cov["pct"] / 100.0),
                        ", ".join(f"{j['name']} ({j['hits']})" for j in cov["jurisdictions"]),
                        " ~ ".join(b.get("aliases") or [b["term"]]), b.get("scope") or "",
                        ", ".join(f'{f["form"]} ({f["n"]})' for f in forms), expl]
                for j, v in enumerate(vals, start=1):
                    c = ws.cell(row=r, column=j)
                    if j == 2:
                        _round_cell(c, v)
                        continue
                    _setv(c, v)
                    c.alignment = WRAP
                    if j == 9:
                        c.number_format = "0.0%"
                        c.font = Font(bold=True)
                c = ws.cell(row=r, column=1)
                c.hyperlink = _sheet_link(sname)
                c.font = Font(color="1F4E79", underline="single", bold=True)
                r += 1
        _add_table(ws, "tSummary", hr, r - 1, len(headers))
    else:
        headers = ["Search term", "Total hits", "Hits per file",
                   "Equivalent terms searched", "How it was searched / warnings"]
        widths = [30, 11, 70, 40, 80]
        _write_header(ws, hr, headers, widths, height=30)
        r = hr + 1
        for b, sname in zip(blocks, sheet_names):
            total = sum(len(fb["rows"]) for fb in b["results"])
            per = ", ".join(f"{fb['file']}: {len(fb['rows'])}" for fb in b["results"])
            expl = " | ".join(list(b.get("notes", [])) + list(b.get("how", []))
                              + ["WARNING: " + w for w in b.get("warn", [])])
            for j, v in enumerate([b["term"], total, per,
                                   " ~ ".join(b.get("aliases") or [b["term"]]), expl], 1):
                c = ws.cell(row=r, column=j)
                _setv(c, v)
                c.alignment = WRAP
            c = ws.cell(row=r, column=1)
            c.hyperlink = _sheet_link(sname)
            c.font = Font(color="1F4E79", underline="single", bold=True)
            r += 1
        _add_table(ws, "tSummary", hr, r - 1, len(headers))
    ws.freeze_panes = ws.cell(row=hr + 1, column=2)

def _write_matrix(ws, p, blocks):
    """Jurisdiction x term: how many paragraphs mention each term."""
    labels = blocks[0]["rounds"] if blocks else []
    df = load_mer(p["round"])
    base_filters = dict(p.get("filters", {}) or {})
    heads = _unique(["Round", "Jurisdiction"] + [b["term"] for b in blocks]
                    + ["Terms mentioned"])
    ncols = len(heads)
    _band(ws, 1, "Country matrix  -  paragraphs mentioning each term, per jurisdiction", ncols)
    _note(ws, 2, "Each number is how many paragraphs of that jurisdiction's report mention "
                 "the term (empty = not mentioned). 'Terms mentioned' counts how many of "
                 "the searched terms appear at all. Rounds are listed separately"
                 + (" - filter the Round column to see one." if len(labels) == 2 else "."),
          ncols, italic=True)
    hr = 4
    _write_header(ws, hr, heads, [11, 26] + [13] * len(blocks) + [12], height=60)
    r = hr + 1
    for lab in labels:
        scope = apply_filters(df[df["Round"] == lab], base_filters)
        juris = sorted(scope["Jurisdiction"].unique().tolist())
        counts = [Counter(x["Jurisdiction"] for x in b["results"] if x["Round"] == lab)
                  for b in blocks]
        rows = []
        for j in juris:
            vals = [c.get(j, 0) for c in counts]
            rows.append((j, vals, sum(1 for v in vals if v)))
        rows.sort(key=lambda t: (-t[2], t[0]))
        for j, vals, nterms in rows:
            _round_cell(ws.cell(row=r, column=1), lab)
            _setv(ws.cell(row=r, column=2), j)
            for k, v in enumerate(vals, start=3):
                c = ws.cell(row=r, column=k, value=v)
                _style(c, "matrix", number_format="0;-0;;@",      # zeros show as empty
                       alignment=Alignment(horizontal="center"))
            c = ws.cell(row=r, column=ncols, value=nterms)
            _style(c, "matrix_total", font=Font(bold=True),
                   alignment=Alignment(horizontal="center"))
            r += 1
    last = r - 1
    if last > hr:
        for k in range(3, 3 + len(blocks)):
            col = get_column_letter(k)
            ws.conditional_formatting.add(
                f"{col}{hr + 1}:{col}{last}",
                ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                               end_type="max", end_color="7FA6C9"))
        _add_table(ws, "tMatrix", hr, last, ncols)
    ws.freeze_panes = ws.cell(row=hr + 1, column=3)

def _write_compare(ws, blocks):
    heads = ["Search term", "Jurisdiction", "4th round hits", "5th round hits",
             "Change", "Pattern"]
    _band(ws, 1, "Round comparison  -  jurisdictions assessed in both rounds", len(heads))
    _note(ws, 2, "For each jurisdiction that has both a 4th and a 5th round report: how "
                 "many paragraphs mention the term in each. A term appearing only in one "
                 "round can reflect a change in the country, in the methodology (2013 vs "
                 "2022), or simply in wording - read the paragraphs before drawing "
                 "conclusions.", len(heads), italic=True)
    ws.row_dimensions[2].height = 45
    ws.cell(row=2, column=1).alignment = WRAP
    hr = 4
    _write_header(ws, hr, heads, [30, 26, 14, 14, 10, 18], height=30)
    r = hr + 1
    fills = {"Both rounds": PatternFill("solid", fgColor="E2F1EA"),
             "4th round only": ROUND_FILL["4th round"],
             "5th round only": ROUND_FILL["5th round"],
             "Neither round": PatternFill("solid", fgColor="EDF0F4")}
    for b in blocks:
        for row in b.get("round_compare", []):
            vals = [b["term"], row["name"], row["r4"], row["r5"],
                    row["r5"] - row["r4"], row["pattern"]]
            for j, v in enumerate(vals, start=1):
                _setv(ws.cell(row=r, column=j), v)
            ws.cell(row=r, column=5).number_format = "+0;-0;0"
            ws.cell(row=r, column=6).fill = fills.get(row["pattern"], NOTE_FILL)
            r += 1
    _add_table(ws, "tCompare", hr, r - 1, len(heads))
    ws.freeze_panes = ws.cell(row=hr + 1, column=3)

MER_HEADERS = ["Round", "Table", "Reference (for citing)", "Jurisdiction", "MEASURE",
               "Indicator", "Rating", "Criteria rating", "Quote (full paragraph)",
               "Term in context"]
MER_WIDTHS  = [11, 19, 36, 18, 12, 32, 8, 10, 95, 70]

def _write_term_sheet(ws, p, b, idx):
    is_mer = p["data_mode"] == "MER"
    sentence_only = p.get("quote_style") == "sentence"
    ncols = len(MER_HEADERS) if is_mer else 6
    _band(ws, 1, f'"{b["term"]}"', ncols, size=14)
    info = []
    if len(b.get("aliases") or []) > 1:
        info.append("Also counted as: " + " ~ ".join(b["aliases"][1:]))
    if b.get("scope"):
        info.append("Restricted to: " + b["scope"])
    if b.get("forms"):
        info.append("Wording found: " + ", ".join(f'{f["form"]} ({f["n"]})'
                                                  for f in b["forms"][:10]))
    info.append(f"Mode: {p.get('search_mode')}   |   Source: {b['source_label']}")
    r = 2
    _note(ws, r, "   |   ".join(info), ncols)
    r += 1
    if is_mer:
        for lab in b["rounds"]:
            c = ws.cell(row=r, column=1)
            _setv(c, _stat_line(b, lab))
            c.font = Font(bold=True, size=10, color=ROUND_INK.get(lab, "2C5F8A"))
            for j in range(1, ncols + 1):
                ws.cell(row=r, column=j).fill = ROUND_FILL.get(lab, NOTE_FILL)
            r += 1
    for w in b.get("warn", [])[:3]:
        _note(ws, r, "Note: " + w, ncols, fill=WARN_FILL, color="7A5A12", italic=True)
        r += 1
    r += 1
    hr = r
    if is_mer:
        rows = b["results"]
        if not rows:
            _setv(ws.cell(row=hr, column=1), "No mentions found.")
            return
        heads = list(MER_HEADERS)
        if sentence_only:
            heads[8] = "Quote (sentence with the term)"
        _write_header(ws, hr, heads, MER_WIDTHS, height=30)
        r = hr + 1
        for x in rows:
            _round_cell(ws.cell(row=r, column=1), x["Round"])
            c = ws.cell(row=r, column=2)
            _setv(c, x["Table"])
            _style(c, "table:" + x["Table"], fill=TABLE_FILL.get(x["Table"], NOTE_FILL),
                   alignment=WRAP)
            vals = [x["ref"], x["Jurisdiction"], x["MEASURE"], x["Indicator"],
                    x["Rating"], x["Criteria rating"]]
            for j, v in enumerate(vals, start=3):
                c = ws.cell(row=r, column=j)
                _setv(c, v)
                _style(c, "wrap", alignment=WRAP)
            quote = x["sentence"] if sentence_only else x["OBS_VALUE"]
            for j, v in ((9, quote), (10, x["context"])):
                c = ws.cell(row=r, column=j)
                _setv(c, _rich(v, x.get("variants") or x["matched"]))
                _style(c, "wrap", alignment=WRAP)
            r += 1
        _add_table(ws, f"tTerm{idx}", hr, r - 1, len(heads))
    else:
        heads = ["File", "Location", "Sentence with term", "Details", "Full text",
                 "Term in context"]
        rows = [x for fb in b["results"] for x in fb["rows"]]
        errs = [fb for fb in b["results"] if fb.get("error")]
        for fb in errs:
            _note(ws, hr, f"NOT SEARCHED: {fb['file']} - {fb['error']}", ncols,
                  fill=WARN_FILL, color="7A5A12")
            hr += 1
        if not rows:
            _setv(ws.cell(row=hr, column=1), "No mentions found.")
            return
        _write_header(ws, hr, heads, [30, 45, 70, 45, 90, 70], height=30)
        r = hr + 1
        for x in rows:
            vals = [x["File"], x["Location"],
                    _rich(x["Sentence with term"], x.get("variants") or x["matched"]),
                    x.get("Details", ""),
                    _rich(x["Full text"], x.get("variants") or x["matched"]),
                    _rich(x["context"], x.get("variants") or x["matched"])]
            for j, v in enumerate(vals, start=1):
                c = ws.cell(row=r, column=j)
                _setv(c, v)
                _style(c, "wrap", alignment=WRAP)
            r += 1
        _add_table(ws, f"tTerm{idx}", hr, r - 1, len(heads))
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)

def _write_all_hits(ws, p, blocks):
    is_mer = p["data_mode"] == "MER"
    if is_mer:
        heads = ["Search term", "Round", "Table", "Reference (for citing)", "Jurisdiction",
                 "MEASURE", "Indicator", "Rating", "Criteria rating", "Matched as",
                 "Sentence with term"]
        widths = [26, 11, 19, 36, 18, 12, 30, 8, 10, 18, 90]
    else:
        heads = ["Search term", "File", "Location", "Matched as", "Sentence with term",
                 "Details"]
        widths = [26, 30, 45, 18, 90, 50]
    _band(ws, 1, "All hits  -  every hit for every term, one row each", len(heads))
    _note(ws, 2, "One flat table for filtering, sorting and pivot tables. Full paragraphs "
                 "are on each term's own sheet.", len(heads), italic=True)
    hr = 4
    _write_header(ws, hr, heads, widths, height=30)
    r = hr + 1
    for b in blocks:
        rows = b["results"] if is_mer else [x for fb in b["results"] for x in fb["rows"]]
        for x in rows:
            if is_mer:
                vals = [b["term"], None, x["Table"], x["ref"], x["Jurisdiction"],
                        x["MEASURE"], x["Indicator"], x["Rating"], x["Criteria rating"],
                        ", ".join(x.get("variants") or [x["matched"]]), x["sentence"]]
            else:
                vals = [b["term"], x["File"], x["Location"],
                        ", ".join(x.get("variants") or [x["matched"]]),
                        x["Sentence with term"], x.get("Details", "")]
            for j, v in enumerate(vals, start=1):
                c = ws.cell(row=r, column=j)
                if is_mer and j == 2:
                    _round_cell(c, x["Round"])
                    continue
                _setv(c, v)
                _style(c, "wrap", alignment=WRAP)
            r += 1
    if r - 1 > hr:
        _add_table(ws, "tAllHits", hr, r - 1, len(heads))
    else:
        _setv(ws.cell(row=hr + 1, column=1), "No hits.")
    ws.freeze_panes = ws.cell(row=hr + 1, column=2)

def export_excel(p, blocks, export_dir):
    os.makedirs(export_dir, exist_ok=True)
    path = _free_path(export_dir, export_filename(p, blocks, "xlsx"))
    _TABLE_NAMES.clear()
    _STYLE_CACHE.clear()          # style records belong to one workbook
    is_mer = p["data_mode"] == "MER"
    labels = blocks[0]["rounds"] if (blocks and is_mer) else []
    wb = Workbook()
    ws0 = wb.active
    ws0.title = "Summary"
    used = set(RESERVED_SHEETS)
    names = [sanitize_sheet(b["term"], used) for b in blocks]
    _write_summary(ws0, p, blocks, names)
    if is_mer and blocks:
        _write_matrix(wb.create_sheet("Country matrix"), p, blocks)
        if len(labels) == 2:
            _write_compare(wb.create_sheet("Round comparison"), blocks)
    for i, (b, n) in enumerate(zip(blocks, names), start=1):
        _write_term_sheet(wb.create_sheet(n), p, b, i)
    _write_all_hits(wb.create_sheet("All hits"), p, blocks)
    wb.save(path)
    return path

def _free_path(folder, name):
    path = os.path.join(folder, name)
    stem, ext = os.path.splitext(name)
    n = 2
    while os.path.exists(path):
        path = os.path.join(folder, f"{stem} ({n}){ext}")
        n += 1
    return path

def export_filename(p, blocks, ext):
    """(date)_(name, else first search term)_FATF Search tool.(ext)"""
    name = str(p.get("report_name") or "").strip()
    if not name and blocks:
        name = blocks[0]["term"]
    name = re.sub(r'[\\/*?:\[\]"<>|\x00-\x1f]+', " ", name).strip(" .") or "search"
    name = re.sub(r"\s+", " ", name)[:60]
    return "%s_%s_FATF Search tool.%s" % (datetime.now().strftime("%Y-%m-%d"), name, ext)


# =============================================================================
# WORD REPORT  (built by hand - no extra packages needed)
# =============================================================================
import zipfile
from xml.sax.saxutils import escape as _esc

WNS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
       'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

def _run(text, bold=False, italic=False, color=None, size=None):
    rpr = ""
    if bold: rpr += "<w:b/>"
    if italic: rpr += "<w:i/>"
    if color: rpr += '<w:color w:val="%s"/>' % color
    if size: rpr += '<w:sz w:val="%s"/><w:szCs w:val="%s"/>' % (size, size)
    rpr = "<w:rPr>%s</w:rPr>" % rpr if rpr else ""
    text = _esc(_XML_BAD.sub("", str(text)))
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, text)

def _para(runs, style=None, after=120, before=0, align=None, shade=None, border=False):
    ppr = ""
    if style: ppr += '<w:pStyle w:val="%s"/>' % style
    if shade: ppr += '<w:shd w:val="clear" w:fill="%s"/>' % shade
    if border:
        ppr += ('<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" '
                'w:color="C9D7E6"/></w:pBdr>')
    if align: ppr += '<w:jc w:val="%s"/>' % align
    ppr += '<w:spacing w:before="%d" w:after="%d" w:line="264" w:lineRule="auto"/>' % (before, after)
    return "<w:p><w:pPr>%s</w:pPr>%s</w:p>" % (ppr, "".join(runs))

def _highlight_runs(text, variants, base_size="21"):
    """Paragraph runs with each matched variant in bold."""
    text = str(text)
    vs = sorted({v for v in (variants or []) if v}, key=len, reverse=True)
    if not vs:
        return [_run(text, size=base_size)]
    pat = re.compile("|".join(re.escape(v) for v in vs), re.IGNORECASE)
    out, last = [], 0
    for m in pat.finditer(text):
        if m.start() > last:
            out.append(_run(text[last:m.start()], size=base_size))
        out.append(_run(m.group(), bold=True, size=base_size))
        last = m.end()
    if last < len(text):
        out.append(_run(text[last:], size=base_size))
    return out

def _table(rows, widths, header=True):
    grid = "".join('<w:gridCol w:w="%d"/>' % w for w in widths)
    borders = "".join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="C9D7E6"/>' % e
                      for e in ("top", "left", "bottom", "right", "insideH", "insideV"))
    xml = ['<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:tblBorders>%s</w:tblBorders>'
           '</w:tblPr><w:tblGrid>%s</w:tblGrid>' % (sum(widths), borders, grid)]
    for ri, row in enumerate(rows):
        xml.append("<w:tr>")
        for ci, cellv in enumerate(row):
            txt, bold = (cellv if isinstance(cellv, tuple) else (cellv, False))
            hdr = header and ri == 0
            shade = "1E2F4D" if hdr else ("F4F8FC" if ri % 2 == 0 else None)
            shd = '<w:shd w:val="clear" w:fill="%s"/>' % shade if shade else ""
            xml.append('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s'
                       '<w:tcMar><w:top w:w="70" w:type="dxa"/><w:left w:w="110" w:type="dxa"/>'
                       '<w:bottom w:w="70" w:type="dxa"/><w:right w:w="110" w:type="dxa"/>'
                       '</w:tcMar></w:tcPr>' % (widths[ci], shd))
            xml.append(_para([_run(txt, bold=bold or hdr,
                                   color="FFFFFF" if hdr else None, size="19")], after=0))
            xml.append("</w:tc>")
        xml.append("</w:tr>")
    xml.append("</w:tbl>")
    xml.append(_para([_run("")], after=80))
    return "".join(xml)

_STYLE_PARTS = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:styles ' + WNS + '>',
    '<w:docDefaults><w:rPrDefault><w:rPr>'
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="21"/>'
    '</w:rPr></w:rPrDefault></w:docDefaults>',
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
    '<w:pPr><w:outlineLvl w:val="0"/></w:pPr>'
    '<w:rPr><w:b/><w:color w:val="1E2F4D"/><w:sz w:val="30"/></w:rPr></w:style>',
    '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>'
    '<w:pPr><w:outlineLvl w:val="1"/></w:pPr>'
    '<w:rPr><w:b/><w:color w:val="2C5F8A"/><w:sz w:val="25"/></w:rPr></w:style>',
    '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/>'
    '<w:pPr><w:outlineLvl w:val="2"/></w:pPr>'
    '<w:rPr><w:b/><w:color w:val="20262E"/><w:sz w:val="22"/></w:rPr></w:style>',
    '</w:styles>',
]
_STYLES = "".join(_STYLE_PARTS)

_XML_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")

def _write_docx(path, body_xml, footer_text=""):
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
             '</Relationships>')
    fr = '<w:rPr><w:color w:val="8B9CB3"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
    footer = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + WNS + '>'
              '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
              '<w:r>' + fr + '<w:t xml:space="preserve">' + _esc(_XML_BAD.sub("", footer_text[:120]))
              + '  \u00b7  page </w:t></w:r>'
              '<w:r>' + fr + '<w:fldChar w:fldCharType="begin"/></w:r>'
              '<w:r>' + fr + '<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
              '<w:r>' + fr + '<w:fldChar w:fldCharType="separate"/></w:r>'
              '<w:r>' + fr + '<w:t>1</w:t></w:r>'
              '<w:r>' + fr + '<w:fldChar w:fldCharType="end"/></w:r>'
              '</w:p></w:ftr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + WNS + '>'
           '<w:body>' + body_xml +
           '<w:sectPr><w:footerReference w:type="default" r:id="rId2"/>'
           '<w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1200" w:right="1200" w:bottom="1200" w:left="1200" '
           'w:header="600" w:footer="600" w:gutter="0"/></w:sectPr>'
           '</w:body></w:document>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/styles.xml", _STYLES)
        z.writestr("word/document.xml", doc)
        z.writestr("word/footer1.xml", footer)
    return path

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


ROUND_COLOR = {"4th round": "5B4B8A", "5th round": "9A6A00"}

def _count(b):
    res = b.get("results") or []
    if res and isinstance(res[0], dict) and "rows" in res[0]:
        return sum(len(fb["rows"]) for fb in res)
    return len(res)

def _round_tag(lab):
    return _run(lab.upper() + "  ", bold=True, size="17", color=ROUND_COLOR.get(lab, "2C5F8A"))

def _cov_txt(cov):
    return "%s of %s (%s%%)" % (cov.get("mentioned", 0), cov.get("total", 0), cov.get("pct", 0))

def build_word_report(p, blocks, export_dir):
    os.makedirs(export_dir, exist_ok=True)
    path = _free_path(export_dir, export_filename(p, blocks, "docx"))
    title = (str(p.get("report_name") or "").strip()
             or (blocks[0]["term"] if blocks else "Search results"))
    src_label = blocks[0]["source_label"] if blocks else ""
    today = datetime.now().strftime("%d %B %Y")
    is_mer = p["data_mode"] == "MER"
    labels = blocks[0]["rounds"] if (blocks and is_mer) else []
    two = len(labels) == 2
    sentence_only = (p.get("quote_style") == "sentence")
    B = []

    # ---------- cover ----------
    B.append(_para([_run("FATF SEARCH TOOL", bold=True, color="2C5F8A", size="19")], after=60))
    B.append(_para([_run(title, bold=True, color="1E2F4D", size="44")], after=80))
    B.append(_para([_run("Mentions in " + src_label.replace("MER ", "")
                         + (" Mutual Evaluation Reports" if is_mer else "")
                         + ", with context and statistics",
                         italic=True, color="54677F", size="23")], after=200, border=True))
    cover = [[("Prepared", True), today], [("Source", True), src_label]]
    if is_mer:
        df_all = load_mer(p["round"])
        gf = dict(p.get("filters") or {})
        for lab in labels:
            tot = int(apply_filters(df_all[df_all["Round"] == lab], gf)["Jurisdiction"].nunique())
            n = sum(1 for b in blocks for x in b["results"] if x["Round"] == lab)
            cover.append([(lab[0].upper() + lab[1:], True),
                          "%d mention%s  ·  %d jurisdictions assessed" %
                          (n, "" if n == 1 else "s", tot)])
    else:
        cover.append([("Total mentions", True), str(sum(_count(b) for b in blocks))])
    cover.append([("Terms searched", True), ", ".join(b["term"] for b in blocks)[:700]])
    cover.append([("Quotes shown as", True),
                  "The sentence containing the term" if sentence_only
                  else "The full paragraph"])
    B.append(_table(cover, [2000, 7400], header=False))

    B.append(_para([_run("How to use this document", bold=True, color="1E2F4D", size="21")],
                   before=200, after=80))
    how = ["Section 1 gives the headline numbers for every term.",
           "Each term then has its own section: statistics first, then every mention "
           "grouped by country, with the search term in bold.",
           "Every quote carries a reference line (round, country, Immediate Outcome or "
           "Recommendation, rating) so it can be cited directly.",
           "A mention means the words appear in the report. It does not by itself mean the "
           "report is discussing the body you have in mind, so read the wording before quoting.",
           "In Word, View > Navigation Pane lists every section for quick jumping."]
    if two:
        how.insert(1, "The 4th and 5th rounds are always reported separately, and every "
                      "quote is labelled with its round. Figures are never added across "
                      "rounds, because the rounds cover different sets of countries under "
                      "different methodologies (2013 and 2022).")
    for line in how:
        B.append(_para([_run("•  ", color="2C5F8A", size="21"), _run(line, size="21")],
                       after=70))

    # ---------- 1. overview ----------
    B.append(_para([_run("1.  Overview", bold=True, color="1E2F4D", size="30")],
                   style="Heading1", before=340, after=140, border=True))
    if is_mer and two:
        rows = [[("Term", True), ("4th round mentions", True), ("4th round countries", True),
                 ("5th round mentions", True), ("5th round countries", True)]]
        for b in blocks:
            row = [b["term"][:40]]
            for lab in labels:
                n = sum(1 for x in b["results"] if x["Round"] == lab)
                row += [str(n), _cov_txt(b["coverage"][lab])]
            rows.append(row)
        B.append(_table(rows, [2800, 1350, 1900, 1350, 1900]))
    elif is_mer:
        lab = labels[0]
        rows = [[("Term", True), ("Mentions", True), ("Countries", True), ("Effect.", True),
                 ("Risk & ctx", True), ("Tech. comp.", True)]]
        for b in blocks:
            n2 = Counter(x["Table"] for x in b["results"])
            rows.append([b["term"][:40], str(len(b["results"])),
                         _cov_txt(b["coverage"][lab]),
                         str(n2.get("Effectiveness", 0)), str(n2.get("Risk and Context", 0)),
                         str(n2.get("Technical Compliance", 0))])
        B.append(_table(rows, [2700, 1100, 1900, 1150, 1250, 1300]))
    else:
        rows = [[("Term", True), ("Mentions", True), ("Where they are", True)]]
        for b in blocks:
            files = ", ".join("%s (%d)" % (fb["file"][:26], len(fb["rows"]))
                              for fb in b["results"] if fb["rows"])
            rows.append([b["term"][:40], str(_count(b)), files[:95]])
        B.append(_table(rows, [2700, 1200, 5500]))

    # ---------- per term ----------
    for si, b in enumerate(blocks, start=2):
        B.append(_para([_run("%d.  %s" % (si, b["term"]), bold=True, color="1E2F4D", size="30")],
                       style="Heading1", before=360, after=120, border=True))
        if len(b.get("aliases", [])) > 1:
            B.append(_para([_run("Also counted as: ", bold=True, size="20", color="54677F"),
                            _run(", ".join(b["aliases"][1:]), size="20", color="54677F")],
                           after=90))
        if b.get("scope"):
            B.append(_para([_run("Restricted to " + b["scope"], italic=True,
                                 size="20", color="54677F")], after=90))

        rows_all = (b["results"] if is_mer
                    else [r for fb in b["results"] for r in fb["rows"]])
        if not rows_all:
            B.append(_para([_run("No mentions found.", italic=True, color="54677F")], after=160))
            continue

        B.append(_para([_run("Statistics", bold=True, color="2C5F8A", size="25")],
                       style="Heading2", before=170, after=90))
        if is_mer:
            head = [("Measure", True)] + [(lab[0].upper() + lab[1:], True) for lab in labels]
            stat_rows = [head]
            def per(lab):
                rr = [x for x in rows_all if x["Round"] == lab]
                cov = b["coverage"][lab]
                n2 = Counter(x["Table"] for x in rr)
                meas = Counter(pretty_measure(x["MEASURE"]) for x in rr)
                rat = Counter(x["Rating"] for x in rr if x.get("Rating") and x["Rating"] != "NA")
                top = Counter(x["Jurisdiction"] for x in rr)
                return [str(len(rr)), _cov_txt(cov),
                        str(n2.get("Effectiveness", 0)), str(n2.get("Risk and Context", 0)),
                        str(n2.get("Technical Compliance", 0)),
                        ", ".join("%s (%d)" % kv for kv in meas.most_common(5)) or "-",
                        ", ".join("%s (%d)" % kv for kv in rat.most_common()) or "-",
                        ", ".join("%s (%d)" % kv for kv in top.most_common(6)) or "-"]
            cols = [per(lab) for lab in labels]
            names = ["Mentions", "Countries mentioning it", "In effectiveness chapters",
                     "In risk and context", "In technical compliance", "Comes up most under",
                     "Ratings on those paragraphs", "Most mentions"]
            for k, nm in enumerate(names):
                stat_rows.append([nm] + [c[k] for c in cols])
            widths = [2600, 3400, 3400] if two else [3000, 6400]
            B.append(_table(stat_rows, widths))
            if two and b.get("round_compare"):
                B.append(_para([_run("Countries assessed in both rounds", bold=True,
                                     size="20", color="1E2F4D")], before=60, after=60))
                cmp_rows = [[("Country", True), ("4th round", True), ("5th round", True),
                             ("Pattern", True)]]
                for c in b["round_compare"]:
                    cmp_rows.append([c["name"], str(c["r4"]), str(c["r5"]), c["pattern"]])
                B.append(_table(cmp_rows, [3200, 1600, 1600, 3000]))
        if b.get("forms"):
            B.append(_para([_run("Wording actually found: ", bold=True, size="20"),
                            _run(", ".join("%s (%d)" % (f["form"], f["n"])
                                           for f in b["forms"][:10]), size="20")], after=80))
        for w in b.get("warn", [])[:3]:
            B.append(_para([_run("Note: " + w, italic=True, size="20", color="A8873D")], after=70))

        if is_mer:
            for lab in labels:
                rr = [x for x in rows_all if x["Round"] == lab]
                heading = "Mentions" + (" — " + lab if two else "")
                B.append(_para([_run(heading, bold=True, size="25",
                                     color=ROUND_COLOR.get(lab, "2C5F8A") if two else "2C5F8A"),
                                _run("   (%d)" % len(rr), size="20", color="54677F")],
                               style="Heading2", before=220, after=90))
                if not rr:
                    B.append(_para([_run("No mentions in the " + lab + ".", italic=True,
                                         color="54677F", size="20")], after=100))
                    continue
                bycountry = defaultdict(list)
                for x in rr:
                    bycountry[x["Jurisdiction"]].append(x)
                for country in sorted(bycountry):
                    items = bycountry[country]
                    B.append(_para([_run(country + "  ", bold=True, size="22"),
                                    _run("(%d mention%s%s)" % (len(items),
                                         "s" if len(items) > 1 else "",
                                         ", " + lab if two else ""),
                                         size="19", color="54677F")],
                                   style="Heading3", before=180, after=70))
                    for x in items:
                        ref = pretty_measure(x["MEASURE"])
                        if x.get("Rating") and x["Rating"] != "NA":
                            ref += "  ·  rated " + str(x["Rating"])
                        if x.get("Criteria rating") and x["Criteria rating"] != "NA":
                            ref += "  ·  criterion " + str(x["Criteria rating"])
                        ref += "  ·  " + str(x["Table"])
                        if x.get("Indicator"):
                            ref += "  ·  " + str(x["Indicator"])[:70]
                        B.append(_para([_round_tag(lab),
                                        _run(ref, size="18", color="2C5F8A", bold=True)],
                                       after=40, before=110))
                        text = x["sentence"] if sentence_only else x["OBS_VALUE"]
                        B.append(_para(_highlight_runs(text, x.get("variants")),
                                       after=110, shade="F7FAFD"))
        else:
            B.append(_para([_run("Mentions", bold=True, color="2C5F8A", size="25")],
                           style="Heading2", before=200, after=90))
            for fb in b["results"]:
                if not fb["rows"]:
                    continue
                B.append(_para([_run(fb["file"], bold=True, size="22")],
                               style="Heading3", before=180, after=70))
                for x in fb["rows"]:
                    B.append(_para([_run(x["Location"], size="18", color="2C5F8A", bold=True)],
                                   after=40, before=110))
                    if x.get("Details"):
                        B.append(_para([_run(str(x["Details"])[:400], size="18",
                                             color="54677F", italic=True)], after=40))
                    text = x["Sentence with term"] if sentence_only else x["Full text"]
                    B.append(_para(_highlight_runs(text, x.get("variants")),
                                   after=110, shade="F7FAFD"))

    # ---------- method ----------
    B.append(_para([_run("How the search was done", bold=True, color="1E2F4D", size="30")],
                   style="Heading1", before=360, after=120, border=True))
    method = [
        "This is a literal text search of the published reports. A paragraph appears here "
        "only if the term genuinely occurs in it; nothing is inferred or summarised.",
        "Capitals, spacing, hyphens and accents are ignored, so “ship-based” and "
        "“ship based” count as the same term.",
        "Abbreviations written in capitals are matched as whole words, so they are not found "
        "inside longer words.",
        "Where a body has more than one name, all its names count as one term; the "
        "alternatives are listed under each section heading.",
        "A term not appearing means the report did not use those words. It does not mean the "
        "country has no such body or framework.",
    ]
    if two:
        method.append("4th round reports were assessed under the 2013 Methodology and 5th "
                      "round reports under the 2022 Methodology. Differences between the "
                      "rounds can reflect the methodology or the wording as much as the "
                      "country, so compare the paragraphs, not only the counts.")
    for line in method:
        B.append(_para([_run("•  ", color="2C5F8A", size="20"), _run(line, size="20")],
                       after=70))
    B.append(_para([_run("Generated by the FATF Search Tool " + VERSION + " on " + today + ".",
                         italic=True, size="18", color="54677F")], before=200, after=0))

    return _write_docx(path, "".join(B), footer_text="FATF Search Tool  ·  " + title)


# =============================================================================
# API ROUTES
# =============================================================================
def _err(e):
    return jsonify({"ok": False, "error": f"{e}", "trace": traceback.format_exc()})

@app.route("/api/load", methods=["POST"])
def api_load():
    try:
        p = request.get_json(force=True)
        set_paths(p.get("paths"))
        if p["data_mode"] == "MER":
            per = []
            for r in rounds_of(p["round"]):
                d = load_round(r)
                per.append({"round": ROUND_LABEL[r], "rows": int(len(d)),
                            "n_jurisdictions": int(d["Jurisdiction"].nunique()),
                            "counts": {k: int(v) for k, v in d["Table"].value_counts().items()},
                            "folder": PATHS["round_" + r]})
            df = load_mer(p["round"])
            return jsonify({
                "ok": True,
                "rounds": per,
                "total": int(len(df)),
                "n_jurisdictions": int(df["Jurisdiction"].nunique()),
                "jurisdictions": sorted(df["Jurisdiction"].unique().tolist()),
                "indicators": sorted(x for x in df["Indicator"].unique().tolist() if x),
                "ratings": [{"code": c, "label": rating_option(c)}
                            for c in sorted(x for x in df["Rating"].unique().tolist() if x)],
                "crit_ratings": [{"code": c, "label": rating_option(c)}
                                 for c in sorted(x for x in df["Criteria rating"].unique().tolist() if x)],
            })
        ds = load_other()
        files = [{"file": d["file"],
                  "rows": (0 if d.get("records") is None else len(d["records"])),
                  "error": d.get("error")} for d in ds]
        return jsonify({"ok": True, "files": files, "folder": PATHS["other"]})
    except Exception as e:
        return _err(e)

def _summary(p, b):
    """Counts and statistics for one term card (rows are fetched separately)."""
    out = {"term": b["term"], "aliases": b.get("aliases", []), "scope": b.get("scope"),
           "notes": b.get("notes", []), "how": b.get("how", []),
           "warn": b.get("warn", []), "forms": b.get("forms", [])}
    if p["data_mode"] == "MER":
        res = b["results"]
        out["total"] = len(res)
        out["rounds"] = b["rounds"]
        out["coverage"] = b["coverage"]
        out["round_compare"] = b.get("round_compare", [])
        out["by_round"] = {lab: sum(1 for x in res if x["Round"] == lab) for lab in b["rounds"]}
        out["sections"] = [{"key": t, "name": t,
                            "total": sum(1 for x in res if x["Table"] == t),
                            "by_round": {lab: sum(1 for x in res if x["Table"] == t
                                                  and x["Round"] == lab)
                                         for lab in b["rounds"]}}
                           for t in TABLES]
    else:
        out["total"] = sum(len(fb["rows"]) for fb in b["results"])
        out["rounds"] = []
        out["sections"] = [{"key": str(i), "name": fb["file"], "total": len(fb["rows"]),
                            "error": fb.get("error"), "by_round": {}}
                           for i, fb in enumerate(b["results"])]
    return out

@app.route("/api/search", methods=["POST"])
def api_search():
    try:
        p = request.get_json(force=True)
        sid = p.get("sid") or uuid.uuid4().hex
        with _LOCK:
            _STORE.pop(sid, None)                 # an explicit search always re-runs
        raw = get_raw(p, sid)
        blocks = finalize(p, raw)
        return jsonify({"ok": True, "sid": sid, "data_mode": p["data_mode"],
                        "rounds": round_labels(p["round"]) if p["data_mode"] == "MER" else [],
                        "page_size": PAGE_SIZE,
                        "results": [_summary(p, b) for b in blocks]})
    except Exception as e:
        return _err(e)

@app.route("/api/progress")
def api_progress():
    sid = request.args.get("sid", "")
    return jsonify(_PROGRESS.get(sid) or {"done": 0, "total": 0, "term": "", "finished": False})

@app.route("/api/rows", methods=["POST"])
def api_rows():
    """One page of hits for one section (table or file) of one term card."""
    try:
        p = request.get_json(force=True)
        with _LOCK:
            ent = _STORE.get(p.get("sid"))
        if not ent:
            return jsonify({"ok": False, "error": "This search has expired - please run it again."})
        b = ent["raw"][int(p.get("term_index", 0))]
        sec = str(p.get("section", ""))
        rnd = p.get("round") or "all"
        if ent["data_mode"] == "MER":
            rows = [x for x in b["results"] if x["Table"] == sec
                    and (rnd == "all" or x["Round"] == rnd)]
        else:
            rows = b["results"][int(sec)]["rows"]
        off = max(0, int(p.get("offset", 0)))
        lim = max(1, min(500, int(p.get("limit", PAGE_SIZE))))
        return jsonify({"ok": True, "total": len(rows),
                        "rows": [_slim_row(x) for x in rows[off:off + lim]]})
    except Exception as e:
        return _err(e)

@app.route("/api/interpret", methods=["POST"])
def api_interpret():
    """Explain how each typed line will be searched, without running a search."""
    try:
        p = request.get_json(force=True)
        known = None
        if p.get("data_mode") == "MER":
            try:
                set_paths(p.get("paths"))
                known = sorted(load_mer(p["round"])["Jurisdiction"].unique().tolist())
            except Exception:
                known = None
        out, seen_labels = [], {}
        for line in [t.strip() for t in p.get("terms", []) if t.strip()]:
            tp = prepare_term(line, p.get("search_mode", "strict"),
                              bool(p.get("caps_only", False)), known)
            if tp is None:
                out.append({"input": line, "label": "", "aliases": [], "scope": None,
                            "notes": [], "how": [],
                            "warn": ["Nothing searchable on this line (no letters or "
                                     "numbers), so it will be skipped."]})
                continue
            warn = list(tp["warn"])
            key = tp["label"].lower()
            if key in seen_labels:
                warn.append(f'Duplicate of an earlier line ("{seen_labels[key]}"). '
                            f'It will get its own sheet named with a _2 suffix.')
            else:
                seen_labels[key] = tp["input"]
            out.append({"input": tp["input"], "label": tp["label"],
                        "aliases": tp["aliases"], "scope": tp["scope"],
                        "notes": tp["notes"], "how": tp["how"], "warn": warn})
        return jsonify({"ok": True, "terms": out})
    except Exception as e:
        return _err(e)

@app.route("/api/export", methods=["POST"])
def api_export():
    try:
        p = request.get_json(force=True)
        sid = p.get("sid")
        raw = get_raw(p, sid)                 # reuses the on-screen search
        blocks = finalize(p, raw, p.get("excluded") or [])
        export_dir = (p.get("export_dir") or "").strip().strip('"') or DEFAULT_EXPORT_DIR
        fmt = (p.get("format") or "xlsx").lower()
        paths = []
        with _EXPORT_LOCK:                    # one export at a time
            if fmt in ("xlsx", "both"):
                paths.append(export_excel(p, blocks, export_dir))
            if fmt in ("docx", "both"):
                paths.append(build_word_report(p, blocks, export_dir))
        return jsonify({"ok": True, "path": "\n".join(paths), "paths": paths,
                        "folder": os.path.abspath(export_dir)})
    except PermissionError as e:
        return jsonify({"ok": False, "error": f"Could not save - is a file with the same "
                        f"name open in Excel or Word? Close it and try again. ({e})"})
    except Exception as e:
        return _err(e)

@app.route("/api/open_folder", methods=["POST"])
def api_open_folder():
    """Open the export folder in Windows Explorer / Finder."""
    try:
        p = request.get_json(force=True)
        folder = (p.get("folder") or "").strip() or DEFAULT_EXPORT_DIR
        os.makedirs(folder, exist_ok=True)
        if hasattr(os, "startfile"):
            os.startfile(folder)                     # Windows
        else:
            import subprocess, sys
            subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", folder])
        return jsonify({"ok": True})
    except Exception as e:
        return _err(e)


# =============================================================================
# FRONT END (single page)
# =============================================================================
PAGE = r"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>MER Search — FATF Mutual Evaluation Intelligence</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Libre+Franklin:wght@600;700;800&family=Public+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root{
  --canvas:#f1f5fa; --panel:#ffffff; --panel2:#f6f9fc;
  --ink:#1e2f4d; --muted:#54677f; --faint:#8b9cb3;
  --border:#dde6f0; --border2:#c9d7e6;
  --accent:#2c5f8a; --accent-strong:#1f4a70; --accent-soft:#e3edf6;
  --steel:#7fa6c9;
  --mark:#fdf0b4; --mark-ink:#54470a;
  --eff:#5b87b5; --eff-soft:#e4eef7;
  --rc:#6fa88a;  --rc-soft:#e7f2ec;
  --tc:#b39a63;  --tc-soft:#f5efe1;
  --oth:#8b99ab; --oth-soft:#eef1f5;
  --good:#4d8b6f; --good-soft:#e4f1ea;
  --mid:#4a7cab;  --mid-soft:#e4edf6;
  --warn:#a8873d; --warn-soft:#f7efdb;
  --bad:#a85a68;  --bad-soft:#f6e7ea;
  --na:#93a0b0;   --na-soft:#edf0f4;
  --shadow:0 1px 2px rgba(30,47,77,.06),0 5px 16px rgba(30,47,77,.06);
  --radius:12px;
}
body.dark{
  --canvas:#0d1725; --panel:#152238; --panel2:#1b2b44;
  --ink:#e6edf6; --muted:#9db0c6; --faint:#6b7f97;
  --border:#24334a; --border2:#31445f;
  --accent:#6ea3d4; --accent-strong:#8db8e0; --accent-soft:#1c3350;
  --steel:#6ea3d4;
  --mark:#5d5013; --mark-ink:#fdf0b4;
  --eff:#7fa9d4; --eff-soft:#1d2f47;
  --rc:#83bd9e;  --rc-soft:#1a3229;
  --tc:#cbb079;  --tc-soft:#37301c;
  --oth:#9db0c6; --oth-soft:#20304a;
  --good:#6cbf97; --good-soft:#183227;
  --mid:#7fa9d4;  --mid-soft:#1d2f47;
  --warn:#d4b25e; --warn-soft:#383018;
  --bad:#d48a97;  --bad-soft:#3a2229;
  --na:#7e8fa3;   --na-soft:#20304a;
  --shadow:0 1px 2px rgba(0,0,0,.4),0 6px 18px rgba(0,0,0,.3);
}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;font-family:'Public Sans','Segoe UI',system-ui,sans-serif;background:var(--canvas);
  color:var(--ink);font-size:13.5px;line-height:1.5;
  background-image:linear-gradient(180deg, rgba(44,95,138,.05), transparent 320px);}
body.dark{background-image:linear-gradient(180deg, rgba(110,163,212,.06), transparent 320px);}

/* ---------- topbar ---------- */
header{display:flex;align-items:center;gap:14px;padding:0 22px;height:58px;
  background:var(--panel);border-bottom:1px solid var(--border);position:sticky;top:0;z-index:20}
.wordmark{display:flex;align-items:baseline;gap:10px}
.wordmark .glyph{width:26px;height:26px;border-radius:6px;align-self:center;flex-shrink:0;
  background:var(--ink);position:relative}
.wordmark .glyph::after{content:'';position:absolute;inset:7px;border:2px solid rgba(255,255,255,.85);
  border-radius:3px;border-right-color:transparent;border-bottom-color:transparent;transform:rotate(-45deg)}
.wordmark h1{font-family:'Libre Franklin','Segoe UI',sans-serif;font-size:16.5px;font-weight:800;
  margin:0;letter-spacing:-.01em}
.wordmark .eyebrow{font-size:10px;font-weight:600;letter-spacing:.14em;color:var(--faint);text-transform:uppercase}
.srcchip{margin-left:auto;display:none;align-items:center;gap:7px;padding:5px 12px;border-radius:99px;
  background:var(--accent-soft);color:var(--accent);font-weight:600;font-size:12px;white-space:nowrap}
.srcchip .dot{width:7px;height:7px;border-radius:50%;background:var(--accent)}
.spinner{display:none;width:15px;height:15px;border:2.5px solid var(--border2);
  border-top-color:var(--accent);border-radius:50%;animation:spin .7s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.iconbtn{border:1px solid var(--border);background:var(--panel);color:var(--muted);
  border-radius:8px;padding:6px 12px;cursor:pointer;font-size:12.5px;font-weight:600;font-family:inherit}
.iconbtn:hover{border-color:var(--border2);color:var(--ink)}

/* ---------- layout ---------- */
.layout{display:flex;gap:16px;padding:18px 22px;align-items:flex-start;max-width:1560px;margin:0 auto}
.sidebar{width:322px;flex-shrink:0;display:flex;flex-direction:column;gap:14px;
  position:sticky;top:76px;max-height:calc(100vh - 94px);overflow-y:auto;padding-right:2px}
.sidebar::-webkit-scrollbar{width:8px}
.sidebar::-webkit-scrollbar-thumb{background:var(--border2);border-radius:8px}
.main{flex:1;min-width:0}
.card{background:var(--panel);border:1px solid var(--border);border-radius:var(--radius);
  padding:16px;box-shadow:var(--shadow)}
.card h3{margin:0 0 12px;font-size:10.5px;font-weight:700;text-transform:uppercase;
  letter-spacing:.13em;color:var(--faint)}
label{display:block;margin:10px 0 5px;font-weight:600;font-size:12px;color:var(--muted)}
select,textarea,input[type=text]{width:100%;padding:8px 10px;border:1px solid var(--border);
  border-radius:8px;background:var(--panel2);color:var(--ink);font-size:13px;font-family:inherit;
  transition:border-color .15s, box-shadow .15s}
select:focus,textarea:focus,input:focus{outline:none;border-color:var(--accent);
  box-shadow:0 0 0 3px color-mix(in srgb, var(--accent) 18%, transparent)}
select[multiple]{height:108px}
textarea{height:180px;resize:vertical;font-family:ui-monospace,Consolas,monospace;font-size:12.5px}
.pathtoggle{font-size:11.5px;color:var(--accent);cursor:pointer;text-decoration:underline;
  text-underline-offset:2px;font-weight:600}
.exbtn{border:0;background:transparent;color:var(--faint);cursor:pointer;font-size:15px;
  line-height:1;padding:2px 5px;border-radius:6px;font-family:inherit}
.exbtn:hover{background:var(--bad-soft);color:var(--bad)}
tr.dropped td{opacity:.32;text-decoration:line-through}
tr.dropped .exbtn{text-decoration:none;opacity:1;color:var(--good)}
.dropnote{font-size:11.5px;color:var(--warn);font-weight:700;padding:6px 20px;
  background:var(--warn-soft);border-bottom:1px solid var(--border);display:none}
.termcount{margin-top:5px;font-size:12px;font-weight:700;color:var(--accent);
  background:var(--accent-soft);display:inline-block;padding:3px 11px;border-radius:99px}
input[type=range]{width:100%;accent-color:var(--accent)}

/* ---------- segmented ---------- */
.seg{display:flex;background:var(--panel2);border:1px solid var(--border);border-radius:9px;padding:3px;gap:3px}
.seg button{flex:1;padding:7px 4px;border:0;border-radius:7px;background:transparent;color:var(--muted);
  cursor:pointer;font-weight:600;font-size:12.5px;font-family:inherit;transition:all .15s}
.seg button:hover{color:var(--ink)}
.seg button.on{background:var(--panel);color:var(--accent);
  box-shadow:0 1px 3px rgba(16,24,38,.12);font-weight:700}

/* ---------- buttons ---------- */
.btn{padding:10px 18px;border:0;border-radius:9px;cursor:pointer;font-weight:700;
  font-size:13.5px;font-family:inherit;transition:all .15s}
.btn.primary{background:var(--accent);color:#fff}
.btn.primary:hover{background:var(--accent-strong);transform:translateY(-1px);
  box-shadow:0 4px 12px color-mix(in srgb, var(--accent) 35%, transparent)}
.btn.ghost{background:transparent;color:var(--accent);border:1.5px solid var(--accent)}
.btn.ghost:hover{background:var(--accent-soft)}
.btn.grey{background:var(--panel2);color:var(--muted);border:1px solid var(--border)}
.btn.grey:hover{color:var(--ink)}
.btn:disabled{opacity:.45;cursor:not-allowed;transform:none!important;box-shadow:none!important}
.btn:focus-visible,.seg button:focus-visible,.iconbtn:focus-visible{outline:2px solid var(--accent);outline-offset:2px}

.row{display:flex;gap:10px;flex-wrap:wrap}
.status{font-size:12.5px;color:var(--muted);margin-top:10px;white-space:pre-wrap;line-height:1.6}
.status.err{color:var(--bad)}
.hint{font-size:11.5px;color:var(--faint);margin-top:5px;line-height:1.5}
.checkline{display:flex;align-items:center;gap:8px;margin:6px 0;font-size:13px;cursor:pointer}
.checkline input{width:15px;height:15px;accent-color:var(--accent);cursor:pointer}
.checkline .swatch{width:9px;height:9px;border-radius:3px}

/* ---------- results ---------- */
.termcard{background:var(--panel);border:1px solid var(--border);border-radius:var(--radius);
  box-shadow:var(--shadow);margin-bottom:18px;overflow:hidden;
  animation:rise .35s ease both}
@keyframes rise{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
@media (prefers-reduced-motion: reduce){
  .termcard{animation:none}
  .covfill{transition:none!important}
  html{scroll-behavior:auto}
}
.termhead{display:flex;align-items:center;gap:14px;padding:15px 20px;cursor:pointer;
  border-bottom:1px solid var(--border);flex-wrap:wrap}
.termhead:hover{background:var(--panel2)}
.termhead .arrow{color:var(--faint);font-size:11px;transition:transform .15s}
.termhead.closed .arrow{transform:rotate(-90deg)}
.termname{font-family:'Libre Franklin','Segoe UI',sans-serif;font-size:18.5px;font-weight:700;
  letter-spacing:-.01em}
.termname::before{content:'“';color:var(--steel)}
.termname::after{content:'”';color:var(--steel)}
.hitpill{padding:3px 12px;border-radius:99px;background:var(--accent);color:#fff;
  font-weight:700;font-size:12px}
.hitpill.zero{background:var(--na-soft);color:var(--na)}
.aliaswrap{display:inline-flex;gap:5px;flex-wrap:wrap;align-items:center}
.scopetag{padding:2px 9px;border-radius:99px;background:var(--tc-soft);color:var(--tc);
  font-size:11.5px;font-weight:700;border:1px solid var(--border2)}
.formsbar{display:flex;gap:6px;flex-wrap:wrap;padding:9px 20px;background:var(--panel2);
  border-bottom:1px solid var(--border);align-items:center}
.formsbar .lbl{font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;
  color:var(--faint);margin-right:2px}
.form{padding:2px 9px;border-radius:6px;background:var(--mark);color:var(--mark-ink);
  font-size:11.5px;font-weight:700;font-family:ui-monospace,Consolas,monospace}
.form i{font-style:normal;font-weight:600;opacity:.65;margin-left:3px}
.warnbar{padding:9px 20px;background:var(--warn-soft);border-bottom:1px solid var(--border);
  color:var(--warn);font-size:12px;font-weight:600;line-height:1.55}
.warnbar div+div{margin-top:3px}
#interp{margin-top:12px;display:none}
.itable{width:100%;border-collapse:separate;border-spacing:0;font-size:12.5px;
  border:1px solid var(--border);border-radius:9px;overflow:hidden;background:var(--panel)}
.itable th{background:var(--panel2);text-align:left;padding:7px 10px;font-size:10.5px;
  text-transform:uppercase;letter-spacing:.08em;color:var(--faint);font-weight:700}
.itable td{padding:8px 10px;border-top:1px solid var(--border);vertical-align:top}
.itable code{background:var(--panel2);padding:1px 5px;border-radius:4px;font-size:12px}
.iwarn{color:var(--warn);font-weight:600}
.inote{color:var(--muted)}
.iok{color:var(--good);font-weight:700}
.aliaswrap::before{content:'=';color:var(--faint);font-weight:700;margin-right:1px}
.alias{padding:2px 9px;border-radius:99px;background:var(--accent-soft);color:var(--accent);
  font-size:11.5px;font-weight:600;border:1px solid var(--border2)}

/* coverage gauge — the signature */
.cov{margin-left:auto;min-width:280px;flex:0 1 340px}
.covtop{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:5px}
.covlabel{font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--faint)}
.covnum{font-family:'Libre Franklin',sans-serif;font-weight:700;font-size:15px;color:var(--accent)}
.covnum small{font-size:11.5px;color:var(--muted);font-weight:600;font-family:'Public Sans',sans-serif}
.covtrack{height:9px;border-radius:99px;background:var(--panel2);border:1px solid var(--border);
  overflow:hidden;position:relative}
.covfill{height:100%;border-radius:99px;width:0;
  background:linear-gradient(90deg,var(--accent),var(--steel));
  transition:width .8s cubic-bezier(.25,.9,.3,1)}
.covcountries{font-size:11.5px;color:var(--accent);cursor:pointer;text-decoration:underline;
  text-underline-offset:2px;margin-top:4px;display:inline-block}
.chipbox{display:none;padding:12px 20px;background:var(--panel2);border-bottom:1px solid var(--border);
  gap:6px;flex-wrap:wrap}
.chipbox.open{display:flex}
.chip{padding:3px 10px;border-radius:99px;background:var(--panel);border:1px solid var(--border2);
  font-size:11.5px;font-weight:600}
.chip b{color:var(--accent)}

.termbody{padding:8px 20px 18px}
.tblsec{margin-top:16px}
.secbar{display:flex;align-items:center;gap:10px;margin-bottom:9px}
.secbar .bar{width:4px;height:18px;border-radius:3px}
.secbar .name{font-weight:700;font-size:13.5px}
.secbar .cnt{font-size:12px;color:var(--muted);font-weight:600}
.secbar .capnote{font-size:11px;color:var(--faint);margin-left:auto}

table.results{width:100%;border-collapse:separate;border-spacing:0;background:var(--panel);
  border:1px solid var(--border);border-radius:9px;overflow:hidden;font-size:12.5px;table-layout:fixed}
table.results th{background:var(--panel2);text-align:left;padding:8px 10px;
  border-bottom:1px solid var(--border2);font-size:10.5px;text-transform:uppercase;
  letter-spacing:.08em;color:var(--faint);font-weight:700}
table.results td{padding:9px 10px;border-bottom:1px solid var(--border);vertical-align:top;
  word-wrap:break-word}
table.results tr:last-child td{border-bottom:0}
table.results tbody tr:hover td{background:color-mix(in srgb, var(--accent-soft) 45%, transparent)}
mark{background:var(--mark);color:var(--mark-ink);font-weight:800;padding:0 2px;border-radius:3px}

.pill{display:inline-block;padding:2px 9px;border-radius:99px;font-size:11px;font-weight:700;
  white-space:nowrap;margin:1px 0}
.pill.good{background:var(--good-soft);color:var(--good)}
.pill.mid{background:var(--mid-soft);color:var(--mid)}
.pill.warn{background:var(--warn-soft);color:var(--warn)}
.pill.bad{background:var(--bad-soft);color:var(--bad)}
.pill.na{background:var(--na-soft);color:var(--na)}

.full{display:none;margin-top:8px;padding:11px;background:var(--panel2);border:1px solid var(--border);
  border-radius:8px;max-height:320px;overflow:auto;line-height:1.65}
.showfull{color:var(--accent);cursor:pointer;font-size:11.5px;font-weight:600;
  text-decoration:underline;text-underline-offset:2px;white-space:nowrap}
.empty{color:var(--faint);font-size:12.5px;padding:10px 2px;font-style:italic}

.hist a{display:block;color:var(--accent);cursor:pointer;padding:4px 0;text-decoration:none;
  font-size:12.5px;border-bottom:1px dashed var(--border);overflow:hidden;
  text-overflow:ellipsis;white-space:nowrap}
.hist a:last-child{border-bottom:0}
.hist a:hover{color:var(--accent-strong)}
.hist .empty{border:0}
#thresholdWrap{display:none;margin-top:8px;padding:10px 12px;background:var(--panel2);
  border-radius:9px;border:1px solid var(--border)}
.welcome{text-align:center;padding:64px 20px;color:var(--faint)}
.welcome .biglyph{font-family:'Libre Franklin',sans-serif;font-size:42px;font-weight:700;
  color:var(--border2);letter-spacing:-.02em;margin-bottom:8px}
@media (max-width:980px){.layout{flex-direction:column}.sidebar{width:100%;position:static;max-height:none}
  .cov{min-width:0;flex-basis:100%;margin-left:0}}
</style>
<style>
/* ---------- v8 additions: rounds, progress, paging ---------- */
:root{--r4:#5b4b8a;--r4-soft:#ece7f7;--r5:#946600;--r5-soft:#fbefcf}
body.dark{--r4:#b9abec;--r4-soft:#2b2546;--r5:#e3bb5f;--r5-soft:#3a2f14}
.rtag{display:inline-block;padding:1px 7px;border-radius:5px;font-size:10.5px;font-weight:800;
  letter-spacing:.03em;white-space:nowrap;line-height:1.6}
.rtag.r4{background:var(--r4-soft);color:var(--r4)}
.rtag.r5{background:var(--r5-soft);color:var(--r5)}
.rpill{padding:2px 10px;border-radius:99px;font-weight:700;font-size:11.5px;border:1px solid var(--border2)}
.rpill.r4{background:var(--r4-soft);color:var(--r4)}
.rpill.r5{background:var(--r5-soft);color:var(--r5)}
.covs{margin-left:auto;display:flex;gap:18px;flex-wrap:wrap;flex:0 1 560px;justify-content:flex-end}
.covs .cov{margin-left:0;min-width:200px;flex:1 1 230px}
.covfill.r4{background:var(--r4)}
.covfill.r5{background:var(--r5)}
.covnum.r4{color:var(--r4)} .covnum.r5{color:var(--r5)}
.roundhint{font-size:11.5px;color:var(--muted);margin-top:7px;line-height:1.5}
.roundnote{padding:10px 16px;margin-bottom:14px;border-radius:10px;background:var(--panel);
  border:1px solid var(--border);box-shadow:var(--shadow);font-size:12.5px;color:var(--muted);
  line-height:1.7}
#roundSeg button{white-space:nowrap;font-size:12px;padding:7px 2px}
body.dark .btn.primary{color:#0d1725}
.roundnote b{color:var(--ink)}
.progwrap{display:none;margin-top:10px}
.progbar{height:8px;border-radius:99px;background:var(--panel2);border:1px solid var(--border);overflow:hidden}
.progfill{height:100%;width:0;background:linear-gradient(90deg,var(--accent),var(--steel));transition:width .35s}
.progtxt{font-size:12px;color:var(--muted);margin-top:5px}
.cardtools{display:flex;align-items:center;gap:10px;padding:12px 0 2px;flex-wrap:wrap}
.cardtools .lbl{font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--faint)}
.seg.mini{display:inline-flex}
.seg.mini button{flex:none;padding:4px 12px;font-size:12px}
.morebar{display:flex;gap:8px;margin-top:8px}
.morebtn{flex:1;padding:8px;border:1px dashed var(--border2);background:var(--panel2);color:var(--accent);
  border-radius:8px;cursor:pointer;font-weight:700;font-family:inherit;font-size:12.5px}
.morebtn:hover{border-style:solid}
.morebtn.all{flex:0 0 auto;padding:8px 16px}
.acts{white-space:nowrap;text-align:right}
.actbtn{border:0;background:transparent;color:var(--faint);cursor:pointer;font-size:14px;
  line-height:1;padding:2px 5px;border-radius:6px;font-family:inherit}
.actbtn:hover{background:var(--accent-soft);color:var(--accent)}
.actbtn.done{color:var(--good)}
tr.dropped .actbtn{opacity:.3}
.chipbox{flex-direction:column;gap:10px}
.chiprow{display:flex;gap:6px;flex-wrap:wrap;align-items:center}
.cmpt{border-collapse:collapse;font-size:12px;margin-top:2px}
.cmpt td,.cmpt th{padding:3px 10px;border-bottom:1px solid var(--border);text-align:left}
.cmpt th{font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:var(--faint)}
.pat{font-size:11px;font-weight:700;padding:1px 8px;border-radius:99px}
.pat.both{background:var(--good-soft);color:var(--good)}
.pat.p4{background:var(--r4-soft);color:var(--r4)}
.pat.p5{background:var(--r5-soft);color:var(--r5)}
.pat.none{background:var(--na-soft);color:var(--na)}
.exportrow{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:10px}
.exportrow select{width:auto;padding:6px 8px;font-size:12.5px}
.linkbtn{background:none;border:0;color:var(--accent);cursor:pointer;font-weight:700;font-size:12.5px;
  text-decoration:underline;text-underline-offset:2px;font-family:inherit;padding:0}
.kbd{font-family:ui-monospace,Consolas,monospace;font-size:11px;border:1px solid var(--border2);
  border-bottom-width:2px;border-radius:4px;padding:0 4px;background:var(--panel2)}
td.jur{font-weight:600}
</style>
</head>
<body>
<header>
  <div class="wordmark">
    <div class="glyph" aria-hidden="true"></div>
    <div>
      <h1>MER Search</h1>
      <div class="eyebrow">FATF Mutual Evaluation Intelligence &nbsp;·&nbsp; __VERSION__</div>
    </div>
  </div>
  <div class="srcchip" id="srcChip"><span class="dot"></span><span id="srcChipTxt"></span></div>
  <div class="spinner" id="spin" role="status" aria-label="working"></div>
  <button class="iconbtn" onclick="toggleDark()" id="darkbtn">Dark mode</button>
</header>

<div class="layout">
  <!-- ============ SIDEBAR ============ -->
  <div class="sidebar">
    <div class="card">
      <h3>Data source</h3>
      <div class="seg" id="modeSeg">
        <button class="on" data-v="MER" onclick="setMode('MER')">MER</button>
        <button data-v="OTHER" onclick="setMode('OTHER')">Other files</button>
      </div>
      <div id="roundWrap" style="margin-top:9px">
        <div class="seg" id="roundSeg">
          <button data-v="4" onclick="setRound('4')">4th round</button>
          <button class="on" data-v="5" onclick="setRound('5')">5th round</button>
          <button data-v="both" onclick="setRound('both')">Both rounds</button>
        </div>
        <div class="roundhint" id="roundHint"></div>
      </div>
      <div id="pathBox" style="margin-top:10px">
        <span class="pathtoggle" onclick="togglePaths()">Change source folders</span>
        <div id="pathFields" style="display:none">
          <label>Round 4 folder</label><input type="text" id="p4" placeholder="leave blank for the default">
          <label>Round 5 folder</label><input type="text" id="p5" placeholder="leave blank for the default">
          <label>Other files folder</label><input type="text" id="po" placeholder="leave blank for the default">
          <div class="hint">Changes take effect when you click Load data.</div>
        </div>
      </div>
      <button class="btn primary" style="width:100%;margin-top:12px" onclick="loadData()">Load data</button>
      <div class="status" id="loadStatus">Not loaded yet.</div>
    </div>

    <div class="card" id="filterCard">
      <h3>Filters</h3>
      <label>Jurisdiction <span style="font-weight:400;color:var(--faint)">(Ctrl+click for multiple)</span></label>
      <select multiple id="fJur"></select>
      <label>Tables</label>
      <label class="checkline"><input type="checkbox" id="tEff" checked><span class="swatch" style="background:var(--eff)"></span><span>Effectiveness</span></label>
      <label class="checkline"><input type="checkbox" id="tRC" checked><span class="swatch" style="background:var(--rc)"></span><span>Risk and Context</span></label>
      <label class="checkline"><input type="checkbox" id="tTC" checked><span class="swatch" style="background:var(--tc)"></span><span>Technical Compliance</span></label>
      <label>Indicator / section</label>
      <select multiple id="fInd"></select>
      <label>Rating <span style="font-weight:400;color:var(--faint)">(effectiveness / compliance)</span></label>
      <select multiple id="fRat"></select>
      <label>Criteria rating</label>
      <select multiple id="fCrit"></select>
      <button class="btn grey" style="width:100%;margin-top:10px" onclick="clearFilters()">Clear filters</button>
    </div>

    <div class="card">
      <h3>Search history</h3>
      <div class="hist" id="hist"><div class="empty">No searches yet.</div></div>
    </div>
  </div>

  <!-- ============ MAIN ============ -->
  <div class="main">
    <div class="card">
      <label style="margin-top:0">Search terms — one per line, no limit. Case, spacing and punctuation are ignored ("ship-based" = "ship based").<br>
      <span style="font-weight:400;color:var(--faint)">Use <b>~</b> to treat different names as the same term: <code>FMLIT~HC fraud and ML intelligence taskforce</code>. Acronyms in brackets expand automatically: <code>JIG (Joint Intelligence Group)</code> also finds JIG and the full name on their own.</span></label>
      <textarea id="terms" placeholder="ship-based casino&#10;cruise&#10;FMLIT~HC fraud and ML intelligence taskforce&#10;JIG (Joint Intelligence Group)"></textarea>
      <div class="termcount" id="termCount">0 terms ready &nbsp;·&nbsp; no limit</div>
      <span class="hint" style="margin-left:8px"><span class="kbd">Ctrl</span> + <span class="kbd">Enter</span> searches</span>

      <label>Search mode</label>
      <div class="seg" id="smSeg" style="max-width:420px">
        <button class="on" data-v="strict" onclick="setSM('strict')">Strict keyword</button>
        <button data-v="exact" onclick="setSM('exact')">Exact phrase</button>
        <button data-v="fuzzy" onclick="setSM('fuzzy')">Fuzzy</button>
      </div>
      <div class="hint" id="smHint">Words must start with your term — "ship" finds ships and shipping, never membership.</div>
      <div id="thresholdWrap">
        <label style="margin-top:0">Fuzzy similarity: <span id="thVal" style="color:var(--accent);font-weight:700">85</span>%</label>
        <input type="range" min="70" max="100" value="85" id="threshold"
               oninput="document.getElementById('thVal').textContent=this.value">
        <div class="hint">Lower = catches more typos but more noise. Filter by jurisdiction to keep fuzzy fast on round 4.</div>
      </div>

      <label class="checkline" style="margin-top:10px"><input type="checkbox" id="capsOnly">
        <span>Match ALL-CAPS terms in capitals only (ASC finds "ASC" but not "ascertain" or "asc")</span></label>

      <label>Report name <span style="font-weight:400;color:var(--faint)">(used in the file name and as the report title; blank uses the first search term)</span></label>
      <input type="text" id="reportName" placeholder="e.g. DPMS in the 4th and 5th rounds">

      <label>Export folder</label>
      <input type="text" id="exportDir" value="__DEFAULT_EXPORT__">
      <div class="hint" id="fileNamePreview"></div>

      <div class="row" style="margin-top:14px">
        <button class="btn primary" id="btnSearch" onclick="doSearch()">Search</button>
        <button class="btn ghost" id="btnCheck" onclick="doInterpret()">Check my terms</button>
        <button class="btn grey" onclick="clearAll()">Clear</button>
      </div>
      <div class="exportrow">
        <button class="btn ghost" id="btnExportX" onclick="doExport('xlsx')" disabled>Export Excel</button>
        <button class="btn ghost" id="btnExportW" onclick="doExport('docx')" disabled>Export Word report</button>
        <button class="btn ghost" id="btnExportB" onclick="doExport('both')" disabled>Both</button>
        <select id="quoteStyle" title="How quotes appear in the Excel and Word exports">
          <option value="paragraph">Quotes in exports: full paragraph</option>
          <option value="sentence">Quotes in exports: sentence only (much smaller files)</option>
        </select>
      </div>
      <div class="progwrap" id="progWrap"><div class="progbar"><div class="progfill" id="progFill"></div></div>
        <div class="progtxt" id="progTxt"></div></div>
      <div class="status" id="searchStatus"></div>
      <div id="interp"></div>
    </div>

    <div id="results" style="margin-top:16px">
      <div class="welcome">
        <div class="biglyph">MER</div>
        Load a data source on the left, enter your terms, and search.<br>
        Every term gets its own results card with jurisdiction coverage for each round.
      </div>
    </div>
  </div>
</div>

<script>
const PAGE=__PAGE_SIZE__;
let dataMode="MER", round="5", searchMode="strict", loaded=false, loadedKey='';
let lastPayload=null, lastSid=null, lastRounds=[], lastResults=[], lastMode='MER';
const excluded=new Set();
const rowCache={};
let cards={};
const RLAB={'4':'4th round','5':'5th round','both':'4th and 5th rounds'};
const RCLS={'4th round':'r4','5th round':'r5'};
const RSHORT={'4th round':'4th','5th round':'5th'};

function $(id){return document.getElementById(id);}
function esc(s){return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function newSid(){return Date.now().toString(36)+Math.random().toString(36).slice(2,10);}
async function post(url,body){
  const r=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
  return r.json();
}
function fileNamePreview(){
  const n=($('reportName').value||'').trim();
  const t=($('terms').value.split('\n').map(x=>x.trim()).filter(Boolean)[0]||'search');
  const d=new Date(); const ds=d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
  $('fileNamePreview').innerHTML='Will save as:  '+esc(ds+'_'+(n||t)+'_FATF Search tool')+'.xlsx / .docx'
    +' &nbsp; <button class="linkbtn" onclick="openFolder()">Open export folder</button>';
}
function togglePaths(){const f=$('pathFields');f.style.display=(f.style.display==='none'?'block':'none');}

function segSet(id,v){document.querySelectorAll('#'+id+' button').forEach(b=>b.classList.toggle('on',b.dataset.v===v));}
function setMode(v){if(v===dataMode)return;dataMode=v;segSet('modeSeg',v);
  $('roundWrap').style.display=(v==='MER')?'':'none';
  $('filterCard').style.display=(v==='MER')?'':'none';
  loaded=false;setChip(null);loadData();}
function setRound(v){if(v===round&&loaded)return;round=v;segSet('roundSeg',v);roundHint();loaded=false;setChip(null);loadData();}
function roundHint(){
  const h={'4':'4th round reports (2013 Methodology).','5':'5th round reports (2022 Methodology) published so far.',
    'both':'Both rounds searched together. Every hit is labelled with its round and every statistic is shown per round &mdash; the rounds are never added together.'};
  $('roundHint').innerHTML=h[round]||'';}
function setSM(v){searchMode=v;segSet('smSeg',v);
  $('thresholdWrap').style.display=(v==='fuzzy')?'':'none';
  const hints={strict:'Words must start with your term — "ship" finds ships and shipping, never membership.',
    exact:'Whole words only — "ship" will not match ships or shipping.',
    fuzzy:'Catches typos and close spellings ("cassino" ≈ "casino"). Slower on large tables.'};
  $('smHint').textContent=hints[v];}
function spin(on){$('spin').style.display=on?'inline-block':'none';}
function toggleDark(){document.body.classList.toggle('dark');
  const d=document.body.classList.contains('dark');
  $('darkbtn').textContent=d?'Light mode':'Dark mode';
  try{localStorage.setItem('fatf_dark',d?'1':'0');}catch(e){}}
try{if(localStorage.getItem('fatf_dark')==='1')toggleDark();}catch(e){}
function setChip(txt){const c=$('srcChip');
  if(!txt){c.style.display='none';return}
  $('srcChipTxt').textContent=txt;c.style.display='flex';}

function fillSelect(id,vals){const s=$(id);const keep=new Set(Array.from(s.selectedOptions).map(o=>o.value));s.innerHTML='';
  vals.forEach(v=>{const o=document.createElement('option');
    if(v&&typeof v==='object'){o.value=v.code;o.textContent=v.label;}
    else{o.value=v;o.textContent=v;}
    if(keep.has(o.value))o.selected=true;
    s.appendChild(o);});}
function selVals(id){return Array.from($(id).selectedOptions).map(o=>o.value);}
function clearFilters(){['fJur','fInd','fRat','fCrit'].forEach(id=>$(id).selectedIndex=-1);
  ['tEff','tRC','tTC'].forEach(id=>$(id).checked=true);}
function paths(){return {round_4:$('p4').value,round_5:$('p5').value,other:$('po').value};}

async function loadData(){
  spin(true);
  const st=$('loadStatus');
  st.classList.remove('err');
  st.textContent='Loading… the very first load of a round takes a while; after that it is quick, even after a restart.';
  const key=dataMode+'|'+round+'|'+JSON.stringify(paths());
  try{
    const j=await post('/api/load',{data_mode:dataMode,round:round,paths:paths()});
    if(!j.ok){st.textContent='Could not load: '+j.error;st.classList.add('err');loaded=false;return;}
    loaded=true;loadedKey=key;
    if(dataMode==='MER'){
      st.textContent=j.rounds.map(r=>'Loaded '+r.round+' — '+r.rows.toLocaleString()+' paragraphs, '
        +r.n_jurisdictions+' jurisdictions\n'
        +Object.entries(r.counts).map(([k,v])=>'  • '+k+': '+v.toLocaleString()).join('\n')).join('\n');
      fillSelect('fJur',j.jurisdictions);fillSelect('fInd',j.indicators);
      fillSelect('fRat',j.ratings);fillSelect('fCrit',j.crit_ratings);
      if(j.rounds.length===2)
        setChip('4th + 5th round · '+j.rounds[0].n_jurisdictions+' + '+j.rounds[1].n_jurisdictions+' jurisdictions');
      else setChip(j.rounds[0].round+' · '+j.n_jurisdictions+' jurisdictions · '+j.total.toLocaleString()+' paragraphs');
    }else{
      st.textContent='Loaded other files:\n'+j.files.map(f=>'  • '+f.file+' — '+
        (f.error?('NOT READABLE: '+f.error):(f.rows.toLocaleString()+' searchable pieces'))).join('\n');
      setChip('Other files · '+j.files.length+' file'+(j.files.length===1?'':'s'));
    }
  }catch(e){st.textContent='Could not load: '+e;st.classList.add('err');loaded=false;}
  finally{spin(false);}
}

function payload(){
  const terms=$('terms').value.split('\n').map(t=>t.trim()).filter(Boolean);
  const tables=[];if($('tEff').checked)tables.push('Effectiveness');
  if($('tRC').checked)tables.push('Risk and Context');
  if($('tTC').checked)tables.push('Technical Compliance');
  return{data_mode:dataMode,round:round,terms:terms,search_mode:searchMode,
    caps_only:$('capsOnly').checked,
    report_name:$('reportName').value,
    excluded:[...excluded],
    paths:paths(),
    threshold:$('threshold').value,
    export_dir:$('exportDir').value,
    quote_style:$('quoteStyle').value,
    filters:{jurisdictions:selVals('fJur'),tables:tables,indicators:selVals('fInd'),
             ratings:selVals('fRat'),crit_ratings:selVals('fCrit')}};
}

function hl(text,matched,variants){
  let vs=(variants&&variants.length)?variants.slice():(matched?[matched]:[]);
  vs=vs.filter(Boolean).sort((a,b)=>b.length-a.length);
  const e=esc(text);
  if(!vs.length)return e;
  const alt=vs.map(v=>esc(v).replace(/[.*+?^${}()|[\]\\]/g,'\\$&')).join('|');
  return e.replace(new RegExp(alt,'gi'),x=>'<mark>'+x+'</mark>');
}
function pillClass(rating){
  const code=(rating||'').split(' ')[0].toUpperCase();
  if(['C','HE','M'].includes(code))return 'good';
  if(['LC','SE','MM'].includes(code))return 'mid';
  if(['PC','ME','PM'].includes(code))return 'warn';
  if(['NC','LE','NM'].includes(code))return 'bad';
  return 'na';
}
function pill(rating){if(!rating)return'';return '<span class="pill '+pillClass(rating)+'">'+esc(rating)+'</span>';}
function rtag(lab){return '<span class="rtag '+(RCLS[lab]||'')+'">'+esc(RSHORT[lab]||lab)+'</span>';}

async function doInterpret(){
  const p=payload();
  if(p.terms.length===0){alert('Enter at least one search term.');return;}
  spin(true);
  try{
    const j=await post('/api/interpret',p);
    const box=$('interp');
    if(!j.ok){box.style.display='block';box.innerHTML='<div class="status err">'+esc(j.error)+'</div>';return;}
    let h='<table class="itable"><thead><tr><th style="width:20%">You typed</th>'
      +'<th style="width:22%">Searched as</th><th>How it will be matched</th></tr></thead><tbody>';
    j.terms.forEach(t=>{
      const al=t.aliases.map(a=>'<code>'+esc(a)+'</code>').join(' = ');
      h+='<tr><td><b>'+esc(t.input)+'</b>'+(t.scope?'<br><span class="scopetag">only '+esc(t.scope)+'</span>':'')+'</td>'
        +'<td>'+al+(t.label?'<div class="inote" style="margin-top:4px">sheet: <b>'+esc(t.label)+'</b></div>':'')+'</td><td>'
        +t.notes.map(n=>'<div class="inote">• '+esc(n)+'</div>').join('')
        +t.how.map(n=>'<div class="inote">• '+esc(n)+'</div>').join('')
        +(t.warn.length?t.warn.map(n=>'<div class="iwarn">⚠ '+esc(n)+'</div>').join('')
                        :'<div class="iok">✓ no ambiguity detected</div>')
        +'</td></tr>';
    });
    h+='</tbody></table>';
    box.style.display='block';box.innerHTML=h;
  }catch(e){$('searchStatus').textContent='Check failed: '+e;}
  finally{spin(false);}
}

/* ---------------- progress ---------------- */
let progTimer=null;
function startProgress(sid,label){
  $('progWrap').style.display='block';$('progFill').style.width='0%';$('progTxt').textContent=label;
  progTimer=setInterval(async()=>{
    try{const r=await fetch('/api/progress?sid='+encodeURIComponent(sid));const j=await r.json();
      if(j.total){const pct=Math.round(100*j.done/j.total);$('progFill').style.width=pct+'%';
        $('progTxt').textContent=(j.finished?'Finishing…':'Searching term '+(j.done+1)+' of '+j.total+(j.term?'  ·  “'+j.term+'”':''));}
    }catch(e){}
  },500);
}
function stopProgress(){if(progTimer)clearInterval(progTimer);progTimer=null;$('progWrap').style.display='none';}

async function doSearch(){
  if(!loaded||loadedKey!==dataMode+'|'+round+'|'+JSON.stringify(paths())){await loadData();if(!loaded)return;}
  const p=payload();
  if(p.terms.length===0){alert('Enter at least one search term.');return;}
  p.sid=newSid();
  spin(true);$('btnSearch').disabled=true;
  const ss=$('searchStatus');ss.classList.remove('err');ss.textContent='';
  $('interp').style.display='none';
  startProgress(p.sid,'Starting…'+(searchMode==='fuzzy'?' Fuzzy mode is slow on the 4th round without filters.':''));
  try{
    const j=await post('/api/search',p);
    if(!j.ok){ss.textContent='Search failed: '+j.error;ss.classList.add('err');return;}
    lastPayload=p;lastSid=j.sid;lastRounds=j.rounds;lastResults=j.results;lastMode=j.data_mode;
    ['btnExportX','btnExportW','btnExportB'].forEach(id=>$(id).disabled=false);
    saveHist(p);
    renderResults(j);
    const tot=j.results.reduce((a,b)=>a+b.total,0);
    let msg='Found '+tot.toLocaleString()+' hit'+(tot===1?'':'s')+' across '+j.results.length+' term'+(j.results.length===1?'':'s')+'.';
    if(j.rounds.length===2){
      const r4=j.results.reduce((a,b)=>a+(b.by_round['4th round']||0),0),r5=j.results.reduce((a,b)=>a+(b.by_round['5th round']||0),0);
      msg+='  4th round: '+r4.toLocaleString()+'  ·  5th round: '+r5.toLocaleString()+'.';
    }
    ss.textContent=msg;
  }catch(e){ss.textContent='Search failed: '+e;ss.classList.add('err');}
  finally{stopProgress();spin(false);$('btnSearch').disabled=false;}
}

const SECS={'Effectiveness':'var(--eff)','Risk and Context':'var(--rc)','Technical Compliance':'var(--tc)'};

function covHtml(block,bi){
  const labs=block.rounds||[];if(!labs.length)return '';
  const two=labs.length===2;
  let h='<div class="covs" onclick="event.stopPropagation()">';
  labs.forEach(lab=>{
    const cv=block.coverage[lab];const cls=two?(RCLS[lab]||''):'';
    const num=cv.total?cv.pct+'% <small>'+cv.mentioned+' of '+cv.total+'</small>'
                      :'<small>no jurisdictions in scope</small>';
    h+='<div class="cov"><div class="covtop"><span class="covlabel">'+(two?esc(lab)+' coverage':'Jurisdiction coverage')+'</span>'
      +'<span class="covnum '+cls+'">'+num+'</span></div>'
      +'<div class="covtrack"><div class="covfill '+cls+'" data-w="'+cv.pct+'"></div></div></div>';
  });
  const any=labs.some(l=>block.coverage[l].mentioned>0);
  if(any)h+='<div style="flex-basis:100%;text-align:right"><span class="covcountries" onclick="toggleChips('+bi+')">Show jurisdictions</span></div>';
  return h+'</div>';
}
function chipsHtml(block){
  const labs=block.rounds||[];const two=labs.length===2;let h='';
  labs.forEach(lab=>{
    const cv=block.coverage[lab];if(!cv.jurisdictions.length)return;
    h+='<div class="chiprow">'+(two?rtag(lab):'')+cv.jurisdictions.map(x=>'<span class="chip">'+esc(x.name)+' <b>'+x.hits+'</b></span>').join('')+'</div>';
  });
  if(two&&block.round_compare&&block.round_compare.length){
    const pc={'Both rounds':'both','4th round only':'p4','5th round only':'p5','Neither round':'none'};
    h+='<div><div class="covlabel" style="margin:2px 0 4px">Jurisdictions assessed in both rounds</div><table class="cmpt"><tr><th>Jurisdiction</th><th>4th</th><th>5th</th><th></th></tr>'
      +block.round_compare.map(c=>'<tr><td>'+esc(c.name)+'</td><td>'+c.r4+'</td><td>'+c.r5+'</td><td><span class="pat '+pc[c.pattern]+'">'+esc(c.pattern)+'</span></td></tr>').join('')
      +'</table></div>';
  }
  return h;
}
function secCount(sec,rnd,two){
  if(rnd!=='all'){const n=sec.by_round[rnd]||0;return n+' hit'+(n===1?'':'s')+' in the '+rnd;}
  let s=sec.total+' hit'+(sec.total===1?'':'s');
  if(two)s+='  ·  4th '+(sec.by_round['4th round']||0)+'  ·  5th '+(sec.by_round['5th round']||0);
  return s;
}

function renderResults(j){
  cards={};
  const R=$('results');
  let top='<div class="dropnote" id="dropNote"></div>';
  if(j.rounds.length===2)
    top+='<div class="roundnote"><b>Both rounds.</b> Hits are tagged '+rtag('4th round')+' or '+rtag('5th round')
      +'. Coverage and counts are always shown per round and never added together, because the rounds cover different countries.</div>';
  R.innerHTML=top;
  const openN=j.results.length<=4?j.results.length:2;
  j.results.forEach((b,bi)=>{
    const two=(b.rounds||[]).length===2;
    cards[bi]={round:'all',shown:{},loaded:false};
    const card=document.createElement('div');card.className='termcard';card.id='card'+bi;
    card.style.animationDelay=Math.min(bi*60,600)+'ms';
    const head=document.createElement('div');head.className='termhead'+(bi<openN?'':' closed');
    let hh='<span class="arrow">▼</span><span class="termname">'+esc(b.term)+'</span>'
      +'<span class="hitpill'+(b.total===0?' zero':'')+'">'+b.total.toLocaleString()+' hit'+(b.total===1?'':'s')+'</span>';
    if(two)hh+=b.rounds.map(l=>'<span class="rpill '+RCLS[l]+'">'+RSHORT[l]+' · '+(b.by_round[l]||0).toLocaleString()+'</span>').join('');
    if(b.aliases&&b.aliases.length>1)
      hh+='<span class="aliaswrap" title="All of these are treated as the same term">'+b.aliases.slice(1).map(a=>'<span class="alias">'+esc(a)+'</span>').join('')+'</span>';
    if(b.scope)hh+='<span class="scopetag">only '+esc(b.scope)+'</span>';
    hh+=covHtml(b,bi);
    head.innerHTML=hh;
    card.appendChild(head);
    if(b.warn&&b.warn.length){
      const wb=document.createElement('div');wb.className='warnbar';
      wb.innerHTML=b.warn.map(w=>'<div>⚠ '+esc(w)+'</div>').join('');card.appendChild(wb);}
    if(b.forms&&b.forms.length){
      const fb=document.createElement('div');fb.className='formsbar';
      fb.innerHTML='<span class="lbl">Matched in the text as</span>'+b.forms.map(f=>'<span class="form">'+esc(f.form)+'<i>'+f.n+'</i></span>').join('');
      card.appendChild(fb);}
    if(b.rounds&&b.rounds.length){
      const cb=document.createElement('div');cb.className='chipbox';cb.id='chips'+bi;cb.innerHTML=chipsHtml(b);card.appendChild(cb);}
    const body=document.createElement('div');body.className='termbody';body.id='body'+bi;
    if(bi>=openN)body.style.display='none';
    let bh='';
    if(two)bh+='<div class="cardtools"><span class="lbl">Show</span><div class="seg mini" id="rs'+bi+'">'
      +'<button class="on" data-v="all" onclick="setCardRound('+bi+',\'all\')">Both rounds</button>'
      +'<button data-v="4th round" onclick="setCardRound('+bi+',\'4th round\')">4th round only</button>'
      +'<button data-v="5th round" onclick="setCardRound('+bi+',\'5th round\')">5th round only</button></div></div>';
    b.sections.forEach((s,si)=>{
      const color=SECS[s.name]||'var(--oth)';
      bh+='<div class="tblsec"><div class="secbar"><span class="bar" style="background:'+color+'"></span>'
        +'<span class="name">'+esc(s.name)+'</span><span class="cnt" id="cnt'+bi+'_'+si+'">'+secCount(s,'all',two)+'</span></div>'
        +'<div id="sb'+bi+'_'+si+'"></div></div>';
    });
    body.innerHTML=bh;
    head.onclick=()=>{head.classList.toggle('closed');const closed=head.classList.contains('closed');
      body.style.display=closed?'none':'';if(!closed)ensureLoaded(bi);};
    card.appendChild(body);
    R.appendChild(card);
    if(bi<openN)ensureLoaded(bi);
  });
  updateDropNote();
  requestAnimationFrame(()=>requestAnimationFrame(()=>{
    document.querySelectorAll('.covfill').forEach(el=>el.style.width=el.dataset.w+'%');}));
}
function ensureLoaded(bi){if(cards[bi].loaded)return;cards[bi].loaded=true;
  lastResults[bi].sections.forEach((s,si)=>loadSection(bi,si,false,PAGE));}
function setCardRound(bi,lab){
  cards[bi].round=lab;segSet('rs'+bi,lab);
  const b=lastResults[bi];
  b.sections.forEach((s,si)=>{$('cnt'+bi+'_'+si).textContent=secCount(s,lab,true);loadSection(bi,si,false,PAGE);});
}
function tableShell(two,isMer){
  if(isMer)return '<table class="results"><colgroup>'+(two?'<col style="width:6%">':'')
      +'<col style="width:11%"><col style="width:8%"><col style="width:17%"><col style="width:9%"><col><col style="width:62px"></colgroup>'
      +'<thead><tr>'+(two?'<th>Round</th>':'')+'<th>Jurisdiction</th><th>Measure</th><th>Indicator</th><th>Rating</th><th>Detected term and context</th><th></th></tr></thead><tbody></tbody></table>';
  return '<table class="results"><colgroup><col style="width:24%"><col style="width:35%"><col><col style="width:62px"></colgroup>'
      +'<thead><tr><th>Where in the file</th><th>Sentence with term</th><th>Detected term and context</th><th></th></tr></thead><tbody></tbody></table>';
}
function actions(id){
  const d=excluded.has(id);
  return '<td class="acts"><button class="actbtn" title="Copy the quote with its citation" onclick="copyCite(\''+id+'\',this)">⧉</button>'
    +'<button class="exbtn" title="'+(d?'Put this hit back':'Remove this hit from the export')+'" onclick="toggleDrop(\''+id+'\',this)">'+(d?'↺':'✕')+'</button></td>';
}
function rowHtml(x,two,isMer){
  const cls=excluded.has(x._id)?' class="dropped"':'';
  if(isMer){
    const crit=(x['Criteria rating']&&x['Criteria rating']!=='NA')?'<br>'+pill(x['Criteria rating']):'';
    return '<tr'+cls+' data-id="'+x._id+'">'+(two?'<td>'+rtag(x.Round)+'</td>':'')
      +'<td class="jur">'+esc(x.Jurisdiction)+'</td><td>'+esc(x.MEASURE)+'</td><td>'+esc(x.Indicator)+'</td>'
      +'<td>'+(x.Rating==='NA'?'<span class="pill na">NA</span>':pill(x.Rating))+crit+'</td>'
      +'<td>'+hl(x.context,x.matched,x.variants)
      +' <span class="showfull" onclick="toggleFull(this)">show full text</span>'
      +'<div class="full">'+hl(x.OBS_VALUE,x.matched,x.variants)+'</div></td>'+actions(x._id)+'</tr>';
  }
  return '<tr'+cls+' data-id="'+x._id+'"><td class="jur">'+esc(x.Location)
    +(x.Details?'<div style="font-weight:400;color:var(--muted);margin-top:3px;font-size:11.5px">'+esc(String(x.Details).slice(0,220))+'</div>':'')
    +'</td><td>'+hl(x['Sentence with term'],x.matched,x.variants)+'</td>'
    +'<td>'+hl(x.context,x.matched,x.variants)
    +' <span class="showfull" onclick="toggleFull(this)">show full text</span>'
    +'<div class="full">'+hl(x['Full text'],x.matched,x.variants)+'</div></td>'+actions(x._id)+'</tr>';
}
function toggleFull(el){const d=el.nextElementSibling;const open=d.style.display==='block';
  d.style.display=open?'none':'block';el.textContent=open?'show full text':'hide full text';}

async function loadSection(bi,si,append,limit){
  const st=cards[bi];const b=lastResults[bi];const sec=b.sections[si];
  const two=(b.rounds||[]).length===2;const isMer=lastMode==='MER';
  const box=$('sb'+bi+'_'+si);
  if(sec.error){box.innerHTML='<div class="empty">Not searched: '+esc(sec.error)+'</div>';return;}
  const off=append?(st.shown[si]||0):0;
  const reqRound=st.round;
  if(!append)box.innerHTML='<div class="empty">Loading…</div>';
  let j;
  try{j=await post('/api/rows',{sid:lastSid,term_index:bi,section:sec.key,round:reqRound,offset:off,limit:limit||PAGE});}
  catch(e){box.innerHTML='<div class="empty">Could not load rows: '+esc(e)+'</div>';return;}
  if(st.round!==reqRound)return;           // user switched round meanwhile
  if(!j.ok){box.innerHTML='<div class="empty">'+esc(j.error)+'</div>';return;}
  j.rows.forEach(x=>rowCache[x._id]=x);
  if(!append){
    if(!j.total){st.shown[si]=0;box.innerHTML='<div class="empty">No results'+(reqRound!=='all'?' in the '+esc(reqRound):' in this '+(isMer?'table':'file'))+'.</div>';return;}
    box.innerHTML=tableShell(two&&reqRound==='all',isMer);
  }
  box.querySelector('tbody').insertAdjacentHTML('beforeend',j.rows.map(x=>rowHtml(x,two&&reqRound==='all',isMer)).join(''));
  st.shown[si]=off+j.rows.length;
  const old=box.querySelector('.morebar');if(old)old.remove();
  const left=j.total-st.shown[si];
  if(left>0){
    box.insertAdjacentHTML('beforeend','<div class="morebar"><button class="morebtn" onclick="loadSection('+bi+','+si+',true,'+PAGE+')">Show '
      +Math.min(PAGE,left)+' more  ·  '+st.shown[si]+' of '+j.total+' shown</button>'
      +'<button class="morebtn all" onclick="loadSection('+bi+','+si+',true,500)">Show '+(left>500?'500 more':'all')+'</button></div>');
  }
}

function updateDropNote(){const note=$('dropNote');if(!note)return;
  note.style.display=excluded.size?'block':'none';
  note.textContent=excluded.size+' hit'+(excluded.size===1?'':'s')+' removed. They will be left out of the export and out of its statistics.';}
function toggleDrop(id,el){
  if(excluded.has(id))excluded.delete(id);else excluded.add(id);
  const d=excluded.has(id);
  document.querySelectorAll('tr[data-id="'+id+'"]').forEach(tr=>{tr.classList.toggle('dropped',d);
    const b=tr.querySelector('.exbtn');if(b){b.textContent=d?'↺':'✕';b.title=d?'Put this hit back':'Remove this hit from the export';}});
  updateDropNote();
}
function citeText(x){
  const q=(x.sentence||x['Sentence with term']||x.context||'').replace(/\s+/g,' ').trim();
  return '“'+q+'” ('+x.ref+')';
}
async function copyCite(id,btn){
  const x=rowCache[id];if(!x)return;const t=citeText(x);
  let ok=false;
  try{await navigator.clipboard.writeText(t);ok=true;}catch(e){
    const ta=document.createElement('textarea');ta.value=t;document.body.appendChild(ta);ta.select();
    try{ok=document.execCommand('copy');}catch(e2){}ta.remove();}
  if(ok){btn.textContent='✓';btn.classList.add('done');setTimeout(()=>{btn.textContent='⧉';btn.classList.remove('done');},1300);}
  else alert(t);
}
function toggleChips(i){const cb=$('chips'+i);cb.classList.toggle('open');
  const link=$('card'+i).querySelector('.covcountries');
  if(link)link.textContent=cb.classList.contains('open')?'Hide jurisdictions':'Show jurisdictions';}

async function doExport(fmt){
  if(!lastPayload){alert('Run a search first.');return;}
  lastPayload.export_dir=$('exportDir').value;
  lastPayload.report_name=$('reportName').value;
  lastPayload.excluded=[...excluded];
  lastPayload.format=fmt||'xlsx';
  lastPayload.quote_style=$('quoteStyle').value;
  lastPayload.sid=lastSid;
  const n=lastResults.reduce((a,b)=>a+b.total,0)-excluded.size;
  const sent=lastPayload.quote_style==='sentence';
  const warn=[];
  if(fmt!=='xlsx'&&n>1500)
    warn.push('The Word report will hold about '+n.toLocaleString()+' quotes (roughly '+Math.round(n/(sent?12:3.5)).toLocaleString()+' pages).');
  if(fmt!=='docx'&&n>15000)
    warn.push('The Excel file will be about '+Math.round(n*(sent?0.36:1.45)/1000)+' MB and may take a minute or two to build.');
  if(warn.length&&!confirm(warn.join('\n\n')
      +(sent?'':'\n\nTip: choosing "Quotes: sentence only" makes both files several times smaller.')+'\n\nBuild anyway?'))return;
  spin(true);
  const ss=$('searchStatus');ss.classList.remove('err');
  ss.textContent='Building the '+(fmt==='docx'?'Word report':fmt==='both'?'files':'spreadsheet')+
    (excluded.size?' (leaving out '+excluded.size+' removed hit'+(excluded.size===1?'':'s')+')':'')+'…';
  try{
    const j=await post('/api/export',lastPayload);
    if(j.ok){ss.innerHTML='Saved:<br>'+j.paths.map(esc).join('<br>')+'<br><button class="linkbtn" onclick="openFolder()">Open export folder</button>';}
    else{ss.textContent='Export failed: '+j.error;ss.classList.add('err');}
  }catch(e){ss.textContent='Export failed: '+e;ss.classList.add('err');}
  finally{spin(false);}
}
async function openFolder(){try{const j=await post('/api/open_folder',{folder:$('exportDir').value});
  if(!j.ok)alert('Could not open the folder: '+j.error);}catch(e){alert('Could not open the folder: '+e);}}

function clearAll(){$('terms').value='';
  $('results').innerHTML='<div class="welcome"><div class="biglyph">MER</div>Load a data source on the left, enter your terms, and search.<br>Every term gets its own results card with jurisdiction coverage for each round.</div>';
  $('searchStatus').textContent='';$('interp').style.display='none';$('interp').innerHTML='';
  ['btnExportX','btnExportW','btnExportB'].forEach(id=>$(id).disabled=true);
  lastPayload=null;excluded.clear();clearFilters();countTerms();fileNamePreview();}

function histGet(){try{return JSON.parse(localStorage.getItem('fatf_hist')||'[]');}catch(e){return [];}}
function histSet(h){try{localStorage.setItem('fatf_hist',JSON.stringify(h));}catch(e){}}
function saveHist(p){
  let h=histGet();
  h.unshift({t:p.terms,mode:p.data_mode,round:p.round,sm:p.search_mode,when:new Date().toLocaleString()});
  histSet(h.slice(0,25));renderHist();}
function histLabel(x){
  const src=(x.mode==='OTHER')?'Other files':(x.round==='both'?'4th + 5th':(x.round==='4'?'4th round':(x.round==='5'?'5th round':'')));
  return (src?'['+src+'] ':'')+x.t.join(', ');}
function renderHist(){
  const h=histGet();const el=$('hist');
  if(!h.length){el.innerHTML='<div class="empty">No searches yet.</div>';return;}
  el.innerHTML=h.map((x,i)=>'<a onclick="reuse('+i+')" title="'+esc(x.when)+'">'+esc(histLabel(x))+'</a>').join('');}
async function reuse(i){const x=histGet()[i];if(!x)return;
  $('terms').value=x.t.join('\n');countTerms();fileNamePreview();
  if(x.sm)setSM(x.sm);
  const needMode=x.mode||'MER';
  if(needMode!==dataMode){dataMode=needMode;segSet('modeSeg',needMode);
    $('roundWrap').style.display=(needMode==='MER')?'':'none';$('filterCard').style.display=(needMode==='MER')?'':'none';loaded=false;}
  if(needMode==='MER'&&x.round&&x.round!==round){round=x.round;segSet('roundSeg',round);roundHint();loaded=false;}
  if(!loaded)await loadData();
  window.scrollTo({top:0});}
function countTerms(){
  const n=$('terms').value.split('\n').map(t=>t.trim()).filter(Boolean).length;
  $('termCount').innerHTML=n+' term'+(n===1?'':'s')+' ready &nbsp;·&nbsp; no limit';}

$('terms').addEventListener('input',()=>{countTerms();fileNamePreview();});
$('terms').addEventListener('keydown',e=>{if(e.key==='Enter'&&(e.ctrlKey||e.metaKey)){e.preventDefault();doSearch();}});
$('reportName').addEventListener('input',fileNamePreview);
countTerms();fileNamePreview();roundHint();renderHist();
loadData();          // load the default round straight away
</script>
</body>
</html>
"""

@app.route("/")
def index():
    page = PAGE.replace("__DEFAULT_EXPORT__", html_escape(DEFAULT_EXPORT_DIR, quote=True))
    page = page.replace("__VERSION__", html_escape(VERSION))
    page = page.replace("__PAGE_SIZE__", str(PAGE_SIZE))
    resp = Response(page, mimetype="text/html")
    # stop the browser serving an old cached copy of the interface
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp

# =============================================================================
if __name__ == "__main__":
    os.makedirs(DEFAULT_EXPORT_DIR, exist_ok=True)
    threading.Timer(1.2, lambda: webbrowser.open(f"http://127.0.0.1:{PORT}")).start()
    print(f"\n  MER Search Tool {VERSION} running at http://127.0.0.1:{PORT}")
    print("  New: search the 4th and 5th rounds together | Round column in every export")
    print(f"  Data folder: {BASE_DIR}")
    print("  Keep this window open while using the tool. Press Ctrl+C to stop.\n")
    app.run(host="127.0.0.1", port=PORT, debug=False, threaded=True)
