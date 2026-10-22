"""
FATF Mutual Evaluation Search Tool  (v9)
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
from html import escape as html_escape, unescape as html_unescape
from html.parser import HTMLParser

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
KID_DIR     = os.path.join(BASE_DIR, "source files", "KID")
OTHER_DIR   = os.path.join(BASE_DIR, "source files", "other")
DEFAULT_EXPORT_DIR = os.path.join(BASE_DIR, "exports")

VERSION = "v9  (15 Oct 2026)"

CONTEXT_CHARS = 250          # characters before/after detected term
PAGE_SIZE     = 25           # rows fetched per "show more" click in the browser
PORT          = 5177

ROUND_LABEL = {"4": "4th round", "5": "5th round"}
TABLES = ["Effectiveness", "Risk and Context", "Technical Compliance"]
KID_TABLES = ["Effectiveness", "Technical Compliance", "General"]
CASE_TABLES = ["Effectiveness"]
STRUCTURED = ("MER", "KID", "CASE")   # sources with rounds, jurisdictions and ratings

def tables_for(mode):
    return {"KID": KID_TABLES, "CASE": CASE_TABLES}.get(mode, TABLES)

def rounds_of(sel):
    """'4' -> ['4'], '5' -> ['5'], 'both' -> ['4', '5']"""
    sel = str(sel or "5").strip().lower()
    if sel in ("both", "all", "4+5", "45"):
        return ["4", "5"]
    return ["4"] if sel == "4" else ["5"]

def round_labels(sel):
    return [ROUND_LABEL[r] for r in rounds_of(sel)]

def source_label(sel, mode="MER"):
    rs = rounds_of(sel)
    if mode == "CASE":
        return "MER case studies, " + (
            "4th and 5th round" if len(rs) == 2 else ROUND_LABEL[rs[0]])
    if mode == "KID":
        return "Key Issue Documents, " + (
            "4th and 5th round" if len(rs) == 2 else ROUND_LABEL[rs[0]])
    if len(rs) == 2:
        return "MER 4th and 5th round"
    return "MER " + ROUND_LABEL[rs[0]]

# folders can be overridden from the interface
PATHS = {"round_4": ROUND4_DIR, "round_5": ROUND5_DIR, "kid": KID_DIR, "other": OTHER_DIR}

def set_paths(p):
    """Accept folder overrides typed by the user; blank means the default."""
    defaults = {"round_4": ROUND4_DIR, "round_5": ROUND5_DIR, "kid": KID_DIR, "other": OTHER_DIR}
    for k in ("round_4", "round_5", "kid", "other"):
        v = (p or {}).get(k)
        newv = str(v).strip().strip('"').strip("'") if v and str(v).strip() else defaults[k]
        if newv != PATHS[k]:
            PATHS[k] = newv
            _CACHE.pop({"round_4": "mer4", "round_5": "mer5", "kid": "kid",
                        "other": "other"}[k], None)
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

# The OECD CSVs carry each paragraph as HTML (paragraphs, lists, tables and
# the shaded "Box" case studies). The search runs on the plain text above, so
# counts never change; this keeps a safe copy of the formatting so the browser
# can show the paragraph as it appears in the report.
_KEEP_TAGS = {"p", "br", "b", "strong", "i", "em", "u", "s", "sub", "sup",
              "ul", "ol", "li", "table", "thead", "tbody", "tfoot", "tr", "th",
              "td", "caption", "colgroup", "col", "div", "span", "blockquote",
              "hr", "h1", "h2", "h3", "h4", "h5", "h6", "a", "dl", "dt", "dd"}
_VOID_TAGS = {"br", "hr", "col"}
_DROP_WITH_CONTENT = {"script", "style", "iframe", "object", "embed", "noscript",
                      "template", "svg", "math", "head", "title"}
_KEEP_STYLE = {"text-align", "vertical-align", "font-weight", "font-style",
               "text-decoration", "list-style-type", "padding-left", "margin-left"}
_SAFE_STYLE_VALUE = re.compile(r"^[#%.,\w\s-]{1,40}$")
_BLOCK_CLOSES_P = (({"p"}, {"td", "th", "li", "div", "blockquote", "dd"}),)
# tag -> ((tags it closes, tags that stop the search), ...)
_IMPLIED_CLOSE = {
    "td": (({"td", "th"}, {"table", "tr"}),),
    "th": (({"td", "th"}, {"table", "tr"}),),
    "tr": (({"tr"}, {"table"}),),
    "thead": (({"tr", "thead", "tbody", "tfoot"}, {"table"}),),
    "tbody": (({"tr", "thead", "tbody", "tfoot"}, {"table"}),),
    "tfoot": (({"tr", "thead", "tbody", "tfoot"}, {"table"}),),
    "li": (({"li"}, {"ul", "ol"}),),
    "dt": (({"dt", "dd"}, {"dl"}),),
    "dd": (({"dt", "dd"}, {"dl"}),),
}
for _t in ("p", "table", "ul", "ol", "div", "blockquote", "h1", "h2", "h3", "h4",
           "h5", "h6", "hr", "dl"):
    _IMPLIED_CLOSE[_t] = _IMPLIED_CLOSE.get(_t, ()) + _BLOCK_CLOSES_P
_HEADING_TAGS = {"h1": "h4", "h2": "h4", "h3": "h5", "h4": "h5", "h5": "h6", "h6": "h6"}

class _Sanitiser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out, self.stack, self.skip = [], [], 0

    def _attrs(self, tag, attrs):
        keep, classes, styles = [], [], []
        a = {k.lower(): (v or "") for k, v in attrs}
        for k in ("colspan", "rowspan"):
            if a.get(k, "").strip().isdigit() and tag in ("td", "th"):
                keep.append(f'{k}="{int(a[k]) if int(a[k]) < 100 else 99}"')
        if tag == "a":
            href = a.get("href", "").strip()
            if re.match(r"^https?://", href, re.I):
                keep.append(f'href="{html_escape(href, quote=True)}" target="_blank" rel="noopener noreferrer"')
        if a.get("align", "").lower() in ("left", "right", "center", "justify"):
            styles.append(f"text-align:{a['align'].lower()}")
        if a.get("valign", "").lower() in ("top", "middle", "bottom"):
            styles.append(f"vertical-align:{a['valign'].lower()}")
        if a.get("bgcolor"):
            classes.append("shaded")
        for decl in a.get("style", "").split(";"):
            if ":" not in decl:
                continue
            prop, val = (x.strip().lower() for x in decl.split(":", 1))
            if prop.startswith("background") and val not in ("", "none", "transparent", "#fff", "#ffffff", "white"):
                classes.append("shaded")
            elif re.match(r"^border(-(top|right|bottom|left))?$", prop) \
                    and tag in ("div", "p", "blockquote") \
                    and not re.search(r"\bnone\b|^0", val):
                classes.append("boxed")
            elif prop in _KEEP_STYLE and _SAFE_STYLE_VALUE.match(val):
                styles.append(f"{prop}:{val}")
        if classes:
            keep.append('class="%s"' % " ".join(sorted(set(classes))))
        if styles:
            keep.append('style="%s"' % html_escape(";".join(styles), quote=True))
        return (" " + " ".join(keep)) if keep else ""

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag in _DROP_WITH_CONTENT:
            self.skip += 1
            return
        if self.skip or tag not in _KEEP_TAGS:
            return
        out_tag = _HEADING_TAGS.get(tag, tag)
        for closes, stops in _IMPLIED_CLOSE.get(tag, ()):
            self._close_implied(closes, stops)
        self.out.append(f"<{out_tag}{self._attrs(tag, attrs)}>")
        if tag not in _VOID_TAGS:
            self.stack.append(out_tag)

    def _close_implied(self, closes, stops):
        """'<td>a<td>b' means two cells, not one inside the other."""
        for i in range(len(self.stack) - 1, -1, -1):
            t = self.stack[i]
            if t in stops:
                return
            if t in closes:
                while len(self.stack) > i:
                    self.out.append(f"</{self.stack.pop()}>")
                return

    def handle_startendtag(self, tag, attrs):
        if tag.lower() in _VOID_TAGS:
            self.handle_starttag(tag, attrs)

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in _DROP_WITH_CONTENT:
            self.skip = max(0, self.skip - 1)
            return
        if self.skip or tag not in _KEEP_TAGS or tag in _VOID_TAGS:
            return
        out_tag = _HEADING_TAGS.get(tag, tag)
        if out_tag in self.stack:              # close anything left open inside it
            while self.stack:
                t = self.stack.pop()
                self.out.append(f"</{t}>")
                if t == out_tag:
                    break

    def handle_data(self, data):
        if not self.skip:
            self.out.append(html_escape(ILLEGAL_XLSX.sub("", data), quote=False))

def sanitize_html(text):
    """Safe, display-only copy of a paragraph's HTML. Returns "" when the
       text has no markup, so the browser simply shows the plain text."""
    if pd.isna(text):
        return ""
    text = str(text)
    if "<" not in text or not re.search(r"<\s*/?\s*[a-zA-Z]", text):
        return ""
    try:
        p = _Sanitiser()
        p.feed(text)
        p.close()
        p.out.extend(f"</{t}>" for t in reversed(p.stack))
        html = "".join(p.out).strip()
    except Exception:
        return ""
    return html if re.search(r"<[a-z]", html) else ""

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
LOADER_VERSION = "9.0"

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
    kfra["OBS_HTML"] = kfra["OBS_VALUE"].apply(sanitize_html)
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
    tc["OBS_HTML"] = tc["OBS_VALUE"].apply(sanitize_html)
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
    df = pd.concat([kfra[UNIFORM_COLS + ["OBS_HTML"]], tc[UNIFORM_COLS + ["OBS_HTML"]]],
                   ignore_index=True)
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

# =============================================================================
# KEY ISSUE DOCUMENTS
# =============================================================================
# One CSV from OECD Data Explorer (DF_KID). Each KID has five parts, one row
# each. The search runs on each part; a hit is the KID, and the browser and
# the exports show the whole KID with the parts that matched marked.
KID_PARTS = OrderedDict([("ISSUE", "Key issue"),
                         ("ECG_CONCL", "ECG discussion"),
                         ("ECG_PROP", "ECG proposal to Plenary"),
                         ("PLEN_RECORD", "Plenary discussion"),
                         ("DISCUSSION_OUTCOME", "Plenary decision")])
KID_PART_NAMES = list(KID_PARTS.values())
_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
           "September", "October", "November", "December"]
KID_COLS = ["KID", "KIDNo", "ThemeCode", "Theme", "Part", "PartNo", "Outcome", "Plenary"]

def plenary_label(tp):
    """'2015-02' -> 'February 2015'"""
    m = re.match(r"^(\d{4})-(\d{2})", str(tp or ""))
    if not m or not 1 <= int(m.group(2)) <= 12:
        return str(tp or "")
    return f"{_MONTHS[int(m.group(2)) - 1]} {m.group(1)}"

def kid_short(kid):
    """'KID n°47 (4th round)' -> 'KID n°47'"""
    return re.sub(r"\s*\([^)]*\)\s*$", "", str(kid or "")).strip()

def html_to_lines(text):
    """Plain text that keeps the paragraph breaks (for Word and Excel)."""
    if pd.isna(text):
        return ""
    t = str(text)
    t = re.sub(r"(?i)<\s*br\s*/?>|<\s*/\s*(p|div|li|tr|h\d)\s*>", "\n", t)
    t = re.sub(r"(?i)<\s*li[^>]*>", "\n\u2022 ", t)
    t = re.sub(r"<[^>]+>", " ", t)
    t = html_unescape(t).replace("\xa0", " ")
    t = ILLEGAL_XLSX.sub("", t)
    lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in t.split("\n")]
    return "\n".join(ln for ln in lines if ln)

def kid_html(text):
    """Display HTML for one KID part. Line breaks in the source are real
       breaks (bullet lists are often typed as lines)."""
    if pd.isna(text) or not str(text).strip():
        return ""
    t = str(text).strip()
    if not re.search(r"<\s*/?\s*[a-zA-Z]", t):
        return "".join(f"<p>{html_escape(ln.strip(), quote=False)}</p>"
                       for ln in t.split("\n") if ln.strip())
    return sanitize_html(re.sub(r"\s*\n\s*", "<br>", t))

def _build_kid(path):
    d = pd.read_csv(path, encoding="utf-8-sig", dtype=str, keep_default_na=False)
    need = ["MER_ROUND", "Jurisdiction", "IOS_REC", "IOs / Recommendations", "THEME",
            "Key Issue theme", "KID_ELEMENT", "Plenary decision", "KID_FINAL_EFF_RATING",
            "Final effectiveness rating", "KID_FINAL_COMP_RATING",
            "Final technical compliance rating", "KID_ID", "TIME_PERIOD", "OBS_VALUE"]
    missing = [c for c in need if c not in d.columns]
    if missing:
        raise ValueError("The Key Issue Documents CSV is missing the column(s) "
                         + ", ".join(missing) + ". Download the full DF_KID table from "
                         "OECD Data Explorer, with labels.")
    register_labels(d["KID_FINAL_EFF_RATING"], d["Final effectiveness rating"])
    register_labels(d["KID_FINAL_COMP_RATING"], d["Final technical compliance rating"])
    rows = []
    for _, x in d.iterrows():
        m = x["IOS_REC"].strip()
        table = ("Effectiveness" if m.startswith("IO") else
                 "Technical Compliance" if re.match(r"^R\d", m) else "General")
        eff = combine_rating(x["KID_FINAL_EFF_RATING"], x["Final effectiveness rating"])
        comp = combine_rating(x["KID_FINAL_COMP_RATING"], x["Final technical compliance rating"])
        rating = eff if table == "Effectiveness" else comp if table == "Technical Compliance" \
            else (eff if eff != "NA" else comp)
        ind = x["IOs / Recommendations"].strip()
        part = KID_PARTS.get(x["KID_ELEMENT"], x["KID_ELEMENT"])
        kid = x["KID_ID"].strip()
        no = re.search(r"(\d+)", kid)
        raw = x["OBS_VALUE"]
        rows.append({
            "Round": ROUND_LABEL.get(x["MER_ROUND"].strip(), x["MER round"] if "MER round" in d.columns else x["MER_ROUND"]),
            "Jurisdiction": x["Jurisdiction"].strip(),
            "Table": table,
            "MEASURE": "" if m in ("", "_Z") else m,
            "Indicator": "No IO or Recommendation" if ind in ("", "Not applicable") else ind,
            "Rating": rating, "Criteria rating": "NA",
            "OBS_VALUE": strip_html(raw), "OBS_HTML": kid_html(raw), "Lines": html_to_lines(raw),
            "KID": kid, "KIDNo": int(no.group(1)) if no else 0,
            "ThemeCode": x["THEME"].strip(), "Theme": x["Key Issue theme"].strip(),
            "Part": part, "PartNo": KID_PART_NAMES.index(part) if part in KID_PART_NAMES else 9,
            "Outcome": x["Plenary decision"].strip() or "Not recorded",
            "Plenary": x["TIME_PERIOD"].strip(),
        })
    df = pd.DataFrame(rows)
    df["_norm"] = [norm_text(t) for t in df["OBS_VALUE"].tolist()]
    index = defaultdict(list)
    for rec in df.drop(columns=["_norm"]).to_dict("records"):
        index[rec["KID"]].append(rec)
    for v in index.values():
        v.sort(key=lambda r: r["PartNo"])
    return df, dict(index)

def kid_file():
    return _find_csv(PATHS["kid"], ["kid"], "Key Issue Documents")

def load_kid_all():
    path = kid_file()
    sig = _file_sig(path)
    ent = _CACHE.get("kid")
    if ent is None or ent[0] != sig:
        df, index = _build_kid(path)
        ent = _CACHE["kid"] = (sig, df, index)
    return ent[1], ent[2]

def load_kid(sel):
    df, _ = load_kid_all()
    return df[df["Round"].isin(round_labels(sel))]

# =============================================================================
# CASE STUDIES (boxes in the effectiveness chapters of the MERs)
# =============================================================================
# MERs set case studies apart in boxes ("Box 3.2. Case example: ..."). In the
# HTML these are single-cell tables, or blocks with shading or a border. Each
# box becomes one record; boxes whose title mentions a case or an example are
# case studies, the others are kept as "Other box" so nothing is lost.
from xml.etree import ElementTree as _ET

CASE_COLS = ["BoxTitle", "BoxKind", "CoreIssue", "BoxNo", "Lines"]
_CASE_TITLE = re.compile(r"\bcase\b|\bcases\b|\bexample|\bcase[- ]stud", re.I)
_BOX_TITLE = re.compile(r"^\s*Box\s+[\dA-Z]+(\.\d+)*", re.I)

def _box_candidates(root):
    """Outermost elements that look like a box."""
    found = []
    def is_box(el):
        tag = el.tag
        cls = el.get("class", "")
        if tag == "table":
            rows = [ch for ch in el if ch.tag == "tr"] + [
                tr for sec in el if sec.tag in ("thead", "tbody", "tfoot")
                for tr in sec if tr.tag == "tr"]
            cells = [c for tr in rows for c in tr if c.tag in ("td", "th")]   # not nested tables
            if len(cells) == 1:
                return True
            if "shaded" in cls and len(cells) <= 2:
                return True
            txt = "".join(el.itertext()).strip()
            return bool(_BOX_TITLE.match(txt)) and len(cells) <= 2
        if tag in ("div", "blockquote", "p") and ("boxed" in cls or "shaded" in cls):
            return True
        return False
    def walk(el):
        for ch in el:
            if is_box(ch):
                found.append(ch)
            else:
                walk(ch)
    walk(root)
    return found

def _box_title(text):
    first = next((ln.strip() for ln in text.split("\n") if ln.strip()), "")
    return first[:240]

def extract_boxes(html):
    """[(title, html, lines, plain_text)] for every box in one paragraph."""
    if not html or ("<table" not in html and "class=" not in html):
        return []
    xml = re.sub(r"<(br|hr|col)(\s[^>]*?)?\s*/?>", r"<\1\2/>", html)
    try:
        root = _ET.fromstring("<root>" + xml + "</root>")
    except _ET.ParseError:
        return []
    out = []
    for el in _box_candidates(root):
        frag = _ET.tostring(el, encoding="unicode", method="html")
        lines = html_to_lines(frag)
        if not lines.strip():
            continue
        out.append((_box_title(lines), frag, lines, strip_html(frag)))
    return out

def _core_issue(measure):
    m = re.match(r"IO(\d+)CI(\d+)", str(measure or ""))
    return f"{m.group(1)}.{m.group(2)}" if m else ""

def _build_cases(df):
    eff = df[(df["Table"] == "Effectiveness") & (df["OBS_HTML"] != "")] \
        if "OBS_HTML" in df.columns else df.iloc[0:0]
    rows = []
    for rec in eff.to_dict("records"):
        for no, (title, frag, lines, text) in enumerate(extract_boxes(rec["OBS_HTML"]), 1):
            r = {c: rec[c] for c in UNIFORM_COLS}
            r.update({"OBS_VALUE": text, "OBS_HTML": frag, "Lines": lines,
                      "BoxTitle": title, "BoxNo": no,
                      "BoxKind": "Case study" if _CASE_TITLE.search(title) else "Other box",
                      "CoreIssue": _core_issue(rec["MEASURE"])})
            rows.append(r)
    out = pd.DataFrame(rows, columns=UNIFORM_COLS + ["OBS_HTML"] + CASE_COLS)
    out["_norm"] = [norm_text(t) for t in out["OBS_VALUE"].tolist()]
    return out

def load_cases_round(rn):
    base = load_round(rn)
    hit = _CACHE.get("case" + rn)
    if hit is not None and hit[0] is base:
        return hit[1]
    cases = _build_cases(base)
    _CACHE["case" + rn] = (base, cases)
    return cases

def load_cases(sel):
    parts = [load_cases_round(r) for r in rounds_of(sel)]
    return parts[0] if len(parts) == 1 else pd.concat(parts, ignore_index=True)

def list_cases(df, filters):
    """Every box matching the filters, for a search with no words."""
    sub = apply_filters(df, filters)
    results = []
    for rec in sub.drop(columns=["_norm"]).to_dict("records"):
        r = {c: rec[c] for c in UNIFORM_COLS}
        r.update({c: rec[c] for c in CASE_COLS})
        r["html"] = rec["OBS_HTML"]
        text = rec["OBS_VALUE"]
        r["context"] = text[:600] + ("..." if len(text) > 600 else "")
        r["sentence"] = rec["BoxTitle"]
        r["matched"] = ""
        r["variants"] = []
        r["ref"] = citation(r)
        r["_id"] = row_id(r)
        results.append(r)
    results.sort(key=lambda r: (r["Jurisdiction"], r["Round"], _measure_key(r["MEASURE"]),
                                r["BoxNo"]))
    return results

def load_struct(p):
    """The table behind a structured source (MER, KID, CASE) for this search."""
    mode = p.get("data_mode")
    if mode == "KID":
        return load_kid(p["round"])
    if mode == "CASE":
        return load_cases(p["round"])
    return load_mer(p["round"])

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
            rawvals = d.iloc[ri].tolist()
            rowvals = [_clean(v) for v in rawvals]
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
                rec = {"Location": where, "Text": val, "Details": details}
                html = sanitize_html(rawvals[ci])
                if html:
                    rec["Html"] = html
                recs.append(rec)
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
    for key, col in (("themes", "ThemeCode"), ("parts", "Part"), ("outcomes", "Outcome"),
                     ("plenaries", "Plenary"), ("boxkinds", "BoxKind")):
        if f.get(key) and col in sub.columns:
            sub = sub[sub[col].isin(f[key])]
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

_TABLE_ORDER = {t: i for i, t in enumerate(TABLES + ["General"])}

def citation(r):
    """'Austria, 5th round MER, IO.4 (rated ME)'"""
    if r.get("KID"):
        ref = f'{r["Jurisdiction"]}, {kid_short(r["KID"])} ({r["Round"]})'
        if r.get("MEASURE"):
            ref += f', {pretty_measure(r["MEASURE"])}'
        ref += f', {plenary_label(r.get("Plenary"))} Plenary'
        if r.get("Outcome"):
            ref += f': {r["Outcome"].lower()}'
        if r.get("Rating") and r["Rating"] != "NA":
            ref += f', final rating {r["Rating"]}'
        return ref
    ref = f'{r["Jurisdiction"]}, {r["Round"]} MER, {pretty_measure(r["MEASURE"])}'
    if r.get("CoreIssue"):
        ref += f', core issue {r["CoreIssue"]}'
    if r.get("Rating") and r["Rating"] != "NA":
        ref += f' (rated {r["Rating"]})'
    if r.get("Criteria rating") and r["Criteria rating"] != "NA":
        ref += f', criterion {r["Criteria rating"]}'
    if r.get("BoxTitle"):
        ref += f', {r["BoxTitle"][:120]}'
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
    htmls = sub["OBS_HTML"].tolist() if "OBS_HTML" in sub.columns else [""] * len(texts)
    extra = {c: sub[c].tolist() for c in CASE_COLS if c in sub.columns}
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
        for c in extra:
            r[c] = extra[c][i]
        r["html"] = htmls[i]
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

def search_kid(df, tp, mode, threshold, filters):
    """Search each part of each KID; one result per KID that matches, carrying
       all five parts (the filter on parts limits only where to search)."""
    sub = apply_filters(df, filters)
    if sub.empty or (tp["pattern"] is None and not any(tp["tokens"])):
        return []
    _, index = load_kid_all()
    texts, kids, pnos = sub["OBS_VALUE"].tolist(), sub["KID"].tolist(), sub["PartNo"].tolist()
    if mode in ("strict", "exact"):
        norms = sub["_norm"].tolist()
        groups = [g for g in tp["ptokens"] if g]
        idx = [i for i, n in enumerate(norms) if any(all(t in n for t in g) for g in groups)]
    else:
        idx = range(len(texts))
    found = defaultdict(dict)
    for i in idx:
        if not texts[i]:
            continue
        hit = find_match(texts[i], tp, mode, threshold)
        if hit is not None:
            found[kids[i]][pnos[i]] = (hit, texts[i])
    results = []
    for kid, parts in found.items():
        recs = index[kid]
        base = recs[0]
        r = {c: base[c] for c in UNIFORM_COLS if c != "OBS_VALUE"}
        r.update({c: base[c] for c in KID_COLS if c not in ("Part", "PartNo")})
        order = sorted(parts)
        variants = []
        for pn in order:
            for v in parts[pn][0][2]:
                if v not in variants:
                    variants.append(v)
        (s0, e0, _), t0 = parts[order[0]]
        names = [KID_PART_NAMES[pn] for pn in order]
        r["matched_parts"] = names
        r["context"] = f"[{names[0]}] " + make_context(t0, s0, e0)
        r["sentence"] = "  \u2026  ".join(
            f"[{KID_PART_NAMES[pn]}] " + sentence_of(parts[pn][1], *parts[pn][0][:2])
            for pn in order)
        r["OBS_VALUE"] = "\n\n".join(f"{KID_PART_NAMES[pn]}: {parts[pn][1]}" for pn in order)
        r["parts"] = [{"part": x["Part"], "html": x["OBS_HTML"], "lines": x["Lines"],
                       "hit": x["PartNo"] in parts} for x in recs]
        r["html"] = ""
        r["matched"] = variants[0]
        r["variants"] = variants
        r["PlenaryLabel"] = plenary_label(r["Plenary"])
        r["ref"] = citation(r)
        r["_id"] = row_id(r)
        results.append(r)
    results.sort(key=lambda r: (_TABLE_ORDER.get(r["Table"], 9), r["Jurisdiction"],
                                r["Round"], r["Plenary"], r["KIDNo"]))
    return results

def coverage_stats(df, filters, results, tables=TABLES):
    """% of jurisdictions (within the current filter scope) that mention the term.
       Always called for ONE round at a time."""
    scope = apply_filters(df, filters)
    total = int(scope["Jurisdiction"].nunique())
    counts = Counter(r["Jurisdiction"] for r in results)
    mentioned = len(counts)
    pct = round(100.0 * mentioned / total, 1) if total else 0.0
    per_table = {tbl: len({r["Jurisdiction"] for r in results if r["Table"] == tbl})
                 for tbl in tables}
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
                "html": rec.get("Html", ""),
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
    if p["data_mode"] in STRUCTURED:
        df = load_struct(p)
        searcher = search_kid if p["data_mode"] == "KID" else search_mer
        known = sorted(df["Jurisdiction"].unique().tolist())
        if p["data_mode"] == "CASE" and not terms:
            filters = dict(p.get("filters", {}) or {})
            raw.append({"term": "All case studies", "aliases": [], "scope": None,
                        "notes": ["No search words: every box that matches the filters."],
                        "how": [], "warn": [], "input": "", "filters": filters,
                        "results": list_cases(df, filters)})
        for i, group in enumerate(terms):
            _set_progress(sid, i, len(terms), group)
            tp = prepare_term(group, mode, csa, known)
            if tp is None:
                continue
            filters = dict(p.get("filters", {}) or {})
            if tp["scope"]:                      # "SAMLIT (South Africa)"
                filters["jurisdictions"] = [tp["scope"]]
            res = searcher(df, tp, mode, threshold, filters)
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
    if p["data_mode"] in STRUCTURED:
        df = load_struct(p)
        tables = tables_for(p["data_mode"])
        labels = round_labels(p["round"])
        per_round = {lab: df[df["Round"] == lab] for lab in labels}
        for b in raw:
            res = [r for r in b["results"] if r["_id"] not in excluded]
            nb = dict(b)
            nb["results"] = res
            nb["rounds"] = labels
            nb["coverage"] = {lab: coverage_stats(per_round[lab], b["filters"],
                                                  [r for r in res if r["Round"] == lab], tables)
                              for lab in labels}
            if len(labels) == 2:
                nb["round_compare"] = round_compare(
                    apply_filters(per_round["4th round"], b["filters"]),
                    apply_filters(per_round["5th round"], b["filters"]), res)
            nb["forms"] = _form_counts(res)
            nb["source_label"] = source_label(p["round"], p["data_mode"])
            nb["tables"] = tables
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
            "Criteria rating", "OBS_VALUE", "html", "context", "sentence", "matched",
            "variants", "ref", "File", "Location", "Sentence with term", "Details",
            "Full text", "KID", "Theme", "Outcome", "Plenary", "PlenaryLabel", "parts",
            "matched_parts", "BoxTitle", "BoxKind", "CoreIssue")
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
    meas = Counter(pretty_measure(x["MEASURE"]) for x in rows if x["MEASURE"])
    rat = Counter(x["Rating"] for x in rows if x.get("Rating") and x["Rating"] != "NA")
    unit = ("KID" if rows and rows[0].get("KID") else
            "case study" if rows and rows[0].get("BoxTitle") else "mention")
    s = (f"{lab.upper()}:  {len(rows)} {_units(len(rows), unit)} in "
         f"{cov['mentioned']} of {cov['total']} jurisdictions ({cov['pct']}%)   |   "
         + ", ".join(f"{t} {nby.get(t, 0)}" for t in b.get("tables", TABLES)))
    if unit == "KID":
        dec = Counter(x["Outcome"] for x in rows)
        s += "   |   Plenary decisions: " + ", ".join(f"{k} ({v})" for k, v in dec.most_common())
    if meas:
        s += "   |   Most under: " + ", ".join(f"{k} ({v})" for k, v in meas.most_common(5))
    if rat:
        s += "   |   Ratings: " + ", ".join(f"{k} ({v})" for k, v in rat.most_common())
    return s

def _write_summary(ws, p, blocks, sheet_names):
    is_mer = p["data_mode"] in STRUCTURED
    is_kid = p["data_mode"] == "KID"
    is_case = p["data_mode"] == "CASE"
    tbls = tables_for(p["data_mode"])
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
    _note(ws, r, ("Each hit is one Key Issue Document (KID) in which the term appears in "
                  "at least one of its five parts. " if is_kid else "")
              + "Click a search term to jump to its sheet. 'Country matrix' shows which "
                "jurisdictions mention which terms; 'All hits' lists every hit in one "
                "filterable table.", ncols, italic=True)
    r += 2
    hr = r
    if is_mer:
        headers = ["Search term", "Round", "KIDs mentioning it" if is_kid else
                   "Case studies" if is_case else "Total hits"] + tbls + [
                   "Jurisdictions mentioned",
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
                vals = [b["term"], lab, len(rows)] + [nby.get(t, 0) for t in tbls] + [
                        cov["mentioned"],
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
    df = load_struct(p)
    unit = {"KID": "KIDs", "CASE": "case studies"}.get(p["data_mode"], "paragraphs")
    base_filters = dict(p.get("filters", {}) or {})
    heads = _unique(["Round", "Jurisdiction"] + [b["term"] for b in blocks]
                    + ["Terms mentioned"])
    ncols = len(heads)
    _band(ws, 1, f"Country matrix  -  {unit} mentioning each term, per jurisdiction", ncols)
    _note(ws, 2, (f"Each number is how many of that jurisdiction's Key Issue Documents mention "
                  if unit == "KIDs" else
                  "Each number is how many case studies (boxes) of that jurisdiction's report "
                  "contain " if unit == "case studies" else
                  "Each number is how many paragraphs of that jurisdiction's report mention ")
              + "the term (empty = not mentioned). 'Terms mentioned' counts how many of "
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

KID_HEADERS = ["Round", "KID", "Reference (for citing)", "Jurisdiction", "IO / Recommendation",
               "Theme", "Plenary", "Plenary outcome", "Final rating", "Term found in"]
KID_WIDTHS = [11, 12, 40, 18, 30, 28, 15, 15, 10, 26]

def _kid_rows(ws, hr, b, sentence_only, idx):
    """One row per KID: its details, then each part with the term in bold."""
    heads = list(KID_HEADERS) + (["Sentences with the term"] if sentence_only
                                 else list(KID_PART_NAMES))
    widths = KID_WIDTHS + ([110] if sentence_only else [60] * len(KID_PART_NAMES))
    _write_header(ws, hr, heads, widths, height=30)
    r = hr + 1
    for x in b["results"]:
        _round_cell(ws.cell(row=r, column=1), x["Round"])
        vals = [kid_short(x["KID"]), x["ref"], x["Jurisdiction"], x["Indicator"], x["Theme"],
                plenary_label(x["Plenary"]), x["Outcome"], x["Rating"],
                ", ".join(x.get("matched_parts") or [])]
        for j, v in enumerate(vals, start=2):
            c = ws.cell(row=r, column=j)
            _setv(c, v)
            _style(c, "wrap", alignment=WRAP)
        variants = x.get("variants") or x["matched"]
        texts = [x["sentence"]] if sentence_only else [pt["lines"] for pt in x["parts"]]
        for j, v in enumerate(texts, start=len(KID_HEADERS) + 1):
            c = ws.cell(row=r, column=j)
            _setv(c, _rich(v, variants) if v else "")
            _style(c, "wrap", alignment=WRAP)
        r += 1
    _add_table(ws, f"tTerm{idx}", hr, r - 1, len(heads))

CASE_HEADERS = ["Round", "Reference (for citing)", "Jurisdiction", "IO", "Core issue",
                "Indicator", "Rating", "Type", "Box title"]
CASE_WIDTHS = [11, 40, 18, 8, 10, 34, 8, 12, 50]

def _case_rows(ws, hr, b, sentence_only, idx):
    """One row per box: where it comes from, then its full text."""
    heads = CASE_HEADERS + ["Sentence with the term" if sentence_only else "Case study (full text)",
                            "Term in context"]
    _write_header(ws, hr, heads, CASE_WIDTHS + [110, 60], height=30)
    r = hr + 1
    for x in b["results"]:
        _round_cell(ws.cell(row=r, column=1), x["Round"])
        vals = [x["ref"], x["Jurisdiction"], pretty_measure(x["MEASURE"]), x.get("CoreIssue", ""),
                x["Indicator"], x["Rating"], x.get("BoxKind", ""), x.get("BoxTitle", "")]
        for j, v in enumerate(vals, start=2):
            c = ws.cell(row=r, column=j)
            _setv(c, v)
            _style(c, "wrap", alignment=WRAP)
        variants = x.get("variants") or x.get("matched")
        body = x["sentence"] if sentence_only else (x.get("Lines") or x["OBS_VALUE"])
        for j, v in ((len(CASE_HEADERS) + 1, body),
                     (len(CASE_HEADERS) + 2, x["context"] if x.get("matched") else "")):
            c = ws.cell(row=r, column=j)
            _setv(c, _rich(v, variants))
            _style(c, "wrap", alignment=WRAP)
        r += 1
    _add_table(ws, f"tTerm{idx}", hr, r - 1, len(heads))

def _write_term_sheet(ws, p, b, idx):
    is_mer = p["data_mode"] in STRUCTURED
    is_kid = p["data_mode"] == "KID"
    is_case = p["data_mode"] == "CASE"
    sentence_only = p.get("quote_style") == "sentence"
    ncols = (len(KID_HEADERS) + 5 if is_kid else len(CASE_HEADERS) + 2 if is_case
             else len(MER_HEADERS)) if is_mer else 6
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
        if is_kid or is_case:
            (_kid_rows if is_kid else _case_rows)(ws, hr, b, sentence_only, idx)
            ws.freeze_panes = ws.cell(row=hr + 1, column=3)
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
    is_mer = p["data_mode"] in STRUCTURED
    is_kid = p["data_mode"] == "KID"
    is_case = p["data_mode"] == "CASE"
    if is_case:
        heads = ["Search term", "Round", "Reference (for citing)", "Jurisdiction", "IO",
                 "Core issue", "Indicator", "Rating", "Type", "Box title", "Matched as"]
        widths = [26, 11, 40, 18, 8, 10, 34, 8, 12, 50, 18]
    elif is_kid:
        heads = ["Search term", "Round", "KID", "Reference (for citing)", "Jurisdiction",
                 "IO / Recommendation", "Theme", "Plenary", "Plenary outcome",
                 "Final rating", "Term found in", "Matched as", "Sentences with the term"]
        widths = [26, 11, 12, 40, 18, 30, 28, 15, 15, 10, 26, 18, 90]
    elif is_mer:
        heads = ["Search term", "Round", "Table", "Reference (for citing)", "Jurisdiction",
                 "MEASURE", "Indicator", "Rating", "Criteria rating", "Matched as",
                 "Sentence with term"]
        widths = [26, 11, 19, 36, 18, 12, 30, 8, 10, 18, 90]
    else:
        heads = ["Search term", "File", "Location", "Matched as", "Sentence with term",
                 "Details"]
        widths = [26, 30, 45, 18, 90, 50]
    _band(ws, 1, "All hits  -  every hit for every term, one row each", len(heads))
    _note(ws, 2, "One flat table for filtering, sorting and pivot tables. "
                 + ("Each KID's full text is on each term's own sheet." if is_kid else
                    "Each case study's full text is on each term's own sheet." if is_case else
                    "Full paragraphs are on each term's own sheet."), len(heads), italic=True)
    hr = 4
    _write_header(ws, hr, heads, widths, height=30)
    r = hr + 1
    for b in blocks:
        rows = b["results"] if is_mer else [x for fb in b["results"] for x in fb["rows"]]
        for x in rows:
            if is_case:
                vals = [b["term"], None, x["ref"], x["Jurisdiction"], pretty_measure(x["MEASURE"]),
                        x.get("CoreIssue", ""), x["Indicator"], x["Rating"], x.get("BoxKind", ""),
                        x.get("BoxTitle", ""), ", ".join(v for v in (x.get("variants") or []) if v)]
            elif is_kid:
                vals = [b["term"], None, kid_short(x["KID"]), x["ref"], x["Jurisdiction"],
                        x["Indicator"], x["Theme"], plenary_label(x["Plenary"]), x["Outcome"],
                        x["Rating"], ", ".join(x.get("matched_parts") or []),
                        ", ".join(x.get("variants") or [x["matched"]]), x["sentence"]]
            elif is_mer:
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
    is_mer = p["data_mode"] in STRUCTURED
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
    mm = re.match(r'R(\d+)(_|$)', m)
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

KID_DECISION_COLOR = {"Upgrade": "2F7A4F", "Downgrade": "A63A2A"}

def _kid_word(x, lab, sentence_only):
    """One KID in the Word report: a heading line, then its parts."""
    out = []
    head = kid_short(x["KID"])
    if x.get("MEASURE"):
        head += "  ·  " + pretty_measure(x["MEASURE"])
    head += "  ·  " + x["Theme"]
    out.append(_para([_round_tag(lab), _run(head, size="19", color="1E2F4D", bold=True)],
                     after=30, before=150))
    meta = [_run(plenary_label(x["Plenary"]) + " Plenary  ·  ", size="18", color="54677F"),
            _run(x["Outcome"], size="18", bold=True,
                 color=KID_DECISION_COLOR.get(x["Outcome"], "54677F"))]
    if x.get("Rating") and x["Rating"] != "NA":
        meta.append(_run("  ·  final rating " + x["Rating"], size="18", color="54677F"))
    meta.append(_run("  ·  term found in: " + ", ".join(x.get("matched_parts") or []),
                     size="18", color="54677F", italic=True))
    out.append(_para(meta, after=60))
    if sentence_only:
        out.append(_para(_highlight_runs(x["sentence"], x.get("variants")),
                         after=110, shade="F7FAFD"))
        return out
    for pt in x["parts"]:
        if not pt["lines"]:
            continue
        out.append(_para([_run(pt["part"] + ("" if pt["hit"] else ""), bold=True, size="17",
                               color="B4472C" if pt["hit"] else "54677F")], after=20, before=40))
        for line in pt["lines"].split("\n"):
            out.append(_para(_highlight_runs(line, x.get("variants"), base_size="20"),
                             after=40, shade="FBF3EE" if pt["hit"] else "F7F7F5"))
    out.append(_para([_run("Reference: " + x["ref"], italic=True, size="16", color="54677F")],
                     after=120))
    return out

def _units(n, unit):
    return {"case study": "case study" if n == 1 else "case studies"}.get(
        unit, unit + ("" if n == 1 else "s"))

def _case_word(x, lab, sentence_only):
    """One case study in the Word report: where it comes from, then the box."""
    head = pretty_measure(x["MEASURE"])
    if x.get("CoreIssue"):
        head += "  ·  core issue " + x["CoreIssue"]
    if x.get("Rating") and x["Rating"] != "NA":
        head += "  ·  rated " + x["Rating"]
    out = [_para([_round_tag(lab), _run(head, size="18", color="2C5F8A", bold=True)],
                 after=30, before=150)]
    if x.get("BoxTitle"):
        out.append(_para([_run(x["BoxTitle"], bold=True, size="21", color="1E2F4D")], after=50))
    if sentence_only and x.get("matched"):
        out.append(_para(_highlight_runs(x["sentence"], x.get("variants")),
                         after=110, shade="F6EBE1"))
    else:
        lines = (x.get("Lines") or x["OBS_VALUE"]).split("\n")
        if lines and lines[0].strip() == (x.get("BoxTitle") or "").strip():
            lines = lines[1:]
        for line in lines:
            out.append(_para(_highlight_runs(line, x.get("variants"), base_size="20"),
                             after=40, shade="F6EBE1"))
    out.append(_para([_run("Reference: " + x["ref"], italic=True, size="16", color="54677F")],
                     after=120))
    return out

def build_word_report(p, blocks, export_dir):
    os.makedirs(export_dir, exist_ok=True)
    path = _free_path(export_dir, export_filename(p, blocks, "docx"))
    title = (str(p.get("report_name") or "").strip()
             or (blocks[0]["term"] if blocks else "Search results"))
    src_label = blocks[0]["source_label"] if blocks else ""
    today = datetime.now().strftime("%d %B %Y")
    is_mer = p["data_mode"] in STRUCTURED
    is_kid = p["data_mode"] == "KID"
    is_case = p["data_mode"] == "CASE"
    tbls = tables_for(p["data_mode"])
    unit = "KID" if is_kid else "case study" if is_case else "mention"
    labels = blocks[0]["rounds"] if (blocks and is_mer) else []
    two = len(labels) == 2
    sentence_only = (p.get("quote_style") == "sentence")
    B = []

    # ---------- cover ----------
    B.append(_para([_run("FATF SEARCH TOOL", bold=True, color="2C5F8A", size="19")], after=60))
    B.append(_para([_run(title, bold=True, color="1E2F4D", size="44")], after=80))
    B.append(_para([_run(("Key Issue Documents that mention the terms, " + src_label.split(", ")[-1]
                          if is_kid else
                          "Case studies from the effectiveness chapters of the MERs, "
                          + src_label.split(", ")[-1] if is_case else "Mentions in " + src_label.replace("MER ", "")
                         + (" Mutual Evaluation Reports" if is_mer else ""))
                         + ", with context and statistics",
                         italic=True, color="54677F", size="23")], after=200, border=True))
    cover = [[("Prepared", True), today], [("Source", True), src_label]]
    if is_mer:
        df_all = load_struct(p)
        gf = dict(p.get("filters") or {})
        for lab in labels:
            tot = int(apply_filters(df_all[df_all["Round"] == lab], gf)["Jurisdiction"].nunique())
            n = sum(1 for b in blocks for x in b["results"] if x["Round"] == lab)
            cover.append([(lab[0].upper() + lab[1:], True),
                          "%d %s  ·  %d jurisdictions %s" %
                          (n, _units(n, unit), tot,
                           "with KIDs" if is_kid else "assessed")])
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
           ] + (["A Key Issue Document (KID) records an issue the Plenary discussed during the "
                 "adoption of a mutual evaluation: the key issue, the ECG discussion, the ECG "
                 "proposal, the Plenary discussion and the Plenary decision. Each KID counts "
                 "once, however many of its parts mention the term; the parts that do are "
                 "marked."] if is_kid else []) + [
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
        rows = [[("Term", True), ("4th round %ss" % unit, True), ("4th round countries", True),
                 ("5th round %ss" % unit, True), ("5th round countries", True)]]
        for b in blocks:
            row = [b["term"][:40]]
            for lab in labels:
                n = sum(1 for x in b["results"] if x["Round"] == lab)
                row += [str(n), _cov_txt(b["coverage"][lab])]
            rows.append(row)
        B.append(_table(rows, [2800, 1350, 1900, 1350, 1900]))
    elif is_mer:
        lab = labels[0]
        short = {"Effectiveness": "Effect.", "Risk and Context": "Risk & ctx",
                 "Technical Compliance": "Tech. comp.", "General": "General"}
        rows = [[("Term", True), ("KIDs" if is_kid else "Mentions", True), ("Countries", True)]
                + [(short[t], True) for t in tbls]]
        for b in blocks:
            n2 = Counter(x["Table"] for x in b["results"])
            rows.append([b["term"][:40], str(len(b["results"])),
                         _cov_txt(b["coverage"][lab])] + [str(n2.get(t, 0)) for t in tbls])
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
                meas = Counter(pretty_measure(x["MEASURE"]) for x in rr if x["MEASURE"])
                rat = Counter(x["Rating"] for x in rr if x.get("Rating") and x["Rating"] != "NA")
                top = Counter(x["Jurisdiction"] for x in rr)
                out = [str(len(rr)), _cov_txt(cov)] + [str(n2.get(t, 0)) for t in tbls] + [
                        ", ".join("%s (%d)" % kv for kv in meas.most_common(5)) or "-",
                        ", ".join("%s (%d)" % kv for kv in rat.most_common()) or "-",
                        ", ".join("%s (%d)" % kv for kv in top.most_common(6)) or "-"]
                if is_kid:
                    dec = Counter(x["Outcome"] for x in rr)
                    out.append(", ".join("%s (%d)" % kv for kv in dec.most_common()) or "-")
                return out
            cols = [per(lab) for lab in labels]
            tname = {"Effectiveness": "In effectiveness " + ("KIDs" if is_kid else "chapters"),
                     "Risk and Context": "In risk and context",
                     "Technical Compliance": "In technical compliance" + (" KIDs" if is_kid else ""),
                     "General": "In general KIDs (no IO or R)"}
            names = (["KIDs mentioning it" if is_kid else "Case studies" if is_case else "Mentions",
                      "Countries mentioning it"]
                     + [tname[t] for t in tbls]
                     + ["Comes up most under",
                        "Final ratings after the Plenary" if is_kid else "Ratings on those paragraphs",
                        "Most mentions"] + (["Plenary decisions"] if is_kid else []))
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
                    if is_kid or is_case:
                        B.append(_para([_run(country + "  ", bold=True, size="22"),
                                        _run("(%d %s%s)" % (len(items), _units(len(items), unit),
                                             ", " + lab if two else ""),
                                             size="19", color="54677F")],
                                       style="Heading3", before=180, after=70))
                        for x in items:
                            B.extend((_kid_word if is_kid else _case_word)(x, lab, sentence_only))
                        continue
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
# =============================================================================
# FILTER PANEL DATA (groups for "select a whole section")
# =============================================================================
# FATF members with their own mutual evaluation (the European Commission and
# the Gulf Co-operation Council are members but are not evaluated). Russia's
# membership has been suspended since February 2023; it keeps its reports.
FATF_MEMBERS = [
    "Argentina", "Australia", "Austria", "Belgium", "Brazil", "Canada", "China",
    "Denmark", "Finland", "France", "Germany", "Greece", "Hong Kong, China",
    "Iceland", "India", "Indonesia", "Ireland", "Israel", "Italy", "Japan", "Korea",
    "Luxembourg", "Malaysia", "Mexico", "Netherlands", "New Zealand", "Norway",
    "Portugal", "Russian Federation", "Saudi Arabia", "Singapore", "South Africa",
    "Spain", "Sweden", "Switzerland", "Türkiye", "United Kingdom", "United States",
]

def _jkey(name):
    k = re.sub(r"[^a-z]", "", deaccent(str(name)).lower())
    return {"hongkongchina": "hongkongchina", "hongkongsarchina": "hongkongchina",
            "hongkong": "hongkongchina", "turkey": "turkiye", "republicofkorea": "korea",
            "korearepublicof": "korea", "russia": "russianfederation",
            "unitedstatesofamerica": "unitedstates", "thenetherlands": "netherlands",
            "netherlandsthe": "netherlands", "kingdomofthenetherlands": "netherlands",
            "kingdomofsaudiarabia": "saudiarabia", "peoplesrepublicofchina": "china"}.get(k, k)

_FATF_KEYS = {_jkey(m) for m in FATF_MEMBERS}

def jurisdiction_facet(df):
    names = sorted(df["Jurisdiction"].unique().tolist())
    members = [n for n in names if _jkey(n) in _FATF_KEYS]
    others = [n for n in names if _jkey(n) not in _FATF_KEYS]
    tree = [{"id": "fatf", "name": "FATF members", "items": members},
            {"id": "other", "name": "Other jurisdictions (FSRB members)", "items": others}]
    presets = []
    rounds = df.groupby("Jurisdiction")["Round"].unique()
    if df["Round"].nunique() == 2:
        has = {j: set(r) for j, r in rounds.items()}
        both = [j for j in names if len(has[j]) == 2]
        only4 = [j for j in names if has[j] == {"4th round"}]
        only5 = [j for j in names if has[j] == {"5th round"}]
        presets = [p for p in (
            {"name": "Assessed in both rounds", "items": both},
            {"name": "4th round only", "items": only4},
            {"name": "5th round only", "items": only5}) if p["items"]]
    presets.append({"name": "FATF members", "items": members})
    return {"tree": [t for t in tree if t["items"]], "presets": presets}

def indicator_facet(df):
    """Indicators grouped by table, then by IO / recommendation."""
    u = df[df["Indicator"] != ""][["Table", "MEASURE", "Indicator"]].drop_duplicates()
    tree = []
    for tbl in TABLES + ["General"]:
        part = u[u["Table"] == tbl]
        if part.empty:
            continue
        by_group = defaultdict(set)
        for m, ind in zip(part["MEASURE"].tolist(), part["Indicator"].tolist()):
            by_group[pretty_measure(m) or "Other"].add(ind)
        names = sorted(by_group, key=_measure_key)
        groups = [{"id": f"{tbl}|{g}", "name": g,
                   "items": sorted(by_group[g], key=_measure_key)} for g in names]
        # one group, or groups of one indicator each, is just noise: flatten
        if len(groups) == 1 or all(len(g["items"]) == 1 for g in groups):
            tree.append({"id": tbl, "name": tbl,
                         "items": sorted({i for g in groups for i in g["items"]},
                                         key=_measure_key)})
        else:
            tree.append({"id": tbl, "name": tbl, "children": groups})
    return {"tree": tree, "presets": []}

_RATING_ORDER = ["HE", "SE", "ME", "LE", "C", "LC", "PC", "NC", "M", "MM", "PM", "NM", "NA"]

def _rating_sort(c):
    return (_RATING_ORDER.index(c) if c in _RATING_ORDER else 50, c)

def rating_facet(df):
    tree = []
    allr = set(df["Rating"]) - {"", "NA"}
    eff = allr & {"HE", "SE", "ME", "LE"}
    tc = allr - eff
    if eff:
        tree.append({"id": "eff", "name": "Effectiveness ratings",
                     "items": sorted(eff, key=_rating_sort)})
    if tc - eff:
        tree.append({"id": "tc", "name": "Technical compliance ratings",
                     "items": sorted(tc - eff, key=_rating_sort)})
    if (df["Rating"] == "NA").any():
        tree.append({"id": "na", "name": "No rating", "items": ["NA"]})
    return {"tree": tree, "presets": []}

def crit_facet(df):
    vals = sorted((x for x in df["Criteria rating"].unique().tolist() if x), key=_rating_sort)
    return {"tree": [{"id": "all", "name": "", "items": vals}], "presets": []}

def kid_theme_facet(df):
    """Themes grouped by table, then by IO / Recommendation."""
    u = df[["Table", "Indicator", "ThemeCode", "Theme"]].drop_duplicates("ThemeCode")
    labels = dict(zip(u["ThemeCode"], u["Theme"]))
    tree = []
    for tbl in KID_TABLES:
        part = u[u["Table"] == tbl]
        if part.empty:
            continue
        groups = []
        for ind in sorted(part["Indicator"].unique().tolist(), key=_measure_key):
            codes = part[part["Indicator"] == ind]
            codes = sorted(codes["ThemeCode"].tolist(), key=lambda c: labels[c].lower())
            groups.append({"id": f"{tbl}|{ind}", "name": ind, "items": codes})
        tree.append({"id": tbl, "name": tbl, "children": groups})
    return {"tree": tree, "presets": [], "labels": labels}

def kid_plenary_facet(df):
    codes = sorted(df["Plenary"].unique().tolist())
    years = OrderedDict()
    for c in codes:
        years.setdefault(c[:4], []).append(c)
    tree = [{"id": "y" + y, "name": y, "items": v} for y, v in reversed(years.items())]
    return {"tree": tree, "presets": [], "labels": {c: plenary_label(c) for c in codes}}

def kid_facets(df):
    outs = [o for o in ("Upgrade", "Downgrade", "Text amendment")
            if o in set(df["Outcome"])] + sorted(set(df["Outcome"]) - {"Upgrade", "Downgrade",
                                                                       "Text amendment"})
    return {"jur": jurisdiction_facet(df), "ind": indicator_facet(df),
            "theme": kid_theme_facet(df),
            "part": {"tree": [{"id": "all", "name": "", "items": KID_PART_NAMES}], "presets": []},
            "outc": {"tree": [{"id": "all", "name": "", "items": outs}], "presets": []},
            "plen": kid_plenary_facet(df), "rat": rating_facet(df)}

def case_facets(df):
    kinds = [k for k in ("Case study", "Other box") if k in set(df["BoxKind"])]
    return {"jur": jurisdiction_facet(df), "ind": indicator_facet(df),
            "rat": rating_facet(df),
            "kind": {"tree": [{"id": "all", "name": "", "items": kinds}], "presets": [],
                     "default": ["Case study"] if "Case study" in kinds else kinds}}

def _err(e):
    return jsonify({"ok": False, "error": f"{e}", "trace": traceback.format_exc()})

@app.route("/api/load", methods=["POST"])
def api_load():
    try:
        p = request.get_json(force=True)
        set_paths(p.get("paths"))
        if p["data_mode"] == "CASE":
            df = load_cases(p["round"])
            per = []
            for r in rounds_of(p["round"]):
                d = df[df["Round"] == ROUND_LABEL[r]]
                per.append({"round": ROUND_LABEL[r], "rows": int(len(d)),
                            "n_jurisdictions": int(d["Jurisdiction"].nunique()),
                            "counts": {{"Case study": "Case studies", "Other box": "Other boxes"}.get(k, k): int(v) for k, v in
                                       d["BoxKind"].value_counts().items()}})
            return jsonify({
                "ok": True, "rounds": per, "unit": "CASE",
                "total": int(len(df)),
                "n_jurisdictions": int(df["Jurisdiction"].nunique()),
                "facets": case_facets(df),
                "labels": {c: rating_option(c) for c in set(df["Rating"].unique()) if c}})
        if p["data_mode"] == "KID":
            df = load_kid(p["round"])
            per = []
            for lab in round_labels(p["round"]):
                d = df[df["Round"] == lab]
                per.append({"round": lab, "rows": int(d["KID"].nunique()),
                            "n_jurisdictions": int(d["Jurisdiction"].nunique()),
                            "counts": {k: int(v) for k, v in
                                       d.drop_duplicates("KID")["Table"].value_counts().items()}})
            return jsonify({
                "ok": True, "rounds": per, "unit": "KID",
                "total": int(df["KID"].nunique()),
                "n_jurisdictions": int(df["Jurisdiction"].nunique()),
                "file": os.path.basename(kid_file()),
                "facets": kid_facets(df),
                "labels": {c: rating_option(c) for c in set(df["Rating"].unique()) if c}})
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
                "facets": {"jur": jurisdiction_facet(df), "ind": indicator_facet(df),
                           "rat": rating_facet(df), "crit": crit_facet(df)},
                "labels": {c: rating_option(c) for c in
                           set(df["Rating"].unique()) | set(df["Criteria rating"].unique()) if c},
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
    if p["data_mode"] in STRUCTURED:
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
                           for t in tables_for(p["data_mode"])]
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
                        "rounds": round_labels(p["round"]) if p["data_mode"] in STRUCTURED else [],
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
        if ent["data_mode"] in STRUCTURED:
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
        if p.get("data_mode") in STRUCTURED:
            try:
                set_paths(p.get("paths"))
                known = sorted(load_struct(p)["Jurisdiction"].unique().tolist())
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
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>FATF Search Tool</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Instrument+Sans:wght@400;500;600;650;700&display=swap" rel="stylesheet">
<style>
/* =====================================================================
   Tokens: warm paper, terracotta, plain professional type
   ===================================================================== */
:root{
  --canvas:#f4eee6; --surface:#fffdfa; --sunk:#f6f0e8; --doc:#fffdfa;
  --ink:#1f1b17; --muted:#665d55; --faint:#978c81;
  --rule:#e8dfd4; --rule2:#d9cbbb;
  --accent:#b4472c; --accent-strong:#933720; --accent-soft:#f8e7de; --on-accent:#ffffff;
  --mark:#f7d6c4; --mark-ink:#3a160a;
  --box:#f6ebe1; --box-rule:#b4472c;
  --eff:#4a6a86; --rc:#62805a; --tc:#b07a2e; --gen:#8a7f74; --oth:#8a8077;
  --good:#2f7a4f; --good-soft:#e2f0e5;
  --mid:#3f6688;  --mid-soft:#e3ebf2;
  --warn:#94660f; --warn-soft:#f8ecd1;
  --bad:#a63a2a;  --bad-soft:#f8e1db;
  --na:#7f756c;   --na-soft:#efe9e2;
  --r4:#6a4c7c; --r4-soft:#efe6f2;
  --r5:#2f6b66; --r5-soft:#dff0ec;
  --shadow:0 1px 2px rgba(60,35,20,.05),0 10px 30px -18px rgba(90,50,25,.25);
  --sans:'Instrument Sans','Segoe UI',system-ui,-apple-system,sans-serif;
  --mono:ui-monospace,'Cascadia Mono',Consolas,monospace;
  --r-sm:6px; --r-md:10px;
  color-scheme:light;
}
body.dark{
  --canvas:#1b1714; --surface:#25201c; --sunk:#2c2621; --doc:#25201c;
  --ink:#f2eae1; --muted:#bbaea0; --faint:#877a6e;
  --rule:#3a322b; --rule2:#4b4037;
  --accent:#e57d5d; --accent-strong:#f39a7e; --accent-soft:#3e251c; --on-accent:#1e0f09;
  --mark:#74361f; --mark-ink:#ffeadf;
  --box:#31261f; --box-rule:#e57d5d;
  --eff:#8fb0cc; --rc:#9cbc92; --tc:#dcae66; --gen:#b3a596; --oth:#a99c90;
  --good:#8cc79a; --good-soft:#1f3123;
  --mid:#90b4d6;  --mid-soft:#1e2c39;
  --warn:#e0b45c; --warn-soft:#382c14;
  --bad:#ef917e;  --bad-soft:#3f221c;
  --na:#a99c90;   --na-soft:#2f2823;
  --r4:#c9a8de; --r4-soft:#34273c;
  --r5:#7fc8bf; --r5-soft:#1b322f;
  --shadow:0 1px 2px rgba(0,0,0,.3),0 12px 30px -18px rgba(0,0,0,.7);
  color-scheme:dark;
}

/* =====================================================================
   Base
   ===================================================================== */
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;font-family:var(--sans);color:var(--ink);font-size:14px;line-height:1.55;-webkit-font-smoothing:antialiased;
  background:radial-gradient(1100px 600px at 100% -5%,#efdfcf 0,transparent 60%),var(--canvas);background-attachment:fixed;
  font-variant-numeric:tabular-nums}
body.dark{background:radial-gradient(1000px 600px at 100% 0,#2c211a 0,transparent 60%),var(--canvas);background-attachment:fixed}
button,input,select,textarea{font:inherit;color:inherit}
:focus-visible{outline:2px solid var(--accent);outline-offset:2px;border-radius:4px}
code{font-family:var(--mono);font-size:.88em;background:var(--sunk);border:1px solid var(--rule);padding:0 .35em;border-radius:4px}
.sr{position:absolute;left:0;top:0;width:1px;height:1px;overflow:hidden;clip-path:inset(50%);white-space:nowrap}
@media (prefers-reduced-motion: reduce){
  *,*::before,*::after{animation:none!important;transition:none!important}
  html{scroll-behavior:auto}
}
mark{color:var(--mark-ink);background:var(--mark);font-weight:600;border-radius:3px;padding:0 .12em;margin:0 -.04em;
  -webkit-box-decoration-break:clone;box-decoration-break:clone}

/* =====================================================================
   Top bar
   ===================================================================== */
.top{position:sticky;top:0;z-index:30;display:flex;align-items:center;gap:14px;height:62px;padding:0 28px;
  background:color-mix(in srgb,var(--canvas) 82%,transparent);backdrop-filter:blur(12px);-webkit-backdrop-filter:blur(12px);
  border-bottom:1px solid var(--rule)}
.top .grow{flex:1}
.brand{display:flex;align-items:center;gap:11px;min-width:0}
.brandmark{width:30px;height:30px;border-radius:8px;background:var(--accent);color:var(--on-accent);display:grid;place-items:center;
  font-size:15px;flex-shrink:0}
.brandmark::before{content:'\2726'}
.brandname{font-weight:700;font-size:16px;letter-spacing:-.01em;line-height:1.15}
.brandsub{font-size:12px;color:var(--faint);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.srcchip{display:none;align-items:center;gap:8px;padding:5px 12px;border-radius:99px;border:1px solid var(--rule2);
  background:var(--surface);color:var(--muted);font-size:12.5px;font-weight:500;white-space:nowrap}
.srcchip .dot{width:7px;height:7px;border-radius:50%;background:var(--good)}
.spinner{display:none;width:16px;height:16px;border:2px solid var(--rule2);border-top-color:var(--accent);border-radius:50%;
  animation:spin .7s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.iconbtn{border:1px solid var(--rule2);background:var(--surface);color:var(--muted);border-radius:var(--r-sm);padding:6px 12px;
  cursor:pointer;font-size:13px;font-weight:500;white-space:nowrap}
.iconbtn:hover{color:var(--ink);border-color:var(--faint)}

/* =====================================================================
   Layout
   ===================================================================== */
.layout{display:grid;grid-template-columns:350px minmax(0,1fr);gap:28px;padding:28px;max-width:1660px;margin:0 auto;align-items:start}
.rail{position:sticky;top:86px;max-height:calc(100vh - 106px);overflow-y:auto;display:flex;flex-direction:column;gap:14px;
  padding:0 4px 8px 0;scrollbar-width:thin}
.panel{background:var(--surface);border:1px solid var(--rule);border-radius:var(--r-md);padding:16px;box-shadow:var(--shadow)}
.main{min-width:0;display:flex;flex-direction:column;gap:20px}
.ptitle{display:flex;align-items:center;gap:9px;margin:0 0 12px;font-size:14.5px;font-weight:650;letter-spacing:-.005em}
.step{width:22px;height:22px;border-radius:50%;background:var(--accent-soft);color:var(--accent);display:inline-grid;place-items:center;
  font-size:12px;font-weight:700;flex-shrink:0}
.ptitle .sub{margin-left:auto;font-size:12px;font-weight:500;color:var(--faint)}

label{display:block;margin:12px 0 5px;font-weight:600;font-size:13px;color:var(--ink)}
label .opt{font-weight:400;color:var(--faint)}
select,textarea,input[type=text],input[type=search]{width:100%;padding:8px 11px;border:1px solid var(--rule2);border-radius:var(--r-sm);
  background:var(--surface);color:var(--ink);font-size:14px;transition:border-color .15s,box-shadow .15s}
select:focus,textarea:focus,input[type=text]:focus,input[type=search]:focus{outline:none;border-color:var(--accent);
  box-shadow:0 0 0 3px color-mix(in srgb,var(--accent) 18%,transparent)}
input[type=range]{width:100%;accent-color:var(--accent)}
input[type=checkbox],input[type=radio]{accent-color:var(--accent)}

/* source picker: one card per source */
.srcpick{display:flex;flex-direction:column;gap:6px}
.srcopt{display:flex;align-items:flex-start;gap:11px;width:100%;text-align:left;padding:10px 12px;border:1px solid var(--rule2);
  border-radius:var(--r-sm);background:var(--surface);cursor:pointer;transition:border-color .15s,background .15s}
.srcopt:hover{border-color:var(--faint)}
.srcopt .radio{width:16px;height:16px;border-radius:50%;border:1.5px solid var(--rule2);margin-top:2px;flex-shrink:0;display:grid;place-items:center}
.srcopt b{display:block;font-size:14px;font-weight:600}
.srcopt span.d{display:block;font-size:12.5px;color:var(--muted);line-height:1.4;margin-top:1px}
.srcopt.on{border-color:var(--accent);background:color-mix(in srgb,var(--accent-soft) 55%,var(--surface));box-shadow:inset 3px 0 0 var(--accent)}
.srcopt.on .radio{border-color:var(--accent)}
.srcopt.on .radio::after{content:'';width:8px;height:8px;border-radius:50%;background:var(--accent)}

.seg{display:flex;background:var(--sunk);border:1px solid var(--rule);border-radius:var(--r-sm);padding:3px;gap:2px}
.seg button{flex:1;padding:7px 8px;border:0;border-radius:4px;background:transparent;color:var(--muted);cursor:pointer;
  font-weight:500;font-size:13px;white-space:nowrap;transition:background .15s,color .15s}
.seg button:hover{color:var(--ink)}
.seg button.on{background:var(--surface);color:var(--accent);font-weight:600;box-shadow:0 1px 2px rgba(60,35,20,.14)}
.seg.mini{display:inline-flex}
.seg.mini button{flex:none;padding:4px 12px;font-size:12.5px}

.btn{display:inline-flex;align-items:center;justify-content:center;gap:7px;padding:9px 18px;border:1px solid transparent;
  border-radius:var(--r-sm);cursor:pointer;font-weight:600;font-size:14px;transition:background .15s,border-color .15s,color .15s,box-shadow .15s}
.btn.primary{background:var(--accent);color:var(--on-accent);box-shadow:0 1px 2px rgba(120,40,20,.25)}
.btn.primary:hover{background:var(--accent-strong)}
.btn.ghost{background:var(--surface);color:var(--accent);border-color:color-mix(in srgb,var(--accent) 45%,var(--rule2))}
.btn.ghost:hover{background:var(--accent-soft)}
.btn.grey{background:transparent;color:var(--muted);border-color:var(--rule2)}
.btn.grey:hover{color:var(--ink);background:var(--sunk)}
.btn.wide{width:100%}
.btn:disabled{opacity:.45;cursor:not-allowed}
.btn svg{width:15px;height:15px}
.linkbtn{background:none;border:0;color:var(--accent);cursor:pointer;font-weight:600;font-size:inherit;padding:0;
  text-decoration:underline;text-decoration-thickness:1px;text-underline-offset:3px}
.linkbtn:hover{color:var(--accent-strong)}

.status{font-size:13px;color:var(--muted);margin-top:10px;white-space:pre-wrap;line-height:1.6}
.status.err{color:var(--bad)}
.status.ok{color:var(--good)}
.hint{font-size:12.5px;color:var(--faint);margin-top:6px;line-height:1.5}
.kbd{font-family:var(--mono);font-size:11px;border:1px solid var(--rule2);border-bottom-width:2px;border-radius:4px;padding:0 4px;
  background:var(--sunk);color:var(--muted)}
.checkline{display:flex;align-items:flex-start;gap:9px;margin:10px 0 0;font-size:13.5px;font-weight:400;color:var(--ink);cursor:pointer}
.checkline input{width:16px;height:16px;margin:2px 0 0;flex-shrink:0;cursor:pointer}
.pathtoggle{margin-top:12px;font-size:13px}
#pathFields{margin-top:4px}
.roundhint{font-size:12.5px;color:var(--muted);margin-top:7px;line-height:1.5}
.sublabel{font-size:12.5px;font-weight:600;color:var(--muted);margin:14px 0 6px}

/* =====================================================================
   Filters (rail)
   ===================================================================== */
.ftop{display:flex;align-items:center;gap:8px;margin-bottom:6px}
.ftop .ptitle{margin:0}
.expandbtn{margin-left:auto;display:inline-flex;align-items:center;gap:6px;border:1px solid var(--rule2);background:var(--surface);
  color:var(--accent);border-radius:var(--r-sm);padding:5px 10px;cursor:pointer;font-size:12.5px;font-weight:600}
.expandbtn:hover{background:var(--accent-soft)}
.expandbtn svg{width:13px;height:13px}
.fsummary{font-size:12.5px;color:var(--muted);margin:0 0 12px;line-height:1.5}
.tabletoggles{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:6px}
.ttog{position:relative;display:inline-flex;align-items:center;gap:7px;margin:0;padding:5px 11px 5px 9px;border-radius:99px;
  border:1px solid var(--rule2);font-size:12.5px;font-weight:500;color:var(--muted);cursor:pointer;background:var(--surface);user-select:none}
.ttog input{position:absolute;opacity:0;pointer-events:none}
.ttog .sw{width:9px;height:9px;border-radius:2px;background:var(--c);opacity:.3}
.ttog:has(input:checked){color:var(--ink);border-color:color-mix(in srgb,var(--c) 55%,var(--rule2));background:color-mix(in srgb,var(--c) 9%,var(--surface))}
.ttog:has(input:checked) .sw{opacity:1}
.ttog:has(input:focus-visible){outline:2px solid var(--accent);outline-offset:2px}

.facet{border:1px solid var(--rule);border-radius:var(--r-sm);background:var(--surface);margin-top:8px;overflow:hidden}
.facet-head{display:flex;align-items:center;gap:8px;width:100%;padding:9px 10px 9px 8px;background:none;border:0;cursor:pointer;
  text-align:left;font-size:13.5px;font-weight:600}
.facet-head:hover{background:var(--sunk)}
.caret{width:16px;height:16px;flex-shrink:0;display:inline-grid;place-items:center;color:var(--faint);transition:transform .15s}
.caret::before{content:'';width:6px;height:6px;border-right:1.5px solid currentColor;border-bottom:1.5px solid currentColor;transform:translate(-1px,-1px) rotate(45deg)}
.closed>.facet-head .caret,.fgroup:not(.open)>.fghead .fcaret{transform:rotate(-90deg)}
.facet-count{margin-left:auto;font-size:12px;font-weight:500;color:var(--faint);white-space:nowrap}
.facet-count.some{color:var(--accent);font-weight:600}
.facet-count.none{color:var(--bad);font-weight:600}
.facet-body{border-top:1px solid var(--rule);padding:8px}
.closed>.facet-body{display:none}
.facet-tools{display:flex;gap:4px;align-items:center;margin-bottom:6px}
.facet-tools input{padding:6px 8px;font-size:13px}
.mini{border:1px solid var(--rule2);background:var(--surface);color:var(--muted);border-radius:5px;padding:5px 9px;cursor:pointer;
  font-size:12.5px;font-weight:500;white-space:nowrap}
.mini:hover{color:var(--ink);background:var(--sunk)}
.flist{max-height:236px;overflow:auto;margin:0 -4px;padding:0 4px;scrollbar-width:thin}
.fempty{font-size:13px;color:var(--faint);padding:8px 4px}

.fgroup{margin:2px 0}
.fghead{display:flex;align-items:center;gap:4px;padding:3px 2px;border-radius:5px}
.fghead:hover{background:var(--sunk)}
.fcaret{border:0;background:none;padding:0;cursor:pointer;color:var(--faint);width:18px;height:18px;display:grid;place-items:center;transition:transform .15s;flex-shrink:0}
.fcaret::before{content:'';width:6px;height:6px;border-right:1.5px solid currentColor;border-bottom:1.5px solid currentColor;transform:translate(-1px,-1px) rotate(45deg)}
.fgcheck{display:flex;align-items:center;gap:7px;margin:0;font-size:13px;font-weight:600;color:var(--ink);cursor:pointer;min-width:0;flex:1}
.fgcheck input,.fitem input{width:15px;height:15px;margin:0;flex-shrink:0;cursor:pointer}
.fgname{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.fgcount{font-size:11.5px;color:var(--faint);padding-right:4px}
.fgbody{padding-left:22px}
.fitems{display:flex;flex-direction:column}
.fitem{display:flex;align-items:flex-start;gap:7px;margin:0;padding:3px 4px;border-radius:5px;font-size:13px;font-weight:400;color:var(--ink);cursor:pointer;line-height:1.35}
.fitem:hover{background:var(--sunk)}
.fitem input{margin-top:1px}
.fitem .q{background:var(--mark);color:var(--mark-ink);border-radius:2px}
.fitem.off span{color:var(--muted)}

/* =====================================================================
   Filters pop-out
   ===================================================================== */
dialog.fdlg{padding:0;border:1px solid var(--rule2);border-radius:12px;background:var(--surface);color:var(--ink);
  width:min(1180px,94vw);height:min(820px,90vh);max-width:none;max-height:none;box-shadow:0 30px 80px -20px rgba(40,20,10,.45);
  display:none;flex-direction:column;overflow:hidden}
dialog.fdlg[open]{display:flex;animation:pop .18s ease-out}
dialog.fdlg.max{width:100vw;height:100vh;border-radius:0;border:0}
dialog.fdlg::backdrop{background:rgba(40,24,14,.35);backdrop-filter:blur(3px)}
@keyframes pop{from{opacity:0;transform:scale(.985)}to{opacity:1;transform:none}}
.fdlg-head{display:flex;align-items:center;gap:12px;padding:16px 20px;border-bottom:1px solid var(--rule)}
.fdlg-head h2{margin:0;font-size:19px;font-weight:700;letter-spacing:-.01em}
.fdlg-head p{margin:2px 0 0;font-size:13px;color:var(--muted)}
.fdlg-head .sp{flex:1}
.fdlg-body{flex:1;min-height:0;display:grid;grid-template-columns:250px minmax(0,1fr)}
.fdlg-nav{border-right:1px solid var(--rule);background:var(--sunk);padding:12px;overflow:auto;display:flex;flex-direction:column;gap:4px}
.navtab{display:flex;flex-direction:column;align-items:flex-start;gap:1px;width:100%;text-align:left;padding:9px 12px;border:1px solid transparent;
  border-radius:8px;background:none;cursor:pointer}
.navtab:hover{background:var(--surface)}
.navtab.on{background:var(--surface);border-color:var(--rule2);box-shadow:inset 3px 0 0 var(--accent)}
.navtab b{font-size:14px;font-weight:600}
.navtab span{font-size:12.5px;color:var(--muted)}
.navtab span.some{color:var(--accent);font-weight:600}
.navtab span.none{color:var(--bad);font-weight:600}
.navsep{margin:14px 4px 6px;font-size:12.5px;font-weight:600;color:var(--muted)}
.fdlg-nav .tabletoggles{flex-direction:column;align-items:stretch;padding:0 2px}
.fdlg-pane{display:flex;flex-direction:column;min-height:0}
.fdlg-tools{display:flex;flex-wrap:wrap;gap:8px;align-items:center;padding:14px 20px 10px}
.fdlg-tools input[type=search]{flex:1 1 260px;width:auto;padding:9px 12px;font-size:14px}
.fdlg-tools .checkline{margin:0 0 0 4px;white-space:nowrap}
.presets{display:flex;flex-wrap:wrap;gap:6px;align-items:center;padding:0 20px 10px;font-size:12.5px;color:var(--muted)}
.presets:empty{display:none}
.preset{border:1px dashed var(--rule2);background:var(--surface);color:var(--ink);border-radius:99px;padding:4px 11px;cursor:pointer;font-size:12.5px}
.preset:hover{border-style:solid;border-color:var(--accent);color:var(--accent)}
.preset small{color:var(--faint);margin-left:3px}
.fdlg-list{flex:1;min-height:0;overflow:auto;padding:4px 20px 20px;border-top:1px solid var(--rule)}
.fdlg-list>.fgroup{padding:10px 0;margin:0;border-bottom:1px solid var(--rule)}
.fdlg-list>.fgroup:last-child{border-bottom:0}
.fdlg-list>.fgroup>.fghead .fgcheck{font-size:14.5px}
.fdlg-list .fitems{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:0 12px}
.fdlg-list .fgbody{padding-left:24px}
.fdlg-list .fgbody:has(>.fgroup){display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:4px 18px;padding-top:4px}
.fdlg-list .fgbody>.fgroup .fitems{grid-template-columns:1fr}
.fdlg-list>.fitems{padding-top:10px}
.fdlg-foot{display:flex;align-items:center;gap:12px;padding:14px 20px;border-top:1px solid var(--rule);background:var(--sunk)}
.fdlg-foot .note{font-size:13px;color:var(--muted);flex:1}

/* =====================================================================
   Composer
   ===================================================================== */
.composer{background:var(--surface);border:1px solid var(--rule);border-radius:var(--r-md);box-shadow:var(--shadow);padding:22px 24px 20px}
.ctitle{font-size:24px;font-weight:700;letter-spacing:-.015em;line-height:1.2;margin:0 0 4px}
.csub{font-size:14px;color:var(--muted);margin:0 0 18px;max-width:78ch}
.chead{display:flex;align-items:baseline;gap:10px;flex-wrap:wrap;margin-bottom:7px}
.chead label{margin:0}
.termcount{font-size:12.5px;font-weight:600;color:var(--accent)}
.chead .hint{margin:0 0 0 auto}
textarea#terms{height:150px;resize:vertical;font-family:var(--mono);font-size:13.5px;line-height:1.6;background:var(--sunk)}
.syntax{margin-top:8px;font-size:13px;color:var(--muted)}
.syntax summary{cursor:pointer;color:var(--accent);font-weight:600;width:max-content}
.syntax table{border-collapse:collapse;margin-top:8px;width:100%;max-width:780px}
.syntax td{padding:6px 12px 6px 0;border-top:1px solid var(--rule);vertical-align:top}
.syntax td:first-child{white-space:nowrap;width:1%}
.optrow{display:grid;grid-template-columns:minmax(0,420px) minmax(0,1fr);gap:8px 28px;margin-top:12px;align-items:start}
.optrow label{margin-top:0}
#thresholdWrap{display:none;margin-top:10px;padding:10px 12px;background:var(--sunk);border-radius:var(--r-sm);border:1px solid var(--rule)}
.actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:18px;align-items:center}
.actions .btn.primary{padding:10px 28px}
.progwrap{display:none;margin-top:12px}
.progbar{height:6px;border-radius:99px;background:var(--sunk);border:1px solid var(--rule);overflow:hidden}
.progfill{height:100%;width:0;background:var(--accent);transition:width .35s}
.progtxt{font-size:13px;color:var(--muted);margin-top:5px}
.exportbox{margin-top:20px;padding-top:18px;border-top:1px solid var(--rule)}
.exportbox .ptitle{margin-bottom:2px}
.exgrid{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:0 20px}
.exportrow{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:12px}
.exportrow select{width:auto;padding:8px 10px;font-size:13.5px}
.exportnote{font-size:12.5px;color:var(--faint);margin-top:8px}

#interp{margin-top:14px;display:none}
.itable{width:100%;border-collapse:separate;border-spacing:0;font-size:13px;border:1px solid var(--rule);border-radius:var(--r-sm);overflow:hidden}
.itable th{background:var(--sunk);text-align:left;padding:8px 10px;font-size:12.5px;color:var(--muted);font-weight:600}
.itable td{padding:9px 10px;border-top:1px solid var(--rule);vertical-align:top}
.iwarn{color:var(--warn);font-weight:600}
.inote{color:var(--muted)}
.iok{color:var(--good);font-weight:600}

/* =====================================================================
   Results
   ===================================================================== */
.welcome{padding:48px 24px;text-align:center;color:var(--muted);border:1px dashed var(--rule2);border-radius:var(--r-md);background:color-mix(in srgb,var(--surface) 50%,transparent)}
.welcome h2{font-weight:700;font-size:19px;color:var(--ink);margin:0 0 8px;letter-spacing:-.01em}
.welcome p{margin:0 auto;max-width:60ch}
.welcome ol{display:inline-flex;gap:22px;flex-wrap:wrap;justify-content:center;list-style:none;padding:0;margin:18px 0 0;font-size:13.5px;color:var(--ink)}
.welcome li{display:flex;align-items:center;gap:8px}
.roundnote,.dropnote{padding:10px 14px;border-radius:var(--r-sm);font-size:13px;line-height:1.6}
.roundnote{background:var(--surface);border:1px solid var(--rule);color:var(--muted);margin-bottom:14px}
.roundnote b{color:var(--ink)}
.dropnote{display:none;background:var(--warn-soft);color:var(--warn);font-weight:600;margin-bottom:14px}

.termcard{background:var(--surface);border:1px solid var(--rule);border-radius:var(--r-md);box-shadow:var(--shadow);margin-bottom:20px;overflow:hidden}
.termhead{display:flex;align-items:center;gap:10px 14px;padding:16px 20px;cursor:pointer;flex-wrap:wrap;border-top:3px solid var(--accent)}
.termhead:hover{background:color-mix(in srgb,var(--sunk) 60%,transparent)}
.termhead .arrow{width:16px;height:16px;display:grid;place-items:center;color:var(--faint);transition:transform .15s;flex-shrink:0}
.termhead .arrow::before{content:'';width:7px;height:7px;border-right:1.8px solid currentColor;border-bottom:1.8px solid currentColor;transform:translate(-1px,-2px) rotate(45deg)}
.termhead.closed .arrow{transform:rotate(-90deg)}
.termname{font-size:20px;font-weight:700;letter-spacing:-.01em;line-height:1.2}
.hitpill{padding:3px 11px;border-radius:99px;background:var(--accent);color:var(--on-accent);font-weight:600;font-size:12.5px}
.hitpill.zero{background:var(--na-soft);color:var(--na)}
.rpill{padding:2px 9px;border-radius:99px;font-weight:600;font-size:12px}
.rpill.r4{background:var(--r4-soft);color:var(--r4)}
.rpill.r5{background:var(--r5-soft);color:var(--r5)}
.aliaswrap{display:inline-flex;gap:5px;flex-wrap:wrap;align-items:center;font-size:12.5px;color:var(--faint)}
.alias{padding:2px 9px;border-radius:99px;background:var(--accent-soft);color:var(--accent);font-size:12px;font-weight:500}
.scopetag{padding:2px 9px;border-radius:99px;background:var(--warn-soft);color:var(--warn);font-size:12px;font-weight:600}

.covs{margin-left:auto;display:flex;gap:20px;flex-wrap:wrap;flex:0 1 560px;justify-content:flex-end}
.cov{min-width:210px;flex:1 1 230px}
.covtop{display:flex;justify-content:space-between;align-items:baseline;gap:10px;margin-bottom:5px}
.covlabel{font-size:12.5px;font-weight:500;color:var(--muted)}
.covnum{font-weight:700;font-size:15px;color:var(--accent)}
.covnum small{font-size:12px;color:var(--muted);font-weight:500}
.covnum.r4{color:var(--r4)} .covnum.r5{color:var(--r5)}
.covtrack{height:6px;border-radius:99px;background:var(--sunk);overflow:hidden;box-shadow:inset 0 0 0 1px var(--rule)}
.covfill{height:100%;border-radius:99px;width:0;background:var(--accent);transition:width .8s cubic-bezier(.25,.9,.3,1)}
.covfill.r4{background:var(--r4)} .covfill.r5{background:var(--r5)}
.covcountries{font-size:12.5px}

.warnbar{padding:9px 20px;background:var(--warn-soft);color:var(--warn);font-size:13px;font-weight:500;line-height:1.55;border-top:1px solid var(--rule)}
.warnbar div+div{margin-top:3px}
.formsbar{display:flex;gap:6px;flex-wrap:wrap;align-items:center;padding:10px 20px;border-top:1px solid var(--rule);background:var(--sunk)}
.formsbar .lbl{font-size:12.5px;color:var(--muted);margin-right:4px}
.form{padding:1px 8px;border-radius:5px;background:var(--surface);border:1px solid var(--rule2);font-size:12.5px;font-weight:600}
.form i{font-style:normal;font-weight:500;color:var(--faint);margin-left:5px}
.chipbox{display:none;padding:14px 20px;background:var(--sunk);border-top:1px solid var(--rule);flex-direction:column;gap:10px}
.chipbox.open{display:flex}
.chiprow{display:flex;gap:6px;flex-wrap:wrap;align-items:center}
.chip{padding:3px 10px;border-radius:99px;background:var(--surface);border:1px solid var(--rule2);font-size:12.5px}
.chip b{color:var(--accent)}
.cmpt{border-collapse:collapse;font-size:12.5px;margin-top:2px}
.cmpt td,.cmpt th{padding:4px 12px 4px 0;border-bottom:1px solid var(--rule);text-align:left}
.cmpt th{font-size:12px;color:var(--muted);font-weight:600}
.pat{font-size:11.5px;font-weight:600;padding:1px 8px;border-radius:99px}
.pat.both{background:var(--good-soft);color:var(--good)}
.pat.p4{background:var(--r4-soft);color:var(--r4)}
.pat.p5{background:var(--r5-soft);color:var(--r5)}
.pat.none{background:var(--na-soft);color:var(--na)}

.termbody{padding:4px 20px 20px;border-top:1px solid var(--rule)}
.cardtools{display:flex;align-items:center;gap:10px;padding:14px 0 0;flex-wrap:wrap;font-size:13px;color:var(--muted)}
.tblsec{margin-top:18px}
.secbar{display:flex;align-items:center;gap:10px;margin-bottom:8px}
.secbar .bar{width:10px;height:10px;border-radius:3px}
.secbar .name{font-weight:650;font-size:15px}
.secbar .cnt{font-size:13px;color:var(--muted)}
.secrows{overflow-x:auto}

table.results{width:100%;border-collapse:separate;border-spacing:0;border:1px solid var(--rule);border-radius:var(--r-sm);overflow:hidden;
  font-size:13.5px;table-layout:fixed;background:var(--surface)}
table.results th{position:relative;background:var(--sunk);text-align:left;padding:8px 10px;border-bottom:1px solid var(--rule);font-size:12.5px;color:var(--muted);font-weight:600}
table.results td{padding:11px 10px;border-bottom:1px solid var(--rule);vertical-align:top;overflow-wrap:anywhere}
table.results tbody tr:last-child td{border-bottom:0}
table.results tbody tr.hit:hover td{background:color-mix(in srgb,var(--accent-soft) 35%,transparent)}
td.jur{font-weight:600}
td.meas{color:var(--muted)}
td.ctx{line-height:1.6}
.smallmeta{display:block;font-size:12.5px;color:var(--muted);font-weight:400;margin-top:2px}
tr.dropped td{opacity:.35}
tr.dropped td.ctx{text-decoration:line-through}
tr.dropped td.acts{opacity:1}
.rowtools{display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin-top:7px;font-size:12.5px}
.readbtn{display:inline-flex;align-items:center;gap:5px;border:1px solid color-mix(in srgb,var(--accent) 45%,var(--rule2));background:var(--surface);color:var(--accent);
  border-radius:5px;padding:3px 10px;cursor:pointer;font-size:12.5px;font-weight:600}
.readbtn:hover{background:var(--accent-soft)}
.readbtn[aria-expanded=true]{background:var(--accent);color:var(--on-accent);border-color:var(--accent)}
.fmt{font-size:12px;color:var(--muted);border:1px solid var(--rule);border-radius:5px;padding:0 6px;background:var(--sunk)}
.parttag{font-size:12px;color:var(--accent);background:var(--accent-soft);border-radius:5px;padding:0 7px;font-weight:600}
.acts{white-space:nowrap;text-align:right}
.actbtn,.exbtn{border:0;background:transparent;color:var(--faint);cursor:pointer;font-size:14px;line-height:1;padding:4px 6px;border-radius:5px}
.actbtn:hover{background:var(--accent-soft);color:var(--accent)}
.exbtn:hover{background:var(--bad-soft);color:var(--bad)}
.actbtn.done{color:var(--good)}
tr.dropped .exbtn{color:var(--good)}

.pill{display:inline-block;padding:1px 8px;border-radius:99px;font-size:12px;font-weight:600;white-space:nowrap;margin:1px 0}
.pill.good{background:var(--good-soft);color:var(--good)}
.pill.mid{background:var(--mid-soft);color:var(--mid)}
.pill.warn{background:var(--warn-soft);color:var(--warn)}
.pill.bad{background:var(--bad-soft);color:var(--bad)}
.pill.na{background:var(--na-soft);color:var(--na)}
.rtag{display:inline-block;padding:1px 7px;border-radius:5px;font-size:12px;font-weight:700;white-space:nowrap}
.rtag.r4{background:var(--r4-soft);color:var(--r4)}
.rtag.r5{background:var(--r5-soft);color:var(--r5)}

/* ---------- the rendered paragraph / KID ---------- */
tr.docrow>td{padding:0 10px 14px;background:var(--sunk)}
.doc{background:var(--doc);border:1px solid var(--rule);border-left:3px solid var(--c,var(--accent));border-radius:4px var(--r-sm) var(--r-sm) 4px;
  padding:16px 22px 18px;margin-top:2px;box-shadow:var(--shadow)}
.dochead{display:flex;gap:10px;align-items:baseline;flex-wrap:wrap;font-size:13px;color:var(--muted);padding-bottom:10px;margin-bottom:12px;border-bottom:1px solid var(--rule)}
.dochead .cite{font-weight:600;color:var(--ink)}
.dochead .sp{flex:1}
.docbody{font-size:15px;line-height:1.7;color:var(--ink);max-width:92ch}
.docbody>*:first-child{margin-top:0}
.docbody>*:last-child{margin-bottom:0}
.docbody p{margin:0 0 .7em}
.docbody ul,.docbody ol{margin:0 0 .7em;padding-left:1.5em}
.docbody li{margin:.2em 0}
.docbody li::marker{color:var(--faint)}
.docbody h4,.docbody h5,.docbody h6{margin:1em 0 .4em;font-size:15px}
.docbody sup,.docbody sub{font-size:.7em;line-height:0}
.docbody sup{color:var(--accent)}
.docbody a{color:var(--accent);overflow-wrap:anywhere}
.docbody blockquote{margin:0 0 .7em;padding-left:14px;border-left:3px solid var(--rule2);color:var(--muted)}
.docbody hr{border:0;border-top:1px solid var(--rule);margin:1em 0}
.docbody .tblwrap{overflow-x:auto;margin:.4em 0 1em;max-width:100%}
.docbody table{border-collapse:collapse;font-size:13.5px;line-height:1.45;min-width:50%;background:var(--surface);border:1px solid var(--rule2)}
.docbody th,.docbody td{border:1px solid var(--rule2);padding:6px 10px;vertical-align:top;text-align:left}
.docbody thead th,.docbody thead td,.docbody tr:first-child>th{background:var(--sunk);font-weight:600}
.docbody td p,.docbody th p{margin:0 0 .4em}
.docbody td p:last-child,.docbody th p:last-child{margin-bottom:0}
.docbody .mbox,.docbody .boxed,.docbody div.shaded{display:block;background:var(--box);border:0;border-left:3px solid var(--box-rule);
  border-radius:2px var(--r-sm) var(--r-sm) 2px;padding:14px 18px;margin:.6em 0 1em;font-size:14.5px}
.docbody table.mbox{display:table;width:100%;border-collapse:separate}
.docbody table.mbox>*>tr>td,.docbody table.mbox>tr>td{border:0;padding:0;background:transparent;font-size:14.5px;line-height:1.65}
.docbody .mbox .boxtitle{font-weight:700;font-size:14px;color:var(--accent-strong);margin-bottom:.5em}
body.dark .docbody .mbox .boxtitle{color:var(--accent)}
.docbody td.shaded,.docbody th.shaded,.docbody tr.shaded>td{background:var(--box)}
.docbody .tblcap{font-size:13.5px;font-weight:600}

/* KID record: details, then the five steps in order */
.kidtop{display:flex;align-items:flex-start;gap:14px;flex-wrap:wrap;padding-bottom:12px;margin-bottom:14px;border-bottom:1px solid var(--rule)}
.kidtitle{font-size:17px;font-weight:700;letter-spacing:-.01em}
.kidtitle small{font-weight:500;color:var(--muted);font-size:13.5px;margin-left:6px}
.kidmeta{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px 22px;width:100%;margin-top:4px}
.kidmeta div{font-size:13.5px}
.kidmeta dt{font-size:12px;color:var(--muted);margin:0 0 1px}
.kidmeta dd{margin:0;font-weight:600}
.kidtop .sp{flex:1}
ol.steps{list-style:none;margin:0;padding:0;position:relative;max-width:96ch}
ol.steps::before{content:'';position:absolute;left:13px;top:12px;bottom:12px;width:2px;background:var(--rule)}
ol.steps>li{position:relative;padding:0 0 16px 44px}
ol.steps>li:last-child{padding-bottom:0}
.stepdot{position:absolute;left:0;top:0;width:28px;height:28px;border-radius:50%;display:grid;place-items:center;font-size:12.5px;font-weight:700;
  background:var(--surface);border:2px solid var(--rule2);color:var(--muted)}
li.hit .stepdot{background:var(--accent);border-color:var(--accent);color:var(--on-accent)}
.steplabel{font-size:13.5px;font-weight:700;margin:3px 0 6px;display:flex;align-items:center;gap:8px}
.steplabel .found{font-size:11.5px;font-weight:600;color:var(--accent);background:var(--accent-soft);border-radius:99px;padding:0 8px}
li.hit .stepbody{background:color-mix(in srgb,var(--accent-soft) 45%,var(--surface));border:1px solid color-mix(in srgb,var(--accent) 25%,var(--rule));border-radius:var(--r-sm);padding:10px 14px}
.stepbody.emptypart{color:var(--faint);font-size:13.5px;font-style:normal}
.decision{display:inline-block;white-space:nowrap;padding:1px 9px;border-radius:99px;font-size:12px;font-weight:700}
.decision.up{background:var(--good-soft);color:var(--good)}
.decision.down{background:var(--bad-soft);color:var(--bad)}
.decision.same{background:var(--na-soft);color:var(--na)}

.morebar{display:flex;gap:8px;margin-top:8px}
.morebtn{flex:1;padding:8px;border:1px dashed var(--rule2);background:var(--surface);color:var(--accent);border-radius:var(--r-sm);cursor:pointer;font-weight:600;font-size:13px}
.morebtn:hover{border-style:solid;background:var(--accent-soft)}
.morebtn.all{flex:0 0 auto;padding:8px 16px}
.empty{color:var(--faint);font-size:13px;padding:10px 2px}

.hist{display:flex;flex-direction:column}
.hist a{display:block;color:var(--ink);cursor:pointer;padding:7px 6px;text-decoration:none;font-size:13px;border-bottom:1px solid var(--rule);
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap;border-radius:4px}
.hist a:last-child{border-bottom:0}
.hist a:hover{color:var(--accent);background:var(--sunk)}
.hist a small{color:var(--faint);margin-right:6px}

@media (max-width:1080px){
  .layout{grid-template-columns:minmax(0,1fr);padding:16px}
  .rail{position:static;max-height:none;overflow:visible}
  .optrow,.exgrid{grid-template-columns:1fr}
  .covs{margin-left:0;flex-basis:100%;justify-content:flex-start}
  .fdlg-body{grid-template-columns:1fr;grid-template-rows:auto minmax(0,1fr)}
  .fdlg-nav{flex-direction:row;flex-wrap:wrap;border-right:0;border-bottom:1px solid var(--rule)}
  .navtab{width:auto}
  .navsep,.fdlg-nav .tabletoggles{display:none}
}
@media (max-width:640px){
  .top{padding:0 16px;gap:10px}
  .brandsub,.srcchip{display:none!important}
  .composer{padding:16px}
  .seg button{white-space:normal;line-height:1.25}
  table.results{min-width:700px}
  .chead .hint{margin-left:0}
  .termhead{padding:14px 16px}
  .termbody{padding:4px 12px 16px}
  .fdlg-foot{flex-wrap:wrap}
  .fdlg-foot .note{display:none}
  .doc{padding:14px}
}
</style>
</head>
<body>
<header class="top">
  <div class="brand">
    <span class="brandmark" aria-hidden="true"></span>
    <div>
      <div class="brandname">FATF Search Tool</div>
      <div class="brandsub">Mutual evaluations and key issues, __VERSION__</div>
    </div>
  </div>
  <span class="grow"></span>
  <div class="srcchip" id="srcChip"><span class="dot"></span><span id="srcChipTxt"></span></div>
  <div class="spinner" id="spin" role="status" aria-label="Working"></div>
  <button class="iconbtn" onclick="toggleDark()" id="darkbtn">Dark mode</button>
</header>

<div class="layout">
  <!-- ======================= RAIL ======================= -->
  <aside class="rail" aria-label="Source and filters">
    <section class="panel">
      <h2 class="ptitle"><span class="step">1</span>Choose what to search</h2>
      <div class="srcpick" id="modeSeg" role="radiogroup" aria-label="Source">
        <button class="srcopt on" data-v="MER" role="radio" onclick="setMode('MER')"><span class="radio"></span>
          <span><b>Mutual evaluations</b><span class="d">Every paragraph of the MERs: effectiveness, risk and context, technical compliance</span></span></button>
        <button class="srcopt" data-v="KID" role="radio" onclick="setMode('KID')"><span class="radio"></span>
          <span><b>Key issues</b><span class="d">Key Issue Documents: what was debated at Plenary and what was decided</span></span></button>
        <button class="srcopt" data-v="CASE" role="radio" onclick="setMode('CASE')"><span class="radio"></span>
          <span><b>Case studies</b><span class="d">The case study boxes in the MER effectiveness chapters, both rounds. Search words, or filter and list them all</span></span></button>
        <button class="srcopt" data-v="OTHER" role="radio" onclick="setMode('OTHER')"><span class="radio"></span>
          <span><b>Other files</b><span class="d">Your own Word, PDF, Excel, CSV and text files</span></span></button>
      </div>
      <div id="roundWrap">
        <div class="sublabel">Round</div>
        <div class="seg" id="roundSeg">
          <button data-v="4" onclick="setRound('4')">4th round</button>
          <button class="on" data-v="5" onclick="setRound('5')">5th round</button>
          <button data-v="both" onclick="setRound('both')">Both</button>
        </div>
        <div class="roundhint" id="roundHint"></div>
      </div>
      <div id="pathBox" class="pathtoggle">
        <button class="linkbtn" onclick="togglePaths()">Change source folders</button>
        <div id="pathFields" style="display:none">
          <label for="p4">4th round folder</label><input type="text" id="p4" placeholder="Leave blank for the default">
          <label for="p5">5th round folder</label><input type="text" id="p5" placeholder="Leave blank for the default">
          <label for="pk">Key Issue Documents folder</label><input type="text" id="pk" placeholder="Leave blank for the default">
          <label for="po">Other files folder</label><input type="text" id="po" placeholder="Leave blank for the default">
          <div class="hint">Paste the full folder path, e.g. C:\Users\you\Documents\MER data\round_4. The tool remembers it on this computer. Changes take effect when you load the data.</div>
        </div>
      </div>
      <button class="btn primary wide" style="margin-top:14px" onclick="loadData()">Load data</button>
      <div class="status" id="loadStatus">Not loaded yet.</div>
    </section>

    <section class="panel" id="filterCard">
      <div class="ftop">
        <h2 class="ptitle"><span class="step">2</span>Narrow it down <span class="sub" style="margin-left:4px">optional</span></h2>
        <button class="expandbtn" onclick="openFilters()" title="Open the filters in a large window">
          <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" aria-hidden="true"><path d="M9.5 2.5h4v4M13.5 2.5 9 7M6.5 13.5h-4v-4M2.5 13.5 7 9"/></svg>
          Expand
        </button>
      </div>
      <div class="fsummary" id="fSummary">Searching everything in the loaded reports.</div>
      <div class="tabletoggles" id="tableToggles"></div>
      <div id="facetRail"></div>
      <button class="btn grey wide" style="margin-top:12px" onclick="clearFilters()">Reset filters</button>
    </section>

    <section class="panel">
      <h2 class="ptitle">Recent searches</h2>
      <div class="hist" id="hist"><div class="empty">No searches yet.</div></div>
    </section>
  </aside>

  <!-- ======================= MAIN ======================= -->
  <main class="main">
    <section class="composer">
      <h1 class="ptitle" style="font-size:16px"><span class="step">3</span>Type your search terms</h1>
      <p class="csub" id="modeIntro"></p>
      <div class="chead">
        <label for="terms">One term per line</label>
        <span class="termcount" id="termCount">0 terms</span>
        <span class="hint"><span class="kbd">Ctrl</span> + <span class="kbd">Enter</span> to search</span>
      </div>
      <textarea id="terms" spellcheck="false" placeholder="beneficial ownership&#10;DPMS~dealers in precious metals and stones&#10;JIG (Joint Intelligence Group)&#10;ship-based casino"></textarea>
      <details class="syntax">
        <summary>How terms are matched</summary>
        <table>
          <tr><td><code>beneficial ownership</code></td><td>The words in this order. Case, spacing, hyphens, punctuation and accents are ignored.</td></tr>
          <tr><td><code>ship</code></td><td>Words that start with "ship" (ships, shipping), never "membership".</td></tr>
          <tr><td><code>DPMS~dealers in precious metals</code></td><td><code>~</code> joins names for the same thing into one result.</td></tr>
          <tr><td><code>JIG (Joint Intelligence Group)</code></td><td>JIG, the full name, or either bracketed form.</td></tr>
          <tr><td><code>SAMLIT (South Africa)</code></td><td>SAMLIT, in South Africa's report only.</td></tr>
          <tr><td><code>JMLIT+</code></td><td>Only "JMLIT+", not plain JMLIT.</td></tr>
        </table>
      </details>

      <div class="optrow">
        <div>
          <label>Search mode</label>
          <div class="seg" id="smSeg">
            <button class="on" data-v="strict" onclick="setSM('strict')">Strict keyword</button>
            <button data-v="exact" onclick="setSM('exact')">Exact phrase</button>
            <button data-v="fuzzy" onclick="setSM('fuzzy')">Fuzzy</button>
          </div>
          <div class="hint" id="smHint">Words must start with your term: "ship" finds ships and shipping, never membership.</div>
          <div id="thresholdWrap">
            <label for="threshold" style="margin-top:0">Fuzzy similarity: <span id="thVal" style="color:var(--accent)">85</span>%</label>
            <input type="range" min="70" max="100" value="85" id="threshold"
                   oninput="document.getElementById('thVal').textContent=this.value">
            <div class="hint">Lower catches more typos but also more noise. Filter by jurisdiction to keep fuzzy fast on the 4th round.</div>
          </div>
        </div>
        <div>
          <label>Acronyms</label>
          <label class="checkline" style="margin-top:4px"><input type="checkbox" id="capsOnly">
            <span>Match all-caps terms in capitals only, so ASC finds "ASC" but not "ascertain"</span></label>
        </div>
      </div>

      <div class="actions">
        <button class="btn primary" id="btnSearch" onclick="doSearch()">Search</button>
        <button class="btn ghost" id="btnCheck" onclick="doInterpret()">Check my terms</button>
        <button class="btn grey" onclick="clearAll()">Clear</button>
      </div>
      <div class="progwrap" id="progWrap"><div class="progbar"><div class="progfill" id="progFill"></div></div>
        <div class="progtxt" id="progTxt"></div></div>
      <div class="status" id="searchStatus"></div>
      <div id="interp"></div>

      <div class="exportbox">
        <h3 class="ptitle"><span class="step">4</span>Export the results</h3>
        <div class="exportnote">Removed hits (✕) are left out. Each quote comes with its citation.</div>
        <div class="exgrid">
          <div>
            <label for="reportName">Report name <span class="opt">(title and file name; blank uses the first term)</span></label>
            <input type="text" id="reportName" placeholder="e.g. DPMS in the 4th and 5th rounds">
          </div>
          <div>
            <label for="exportDir">Export folder</label>
            <input type="text" id="exportDir" value="__DEFAULT_EXPORT__">
          </div>
        </div>
        <div class="hint" id="fileNamePreview"></div>
        <div class="exportrow">
          <button class="btn ghost" id="btnExportX" onclick="doExport('xlsx')" disabled>Excel workbook</button>
          <button class="btn ghost" id="btnExportW" onclick="doExport('docx')" disabled>Word report</button>
          <button class="btn ghost" id="btnExportB" onclick="doExport('both')" disabled>Both</button>
          <select id="quoteStyle" title="How quotes appear in the Excel and Word exports" aria-label="Quote style in exports">
            <option value="paragraph" id="qsFull">Quotes: full paragraph</option>
            <option value="sentence">Quotes: sentence with the term only (smaller files)</option>
          </select>
        </div>
      </div>
    </section>

    <div id="results"></div>
  </main>
</div>

<!-- ======================= FILTERS POP-OUT ======================= -->
<dialog class="fdlg" id="fdlg" aria-labelledby="fdlgTitle">
  <div class="fdlg-head">
    <div>
      <h2 id="fdlgTitle">Filters</h2>
      <p id="dlgSummary"></p>
    </div>
    <span class="sp"></span>
    <button class="iconbtn" id="dlgMax" onclick="toggleMax()">Full screen</button>
    <button class="iconbtn" onclick="closeFilters()" aria-label="Close filters">Close</button>
  </div>
  <div class="fdlg-body">
    <nav class="fdlg-nav" id="dlgNav" aria-label="Filter groups"></nav>
    <section class="fdlg-pane">
      <div class="fdlg-tools">
        <input type="search" id="dlgSearch" placeholder="Find…" aria-label="Find in this filter">
        <button class="mini" id="dlgAll" onclick="facetAll(curTab,true)">Select all</button>
        <button class="mini" id="dlgNone" onclick="facetAll(curTab,false)">Deselect all</button>
        <button class="mini" onclick="facetInvert(curTab)">Invert</button>
        <button class="mini" onclick="setTree(curTab,true)">Expand groups</button>
        <button class="mini" onclick="setTree(curTab,false)">Collapse groups</button>
        <label class="checkline"><input type="checkbox" id="dlgOnly"> Show selected only</label>
      </div>
      <div class="presets" id="dlgPresets"></div>
      <div class="fdlg-list" id="dlgList"></div>
    </section>
  </div>
  <div class="fdlg-foot">
    <button class="btn grey" onclick="clearFilters()">Reset all filters</button>
    <span class="note" id="dlgFootNote">Filters apply to your next search.</span>
    <button class="btn ghost" id="dlgRerun" onclick="closeFilters();doSearch()" style="display:none">Search again with these filters</button>
    <button class="btn primary" onclick="closeFilters()">Done</button>
  </div>
</dialog>

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
const INTRO={
  MER:'Searches every paragraph of the mutual evaluation reports. Each hit opens to the paragraph as it appears in the report, tables and boxes included.',
  KID:'Searches the Key Issue Documents: the key issue, the ECG discussion and proposal, and the Plenary discussion and decision. Each hit is one KID, and opens to the full record.',
  CASE:'Searches the case study boxes in the effectiveness chapters of the MERs. Type words to narrow them down, or leave the box empty and press Search to list every case study that matches your filters.',
  OTHER:'Searches your own files in the "other" folder, and shows exactly where each hit is.'};
function welcomeHtml(){
  return '<div class="welcome"><h2>Your results will appear here</h2><p>'+INTRO[dataMode]+'</p>'
    +'<ol><li><span class="step">1</span>Choose a source</li><li><span class="step">2</span>Narrow it down if you like</li>'
    +'<li><span class="step">3</span>Type terms and search</li><li><span class="step">4</span>Export</li></ol></div>';}

function $(id){return document.getElementById(id);}
function esc(s){return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function newSid(){return Date.now().toString(36)+Math.random().toString(36).slice(2,10);}
async function post(url,body){
  const r=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
  return r.json();
}
const PLURALS={'case study':'case studies'};
function plural(n,w){return n.toLocaleString()+' '+(n===1?w:(PLURALS[w]||w+'s'));}
function fileNamePreview(){
  const n=($('reportName').value||'').trim();
  const t=($('terms').value.split('\n').map(x=>x.trim()).filter(Boolean)[0]||(dataMode==='CASE'?'All case studies':'search'));
  const d=new Date(); const ds=d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
  $('fileNamePreview').innerHTML='Saves as '+esc(ds+'_'+(n||t)+'_FATF Search tool')+'.xlsx / .docx. '
    +'<button class="linkbtn" onclick="openFolder()">Open export folder</button>';
}
function togglePaths(){const f=$('pathFields');f.style.display=(f.style.display==='none'?'block':'none');}

function segSet(id,v){document.querySelectorAll('#'+id+' button').forEach(b=>b.classList.toggle('on',b.dataset.v===v));}
function applyModeUI(){
  segSet('modeSeg',dataMode);
  document.querySelectorAll('#modeSeg button').forEach(b=>b.setAttribute('aria-checked',b.dataset.v===dataMode));
  const st=dataMode!=='OTHER';
  $('roundWrap').style.display=st?'':'none';
  $('filterCard').style.display=st?'':'none';
  $('modeIntro').textContent=INTRO[dataMode];
  $('qsFull').textContent=dataMode==='KID'?'Quotes: the whole KID (all five parts)':
    dataMode==='CASE'?'Quotes: the whole case study':'Quotes: full paragraph';
  initFacets();roundHint();
  if(!lastPayload)$('results').innerHTML=welcomeHtml();}
function setMode(v){if(v===dataMode)return;dataMode=v;
  if(v==='CASE'&&round!=='both'){round='both';segSet('roundSeg','both');}
  applyModeUI();loaded=false;setChip(null);countTerms();fileNamePreview();loadData();}
function setRound(v){if(v===round&&loaded)return;round=v;segSet('roundSeg',v);roundHint();loaded=false;setChip(null);loadData();}
function roundHint(){
  const h=dataMode==='CASE'?{'4':'Case studies from 4th round reports.','5':'Case studies from 5th round reports so far.',
    'both':'Both rounds together. Every case study is labelled with its round, country, IO, core issue and rating.'}:dataMode==='KID'?{'4':'Key issues from 4th round evaluations.','5':'Key issues from 5th round evaluations so far.',
    'both':'Both rounds together. Every KID is labelled with its round, and statistics are shown per round, never added together.'}:{'4':'4th round reports (2013 Methodology).','5':'5th round reports (2022 Methodology) published so far.',
    'both':'Both rounds searched together. Every hit is labelled with its round, and every statistic is shown per round, never added together.'};
  $('roundHint').textContent=h[round]||'';}
function setSM(v){searchMode=v;segSet('smSeg',v);
  $('thresholdWrap').style.display=(v==='fuzzy')?'block':'none';
  const hints={strict:'Words must start with your term: "ship" finds ships and shipping, never membership.',
    exact:'Whole words only: "ship" will not match ships or shipping.',
    fuzzy:'Catches typos and close spellings ("cassino" finds "casino"). Slower on large tables.'};
  $('smHint').textContent=hints[v];}
function spin(on){$('spin').style.display=on?'inline-block':'none';}
function toggleDark(){document.body.classList.toggle('dark');
  const d=document.body.classList.contains('dark');
  $('darkbtn').textContent=d?'Light mode':'Dark mode';
  try{localStorage.setItem('fatf_dark',d?'1':'0');}catch(e){}}
try{const pref=localStorage.getItem('fatf_dark');
  if(pref==='1'||(pref===null&&window.matchMedia&&matchMedia('(prefers-color-scheme: dark)').matches))toggleDark();}catch(e){}
function setChip(txt){const c=$('srcChip');
  if(!txt){c.style.display='none';return}
  $('srcChipTxt').textContent=txt;c.style.display='flex';}
function paths(){return {round_4:$('p4').value.trim(),round_5:$('p5').value.trim(),kid:$('pk').value.trim(),other:$('po').value.trim()};}
/* folders typed here are remembered on this computer, so data kept outside
   the tool's folder only has to be pointed to once */
function restorePaths(){
  try{const p=JSON.parse(localStorage.getItem('fatf_paths')||'{}');
    $('p4').value=p.round_4||'';$('p5').value=p.round_5||'';$('pk').value=p.kid||'';$('po').value=p.other||'';
    if(p.round_4||p.round_5||p.kid||p.other)$('pathFields').style.display='block';}catch(e){}
  ['p4','p5','pk','po'].forEach(id=>$(id).addEventListener('change',()=>{
    try{localStorage.setItem('fatf_paths',JSON.stringify(paths()));}catch(e){}}));}

/* =====================================================================
   Filters: one state, drawn in the rail and in the pop-out
   ===================================================================== */
const FDEFS={
  MER:[
    {k:'jur', fk:'jurisdictions', title:'Jurisdictions', noun:'jurisdiction', ph:'Find a jurisdiction'},
    {k:'ind', fk:'indicators', title:'Sections', noun:'section', ph:'Find an IO, recommendation or criterion, e.g. IO.6 or C.20'},
    {k:'rat', fk:'ratings', title:'Ratings', noun:'rating', ph:'Find a rating'},
    {k:'crit',fk:'crit_ratings', title:'Criterion ratings', noun:'criterion rating', ph:'Find a criterion rating'}],
  KID:[
    {k:'jur', fk:'jurisdictions', title:'Jurisdictions', noun:'jurisdiction', ph:'Find a jurisdiction'},
    {k:'ind', fk:'indicators', title:'IOs and Recommendations', noun:'IO or Recommendation', ph:'Find an IO or Recommendation, e.g. IO5 or R.24'},
    {k:'theme',fk:'themes', title:'Themes', noun:'theme', ph:'Find a theme, e.g. risk, beneficial ownership'},
    {k:'part', fk:'parts', title:'Where to look in the KID', noun:'part', ph:'Find a part'},
    {k:'outc', fk:'outcomes', title:'Plenary decision', noun:'decision', ph:'Find a decision'},
    {k:'plen', fk:'plenaries', title:'Plenary meeting', noun:'Plenary meeting', ph:'Find a year or month, e.g. 2019'},
    {k:'rat', fk:'ratings', title:'Final rating', noun:'rating', ph:'Find a rating'}],
  CASE:[
    {k:'kind', fk:'boxkinds', title:'Type of box', noun:'box type', ph:'Find a type'},
    {k:'jur', fk:'jurisdictions', title:'Jurisdictions', noun:'jurisdiction', ph:'Find a jurisdiction'},
    {k:'ind', fk:'indicators', title:'IOs and core issues', noun:'core issue', ph:'Find an IO or core issue, e.g. IO.6'},
    {k:'rat', fk:'ratings', title:'Effectiveness rating', noun:'rating', ph:'Find a rating'}]};
const TDEFS={
  MER:[{id:'tEff',v:'Effectiveness',t:'Effectiveness',c:'--eff'},{id:'tRC',v:'Risk and Context',t:'Risk and context',c:'--rc'},
       {id:'tTC',v:'Technical Compliance',t:'Technical compliance',c:'--tc'}],
  KID:[{id:'tEff',v:'Effectiveness',t:'Effectiveness (IOs)',c:'--eff'},{id:'tTC',v:'Technical Compliance',t:'Technical compliance (Rs)',c:'--tc'},
       {id:'tGen',v:'General',t:'General',c:'--gen'}],
  CASE:[]};
let FDEF=[], TDEF=[];
const F={};
let LABELS={}, curTab='jur', gseq=0;
function initFacets(){
  FDEF=FDEFS[dataMode]||[];TDEF=TDEFS[dataMode]||[];
  Object.keys(F).forEach(k=>delete F[k]);
  FDEF.forEach(d=>F[d.k]=Object.assign({},d,{tree:[],values:[],sel:new Set(),q:'',presets:[],labels:{},gidx:{},openI:{},openD:{}}));
  curTab='jur';
  $('tableToggles').innerHTML=TDEF.map(t=>'<label class="ttog" style="--c:var('+t.c+')"><input type="checkbox" id="'+t.id+'" checked onchange="syncTables(this)"><span class="sw"></span>'+esc(t.t)+'</label>').join('');
  buildRail();refreshAll();}

function lab(k,v){const f=F[k];return (f&&f.labels&&f.labels[v])||((k==='rat'||k==='crit')?LABELS[v]:null)||v;}
function setFacet(k,data){
  const f=F[k];
  const wasAll=!f.values.length||f.sel.size===f.values.length;
  f.tree=(data&&data.tree)||[];f.presets=(data&&data.presets)||[];f.labels=(data&&data.labels)||{};f.gidx={};
  const vals=[];
  const walk=n=>{n._g='g'+(++gseq);let its=[];
    (n.items||[]).forEach(v=>{its.push(v);vals.push(v);});
    (n.children||[]).forEach(c=>{its=its.concat(walk(c));});
    f.gidx[n._g]=its;return its;};
  f.tree.forEach(walk);
  f.values=[...new Set(vals)];
  if(wasAll&&data&&data.default&&!f.touched)f.sel=new Set(data.default.filter(v=>f.values.includes(v)));
  else if(wasAll)f.sel=new Set(f.values);
  else{f.sel=new Set(f.values.filter(v=>f.sel.has(v)));if(!f.sel.size)f.sel=new Set(f.values);}
}
function facetVals(k){const f=F[k];return (!f.values.length||f.sel.size===f.values.length)?[]:[...f.sel];}
function countText(k){const f=F[k],n=f.sel.size,t=f.values.length;
  if(!t)return {t:'None loaded',c:''};
  if(n===t)return {t:'All '+t,c:''};
  if(!n)return {t:'None selected',c:'none'};
  return {t:n+' of '+t,c:'some'};}
function qmatch(s,q){return !q||String(s).toLowerCase().includes(q);}
function hlq(s,q){const e=esc(s);if(!q)return e;
  const i=String(s).toLowerCase().indexOf(q);if(i<0)return e;
  return esc(String(s).slice(0,i))+'<span class="q">'+esc(String(s).slice(i,i+q.length))+'</span>'+esc(String(s).slice(i+q.length));}

function itemVisible(k,v,q,groupHit,ctx){
  const f=F[k];
  if(ctx==='D'&&$('dlgOnly').checked&&!f.sel.has(v))return false;
  return groupHit||qmatch(lab(k,v),q);}
function itemHtml(k,v,q){
  const on=F[k].sel.has(v);
  return '<label class="fitem'+(on?'':' off')+'"><input type="checkbox" data-v="'+esc(v)+'"'+(on?' checked':'')+'><span>'+hlq(lab(k,v),q)+'</span></label>';}
function nodeHtml(k,n,depth,ctx,q,parentHit){
  const f=F[k];
  const hit=parentHit||!!(q&&n.name&&qmatch(n.name,q));
  const items=(n.items||[]).filter(v=>itemVisible(k,v,q,hit,ctx));
  const kids=(n.children||[]).map(c=>nodeHtml(k,c,depth+1,ctx,q,hit)).filter(Boolean);
  if(!items.length&&!kids.length)return '';
  const itemsHtml=items.length?'<div class="fitems">'+items.map(v=>itemHtml(k,v,q)).join('')+'</div>':'';
  if(!n.name)return itemsHtml+kids.join('');
  const all=f.gidx[n._g]||[];const nsel=all.filter(v=>f.sel.has(v)).length;
  const om=ctx==='D'?f.openD:f.openI;
  const def=ctx==='D'?true:(depth===0);
  const open=(q||(ctx==='D'&&$('dlgOnly').checked))?true:(n._g in om?om[n._g]:def);
  return '<div class="fgroup'+(open?' open':'')+'"><div class="fghead">'
    +'<button class="fcaret" data-tog="'+n._g+'" aria-expanded="'+open+'" aria-label="Show or hide '+esc(n.name)+'"></button>'
    +'<label class="fgcheck"><input type="checkbox" data-g="'+n._g+'"'+(nsel===all.length?' checked':'')+(nsel&&nsel<all.length?' data-ind="1"':'')+'>'
    +'<span class="fgname">'+hlq(n.name,hit&&!parentHit?q:'')+'</span></label>'
    +'<span class="fgcount">'+nsel+'/'+all.length+'</span></div>'
    +(open?'<div class="fgbody">'+itemsHtml+kids.join('')+'</div>':'')+'</div>';}
function treeHtml(k,ctx){
  const f=F[k];const q=(f.q||'').trim().toLowerCase();
  const h=f.tree.map(n=>nodeHtml(k,n,0,ctx,q,false)).join('');
  if(h)return h;
  if(!f.values.length)return '<div class="fempty">Load the data to see the options.</div>';
  return '<div class="fempty">Nothing matches '+(q?'"'+esc(q)+'"':'this view')+'.</div>';}
function paintIndeterminate(el){el.querySelectorAll('input[data-ind]').forEach(i=>i.indeterminate=true);}

function buildRail(){
  $('facetRail').innerHTML=FDEF.map(d=>
    '<div class="facet'+(d.k==='jur'?'':' closed')+'" id="fx_'+d.k+'">'
    +'<button class="facet-head" onclick="this.parentNode.classList.toggle(\'closed\')"><span class="caret"></span>'
    +'<span>'+esc(d.title)+'</span><span class="facet-count" id="fc_'+d.k+'"></span></button>'
    +'<div class="facet-body"><div class="facet-tools">'
    +'<input type="search" class="fsearch" data-f="'+d.k+'" placeholder="'+esc(d.ph.split(',')[0])+'" aria-label="'+esc(d.ph)+'">'
    +'<button class="mini" onclick="facetAll(\''+d.k+'\',true)" title="Select all shown">All</button>'
    +'<button class="mini" onclick="facetAll(\''+d.k+'\',false)" title="Deselect all shown">None</button>'
    +'<button class="mini" onclick="openFilters(\''+d.k+'\')" title="Open in the large window">Expand</button></div>'
    +'<div class="flist" id="fl_'+d.k+'" data-f="'+d.k+'" data-ctx="I"></div></div></div>').join('');
  document.querySelectorAll('.fsearch').forEach(inp=>inp.addEventListener('input',()=>{
    const k=inp.dataset.f;F[k].q=inp.value;if(curTab===k&&$('fdlg').open)$('dlgSearch').value=inp.value;refreshFacet(k);}));
  FDEF.forEach(d=>bindList($('fl_'+d.k)));
}
function bindList(el){
  el.addEventListener('change',e=>{
    const t=e.target;const k=el.dataset.f;const f=F[k];if(!f||t.type!=='checkbox')return;
    f.touched=true;
    if(t.dataset.v!==undefined){t.checked?f.sel.add(t.dataset.v):f.sel.delete(t.dataset.v);}
    else if(t.dataset.g){const its=f.gidx[t.dataset.g]||[];
      const allOn=its.every(v=>f.sel.has(v));its.forEach(v=>allOn?f.sel.delete(v):f.sel.add(v));}
    refreshFacet(k);});
  el.addEventListener('click',e=>{
    const b=e.target.closest('[data-tog]');if(!b)return;
    const k=el.dataset.f;const f=F[k];const om=el.dataset.ctx==='D'?f.openD:f.openI;
    om[b.dataset.tog]=b.getAttribute('aria-expanded')!=='true';refreshFacet(k);});
}
function refreshFacet(k){
  const rl=$('fl_'+k);
  if(rl){const st=rl.scrollTop;rl.innerHTML=treeHtml(k,'I');paintIndeterminate(rl);rl.scrollTop=st;}
  const c=countText(k);const fc=$('fc_'+k);if(fc){fc.textContent=c.t;fc.className='facet-count '+c.c;}
  if($('fdlg').open){
    if(k===curTab){const L=$('dlgList');const st=L.scrollTop;L.dataset.f=k;L.innerHTML=treeHtml(k,'D');paintIndeterminate(L);L.scrollTop=st;}
    renderNav();
  }
  renderSummary();
}
function refreshAll(){FDEF.forEach(d=>refreshFacet(d.k));}
function shownValues(k,ctx){
  const f=F[k];const q=(f.q||'').trim().toLowerCase();const out=new Set();
  const walk=(n,ph)=>{const hit=ph||!!(q&&n.name&&qmatch(n.name,q));
    (n.items||[]).forEach(v=>{if(itemVisible(k,v,q,hit,ctx))out.add(v);});(n.children||[]).forEach(c=>walk(c,hit));};
  f.tree.forEach(n=>walk(n,false));return out;}
function facetAll(k,on){
  const f=F[k];const ctx=($('fdlg').open&&k===curTab)?'D':'I';
  const vis=(f.q||(ctx==='D'&&$('dlgOnly').checked))?shownValues(k,ctx):new Set(f.values);
  vis.forEach(v=>on?f.sel.add(v):f.sel.delete(v));refreshFacet(k);}
function facetInvert(k){const f=F[k];f.sel=new Set(f.values.filter(v=>!f.sel.has(v)));refreshFacet(k);}
function setTree(k,open){const f=F[k];Object.keys(f.gidx).forEach(g=>f.openD[g]=open);refreshFacet(k);}
function applyPreset(k,i){const f=F[k];const p=f.presets[i];if(!p)return;
  f.sel=new Set(p.items.filter(v=>f.values.includes(v)));refreshFacet(k);}
function syncTables(src){
  if(src){document.querySelectorAll('input[data-mirror="'+src.id+'"]').forEach(m=>m.checked=src.checked);}
  renderSummary();}
function mirrorTable(m){const src=$(m.dataset.mirror);src.checked=m.checked;syncTables(src);}
function tablesOn(){return TDEF.filter(t=>$(t.id)&&$(t.id).checked).map(t=>t.v);}
function summaryText(){
  const parts=[];
  FDEF.forEach(d=>{const f=F[d.k];if(f.values.length&&f.sel.size!==f.values.length)
    parts.push(f.sel.size?plural(f.sel.size,d.noun)+' of '+f.values.length:'no '+d.noun+'s');});
  const tb=tablesOn();if(tb.length<TDEF.length)parts.push(tb.length?tb.join(' and ').toLowerCase()+' only':'no tables');
  return parts.length?'Narrowed to '+parts.join(', ')+'.':(dataMode==='KID'?'Searching every Key Issue Document loaded.':'Searching everything in the loaded reports.');}
function renderSummary(){const s=summaryText();$('fSummary').textContent=s;if($('fdlg').open)$('dlgSummary').textContent=s;}
function renderNav(){
  $('dlgNav').innerHTML=FDEF.map(d=>{const c=countText(d.k);
    return '<button class="navtab'+(d.k===curTab?' on':'')+'" onclick="showTab(\''+d.k+'\')"><b>'+esc(d.title)+'</b>'
      +'<span class="'+c.c+'">'+esc(c.t)+'</span></button>';}).join('')
    +'<div class="navsep">Tables</div><div class="tabletoggles">'
    +TDEF.map(t=>'<label class="ttog" style="--c:var('+t.c+')"><input type="checkbox" data-mirror="'+t.id+'"'+($(t.id).checked?' checked':'')
      +' onchange="mirrorTable(this)"><span class="sw"></span>'+esc(t.t)+'</label>').join('')+'</div>';}
function showTab(k){
  curTab=k;const f=F[k];const d=FDEF.find(x=>x.k===k);
  $('dlgSearch').value=f.q||'';$('dlgSearch').placeholder=d.ph;
  $('dlgPresets').innerHTML=f.presets.length?'Quick picks: '+f.presets.map((p,i)=>
    '<button class="preset" onclick="applyPreset(\''+k+'\','+i+')" title="Select exactly these">'+esc(p.name)+'<small>'+p.items.length+'</small></button>').join(''):'';
  refreshFacet(k);}
function openFilters(k){
  const dlg=$('fdlg');
  if(!dlg.open){if(dlg.showModal)dlg.showModal();else dlg.setAttribute('open','');}
  $('dlgRerun').style.display=lastPayload?'':'none';
  showTab(k||curTab);renderSummary();
  setTimeout(()=>$('dlgSearch').focus(),30);}
function closeFilters(){const dlg=$('fdlg');if(dlg.open){if(dlg.close)dlg.close();else dlg.removeAttribute('open');}}
function toggleMax(){const d=$('fdlg');d.classList.toggle('max');$('dlgMax').textContent=d.classList.contains('max')?'Exit full screen':'Full screen';}
function clearFilters(){
  FDEF.forEach(d=>{const f=F[d.k];f.sel=new Set(f.values);f.q='';});
  document.querySelectorAll('.fsearch').forEach(i=>i.value='');$('dlgSearch').value='';
  TDEF.forEach(t=>{if($(t.id))$(t.id).checked=true;});
  if($('fdlg').open)renderNav();
  refreshAll();}

/* =====================================================================
   Loading
   ===================================================================== */
async function loadData(){
  spin(true);
  const st=$('loadStatus');
  st.classList.remove('err','ok');
  st.textContent='Loading… the very first load of a round takes a while; after that it is quick, even after a restart.';
  const key=dataMode+'|'+round+'|'+JSON.stringify(paths());
  try{
    const j=await post('/api/load',{data_mode:dataMode,round:round,paths:paths()});
    if(!j.ok){st.textContent='Could not load: '+j.error;st.classList.add('err');loaded=false;return;}
    loaded=true;loadedKey=key;
    if(dataMode!=='OTHER'){
      const unit=j.unit==='KID'?'KIDs':j.unit==='CASE'?'boxes':'paragraphs';
      st.classList.add('ok');
      st.textContent=(j.file?'Loaded '+j.file+'\n':'')+j.rounds.map(r=>r.round+': '+r.rows.toLocaleString()+' '+unit+', '
        +r.n_jurisdictions+' jurisdictions\n'
        +Object.entries(r.counts).map(([k,v])=>'  '+k+': '+v.toLocaleString()).join('\n')).join('\n');
      LABELS=j.labels||{};
      const fx=j.facets||{
        jur:{tree:[{name:'',items:j.jurisdictions}]},ind:{tree:[{name:'',items:j.indicators}]},
        rat:{tree:[{name:'',items:j.ratings.map(r=>r.code)}]},crit:{tree:[{name:'',items:j.crit_ratings.map(r=>r.code)}]}};
      FDEF.forEach(d=>setFacet(d.k,fx[d.k]));
      refreshAll();
      if($('fdlg').open)showTab(curTab);
      const src=dataMode==='KID'?'Key issues':dataMode==='CASE'?'Case studies':'Mutual evaluations';
      if(j.rounds.length===2)
        setChip(src+', 4th and 5th rounds, '+j.total.toLocaleString()+' '+unit);
      else setChip(src+', '+j.rounds[0].round+', '+j.n_jurisdictions+' jurisdictions, '+j.total.toLocaleString()+' '+unit);
    }else{
      st.textContent='Other files:\n'+j.files.map(f=>'  '+f.file+': '+
        (f.error?('not readable ('+f.error+')'):(f.rows.toLocaleString()+' searchable pieces'))).join('\n');
      setChip('Other files, '+plural(j.files.length,'file'));
    }
  }catch(e){st.textContent='Could not load: '+e;st.classList.add('err');loaded=false;}
  finally{spin(false);}
}

function payload(){
  const terms=$('terms').value.split('\n').map(t=>t.trim()).filter(Boolean);
  return{data_mode:dataMode,round:round,terms:terms,search_mode:searchMode,
    caps_only:$('capsOnly').checked,
    report_name:$('reportName').value,
    excluded:[...excluded],
    paths:paths(),
    threshold:$('threshold').value,
    export_dir:$('exportDir').value,
    quote_style:$('quoteStyle').value,
    filters:Object.assign({tables:tablesOn().length===TDEF.length?[]:tablesOn()},
      Object.fromEntries(FDEF.map(d=>[d.fk,facetVals(d.k)])))};
}
function filterProblem(){
  if(dataMode==='OTHER')return null;
  for(const d of FDEF){const f=F[d.k];if(f.values.length&&!f.sel.size)return {k:d.k,msg:'No '+d.noun+'s are selected. Tick at least one, or select all.'};}
  if(TDEF.length&&!tablesOn().length)return {k:curTab,msg:'No tables are selected. Turn on at least one of '+TDEF.map(t=>t.t).join(', ')+'.'};
  return null;}

/* =====================================================================
   Highlighting
   ===================================================================== */
function reEsc(s){return s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');}
function hl(text,matched,variants){
  let vs=(variants&&variants.length)?variants.slice():(matched?[matched]:[]);
  vs=vs.filter(Boolean).sort((a,b)=>b.length-a.length);
  const e=esc(text);
  if(!vs.length)return e;
  const alt=vs.map(v=>reEsc(esc(v))).join('|');
  return e.replace(new RegExp(alt,'gi'),x=>'<mark>'+x+'</mark>');
}
/* Mark the matched wording inside rendered HTML, even where it crosses
   a tag (a bold word, a table cell) or the report uses a non-breaking space. */
function markDom(root,matched,variants){
  let vs=(variants&&variants.length)?variants.slice():(matched?[matched]:[]);
  vs=vs.filter(Boolean).sort((a,b)=>b.length-a.length);
  if(!vs.length)return 0;
  const re=new RegExp(vs.map(v=>reEsc(v.trim()).replace(/\s+/g,'[\\s\\u00a0]*')).join('|'),'gi');
  const nodes=[];let full='';
  const w=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);let n;
  while((n=w.nextNode())){nodes.push({n:n,s:full.length});full+=n.nodeValue;}
  const ranges=[];let m;
  while((m=re.exec(full))){if(!m[0].length){re.lastIndex++;continue;}ranges.push([m.index,m.index+m[0].length]);}
  for(let i=nodes.length-1;i>=0;i--){
    const node=nodes[i].n,s=nodes[i].s,e=s+node.nodeValue.length;
    const segs=ranges.filter(r=>r[0]<e&&r[1]>s).map(r=>[Math.max(r[0],s)-s,Math.min(r[1],e)-s]);
    for(let j=segs.length-1;j>=0;j--){
      const a=segs[j][0],b=segs[j][1];if(b<=a)continue;
      node.splitText(b);const mid=node.splitText(a);
      if(!mid.nodeValue.trim())continue;
      const mk=document.createElement('mark');mid.parentNode.insertBefore(mk,mid);mk.appendChild(mid);}
  }
  return ranges.length;
}
/* Give the report's own structures their shape: single-cell tables are
   FATF "boxes" (case studies, examples); wide tables scroll sideways. */
function shapeDoc(body){
  body.querySelectorAll('table').forEach(t=>{
    const cells=t.querySelectorAll(':scope>tr>td,:scope>tr>th,:scope>*>tr>td,:scope>*>tr>th');
    const firstText=(t.textContent||'').trim();
    if(cells.length===1||(t.classList.contains('shaded')&&cells.length<=2)||/^Box\s+\d/i.test(firstText)&&cells.length<=2){
      t.classList.add('mbox');
    }else if(!t.closest('.tblwrap')){
      const wrap=document.createElement('div');wrap.className='tblwrap';t.parentNode.insertBefore(wrap,t);wrap.appendChild(t);
    }
  });
  body.querySelectorAll('.mbox,.boxed,div.shaded').forEach(b=>{
    const first=b.querySelector('p,div,b,strong');
    if(first&&/^\s*Box\s+\d/i.test(first.textContent||''))first.classList.add('boxtitle');
  });
  body.querySelectorAll('p').forEach(p=>{
    const t=(p.textContent||'').trim();
    if(/^(Table|Figure|Chart)\s+\d[\d.]*\.?\s/i.test(t)&&t.length<220&&p.nextElementSibling&&
       (p.nextElementSibling.matches('table,.tblwrap')))p.classList.add('tblcap');
  });
}
function fmtTags(html){
  if(!html)return '';
  const tags=[];
  const nTables=(html.match(/<table/gi)||[]).length;
  const isBox=/class="[^"]*boxed/.test(html)||/Box\s+\d/.test(html);
  if(isBox)tags.push('Box');
  if(nTables&&!(isBox&&nTables===1))tags.push(nTables>1?nTables+' tables':'Table');
  if(/<(ul|ol)[\s>]/i.test(html))tags.push('List');
  return tags.map(t=>'<span class="fmt">'+t+'</span>').join('');
}

function pillClass(rating){
  const code=(rating||'').split(' ')[0].toUpperCase();
  if(['C','HE','M'].includes(code))return 'good';
  if(['LC','SE','MM'].includes(code))return 'mid';
  if(['PC','ME','PM'].includes(code))return 'warn';
  if(['NC','LE','NM'].includes(code))return 'bad';
  return 'na';
}
function pill(rating){if(!rating)return'';return '<span class="pill '+pillClass(rating)+'" title="'+esc(LABELS[rating]||rating)+'">'+esc(rating)+'</span>';}
function rtag(lab){return '<span class="rtag '+(RCLS[lab]||'')+'">'+esc(RSHORT[lab]||lab)+'</span>';}

async function doInterpret(){
  const p=payload();
  if(p.terms.length===0){alert('Enter at least one search term.');return;}
  spin(true);
  try{
    const j=await post('/api/interpret',p);
    const box=$('interp');
    if(!j.ok){box.style.display='block';box.innerHTML='<div class="status err">'+esc(j.error)+'</div>';return;}
    let h='<table class="itable"><thead><tr><th style="width:22%">You typed</th>'
      +'<th style="width:24%">Searched as</th><th>How it will be matched</th></tr></thead><tbody>';
    j.terms.forEach(t=>{
      const al=t.aliases.map(a=>'<code>'+esc(a)+'</code>').join(' or ');
      h+='<tr><td><b>'+esc(t.input)+'</b>'+(t.scope?'<br><span class="scopetag">only '+esc(t.scope)+'</span>':'')+'</td>'
        +'<td>'+al+(t.label?'<div class="inote" style="margin-top:4px">Sheet: <b>'+esc(t.label)+'</b></div>':'')+'</td><td>'
        +t.notes.map(n=>'<div class="inote">'+esc(n)+'</div>').join('')
        +t.how.map(n=>'<div class="inote">'+esc(n)+'</div>').join('')
        +(t.warn.length?t.warn.map(n=>'<div class="iwarn">'+esc(n)+'</div>').join('')
                        :'<div class="iok">No ambiguity found</div>')
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
        $('progTxt').textContent=(j.finished?'Finishing…':'Searching term '+(j.done+1)+' of '+j.total+(j.term?': “'+j.term+'”':''));}
    }catch(e){}
  },500);
}
function stopProgress(){if(progTimer)clearInterval(progTimer);progTimer=null;$('progWrap').style.display='none';}

async function doSearch(){
  if(!loaded||loadedKey!==dataMode+'|'+round+'|'+JSON.stringify(paths())){await loadData();if(!loaded)return;}
  const p=payload();
  const ss=$('searchStatus');ss.classList.remove('err');ss.textContent='';
  if(p.terms.length===0&&dataMode!=='CASE'){ss.textContent='Enter at least one search term.';ss.classList.add('err');$('terms').focus();return;}
  const fp=filterProblem();
  if(fp){ss.textContent=fp.msg;ss.classList.add('err');openFilters(fp.k);return;}
  p.sid=newSid();
  spin(true);$('btnSearch').disabled=true;
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
    let msg='Found '+plural(tot,unitOf())+' across '+plural(j.results.length,'term')+'.';
    if(j.rounds.length===2){
      const r4=j.results.reduce((a,b)=>a+(b.by_round['4th round']||0),0),r5=j.results.reduce((a,b)=>a+(b.by_round['5th round']||0),0);
      msg+=' 4th round: '+r4.toLocaleString()+'. 5th round: '+r5.toLocaleString()+'.';
    }
    const fs=summaryText();if(!fs.startsWith('Searching everything'))msg+=' '+fs;
    ss.textContent=msg;
  }catch(e){ss.textContent='Search failed: '+e;ss.classList.add('err');}
  finally{stopProgress();spin(false);$('btnSearch').disabled=false;}
}

const SECS={'Effectiveness':'var(--eff)','Risk and Context':'var(--rc)','Technical Compliance':'var(--tc)','General':'var(--gen)'};

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
  if(any)h+='<div style="flex-basis:100%;text-align:right"><button class="linkbtn covcountries" onclick="toggleChips('+bi+')">Show jurisdictions</button></div>';
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
    h+='<div><div class="covlabel" style="margin:2px 0 6px">Jurisdictions assessed in both rounds</div><table class="cmpt"><tr><th>Jurisdiction</th><th>4th</th><th>5th</th><th></th></tr>'
      +block.round_compare.map(c=>'<tr><td>'+esc(c.name)+'</td><td>'+c.r4+'</td><td>'+c.r5+'</td><td><span class="pat '+pc[c.pattern]+'">'+esc(c.pattern)+'</span></td></tr>').join('')
      +'</table></div>';
  }
  return h;
}
function secCount(sec,rnd,two){
  if(rnd!=='all'){const n=sec.by_round[rnd]||0;return plural(n,unitOf())+' in the '+rnd;}
  let s=plural(sec.total,unitOf());
  if(two)s+=', 4th '+(sec.by_round['4th round']||0)+', 5th '+(sec.by_round['5th round']||0);
  return s;
}

function renderResults(j){
  cards={};
  const R=$('results');
  let top='<div class="dropnote" id="dropNote"></div>';
  if(j.rounds.length===2)
    top+='<div class="roundnote"><b>Both rounds.</b> Hits are tagged '+rtag('4th round')+' or '+rtag('5th round')
      +'. Coverage and counts are shown per round and never added together, because the rounds cover different countries.</div>';
  R.innerHTML=top;
  const openN=j.results.length<=4?j.results.length:2;
  j.results.forEach((b,bi)=>{
    const two=(b.rounds||[]).length===2;
    cards[bi]={round:'all',shown:{},loaded:false};
    const card=document.createElement('article');card.className='termcard';card.id='card'+bi;
    const head=document.createElement('div');head.className='termhead'+(bi<openN?'':' closed');
    head.setAttribute('role','button');head.tabIndex=0;head.setAttribute('aria-expanded',bi<openN);
    let hh='<span class="arrow" aria-hidden="true"></span><span class="termname">'+esc(b.term)+'</span>'
      +'<span class="hitpill'+(b.total===0?' zero':'')+'">'+plural(b.total,unitOf())+'</span>';
    if(two)hh+=b.rounds.map(l=>'<span class="rpill '+RCLS[l]+'">'+RSHORT[l]+' '+(b.by_round[l]||0).toLocaleString()+'</span>').join('');
    if(b.aliases&&b.aliases.length>1)
      hh+='<span class="aliaswrap" title="All of these are treated as the same term">also '+b.aliases.slice(1).map(a=>'<span class="alias">'+esc(a)+'</span>').join('')+'</span>';
    if(b.scope)hh+='<span class="scopetag">only '+esc(b.scope)+'</span>';
    hh+=covHtml(b,bi);
    head.innerHTML=hh;
    card.appendChild(head);
    if(b.warn&&b.warn.length){
      const wb=document.createElement('div');wb.className='warnbar';
      wb.innerHTML=b.warn.map(w=>'<div>'+esc(w)+'</div>').join('');card.appendChild(wb);}
    if(b.forms&&b.forms.length){
      const fb=document.createElement('div');fb.className='formsbar';
      fb.innerHTML='<span class="lbl">Matched in the text as</span>'+b.forms.map(f=>'<span class="form">'+esc(f.form)+'<i>'+f.n+'</i></span>').join('');
      card.appendChild(fb);}
    if(b.rounds&&b.rounds.length){
      const cb=document.createElement('div');cb.className='chipbox';cb.id='chips'+bi;cb.innerHTML=chipsHtml(b);card.appendChild(cb);}
    const body=document.createElement('div');body.className='termbody';body.id='body'+bi;
    if(bi>=openN)body.style.display='none';
    let bh='';
    if(two)bh+='<div class="cardtools">Show <div class="seg mini" id="rs'+bi+'">'
      +'<button class="on" data-v="all" onclick="setCardRound('+bi+',\'all\')">Both rounds</button>'
      +'<button data-v="4th round" onclick="setCardRound('+bi+',\'4th round\')">4th round only</button>'
      +'<button data-v="5th round" onclick="setCardRound('+bi+',\'5th round\')">5th round only</button></div></div>';
    b.sections.forEach((s,si)=>{
      const color=SECS[s.name]||'var(--oth)';
      bh+='<section class="tblsec" style="--c:'+color+'"><div class="secbar"><span class="bar" style="background:'+color+'"></span>'
        +'<span class="name">'+esc(s.name)+'</span><span class="cnt" id="cnt'+bi+'_'+si+'">'+secCount(s,'all',two)+'</span></div>'
        +'<div class="secrows" id="sb'+bi+'_'+si+'"></div></section>';
    });
    body.innerHTML=bh;
    const toggle=()=>{head.classList.toggle('closed');const closed=head.classList.contains('closed');
      head.setAttribute('aria-expanded',!closed);body.style.display=closed?'none':'';if(!closed)ensureLoaded(bi);};
    head.onclick=e=>{if(e.target.closest('button'))return;toggle();};
    head.onkeydown=e=>{if((e.key==='Enter'||e.key===' ')&&e.target===head){e.preventDefault();toggle();}};
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
function unitOf(){return lastMode==='KID'?'KID':lastMode==='CASE'?'case study':'hit';}
function tableShell(two,kind){
  if(kind==='KID')return '<table class="results"><colgroup>'+(two?'<col style="width:56px">':'')
      +'<col style="width:13%"><col style="width:104px"><col style="width:20%"><col style="width:128px"><col><col style="width:68px"></colgroup>'
      +'<thead><tr>'+(two?'<th>Round</th>':'')+'<th>Jurisdiction</th><th>KID</th><th>IO or Recommendation, theme</th><th>Plenary decision</th><th>Where the term appears</th><th><span class="sr">Actions</span></th></tr></thead><tbody></tbody></table>';
  if(kind==='CASE')return '<table class="results"><colgroup>'+(two?'<col style="width:56px">':'')
      +'<col style="width:13%"><col style="width:20%"><col style="width:76px"><col><col style="width:68px"></colgroup>'
      +'<thead><tr>'+(two?'<th>Round</th>':'')+'<th>Jurisdiction</th><th>IO and core issue</th><th>Rating</th><th>Case study</th><th><span class="sr">Actions</span></th></tr></thead><tbody></tbody></table>';
  if(kind==='MER')return '<table class="results"><colgroup>'+(two?'<col style="width:56px">':'')
      +'<col style="width:13%"><col style="width:84px"><col style="width:17%"><col style="width:76px"><col><col style="width:68px"></colgroup>'
      +'<thead><tr>'+(two?'<th>Round</th>':'')+'<th>Jurisdiction</th><th>Measure</th><th>Indicator</th><th>Rating</th><th>Term in context</th><th><span class="sr">Actions</span></th></tr></thead><tbody></tbody></table>';
  return '<table class="results"><colgroup><col style="width:24%"><col style="width:32%"><col><col style="width:68px"></colgroup>'
      +'<thead><tr><th>Where in the file</th><th>Sentence with the term</th><th>Term in context</th><th><span class="sr">Actions</span></th></tr></thead><tbody></tbody></table>';
}
function actions(id){
  const d=excluded.has(id);
  return '<td class="acts"><button class="actbtn" title="Copy the quote with its citation" aria-label="Copy quote" onclick="copyCite(\''+id+'\',this)">⧉</button>'
    +'<button class="exbtn" title="'+(d?'Put this hit back':'Remove this hit from the export')+'" aria-label="'+(d?'Put back':'Remove')+'" onclick="toggleDrop(\''+id+'\',this)">'+(d?'↺':'✕')+'</button></td>';
}
function readTools(x){
  return '<div class="rowtools"><button class="readbtn" aria-expanded="false" onclick="toggleDoc(this,\''+x._id+'\')">Read paragraph</button>'
    +fmtTags(x.html)+'</div>';
}
function decisionPill(o){const c=o==='Upgrade'?'up':o==='Downgrade'?'down':'same';
  return '<span class="decision '+c+'">'+esc(o||'Not recorded')+'</span>';}
function kidShort(k){return String(k||'').replace(/\s*\([^)]*\)\s*$/,'');}
function rowHtml(x,two,kind){
  const cls=excluded.has(x._id)?' dropped':'';
  if(kind==='KID'){
    return '<tr class="hit'+cls+'" data-id="'+x._id+'">'+(two?'<td>'+rtag(x.Round)+'</td>':'')
      +'<td class="jur">'+esc(x.Jurisdiction)+'</td>'
      +'<td><b>'+esc(kidShort(x.KID))+'</b><span class="smallmeta">'+esc(x.PlenaryLabel||'')+' Plenary</span></td>'
      +'<td>'+esc(x.Indicator)+'<span class="smallmeta">'+esc(x.Theme||'')+'</span></td>'
      +'<td>'+decisionPill(x.Outcome)+(x.Rating&&x.Rating!=='NA'?'<span class="smallmeta">Final rating '+pill(x.Rating)+'</span>':'')+'</td>'
      +'<td class="ctx">'+hl(x.context.replace(/^\[[^\]]*\]\s*/,''),x.matched,x.variants)
      +'<div class="rowtools"><button class="readbtn" aria-expanded="false" onclick="toggleDoc(this,\''+x._id+'\')">Open KID</button>'
      +'<span class="smallmeta" style="display:inline;margin:0 2px 0 4px">Found in</span>'
      +(x.matched_parts||[]).map(n=>'<span class="parttag">'+esc(n)+'</span>').join('')+'</div></td>'+actions(x._id)+'</tr>';
  }
  if(kind==='CASE'){
    const io=(x.MEASURE||'').replace(/^IO(\d+).*/,'IO.$1');
    return '<tr class="hit'+cls+'" data-id="'+x._id+'">'+(two?'<td>'+rtag(x.Round)+'</td>':'')
      +'<td class="jur">'+esc(x.Jurisdiction)+'</td>'
      +'<td><b>'+esc(io)+(x.CoreIssue?', core issue '+esc(x.CoreIssue):'')+'</b><span class="smallmeta">'+esc(x.Indicator)+'</span></td>'
      +'<td>'+(x.Rating==='NA'?'<span class="pill na">NA</span>':pill(x.Rating))+'</td>'
      +'<td class="ctx"><b>'+hl(x.BoxTitle||'',x.matched,x.variants)+'</b>'
      +(x.BoxKind&&x.BoxKind!=='Case study'?' <span class="fmt">'+esc(x.BoxKind)+'</span>':'')
      +'<div>'+hl(x.matched?x.context:(x.context||'').replace(x.BoxTitle||'','').trim(),x.matched,x.variants)+'</div>'
      +'<div class="rowtools"><button class="readbtn" aria-expanded="false" onclick="toggleDoc(this,\''+x._id+'\')">Open case study</button></div></td>'
      +actions(x._id)+'</tr>';
  }
  if(kind==='MER'){
    const crit=(x['Criteria rating']&&x['Criteria rating']!=='NA')?'<br>'+pill(x['Criteria rating']):'';
    return '<tr class="hit'+cls+'" data-id="'+x._id+'">'+(two?'<td>'+rtag(x.Round)+'</td>':'')
      +'<td class="jur">'+esc(x.Jurisdiction)+'</td><td class="meas">'+esc(x.MEASURE)+'</td><td>'+esc(x.Indicator)+'</td>'
      +'<td>'+(x.Rating==='NA'?'<span class="pill na">NA</span>':pill(x.Rating))+crit+'</td>'
      +'<td class="ctx">'+hl(x.context,x.matched,x.variants)+readTools(x)+'</td>'+actions(x._id)+'</tr>';
  }
  return '<tr class="hit'+cls+'" data-id="'+x._id+'"><td class="jur">'+esc(x.Location)
    +(x.Details?'<div style="font-weight:400;color:var(--muted);margin-top:3px;font-size:12px">'+esc(String(x.Details).slice(0,220))+'</div>':'')
    +'</td><td class="ctx">'+hl(x['Sentence with term'],x.matched,x.variants)+'</td>'
    +'<td class="ctx">'+hl(x.context,x.matched,x.variants)+readTools(x)+'</td>'+actions(x._id)+'</tr>';
}
function kidNode(x,color){
  const art=document.createElement('div');art.className='doc';art.style.setProperty('--c',color);
  const top=document.createElement('div');top.className='kidtop';
  top.innerHTML='<div class="kidtitle">'+esc(kidShort(x.KID))+'<small>'+esc(x.Jurisdiction)+', '+esc(x.Round)+'</small></div>'
    +'<span class="sp"></span><button class="linkbtn" onclick="copyCite(\''+x._id+'\',this)">Copy quote</button>'
    +'<dl class="kidmeta">'
    +'<div><dt>IO or Recommendation</dt><dd>'+esc(x.Indicator)+'</dd></div>'
    +'<div><dt>Theme</dt><dd>'+esc(x.Theme||'')+'</dd></div>'
    +'<div><dt>Plenary</dt><dd>'+esc(x.PlenaryLabel||'')+'</dd></div>'
    +'<div><dt>Decision</dt><dd>'+decisionPill(x.Outcome)+(x.Rating&&x.Rating!=='NA'?' &nbsp;final rating '+pill(x.Rating):'')+'</dd></div>'
    +'</dl>';
  const ol=document.createElement('ol');ol.className='steps';
  (x.parts||[]).forEach((pt,i)=>{
    const li=document.createElement('li');if(pt.hit)li.className='hit';
    li.innerHTML='<span class="stepdot">'+(i+1)+'</span><div class="steplabel">'+esc(pt.part)
      +(pt.hit?'<span class="found">term found here</span>':'')+'</div>';
    const body=document.createElement('div');body.className='stepbody docbody';
    if(pt.html){body.innerHTML=pt.html;shapeDoc(body);if(pt.hit)markDom(body,x.matched,x.variants);}
    else{body.classList.add('emptypart');body.textContent='Nothing recorded.';}
    li.appendChild(body);ol.appendChild(li);});
  const foot=document.createElement('div');foot.className='hint';foot.style.marginTop='14px';
  foot.textContent='Cite as: '+(x.ref||'');
  art.appendChild(top);art.appendChild(ol);art.appendChild(foot);return art;
}
function docNode(x,color){
  if(x.parts)return kidNode(x,color);
  const art=document.createElement('div');art.className='doc';art.style.setProperty('--c',color);
  const head=document.createElement('div');head.className='dochead';
  head.innerHTML='<span class="cite">'+esc(x.ref||'')+'</span>'+(x.Indicator?'<span>'+esc(x.Indicator)+'</span>':'')
    +'<span class="sp"></span><button class="linkbtn" onclick="copyCite(\''+x._id+'\',this)">Copy quote</button>';
  const body=document.createElement('div');body.className='docbody';
  if(x.html){body.innerHTML=x.html;shapeDoc(body);markDom(body,x.matched,x.variants);}
  else{body.innerHTML='<p>'+hl(x.OBS_VALUE||x['Full text']||'',x.matched,x.variants)+'</p>';}
  art.appendChild(head);art.appendChild(body);return art;
}
function toggleDoc(btn,id){
  const tr=btn.closest('tr');const nx=tr.nextElementSibling;
  const isKid=!!(rowCache[id]&&rowCache[id].parts);const what=isKid?'KID':lastMode==='CASE'?'case study':'paragraph';
  if(nx&&nx.classList.contains('docrow')&&nx.dataset.for===id){nx.remove();btn.setAttribute('aria-expanded','false');btn.textContent=(isKid||lastMode==='CASE'?'Open ':'Read ')+what;return;}
  const x=rowCache[id];if(!x)return;
  const row=document.createElement('tr');row.className='docrow';row.dataset.for=id;
  const td=document.createElement('td');td.colSpan=tr.children.length;
  const sec=tr.closest('.tblsec');const color=sec?getComputedStyle(sec).getPropertyValue('--c')||'var(--accent)':'var(--accent)';
  td.appendChild(docNode(x,color));row.appendChild(td);tr.after(row);
  btn.setAttribute('aria-expanded','true');btn.textContent='Close '+what;
}

async function loadSection(bi,si,append,limit){
  const st=cards[bi];const b=lastResults[bi];const sec=b.sections[si];
  const two=(b.rounds||[]).length===2;const isMer=lastMode!=='OTHER';const kind=lastMode;
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
    box.innerHTML=tableShell(two&&reqRound==='all',kind);
  }
  box.querySelector('tbody').insertAdjacentHTML('beforeend',j.rows.map(x=>rowHtml(x,two&&reqRound==='all',kind)).join(''));
  st.shown[si]=off+j.rows.length;
  const old=box.querySelector('.morebar');if(old)old.remove();
  const left=j.total-st.shown[si];
  if(left>0){
    box.insertAdjacentHTML('beforeend','<div class="morebar"><button class="morebtn" onclick="loadSection('+bi+','+si+',true,'+PAGE+')">Show '
      +Math.min(PAGE,left)+' more ('+st.shown[si]+' of '+j.total+' shown)</button>'
      +'<button class="morebtn all" onclick="loadSection('+bi+','+si+',true,500)">Show '+(left>500?'500 more':'all')+'</button></div>');
  }
}

function updateDropNote(){const note=$('dropNote');if(!note)return;
  note.style.display=excluded.size?'block':'none';
  note.textContent=plural(excluded.size,'hit')+' removed. They will be left out of the export and its statistics.';}
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
  if(ok){const old=btn.textContent;btn.textContent=btn.classList.contains('linkbtn')?'Copied':'✓';btn.classList.add('done');
    setTimeout(()=>{btn.textContent=old;btn.classList.remove('done');},1300);}
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
      +(sent?'':'\n\nTip: "Quotes: sentence only" makes both files several times smaller.')+'\n\nBuild anyway?'))return;
  spin(true);
  const ss=$('searchStatus');ss.classList.remove('err');
  ss.textContent='Building the '+(fmt==='docx'?'Word report':fmt==='both'?'files':'spreadsheet')+
    (excluded.size?' (leaving out '+plural(excluded.size,'removed hit')+')':'')+'…';
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
  lastPayload=null;$('results').innerHTML=welcomeHtml();
  $('searchStatus').textContent='';$('interp').style.display='none';$('interp').innerHTML='';
  ['btnExportX','btnExportW','btnExportB'].forEach(id=>$(id).disabled=true);
  lastPayload=null;excluded.clear();clearFilters();countTerms();fileNamePreview();}

function histGet(){try{return JSON.parse(localStorage.getItem('fatf_hist')||'[]');}catch(e){return [];}}
function histSet(h){try{localStorage.setItem('fatf_hist',JSON.stringify(h));}catch(e){}}
function saveHist(p){
  let h=histGet();
  h.unshift({t:p.terms,mode:p.data_mode,round:p.round,sm:p.search_mode,when:new Date().toLocaleString()});
  histSet(h.slice(0,25));renderHist();}
function histSrc(x){if(x.mode==='OTHER')return 'Other files';
  const r=x.round==='both'?'both rounds':(x.round==='4'?'4th round':(x.round==='5'?'5th round':''));
  return (x.mode==='KID'?'Key issues, ':x.mode==='CASE'?'Case studies, ':'MER, ')+r;}
function renderHist(){
  const h=histGet();const el=$('hist');
  if(!h.length){el.innerHTML='<div class="empty">No searches yet.</div>';return;}
  el.innerHTML=h.map((x,i)=>'<a onclick="reuse('+i+')" title="'+esc(x.when)+'"><small>'+esc(histSrc(x))+'</small>'+esc(x.t.join(', ')||'all, by filters')+'</a>').join('');}
async function reuse(i){const x=histGet()[i];if(!x)return;
  $('terms').value=x.t.join('\n');countTerms();fileNamePreview();
  if(x.sm)setSM(x.sm);
  const needMode=x.mode||'MER';
  if(needMode!==dataMode){dataMode=needMode;applyModeUI();loaded=false;}
  if(needMode!=='OTHER'&&x.round&&x.round!==round){round=x.round;segSet('roundSeg',round);roundHint();loaded=false;}
  if(!loaded)await loadData();
  window.scrollTo({top:0});}
function countTerms(){
  const n=$('terms').value.split('\n').map(t=>t.trim()).filter(Boolean).length;
  $('termCount').textContent=n?plural(n,'term'):(dataMode==='CASE'?'No words: Search lists every case study matching the filters':'No terms yet');}

$('terms').addEventListener('input',()=>{countTerms();fileNamePreview();});
$('terms').addEventListener('keydown',e=>{if(e.key==='Enter'&&(e.ctrlKey||e.metaKey)){e.preventDefault();doSearch();}});
$('reportName').addEventListener('input',fileNamePreview);
$('dlgSearch').addEventListener('input',()=>{const f=F[curTab];f.q=$('dlgSearch').value;
  const rs=document.querySelector('.fsearch[data-f="'+curTab+'"]');if(rs)rs.value=f.q;refreshFacet(curTab);});
$('dlgOnly').addEventListener('change',()=>refreshFacet(curTab));
$('fdlg').addEventListener('click',e=>{if(e.target===$('fdlg'))closeFilters();});   // click on the backdrop
restorePaths();
bindList($('dlgList'));
applyModeUI();
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
    print("  New: paragraphs shown with their tables and boxes | new filter panel")
    print(f"  Data folder: {BASE_DIR}")
    print("  Keep this window open while using the tool. Press Ctrl+C to stop.\n")
    app.run(host="127.0.0.1", port=PORT, debug=False, threaded=True)
