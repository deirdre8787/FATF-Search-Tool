"""
Check how case study boxes are marked up in the MER effectiveness text.

Run from the tool's folder:
    python tools/box_check.py            (4th round)
    python tools/box_check.py 5          (5th round)

It prints what the tool finds, and saves exports/box_sample.txt with a few
raw paragraphs that mention a box, exactly as they are in the CSV. That
file is what to send if the boxes are not being picked up.
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import pandas as pd   # noqa: E402
import app            # noqa: E402

rn = sys.argv[1] if len(sys.argv) > 1 else "4"
folder = app.PATHS["round_" + rn]
path = app._find_csv(folder, ["rc_io", "kfra"], "effectiveness / risk and context")
print(f"Reading {path}")
raw = pd.read_csv(path, encoding="utf-8-sig", dtype=str,
                  usecols=["Jurisdiction", "MEASURE", "OBS_VALUE"]).fillna("")
eff = raw[raw["MEASURE"].str.startswith("IO")]
v = eff["OBS_VALUE"]
print(f"Effectiveness paragraphs:            {len(eff):>7,}")
checks = [
    ("with any HTML tag", r"<[a-zA-Z]"),
    ("with a <table>", r"(?i)<table"),
    ("with a class= attribute", r"(?i)class="),
    ("with a border style", r"(?i)border"),
    ("with a background style", r"(?i)background"),
    ("mentioning 'Box 1.2'-style titles", r"Box\s+\d+\.\d+"),
    ("mentioning 'case study/example'", r"(?i)case\s+(stud|example)"),
]
for label, pat in checks:
    n = int(v.str.contains(pat, regex=True).sum())
    print("  " + (label + ":").ljust(36) + "{:>7,}".format(n))

cases = app.load_cases_round(rn)
print(f"\nBoxes the tool finds: {len(cases):,}")
for k, n in cases["BoxKind"].value_counts().items():
    print(f"  {k}: {n:,}")

sample = eff[v.str.contains(r"Box\s+\d+\.\d+", regex=True)].head(3)
if sample.empty:
    sample = eff[v.str.contains(r"(?i)case\s+(stud|example)", regex=True)].head(3)
out_dir = os.path.join(os.path.dirname(HERE), "exports")
os.makedirs(out_dir, exist_ok=True)
out = os.path.join(out_dir, "box_sample.txt")
with open(out, "w", encoding="utf-8") as fh:
    for _, r in sample.iterrows():
        fh.write(f"===== {r['Jurisdiction']} | {r['MEASURE']}\n{r['OBS_VALUE'][:6000]}\n\n")
print(f"\nSaved {len(sample)} raw sample paragraph(s) to {out}")
