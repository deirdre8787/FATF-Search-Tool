# FATF Search Tool

**Version 9 (15 October 2026)**

A keyword-search tool for FATF Mutual Evaluation Reports (MERs) and Key Issue Documents (KIDs). It searches the 4th round, the 5th round, or both together, and can also search your own Word, PDF and Excel documents. Results export to Excel and to a Word report, with a citation for every quote.

The tool runs on your own computer. Nothing is uploaded anywhere: it opens as a page in your browser, served from your own computer.

---

## Contents

- [What it can do](#what-it-can-do)
- [Quick start](#quick-start)
- [Getting the MER data](#getting-the-mer-data)
- [Writing search terms](#writing-search-terms)
- [Exports](#exports)
- [Searching other documents](#searching-other-documents)
- [Command-line version](#command-line-version)
- [Reading the results correctly](#reading-the-results-correctly)
- [Troubleshooting](#troubleshooting)
- [Files in this repository](#files-in-this-repository)

---

## What it can do

- **Search the Key Issue Documents.** Each KID records an issue debated at Plenary in five parts: the key issue, the ECG discussion, the ECG proposal, the Plenary discussion and the Plenary decision. A hit is one KID; it opens to the full record as a five-step timeline with the parts that matched highlighted. Filter by jurisdiction, IO or Recommendation, theme, part of the KID, Plenary decision (upgrade, downgrade, text amendment), Plenary meeting and final rating. Exports have one row per KID with each part in its own column (Excel) and the whole KID under each country (Word).
- **Case studies.** The case study boxes in the effectiveness chapters of the MERs, both rounds, each labelled with its round, country, IO, core issue and rating. Type words to narrow them down, or leave the search box empty and press Search to list every case study that matches the filters, ready to export in full. Boxes are found from the report's formatting (single-cell or shaded tables, bordered blocks); boxes written as plain paragraphs starting "Box 3.1 ..." are picked up too. A box counts as a case study when its title mentions a case or an example; both kinds are shown by default and the "Type of box" filter separates them. If a round shows no boxes, run `python tools/box_check.py 4` (or `5`): it reports how that round's text is marked up and saves a few raw sample paragraphs to `exports/box_sample.txt`.
- **Search every MER paragraph** in the OECD Data Explorer datasets: Effectiveness (IO.1 to IO.11), Risk and Context (including the executive summary) and Technical Compliance (R.1 to R.40, criterion by criterion).
- **Search the 4th round, the 5th round, or both.** In "Both rounds", every hit is tagged 4th or 5th round, and counts and coverage are always shown per round, never added together. The rounds cover different countries and use different methodologies (2013 and 2022).
- **Read each hit as it appears in the report.** "Read paragraph" opens the full paragraph with its original formatting: tables, the shaded case-study boxes, lists, bold text and footnote marks, with the matched words highlighted. Hits that contain a table, box or list are labelled.
- **Filter** by jurisdiction, table, section (for example IO.6 or C.20.1), rating and criterion rating. Each filter is a checklist with search, *All* and *None*, and a tick box on every group, so you can select a whole section (all of IO.6, all of R.20, all FATF members) in one click. **Expand** opens the filters in a large window (with a full-screen option) for long lists. Filters also set the denominator for "x of y jurisdictions".
- **See what actually matched.** Each term shows the spellings found in the text and how many times each appears. This is the quickest way to spot false hits, which you can remove before exporting.
- **Compare rounds.** For countries with both a 4th and a 5th round report, the tool shows whether a term appears in one round, both or neither.
- **Export** to Excel (filterable tables) and to a Word report (quotes grouped by round and country, each with a citation such as *Austria, 5th round MER, IO.4 (rated ME)*).
- **Search your own files** (Word, PDF, Excel, CSV, text), with the exact location of each hit.

## Quick start

**You need:** Windows and Python 3.9 or later. [Anaconda](https://www.anaconda.com/download) is the easiest option, because it already includes pandas and openpyxl.

1. Download this repository: green **Code** button, then **Download ZIP**. Unzip it, for example to `Documents\Search_tool`.
2. Put the MER CSV files in the `source files` folders (see [Getting the MER data](#getting-the-mer-data)).
3. Double-click **`run_search_tool.bat`**. The first time, it installs Flask, which the tool needs.
4. Your browser opens at **http://127.0.0.1:5177**. Keep the black window open while you use the tool, and close it to stop.

The header should say **v9 (15 Oct 2026)**. If it shows an older version, close every black window, run the .bat again and press Ctrl+F5 in the browser.

On Mac or Linux, or without the .bat file:

```bash
pip install -r requirements.txt
python app.py
```

**Optional extras:** `rapidfuzz` speeds up fuzzy mode, and `pypdf` is needed to search PDFs. Both are in `requirements.txt`.

### Staying up to date automatically

Instead of downloading a ZIP each time, install the tool once with git and let the launcher update it:

1. Install [Git for Windows](https://git-scm.com/download/win) (all the default options are fine).
2. Open a Command Prompt in your Documents folder and run:
   ```
   git clone https://github.com/deirdre8787/fatf-search-tool.git Search_tool
   ```
3. Put the MER CSVs in `Search_tool\source files\...` as below, and double-click `run_search_tool.bat`.

Every time you start it, `run_search_tool.bat` fetches the newest version from GitHub (`git pull`) before opening the tool. Your CSVs, your own documents and your exports are never uploaded or overwritten, because git ignores them. If you are offline, it starts the copy you already have.

Do not edit the tool's own files in that folder: a local edit stops the update. (To undo local edits: `git checkout -- .`)

**Keeping the data somewhere else.** If you prefer to keep the CSVs in another folder (for example a shared drive), click *Change source folders* in the tool and paste the folder paths. The tool remembers them on that computer.

## Getting the MER data

The MER data is not included in this repository. Download it from OECD Data Explorer, which has two CSV files for each round:

| Round | File name contains | Holds |
| --- | --- | --- |
| 4th | `KFRA` / `RC_IO` | Effectiveness and Risk and Context |
| 4th | `TCOMP` / `TECH_COMP` | Technical Compliance |
| 5th | `KFRA_5` / `RC_IO` | Effectiveness and Risk and Context |
| 5th | `TCOMP_5` / `TECH_COMP` | Technical Compliance |

Put them here:

```
source files/
  MER/
    round_4/   <- the two 4th round CSVs
    round_5/   <- the two 5th round CSVs
  KID/         <- the Key Issue Documents CSV (name contains KID)
  other/       <- your own Word / PDF / Excel / CSV / text files
```

The tool finds the files by name, so a newer download only needs dropping in to replace the old one. The first load of the 4th round takes 20 to 60 seconds (and once more after updating to a new version). After that the tool keeps a cache in `%LOCALAPPDATA%\MER_Search_cache`, so later loads take about a second. You can delete the cache at any time and it will rebuild itself.

For the Key Issue Documents, download the **DF_KID** table from OECD Data Explorer (the full table, with labels) and put the CSV in `source files/KID/`. One file holds both rounds.

## Writing search terms

Enter one term per line, as many as you like. Click **Check my terms** to see exactly how each line will be searched, along with any warnings.

| You type | What it finds |
| --- | --- |
| `beneficial ownership` | The words in this order. Case, spacing, hyphens, punctuation and accents are ignored |
| `ship` | Words that *start* with "ship" (ships, shipping), never "membership" |
| `DPMS~dealers in precious metals and stones` | Either name. `~` joins names for the same thing into one result |
| `JIG (Joint Intelligence Group)` | JIG, the full name, or either bracketed form |
| `SAMLIT (South Africa)` | SAMLIT, in South Africa's report only |
| `ASC` | ASC as a whole word, never "ascertain". Tick **capitals only** to skip lower-case "asc" |
| `JMLIT+` | Only "JMLIT+", not plain JMLIT |
| `Section 314` | Exactly 314, not 3140 |

**Search modes:**

- **Strict keyword** (the default): words must start with the term.
- **Exact phrase**: whole words only.
- **Fuzzy**: tolerates typos. It is slow on the full 4th round, so filter by jurisdiction first.

**Tips:**

- Spell out short abbreviations that appear everywhere. Searching `TF` or `PF` alone matches almost every paragraph, so use `proliferation financing` instead.
- Tick **capitals only** for acronym lists (FTZ, TBML, MVTS, NPO, PPP).
- Use the Indicator filter to target sections. For example, searching `fraud` across all MERs returns about 3,100 paragraphs, but far fewer within IO.6 and R.20.

## Exports

Choose a report name and the quote style (*full paragraph* or *sentence only*), then export to Excel, Word or both. Files are saved to `exports/` as `YYYY-MM-DD_<report name>_FATF Search tool.xlsx` / `.docx`.

**Excel.** Every sheet is a real Excel table with a filter on each column, including Round.

| Sheet | Contents |
| --- | --- |
| Summary | One row per term per round: hits by table, jurisdictions mentioned and in scope, coverage %, matched spellings, warnings. Click a term to jump to its sheet |
| Country matrix | Every jurisdiction against every term, with the number of paragraphs mentioning it (shaded) |
| Round comparison | Countries with both a 4th and 5th round report: hits in each round and the change |
| One sheet per term | Round, Table, Reference (for citing), Jurisdiction, MEASURE, Indicator, Rating, Quote (term in bold), Term in context |
| All hits | Every hit of every term in one flat table, ready for pivot tables |

**Word.** A cover page, then for each term: a statistics table with a column per round, a comparison table for countries assessed in both rounds, the wording matched, and every quote grouped by round and country. Each quote has a coloured round tag and its citation. The report has page numbers, and Word's navigation pane works.

**Methodology note.** `Methodology_note_TEMPLATE.docx` is a plain-English note for colleagues who receive your exports: what was searched, how, and how to read the files. Fill in the bracketed fields, save it under a new name and send it with the exports.

## Searching other documents

Choose **Other files** on the left, then drop `.docx`, `.pdf`, `.xlsx` / `.xls`, `.csv` or `.txt` files into `source files/other/`. Each hit shows where it was found:

- **Word:** heading trail plus paragraph number, or table / row / column
- **Excel:** sheet, cell and column heading, with the whole row as context
- **PDF:** page and paragraph

Old `.doc` files need saving as `.docx` first. Scanned PDFs with no text layer cannot be searched.

## Command-line version

`tools/mer_search.py` runs the same search engine without the browser and writes an Excel workbook (Summary, one sheet per term, All hits). It is useful for scripted or repeated searches.

```bash
python tools/mer_search.py \
    --r4 "source files/MER/round_4" --r5 "source files/MER/round_5" \
    --round both --terms terms.txt --out results.xlsx \
    --caps-only --quotes sentence
```

Options: `--round 4|5|both`, `--mode strict|exact|fuzzy`, `--jurisdictions "Bahrain,Qatar"`, `--tables "Effectiveness,Risk and Context"`, `--quotes paragraph|sentence`.

Its hit counts match the app exactly. The Word report, Country matrix, Round comparison and Other files mode are only in the app.

## Reading the results correctly

- **No hit does not mean the thing is absent.** It means the report does not use that wording.
- **Never add the rounds together.** The 5th round covers only a few countries so far, so always give the denominator ("3 of 7 jurisdictions").
- **Counts can include repeats.** Key findings often reappear word for word in the executive summary.
- **MERs are snapshots.** Each report describes the country at the time of its on-site visit, so cite the assessment year with each quote.
- **Read the hits.** Check "Matched in the text as" for false hits before using any figure.

## Troubleshooting

| Problem | Fix |
| --- | --- |
| Old version or layout in the browser | Close every black window, run the .bat again, press Ctrl+F5 |
| "No … CSV found" | The file name must contain `RC_IO`/`KFRA` or `TECH_COMP`/`TCOMP`, and it must be in the right round folder |
| "Could not save" | A file with that name is open in Excel or Word. Close it and export again |
| "This search has expired" | Search again. The tool keeps the last 6 searches in memory |
| Port 5177 already in use | Another copy of the tool is running. Close its window |
| "Python was not found" | Install Anaconda, or open Anaconda Prompt and run `python app.py` from the tool's folder |

The full user guide, including a list of what changed in each version, is in [README.txt](README.txt).

## Files in this repository

| File | What it is |
| --- | --- |
| `app.py` | The whole application: Flask server, search engine and web interface |
| `run_search_tool.bat` | Windows launcher: finds Python, installs Flask, starts the tool |
| `requirements.txt` | Python packages |
| `README.txt` | Full user guide |
| `Methodology_note_TEMPLATE.docx` | Methodology note to send with exports |
| `tools/mer_search.py` | Command-line version of the search |
| `source files/` | Where the MER CSVs and your own documents go. Not uploaded to GitHub |
| `exports/` | Where exports are saved. Not uploaded to GitHub |

`.gitignore` keeps your data and exports out of the repository, so only the tool itself is shared.
