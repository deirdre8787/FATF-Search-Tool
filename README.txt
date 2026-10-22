FATF SEARCH TOOL  -  v10 (22 Oct 2026)
=====================================
Search FATF Mutual Evaluation Reports (4th round, 5th round, or both
together) and other documents, then export the results to Excel and Word.

CHECK YOU ARE ON v10: the header in the browser says "v10 (22 Oct 2026)".
If it doesn't, close every black console window, run the .bat again and
press Ctrl+F5 in the browser.


WHAT IS NEW IN v10
------------------
* 4TH ROUND CASE STUDIES FROM THE EXTRACTED BOXES FILE. Most 4th round
  boxes are missing from the Data Explorer text, so they now come from
  MER_boxes_4th_round_....csv (boxes extracted from the published MER
  PDFs). Put it in "source files\MER\round_4" next to the other two
  CSVs; the tool finds it by "box" in the file name. Case studies then
  shows the 4th round from that file and the 5th round from the text,
  together. Each box shows its IO, rating, the page numbers in the MER
  and any note from the extraction (for example "check where it ends").
  Boxes outside the IO chapters are under "Other chapters".


WHAT IS NEW IN v9
-----------------
* PARAGRAPHS KEEP THEIR FORMATTING. The MER data arrives as HTML. The tool
  used to throw the formatting away; it now keeps a safe copy of it.
  "Read paragraph" under each hit opens the full paragraph as it looks in
  the report: tables as tables, the shaded "Box" case studies as boxes,
  lists, bold, italics and footnote marks, with the matched words
  highlighted. Hits containing a table, box or list are labelled.
  The search itself still runs on the plain text, so hit counts are
  exactly the same as in v8. Exports are unchanged (plain-text quotes).
  HTML in your own CSV / Excel cells (Other files) is shown the same way.

* KEY ISSUE DOCUMENTS. A third source next to the mutual evaluations
  and your own files. Put the DF_KID CSV from OECD Data Explorer in
  "source files\KID". Each hit is one KID; "Open KID" shows the whole
  record as five steps (key issue, ECG discussion, ECG proposal, Plenary
  discussion, Plenary decision) with the parts that mention the term
  highlighted, and tracked text amendments shown as in the original.
  Extra filters: theme, part of the KID, Plenary decision (upgrade /
  downgrade / text amendment), Plenary meeting and final rating.
  Excel: one row per KID with each part in its own column, term in bold.
  Word: the KIDs grouped by round and country, each with all five parts.

* CASE STUDIES. A fourth source: the case study boxes in the
  effectiveness chapters of the MERs (4th and 5th round). Each one shows
  its round, country, IO, core issue and rating. Leave the search box
  empty and press Search to list every case study that matches your
  filters, or type words to narrow them down. Excel has one row per case
  study with its full text; Word groups them by round and country.
  A box counts as a case study when its title mentions a case or an
  example; other boxes can be included with the "Type of box" filter.

* NEW FILTER PANEL. Each filter (jurisdictions, sections, ratings,
  criterion ratings) is now a checklist instead of a Ctrl+click list:
    - a search box, plus All / None for whatever is shown;
    - a tick box on every group, to select a whole section at once:
      all FATF members, all of Effectiveness, all of IO.6, all of R.20,
      all effectiveness ratings, and so on;
    - "Expand" opens the filters in a large window, with a full-screen
      option, Invert, "show selected only" and quick picks such as
      "Assessed in both rounds";
    - "Search again with these filters" re-runs the last search.
  Everything starts ticked (= no filter). If you untick everything in a
  filter, the tool asks you to tick at least one before searching.

* NEW LOOK, with a dark mode that follows your computer's setting until
  you choose one.

* The first load of each round after updating takes the usual 20 to 60
  seconds once, because the cache is rebuilt to hold the formatting.


WHAT IS NEW IN v8
-----------------
* SEARCH BOTH ROUNDS AT ONCE. The round switch now has three options:
  4th round | 5th round | Both rounds.
  In "Both rounds":
    - every hit carries a round label (purple "4th", amber "5th");
    - counts and jurisdiction coverage are ALWAYS shown per round, never
      added together - the rounds cover different countries (199 vs 7 so
      far) under different methodologies (2013 vs 2022), so a combined
      percentage would mislead;
    - each result card has a "Show: Both rounds / 4th only / 5th only"
      switch, and hits from the same country sit next to each other so
      you can compare a country's 4th and 5th round wording directly;
    - "Show jurisdictions" also lists the countries assessed in BOTH
      rounds and whether the term appears in the 4th, the 5th, both or
      neither.

* EXCEL EXPORT, REBUILT
    Summary           one row per term PER ROUND (filter the Round column).
                      Click a term to jump to its sheet.
    Country matrix    every jurisdiction x every term: how many paragraphs
                      mention it (colour-shaded; empty = not mentioned).
    Round comparison  (both rounds only) countries with both a 4th and a
                      5th round report: hits in each round and the pattern.
    One sheet per term  one table with a Round column and a Table column
                      (Effectiveness / Risk and Context / Technical
                      Compliance), a citation column, the quote and the
                      term in context, term in bold.
    All hits          every hit for every term in one flat table - ready
                      for pivot tables.
  Every table is a real Excel table: every column, including Round, has
  its own filter arrow. (In v7 only the first block of each sheet did.)

* WORD REPORT, ROUND BY ROUND
  Per term: a statistics table with one column per round, the "assessed
  in both rounds" comparison, then the mentions - first the 4th round,
  then the 5th, grouped by country. Every quote starts with a coloured
  4TH ROUND / 5TH ROUND tag and its citation. Page numbers in the footer.

* QUOTES IN EXPORTS: full paragraph, or sentence only.
  Choose next to the export buttons. "Sentence only" makes both the Excel
  and Word files several times smaller - use it for big searches. The
  tool warns you before building a very large file.

* FASTER
    - searches run about 2x faster than v7 (same results - checked);
    - after the first ever load, both rounds load in about a second, even
      after a restart (a cache is kept in your local AppData folder, not
      in Documents or OneDrive; delete it any time, it rebuilds itself);
    - exports reuse the search already on screen instead of re-running it;
    - a progress bar shows which term is being searched.

* EASIER TO WORK WITH RESULTS
    - no more 200-row limit: "Show 25 more" / "Show all" on every table,
      so every hit can be read and removed before exporting;
    - copy button on every hit: copies the sentence with its citation, e.g.
      "...all DNFBPs have a licensing requirement." (Albania, 4th round
      MER, IO.3 (rated ME))
    - Search history remembers the round and mode, and restores them;
    - Ctrl+Enter in the search box runs the search;
    - "Open export folder" link; the default round loads on start-up.

* BUGS FIXED FROM v7
    - Excel filter arrows only worked on the first block of each sheet.
    - "Change source folders" was ignored when clicking Load data.
    - The export folder box showed doubled backslashes.
    - Text starting with "=" could be turned into an Excel formula.
    - Hits beyond the first 200 per table could not be seen or removed.
    - Word tables had a formatting error (Word repaired it silently).
    - The "Check my terms" table stayed on screen after a new search.
    - Hits in two different "other" files with the same location could
      be confused with each other when removing one.


SETUP (one time)
----------------
1. Unzip to:  C:\Users\LUO_D\Documents\Search_tool
   (If you already have a Search_tool folder, just REPLACE app.py,
    README.txt and run_search_tool.bat - your data and exports stay put.)

     Search_tool\app.py
     Search_tool\run_search_tool.bat
     Search_tool\README.txt
     Search_tool\Methodology_note_TEMPLATE.docx
     Search_tool\exports\                      <- exports land here
     Search_tool\source files\MER\round_4\    <- the 2 round-4 CSVs
     Search_tool\source files\MER\round_5\    <- the 2 round-5 CSVs
     Search_tool\source files\other\          <- docx / pdf / xlsx / csv / txt

   The tool now looks for its data next to app.py, so the folder can live
   anywhere (it falls back to the path above if needed).

2. Packages: pandas and openpyxl come with Anaconda. The .bat installs
   Flask the first time. Optional: rapidfuzz (faster fuzzy mode) and pypdf
   (needed only to search PDFs in Other files mode):
       pip install rapidfuzz pypdf


RUN
---
Double-click run_search_tool.bat. The browser opens at http://127.0.0.1:5177
and the 5th round loads automatically. Keep the black window open.
The very first load of the 4th round takes 20-60 seconds; after that it is
about a second, even after a restart.


USE
---
1. Pick MER (4th / 5th / Both rounds) or Other files on the left.
2. Type search terms, one per line, no limit.
     - Case, spaces, punctuation and accents are ignored.
     - "~" joins names for the same thing:  FMLIT~HC fraud and ML intelligence taskforce
     - Brackets are read intelligently:
         JIG (Joint Intelligence Group) -> JIG, the full name and both
                                           bracketed forms all count
         SAMLIT (South Africa)          -> searches SAMLIT, South Africa only
     - ALL-CAPS acronyms match whole words only (ASC never matches
       "ascertain"); tick "capitals only" to ignore lower-case "asc".
3. Click "Check my terms" to see exactly how each line will be searched
   and any warnings (everyday words, very long phrases, duplicates).
4. Optional filters: jurisdiction, table, indicator, ratings.
5. Search. Each term gets a card: hits per round, coverage per round,
   the exact wording found, and the hits table (25 at a time).
   Use X to remove a hit you can see is irrelevant (it is left out of the
   exports and their statistics); the arrow puts it back.
6. Name the report, choose the quote style, and export Excel, Word or
   both. Files are saved as  (date)_(report name)_FATF Search tool.xlsx/.docx


SENDING RESULTS TO COLLEAGUES
-----------------------------
Methodology_note_TEMPLATE.docx is a plain-English note on what the search
is, how it was done and how to read the exports (now including the Round
column and the new sheets). Fill in the bracketed fields on the front page,
save under a new name and send it with the files. Keep the template itself
unchanged.


OTHER FILES MODE
----------------
Drop Word (.docx), PDF, Excel, .csv or .txt files into source files\other.
Everything in them is searched, and each hit gives its exact location:
  Word:  heading trail + paragraph number or table/row/column
  Excel: sheet, cell and column heading, plus the whole row as Details
  PDF:   page and paragraph number
Old .doc files are not supported - open them in Word and save as .docx.


NOTES
-----
- The tool finds the two CSVs in each round folder by name (one containing
  "RC_IO" or "KFRA", one containing "TECH_COMP" or "TCOMP"), so new
  downloads with different date-stamps just need dropping in. The cache
  notices new files automatically.
- Executive summary rows are counted under "Risk and Context" (use the
  Indicator filter to include or exclude them).
- If a search result says "This search has expired", simply search again
  (the tool keeps the last 6 searches in memory).
