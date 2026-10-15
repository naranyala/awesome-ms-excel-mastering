# Awesome Microsoft Excel [![Awesome](https://awesome.re/badge.svg)](https://awesome.re)

> A curated catalog and reference for everything Microsoft Excel: how to master it step by step, formulas and functions, data analysis, dashboards, automation, AI features, professional use cases, financial modeling, spreadsheet risk, career resources, and communities.

Excel is still the most widely used analysis tool in the world. This list organizes the best resources — plus the knowledge checklists you'd normally have to piece together from a dozen books — into a single path from your first workbook to expert-level models, automated reporting pipelines, and VBA/Python solutions.

## Contents

- [Official Microsoft Resources](#official-microsoft-resources)
- [Mastery Roadmap](#mastery-roadmap)
- [Courses and Learning](#courses-and-learning)
- [Practice and Challenges](#practice-and-challenges)
- [Spreadsheet Fundamentals](#spreadsheet-fundamentals)
- [Formulas and Functions](#formulas-and-functions)
- [Errors and Troubleshooting](#errors-and-troubleshooting)
- [Data Analysis](#data-analysis)
- [Visualization and Dashboards](#visualization-and-dashboards)
- [Automation and Programming](#automation-and-programming)
- [AI in Excel](#ai-in-excel)
- [Professional Use Cases](#professional-use-cases)
- [Financial Modeling](#financial-modeling)
- [Spreadsheet Risk, Governance, and Security](#spreadsheet-risk-governance-and-security)
- [Performance, File Formats, and Platforms](#performance-file-formats-and-platforms)
- [Collaboration and Sharing](#collaboration-and-sharing)
- [Keyboard Shortcuts](#keyboard-shortcuts)
- [Templates](#templates)
- [Books](#books)
- [Blogs, Videos, and Newsletters](#blogs-videos-and-newsletters)
- [Add-ins, Tools, and Libraries](#add-ins-tools-and-libraries)
- [Certification and Career](#certification-and-career)
- [Communities and Forums](#communities-and-forums)
- [Related Lists](#related-lists)
- [Glossary](#glossary)
- [Contributing](#contributing)
- [License](#license)

---

## Official Microsoft Resources

Start here — first-party documentation is always current.

- [Excel help and learning](https://support.microsoft.com/en-us/excel) - Complete support hub: how-tos, troubleshooting, feature guides.
- [Microsoft Learn training catalog](https://learn.microsoft.com/en-us/training/) - Free modules and learning paths, including Excel and Copilot topics.
- [Microsoft Learn — Excel search results](https://learn.microsoft.com/en-us/training/browse/?terms=excel) - Every Excel-related module and learning path in one filtered list.
- [Excel blog (official)](https://techcommunity.microsoft.com/t5/excel-blog/bg-p/ExcelBlog) - What's new: rolling feature releases, function launches, Insider news.
- [Excel (Microsoft 365 product page)](https://www.microsoft.com/en-us/microsoft-365/excel) - Feature overview, what's new, and plans.
- [Excel for the web (launch)](https://www.microsoft365.com/launch/excel) - Browser-based Excel; no install needed.
- [Power Query documentation](https://learn.microsoft.com/en-us/power-query/) - Get data, transform, M language reference, query folding.
- [Power Query in Excel](https://support.microsoft.com/en-us/excel/about-power-query-in-excel) - Get & Transform basics inside Excel.
- [Learn Power Query and Power Pivot](https://support.microsoft.com/en-us/excel/learn-to-use-power-query-and-power-pivot-in-excel) - Microsoft's guided intro to the modern BI stack in Excel.
- [DAX reference](https://learn.microsoft.com/en-us/dax/) - Data Analysis Expressions for the data model (shared with Power BI).
- [Excel VBA reference](https://learn.microsoft.com/en-us/office/vba/api/overview/excel) - Concepts, samples, and object model docs.
- [Excel object model](https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model) - Every VBA object, property, method, and event.
- [Office Scripts documentation](https://learn.microsoft.com/en-us/office/dev/scripts) - TypeScript automation for Excel on the web, Windows, and Mac.
- [Office Scripts API reference](https://learn.microsoft.com/en-us/javascript/api/office-scripts/overview?view=office-scripts) - Full scriptable API surface.
- [Office Scripts samples](https://learn.microsoft.com/en-us/office/dev/scripts/resources/samples/samples-overview) - Real-world scripts: folders, email, charts, Power Automate flows.
- [Office add-ins development](https://learn.microsoft.com/en-us/office/dev/add-ins/) - Build cross-platform task pane and content add-ins (web tech, not VBA).
- [Python in Excel](https://support.microsoft.com/en-us/excel/python/introduction-to-python-in-excel) - Official intro to running Python (pandas, seaborn, scikit-learn) in cells.

## Mastery Roadmap

A realistic path from beginner to expert. Do not skip the "proof of work" column — building things is what makes skills stick.

| Stage | You can already… | Focus next | Proof-of-work milestone |
|---|---|---|---|
| **1. Foundations** (0–4 weeks) | Open and save files, type data | Navigation, selection, formatting, cell references (relative/absolute), `SUM`/`AVERAGE`/`MIN`/`MAX`/`COUNT`, sort and filter, basic charts, print setup, top 10 shortcuts | Build a personal budget with totals and a pie chart |
| **2. Intermediate** (1–3 months) | Write basic formulas confidently | Tables and structured references, `SUMIFS`/`COUNTIFS`/`AVERAGEIFS`, `XLOOKUP` + `INDEX/MATCH`, text and date functions, PivotTables, conditional formatting, data validation, `IFERROR`, Flash Fill | Monthly sales report driven by a PivotTable and slicers |
| **3. Advanced** (3–6 months) | Build reports without help | Power Query (clean, merge, append, refresh), data model + Power Pivot + DAX, dynamic arrays (`FILTER`, `SORT`, `UNIQUE`), `LET`/`LAMBDA`, named ranges, scenario and what-if analysis, workbook design and protection, performance tuning | An automated report that refreshes messy source files with one click |
| **4. Expert** (6–12 months) | Design models others rely on | Financial/statistical modeling, VBA macros and user forms, Office Scripts + Power Automate, Python in Excel, add-in development, error handling and audit practices, teaching/reviewing others' work | End-to-end automated solution + MOS Expert certification (see [Certification and Career](#certification-and-career)) |

**Role-based tracks** (what to study, in order):

| Track | Sequence |
|---|---|
| Business / operations analyst | Fundamentals → tables + `XLOOKUP`/`SUMIFS` → PivotTables → Power Query consolidation → dashboards → light DAX → Copilot |
| Finance / FP&A | Fundamentals → financial functions → modeling conventions → scenarios + Goal Seek/Solver → Power Query (ERP extracts) → variance dashboards |
| Data analyst | Tidy-data rules → Power Query → data model + DAX → visualization rules → Power BI handoff (same skills) |
| Admin / executive assistant | Templates + data validation → protection → sorting/filtering mastery → PivotTable summaries → Office Scripts + Power Automate |
| Developer | VBA object model → debugging + error handling → event-driven workbooks → `.xlam` add-ins → Office Scripts → Python in Excel |

**Learning principles that actually work:**

1. **Rebuild, don't watch.** Pause videos and reproduce the result from scratch.
2. **Learn on real data.** Export files from your job, studies, or hobby and analyze them.
3. **Master 20% first.** Tables, `XLOOKUP`-family lookups, `SUMIFS`, PivotTables, and Power Query cover most office work.
4. **Type formulas by hand** until you can write them without looking — formula fluency beats menu memorization.
5. **Automate last.** Record a macro only after you can do the task manually and know the manual steps.
6. **Teach it.** Answer one question per week on r/excel or an internal team wiki.

**Signature practice projects:**

- Consolidate 12 monthly CSV files into one refreshable master table (Power Query).
- Build a KPI dashboard with a data model, measures, and slicers (Power Pivot/DAX).
- Recreate a manual weekly report as a one-button VBA macro or Office Script.
- Build a 3-statement financial model with scenario switching and an amortization schedule.
- Design a data-entry template with validation, dropdowns, protection, and error checks.

## Courses and Learning

**Free**

- [Excel Easy](https://www.excel-easy.com/) - Well-structured beginner-to-advanced tutorial with screenshots.
- [Chandoo's free Excel course](https://chandoo.org/wp/excel-basics/) - Complete free course with practice files.
- [Excel Exposure](https://www.excelexposure.com/) - Free organized course built for self-paced learning.
- [Contextures](https://contextures.com/) - Deep tutorials on validation, pivots, Power Query, and cleanup.
- [GoSkills Excel catalog](https://www.goskills.com/excel/courses) - Bite-sized lessons (free tier available) across 25 Excel courses.

**Paid and university**

- [Excel Skills for Business (Coursera, Macquarie University)](https://www.coursera.org/specializations/excel) - The most respected structured specialization: Essentials → Advanced, 4 courses with practice workbooks.
- [XelPlus (Leila Gharani)](https://www.xelplus.com/) - Career-focused Excel, Power Query, VBA, Python in Excel, and finance courses.
- [Excel Campus (Jon Acampora)](https://www.excelcampus.com/) - Formulas, Power Query, VBA, and productivity training.
- [Excel Campus — formulas track](https://www.excelcampus.com/formulas/) - Dedicated lookup, logic, and dynamic-array formula training.
- [My Online Training Hub (Mynda Treacy)](https://www.myonlinetraininghub.com/) - Free tutorials plus paid courses on PivotTables, dashboards, Power Query, and DAX; Microsoft MVP-led.
- [Chandoo dashboard training](https://chandoo.org/wp/dashboard/) - Dashboard-focused curriculum with free practice files.
- [Maven Analytics](https://mavenanalytics.io/) - Excel and data analysis paths with real-world projects.
- [LinkedIn Learning](https://www.linkedin.com/learning/) - Broad Excel library (Excel Essential Training, Power Query, VBA) with certificates.
- [Udemy](https://www.udemy.com/) - Affordable beginner-to-advanced courses; check ratings and recency.

## Practice and Challenges

Reading about Excel is not the same as doing Excel.

- [Excel Practice Online](https://excel-practice-online.com/) - Free in-browser exercises: functions from `SUM` to `XLOOKUP`, pivots, text functions.
- [Excel Exercice](https://excel-exercice.com/) - Interactive grid: type the formula, get instant correction; quizzes by topic.
- [Spreadsheet Challenges](https://spreadsheetchallenges.com/) - Real-world problem sets with answer validation and leaderboards.
- [Kaggle datasets](https://www.kaggle.com/datasets) - Free real datasets to practice cleaning, pivoting, and charting.
- **Self-directed practice** - Rebuild a report you receive at work; type along with one video per day; solve one fresh r/excel question nightly (see [Communities and Forums](#communities-and-forums)); keep a `practice/` folder with date-stamped files.

## Spreadsheet Fundamentals

The reference layer most tutorials rush past. Master these and advanced features become much easier.

**Workbook anatomy**

- A workbook contains worksheets (and optionally chart sheets); a cell sits at the intersection of one row and one column.
- Limits: **1,048,576 rows × 16,384 columns** (`A` to `XFD`) in `.xlsx`/`.xlsm`/`.xlsb`; only **65,536 × 256** in legacy `.xls`.
- Modern data hygiene: **one record per row, one variable per column, one table per sheet, no subtotals inside the data, no merged cells** (merged cells break sorting, pasting, and formula references).

**Cell references — the four kinds**

| Example | Meaning | Typical use |
|---|---|---|
| `A1` | Relative | Fills adjust as you copy the formula |
| `$A$1` | Absolute | Fixed assumption (tax rate, FX rate) |
| `$A1` / `A$1` | Mixed | Lock column or row when filling across/down |
| `Sheet2!A1`, `'My Sheet'!A1` | Sheet-qualified | Cross-sheet link (quotes needed if name has spaces) |
| `SUM(Sheet1:Sheet3!A1)` | 3D reference | Same cell across a block of consecutive sheets |
| `R1C1` style (`R[-2]C[-1]`) | Row/Column addressing | Enable in File → Options → Formulas; used by VBA and auditors |

- Press `F4` while editing a reference to cycle `$` combinations; `Ctrl+[` jumps to precedents, `Ctrl+]` to dependents.
- A **circular reference** (formula depends on itself) triggers a warning; intentional circularity requires enabling iterative calculation (File → Options → Formulas) — know it's on, and why.

**Data types**

- Number vs **text-stored-as-number** is the #1 hidden bug: numbers right-align, text left-aligns; `ISTEXT()`/`ISNUMBER()` diagnose; fix with `VALUE()`, `*1`, Paste Special → Multiply, or Text to Columns.
- Dates are numbers (see [Formulas and Functions](#formulas-and-functions)); time is a fraction of a day.
- Errors are their own type (`#N/A`, `#REF!`…), see [Errors and Troubleshooting](#errors-and-troubleshooting).
- **Linked data types** (Stocks, Geography, Organizations, and more) pull live reference data into a cell; built-in advanced charting (maps) can consume them.
- Booleans `TRUE`/`FALSE` coerce to `1`/`0` in math — handy for counting conditions (`SUM(--(A:A="x"))` or `COUNTIF`).

**Number formats** (display only — underlying value never changes)

| Code | Result |
|---|---|
| `#,##0.00` | Thousands separator + 2 decimals: `1,234.56` |
| `0.00%` | Percent: `12.34%` (stored as `0.1234`) |
| `#,##0.00;[Red](#,##0.00)` | Negatives in red parentheses — finance standard |
| `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)` | Accounting alignment (currency, dashes, aligned parens) |
| `d-mmm-yy`, `yyyy-mm-dd`, `m/d/yyyy h:mm AM/PM` | Dates and timestamps |
| `[h]:mm` | Elapsed hours (`26:15`), unlike `h:mm` which wraps at 24 |
| `0.00E+00` | Scientific |
| `# ?/?`, `??/??` | Fractions |
| `00000` | Keep leading zeros (IDs) |
| `"Qty: "0` | Literal text + number |
| `;;;` | Hide everything (blank-looking cell with value) |
| `[Red][<=100];[Blue][>100]` | Conditional format codes by value |

**Tables vs ranges**

- Convert with `Ctrl+T`; tables auto-expand, auto-fill formulas, support slicers, structured references (`Table1[Revenue]`), and refreshable Power Query sources.
- [Overview of Excel tables](https://support.microsoft.com/en-us/excel/overview-of-excel-tables) · [Structured references](https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables)
- Avoid mixing a table with the rest of a formatted sheet; one table per data block.

**Named ranges**

- `Ctrl+F3` opens Name Manager; names make formulas readable (`=Revenue*TaxRate`), define dropdown sources, print areas, and chart series.
- Scope: workbook-wide (default) or sheet-local; names can hold constants or formulas (`TaxRate := 0.21`), not just ranges.
- Create fast with `Ctrl+Shift+F3` (Create from Selection); never let names collide with cell references (avoid `A1`-style names such as `Q1` — Name Manager flags collisions).

**Sorting, filtering, and cleanup**

- Multi-level sort (Data → Sort), sort by color/icon, and **custom lists** (create `Jan, Feb, Mar…` or your own regions order).
- Filter: `Ctrl+Shift+L`; custom number/text filters (contains, begins with, top 10, between dates); filter to unique values.
- **Advanced Filter** for OR-logic across columns; **Remove Duplicates** for quick dedupe; `UNIQUE()` when you want it dynamic.
- **Flash Fill** (`Ctrl+E`) learns patterns (split names, reformat phones); **Text to Columns** for delimited/width splits; **Find & Replace** supports wildcards `* ? ~`.
- **Go To Special**: blanks, formulas, constants, objects, **visible cells only (`Alt+;`)** — essential before copy/paste of filtered data.

**Data validation (input control)**

- Types: list, whole/decimal number, date/time, text length, custom formula.
- Dropdown from a range best practice: put options in a table, name it, reference the name; **dependent dropdowns** via `INDIRECT`/`VLOOKUP`/`XLOOKUP` (modern: `TEXTSPLIT` for multi-level).
- Error alerts have three styles: **Stop** (blocks), Warning (allows), Information (notifies); input messages guide users.
- Validation does **not** stop paste-in of bad data — pair with sheet protection; use Data → Circle Invalid Data to audit existing sheets.
- Modern cell controls: checkboxes and in-cell dropdowns (Microsoft 365).

**Conditional formatting**

- Rule types: cell value/text/date conditions, top/bottom N, above/below average, unique/duplicate, **formula rules** (most powerful), plus data bars, color scales, icon sets.
- Formula rule locking pattern: apply to `$B$2:$Z$100` with rule `=$B2>100` (lock the column, free the row).
- Manage priority and **Stop If True** in Manage Rules; use one highlight color system (red = attention, green = target met).
- Too many overlapping rules slow recalculation — consolidate, or move logic into a helper column.
- Conditional formatting is not stored in the cell — never rely on it for data (export loses it).

**Presentation, navigation, and printing**

- Cell styles (Good/Bad/Neutral, headings) keep workbooks consistent; Format Painter double-click locks for repeated pasting.
- Freeze panes, split, outline/group (rows/columns), hide gridlines/headings for reports.
- Print setup: set print area, **repeat header rows** (Page Layout → Print Titles), fit to one page wide, headers/footers with page numbers, export via Print → Save as PDF.
- The **Camera tool** (add to Quick Access Toolbar) creates a live picture of a range — classic dashboard technique.
- Hidden tools worth knowing: Watch Window, Custom Views, Document Themes, Accessibility Checker.

## Formulas and Functions

**How formulas evaluate**

- Operator precedence applies; use parentheses aggressively. Partial-evaluate any fragment of a formula with **`F9`** (Esc cancels the edit).
- Show every formula at once with `` Ctrl+` ``; navigate with Trace Precedents/Dependents (`Ctrl+[` / `Ctrl+]`), Evaluate Formula dialog, and Watch Window.
- Modern grid behavior: formulas either **spill** (dynamic arrays) or **implicitly intersect** (`@` forces intersection when needed).

**Official references**

- [Excel functions (by category)](https://support.microsoft.com/en-us/excel/excel-functions-by-category) - Every function grouped by purpose.
- [Excel functions (alphabetical)](https://support.microsoft.com/en-us/excel/excel-functions-alphabetical) - Fast lookup when you know the name.
- [ExcelJet function list](https://www.exceljet.net/functions) - Function docs with copy-paste examples.

**The must-know function map**

| Category | Learn these first | Typical use |
|---|---|---|
| Aggregation | `SUM`, `SUMIFS`, `COUNTIFS`, `COUNTA`, `AVERAGEIFS`, `MAXIFS`, `AGGREGATE` | Conditional totals, KPI counts, robust averages |
| Lookup | `XLOOKUP`, `INDEX`+`MATCH`, `FILTER`, `SORT`, `UNIQUE`, `CHOOSECOLS`, `XMATCH` | Join tables, extract subsets, two-way lookups |
| Logic | `IF`, `IFS`, `SWITCH`, `AND`/`OR`, `IFERROR`, `LET` | Decision logic, clean error handling, readable formulas |
| Text | `TEXTJOIN`, `TEXTSPLIT`, `TEXTBEFORE`/`TEXTAFTER`, `TRIM`, `CLEAN`, `PROPER`, `TEXT` | Clean messy strings, split/merge columns, format numbers |
| Date and time | `TODAY`, `EOMONTH`, `EDATE`, `YEARFRAC`, `NETWORKDAYS`, `DATE`, `WEEKDAY` | Ageing, schedules, fiscal calendars, working days |
| Financial | `PMT`, `IPMT`/`PPMT`, `NPV`, `XNPV`, `IRR`/`XIRR`, `FV`, `PV`, `RATE`, `SLN`, `DDB`, `SYD` | Loans, depreciation, investment appraisal |
| Statistical | `MEDIAN`, `PERCENTILE`, `STDEV.S`, `CORREL`, `FORECAST.LINEAR`, `TREND`, `RANK` | Distribution, correlation, forecasting |
| Information | `ISNUMBER`, `ISBLANK`, `NA`, `FORMULATEXT`, `CELL` | Validation, error audits, reverse-engineering models |
| Dynamic arrays | `LAMBDA`, `MAP`, `BYROW`, `BYCOL`, `REDUCE`, `SCAN` | Reusable custom logic without VBA |

**Criteria syntax for `SUMIFS`/`COUNTIFS`/`AVERAGEIFS`** — the most copied wrong part of Excel:

| Goal | Criteria written as |
|---|---|
| Greater than 100 | `">100"` |
| Reference a cell (e.g. `$F$1`) | `">="&$F$1` (**always concatenate with `&`**) |
| Date on/after a date cell | `">="&$D$1` or `">="&DATE(2026,1,1)` |
| Exact text | `"North"` or `"=North"` |
| Not equal / not blank | `"<>Failed"`; any non-blank cell: `"<>"` |
| Starts with | `"North*"` |
| Contains | `"*"&$A1&"*"` |
| Wildcards | `*` any run, `?` one char, `~` escapes them (`~*` literal star) |
| Between | two criteria pairs: `">="&10, "<="&100` |
| Multiple AND conditions | additional range/criteria pairs (up to 127) |

**Lookup decision tree**

| Need | Reach for |
|---|---|
| Single value, left-to-right or right-to-left | **`XLOOKUP`** (default for Microsoft 365) |
| Compatibility with old files | `INDEX`+`MATCH`, then legacy `VLOOKUP`/`HLOOKUP` |
| Multiple criteria | `XLOOKUP(1,(rng1=a)*(rng2=b),result)` or `SUMIFS` when the result is numeric |
| Return several columns | `XLOOKUP` with `CHOOSECOLS`, or `INDEX`/`MATCH` column numbers — or move it to Power Query |
| Partial / fuzzy match | `XLOOKUP(...,match_mode:2)` (wildcards), `XMATCH` with wildcard, or `FILTER` |
| Two-way (row + column) | nested `XLOOKUP` or `INDEX(r, MATCH, MATCH)` |
| Extract many rows | **`FILTER`** (dynamic array — one formula, spills results) |
| Does a value exist at all? | `COUNTIF>0`, `ISNUMBER(MATCH(...))`, or `XLOOKUP` with `if_not_found` |
| Lookup returns `#N/A` handling | `IFNA(...,"Not found")` — prefer `IFNA` over `IFERROR` for lookups |

**Modern Microsoft 365 function families** (rolling out continuously — check the Excel blog under [Official Microsoft Resources](#official-microsoft-resources)):

- **Dynamic arrays:** `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `RANDARRAY`
- **Text:** `TEXTSPLIT`, `TEXTBEFORE`, `TEXTAFTER`, `TEXTJOIN`, plus newer `REGEXEXTRACT`/`REGEXTEST`/`REGEXREPLACE`
- **Reshaping:** `VSTACK`, `HSTACK`, `TOCOL`, `TOROW`, `WRAPROWS`, `WRAPCOLS`, `TAKE`, `DROP`, `EXPAND`, `CHOOSEROWS`, `CHOOSECOLS`
- **Aggregating arrays:** `GROUPBY`, `PIVOTBY` (pivot-table-like results in a spill range)
- **Programming:** `LET` (name intermediate results), `LAMBDA` (write your own functions), `MAP`, `BYROW`, `BYCOL`, `SCAN`, `REDUCE`
- **Spill syntax:** `range#` (whole spill), `@` (implicit intersection), `A1:B10` normal ranges
- **Availability caveat:** dynamic arrays, `LAMBDA`, Copilot, and Python in Excel require Microsoft 365; perpetual releases (Excel 2024/2021/2019) lag or lack them.

**Reference and volatility**

- **Volatile functions** recalculate on *every* change: `NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `OFFSET`, `INDIRECT`, `CELL`, `INFO`. Replace `OFFSET` with `INDEX`, `INDIRECT` with direct references or `XLOOKUP`.
- Prefer bounded ranges (`$A$2:$A$5000`) or table columns (`Table1[Amount]`) over whole-column refs in volatile-heavy models.
- `INDIRECT` breaks when sheets are renamed and doesn't work in closed linked workbooks — avoid in distributed models.

**Date and time essentials**

- Serial `1` = 1900-01-01; today's date is ~46,000+; time is a fraction (`0.5` = 12:00). Add days to dates directly (`date + 30`).
- Excel deliberately treats **1900 as a leap year** (serial 60 = fake 1900-02-29) for Lotus compatibility — beware when doing day counts across early 1900 dates.
- Legacy Mac workbooks may use the **1904 date system** (File → Options → Advanced → When calculating this workbook) — dates shift by 4 years and 1 day if toggled.
- Workhorses: `EOMONTH` (month end), `EDATE`, `YEARFRAC` (day-count for finance), `NETWORKDAYS`/`WORKDAY` (+`.INTL` for custom weekends), `DATEDIF` (hidden but works), `WEEKDAY`/`ISOWEEKNUM`.
- Fiscal year: `YEAR(date+(12-month_offset))` patterns or a fiscal-month helper column feeding `SUMIFS`.

**Text cleaning recipes**

- Trim spaces: `TRIM`; remove non-printables: `CLEAN`; fix case: `PROPER`/`UPPER`/`LOWER`.
- Keep leading zeros: format the column as Text *before* pasting, or `TEXT(A1,"00000")`.
- Numbers as text → `VALUE`, `--A1`, or multiply by 1 via Paste Special.
- Split full name: `TEXTBEFORE`/`TEXTAFTER`, `TEXTSPLIT(A1,," ")`, or Text to Columns / Flash Fill for one-offs.
- Extract digits: modern `REGEXEXTRACT(A1,"\d+")`; classic: `MID`+`SEARCH` gymnastics or helper columns.

**Error handling patterns**

- Lookups: `=IFNA(XLOOKUP(...),"Not found in master")`.
- Division: `=IFERROR(revenue/volume,0)` — but only where `0` is meaningful; otherwise show `N/A` state.
- `LET` to compute once: `=LET(r,XLOOKUP(...),IF(r>target,r,0))`.
- Audit formulas: `FORMULATEXT`, `ISFORMULA`, and conditional formatting with `=ISERROR($B2)` across report cells.

**Financial function conventions**

- Cash-flow **signs matter**: money out is negative, money in positive (Excel returns an error if all flows share a sign).
- Match units: `rate` per period and `nper` in periods (monthly rate for monthly payments).
- `NPV` assumes period-end cash flows; **`XNPV`/`XIRR` use actual dates** — prefer them for real-world valuations.
- `IRR` may need a `guess` for unusual flows; use `MIRR` when reinvestment assumptions matter.
- `SLN`, `DDB`, `SYD` cover straight-line and accelerated depreciation; link schedules to a live asset register.

**`LAMBDA` patterns**

- Define once in Name Manager: `SquareRoot = LAMBDA(x, SQRT(x))` or business rules like `TieredCommission = LAMBDA(amount, IFS(amount>100000,amount*0.1, amount>50000,amount*0.07, TRUE,amount*0.05))`.
- Document parameters in the name description; test with a small range before rolling out.

## Errors and Troubleshooting

| Error | Meaning | Typical cause | Fix |
|---|---|---|---|
| `#DIV/0!` | Division by zero | Denominator empty/zero | `IFERROR`/guard with `IF(denominator=0,…)` |
| `#NAME?` | Unknown name | Typo, missing quotes, function not in this version | Check spelling, quotes around text, version support |
| `#N/A` | Value not found | Lookup miss | `IFNA` + verify lookup column is clean (spaces, text vs number) |
| `#REF!` | Invalid reference | Referenced row/column/sheet deleted | Rebuild references; audit with Find `#REF!` |
| `#VALUE!` | Wrong data type | Text in math, wrong argument type | `ISNUMBER` check; `VALUE`/`TRIM` the inputs |
| `#NULL!` | No intersection | Space used instead of comma between ranges | Replace space with `,` or `:` |
| `#NUM!` | Invalid number | Out-of-domain values (negative rate periods, `IRR` fails) | Check inputs and assumptions |
| `#SPILL!` | Spill blocked | Cells occupied in the spill range | Clear the obstructing cells |
| `#CALC!` | Empty array result | `FILTER`/calculation returned nothing | Supply `if_empty` argument |
| `#GETTING_DATA` | Live data loading | Linked data type still fetching | Wait or refresh; check connectivity |
| `#PYTHON!` `#BUSY!` `#CONNECT!` | Python in Excel states | Python computing / connection issue | See the Python in Excel docs under [Official Microsoft Resources](#official-microsoft-resources) |

**Debugging toolkit (in order of use):**

1. `` Ctrl+` `` — see all formulas and spot copy/paste mistakes.
2. Select a formula fragment → `F9` — evaluate just that part (Esc to cancel).
3. **Evaluate Formula** dialog (Formulas tab) — step through nested evaluation.
4. Trace Precedents/Dependents (`Ctrl+[`, `Ctrl+]`) and Watch Window.
5. Conditional-format `=ISERROR($B2)` over reports so errors can't hide.
6. Data → Text to Columns / `VALUE` sweep to find text-stored numbers.
7. Name Manager (`Ctrl+F3`) to spot broken names (`#REF!` in the Refers To box).
8. For slow workbooks: Formulas → Calculation Options → Manual, then switch to Automatic and watch what recalculates.

## Data Analysis

**Tidy data rules** (the precondition for everything below): one header row; unique column names; no blank rows/columns; one fact per cell; consistent data types per column; dates as dates, numbers as numbers.

**PivotTables — the 20% that does 80%**

- Areas: Rows, Columns, Values, Filters. Right-click any value → **Value Field Settings** → *Show Values As*: % of grand total, % of column, difference from, running total in, index.
- Group dates (days/months/quarters/years) and numeric bands in place; grouped fields live in the PivotTable cache.
- **Calculated Field** (formula over value fields) vs **Calculated Item** (formula over a row/column item) — most people need the former.
- Slicers and timelines connect multiple PivotTables (PivotTable Analyze → Report Connections) — this is how self-service dashboards work.
- Double-click any value to **drill down** to the underlying records; `GETPIVOTDATA` keeps a stable reference when rows move.
- Keep source in an **Excel Table** (`Ctrl+T`) so refresh picks up new rows; turn on "Refresh data when opening the file" for always-current reports.
- Classic pitfalls: duplicate/blank headers, numbers stored as text, merged cells above the data, filtered source hiding rows, and forgetting Refresh after edits.
- Big data: load to the **data model** (Power Pivot) instead of a worksheet range — handles millions of rows with relationships and DAX measures.

**Power Query (Get & Transform) — repeatable ETL**

- Workflow: **Get data → transform (applied steps) → load → refresh**. Every step is recorded as M code you can edit in the Advanced Editor.
- Core operations: split/merge columns, fill down, change types, replace values, unpivot columns, group by, merge (joins), append (unions), conditional column, index column, custom column.
- **Parameters** (e.g. `FolderPath`, `AsOfDate`) drive dynamic queries; **reference queries** build staging layers without duplication; **function queries** process each file in a folder (the classic "combine 12 monthly files" pattern).
- **Query folding** pushes steps back to the source (SQL, OData) for speed — visible in "Query folding" preview; some steps break folding and force full data pull locally.
- Error handling in M: `try … otherwise`, `Table.ReplaceErrorRows`, and setting data types explicitly early.
- Performance: don't load intermediate queries to the worksheet (load to Connection only), type columns once at the start, prefer hash-friendly merges, and filter rows at the top of the query.
- Use **PivotTable-style aggregation in Power Query** when the output never needs interactive slicing; keep pivots on top of clean loaded data.
- Docs: Power Query documentation and guides under [Official Microsoft Resources](#official-microsoft-resources).

**Data model, Power Pivot, and DAX**

- Load tables to the data model, relate them like a **star schema** (fact tables + dimension tables), keep **filter flow** one-directional unless you have a measured reason not to.
- **Calculated column** = row-by-row value (stored, good for slicers/grouping); **Measure** = aggregate evaluated in context (flexible, compact). Default to measures for KPIs.
- `CALCULATE` is the one function to truly learn: it modifies filter context (e.g. `CALCULATE([Sales], ALL('Date'))`).
- Time intelligence (`SAMEPERIODLASTYEAR`, `TOTALYTD`, `DATEADD`) requires a marked **Date table** covering all periods.
- Common mistakes: bidirectional filters on large tables, summing a column of ratios, ignoring filter context, measures that scan whole tables unnecessarily.
- Study: [DAX Guide](https://dax.guide/), [SQLBI](https://www.sqlbi.com/) (deep articles and videos), plus the official DAX reference under [Official Microsoft Resources](#official-microsoft-resources).

**What-if, optimization, and forecasting**

- **Scenario Manager**: save named input sets (Base/Up/Bad) and show summaries side by side.
- **Data Tables**: one- and two-variable sensitivity grids (input rows/columns + formula in the corner).
- **Goal Seek**: find the input that hits a target result (small models only — no dependents).
- **Solver** (File → Options → Add-ins): three engines — GRG Nonlinear, **Simplex LP** (linear problems: resource allocation, product mix), Evolutionary (non-smooth). Always set constraints explicitly and keep an answer scenario.
- **Forecast Sheet** (Data tab) and `FORECAST.ETS` for seasonal univariate forecasts; backtest against a holdout period.

**Statistics**

- **Analysis ToolPak** (add-in): descriptive stats, histograms (with Pareto), correlation, covariance, t/F/z tests, ANOVA, regression, sampling, Fourier.
- `CORREL`, `SLOPE`/`INTERCEPT`, `RSQ`, `STDEV.S`, `PERCENTILE.INC`, `RANK.EQ` for lighter work.
- Beyond ToolPak: Python in Excel (scikit-learn, statsmodels) or export to Power BI/R/Python for serious modeling.

**Data connections**

- Get Data sources: local Excel/CSV/Text, folder, SharePoint/OneDrive, web pages, OData, SQL, Dataverse, and many more (Home → Get Data).
- Manage connections (Data → Queries & Connections), **Refresh All** (`Ctrl+Alt+F5`), background vs foreground refresh, and refresh-on-open.
- Live linked types: `STOCKHISTORY`, Stocks/Geography data types, and RTD-based feeds (finance teams).

## Visualization and Dashboards

**Chart chooser — match the message to the form**

| You need to show… | Use |
|---|---|
| Comparison across categories | Clustered bar/column (bar if labels are long) |
| Ranking | Sorted bar chart |
| Trend over time | Line chart (multi-series for few entities) |
| Part-to-whole | Stacked bar/100% stacked column; pie only for ≤5 slices, one series, one moment |
| Deviation / bridge | **Waterfall** (native) |
| Distribution | Histogram, box & whisker, Pareto (native chart types) |
| Relationship / correlation | Scatter with fitted trend line |
| Composition hierarchy | Treemap or sunburst (native) |
| Pipeline stages | Funnel (native) |
| Status vs target | KPI card, data bars, icon sets (in-cell) |
| Geospatial | Filled map (connected data types) |
| Schedule | Stacked bar Gantt (invisible base series) |
| Big number | Callout / KPI text — not a chart |

**Native modern chart types** (Insert → Recommended/All Charts): column, bar, line, area, pie/doughnut, radar, bubble, scatter, stock, waterfall, funnel, treemap, sunburst, histogram, Pareto, box & whisker, map, combo + secondary axis, sparklines, in-cell data bars.

**Charts to avoid (default advice):** 3D anything (especially 3D pie), truncated y-axes on bar charts, dual axes without a clear second scale, rainbow palettes, heavy gridlines, chartjunk legends when direct labels work, and charts whose message needs a paragraph of explanation.

**Design rules**

- One chart = one message; title it as an assertion ("Revenue up 12% QoQ"), not "Revenue Chart".
- Sort bars by value (unless time order matters); label directly and delete redundant legends.
- Bars start at zero; lines may not, but annotate the range.
- 2–4 colors max; reserve red/green for status only; ensure colorblind-safe palettes.
- Consistent fonts, number formats (`$1.2M` not `$1,234,567.89`), and units on axes.
- Reference sites: Peltier Tech (advanced charting), Excel Dashboard School, and Storytelling with Data (design thinking) — all linked under [Blogs, Videos, and Newsletters](#blogs-videos-and-newsletters).

**In-cell visualization**

- **Sparklines**: tiny in-cell charts (line/column/win-loss), one per row, resizable with row height; group formatting applies to all.
- **Data bars** for magnitude, **color scales** for heat, **icon sets** for status — often better than a full chart in tables.
- Conditional-format-driven KPI cells (`=IF(target met,"✔","▲")`) for scannable status rows.

**Dashboard anatomy**

1. Header: title, reporting period, "data refreshed" timestamp, owner.
2. KPI strip: 4–6 headline metrics with variance vs target/period.
3. Main visuals: 1–3 charts driven by the data model.
4. Filter panel: slicers/timelines (connected to everything on the page).
5. Detail table: supporting numbers, zebra-free, right-aligned numerics.
6. Footer: metric definitions, caveats, contact.

**Build checklist**

- [ ] Single refreshable source (table/query/data model) — zero manual re-typing
- [ ] All slicers connected; default state looks right
- [ ] Readable at 100% zoom; prints to PDF legibly
- [ ] No visible `#N/A`/`#REF!` anywhere (search before publishing)
- [ ] Units, period, and currency stated once in the header
- [ ] Input cells protected; calculation cells unlocked for review
- [ ] Metrics defined in a Notes/Definitions sheet
- [ ] Tested with extreme values (0, blanks, negatives, huge numbers)
- [ ] File size sane; heavy formatting not applied to whole rows
- [ ] Stakeholder walkthrough completed before automation

## Automation and Programming

**Pick the right tool**

| Need | Use | Notes |
|---|---|---|
| Quick desktop-only repeat task | Macro recorder + VBA | Fastest start; recorder generates working code to edit |
| Complex logic, forms, events, add-ins | VBA | Full object model; still the deepest Excel automation |
| Cloud/web, share with team, no macro prompts | Office Scripts (TypeScript) + Power Automate | Works in Excel for the web, Windows, Mac; no macro security dialogs |
| Data science, stats, ML inside the grid | Python in Excel (`=PY`) | pandas, seaborn, scikit-learn run in Microsoft's cloud |
| Reusable custom worksheet function | `LAMBDA` and names | No code, spreadsheet-native |
| Cross-app workflows (email, Teams, forms, folders) | Power Automate (with or without scripts) | Scheduled flows, approvals, file automation |

**VBA essentials**

- Object path: `Application` → `Workbooks` → `Worksheets` → `Range`. Everything hangs off `Range`.
- Always start modules with `Option Explicit` (forces variable declaration).
- Avoid `Select`/`Activate` — reference ranges directly (`Worksheets("Data").Range("A1").Value = x`); use `With…End With` for grouped properties.
- Error handling skeleton:

```vb
Sub DoThing()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    ' ... work ...
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    MsgBox "Failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub
```

- **Events** (code that runs automatically): `Workbook_Open`, `Workbook_BeforeClose`, `Workbook_SheetChange`, `Worksheet_BeforeDoubleClick`, `Workbook_SheetActivate`.
- **UDFs** (user-defined functions) must be in a standard module; note they are not available in Excel for the web and can be volatile.
- Debugging: `F8` step through, `F9` breakpoint, Locals window, Immediate window (`? Range("A1").Value`), Watch Window, `Application.StatusBar` for progress.
- Common runtime errors: `1004` (application/object-defined — usually bad range reference), `9` (subscript out of range — missing sheet), `13` (type mismatch).
- Distribution: save as `.xlsm`, add a digital signature for controlled environments, or package as an `.xlam` add-in.
- Study: the Excel VBA reference under [Official Microsoft Resources](#official-microsoft-resources), plus [Excel Macro Mastery](https://www.excelmacromastery.com/), [Automate Excel](https://www.automateexcel.com/), [Wise Owl](https://www.wiseowl.co.uk/), [Rubberduck VBA](https://github.com/rubberduck-vba/Rubberduck).

**Office Scripts**

- The **Action Recorder** captures your clicks as TypeScript you can edit; scripts live in OneDrive and can be shared in your organization or bound to workbook buttons.
- Runs from Excel (web/Windows/Mac) or triggered by **Power Automate** (recurrence, file added, email arrived, Teams message) — this is the modern replacement for "macro that emails a report".
- Start with the Office Scripts docs and samples under [Official Microsoft Resources](#official-microsoft-resources); trigger flows with [Power Automate](https://powerautomate.microsoft.com/).

**Python in Excel**

- `=PY()` turns a cell into Python; `xl("A1:B10")` bridges ranges; results return as Python objects or Excel values.
- Runs in Microsoft's cloud with Anaconda libs (pandas, NumPy, scikit-learn, seaborn, statsmodels); external reads must go through Power Query.
- Great for: forecasting, statistical tests, ML-style scoring, and plots Excel can't make — while staying in the workbook everyone already uses.

**Add-in development**

- Classic route: VBA `.xlam` add-ins; modern route: **Office add-ins** (task panes built with HTML/JS — cross-platform, centrally deployed; docs under [Official Microsoft Resources](#official-microsoft-resources)); high-performance .NET UDFs via [Excel-DNA](https://excel-dna.github.io/).

## AI in Excel

- [Microsoft Copilot](https://www.microsoft.com/en-us/microsoft-365/copilot) - Official overview of Copilot across Microsoft 365.
- [Work smarter in Excel with Copilot Chat](https://learn.microsoft.com/en-us/training/modules/work-smarter-excel) - Free Microsoft Learn module: generating formulas, summarizing data, creating visualizations.
- **Python in Excel** - combine natural-language prompting with pandas-level analytics in cells (see [Official Microsoft Resources](#official-microsoft-resources)).
- **Also built in:** *Analyze Data* (suggests pivots/charts from a range), *Ideas* in older builds, formula suggestions, and Copilot-assisted conditional formatting.
- **Prompting tips:** describe the outcome and the range ("From Table1, show monthly revenue by region as a column chart with YoV %"); ask Copilot to explain a formula before accepting it; iterate in a copy of the workbook.
- **Guardrails:** verify AI formulas against a small hand-checked sample; never paste confidential data into unapproved tools; keep a known-good version of critical workbooks; treat generated VBA as untrusted code.

## Professional Use Cases

| Role / industry | Typical deliverables | Core Excel toolkit |
|---|---|---|
| FP&A / Finance | Budgets, forecasts, variance analysis, cash-flow models | Financial functions, scenario manager, `XLOOKUP`/`SUMIFS`, linked schedules, Power Query for ERP extracts |
| Accounting | Reconciliations, amortization and depreciation schedules, close checklists | Tables, `SUMIFS`, date logic, protection, audit checks, pivot summaries |
| Sales / RevOps | Pipeline tracker, quota vs. actual, commission calculators, funnel analysis | PivotTables, `XLOOKUP`, conditional formatting, data validation, slicers |
| Marketing | Campaign performance, attribution models, budget pacing | Power Query (ad platforms/CSV), charts, `SUMIFS` per channel, cohort tables |
| HR / People Ops | Headcount plan, turnover and retention, payroll prep | Data validation dropdowns, `COUNTIFS`, date functions, dashboards, protected templates |
| Operations / Supply chain | Inventory counts, reorder points, EOQ, capacity planning | Solver, what-if tables, `INDEX/MATCH`, dashboard alerts, Power Query consolidation |
| Data analyst / BI | Refreshable datasets, KPI models, self-serve reports | Power Query, data model + DAX, PivotTables, Power BI handoff |
| Engineering | Unit conversions, tolerances, measurement analysis | Engineering functions, regression, charts with error bars, templates |
| Research / Academia | Survey analysis, statistics, experiment tracking | Analysis ToolPak, `CORREL`, regression, clean data structure, reproducible formulas |
| Project management | Gantt charts, RACI, status dashboards, risk registers | Conditional formatting timelines, charts, dropdowns, filters |
| Small business / admin | Invoices, inventory, CRM-lite, expense tracking | Templates, tables, validation, simple formulas, print-ready layouts |

**Recurring real-world tasks this catalog prepares you for:**

1. Consolidate messy monthly files from 12 sources → Power Query folder query.
2. Turn raw transactions into an executive dashboard → data model + DAX + PivotChart.
3. Replace a fragile VLOOKUP spreadsheet → `XLOOKUP` with error handling and tables.
4. Automate a weekly report email → Office Script + Power Automate (or VBA + Outlook).
5. Build a budget model with downside/base/upside cases → scenarios + goal seek.

## Financial Modeling

**Model anatomy** (arrive in this order; never interleave):

1. **Cover / README** — title, purpose, version, author, date, contents, how to refresh.
2. **Assumptions** — every driver on one sheet (rates, growth, headcount, prices), blue inputs only.
3. **Inputs / raw data** — imported actuals (prefer Power Query), clearly marked.
4. **Calculation engine** — dated columns (12 months + 4 quarters + 4 years typical), no formatting distractions.
5. **Outputs** — statements, KPIs, charts, summaries; only formulas, no inputs.
6. **Checks** — balance checks, tie-outs, error scans: `=IF(ABS(a-b)<tol,"OK","CHECK")`.

**Conventions that separate pros from amateurs**

- **Color code:** blue = hardcoded input, black = calculation, green = link from another sheet; legend on the cover.
- **Sign convention:** outflows negative, consistently, everywhere.
- **Units row** on top of every column ("$ in thousands") — prevents the classic scale error.
- One assumption per cell; never bury an input inside a formula.
- Don't break formulas when inserting rows/columns: use tables and named ranges; protect finished sheets.
- Keep periods in columns, items in rows; place new assumptions in the assumptions block, not beside the calculation.
- External workbook links are a last resort — they break silently; import via Power Query instead.

**Model types and what they need**

| Model | Key techniques |
|---|---|
| Three-statement (IS/BS/CF) | Linked schedules (PP&E, debt, equity), circularity management (interest & taxes), balance check |
| DCF / valuation | `XNPV`/`XIRR`, WACC assumptions, terminal value, sensitivity table on exit multiple and WACC |
| Operating budget / forecast | Driver-based build (volume × price), scenario switches, variance vs actual |
| Cash forecast | Direct method by week, receivable/payable ageing with `SUMIFS`, minimum cash alert |
| Loan / debt schedule | `IPMT`/`PPMT`, amortization roll-forward, covenant tests |
| Depreciation schedule | `SLN`/`DDB`/`SYD` per asset class, roll-forward (opening + additions − disposals − charge) |
| SaaS / unit economics | MRR/ARR roll-forward, churn (`LTV/CAC`, payback), cohort tables via `SUMIFS` or Power Query |
| Sensitivity / scenario | Two-way data tables, `CHOOSE`/switch cells for Base/Up/Down, tornado-style one-at-a-time |

**Standards and references**

- [FAST Standard](https://fast-standard.org/) - The widely used financial modelling standard (flexible, auditable, transparent structure).
- General good practice: FAST or equivalent house rules, version control in the file name (`FY26_Budget_v3.2_2026-10-15.xlsx`), an audit trail of changes, and a named model owner.

**Common modeling errors (check for these)**

- Sign flips (negative debt draw, positive interest expense)
- Unit mismatches (thousands vs units; a missing `/1000`)
- Wrong day-count or period offsets (day 0 vs day 1 payments)
- Copy/paste that silently moved an absolute reference
- Off-by-one periods (opening balance ≠ prior closing)
- Broken links after a file move (`#REF!` in Name Manager)
- Circular references left unintended (iterative calc hiding real errors)
- Stale outputs: forgetting to refresh queries or recalculate in manual mode

**Pre-publish review checklist**

- [ ] All checks sheet items = OK
- [ ] Balance sheet balances / cash reconciles to source
- [ ] Changing each major assumption flows end-to-end
- [ ] Sensitivity outputs are economically sensible
- [ ] No `#REF!`, `#N/A`, or hardcoded numbers in the engine
- [ ] No external links (or documented and tested)
- [ ] Units, signs, and dates consistent across sheets
- [ ] Documented version history and owner
- [ ] Protected with inputs unlocked
- [ ] Second-person review completed

## Spreadsheet Risk, Governance, and Security

**Why this matters**

- Peer-reviewed spreadsheet research (notably Raymond Panko's work on spreadsheet errors) consistently finds that **manually developed, unaudited workbooks contain human error at very high rates** — often more than half in studies — and that errors are hard to detect even with review.
- Well-documented public failures:
  - **Reinhart–Rogoff (2013)** — a widely cited austerity study's headline result traced to an Excel aggregation error (missing rows in an `AVERAGE` range).
  - **JPMorgan "London Whale" (2012)** — the internal review criticized a risk model's reliance on fragile, error-prone spreadsheets.
  - **UK NHS Test and Trace (2020)** — thousands of COVID cases dropped when a legacy `.xls` file exceeded its row/column limits and truncated data.
- Lesson: spreadsheets are *applications written by non-developers*. Treat them with engineering discipline.

**Build-time controls**

- Design first: inputs / calculations / outputs separated; no merged cells; one table per data block; documented assumptions.
- Data validation on all entry cells; protection on all calculation cells; error handling in every lookup.
- Prefer Tables + Power Query over manual paste chains; make refresh one click, not fifteen.
- Version in file names; keep a "what changed" log inside the workbook; never overwrite the master via email attachment.

**Testing techniques**

- **Reconciliation:** totals must tie to a trusted source (bank, ERP, control total).
- **Dual calculation:** rebuild a key output a second, different way and compare.
- **Edge cases:** zeros, blanks, negatives, text-in-number columns, very large values, missing categories.
- **Dependency stress test:** move one assumption ±20% and confirm every downstream sheet reacts.
- **Size test:** paste the full expected data volume; confirm no row limits or slowdowns.
- **Error sweep:** find `#REF!`/`#N/A`/hardcodes (Ctrl+\` and conditional formatting).

**Audit and review tools**

- **Inquire** add-in (enable in Excel Options → Add-ins): workbook structure, cell relationships, style comparison.
- **Spreadsheet Compare** (ships with Office): line-by-line cell, formula, and formatting differences between two versions.
- Form auditing: Trace Precedents/Dependents, Watch Window, Name Manager, `FORMULATEXT`.
- Independent review: walkthrough of the logic with someone who did not build it.

**Protection and security layers**

| Layer | What it does | Limits |
|---|---|---|
| Cell locking + sheet protection | Stops accidental edits to formulas | Password is not encryption; users can unprotect if no password |
| Workbook structure protection | Prevents sheet add/delete/move | Weak — do not rely on alone |
| File password (open) | Encrypts the whole file | Lost password = lost file; use org key management |
| Protected View | Isolates files from internet/email origins | Users can click through — train them |
| Macro settings (Trust Center) | Disable with notification; signed macros only | Main malware vector for `.xlsm` |
| Sensitivity labels / encryption (Microsoft Purview) | Persistent protection, DLP, access controls | Requires organization licensing |

- Treat `.xlsm` from outside your org like an executable: only enable macros from trusted sources.
- Keep `Show formulas` on during review; turn off before distributing.
- Consider a "release" step: copy to a read-only/PDF distribution file from the protected master.

## Performance, File Formats, and Platforms

**File formats**

| Format | Macros | Notes / when to use |
|---|---|---|
| `.xlsx` | No | Default, open standard, safest to share |
| `.xlsm` | Yes | Only when automation is required |
| `.xlsb` | Yes | Binary — often faster open/save and smaller for large, calc-heavy workbooks; proprietary |
| `.xls` | Yes | Legacy 2003: 65,536 rows × 256 cols — avoid; truncation causes data loss |
| `.csv` / `.txt` | No | One sheet, values only, no formats; watch encoding (UTF-8 vs ANSI) and local list separator (`,` vs `;`) |
| `.xltx` / `.xltm` | Yes in `.xltm` | Template formats (File → Save As) |
| `.xlam` | Yes | VBA add-in |
| `.ods` | No | OpenDocument interchange |
| `.pdf` | No | Distribution/print fidelity, not for re-use of data |

**Performance levers (in rough order of impact)**

1. Formulas → Calculation Options → **Automatic / Manual / Automatic except data tables**; use Manual during heavy editing.
2. Eliminate volatile functions (`INDIRECT`, `OFFSET`, `NOW` chains) — each one forces global recalculation.
3. Keep ranges tight: don't format entire rows/sheets; delete unused rows/columns beyond the used range (bloats file size and `Ctrl+End`).
4. Convert flat ranges to Tables so structured references replace `A2:A100000` patterns that get copied everywhere.
5. Reduce conditional formatting rule count and overlapping ranges.
6. Pivot caches and chart source ranges: point at Tables; clear unused caches.
7. Minimize external workbook links (they force open/recalc of other files); bring data in via Power Query instead.
8. For heavy models: save as **`.xlsb`**; consider moving detail to the data model.
9. Use **64-bit Office** with plenty of RAM for large workbooks.
10. Compress images, avoid array formulas over entire columns, and split rarely-coherent mega-workbooks with Power Query consolidation.
11. Check `Ctrl+End` — if it lands far below your data, clear the phantom range.
12. If Excel must stay open next to other heavy apps: disable previews/handlers only as a last resort; fix the model first.

**Platforms and versions**

| Platform | Highlights and gaps |
|---|---|
| Windows | Full feature set: VBA, Power Query, Power Pivot, Solver, Inquire, all chart types |
| Mac | Most core features + VBA (some API differences), Power Query; check Microsoft's platform feature documentation before relying on an add-in |
| Web | No VBA; has Office Scripts, co-authoring, and core formulas/pivots; use for review/light edits |
| Mobile (iOS/Android) | View and light editing; not a modeling environment |

- **Licensing:** Microsoft 365 subscription = rolling features (dynamic arrays, `LAMBDA`, Copilot, Python in Excel, new chart types). Perpetual releases (Excel 2024/2021/2019) are frozen — check the feature matrix before designing for a specific audience.
- **Update channels:** Current Channel / Monthly Enterprise / Semi-annual, plus the Microsoft 365 Insider program for previews — organizations may lag public releases.
- Old files: `.xls` (65k rows) and 1904 date setting are the two classic cross-platform traps.

## Collaboration and Sharing

- **One source of truth:** store the master in OneDrive/SharePoint and share a *link*, never an email attachment; use *Version history* to restore previous states.
- **Co-authoring:** simultaneous editing with AutoSave; conflicts are rare but still test your validation/protection setup.
- **Permissions:** view vs edit vs review links; block download for sensitive reports; audit external sharing.
- **Threaded comments and @mentions** replace "final_v7_USE_THIS.xlsx" email chains; keep discussions attached to the cells.
- **Compare versions:** Review → Compare (desktop) or Spreadsheet Compare for formula-level diffs.
- **Distribution:** Publish to PDF/Power BI for read-only consumers; reserve `.xlsx` for people who need to interact.
- **Automation handoff:** scheduled refresh (query or Power Automate flow) means nobody has to ask "is this the latest file?".

## Keyboard Shortcuts

Full official references: [Keyboard shortcuts in Excel (Windows)](https://support.microsoft.com/en-us/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f) · [Common Office for Mac shortcuts](https://support.microsoft.com/en-us/office/common-office-for-mac-keyboard-shortcuts-307f469a-1a05-4263-9b55-49df19a8b8ff) · [Customize keyboard shortcuts in Excel for Mac](https://support.microsoft.com/en-us/office/create-a-custom-keyboard-shortcut-for-office-for-mac-6bbeb90e-96d9-4e03-b199-fc026ebdc321)

| Action | Shortcut |
|---|---|
| Save / Open / Print / Find / Replace | `Ctrl+S` / `Ctrl+O` / `Ctrl+P` / `Ctrl+F` / `Ctrl+H` |
| Undo / Redo | `Ctrl+Z` / `Ctrl+Y` |
| Cut / Copy / Paste / Paste Special | `Ctrl+X` / `Ctrl+C` / `Ctrl+V` / `Ctrl+Alt+V` |
| Paste values fast | `Ctrl+Alt+V`, then `V`, then `Enter` |
| Create Table | `Ctrl+T` |
| Toggle filter | `Ctrl+Shift+L` |
| AutoSum | `Alt+=` |
| Quick Analysis | `Ctrl+Q` |
| Flash Fill | `Ctrl+E` |
| Format Cells dialog | `Ctrl+1` |
| Insert Function dialog | `Shift+F3` |
| Edit active cell | `F2` |
| Insert or edit a note (legacy comment) | `Shift+F2` |
| Repeat last action / cycle `$` in references | `F4` |
| Evaluate selected formula fragment | `F9` |
| Jump to edge of data | `Ctrl` + arrow key |
| Extend selection | `Ctrl+Shift` + arrow key |
| Select entire column / row / current region | `Ctrl+Space` / `Shift+Space` / `Ctrl+A` |
| Select visible cells only (skip hidden rows) | `Alt+;` |
| Fill down / Fill right | `Ctrl+D` / `Ctrl+R` |
| Insert / delete cells, rows, columns | `Ctrl+Shift++` / `Ctrl+-` |
| Insert today's date / current time | `Ctrl+;` / `Ctrl+Shift+;` |
| Show all formulas | `` Ctrl+` `` |
| Go To / Name Manager / Create names from selection | `F5` / `Ctrl+F3` / `Ctrl+Shift+F3` |
| Trace precedents / dependents | `Ctrl+[` / `Ctrl+]` |
| Hide rows / columns | `Ctrl+9` / `Ctrl+0` |
| Next / previous worksheet | `Ctrl+Page Down` / `Ctrl+Page Up` |
| First / last cell of sheet | `Ctrl+Home` / `Ctrl+End` |
| Expand formula bar | `Ctrl+Shift+U` |
| Refresh data in sheet / refresh all | `Ctrl+F5` / `Ctrl+Alt+F5` |
| Insert hyperlink | `Ctrl+K` |
| Ribbon key tips (keyboard-only menus) | `Alt` |
| Macro dialog / VBA editor | `Alt+F8` / `Alt+F11` |

**Mac:** most shortcuts swap `Ctrl` → `⌘` (e.g. `⌘T` Table), but several differ — use the Mac references above rather than muscle memory; Excel for Mac also supports Tools → Customize Keyboard for rebinding commands.

**Cheat sheets:** search the official shortcut pages above with `Ctrl+F`; ExcelJet and Excel Dashboard School also publish compact printable versions (both linked under [Blogs, Videos, and Newsletters](#blogs-videos-and-newsletters)).

## Templates

- [Microsoft Office templates](https://templates.office.com/) - Official budgets, trackers, invoices, and calendars.
- [Vertex42 templates](https://www.vertex42.com/ExcelTemplates/) - High-quality financial, calendar, and business templates with clear formulas.
- [Spreadsheet123](https://spreadsheet123.com/) - Free and premium templates: schedules, lists, calculators.
- **Build your own** - The fastest way to internalize best practices: download a template, then rebuild it from a blank sheet.

## Books

- *Excel 2021/365 Bible* — John Walkenbach (the standard comprehensive reference)
- *Microsoft Excel Data Analysis and Business Modeling* — Wayne Winston (analytics-first, case driven)
- *Excel Pivot Table Data Crunching* — Bill Jelen & Michael Alexander (the definitive PivotTable book)
- *Financial Modeling in Excel For Dummies* — Danielle Stein Fairhurst (modeling from scratch)
- *Financial Modeling* — Simon Benninga (rigorous, finance-faculty standard for valuation and corporate models)
- *Power Pivot and Power BI* — Rob Collie & Avi Singh (data model and DAX for Excel users)
- *Professional Excel Development* — Bovey, Bullen, Green & Wallentin (serious application architecture)
- *Excel VBA Programming For Dummies* — Michael Alexander & John Walkenbach (approachable VBA)
- *Spreadsheet Modeling and Decision Analysis* — Cliff Ragsdale (Excel-based decision/optimization modeling)
- *Data Smart* — John W. Foreman (data science techniques taught through spreadsheets)
- *Excel For Dummies* — Greg Harvey (friendly beginner on-ramp)

## Blogs, Videos, and Newsletters

**Conference**

- [Global Excel Summit](https://www.globalexcelsummit.com/) - Annual online conference: modeling, Power BI, VBA, leadership sessions.

**YouTube channels**

- [Leila Gharani (XelPlus)](https://www.youtube.com/@LeilaGharani) - Business-focused Excel, Power Query, dashboards, Python in Excel; Microsoft MVP.
- [Chandoo](https://www.youtube.com/@chandoo_) - Dashboards, data analysis, Power Query/VBA masterclasses with free files.
- [ExcelIsFun (Mike Girvin)](https://www.youtube.com/@excelisfun) - Enormous, meticulous library: every function, array formula, and Power Query scenario.
- [Excel Campus (Jon Acampora)](https://www.youtube.com/@ExcelCampus) - VBA, shortcuts, and workflow productivity.
- [My Online Training Hub (Mynda Treacy)](https://www.youtube.com/@MyndaTreacy) - Dashboards, Power Query, and DAX explained by a Microsoft MVP.

**Blogs and reference sites**

- [Chandoo](https://chandoo.org/) - Tutorials, courses, and dashboard techniques.
- [ExcelJet](https://www.exceljet.net/) - Clean examples for 500+ functions and shortcuts.
- [MrExcel](https://www.mrexcel.com/) - Tips, solutions, and the long-running community.
- [Trump Excel](https://trumpexcel.com/) - Formulas, charts, VBA, and dashboard tutorials.
- [Ablebits blog](https://www.ablebits.com/office-addins-blog/) - Practical how-tos for everyday Excel problems.
- [The Spreadsheet Guru](https://www.thespreadsheetguru.com/) - Power Query, VBA, and analytics tutorials.
- [Excel Dashboard School](https://www.exceldashboardschool.com/) - Formulas, charting, and dashboard guides.
- [Peltier Tech](https://peltiertech.com/) - Advanced charting and VBA deep dives.
- [Storytelling with Data](https://www.storytellingwithdata.com/) - Data visualization design principles for business charts.

Most of the above run weekly newsletters — subscribe to one (Chandoo or XelPlus are excellent) and read it while you eat lunch.

## Add-ins, Tools, and Libraries

**Productivity add-ins**

- [AppSource — Office add-ins](https://appsource.microsoft.com/en-us/marketplace/apps?product=office) - Marketplace of commercial and free add-ins installed inside Excel.
- [ASAP Utilities](https://www.asap-utilities.com/) - Classic free/paid add-in with hundreds of batch operations.
- [Ablebits Ultimate Suite](https://www.ablebits.com/) - Powerful merge, dedupe, and cleanup tools.
- [Kutools for Excel](https://www.extendoffice.com/) - Large collection of one-click utilities.
- [Power-user](https://power-user.com/) - Reporting, charting, and presentation productivity add-in.
- [think-cell](https://www.think-cell.com/) - The standard charting add-in for consulting and finance decks.

**Built-in tools people forget** (enable via File → Options → Add-ins): **Solver** (optimization), **Analysis ToolPak** (statistics), **Inquire** (workbook audit), **Forecast Sheet**, plus the **Camera tool**, **Watch Window**, and **Accessibility Checker**.

**Alternatives and ecosystems**

- [LibreOffice Calc](https://www.libreoffice.org/) - Free open-source spreadsheet; good for format-agnostic work.
- [Google Sheets](https://sheets.google.com/) - Cloud-native collaboration; learn its equivalents (`QUERY`, `IMPORTRANGE`) if your team mixes tools.
- [Power BI](https://powerbi.microsoft.com/) - The natural next step when Excel reports outgrow the workbook.

**Developer libraries**

- [ClosedXML](https://github.com/ClosedXML/ClosedXML) - Create Excel files from .NET without Excel installed.
- [EPPlus](https://epplussoftware.com/) - .NET read/write/edit library.
- [SheetJS](https://sheetjs.com/) - JavaScript/TypeScript spreadsheet parser and writer.
- [openpyxl](https://openpyxl.readthedocs.io/) - Python library for reading and writing .xlsx files.

## Certification and Career

- [Microsoft Office Specialist: Excel Associate (Microsoft 365 Apps)](https://learn.microsoft.com/en-us/credentials/certifications/mos-excel-associate-m365-apps/) - Industry-standard credential; ~150 hours prep recommended.
- [Microsoft Office Specialist: Excel Expert (Microsoft 365 Apps)](https://learn.microsoft.com/en-us/credentials/certifications/mos-excel-expert-m365-apps/) - Advanced modeling, macros, and workbook customization.
- [Certiport](https://www.certiport.com/) - Official exam delivery partner for MOS exams.
- **Related credential:** [PL-300 (Power BI Data Analyst)](https://learn.microsoft.com/en-us/credentials/certifications/power-bi-data-analyst-associate/) extends the same data modeling skills.

**What the exams actually assess**

| Exam | Domains |
|---|---|
| Associate | Manage worksheets/workbooks; manage cells and ranges; manage tables and table data; formulas and functions; manage charts |
| Expert | Workbook options and settings; manage and format data; advanced formulas and macros; advanced charts and tables |

**Interview prep — questions you should be able to answer cold:**

1. `VLOOKUP` vs `XLOOKUP`: limitations of each, and when `INDEX/MATCH` still wins.
2. Relative vs absolute references; how to fill a formula across a table safely.
3. What is a PivotTable cache, and why can a pivot show stale values?
4. How would you clean a CSV with inconsistent date formats, duplicates, and blank rows?
5. `SUMIFS` vs a PivotTable — when does each perform better?
6. How do you make a report refreshable for a non-technical stakeholder?
7. Difference between a workbook-level named range and a table reference.
8. How do you handle `#N/A`, `#DIV/0!`, and `#REF!` professionally?
9. Explain the difference between calculated columns and measures (data model).
10. Walk through a workbook you built: design, risks, and how you tested it.
11. How do you prevent spreadsheet errors in a model others will rely on?
12. When would you leave Excel entirely (and what would you use instead)?

**Career tips** - Quantify impact ("cut month-end close from 3 days to 4 hours"), publish a small portfolio (anonymized dashboard PDFs + a rebuildable sample file), and get comfortable explaining your model out loud.

## Communities and Forums

- [r/excel](https://www.reddit.com/r/excel/) - Fast, friendly help with solution-point etiquette.
- [Stack Overflow — excel/tagged](https://stackoverflow.com/questions/tagged/excel) - Technical questions, formulas, VBA, Power Query.
- [Microsoft Excel Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/excelgeneral) - Official forum staffed by MVPs and product team members.
- [Microsoft community answers](https://answers.microsoft.com/) - Microsoft's general product Q&A (Excel category).
- [MrExcel Board](https://www.mrexcel.com/board/) - Decades of archived solutions from Excel experts.
- [Excel Forum](https://www.excelforum.com/) - Long-running independent Excel forum.
- [Chandoo forum](https://chandoo.org/forum/) - Dashboards, modeling, and challenge threads.
- [Microsoft MVP program](https://mvp.microsoft.com/) - The community experts behind many of these resources.

## Related Lists

- [awesome](https://github.com/sindresorhus/awesome) - The master list of awesome lists.
- [sancarn/awesome-excel](https://github.com/sancarn/awesome-excel) - Curated Excel resources with paid/free markers.
- [brandonhimpfen/awesome-excel](https://github.com/brandonhimpfen/awesome-excel) - Tools, templates, automation, and learning resources.
- [ebukari/awesome-excel](https://github.com/ebukari/awesome-excel) - Early, compact Excel resource list.

## Glossary

| Term | Meaning |
|---|---|
| Absolute reference | `$A$1` — does not shift when copied |
| Active cell | The selected cell currently receiving input |
| Applied steps | The recorded transformation list of a Power Query |
| Array (dynamic) | A formula returning multiple values that "spill" |
| Calculation column | Formulas auto-filled down a Table column |
| Calculated column | Column in the data model computed row-by-row by DAX |
| Calculated field | PivotTable formula computed over value fields |
| Cardinality | Number of unique values in a column (matters for relationships) |
| Circular reference | A formula that depends, directly or indirectly, on itself |
| Co-authoring | Multiple people editing one workbook simultaneously |
| Criteria | The condition part of `SUMIFS`/`COUNTIFS`/filters (e.g. `">100"`) |
| Data model | In-memory relational model (Power Pivot) hosting tables and relationships |
| Data validation | Rules controlling what may be typed into a cell |
| DAX | Data Analysis Expressions — formula language of the data model |
| ETL | Extract, Transform, Load — Power Query's job |
| Fact table | Row-per-transaction table at the center of a star schema |
| Filter context | The set of filters active when a DAX measure evaluates |
| Flash Fill | Pattern-learning auto-completion (`Ctrl+E`) |
| Grain | What one row in a table represents (e.g. one order line) |
| KPI | Key performance indicator (target vs actual visualization) |
| Measure | DAX aggregate evaluated dynamically in context |
| M language | The scripting language behind Power Query |
| Merged cell | A cell spanning multiple columns — avoid in data |
| Named range | A user-defined name for a cell, range, constant, or formula |
| Pivot cache | The stored copy of source data a PivotTable reads |
| Query folding | Pushing Power Query steps back to the source system |
| R1C1 | Reference style using relative row/column offsets |
| Range | A rectangular block of one or more cells |
| Row context | The "current row" context in calculated columns/iterators |
| Scalar | A single value (as opposed to an array/table) |
| Spill | The range a dynamic array formula automatically fills |
| Spill range | The full output area, addressable with `range#` |
| Star schema | Dimension tables surrounding a fact table |
| Structured reference | Table-style reference: `Table1[Column]` |
| Tidy data | One row per observation, one column per variable |
| Value field | The numeric field a PivotTable aggregates |
| Volatile function | Recalculates on every workbook change |
| What-if analysis | Scenario Manager, Data Tables, Goal Seek, Solver |
| Wildcards | `*`, `?`, `~` used in criteria and Find |
| Workbook / Worksheet | The file / an individual grid inside it |
| XLSB | Binary workbook format (faster for heavy files) |

## Contributing

PRs welcome. To add a resource:

1. Check it isn't already listed and that the link resolves.
2. Place it in the most relevant section; keep descriptions to one line (name + what makes it useful).
3. Prefer free or high-quality resources; note if content is paid.
4. Knowledge additions (tables, checklists) should be factual, version-aware (say which Excel version a feature needs), and non-promotional.
5. Keep alphabetical or thematic order consistent within the section.
6. Open a pull request with a short rationale.

## License

[![CC0](https://mirrors.creativecommons.org/presskit/buttons/88x31/svg/cc-zero.svg)](https://creativecommons.org/publicdomain/zero/1.0/)

To the extent possible under law, this list has waived all copyright and related or neighboring rights to the material in this repository.
