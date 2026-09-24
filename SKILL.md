---
name: nk-model
description: Turn a description of a small business and a few assumptions — price, units sold, costs, the days customers take to pay, equipment, a loan, tax — into a three-statement financial model in one .xlsx: an income statement, a balance sheet and a cash flow that tie to each other, a checks sheet, charts, and a sheet stating every assumption and the method. Use when someone needs to see how a business plan hangs together (when the cash runs out, what a loan or slower-paying customers do to it) or wants a model they can open in a spreadsheet and change. scripts/make_model.py builds the workbook with standard-library Python, every number a live formula with its value stored beside it, and writes nothing unless it balances to the penny; scripts/model_check.py re-checks a workbook from this skill, edited or not, cell by cell against the one it would write: plugs, changed formulas, stored values that disagree, statements that do not tie; scripts/preview.py shows it without a spreadsheet program.
license: MIT
compatibility: standard library only, no packages and no build step. The workbook stores every formula with its value, so a viewer that does not recalculate shows the same numbers; nothing in it runs or links out.
metadata:
  provenance: own practice (2026-09) — the reconciliation discipline of a bank-statement categoriser of mine (exact to the penny, rebuild and compare, a deliberate break must turn red), carried over to a financial model; see Provenance
  version: 0.1.1
---
# Three-statement model

**A business plan is three statements that have to agree.** Profit is on the income statement, cash is on the cash
flow, and the balance sheet is where the two meet: if the model is right, assets equal liabilities plus equity, to the
penny, every year. A model that does not balance has a mistake in it somewhere, and a model that balances only because
someone typed a number into the balance sheet (a plug) is worse, because it looks right. This skill builds the model
from a handful of assumptions, with every number a formula a person can follow, and refuses to hand over one that
does not balance.

> **Paths.** Commands in this skill start with `${…SKILL_DIR}`: this skill's own folder, the one that contains this SKILL.md. Claude Code fills it in. If your agent shows the placeholder as written (Codex, Cursor, Gemini CLI and others), replace it with that folder's absolute path before you run the command. Left as it is, it expands to nothing and the path breaks.

## When this applies

- Someone describes a business — a workshop, a stall, a studio, a small online shop — and asks what it would earn,
  whether it would run out of cash, or what a loan would do.
- A plan needs to be shown to a partner, a lender or a class as a spreadsheet they can open and change.
- Someone has a model already and wants to know whether it holds together: `model_check.py` answers that for a
  workbook this skill wrote, edited or not, cell by cell (what it does not look at is under Boundaries).

## Procedure

1. **Write the business down**: `python3 ${CLAUDE_SKILL_DIR}/scripts/make_model.py --init business.json` gives a
   filled-in example to edit (every field is described in `references/business-format.md`). Ask for the numbers the
   person knows; where they do not know one, say what you assumed and why.
2. **Say whether the numbers are real.** Made-up numbers are marked `"illustrative": true`; real ones need a
   `"sources"` line saying where they come from. The workbook's first sheet repeats it, and says it is not advice.
3. **Build the workbook**: `python3 ${CLAUDE_SKILL_DIR}/scripts/make_model.py business.json -o model.xlsx`. It refuses
   a business that breaks a rule (B01 the file, including text with a control character or an emoji; B02 numbers
   outside sensible ranges, and no amount above a billion; B03 real numbers with no source),
   works the model out twice — once through the workbook's own formulas, once in plain Python without them — and
   writes nothing if the two disagree anywhere or the balance sheet is off by a penny in any year.
4. **Check it**: `python3 ${CLAUDE_SKILL_DIR}/scripts/model_check.py model.xlsx` reads the file back and checks
   eleven things. It builds again the workbook this skill would write for the same years and compares every cell: K01
   the layout (the six sheets, none hidden, every label and line of the method word for word, every cell's number
   format, no cell filled that this skill leaves empty) · K02 no plugs (every statement cell holds exactly this
   skill's formula, nothing is typed where it writes nothing, and the inputs are plain numbers) · K03 every stored
   value is what its formula gives · K04 the balance sheet adds up and balances · K05 the three statements tie · K06
   the inputs are ones the builder accepts, and the numbers follow the stated method · K07 the Checks sheet tells the
   truth · K08 the honesty sentence · K09 no macros, external links, embedded objects or hyperlinks, and no link from
   the package to outside it · K10 the two charts draw the right cells under the right names, and the charts' copies
   of the numbers match · K11 no emoji or icon characters from the builder's set. What it does not look at is under
   Boundaries. Run it again after anyone edits the workbook. It reads the file with
   `${CLAUDE_SKILL_DIR}/scripts/ooxml.py` and evaluates each formula with `${CLAUDE_SKILL_DIR}/scripts/xlformula.py`,
   which follows Excel's own order of operations (`python3 ${CLAUDE_SKILL_DIR}/scripts/xlformula.py "2^3^2"` gives 64).
5. **Read the cash line.** If cash goes below zero at the start or at a year end, both scripts say so: the plan needs
   money the model does not include. That is usually the most useful thing the model says; tell the person in plain
   words.
6. **Look at it without a spreadsheet** if you cannot open one:
   `python3 ${CLAUDE_SKILL_DIR}/scripts/preview.py model.xlsx -o preview.html` shows every sheet as stored, and the
   charts from the numbers they carry.
7. **Hand it over** with the three facts that matter: whether it balances (it always will, or it would not exist),
   when the cash is lowest, and what the model leaves out (it says on the Assumptions sheet).

## Rules that keep it honest

- **No plugs.** Every number on the statements is a formula pointing back at the assumptions or at another
  statement. A number typed into a statement is refused by the checker, even when it is the right number — in a
  cell that should hold a formula, in a cell this skill leaves empty, or written as a formula (=14400.30); so is a
  formula edited so that the totals still come out the same.
- **Balanced to the penny, or nothing.** The builder does not write a workbook that is off by half a penny anywhere,
  and the checker recomputes the balance from the lines themselves, not from the totals, so two totals that are wrong
  by the same amount cannot hide each other.
- **The method is written down and checked.** The Assumptions sheet says how every line is worked out; the checker
  works the model out again from those inputs in plain Python and compares.
- **Stored values are real values.** Each formula's value is stored beside it, so a viewer that does not
  recalculate shows the numbers the formulas give — and the checker recomputes every one. Where this skill's
  arithmetic can differ from Excel's by a hair is under Boundaries.
- **Made up or sourced, and never advice.** The first sheet says which, and says the model is not financial advice.

## Boundaries

- **Yearly, one to ten years, one product line.** Units, a price, a cost per unit and fixed costs, each growing at
  one rate. No months, no seasons, no product mix.
- **One way of doing each thing**: reducing-balance depreciation on equipment, working capital by days, one loan
  repaid in equal yearly amounts with interest on the opening balance, tax only on a profit with no losses carried
  forward, nothing paid out to the owners. Anything else means editing the formulas, and then `model_check.py` will
  say, correctly, that they are no longer this skill's formulas (K02); depending on the change it also says that the
  statements no longer tie (K05, as when money is paid out to the owners) or that the numbers no longer follow the
  stated method (K06, as with straight-line depreciation). Changing the inputs on the Assumptions sheet is the edit
  the model is for.
- **Not advice, and not a forecast of anything.** It shows how the assumptions fit together; whether the assumptions
  are right is the person's to judge.
- **Checked by reading the file, not by Excel.** The checks here read the workbook back with this skill's own reader,
  and the files were also opened by macOS Quick Look and by openpyxl; they have not been opened in Microsoft Excel by
  these checks. Edits are judged cell by cell, as the file stores them; a workbook saved again by a spreadsheet
  program has not been tested, and may be reported as edited where only its storage differs. A shared formula (how
  Excel stores a filled-down run) is read as the formula it stands for when it is stored the usual way; in a group
  where more than one cell carries formula text, this reader and openpyxl can expand a cell differently, so such a
  file can pass here while openpyxl reads other formulas. Fonts named in the file (Fraunces, Inter, Space Mono) fall
  back to the viewer's own where they are not installed.
- **What the checker does not look at.** It compares the cells and their number formats, what the two charts draw and
  their copies of the numbers, and the package's links, macros and embedded objects. An edit anywhere else passes,
  even one that changes what a reader sees: a conditional format that shows a number as fixed text, a third chart, a
  text box laid over a number, a chart's title or its copy of the year labels, hidden rows or columns, merged cells, a
  font colour, a data connection such as a web query (`xl/connections.xml`), text moved into a phonetic guide (read
  here as part of the label; openpyxl leaves it out), Excel's "precision as displayed" setting, or the note under a
  sheet's title, which is free text like the business's name. Open the file and look before it goes anywhere
  (`references/acceptance.md`).
- **Where its arithmetic is not Excel's.** The formula evaluator follows Excel's order of operations and refuses what
  it could not test (`&`, `ROUND`, `TRUE` or text typed into `SUM`, `MIN` or `MAX`, a range inside `AND` or `OR`). It
  does not refuse four cases where its answer differs from what Microsoft documents for Excel, or may: a sum that
  comes out a hair from zero (`1.333 + 1.225 - 1.333 - 1.225` is -2.2e-16 here; Excel 97 and later show 0); text in
  arithmetic (`"1"+"2"` is an error here, 3 in Excel); text reached through a reference inside `AND` or `OR` (an error
  here; Excel ignores it); and a TRUE worked out inside `SUM`, `MIN` or `MAX` (`SUM(1=1)` is 0 here; Microsoft's pages
  do not say). None was tried in Excel. Only the first can come from this skill's own formulas.
- **Cash of exactly nothing at a year end.** The start cash and the check lines are written as 0 when they are within
  half a penny, so owners' money and a loan that exactly pay for the equipment are not called a shortfall. A year-end
  cash that works out to exactly nothing can still be stored a hair below zero (-0.0000000000036), and then the
  Checks sheet and both scripts say the model needs money it does not include. It takes a year whose cash flow comes
  to whole pennies and owners' money set to cancel it exactly; of 7,144 one-year businesses built that way, 63 were
  called short.
- **Text XML cannot hold.** The builder refuses control characters and emoji, but not U+FFFE or U+FFFF: a name or
  description holding one is written into a workbook that is not valid XML, which `model_check.py` and openpyxl
  cannot read. Nor are ⤴ and ⤵ refused, though Unicode lists them as emoji.
- **Very large sums.** Every input can be within its range and the sums still reach the trillions (large amounts
  growing by hundreds of percent a year), where floating point is not exact to the penny. The builder writes such a
  workbook, and the checker, which adds and compounds its own way, can then report a difference at the last penny
  (K05 or K06).
- **Rounded for display.** The statements show money in whole units, and prices, unit costs and the check lines to
  the penny; every total is worked out before rounding, so a total can be a few units off the sum of the lines it
  shows (at most half a unit for each number in the sum, the total included), and units, which grow by a percentage,
  are not whole numbers. Each statement's note says it is rounded, and the Assumptions sheet's method says the rest.
- **The checker knows this skill's layout.** It checks workbooks this skill wrote; a model from anywhere else is
  reported as not from this skill (K01), not checked.

## Provenance

Own practice, 2026-09. The discipline comes from a bank-statement categoriser of mine: totals that must reconcile to
the penny, output that is rebuilt and compared rather than trusted, and a deliberate break that has to turn a check
red before the check counts. This skill carries it to a financial model and keeps none of that code; the workbook
writer, the formula checker and the model are new here. The rule against plugs, and the decision to recompute the
balance from the lines instead of from the totals, come from how a model most often looks right while being wrong. The
colours come from the token file shared with the other work in this family.
