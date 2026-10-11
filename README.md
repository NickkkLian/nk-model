# nk-model

Version **0.2.0**.

For owners and analysts using a coding agent to make an annual Excel income statement, balance sheet and cash
flow for **one sales line**, or subscriptions with **uniform monthly churn and constant month-end additions**. Supports straight-line equipment depreciation over a whole-year useful
life with zero residual, or reducing balance. The package selects Summary first: revenue, profit, cash, minimum funding at tested dates,
warnings and two charts; Assumptions is second. Opening in Excel itself was not checked.

![nk-model: the preview of a workbook's Summary sheet: funding status first, then revenue, net profit and cash by year](https://raw.githubusercontent.com/NickkkLian/nickkk-skills/main/gallery/results/nk-model.png)

Shown in the skill's own HTML preview page.

The agent must classify **every material requirement** and refuse cohort-specific churn, seasonality, monthly
cash forecasts, product mix, prepaid/deferred revenue, custom depreciation, opening working capital and
nonstandard financing. The code validates the **declared features and a limited set of wording patterns**;
an unsupported need expressed in other words can pass if it is omitted from the declaration. In the outsider's
probes, a monthly-cash request declared only as annual sales passed preflight, build and checker, producing an
annual workbook; a sentence denying seasonality was refused because the wording check did not understand
negation. Preserve the original prose, disclose the classification and resolve mismatches rather than deleting
words to obtain a pass.

Working capital uses annual revenue for receivables and annual cost of sales for inventory/payables, each
multiplied by its days / 365. Cash checks cover opening and **year-end** only: reported funding covers those
tested dates and can miss any size of intra-year shortfall. The agent procedure requires fit assessment before
assumption translation; the CLI does not enforce that history. This is not financial advice.

## Build a model

Read `SKILL.md` and `references/business-format.md`. Classify original requirements into `request.json` before
translating assumptions. Preflight validates declared capabilities and a few prose patterns; the agent must assess
everything else honestly. The code refuses unknown or unsupported feature names; the agent must identify and
refuse unsupported requirements that the limited prose checks miss.

The bundled illustrative bike-repair/parts example aggregates repairs and counter parts into **one average
sales unit**: 2,400 units in year one, £45 revenue and £18 variable cost per unit. It does not model separate
repair and counter-parts quantities, prices or margins. A request needing those separate lines must be classified
as `multiple_products` and refused; accepting this aggregate requires explicit agreement to that simplification.

For this bundled example, run from the skill folder:

```bash
cp assets/requests/bike-repair.json request.json
python3 scripts/make_model.py --preflight request.json
python3 scripts/make_model.py --init business.json
python3 scripts/make_model.py business.json -o model.xlsx
python3 scripts/model_check.py model.xlsx
python3 scripts/preview.py model.xlsx -o preview.html
```

![nk-model](https://raw.githubusercontent.com/NickkkLian/nickkk-skills/main/gallery/social/nk-model.png)

Replace all example inputs. Old inputs need explicit migration to scope, sales timing and depreciation fields. No
network, installs or spreadsheet application is required to build/check files. Python 3.9+ standard library only.

Every calculated statement, monthly-schedule and Summary number is a formula with its value stored. The builder
compares selected statement values against a second plain-Python calculation inside the same builder, refusing
differences of half a penny or more; it does not independently verify every monthly or Summary cell. The balance
sheet balances at opening and every year, cash ties to cash flow, and profit reaches equity. Formulas are intended
to update dependent statements after numeric edits and recalculation; native Excel recalculation was not checked.
Statement totals display whole currency units; prices and unit costs display two decimals and Checks show pennies.
Computations generally retain precision, but near-zero opening cash and accounting-check differences with absolute
value below 0.005 are normalized to zero.

The user or agent supplies `illustrative` and, for real inputs, a `sources` sentence. The workbook repeats that
supplied provenance; the code validates its presence and consistency, not whether the numbers or sources are authentic.

Subscriptions bill opening-month customers, apply uniform churn, then add customers at month end. New customers
pay next month; fractional customers are expectations. Straight line starts opening new assets in year one and
year-end purchases the following year; depreciation stops after useful life. One opening loan uses equal annual
principal and interest on opening balance. Tax is on positive annual profit, with no loss carryforward. These
conventions appear in the workbook.

Summary has the business name in row 1 and the modeled sales description in row 2. Its method descriptions
are formulas based on live inputs: equipment and loan descriptions appear only when funded amounts exist,
and start-of-year equipment is distinguished from recurring year-end purchases. Their native Excel updates
were not checked. The sales description is generated from the selected model. Original description, defaults
and scope notes remain on Assumptions.

Funding and its warning use the lowest opening/year-end cash: below −0.005 currency units, funding is the
absolute shortfall; otherwise it is zero. A shortfall of exactly half a penny, or less, produces zero funding
and no funding warning. This threshold does not measure within-year funding needs.

## Check before sharing

```bash
python3 scripts/make_model.py --selftest
python3 scripts/model_check.py --selftest
python3 scripts/numsrc.py --selftest
python3 scripts/ooxml.py --selftest
python3 scripts/preview.py --selftest
python3 scripts/xlformula.py --selftest
```

All six existing suites pass in 0.2.0. The HTML preview is offline and shows stored values; click a number for its
cell and formula. It does not recalculate. The checker uses the shipped evaluator and rebuilds a reference from
builder code, so a separate business-method calculation is necessary.

## How this version was tried

Two agents who did not build the skill each chose three business requests. The first tried a subscription
business, a shop and a consulting firm: all three workbooks balanced, but the report found the subscription's
churn and consulting equipment's write-off did not match the requests, and profit/cash warnings were buried.
The build adds preflight scope checks, explicit uniform monthly subscription and whole-year straight-line
methods, and a first-sheet Summary with funding warnings.

The second agent tried a subscription business, a laundry and seasonal kayak rentals. The agent reported correct
finances in the first two and an honest refusal for the seasonal/monthly-cash request, but Summary words
described nonexistent equipment/loans or the wrong purchase timing. The builder then made those descriptions
follow live inputs and added checks. Detailed inputs, workbooks and logs for that second trial were not supplied
with this integration; its findings are recorded reports, not newly reproduced calculations. Fresh local checks
below run the bundled example and all six self-tests. Opening, saving, recalculating, refreshing charts and
checking visual fit in Excel itself were not checked.

<details><summary>Complete command output recorded on 2026-10-11</summary>

The copy command prints nothing. Each block is the full combined output of its stated command, run on a fresh
copy of the adopted lab source. The bundled inputs are illustrative; replace them and assess fit for your request.

`cp assets/requests/bike-repair.json request.json` (exit 0):

```text
```

`python3 scripts/make_model.py --preflight request.json` (exit 0):

```text
FIT assessed before build:
  annual_sales: One annual units/price line; constant annual growth rates.
  reducing_balance: Opening net equipment times a constant annual depreciation rate; does not imply full write-off by a stated life.
  working_capital_days: Year-end receivables, inventory and payables by annual flows times days / 365; opening balances zero.
  equal_principal_loan: One opening loan, equal annual principal, interest on opening balance.
  Annual single-line planning. New equipment bought at year end; cash not scheduled within a year.
  WARNING: cash tested only at opening and year ends. Intra-year shortfalls are unknown and have no bounded error.
  WARNING: tax on annual profit only; no loss carryforward. Disclose tax assumptions before translation; actual-tax error is unknown.
Proceed only if these methods and timing match the original request. No workbook written.
```

`python3 scripts/make_model.py --init business.json` (exit 0):

```text
wrote business.json: an illustrative bike-repair business to edit
```

`python3 scripts/make_model.py business.json -o model.xlsx` (exit 0):

```text
FIT assessed before build:
  annual_sales: One annual units/price line; constant annual growth rates.
  reducing_balance: Opening net equipment times a constant annual depreciation rate; does not imply full write-off by a stated life.
  working_capital_days: Year-end receivables, inventory and payables by annual flows times days / 365; opening balances zero.
  equal_principal_loan: One opening loan, equal annual principal, interest on opening balance.
  Annual single-line planning. New equipment bought at year end; cash not scheduled within a year.
  WARNING: cash tested only at opening and year ends. Intra-year shortfalls are unknown and have no bounded error.
  WARNING: tax on annual profit only; no loss carryforward. Disclose tax assumptions before translation; actual-tax error is unknown.
wrote model.xlsx: Harbour Bike Repairs, 2027-2029 · 19,673 bytes · balanced to the penny in all 4 balance sheets
net profit 7,614, 15,452, 24,545 · cash at the end of each year 14,400, 26,518, 47,268 (£)
next: python3 scripts/model_check.py model.xlsx
```

`python3 scripts/model_check.py model.xlsx` (exit 0):

```text
✔ model.xlsx: 0 findings · balanced to the penny, every formula's stored value recomputed
```

`python3 scripts/preview.py model.xlsx -o preview.html` (exit 0):

```text
wrote preview.html: 7 sheets; open it, or add #s2, #s3 ... to the address for the others
```

`python3 scripts/make_model.py --selftest` (exit 0):

```text
  PASS  the example business meets every rule
  PASS  the example balances to the penny and the direct working agrees
  PASS  the checks sheet says Balanced
  PASS  Summary: business with no equipment and no loan does not describe either
  PASS  Summary: equipment bought at the start of year one does not imply year-end purchases
  PASS  every line this skill writes in the Assumptions sheet's first column is at most 138 characters
  PASS  a business that loses money every year: balanced and agreed
  PASS  a business with no loan: balanced and agreed
  PASS  ten years and a loan repaid in two: balanced and agreed
  PASS  no stock, instant payment: balanced and agreed
  PASS  a formula bent by one percent is caught: the model no longer balances or agrees
  PASS  the example writes a workbook
  PASS  seven sheets, in order: Summary, Assumptions, Income statement, Balance sheet, Cash flow, Checks, Monthly sales
  PASS  400 days to be paid is refused and nothing is written
  PASS  real numbers with no sources are refused
  PASS  a key written twice is refused
  PASS  equity and loan that exactly cover the equipment: the start cash is 0, and nothing says the model needs more money
  PASS  owners' money of 27,567.20, so the year ends with exactly nothing: neither the note nor the Checks sheet says it needs money, and the cash prints as 0, not -0 (Checks!B11 'No')
  PASS  owners' money of 27,567.19, so the year ends with a penny short: both the note and the Checks sheet say it needs money (Checks!B11 'Yes: the model needs money it does not include')
  PASS  a star in the name is refused, and nothing is written
  PASS  a bell character in the name is refused, and nothing is written
  PASS  fixed costs of ten trillion is refused, and nothing is written
  PASS  a 400-digit number of units is refused, and nothing is written
make_model selftest: 23/23 passed
```

`python3 scripts/model_check.py --selftest` (exit 0):

```text
  ✔ clean → want ['clean'], got ['clean']
  ✔ K01 a sheet renamed → want ['K01'], got ['K01']
  ✔ K01 a row label changed → want ['K01'], got ['K01']
  ✔ K01 an eighth sheet added → want ['K01'], got ['K01']
  ✔ K01 eleven years (the skill writes 1 to 10) → want ['K01'], got ['K01']
  ✔ K01 no year columns on the income statement → want ['K01'], got ['K01']
  ✔ K01 an assumption's label changed → want ['K01'], got ['K01']
  ✔ K01 year headers that disagree between statements → want ['K01'], got ['K01']
  ✔ K02 a number typed in where a formula belongs (a plug) → want ['K02'], got ['K02']
  ✔ K02 a formula outside what the checker can verify → want ['K02'], got ['K02']
  ✔ K03 a stored value that is not its formula's result → want ['K03'], got ['K03']
  ✔ K03 a formula that divides by zero, with a number stored beside it → want ['K02', 'K03'], got ['K02', 'K03']
  ✔ K04 a total that leaves a line out → want ['K02', 'K04'], got ['K02', 'K04']
  ✔ K04 two totals out by the same amount, so the difference row still reads zero → want ['K02', 'K04'], got ['K02', 'K04']
  ✔ K04 money owed by customers at the start, from nowhere → want ['K02', 'K04'], got ['K02', 'K04']
  ✔ K05 two errors that cancel: the cash flow's profit is wrong and its stock line hides it → want ['K02', 'K05'], got ['K02', 'K05']
  ✔ K06 interest at twice the stated rate, everything else consistent → want ['K02', 'K06'], got ['K02', 'K06']
  ✔ K07 a Checks sheet that says NOT balanced about a balanced model → want ['K02', 'K07'], got ['K02', 'K07']
  ✔ K07 a Checks sheet that says cash never goes below zero when it does → want ['K02', 'K07'], got ['K02', 'K07']
  ✔ K08 the not-advice sentence removed → want ['K08'], got ['K08']
  ✔ K09 a macro in the package → want ['K09'], got ['K09']
  ✔ K09 a hyperlink in a sheet → want ['K09'], got ['K09']
  ✔ K10 a chart whose copy of the numbers is stale → want ['K10'], got ['K10']
  ✔ K10 a chart series that points at no cells → want ['K10'], got ['K10']
  ✔ K11 an icon in the description → want ['K06', 'K11'], got ['K06', 'K11']
  ✔ a business whose cash goes below zero: no finding, one warning → want ['clean'], got ['clean']
  ✔ a business whose start cash is exactly nothing: no finding and no warning → want ['clean'], got ['clean']
  ✔ a business whose year-end cash is exactly nothing: no finding and no warning → want ['clean'], got ['clean']
  ✔ the same business a penny short at the year end: no finding, one warning → want ['clean'], got ['clean']
  ✔ K01 the Checks sheet hidden → want ['K01'], got ['K01']
  ✔ K01 losses shown without their brackets → want ['K01'], got ['K01']
  ✔ K01 a line of the method rewritten → want ['K01'], got ['K01']
  ✔ K01 a formula added beside the inputs → want ['K01'], got ['K01']
  ✔ K02 a plug written as a formula → want ['K02'], got ['K02']
  ✔ K02 opening balances typed into the blank start cells → want ['K02', 'K03', 'K04', 'K05'], got ['K02', 'K03', 'K04', 'K05']
  ✔ K02 an input written as a formula → want ['K02'], got ['K02']
  ✔ K02 formulas edited so every total stays the same → want ['K02'], got ['K02']
  ✔ clean: a run of formulas stored the way Excel shares them → want ['clean'], got ['clean']
  ✔ K06 an input the builder would refuse (depreciation at 150%) → want ['K06'], got ['K06']
  ✔ K06 the loan's years typed as 0 → want ['K03', 'K06'], got ['K03', 'K06']
  ✔ K09 a link on a chart that points outside the file → want ['K09'], got ['K09']
  ✔ K10 a chart's two series under each other's names → want ['K10'], got ['K10']
  ✔ K01 a heading replaced by a number → want ['K01'], got ['K01']
  ✔ K10 a chart missing from the file → want ['K10'], got ['K10']
  ✔ K11 an emoji punctuation mark in the description → want ['K06', 'K11'], got ['K06', 'K11']
model_check selftest: 45/45 passed
```

`python3 scripts/numsrc.py --selftest` (exit 0):

```text
selftest: 50/50
```

`python3 scripts/ooxml.py --selftest` (exit 0):

```text
  PASS  column letters: 1 A, 26 Z, 27 AA, 52 AZ, 703 AAA
  PASS  letters and numbers round-trip for the first 1999 columns
  PASS  sheet names are quoted in formulas when they need it
  PASS  numbers are stored in their shortest exact form
  PASS  the same workbook saved twice gives the same bytes
  PASS  a formula is stored with its value, and both read back
  PASS  text (including empty text) and true/false results of formulas read back
  PASS  plain text and numbers read back, with no formula
  PASS  all 14 parts are well-formed XML
  PASS  the charts are in the package and carry a cache of the numbers they draw
  PASS  every xl/ part has a content type
  PASS  Excel is told to recalculate everything when it opens the file
  PASS  a sheet name Excel would refuse is refused here
ooxml selftest: 13/13 passed
```

`python3 scripts/preview.py --selftest` (exit 0):

```text
  PASS  brackets for a negative, the literal for zero
  PASS  percent, thousands and decimals
  PASS  the check format: 0.00, and a real difference keeps its sign
  PASS  one section per sheet
  PASS  a stored value appears as the sheet formats it (net profit 7,614)
  PASS  both charts are drawn, bars and a line
  PASS  the page fetches nothing
  PASS  the three faces the workbook names travel inside the page
  PASS  each inlined face is its whole font file
  PASS  the fonts' licence notice travels with them
  PASS  the title cell's font survives inside its style attribute (Fraunces,Georgia,'Times New Roman',serif)
  PASS  every number on the page has a source the shared checker accepts ([])
  PASS  net profit's source is its own cell and formula (=B14-B15)
  PASS  a number that loses its source is refused by the shared checker (N06)
  PASS  the revenue bars are drawn from the chart's own copy, which holds the stored revenue (3 years)
  PASS  six bars: two series over three years (6)
preview selftest: 16/16 passed
```

`python3 scripts/xlformula.py --selftest` (exit 0):

```text
  PASS  Excel's order: ^ left to right, a leading minus before ^
  PASS  SUM over a range
  PASS  references to other sheets, quoted and not
  PASS  absolute and mixed references
  PASS  MAX and MIN
  PASS  IF with text results
  PASS  AND and comparisons
  PASS  an empty cell counts as 0
  PASS  refused: an unknown function, ROUND, a defined name, a bare range, division by zero, an unfinished call, two values in a row
  PASS  no answer rather than a guess: 0^-1, 0^0, 10^400, (-8)^(1/3), &, TRUE or text typed into MAX/SUM, a range in AND, text as IF's condition
  PASS  comparisons as Excel makes them: types never equal, number < text < logical, an empty cell matches "", 0 and FALSE
  PASS  SUM adds left to right, as + does (1e16 + 1 - 1e16 is 0 both ways)
  PASS  in a reference, TRUE is skipped by MAX and an empty cell by AND, as Excel does
  PASS  the cells a formula reads, ranges expanded
xlformula selftest: 14/14 passed
```

</details>

## What was not checked

- Opening or saving this build in Microsoft Excel, repair-free opening, native Excel recalculation, and chart refresh
  after native input edits. See `references/acceptance.md` for those checks.
- Other spreadsheet engines and arbitrary prose comprehension.
- Formatting changes, hidden rows/columns, text boxes, extra charts, chart titles/year caches, data connections and
  precision-as-displayed are not fully inspected by the checker.
- Within-year cash and funding needs. Annual tested cash points can miss any size of intra-year shortfall.
- Independent confirmation of the earlier hand-back's test history and historical before/after figures. The supplied
  Round 2 report does not include the audit/logs it mentions. Full visual fit and clipping in Excel remain unchecked;
  the offline raster renderer's width/height heuristic is not a complete visual audit.

No GUI automation was used. Recalculation evidence comes from Python, not Excel itself. Large compounded sums can
exceed floating-point penny precision. Fonts fall back where not installed. This shows how stated assumptions fit
together, not whether they are realistic.
