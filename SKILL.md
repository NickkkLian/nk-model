---
name: nk-model
description: "Use when a business needs annual three-statement Excel models for one sales line, or subscriptions with uniform monthly churn and constant monthly additions. Assess and classify every requirement before translating assumptions; agents must refuse unsupported needs, while code checks declared features and limited wording patterns. Package selects Summary with profit, cash, funding warnings and charts; opening in Excel itself is untested. Standard-library Python, live formulas, stored values and an offline preview."
license: MIT
metadata:
  compatibility: "Python 3.9+ standard library only; no installs, network, macros or external links needed."
  version: 0.2.0
---
# Three-statement model

Use for annual single-line businesses with units, price, variable and fixed costs growing at constant annual rates.
Also supports subscriptions with one uniform monthly churn rate and constant additions at month end.
Equipment uses reducing balance or straight line over a whole-year useful life, with zero residual value.
Statements and cash checks are annual. Monthly subscription revenue is exact for the stated convention;
monthly cash is not computed. Opening and year-end cash can miss any size of intra-year shortfall.

The agent must classify every material requirement and refuse cohort-specific churn, seasonality, monthly cash
forecasts, multiple products, prepaid/deferred revenue, custom depreciation, used assets, partial years, residual
values, multiple or nonstandard loans, opening working capital, owner distributions and tax loss carryforwards.
Code validates declared features and limited wording patterns, not arbitrary prose: an unsupported need written
differently can pass when classification omits it. An outsider's monthly-cash prose request passed with only annual
features declared; a sentence denying seasonality was refused by the wording check. Keep original prose and
disclose classification; do not remove words to force acceptance. Resolve unknown requirements before building.
A model that balances can still describe the wrong business.

> **Paths.** `${…SKILL_DIR}` means this skill's own folder. Claude Code fills in `${CLAUDE_SKILL_DIR}`; another
> agent must replace it with this folder's absolute path before running the commands below.

## Procedure

1. **Assess the original description before translating numbers.** Read `references/business-format.md`. Write
   `request.json` with the original `description` and a `scope` containing `features` and plain-word `notes`.
   Classify every material requirement, including timing and omitted inputs. Run:
   `python3 ${CLAUDE_SKILL_DIR}/scripts/make_model.py --preflight request.json`.
   Show its fit assessment to the user before creating `business.json`. Unsupported or unknown features return B04
   and no workbook. Code validates declared features and a few explicit prose patterns; it cannot understand arbitrary
   prose. You must not omit requirements to obtain a pass or flatten monthly churn into annual growth.
2. **Resolve mismatches before translating.** Explain what is unsupported and what would change. The code refuses declared
   unsupported features; it does not offer a bounded approximation for them. Uniform churn can differ without a
   finite bound from an arbitrary cohort-retention curve; annual cash can miss any size of seasonal funding need.
   Do not build the original request in those cases. A user-authorized simpler request needs a new preflight and
   an explicit record of the change. Existing authorization to simplify is sufficient; do not ask again.
3. **Translate an accepted request.** `python3 ${CLAUDE_SKILL_DIR}/scripts/make_model.py --init business.json`
   writes a filled example. Replace all example inputs, copy assessed `scope`, retain the original `description`, and
   write a `business_line` of at most 110 characters as input metadata. Summary generates its sales description
   from the selected model rather than copying that prose. List every missing-number assumption and why in
   `scope.notes`, retained on Assumptions. In subscription mode set
   unused annual units and unit growth to zero. Bill opening-month customers, apply churn to them, then add customers
   at month end. New customers first pay next month. Price and unit cost change only at year boundaries. Expected
   customer counts may be fractional.
4. **State provenance.** Invented scenarios use `illustrative: true`. Real inputs use `illustrative: false` and a
   `sources` sentence. Summary and Assumptions repeat that supplied provenance. The code checks its presence
   and consistency, not the authenticity of inputs or sources. This is not financial advice.
5. **Build and check**, using separate commands:

   ```bash
   python3 ${CLAUDE_SKILL_DIR}/scripts/make_model.py business.json -o model.xlsx
   python3 ${CLAUDE_SKILL_DIR}/scripts/model_check.py model.xlsx
   python3 ${CLAUDE_SKILL_DIR}/scripts/preview.py model.xlsx -o preview.html
   ```

   B01 validates schema/types, duplicate keys and XML-safe text; B02 ranges; B03 provenance; B04 scope and method
   agreement. Old inputs without scope and the new timing/method fields are refused; migrate explicitly. The builder
   evaluates formulas and compares selected statement values against a separate plain-Python calculation inside
   the same builder; this does not independently check every monthly or Summary cell. It writes nothing if compared
   values differ by half a penny or more, or accounting checks fail.
6. **Read Summary first.** It is selected on open; Assumptions comes second. It shows the name in row 1 and modeled sales in row 2,
   revenue, net profit, year-end cash, lowest opening/year-end cash and earliest matching date, minimum additional
   funding, applicable timing/method descriptions and two charts. These descriptions are live formulas; absent
   equipment and loans have blank descriptions. Opening equipment is distinguished from recurring year-end purchases.
   Read the original request and all scope notes on Assumptions, and disclose them at handover. Tested cash below −0.005 currency units puts “Funding needed: cash
   below zero” at the top. Funding is `-lowest tested cash` when that cash is below −0.005, otherwise zero:
   a lower bound at tested points, not a monthly forecast or an automatic borrowing plug. Shortfalls of exactly
   half a penny or less count as zero funding and do not trigger this warning.
7. **Hand over workbook and preview.** State lowest tested cash, date, funding need and all simplifications. Follow
   `references/acceptance.md` and report native-application checks actually performed. Change numeric assumptions
   to model alternatives. Editing formulas is outside the supported model.

Each of the six scripts accepts `--selftest`. The number-source check is
`python3 ${CLAUDE_SKILL_DIR}/scripts/numsrc.py check preview.html`; clicking a number shows its cell, formula and
verification limits. `${CLAUDE_SKILL_DIR}/scripts/ooxml.py` writes and reads workbook packages;
`${CLAUDE_SKILL_DIR}/scripts/xlformula.py` evaluates the supported formulas. Their self-tests check the package
reader/writer and formula semantics separately from the builder and checker suites.

## Methods and limits

- One to ten annual periods; one sales line. Subscription churn and additions stay constant across the horizon.
- Straight line tracks opening cost and equal yearly purchases by age: opening assets start in year one; year-end
  purchases start next year; depreciation stops after useful life. Zero residual; new equipment only. Reducing balance
  applies a rate to opening net equipment and does not promise full write-off by a fixed life.
- Receivables = annual revenue × days / 365. Inventory and payables = annual cost of sales × days / 365. Opening
  balances are zero. This is an annual convention, not a procurement or collections schedule.
- One opening loan, equal annual principal, interest on opening balance. Tax on positive annual pretax profit only;
  no loss carryforward or deferred tax, no owner payouts, no automatic funding of negative cash.
- Calculated statement, monthly-schedule and Summary numbers are formulas with stored values. Assumptions contain
  editable numbers. Chart references stay live; chart caches support viewers without recalculation.
- Summary business-method sentences are generated from numeric inputs and statement flows. Free-text business
  and scope descriptions are not copied into those sentences. Provenance is repeated from Assumptions.
- Statement totals display whole currency units; prices and unit costs display two decimals and Checks show
  pennies. Computations generally retain precision, but near-zero opening cash and accounting-check differences
  with absolute value below 0.005 are normalized to zero. Arithmetic acceptance tolerance
  is absolute error less than 0.005 currency units for identities and supported business calculations, not forecast
  accuracy. Inputs are capped at a billion; compounded sums can exceed float precision, causing builder/checker
  disagreement. Resolve any finding before handover.
- K01–K11 check layout/formats, exact formulas/no plugs, caches, balance-sheet sums, statement ties, methods,
  numeric inputs/scope, truthful Checks/Summary, provenance, unsafe links/packages, chart references/caches and the
  configured icon set. The checker rebuilds a reference using builder code: it is not an independent business-method
  oracle. Use a separate calculation to verify the description.
- Presentation is not fully audited: hidden rows/columns, extra charts, drawings/text boxes, conditional formats,
  merges, fonts/colors, chart titles/year caches, data connections and precision-as-displayed can escape checks.
  Descriptions and scope notes are free text with safe-character and scope checks. Look at the file.
- The included formula evaluator is not Excel. Near-zero cancellation, text arithmetic/logical functions and logical
  results inside SUM/MIN/MAX can differ, as described in `${CLAUDE_SKILL_DIR}/scripts/xlformula.py`. Preview and caches do not establish
  native Excel recalculation. This upgrade has not opened or saved the build in Microsoft Excel.
- Spreadsheet-program saves can change storage in ways the checker flags. Shared formulas work in the usual
  one-master form; ambiguous multi-master groups are not reliable. Fonts fall back where not installed.

## Provenance

The author's reconciliation approach comes from bank-statement work: no plugs, add the lines independently and
prove deliberate breaks are detected. See `references/invariants.md` for the accounting identities.
