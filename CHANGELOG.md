# Changes

## 0.2.1 — 2026-10-14

README now leads with an example request, the resulting workbook and a real K03 wrong-cell finding (checker exit 1).
Its first picture is a local capture of the bundled example's generated Summary preview, replacing the Assumptions
picture whose caption promised Summary. Scope limits and prior command evidence remain below the result.

README records the dated manual Excel opening report once; current skill wording no longer says opening is untested.
Saving, native recalculation after input edits, chart refresh and full visual fit remain unchecked. The checker
explanation now identifies who must independently verify that the business method fits the original description.
Current skill, builder, preview-generator and business-format versions are 0.2.1; model calculations are unchanged.
All six self-test suites, bundled-example preflight/build/check/preview and preview number-source checks pass before
and after; the stored-value edit to Income statement!B4 is refused with four K03 findings in both runs.

## 0.2.0 — 2026-10-11

A first-time user's run found workbooks that balanced but represented uniform subscription churn and equipment
write-off incorrectly, and hid a cash shortfall behind the opening Assumptions sheet. Scope is now assessed before
translation: declared unsupported business methods are refused, while uniform monthly subscription churn/additions and
whole-year straight-line equipment depreciation have explicit conventions and checks.

Summary now opens first with revenue, profit, cash, minimum funding at opening/year-end points, warnings and two
charts. A second first-time user's run reported correct finances but misleading Summary descriptions of equipment
and loans. Summary method descriptions now follow live inputs, omit absent equipment/loans and distinguish
opening purchases from recurring year-end purchases; the checker verifies their formulas and stored text values.

The README adopts the replacement text and records fresh local commands and the two three-request agent trials.
Excel opening, native recalculation, chart refresh and visual fit were untested at that integration. Integration provides a bundled
preflight request and relative printed next-step paths; the description states when to use the skill and preserves
compatibility wording in metadata for frontmatter validators.

An outsider's first read demonstrated that undeclared monthly-cash prose can pass and negated seasonality can
be refused. This text-only fix distinguishes agent classification from declared-feature/limited-wording checks,
corrects the two Summary rows, half-penny funding boundary and display/zero normalization wording, and explains
the example's aggregation and supplied provenance. Opening in Excel itself was still unchecked at that time;
scripts and assets were unchanged.
