# Changes

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
Excel opening, native recalculation, chart refresh and visual fit remain untested. Integration provides a bundled
preflight request and relative printed next-step paths; the description states when to use the skill and preserves
compatibility wording in metadata for frontmatter validators.

An outsider's first read demonstrated that undeclared monthly-cash prose can pass and negated seasonality can
be refused. This text-only fix distinguishes agent classification from declared-feature/limited-wording checks,
corrects the two Summary rows, half-penny funding boundary and display/zero normalization wording, and explains
the example's aggregation and supplied provenance. Opening in Excel itself remains unchecked; scripts and assets
are unchanged.
