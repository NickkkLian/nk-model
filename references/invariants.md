# The five invariants

A three-statement model is three views of one business that have to agree. The ways it goes wrong are quiet: a
number typed into the balance sheet to make it balance, a total that leaves a line out, a formula changed in one
year and not the next, a stored value that no longer matches its formula. Each one produces a workbook that looks
finished.

These five hold for every workbook from this skill. The **Checked by** column names the rules: B rules are the
builder's (`make_model.py`, before it writes), K rules the checker's (`model_check.py`, on any workbook).

| # | Invariant | How you can tell by looking | Checked by |
|---|---|---|---|
| 1 | **No plugs.** Every number on the income statement, the balance sheet, the cash flow and the checks is a formula that points back at the Assumptions sheet or at another statement. | Click any number on a statement: the formula bar shows a formula, never a typed value. The pink cells on the Assumptions sheet are the only inputs. | K02 |
| 2 | **The balance sheet balances to the penny, at the start and every year.** Total assets equal liabilities plus equity, and each total is the sum of its own lines. | The last row of the balance sheet, and every line of the Checks sheet, read 0.00. | K04, K07 |
| 3 | **The three statements tie.** Net profit and depreciation carry to the cash flow; the working-capital lines are the balance sheet's changes; closing cash is the balance sheet's cash; profit kept, equipment, the loan and owners' money roll forward from one year to the next. | On the cash flow, "Cash at the end of the year" equals "Cash" on the balance sheet for the same year. | K05, K07 |
| 4 | **The numbers follow the stated method.** The Assumptions sheet says how every line is worked out, and the statements do exactly that. | Change one input (a price, a rate, a number of days) and the effect runs through every statement that depends on it — a price reaches all three; the days customers take to pay reach the balance sheet and the cash flow, not profit — and the Checks sheet still says Balanced. | K06, B02 |
| 5 | **Stored values are real values, and the file says what it is.** Every formula's value is stored beside it, so a viewer that does not recalculate shows the numbers the formulas give; the first sheet says whether the numbers are made up or where they come from, and that the model is not advice; nothing in the file runs or links out. | Open the file in a previewer that does not recalculate: the numbers are there. Read the third line of the Assumptions sheet. | K03, K08, K09, K10 |

## What is not in the family

- **A balancing figure.** No "suspense", "adjustment" or "other" line that absorbs a difference. If it does not
  balance, something is wrong, and the builder does not write it.
- **A forecast presented as fact.** The model shows how the assumptions fit together. Made-up numbers are marked as
  made up; real ones say where they come from.
- **Hidden inputs.** Every input is on the Assumptions sheet, coloured, with its unit. A number that changes the
  model and lives anywhere else is a plug.
