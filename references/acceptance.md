# Acceptance: what a person checks before the model goes anywhere

The checker reads the file and recomputes it. These are things it cannot see; SKILL.md's Boundaries also list the
edits it does not look at, such as a hidden row or a text box laid over a number. Walk them in order with the
workbook open; each row says what to do and what counts as a pass. The **Rule** column names the machine rule that
covers part of the same ground, where there is one.

| # | Do this | Passes when | Rule |
|---|---|---|---|
| 1 | Read Summary before scrolling. | Business, annual revenue/profit/cash, lowest tested cash/date, minimum funding and warnings are visible. Assumptions is second. | K01, K07 |
| 2 | Read each input next to its label. | Each is a number the person gave you or one you can say why you chose. None is a placeholder. | — |
| 3 | Read the third line of the Assumptions sheet. | It says the numbers are made up, or where they come from, and that this is not financial advice. | K08 |
| 4 | Click three numbers on each statement. | Each shows a formula in the formula bar, never a typed number. | K02 |
| 5 | Read the last row of the balance sheet and every line of the Checks sheet. | All read 0.00, and the Checks sheet says Balanced. | K04, K07 |
| 6 | Put "Cash at the end of the year" from the cash flow beside "Cash" on the balance sheet. | They are the same number in every year. | K05 |
| 7 | Change one input — a price, the days customers take to pay — and watch the three statements. | The change runs through each statement that depends on it (a price reaches all three; the days customers take to pay reach the balance sheet and the cash flow, not profit), and the Checks sheet still says Balanced. Change it back afterwards. | K06 |
| 8 | Look at lowest opening/year-end cash. | If below −0.005 currency units, Summary says Funding needed at the top, with amount/date; shortfalls of exactly half a penny or less give zero funding/no warning. Explain intra-year cash is unknown. | K07 |
| 9 | Read the "How the model works" lines. | They describe what this business does — or you have told the person which of them does not fit (a business with seasons, stock bought in bulk, several loans). | — |
| 10 | Look at the two Summary charts. | Revenue, net profit and year-end cash match statements; charts refresh after recalculated edits in the intended app. | K10 (file caches only) |
| 11 | Open the file in a viewer that does not recalculate (a phone preview, Quick Look), or run `scripts/preview.py`. | The numbers are there and match what Excel shows. | K03 |
| 12 | Open the file in the spreadsheet program the person will use. | It opens without a repair prompt, and the Checks sheet says Balanced after it recalculates. | — |

Nine of these twelve rows name a machine rule; three (2, 9 and 12) have none.
The checker can say the model balances and follows its method; it cannot say the inputs are sensible or that the method
fits the business, which is why rows 2 and 9 matter, and row 12 tries the spreadsheet program itself.

Before numeric translation, preflight the original request and read the fit assessment. Check every requirement
is declared and that subscription billing and equipment purchase timing match. A balanced workbook does not prove fit.
