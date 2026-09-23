# Acceptance: what a person checks before the model goes anywhere

The checker reads the file and recomputes it. These are things it cannot see; SKILL.md's Boundaries also list the
edits it does not look at, such as a hidden row or a text box laid over a number. Walk them in order with the
workbook open; each row says what to do and what counts as a pass. The **Rule** column names the machine rule that
covers part of the same ground, where there is one.

| # | Do this | Passes when | Rule |
|---|---|---|---|
| 1 | Read the Assumptions sheet from the top, without looking at the statements. | You could tell someone what the business sells, for how much, and what it costs to run, from memory. | — |
| 2 | Read each input next to its label. | Each is a number the person gave you or one you can say why you chose. None is a placeholder. | — |
| 3 | Read the third line of the Assumptions sheet. | It says the numbers are made up, or where they come from, and that this is not financial advice. | K08 |
| 4 | Click three numbers on each statement. | Each shows a formula in the formula bar, never a typed number. | K02 |
| 5 | Read the last row of the balance sheet and every line of the Checks sheet. | All read 0.00, and the Checks sheet says Balanced. | K04, K07 |
| 6 | Put "Cash at the end of the year" from the cash flow beside "Cash" on the balance sheet. | They are the same number in every year. | K05 |
| 7 | Change one input — a price, the days customers take to pay — and watch the three statements. | The change runs through each statement that depends on it (a price reaches all three; the days customers take to pay reach the balance sheet and the cash flow, not profit), and the Checks sheet still says Balanced. Change it back afterwards. | K06 |
| 8 | Look at the lowest cash in any year. | If it is below zero, the Checks sheet says so, and you have told the person the plan needs money the model does not include. | K07 |
| 9 | Read the "How the model works" lines. | They describe what this business does — or you have told the person which of them does not fit (a business with seasons, stock bought in bulk, several loans). | — |
| 10 | Open the Charts sheet. | Revenue and net profit, and cash at the end of each year, match the statements. | K10 |
| 11 | Open the file in a viewer that does not recalculate (a phone preview, Quick Look), or run `scripts/preview.py`. | The numbers are there and match what Excel shows. | K03 |
| 12 | Open the file in the spreadsheet program the person will use. | It opens without a repair prompt, and the Checks sheet says Balanced after it recalculates. | — |

Eight of these twelve rows name a machine rule; four (1, 2, 9 and 12) have none.
The checker can say the model balances and follows its method; it cannot say the inputs are sensible or that the method
fits the business, which is why rows 1, 2 and 9 matter most, and row 12 is the only one that tries the spreadsheet program itself.
