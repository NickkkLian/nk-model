# The business file

One JSON object is the whole input. `scripts/make_model.py --init business.json` writes a filled-in example to edit.
Percentages are written as fractions (`0.12` is 12%). A field this list does not name is refused, and so is a key
written twice (B01) — a misspelt field would otherwise be ignored and the model built without it. Text fields may not
hold a control character (the workbook's XML cannot) or an emoji or icon character (the checker's K11 refuses them in
the workbook): the builder refuses them first (B01). Leave out U+FFFE and U+FFFF too: the XML cannot hold them either,
and the builder does not refuse them. No amount, count or price may be above a billion (B02): this is a model of a small
business. The limit does not by itself keep every sum exact to the penny: large amounts growing by hundreds of percent a
year can still reach sums where floating point is not, and the checker may then report a difference at the last penny
(K05 or K06).

## Top level

| Field | What it is | Rule |
|---|---|---|
| `name` | The business's name | text; it titles every sheet |
| `description` | A sentence or two about what it does | text; under the title of the Assumptions sheet |
| `currency` | `£`, `$`, `€`, or a three-letter code | at most three characters (B02) |
| `first_year` | The first year of the model | a whole number, 1900 to 2200 |
| `years` | How many years | a whole number, 1 to 10 (B02) |
| `illustrative` | `true` when the numbers are made up | true or false (B01) |
| `sources` | Where real numbers come from | text; required when `illustrative` is false, and not allowed when it is true (B03) |
| `sales` | `units`, `unit_growth`, `price`, `price_growth` | below |
| `costs` | `unit_cost`, `unit_cost_growth`, `fixed_costs`, `fixed_cost_growth` | below |
| `equipment` | `initial`, `yearly`, `depreciation_rate` | below |
| `working_capital` | `receivable_days`, `inventory_days`, `payable_days` | below |
| `funding` | `equity`, `loan`, `interest_rate`, `loan_years` | below |
| `tax_rate` | The tax rate on profit | 0 to 0.999 |

## The assumptions (B02 gives each range)

| Field | Meaning | Range |
|---|---|---|
| `sales.units` | Units sold in the first year (jobs, items, subscriptions: whatever is sold) | 0 to a billion |
| `sales.unit_growth` | Growth in units each year | -0.9 to 5 |
| `sales.price` | Price of one unit in the first year | above 0, at most a billion |
| `sales.price_growth` | Growth in price each year | -0.9 to 5 |
| `costs.unit_cost` | What one unit costs to make or deliver in the first year | 0 to a billion |
| `costs.unit_cost_growth` | Growth in that cost each year | -0.9 to 5 |
| `costs.fixed_costs` | Costs that do not depend on units (rent, wages, insurance) in the first year | 0 to a billion |
| `costs.fixed_cost_growth` | Growth in fixed costs each year | -0.9 to 5 |
| `equipment.initial` | Equipment bought before the first year | 0 to a billion |
| `equipment.yearly` | Equipment bought during each year | 0 to a billion |
| `equipment.depreciation_rate` | Share of the equipment's value written off each year (reducing balance) | above 0, at most 1 |
| `working_capital.receivable_days` | Days customers take to pay | 0 to 365 |
| `working_capital.inventory_days` | Days of stock held | 0 to 365 |
| `working_capital.payable_days` | Days taken to pay suppliers | 0 to 365 |
| `funding.equity` | Owners' money put in at the start | 0 to a billion |
| `funding.loan` | Loan taken at the start | 0 to a billion |
| `funding.interest_rate` | Interest rate on the loan | 0 to 0.999 |
| `funding.loan_years` | Years to repay the loan in equal amounts | a whole number, 1 to 50 |

## How each line is worked out

This is what the Assumptions sheet says under "How the model works", and what `model_check.py` checks (K06).

- Revenue = units × price; cost of sales = units × cost of one unit. Each grows at its own rate from the first year.
- Depreciation = the equipment's value at the start of the year × the depreciation rate.
- Operating profit = revenue − cost of sales − fixed costs − depreciation.
- Interest = the loan at the start of the year × the interest rate.
- Tax = the tax rate × profit before tax, when there is a profit; nothing on a loss, and a loss is not carried forward.
- Money owed by customers = revenue × days ÷ 365; stock and money owed to suppliers = cost of sales × days ÷ 365.
- The loan is repaid in equal yearly amounts until it is gone; nothing is paid out to the owners.
- Cash at the start = owners' money + loan − equipment bought at the start. After that, the cash flow: net profit,
  plus depreciation, less the money newly tied up in customers and stock, plus the money newly owed to suppliers,
  less equipment bought, less the loan repaid.

## An example

```json
{
  "name": "Harbour Bike Repairs",
  "description": "A one-workshop bike repair and parts business on a harbour road: repairs and servicing, with parts sold at the counter.",
  "currency": "£", "first_year": 2027, "years": 3, "illustrative": true,
  "sales": {"units": 2400, "unit_growth": 0.12, "price": 45, "price_growth": 0.03},
  "costs": {"unit_cost": 18, "unit_cost_growth": 0.02, "fixed_costs": 48000, "fixed_cost_growth": 0.03},
  "equipment": {"initial": 30000, "yearly": 4000, "depreciation_rate": 0.2},
  "working_capital": {"receivable_days": 15, "inventory_days": 45, "payable_days": 30},
  "funding": {"equity": 25000, "loan": 20000, "interest_rate": 0.07, "loan_years": 5},
  "tax_rate": 0.19
}
```
