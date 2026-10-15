# The business file (0.2.1)

## Fit assessment before translation

First write `request.json`: `{"description": "Original request", "scope": {"features": ["annual_sales", "straight_line"],
"notes": "Explain timing, conventions and omitted inputs."}}`. Run `make_model.py --preflight request.json` before
creating numeric assumptions. This checks declared capabilities and a few prose patterns, not arbitrary prose.
The agent must classify all material requirements. Missing scope and unknown/unsupported declared features are
B04. Unsupported needs omitted from the declaration can pass when their wording misses the limited patterns;
negation is not understood by those patterns. Preserve original prose and disclose the classification.

Supported names: `annual_sales`, `monthly_subscription`, `straight_line`, `reducing_balance`, `working_capital_days`,
`equal_principal_loan`, `profit_tax`. Business scope must include exactly one sales and one depreciation method,
matching the numeric mode fields, plus applicable working-capital, loan and tax capabilities.

Refused names: `cohort_churn`, `seasonality`, `monthly_cash`, `multiple_products`, `deferred_revenue`,
`custom_depreciation`, `multiple_loans`, `opening_working_capital`, `loss_carryforward`, `owner_distributions`.
All other feature names are refused too. Custom depreciation includes residuals, used assets, partial years,
disposals and nonconstant purchases. Nonstandard loans include annuity payments, balloons and grace periods.
Refused requests do not become annual approximations. A user-authorized simpler request needs a new preflight
and explicit notes explaining the change.

Cash is tested at opening and year ends only. Intra-year funding error is unbounded. Annual profit tax has no
carryforward; working capital uses annual flows × days / 365. Disclose these before assumption translation.

One JSON object is the whole input. `scripts/make_model.py --init business.json` writes a filled-in example to edit.
Percentages are written as fractions (`0.12` is 12%). A field this list does not name is refused, and so is a key
written twice (B01) — a misspelt field would otherwise be ignored and the model built without it. Text fields may not
hold a control character (the workbook's XML cannot) or an emoji or icon character (the checker's K11 refuses them in
the workbook): the builder refuses them first (B01), including U+FFFE and U+FFFF. No amount, count or price may be above a billion (B02): this is a model of a small
business. The limit does not by itself keep every sum exact to the penny: large amounts growing by hundreds of percent a
year can still reach sums where floating point is not, and the checker may then report a difference at the last penny
(K05 or K06).

## Top level

| Field | What it is | Rule |
|---|---|---|
| `name` | The business's name | text; it titles every sheet |
| `business_line` | One-line input metadata; Summary generates its sales description from the model instead | text; at most 110 characters, no newlines |
| `description` | A sentence or two about what it does | text; under the title of the Assumptions sheet |
| `currency` | `£`, `$`, `€`, or a three-letter code | at most three characters (B02) |
| `first_year` | The first year of the model | a whole number, 1900 to 2200 |
| `years` | How many years | a whole number, 1 to 10 (B02) |
| `illustrative` | `true` when the numbers are made up | true or false (B01) |
| `sources` | Supplied declaration of where real numbers come from; not authenticated | text; required when `illustrative` is false, and not allowed when it is true (B03) |
| `scope` | `features` list and `notes` text, copied from preflight with all added assumptions | required; unsupported/unknown features refused (B04) |
| `sales` | `units`, `unit_growth`, `price`, `price_growth`, `mode`, `opening_customers`, `monthly_additions`, `monthly_churn` | below |
| `costs` | `unit_cost`, `unit_cost_growth`, `fixed_costs`, `fixed_cost_growth` | below |
| `equipment` | `initial`, `yearly`, `depreciation_rate`, `method`, `useful_life` | below |
| `working_capital` | `receivable_days`, `inventory_days`, `payable_days` | below |
| `funding` | `equity`, `loan`, `interest_rate`, `loan_years` | below |
| `tax_rate` | The tax rate on profit | 0 to 0.999 |

## The assumptions (B02 gives each range)

| Field | Meaning | Range |
|---|---|---|
| `sales.units` | Units sold in the first year (jobs, items, subscriptions: whatever is sold) | 0 to a billion |
| `sales.mode` | 0 annual units/price; 1 monthly subscriptions | integer 0 or 1 |
| `sales.opening_customers` | Paying customers before first month | 0 to a billion; zero in mode 0 |
| `sales.monthly_additions` | Customers added at each month end after churn | 0 to a billion; zero in mode 0 |
| `sales.monthly_churn` | Fraction of opening customers lost each month | 0 to 1; zero in mode 0 |
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
| `equipment.method` | 0 reducing balance; 1 straight line | integer 0 or 1 |
| `equipment.useful_life` | Whole-year life, new equipment, zero residual, used only in method 1 | integer 1 to 50 |
| `working_capital.receivable_days` | Days customers take to pay | 0 to 365 |
| `working_capital.inventory_days` | Days of stock held | 0 to 365 |
| `working_capital.payable_days` | Days taken to pay suppliers | 0 to 365 |
| `funding.equity` | Owners' money put in at the start | 0 to a billion |
| `funding.loan` | Loan taken at the start | 0 to a billion |
| `funding.interest_rate` | Interest rate on the loan | 0 to 0.999 |
| `funding.loan_years` | Years to repay the loan in equal amounts | a whole number, 1 to 50 |

## How each line is worked out

This is what the Assumptions sheet says under "How the model works", and what `model_check.py` checks (K06).

- Annual mode: revenue = units × price; cost of sales = units × unit cost, each with its own annual growth rate.
- Subscription mode: bill opening customers, churn them, then add customers at month end. Closing customers =
  opening × (1 − churn) + additions. Annual units = sum of 12 opening counts. Set unused annual units and growth
  to zero. Price and cost per subscriber-month change at year boundaries. Counts are fractional expectations.
- Reducing balance: opening net equipment × depreciation rate. Straight line: each purchase depreciates cost / life
  from year one for opening new assets, and the next year for year-end purchases, stopping after useful life.
- Operating profit = revenue − cost of sales − fixed costs − depreciation.
- Interest = the loan at the start of the year × the interest rate.
- Tax = the tax rate × profit before tax, when there is a profit; nothing on a loss, and a loss is not carried forward.
- Money owed by customers = revenue × days ÷ 365; stock and money owed to suppliers = cost of sales × days ÷ 365.
- The loan is repaid in equal yearly amounts until it is gone; nothing is paid out to the owners.
- Cash at the start = owners' money + loan − equipment bought at the start. After that, the cash flow: net profit,
  plus depreciation, less the money newly tied up in customers and stock, plus the money newly owed to suppliers,
  less equipment bought, less the loan repaid.

## An example

This illustrative bike-repair/parts scenario treats all activity as one average sales unit: 2,400 first-year units,
£45 revenue and £18 variable cost per unit. Separate repair and counter-parts prices, volumes and margins are
not represented. A request needing them must declare `multiple_products` and be refused; using this aggregate
requires explicit agreement to the simplification. The illustrative flag and any sources are supplied declarations,
not independently verified provenance.

```json
{
  "name": "Harbour Bike Repairs",
  "business_line": "A bike workshop earning repair and parts revenue.",
  "description": "A one-workshop bike repair and parts business on a harbour road: repairs and servicing, with parts sold at the counter.",
  "currency": "£", "first_year": 2027, "years": 3, "illustrative": true,
  "scope": {"features": ["annual_sales", "reducing_balance", "working_capital_days", "equal_principal_loan"],
    "notes": "Annual single-line planning. New equipment bought at year end; cash not scheduled within a year."},
  "sales": {"units": 2400, "unit_growth": 0.12, "price": 45, "price_growth": 0.03,
    "mode": 0, "opening_customers": 0, "monthly_additions": 0, "monthly_churn": 0},
  "costs": {"unit_cost": 18, "unit_cost_growth": 0.02, "fixed_costs": 48000, "fixed_cost_growth": 0.03},
  "equipment": {"initial": 30000, "yearly": 4000, "depreciation_rate": 0.2, "method": 0, "useful_life": 5},
  "working_capital": {"receivable_days": 15, "inventory_days": 45, "payable_days": 30},
  "funding": {"equity": 25000, "loan": 20000, "interest_rate": 0.07, "loan_years": 5},
  "tax_rate": 0.19
}
```
