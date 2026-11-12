# GTA Industrial Acquisition Model

A 10-year acquisition model for a **hypothetical 200,000 SF multi-tenant industrial building in Brampton, Ontario**. It is built in Excel with live formulas, and a Python script generates the whole workbook.

The goal is to answer the questions an investments team asks before buying a building: what the property earns today, what happens when each lease rolls, how debt changes the return, and what price still makes sense.

> The property and tenants are made up for learning. The market inputs (rents, operating costs, cap rates, interest rates) come from published Q1 to Q3 2026 GTA market reports, listed on the `Market Data` sheet and below. This is not investment advice.

## The deal

| | |
|---|---|
| Building | 200,000 SF, 28 ft clear, built 2005 (Class B), 4 tenants |
| Purchase price | $54.0M ($270/SF) plus 2.75% closing costs |
| Financing | 55% LTV, 5.50% fixed, 25-year amortization |
| Hold | 10 years (2027 to 2036), sale priced on Year 11 NOI at a 5.75% cap |

**Rent roll**

| Tenant | SF | Year 1 rent ($/SF) | Expiry | Note |
|---|---|---|---|---|
| A: Third-party logistics | 80,000 | $13.25 | End of Year 2 | Below market: upside at rollover |
| B: Food distribution | 60,000 | $18.50 | End of Year 5 | Signed near the 2023 peak: above market |
| C: Auto parts | 35,000 | $15.00 | End of Year 4 | |
| D: E-commerce fulfillment | 25,000 | $17.25 | End of Year 7 | |

## Base case results

| Metric | Result |
|---|---|
| Going-in cap rate | 5.48% |
| Year 1 NOI | $2.96M |
| Unlevered IRR | 6.25% |
| **Levered IRR** | **6.98%** |
| Equity multiple | 1.85x |
| Year 1 DSCR | 1.35x |
| Exit value | $62.8M ($314/SF) |

**What the model says:** at $54M the deal earns about 7% on equity, which is thin for an older building with this much rollover. The sensitivity grid shows the levered IRR only reaches about 8.8% if the price drops to $50M ($250/SF), holding a 5.75% exit cap. Timing also matters: a 7-year hold drops the levered IRR to about 4.4%, because the buyer would price the building on Year 8 NOI, a year with Tenant A and Tenant D both rolling.

## What the model does

| Sheet | Contents |
|---|---|
| `Summary` | Sources and uses, returns, debt metrics, exit, key risks, NOI and cash flow chart |
| `Assumptions` | Every input in one place (blue = input, black = formula, yellow = key assumption), the rent roll, and year-by-year market rent growth |
| `Cash Flow` | 11-year projection: base rent by tenant, rollover downtime, recoveries, vacancy, operating costs, NOI, TI, leasing commissions, reserves, debt service, sale, and the unlevered and levered cash flows |
| `Debt` | Amortization schedule, interest and principal, DSCR and debt yield by year |
| `Sensitivity` | Three live grids: levered and unlevered IRR (purchase price vs exit cap) and levered IRR (LTV vs interest rate) |
| `Sens_Calc` | Helper cash flows behind each grid cell, so the grids update when any input changes (no Excel data tables needed) |
| `Market Data` | The market figures behind the assumptions, with links |

**Lease rollover logic.** Each tenant pays its contract rent with annual escalations until its lease expires. At expiry the space re-leases at that year's market rent, then rolls again every 5 years. Downtime, tenant improvements and leasing commissions are weighted by a 65% renewal probability, similar to how market leasing assumptions work in Argus Enterprise.

**Flexible hold.** Change the hold period (1 to 10 years) on `Assumptions` and the sale year, loan payoff, returns and every sensitivity grid update.

## Market data used

| Input | Model value | Market reference |
|---|---|---|
| Market net rent | $15.75/SF | GTA average asking net rent $16.06/SF in Q3 2026; $16.84 for 30 ft+ clear (Avison Young) |
| Rent growth | 0.5% (2028), 2% (2029), then 3% | Rents fell for 7 straight quarters to Q1 2026, about 9% below the 2024 peak (WareCRE) |
| Operating costs (TMI) | $3.75/SF | GTA TMI averages $3.50 to $4.00/SF (WareCRE) |
| Cap rates | 5.48% going-in, 5.75% exit | Toronto industrial Class B 5.25% to 6.00% (CBRE Q3 2026); 5.10% to 5.50% (Avison Young Q3 2026) |
| Price | $270/SF | GTA average asking sale price $374.53/SF in Q2 2026 (CBRE), which includes newer stock |
| Loan rate | 5.50% | 5-year Government of Canada yield 3.59% on Oct 6, 2026, plus a lender spread |
| Closing costs | 2.75% | Ontario land transfer tax tops out at 2.0% for commercial property, plus legal and due diligence |

Full links are on the `Market Data` sheet.

## Check the math

The returns were checked outside Excel: an independent Python calculation of the IRRs, the loan payment, the Year 10 loan balance and the sale proceeds matched the workbook to the cent.

## Rebuild it

```bash
pip install openpyxl
python build_model.py
```

This writes `GTA_Industrial_Acquisition_Model.xlsx`. Excel recalculates every formula when the file opens.

## Tools

Excel, Python (openpyxl), real estate underwriting concepts: NOI, cap rates, IRR, equity multiple, DSCR, debt yield, lease rollover, TI and leasing commissions.

Built by Ei Htoo Khaing, Data Science and Analytics student at Seneca Polytechnic.
