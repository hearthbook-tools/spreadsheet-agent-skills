---
name: break-even-pricing
description: "Builds a break-even and pricing workbook for a product or service: contribution margin, break-even units, price for a target margin and what-if table. Use for pricing, margins or break-even questions."
license: "MIT"
---

# Break-even & pricing

> **Always run the bundled script first.** It is tested against hand-calculated results; do not rebuild the workbook from scratch. Extra tabs, charts or formatting may be added afterwards only if the user asks, and the script's tabs and formulas must be kept unchanged.

## Workflow
1. Collect: price per unit (or per hour/project), variable cost per unit (materials, platform/payment fees, contractor cost), monthly fixed costs (rent, software, insurance, owner salary if desired), and a target profit or target margin. Ask for platform fee percentages separately if they sell online (e.g. a 10% marketplace fee).
2. Run:
   ```bash
   python scripts/break_even.py --price 49 --var-cost 12 --fee-pct 0.10 --fixed 3200 --target-profit 2000 --out pricing.xlsx
   ```
3. Report: contribution margin per unit and %, break-even units per month, units needed for the target profit, and the price needed for a target margin. Point out that the what-if table shows profit at different prices and volumes; it does not predict demand.

## Formulas (all live in the workbook)
- Fee per unit = price × fee %
- Contribution per unit = price − variable cost − fee per unit
- Break-even units = fixed costs ÷ contribution per unit (rounded up)
- Units for target = (fixed + target profit) ÷ contribution per unit (rounded up)
- Price for target margin m (margin on price, after variable costs and fees, before fixed costs) = variable cost ÷ (1 − fee % − m)
