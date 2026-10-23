# Build rules (formula-first)

1. Separate tabs: `Inputs` → `Data` (raw entries) → `Calc` → `Summary` → `Checks`.
2. Every assumption lives on Inputs with label, value, unit, source/date. Name important cells (Defined Names) when the workbook is for Excel users.
3. One formula pattern per column. Fill it down to the end of the data area plus a buffer (e.g. 500 rows) guarded with IF(A2="","",...).
4. Totals: `=SUM(C2:C501)` over the full buffer, never just the current data rows.
5. Percentages stored as decimals (0.25) formatted as %, never 25 with a "%" label.
6. Dates as real dates; months via EOMONTH or DATE(YEAR(),MONTH(),1), not text.
7. Rounding only at presentation or where money is actually rounded (ROUND(x,2)).
8. Errors guarded where expected (IFERROR around lookups) but never used to hide real mistakes in totals.
9. Checks tab: totals reconcile, counts match, no blank required inputs, no negative quantities. Show a single overall status cell.
10. Before delivering: run `scripts/lint_workbook.py`; zero ERROR lines.
