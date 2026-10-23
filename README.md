# Spreadsheet Agent Skills

Three free [Agent Skills](https://agentskills.io) (SKILL.md format) that make your AI assistant build spreadsheets that keep working, and catch the mistakes that silently break them. Every output is an Excel workbook with **live formulas**, not pasted numbers.

| Skill | What it does |
|---|---|
| [`formula-first-workbook`](skills/formula-first-workbook) | Rules for building workbooks (inputs tab, one formula per column, checks) **plus a linter** that finds overwritten formulas, `SUM`s that stop short of the data, links to deleted sheets and magic numbers buried in formulas. |
| [`break-even-pricing`](skills/break-even-pricing) | Contribution margin, break-even units, units for a profit target, price for a target margin, and an editable what-if table. |
| [`month-end-close`](skills/month-end-close) | A small-business month-end checklist tailored to how you work, worked through step by step. |

## Linter example

```
$ python skills/formula-first-workbook/scripts/lint_workbook.py messy.xlsx
messy.xlsx: 2 error(s), 2 warning(s)
  ERROR Calc!A8: SUM stops at row 4 but values continue to row 7
  ERROR Calc!C2: Refers to missing sheet 'Budget'
  WARN  Calc!B4: Typed number inside a column of formulas (formula overwritten?)
  WARN  Calc!C1: Number(s) typed inside formula: 0.0725 - move to an Inputs cell
```

## Install

**Claude.ai / desktop app** (paid plan with custom skills): turn on *Settings → Capabilities → Code execution and file creation*, then *Customize → Skills → + → Upload a skill* and upload a file from [`dist/`](dist) (one `.skill` file per skill).

**Claude Code:**
```bash
git clone https://github.com/hearthbook-tools/spreadsheet-agent-skills
cp -R spreadsheet-agent-skills/skills/* ~/.claude/skills/
```
Scripts need Python 3 and `openpyxl` (already available in Claude's code-execution sandbox).

## Tested
Each script is tested end to end against hand-calculated results and validated with Anthropic's skill-creator validator. The break-even skill was also verified by uploading to Claude.ai and comparing the produced workbook cell by cell with the reference output.

## Want the bookkeeping workhorses?
The full **Bookkeeping & Spreadsheet Agent Skills** pack adds five more tested skills: bank-statement categorizer (any bank CSV), books reconciler, AR aging, 13-week cash-flow forecast and quarterly tax set-aside tracker, with sample data and an install guide → **[hearthbook.gumroad.com/l/bookkeeping-agent-skills](https://hearthbook.gumroad.com/l/bookkeeping-agent-skills?utm_source=github&utm_medium=readme)**

---
MIT licensed. Outputs organize and calculate your data; they are not accounting, tax or legal advice. Not affiliated with or endorsed by Anthropic; "Claude" is used only to describe compatibility.
