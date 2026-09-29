# SFDC Report

Generates a fictitious Salesforce opportunity export for a fictional public sector software vendor ("Amazing Software"), then builds a weekly summary report from that export.

The project simulates a real workflow: export opportunities from Salesforce, then summarize them for sales leadership. All data is synthetic.

## What it produces

| File | Description |
|------|-------------|
| `SFDC_Opportunities_Export.xlsx` | 100 synthetic sales opportunities, one per row |
| `SFDC_Summary_Report.docx` | Summary report built from the spreadsheet |

## Requirements documents

- `SFDCdatagen.pdf` defines the spreadsheet fields and allowed values.
- `Sample SFDC report.pdf` defines the report format.

## Usage

```bash
pip install openpyxl python-docx
python3 gen.py
```

Running `gen.py` regenerates both output files. The data is random, so each run produces a different spreadsheet and report. The report date is the current date on the machine running the script.

## How `gen.py` works

1. **Entities.** Builds the list of government entities, each with its market segment:
   - Federal Civilian (7 agencies)
   - Federal Defense (5 branches)
   - State and Local (50 states and 50 cities)

   It picks 100 distinct entities at random.
2. **Opportunity fields.** For each entity it generates:
   - a random US-style contact name;
   - a list price (log-normal distribution, minimum $25,000, maximum $100,000,000);
   - a discount, an opportunity type, a deal stage, a proposed product, recent and next solutions engineer activity, a product gap, a competitor, and a Technical Win flag.
3. **Realism rules.**
   - Products are matched to the opportunity type. For example, Greenfield Observability deals get Observe or DEM.
   - Closed Won deals always have Technical Win = Yes.
   - Early-stage deals (Prospecting, Discovery, Proposal) are mostly Technical Win = No.
   - Deals without a Technical Win are more likely to have a product gap.
4. **Spreadsheet.** Writes the rows to `SFDC_Opportunities_Export.xlsx` with a bold header, a frozen header row and currency formatting.
5. **Report.** Reads the generated rows and writes `SFDC_Summary_Report.docx` containing:
   - a title and the current date;
   - the total number of opportunities;
   - the number of deals with Technical Win = No;
   - the total list price of those deals;
   - the top 15 deals with Technical Win = No, ranked by list price from largest to smallest. Each is one line: entity, market, price, opportunity type, next activity, product gap (omitted if "none"), and competitor.

## Notes

- The revenue total sums list price only. Discount is not applied.
- To get repeatable output, add `random.seed(<number>)` after the imports in `gen.py`.
- To change the possible values, edit the lists at the top of `gen.py`.
