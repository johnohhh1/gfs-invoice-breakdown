# GFS Invoice Breakdown

Turns restaurant purchasing PDFs into a categorized Excel workbook, so you can see in one place what a truck or a week of receiving cost by category. No more hand-sorting line items.

| Script | Input | Output |
|---|---|---|
| `gfs_breakdown.py` | A Gordon Food Service purchase order or receiving PDF | `.xlsx` with an F&B sheet broken into meat/seafood, poultry, dairy/eggs, produce/frozen, sauces, dry goods, desserts, oils, beverages and beer/liquor, plus separate To-Go, Misc Supplies and Chemicals sheets with case counts and totals |
| `menulink_cost_journal.py` | A MenuLink "2.05 Cost Journal – Receiving by Vendor" PDF | `.xlsx` with a summary sheet and per-category sheets. The categories come from the G/L codes printed in the report, so no keyword guessing is needed |

## Install

```bash
pip install -r requirements.txt
```

Requires Python 3.9+.

## Usage

```bash
python gfs_breakdown.py "Purchase Order 6439.pdf"
# -> Purchase Order 6439_Breakdown.xlsx in the current folder

python menulink_cost_journal.py cost_journal_p11w1.pdf p11w1_costs.xlsx
```

Both scripts also print a category summary to the console. The store name is read from the PDF header when present. Otherwise set `GFS_STORE_NAME`.

## Customizing categories

GFS items are sorted by keyword rules in `CATEGORY_RULES` at the top of `gfs_breakdown.py`. The first match wins, and the non-food categories are checked first. If an item lands in the wrong place, add its description keyword to the right list. Items that match nothing are printed as `UNCATEGORIZED` warnings so you can spot them.

## Privacy

Invoices are your business data. `.gitignore` blocks `*.pdf` and `*.xlsx` so they don't end up in a commit.

## Contributing

PRs are welcome, especially:

- Keyword rules for items your distributor sells that land in the wrong category
- Support for other distributors' PO layouts (Sysco, US Foods, ...)
- Tests built from synthetic invoice text

Please don't attach real invoices to issues. Paste a few anonymized lines instead.

## License

MIT. See [LICENSE](LICENSE). Not affiliated with Gordon Food Service or MenuLink.
