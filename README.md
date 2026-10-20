# HOA Payment History Converter

Converts HOA "Payment History Report" text exports into per-owner PDF, Excel, or CSV ledgers.

Some legacy HOA accounting systems print the *Payment History Report* to a fixed-width text file. That's often all a departing management company hands over at transition. This tool reads that file and extracts each unit, account, lot, owner (current or previous), and transaction, including the continuation notes. It can export:

- one combined spreadsheet (Excel or CSV), or the raw report as a PDF
- one file per unit and owner, organized in folders (PDF, Excel, or CSV)
- a single account, for its current owner, previous owner, or both (one file each)

## Supported input

The parser expects the layout below. If your report doesn't match, the output will be wrong or empty.

- `UNIT: … ACCT: … LOT: … DELINQUENCY STATUS:` header lines, each followed by a `CURRENT OWNER:` or `PREVIOUS OWNER:` line
- Fixed-width transaction lines starting with an `MM/DD/YY` date, with columns for Payments, Description, Charges, Assessments (A1), Late Fees (01), Special Assessments (08), Prepaids (PP), and Other Accounts (code / amount / balance)
- Lines that aren't transactions (e.g. `Belongs to previous owner.`) are attached as notes to the transaction above them
- `~formfeed~` page-break markers, repeated page headers, and `----` separators are skipped

The software that produced these reports hasn't been confirmed; the export contains no vendor name.

## Demo data

[demo/sample_payment_history.txt](demo/sample_payment_history.txt) is a fictional report in the supported layout, for demos and testing. It covers 24 units, 10 of them with a previous owner, about 2,300 transactions, page breaks with repeated headers, late-notice and previous-owner notes, credit balances, prepayments, a special assessment, and a `1/2` address. All names, addresses, and amounts are made up. To regenerate it:

```bash
python demo/generate_demo_report.py
```

## Usage

1. **Browse…** to pick the `.txt` report. The account list loads automatically.
2. Choose an output format: PDF, Excel, or CSV.
3. Pick a mode:
   - **Default:** one combined file. The PDF option renders the raw report (font size and fixed-width formatting are adjustable). Excel and CSV are parsed into columns.
   - **Process only a single account:** pick a unit and export its current owner, previous owner, or both (one file per owner).
   - **Create separate files for each address:** produces a folder per unit with one file per owner.
4. **Preview Parse** shows what was extracted before you convert.
5. **Convert.**

Relative output names are saved next to the input file. If an output file is open in another app and can't be overwritten, a numbered copy (`name (1).xlsx`) is written instead.

### Output naming

| Mode | Name |
| --- | --- |
| Split by address | `<Output folder>/<Unit>/<Unit>_<Owner>_Current.ext`, `…_Previous.ext` |
| Single account, current or previous | `<Unit>_<Owner>_Current.ext` / `<Unit>_<Owner>_Previous.ext` |
| Single account, both | Two files: `<Unit>_<Owner>_Current.ext` and `<Unit>_<Owner>_Previous.ext` |

Units with no owner line are named `Unknown`.

## Running from source

Requires Python 3.13 with Tk (for example, Homebrew's `python@3.13` plus `python-tk@3.13`).

```bash
python3.13 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python payment_history_converter.py
```

## Building the macOS app

```bash
source venv/bin/activate
./build_mac.sh
```

This regenerates `icon.icns` from `icon.png` and builds `dist/HOA Payment History Converter.app` from `payment_history_converter.spec`. **The script deletes `build/` and `dist/` first,** so move any app you want to keep out of `dist/` before running it.
