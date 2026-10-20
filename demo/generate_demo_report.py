"""Generate a fictional Payment History Report in the fixed-width layout the converter reads.

All names, addresses, and amounts are made up. Output is deterministic (fixed seed).

    python demo/generate_demo_report.py            # writes demo/sample_payment_history.txt
"""
import os
import random
from datetime import date

HOA_NAME = "Sample Meadows HOA (DEMO DATA)"
REPORT_DATE = date(2025, 2, 28)
LINES_PER_PAGE = 58

FIRST = ["Alex", "Jordan", "Taylor", "Morgan", "Casey", "Riley", "Jamie", "Avery", "Quinn", "Drew",
         "Reese", "Skyler", "Rowan", "Emerson", "Hayden", "Parker", "Sawyer", "Logan"]
LAST = ["Example", "Sample", "Placeholder", "Demo", "Testwell", "Mockford", "Fictus", "Specimen",
        "Dummett", "Prototype", "Templeton", "Modelle"]
STREETS = [("Sample Way", "SW"), ("Example Court", "EC")]

ASSESSMENT = 185.00
LATE_FEE = 25.00
SPECIAL_ASSESSMENT = 500.00
OTHER_CODES = ["09", "10", "12"]


def money(value):
    """Positive amounts get a trailing space so they line up with (negative) amounts"""
    if value is None:
        return ""
    return f"({abs(value):,.2f})" if value < 0 else f"{value:,.2f} "


def row(txn_date, payments=None, description="", charges=None, a1=None, late=None, special=None,
        prepaid=None, code="", amount=None, balance=None):
    return (f"{txn_date:%m/%d/%y} {money(payments):>11} {description:<17} {money(charges):>11} "
            f"{money(a1):>11} {money(late):>11} {money(special):>11} {money(prepaid):>11} "
            f"  {code:<4}   {money(amount):>11} {money(balance):>12}").rstrip()


def page_header(page):
    return [
        f"DATE:  {REPORT_DATE.month}/{REPORT_DATE.day}/{REPORT_DATE:%y}",
        f"TIME: 10:30 AM{'':33}{HOA_NAME:<59}PAGE {page:>4}",
        f"{'':47}PAYMENT HISTORY REPORT - {REPORT_DATE:%b}. {REPORT_DATE.day}, {REPORT_DATE.year}",
        "",
    ]


def account_header(unit, acct, delinquency, owner_type, owner):
    label = "CURRENT OWNER: " if owner_type == "CURRENT" else "PREVIOUS OWNER:"
    return [
        f"UNIT:   {unit:<32}ACCT: {acct:<20}LOT: {acct:<12}DELINQUENCY STATUS: {delinquency}".rstrip(),
        f"{label} {owner}",
        "",
        "  TXN                   TRANSACTION                ASSESSMENTS  LATE FEES  SPEC ASSMTS  PREPAIDS   <- OTHER ACCOUNTS ->",
        "  DATE     PAYMENTS     DESCRIPTION      CHARGES       A1          01          08          PP       CODE       AMOUNT     BALANCE",
        "-------- ----------- ----------------- ----------- ----------- ----------- ----------- ----------- ------   ----------- ------------",
    ]


def owner_history(rng, start, end, special_month):
    """Monthly transaction blocks for one owner, each a list of lines"""
    blocks = []
    balance = 0.0
    if rng.random() < 0.3:
        credit = rng.choice([50.0, 92.5, 185.0])
        balance -= credit
        blocks.append([row(start, payments=credit, description="INIT CREDIT BAL", prepaid=credit, balance=balance)])

    month = start
    while month <= end:
        lines = []
        balance += ASSESSMENT
        if balance <= 0:
            # A credit on the account covers the charge
            lines.append(row(month, description="APPLY CHARGES", charges=ASSESSMENT, a1=ASSESSMENT, balance=balance))
            lines.append(row(month, description="APPLY PREPAYMNT", a1=-ASSESSMENT, balance=balance))
        else:
            lines.append(row(month, description="APPLY CHARGES", charges=ASSESSMENT, a1=ASSESSMENT, balance=balance))

        if (month.year, month.month) == special_month:
            balance += SPECIAL_ASSESSMENT
            lines.append(row(month, description="APPLY SPEC ASSMT", charges=SPECIAL_ASSESSMENT,
                             special=SPECIAL_ASSESSMENT, balance=balance))

        if rng.random() < 0.08:
            code = rng.choice(OTHER_CODES)
            fee = rng.choice([35.0, 50.0, 125.0])
            balance += fee
            lines.append(row(month, description="EXPENSE ADJ", charges=fee, code=code, amount=fee, balance=balance))

        roll = rng.random()
        check = f"{rng.randint(1000, 9999)}"
        ref = f"{rng.randint(100000, 999999)}"
        pay_day = date(month.year, month.month, rng.randint(2, 12))
        if roll < 0.12:
            # Paid late: late fee, then a note
            balance += LATE_FEE
            late_day = date(month.year, month.month, 16)
            lines.append(row(late_day, description="APPLY LATE FEE", charges=LATE_FEE, late=LATE_FEE, balance=balance))
            lines.append(f"{late_day:%m/%d/%y} Action taken: 01 - Late Notice")
        elif roll < 0.90:
            amount = round(balance, 2) if balance > 0 else ASSESSMENT
            balance -= amount
            lines.append(row(pay_day, payments=amount, description=f"{check:<10} {ref}",
                             a1=-amount if balance >= 0 else -ASSESSMENT,
                             prepaid=-(amount - ASSESSMENT) if balance < 0 else None, balance=balance))
        else:
            # Prepaying the next few months
            amount = ASSESSMENT * 3
            balance -= amount
            lines.append(row(pay_day, payments=amount, description=f"{check:<10} {ref}", prepaid=-amount, balance=balance))

        blocks.append(lines)
        month = date(month.year + month.month // 12, month.month % 12 + 1, 1)
    return blocks


def build_accounts(rng):
    names = [f"{first} {last}" for first in FIRST for last in LAST]
    rng.shuffle(names)
    accounts = []
    for street, abbrev in STREETS:
        for number in range(1, 13):
            house = f"{number} 1/2" if (street, number) == ("Example Court", 7) else str(number)
            unit = f"{house} {street}"
            acct = f"{number}{abbrev}" + ("H" if "1/2" in house else "")
            delinquency = "DELINQUENT" if rng.random() < 0.1 else ""
            owners = []
            if rng.random() < 0.35:
                # Sold partway through: previous owner's history, then the current owner's
                sale = date(2023, rng.randint(1, 12), 1)
                owners.append(("CURRENT", names.pop(), sale, None))
                owners.append(("PREVIOUS", names.pop(), date(2022, 1, 1), sale))
            else:
                owners.append(("CURRENT", names.pop(), date(2022, 1, 1), None))
            accounts.append((unit, acct, delinquency, owners))
    return accounts


def generate(seed=7):
    rng = random.Random(seed)
    end_month = date(REPORT_DATE.year, REPORT_DATE.month, 1)
    lines, page, page_lines = [], 0, LINES_PER_PAGE

    for unit, acct, delinquency, owners in build_accounts(rng):
        for owner_type, owner, start, sold in owners:
            last_month = date(sold.year - (sold.month == 1), (sold.month - 2) % 12 + 1, 1) if sold else end_month
            blocks = owner_history(rng, start, last_month, special_month=(2023, 6))
            if owner_type == "PREVIOUS":
                blocks[-1].append(f"{last_month:%m/%d/%y} Belongs to previous owner.")
            header = account_header(unit, acct, delinquency, owner_type, owner)
            first_block = True
            for block in blocks:
                needed = len(block) + (0 if first_block else 1)
                if page_lines + needed > LINES_PER_PAGE:
                    if page:
                        lines.append("~formfeed~")
                    page += 1
                    lines += page_header(page) + header
                    page_lines = 4 + len(header)
                elif first_block:
                    lines += [""] + header
                    page_lines += 1 + len(header)
                if not first_block:
                    lines.append("--------")
                    page_lines += 1
                lines += block
                page_lines += len(block)
                first_block = False
    return "\n".join(lines) + "\n"


if __name__ == "__main__":
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample_payment_history.txt")
    with open(out, "w", encoding="utf-8") as f:
        f.write(generate())
    print(f"Wrote {out}")
