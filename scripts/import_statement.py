#!/usr/bin/env python3
"""Import the Godrej Nest Navratri workbook into the dashboard JSON ledger."""

import argparse
import datetime as dt
import json
import re
import sys
import zipfile
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = (
    Path.home()
    / "OneDrive"
    / "Desktop"
    / "Cultural Godrej Nest"
    / "Navratri_Day Contribution Statement 2026.xlsx"
)
DEFAULT_OUTPUT = ROOT / "src" / "statement.json"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"m": MAIN_NS, "r": REL_NS}
EXCEL_EPOCH = dt.datetime(1899, 12, 30)


def column_number(cell_reference):
    letters = re.match(r"[A-Z]+", cell_reference).group()
    number = 0
    for letter in letters:
        number = number * 26 + ord(letter) - ord("A") + 1
    return number


def read_shared_strings(archive):
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return [
        "".join(text.text or "" for text in item.iter(f"{{{MAIN_NS}}}t"))
        for item in root
    ]


def decode_cell(cell, shared_strings):
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(text.text or "" for text in cell.iter(f"{{{MAIN_NS}}}t"))

    value = cell.find("m:v", NS)
    text = value.text if value is not None else ""
    if cell_type == "s" and text:
        return shared_strings[int(text)]
    return text


def read_sheet(archive, target, shared_strings):
    target = target.lstrip("/")
    if not target.startswith("xl/"):
        target = f"xl/{target}"
    root = ET.fromstring(archive.read(target))
    rows = []
    for row in root.findall(".//m:sheetData/m:row", NS):
        values = {}
        for cell in row.findall("m:c", NS):
            values[column_number(cell.attrib["r"])] = decode_cell(cell, shared_strings)
        if any(str(value).strip() for value in values.values()):
            rows.append((int(row.attrib.get("r", len(rows) + 1)), values))
    return rows


def load_sheets(workbook_path):
    with zipfile.ZipFile(workbook_path) as archive:
        workbook_xml = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relationships = {
            item.attrib["Id"]: item.attrib["Target"]
            for item in rels_xml.findall(f"{{{PKG_REL_NS}}}Relationship")
        }
        shared_strings = read_shared_strings(archive)
        sheets = {}
        for sheet in workbook_xml.find("m:sheets", NS):
            rel_id = sheet.attrib[f"{{{REL_NS}}}id"]
            sheets[sheet.attrib["name"]] = read_sheet(
                archive, relationships[rel_id], shared_strings
            )
        return sheets


def row_value(row, column):
    return str(row.get(column, "") or "").strip()


def check_headers(sheet_name, rows, expected):
    if not rows:
        raise ValueError(f"The '{sheet_name}' tab is empty.")
    actual = [str(value).strip().casefold() for _, row in rows[:1] for _, value in sorted(row.items())]
    expected_normalized = [header.casefold() for header in expected]
    if actual[: len(expected_normalized)] != expected_normalized:
        raise ValueError(
            f"Unexpected headers in '{sheet_name}'. Expected: {', '.join(expected)}."
        )


def parse_date(value, sheet_name, row_number):
    raw = str(value or "").strip()
    try:
        date = EXCEL_EPOCH.date() + dt.timedelta(days=int(Decimal(raw)))
    except (InvalidOperation, ValueError, OverflowError):
        date = None
    if date is None:
        for pattern in ("%Y-%m-%d", "%d %b %Y", "%d %B %Y", "%m/%d/%Y", "%d/%m/%Y"):
            try:
                date = dt.datetime.strptime(raw, pattern).date()
                break
            except ValueError:
                continue
    if date is None:
        raise ValueError(f"Invalid date in '{sheet_name}' row {row_number}: {raw!r}.")
    return f"{date.day} {date.strftime('%b')} {date.year}"


def parse_amount(value, sheet_name, row_number):
    raw = str(value or "").strip().replace(",", "").replace("₹", "")
    try:
        amount = Decimal(raw)
    except InvalidOperation as error:
        raise ValueError(
            f"Invalid amount in '{sheet_name}' row {row_number}: {raw!r}."
        ) from error
    if amount <= 0:
        raise ValueError(f"Amount must be positive in '{sheet_name}' row {row_number}.")
    normalized = format(amount.normalize(), "f")
    return normalized[:-2] if normalized.endswith(".0") else normalized


def parse_time(value):
    raw = str(value or "").strip()
    try:
        fraction = Decimal(raw)
        if Decimal(0) <= fraction < Decimal(1):
            total_minutes = int((fraction * Decimal(1440)).quantize(Decimal("1"))) % 1440
            hour, minute = divmod(total_minutes, 60)
            return f"{hour % 12 or 12}:{minute:02d} {'AM' if hour < 12 else 'PM'}"
    except InvalidOperation:
        pass
    match = re.fullmatch(r"(\d{1,2})[.:](\d{2})\s*(AM|PM)", raw, re.IGNORECASE)
    if match:
        return f"{int(match[1])}:{match[2]} {match[3].upper()}"
    return raw


def parse_rows(sheets):
    required_sheets = {"Total Receipts", "Stall Payment", "Total Expenses"}
    missing = required_sheets - sheets.keys()
    if missing:
        raise ValueError(f"Missing workbook tabs: {', '.join(sorted(missing))}.")

    receipt_rows = sheets["Total Receipts"]
    stall_rows = sheets["Stall Payment"]
    expense_rows = sheets["Total Expenses"]
    check_headers("Total Receipts", receipt_rows, ["Name", "Flat No", "Amount", "Date", "Time", "Remarks"])
    check_headers("Stall Payment", stall_rows, ["Name", "Flat No", "Amount", "Date", "Time", "Remarks"])
    check_headers("Total Expenses", expense_rows, ["Date", "Particular", "Amount", "Paid to", "Remarks"])

    receipts = []
    for row_number, row in receipt_rows[1:]:
        if not any(row_value(row, col) for col in range(1, 7)):
            continue
        name, amount, date = row_value(row, 1), row_value(row, 3), row_value(row, 4)
        if not name or not amount or not date:
            raise ValueError(f"Incomplete receipt in 'Total Receipts' row {row_number}.")
        remarks = row_value(row, 6)
        if name.casefold() == "anonymous":
            remarks = ""
        receipts.append({
            "name": name,
            "flat": row_value(row, 2),
            "amount": parse_amount(amount, "Total Receipts", row_number),
            "date": parse_date(date, "Total Receipts", row_number),
            "time": parse_time(row_value(row, 5)),
            "remarks": remarks,
        })

    stalls = []
    for row_number, row in stall_rows[1:]:
        if not any(row_value(row, col) for col in range(1, 7)):
            continue
        name, amount, date = row_value(row, 1), row_value(row, 3), row_value(row, 4)
        if not name or not amount or not date:
            raise ValueError(f"Incomplete stall payment in 'Stall Payment' row {row_number}.")
        stalls.append({
            "name": name,
            "flat": row_value(row, 2),
            "amount": parse_amount(amount, "Stall Payment", row_number),
            "date": parse_date(date, "Stall Payment", row_number),
            "time": parse_time(row_value(row, 5)),
            "remarks": row_value(row, 6),
        })

    expenses = []
    reported_expense_total = None
    for row_number, row in expense_rows[1:]:
        item = row_value(row, 2)
        amount, date = row_value(row, 3), row_value(row, 1)
        if item.casefold() == "total expenses":
            if amount:
                reported_expense_total = parse_amount(amount, "Total Expenses", row_number)
            continue
        if not any(row_value(row, col) for col in range(1, 6)):
            continue
        if not item or not amount or not date:
            raise ValueError(f"Incomplete expense in 'Total Expenses' row {row_number}.")
        expenses.append({
            "date": parse_date(date, "Total Expenses", row_number),
            "item": item,
            "amount": parse_amount(amount, "Total Expenses", row_number),
            "paidTo": row_value(row, 4),
            "remarks": row_value(row, 5),
        })

    totals = {
        "contributions": sum(Decimal(row["amount"]) for row in receipts),
        "stallFees": sum(Decimal(row["amount"]) for row in stalls),
        "expenses": sum(Decimal(row["amount"]) for row in expenses),
    }
    if reported_expense_total is not None and totals["expenses"] != Decimal(reported_expense_total):
        raise ValueError(
            f"Expense rows total {totals['expenses']}, but the workbook's Total Expenses row says {reported_expense_total}."
        )
    totals = {key: int(value) if value == value.to_integral_value() else float(value) for key, value in totals.items()}
    return {"receipts": receipts, "stalls": stalls, "expenses": expenses, "totals": totals}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", nargs="?", type=Path, default=DEFAULT_INPUT,
                        help=f"source .xlsx file (default: {DEFAULT_INPUT})")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT,
                        help=f"dashboard JSON output (default: {DEFAULT_OUTPUT})")
    args = parser.parse_args()

    if not args.workbook.is_file():
        raise FileNotFoundError(f"Workbook not found: {args.workbook}")
    if args.workbook.suffix.lower() != ".xlsx":
        raise ValueError("Input must be an .xlsx workbook.")

    data = parse_rows(load_sheets(args.workbook))
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(
        f"Imported {len(data['receipts'])} receipts, {len(data['stalls'])} stall payments, "
        f"and {len(data['expenses'])} expenses. Totals: contributions ₹{data['totals']['contributions']:,}; "
        f"stall fees ₹{data['totals']['stallFees']:,}; expenses ₹{data['totals']['expenses']:,}."
    )


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, KeyError, zipfile.BadZipFile) as error:
        print(f"Import failed: {error}", file=sys.stderr)
        sys.exit(1)
