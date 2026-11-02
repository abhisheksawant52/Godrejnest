#!/usr/bin/env python3
"""Import a Godrej Nest event workbook into its dashboard event JSON file."""

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
EVENTS_DIR = ROOT / "src" / "events"
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


HEADER_ALIASES = {
    "name": {"name", "donor", "donor name", "contributor", "contributor name", "person", "stallholder", "stall owner"},
    "flat": {"flat", "flat no", "flat number", "flatno", "unit", "unit no", "apartment"},
    "amount": {"amount", "amount paid", "payment", "paid amount", "total amount", "amount inr", "inr"},
    "date": {"date", "payment date", "transaction date"},
    "time": {"time", "payment time", "transaction time"},
    "remarks": {"remarks", "remark", "notes", "note", "comments"},
    "item": {"particular", "particulars", "item", "expense", "expense item", "details", "description"},
    "paidTo": {"paid to", "paidto", "payee", "vendor", "recipient"},
}


def normalize_header(value):
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").casefold()).strip()


def header_map(row):
    aliases = {
        normalize_header(alias): field
        for field, values in HEADER_ALIASES.items()
        for alias in values
    }
    found = {}
    for column, value in row.items():
        normalized = normalize_header(value)
        field = aliases.get(normalized)
        if field is None:
            if normalized.startswith("amount ") or normalized.startswith("total amount "):
                field = "amount"
            elif normalized.startswith("date "):
                field = "date"
            elif normalized.startswith("time "):
                field = "time"
            elif normalized.startswith("name "):
                field = "name"
        if field and field not in found:
            found[field] = column
    return found


def find_header(rows):
    # Allow a title or blank rows above the actual column headings.
    for index, (row_number, row) in enumerate(rows[:10]):
        columns = header_map(row)
        if {"amount", "date"}.issubset(columns) and (
            "name" in columns or "item" in columns
        ):
            return index, columns
    return None, {}


def value_for(row, columns, field):
    column = columns.get(field)
    return row_value(row, column) if column is not None else ""


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
    incoming_sheets = []
    expense_sheets = []
    unrecognized = []
    for sheet_name, rows in sheets.items():
        header_index, columns = find_header(rows)
        if header_index is None:
            unrecognized.append(sheet_name)
            continue
        if "name" in columns and {"amount", "date"}.issubset(columns):
            incoming_sheets.append((sheet_name, rows, header_index, columns))
        elif "item" in columns and {"amount", "date"}.issubset(columns):
            expense_sheets.append((sheet_name, rows, header_index, columns))

    if not incoming_sheets:
        raise ValueError(
            "Could not find an incoming payments tab. Include columns for Name, Amount, and Date."
        )
    if not expense_sheets:
        raise ValueError(
            "Could not find an expenses tab. Include columns for Particular/Item, Amount, and Date."
        )

    # Stalls keep their own dashboard total. Other incoming tabs (for example,
    # owner registrations) are included in contributions.
    def is_stall(sheet):
        return "stall" in normalize_header(sheet[0]).split()

    receipts = []
    stalls = []
    for sheet_name, rows, header_index, columns in incoming_sheets:
        destination = stalls if is_stall((sheet_name,)) else receipts
        for row_number, row in rows[header_index + 1:]:
            name = value_for(row, columns, "name")
            amount = value_for(row, columns, "amount")
            date = value_for(row, columns, "date")
            if not any((name, amount, date)):
                continue
            if not name or not amount or not date:
                raise ValueError(f"Incomplete payment in '{sheet_name}' row {row_number}.")
            remarks = value_for(row, columns, "remarks")
            if name.casefold() == "anonymous":
                remarks = ""
            destination.append({
                "name": name,
                "sourceSheet": sheet_name,
                "flat": value_for(row, columns, "flat"),
                "amount": parse_amount(amount, sheet_name, row_number),
                "date": parse_date(date, sheet_name, row_number),
                "time": parse_time(value_for(row, columns, "time")),
                "remarks": remarks,
            })

    expenses = []
    reported_expense_total = None
    for sheet_name, rows, header_index, columns in expense_sheets:
        for row_number, row in rows[header_index + 1:]:
            item = value_for(row, columns, "item")
            amount = value_for(row, columns, "amount")
            date = value_for(row, columns, "date")
            if normalize_header(item) == "total expenses":
                if amount:
                    try:
                        reported_expense_total = Decimal(amount.replace(",", "").strip())
                    except InvalidOperation as error:
                        raise ValueError(
                            f"Invalid amount in '{sheet_name}' row {row_number}: {amount!r}."
                        ) from error
                    if reported_expense_total < 0:
                        raise ValueError("The reported expense total cannot be negative.")
                continue
            if not any((item, amount, date)):
                continue
            if not item or not amount or not date:
                raise ValueError(f"Incomplete expense in '{sheet_name}' row {row_number}.")
            expenses.append({
                "sourceSheet": sheet_name,
                "date": parse_date(date, sheet_name, row_number),
                "item": item,
                "amount": parse_amount(amount, sheet_name, row_number),
                "paidTo": value_for(row, columns, "paidTo"),
                "remarks": value_for(row, columns, "remarks"),
            })

    totals = {
        "contributions": sum((Decimal(row["amount"]) for row in receipts), Decimal(0)),
        "stallFees": sum((Decimal(row["amount"]) for row in stalls), Decimal(0)),
        "expenses": sum((Decimal(row["amount"]) for row in expenses), Decimal(0)),
    }
    if reported_expense_total is not None and totals["expenses"] != reported_expense_total:
        raise ValueError(
            f"Expense rows total {totals['expenses']}, but the workbook's Total Expenses row says {reported_expense_total}."
        )
    totals = {key: int(value) if value == value.to_integral_value() else float(value) for key, value in totals.items()}
    if unrecognized:
        print(f"Skipped tabs without recognizable payment headers: {', '.join(unrecognized)}", file=sys.stderr)
    return {
        "sourceSheets": list(sheets),
        "receipts": receipts,
        "stalls": stalls,
        "expenses": expenses,
        "totals": totals,
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", nargs="?", type=Path, default=DEFAULT_INPUT,
                        help=f"source .xlsx file (default: {DEFAULT_INPUT})")
    parser.add_argument("event_id", nargs="?", help="event identifier (defaults to a slug from the workbook filename)")
    parser.add_argument("--name", help="event name shown in the event selector")
    parser.add_argument("--output", type=Path, help="optional explicit JSON output path")
    args = parser.parse_args()

    if not args.workbook.is_file():
        raise FileNotFoundError(f"Workbook not found: {args.workbook}")
    if args.workbook.suffix.lower() != ".xlsx":
        raise ValueError("Input must be an .xlsx workbook.")

    event_id = args.event_id or re.sub(r"[^a-z0-9]+", "-", args.workbook.stem.lower()).strip("-")
    if not re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", event_id):
        raise ValueError("Event ID must contain lowercase letters, numbers, and single hyphens only.")
    event_name = args.name or re.sub(r"[_-]+", " ", event_id).title()
    year_match = re.search(r"(?:19|20)\d{2}", f"{event_name} {event_id}")
    event_year = int(year_match.group()) if year_match else dt.date.today().year
    data = parse_rows(load_sheets(args.workbook))
    data = {"id": event_id, "name": event_name, "year": event_year, **data}
    output = args.output or EVENTS_DIR / f"{event_id}.json"
    if output.parent.resolve() == EVENTS_DIR.resolve() and output.exists():
        print(f"Updating existing event data: {output.name}")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(
        f"Imported {len(data['receipts'])} receipts, {len(data['stalls'])} stall payments, "
        f"and {len(data['expenses'])} expenses. Totals: contributions ₹{data['totals']['contributions']:,}; "
        f"stall fees ₹{data['totals']['stallFees']:,}; expenses ₹{data['totals']['expenses']:,}."
    )
    print(f"Event: {event_name} ({event_year}), saved to {output}.")


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, KeyError, zipfile.BadZipFile) as error:
        print(f"Import failed: {error}", file=sys.stderr)
        sys.exit(1)
