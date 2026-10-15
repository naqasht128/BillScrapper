"""Desktop Bill Scrapper.

The selected .xlsx or .ods workbook is updated in place using its existing headings.
Excluded fields are removed before requests are made, and returned values are
written strictly below their matching headings without column shifting.
"""

from __future__ import annotations

import hashlib
import io
import json
import os
import platform
import queue
import re
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import uuid
import webbrowser
import zipfile
from datetime import datetime, timezone

try:
    import winreg
except ImportError:
    winreg = None
from decimal import Decimal
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import requests
from bs4 import BeautifulSoup
from lxml import etree
from openpyxl import load_workbook

URL = "https://bill.pitc.com.pk/pescobill"
LICENSE_API_URL = os.environ.get("LICENSE_API_URL", "https://bill-scrapper.vercel.app").rstrip("/")


def resource_path(relative_path):
    """Resolve bundled resources in source runs, PyInstaller, and Nuitka onefile builds."""
    bundle_root = os.environ.get("NUITKA_ONEFILE_DIRECTORY") or getattr(sys, "_MEIPASS", None) or Path(__file__).resolve().parent
    return str(Path(bundle_root) / relative_path)

# Fields that are intentionally never requested/exported.
DROP_LETTERS = {"B", "C", "E", "H", "K", "L", "M", "N", "P", "Q", "S"}
DROP_RANGES = ((24, 50), (28, 55), (36, 64), (67, 67), (98, 214), (217, 244))  # X:AX, AB:BC, AJ:BL, BO, CT:HF and HI:IJ
DROP_EXACT = {
    "reference_no", "consumer_id", "sub_division", "tariff_cat", "transformer",
    "energy_charges", "taxes", "arrears", "installment",
    "tracking_id",
    "total_electricity_charges", "net_electricity_charges", "subsidies", "category", "26",
    "kwh_units_44", "kwh_units_45", "kwh_units_46",
    "table_1_kwh_type_1",
    "table_1_kvarh_meter_1", "table_1_kvarh_meter_2", "table_1_kvarh_meter_3", "table_1_kvarh_meter_4",
    "table_1_kvarh_type_1", "table_1_mdi_meter_1", "table_1_mdi_meter_2", "table_1_mdi_meter_3", "table_1_mdi_meter_4",
    "table_1_mdi_type_1", "table_1_mdi_previous_1",
    "table_3_787_1", "table_3_787_off_peak_1", "table_3_787_peak_1",
    "table_3_2402_1", "table_3_2402_off_peak_1", "table_3_2402_peak_1",
    "table_3_1615_1", "table_3_1615_off_peak_1", "table_3_1615_peak_1",
}


def excel_col(n):
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def excel_numeric_value(value):
    """Convert bill numbers such as '41,839' to values Excel treats as numeric."""
    if value is None or isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value
    text = clean(value).replace(",", "")
    text = re.sub(r"\s*(?:kw|kva)\s*$", "", text, flags=re.I)
    if not re.fullmatch(r"[-+]?(?:\d+(?:\.\d+)?|\.\d+)", text):
        return value
    number = float(text)
    return int(number) if number.is_integer() else number


def should_be_numeric(key):
    """Identify bill fields that should be written as numbers, not text."""
    normalized = normalized_key(key).lower()
    if normalized in {
        "payment", "grand_total", "adjustments", "san_load",
        "current bill", "current_bill", "total fpa", "total_fpa",
        "prv_cum_mdi_1", "prv_cum_mdi_2", "cum_mdi_1", "cum_mdi_2",
    }:
        return True
    return bool(re.fullmatch(r"(?:kwh|kvarh|mdi)_(?:present|read|mf|units)(?:_\d+)?", normalized))


def clean(value):
    return re.sub(r"\s+", " ", str(value) if value is not None else "").strip()


def field_key(value):
    return re.sub(r"[^a-zA-Z0-9]+", "_", clean(value).lower()).strip("_") or "field"


def normalized_key(value):
    key = field_key(value)
    renamed = {
        "kwh_units_30": "kvarh_units_1",
        "kwh_units_31": "mdi_read_1", "kwh_units_32": "mdi_read_2",
        "kwh_units_33": "mdi_read_3", "kwh_units_34": "mdi_read_4",
        "kwh_units_35": "mdi_mf_1", "kwh_units_36": "mdi_mf_2",
        "kwh_units_37": "mdi_mf_3", "kwh_units_38": "mdi_mf_4",
        "kwh_units_39": "mdi_units_1",
        "kwh_units_40": "prv_cum_mdi_1", "kwh_units_41": "prv_cum_mdi_2",
        "kwh_units_42": "cum_mdi_1", "kwh_units_43": "cum_mdi_2",
        "mdi_present_1": "mdi_read_1", "mdi_present_2": "mdi_read_2",
        "mdi_present_3": "mdi_read_3", "mdi_present_4": "mdi_read_4",
        "prv_cumm_mdi1": "prv_cum_mdi_1", "prv_cumm_mdi2": "prv_cum_mdi_2",
        "prv_cumm_mdi_1": "prv_cum_mdi_1", "prv_cumm_mdi_2": "prv_cum_mdi_2",
        "prv_cum_mdi1": "prv_cum_mdi_1", "prv_cum_mdi2": "prv_cum_mdi_2",
        "previous_cumm_mdi1": "prv_cum_mdi_1", "previous_cumm_mdi2": "prv_cum_mdi_2",
        "previous_cumm_mdi_1": "prv_cum_mdi_1", "previous_cumm_mdi_2": "prv_cum_mdi_2",
        "previous_cum_mdi1": "prv_cum_mdi_1", "previous_cum_mdi2": "prv_cum_mdi_2",
        "previous_cum_mdi_1": "prv_cum_mdi_1", "previous_cum_mdi_2": "prv_cum_mdi_2",
        "cum_mdi1": "cum_mdi_1", "cum_mdi2": "cum_mdi_2",
        "prs_cum_mdi1": "cum_mdi_1", "prs_cum_mdi2": "cum_mdi_2",
        "prs_cumm_mdi1": "cum_mdi_1", "prs_cumm_mdi2": "cum_mdi_2",
        "diff_cum_mdi_1": "Diff_cum_mdi1", "diff_cum_mdi1": "Diff_cum_mdi1",
        "diff_cum_mdi_2": "diff_cum_mdi2", "diff_cum_mdi2": "diff_cum_mdi2",
        "actual_mdi": "Actual MDI",
        "actual_fix_charges": "Actual_Fix_Charges",
        "in_bill_fix_charges": "In_Bill_Fix_Charges (Estimated)",
        "in_bill_fix_charges_estimated": "In_Bill_Fix_Charges (Estimated)",
        "current_bill": "Current Bill",
        "current bill": "Current Bill",
        "total_fpa": "Total FPA",
        "total fpa": "Total FPA",
        "kvarh_units": "kvarh_units_1",
        "mdi_units": "mdi_units_1",
        "scrape_status": "Remarks",
        "remarks": "Remarks",
    }
    if key in renamed:
        return renamed[key]
    if key == "table_3_787_mdi_1":
        return "export_mdi"
    if key == "table_3_2402_mdi_1":
        return "import_mdi"
    key = re.sub(r"^table_\d+_", "", key)
    return re.sub(r"^\d+(?:_\d+)*_", "", key)


def is_preserved_meter_key(key):
    k_lower = key.lower()
    return (
        k_lower.startswith("kvarh_present_")
        or k_lower.startswith("kvarh_units")
        or k_lower.startswith("kwh_units")
        or k_lower.startswith("mdi_read_")
        or k_lower.startswith("mdi_mf_")
        or k_lower.startswith("mdi_units")
        or k_lower.startswith("prv_cum_mdi_")
        or k_lower.startswith("cum_mdi_")
        or k_lower == "status"
        or k_lower.startswith("in_bill_fix_charges")
        or k_lower in {
            "diff_cum_mdi1", "diff_cum_mdi2",
            "actual mdi", "actual_mdi",
            "actual_fix_charges",
            "current bill", "current_bill",
            "total fpa", "total_fpa",
        }
    )


def is_disc_status(status_str):
    """Check if status is Disconnected (P.Disc or T.Disc), where minus and actual fixed charges are skipped."""
    if not status_str:
        return False
    s = str(status_str).lower().strip()
    if re.search(r"\bp\.?\s*disc", s) or "p-disc" in s or "pdisc" in s or "permanent" in s:
        return True
    if re.search(r"\bt\.?\s*disc", s) or "t-disc" in s or "tdisc" in s or "temporary" in s:
        return True
    return False


def is_replaced_status(status_str):
    """Check if status is Replaced."""
    if not status_str:
        return False
    s = str(status_str).lower().strip()
    return "replace" in s or re.search(r"\brep\b", s)


def is_dropped(name, source_col=None, legacy_columns=False):
    key = field_key(name)
    normalized = normalized_key(name)
    if is_preserved_meter_key(normalized) or is_preserved_meter_key(key):
        return False
    if key in {"remarks", "scrape_status"} or normalized == "Remarks":
        return False
    if source_col is not None:
        in_dropped_range = any(start <= source_col <= end for start, end in DROP_RANGES)
        if (legacy_columns and excel_col(source_col) in DROP_LETTERS) or in_dropped_range:
            return True
    match = re.fullmatch(r"kwh_units_(\d+)", key)
    if match and int(match.group(1)) > 4:
        return True
    return (
        key in DROP_EXACT
        or "previous" in key
        or "bill_history" in key
        or key.startswith(("kvarh_meter_", "kvarh_mf_", "mdi_meter_"))
        or key.startswith("bill_history")
        or key.startswith(("table_4_", "table_5_", "table_6_"))
        or key in {"account_no", "raw_bill_text"}
    )


GENERAL_LABELS = {
    "san load": "san_load",
    "san load kw": "san_load",
    "tariff": "tariff",
    "status": "status",
    "feeder": "feeder",
    "name address": "name_address",
    "name and address": "name_address",
    "energy charges": "energy_charges",
    "taxes": "taxes",
    "current bill": "Current Bill",
    "current bill amount": "Current Bill",
    "total fpa": "Total FPA",
    "total f p a": "Total FPA",
    "arrears": "arrears",
    "installment": "installment",
    "adjustment": "adjustments",
    "adjustments": "adjustments",
    "grand total": "grand_total",
    "amount paid": "payment",
    "paid amount": "payment",
    "payment": "payment",
    "payment date": "payment_date",
    "ntn no": "ntn_no",
    "bill month": "bill_month",
    "due date": "due_date",
    "transformer": "transformer",
}


def canonical_label(value):
    label = re.sub(r"[^a-z0-9]+", " ", clean(value).lower()).strip()
    return GENERAL_LABELS.get(label, normalized_key(value))


def order_output_headers(headers):
    """Keep tariff/san-load stable and keep meter readings under fixed names."""
    first = [
        "AC No.", "tariff", "san_load", "status", "Current Bill", "Total FPA", "adjustments", "grand_total",
        "payment", "payment_date",
        "ntn_no", "bill_month", "name_address", "due_date",
    ]
    ordered = [name for name in first if name in headers]
    meter_order = [
        *(f"kwh_meter_{i}" for i in range(1, 5)),
        *(f"kwh_present_{i}" for i in range(1, 5)),
        *(f"kwh_mf_{i}" for i in range(1, 5)),
        *(f"kwh_units_{i}" for i in range(1, 5)),
        *(f"kvarh_present_{i}" for i in range(1, 5)),
        *(f"kvarh_units_{i}" for i in range(1, 5)),
        *(f"mdi_read_{i}" for i in range(1, 5)),
        *(f"mdi_mf_{i}" for i in range(1, 5)),
        *(f"mdi_units_{i}" for i in range(1, 5)),
        *(f"prv_cum_mdi_{i}" for i in range(1, 5)),
        *(f"cum_mdi_{i}" for i in range(1, 5)),
        "Diff_cum_mdi1",
        "diff_cum_mdi2",
        "Actual MDI",
        "Actual_Fix_Charges",
        "In_Bill_Fix_Charges (Estimated)",
    ]
    ordered += [name for name in meter_order if name in headers and name not in ordered]
    ordered += [name for name in headers if name not in ordered and name not in {"feeder", "Remarks", "scrape_status"}]
    if "feeder" in headers:
        ordered.append("feeder")
    if "Remarks" in headers or "scrape_status" in headers:
        ordered.append("Remarks")
    return ordered


def text_lines(element):
    """Parse distinct value lines from a table cell, separating (O)/(P) and space-separated units."""
    for br in element.find_all("br"):
        br.replace_with("\n")
    text = element.get_text("\n")
    # Insert newline before indicators like (O), (P), (T) if preceded by text/digits
    t = re.sub(r"(?<=\S)\s+(\([A-Za-z0-9]+\))", r"\n\1", text)
    raw_lines = [clean(x) for x in t.splitlines() if clean(x)]
    result = []
    for line in raw_lines:
        # If line contains multiple space-separated numbers without letters (e.g. '1673 610')
        nums = re.findall(r"[-+]?(?:\d+(?:\.\d+)?|\.\d+)", line)
        if len(nums) > 1 and not re.search(r"[A-Za-z]", line):
            result.extend(nums)
        else:
            result.append(line)
    return result


OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
SUPPORTED_EXTENSIONS = {".xlsx", ".ods"}


def _extract_ods_cell_value(cell):
    """Extract cell value from an ODF <table:table-cell> element."""
    val_type = cell.get(f"{{{OFFICE_NS}}}value-type")
    p_els = cell.findall(f"{{{TEXT_NS}}}p")
    text_val = "\n".join("".join(p.itertext()) for p in p_els).strip() if p_els else ""

    if val_type in {"float", "currency", "percentage"}:
        if re.fullmatch(r"0\d+", text_val):
            return text_val
        raw_val = cell.get(f"{{{OFFICE_NS}}}value")
        if raw_val is not None:
            try:
                dec = Decimal(raw_val)
                return int(dec) if dec == dec.to_integral_value() else float(dec)
            except Exception:
                pass
        return text_val if text_val else None

    if val_type == "boolean":
        bool_val = cell.get(f"{{{OFFICE_NS}}}boolean-value")
        if bool_val is not None:
            return bool_val.lower() == "true"
        return text_val if text_val else None

    if val_type in {"date", "time"}:
        if text_val:
            return text_val
        return cell.get(f"{{{OFFICE_NS}}}date-value") or cell.get(f"{{{OFFICE_NS}}}time-value")

    return text_val if text_val != "" else None


def _iter_ods_rows(container):
    """Yield <table:table-row> elements in document order without descending into nested tables."""
    row_tag = f"{{{TABLE_NS}}}table-row"
    group_tags = {
        f"{{{TABLE_NS}}}table-row-group",
        f"{{{TABLE_NS}}}table-header-rows",
        f"{{{TABLE_NS}}}table-rows",
    }
    for child in container:
        if child.tag == row_tag:
            yield child
        elif child.tag in group_tags:
            yield from _iter_ods_rows(child)


class _ODSCell:
    def __init__(self, sheet: ODSSheet, row: int, col: int):
        self._sheet = sheet
        self._row = row
        self._col = col

    @property
    def value(self):
        return self._sheet._cells.get((self._row, self._col))

    @value.setter
    def value(self, val):
        if val is None or val == "":
            self._sheet._cells.pop((self._row, self._col), None)
        else:
            self._sheet._cells[(self._row, self._col)] = val


class ODSSheet:
    """Worksheet-compatible interface for an OpenDocument Spreadsheet (.ods) table."""

    def __init__(self):
        self._cells = {}

    @property
    def max_row(self):
        return max((r for (r, _) in self._cells), default=0)

    @property
    def max_column(self):
        return max((c for (_, c) in self._cells), default=0)

    def cell(self, row: int, column: int) -> _ODSCell:
        return _ODSCell(self, row, column)

    def delete_cols(self, idx: int, amount: int = 1):
        if amount <= 0:
            return
        updated = {}
        for (r, c), val in self._cells.items():
            if c < idx:
                updated[(r, c)] = val
            elif c >= idx + amount:
                updated[(r, c - amount)] = val
        self._cells = updated

    def delete_rows(self, idx: int, amount: int = 1):
        if amount <= 0:
            return
        updated = {}
        for (r, c), val in self._cells.items():
            if r < idx:
                updated[(r, c)] = val
            elif r >= idx + amount:
                updated[(r - amount, c)] = val
        self._cells = updated


class ODSWorkbook:
    """Workbook-compatible reader/writer for LibreOffice / ODF .ods files."""

    def __init__(self, path: Path):
        try:
            with zipfile.ZipFile(path, "r") as zf:
                self._entries = [(info, zf.read(info.filename)) for info in zf.infolist()]
        except zipfile.BadZipFile as exc:
            raise RuntimeError(f"Invalid or corrupted .ods file: {path.name}") from exc

        content_bytes = next((data for info, data in self._entries if info.filename == "content.xml"), None)
        if content_bytes is None:
            raise RuntimeError(f"Invalid .ods file (missing content.xml): {path.name}")

        self._root = etree.fromstring(content_bytes)
        self._table_el = self._root.find(f".//{{{TABLE_NS}}}table")
        if self._table_el is None:
            raise RuntimeError(f"No spreadsheet sheet found in {path.name}")

        self.active = ODSSheet()
        self._load_cells()

    def _load_cells(self):
        cell_tags = {f"{{{TABLE_NS}}}table-cell", f"{{{TABLE_NS}}}covered-table-cell"}
        col_rep_attr = f"{{{TABLE_NS}}}number-columns-repeated"
        row_rep_attr = f"{{{TABLE_NS}}}number-rows-repeated"

        current_row = 1
        for row_el in _iter_ods_rows(self._table_el):
            try:
                row_repeat = max(1, int(row_el.get(row_rep_attr, "1")))
            except ValueError:
                row_repeat = 1

            col_idx = 1
            row_vals = {}
            for cell_el in row_el:
                if cell_el.tag not in cell_tags:
                    continue
                try:
                    col_repeat = max(1, int(cell_el.get(col_rep_attr, "1")))
                except ValueError:
                    col_repeat = 1

                val = _extract_ods_cell_value(cell_el)
                if val is not None and str(val) != "":
                    for offset in range(min(col_repeat, 1024)):
                        row_vals[col_idx + offset] = val
                col_idx += col_repeat

            if row_vals:
                for r_offset in range(min(row_repeat, 10000)):
                    for c, v in row_vals.items():
                        self.active._cells[(current_row + r_offset, c)] = v
            current_row += row_repeat

    def save(self, path: Path):
        max_r = self.active.max_row
        max_c = self.active.max_column

        row_containers = {
            f"{{{TABLE_NS}}}table-row",
            f"{{{TABLE_NS}}}table-row-group",
            f"{{{TABLE_NS}}}table-header-rows",
            f"{{{TABLE_NS}}}table-rows",
        }
        col_containers = {
            f"{{{TABLE_NS}}}table-column",
            f"{{{TABLE_NS}}}table-column-group",
            f"{{{TABLE_NS}}}table-header-columns",
            f"{{{TABLE_NS}}}table-columns",
        }

        insert_idx = None
        for idx, child in enumerate(list(self._table_el)):
            if child.tag in row_containers:
                if insert_idx is None:
                    insert_idx = idx
                self._table_el.remove(child)

        if insert_idx is None:
            insert_idx = 0
            for idx, child in enumerate(self._table_el):
                if child.tag in col_containers:
                    insert_idx = idx + 1

        col_tag = f"{{{TABLE_NS}}}table-column"
        col_rep_attr = f"{{{TABLE_NS}}}number-columns-repeated"
        col_els = list(self._table_el.iter(col_tag))
        covered_cols = 0
        for el in col_els:
            try:
                covered_cols += max(1, int(el.get(col_rep_attr, "1")))
            except ValueError:
                covered_cols += 1

        if max_c > covered_cols:
            extra = max_c - covered_cols
            if col_els:
                last_col = col_els[-1]
                try:
                    last_rep = max(1, int(last_col.get(col_rep_attr, "1")))
                except ValueError:
                    last_rep = 1
                last_col.set(col_rep_attr, str(last_rep + extra))
            else:
                new_col = etree.Element(col_tag)
                new_col.set(col_rep_attr, str(max_c))
                self._table_el.insert(0, new_col)
                insert_idx += 1

        row_tag = f"{{{TABLE_NS}}}table-row"
        cell_tag = f"{{{TABLE_NS}}}table-cell"
        p_tag = f"{{{TEXT_NS}}}p"
        val_type_attr = f"{{{OFFICE_NS}}}value-type"
        val_attr = f"{{{OFFICE_NS}}}value"
        bool_attr = f"{{{OFFICE_NS}}}boolean-value"

        for r in range(1, max_r + 1):
            row_el = etree.Element(row_tag)
            for c in range(1, max_c + 1):
                cell_el = etree.SubElement(row_el, cell_tag)
                val = self.active.cell(r, c).value
                if val is None or val == "":
                    continue
                if isinstance(val, bool):
                    cell_el.set(val_type_attr, "boolean")
                    cell_el.set(bool_attr, "true" if val else "false")
                    p_el = etree.SubElement(cell_el, p_tag)
                    p_el.text = "TRUE" if val else "FALSE"
                elif isinstance(val, (int, float, Decimal)):
                    cell_el.set(val_type_attr, "float")
                    cell_el.set(val_attr, str(val))
                    p_el = etree.SubElement(cell_el, p_tag)
                    p_el.text = str(val)
                else:
                    cell_el.set(val_type_attr, "string")
                    for line in str(val).split("\n"):
                        p_el = etree.SubElement(cell_el, p_tag)
                        p_el.text = line
            self._table_el.insert(insert_idx, row_el)
            insert_idx += 1

        new_content = etree.tostring(self._root, xml_declaration=True, encoding="UTF-8")

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf_out:
            for info, data in self._entries:
                if info.filename == "mimetype":
                    zf_out.writestr("mimetype", data, compress_type=zipfile.ZIP_STORED)
            for info, data in self._entries:
                if info.filename == "mimetype":
                    continue
                payload = new_content if info.filename == "content.xml" else data
                zf_out.writestr(info.filename, payload, compress_type=zipfile.ZIP_DEFLATED)

        Path(path).write_bytes(buf.getvalue())


def load_spreadsheet(path: Path):
    ext = path.suffix.lower()
    if ext == ".xlsx":
        return load_workbook(path)
    if ext == ".ods":
        return ODSWorkbook(path)
    raise RuntimeError("Please select an .xlsx or .ods workbook. The selected workbook is updated in place.")


class PITCBillScraper:
    def __init__(self, delay=2.0, timeout=90):
        self.delay = delay
        self.timeout = timeout

    @staticmethod
    def put(output, name, value):
        value = clean(value)
        if not value or is_dropped(name):
            return
        key = canonical_label(name)

        if key == "adjustments":
            numeric_adjustment = re.search(r"[-+]?\d[\d,]*(?:\.\d+)?", value)
            if not numeric_adjustment:
                return
            value = numeric_adjustment.group(0)

        if re.search(r"(?:meter|present|mdi|kwh|kvarh|reading)", key, re.I):
            value = re.sub(r"^\s*[IE]\s+", "", value, flags=re.I).strip()

        if "units" in key:
            value = re.sub(r"\(\s*[OP]\s*\)\s*", "", value, flags=re.I).strip()

        # If a units field receives multiple space-separated numbers (e.g. '1673 610'),
        # automatically distribute them into index and index+1 so they are never merged into one cell.
        units_match = re.match(r"^(kwh|kvarh|mdi)_units_?(\d+)?$", key)
        if units_match:
            prefix = units_match.group(1)
            start_idx = int(units_match.group(2)) if units_match.group(2) else 1
            nums = [clean(n) for n in re.findall(r"[-+]?(?:\d+(?:\.\d+)?|\.\d+)", value)]
            if len(nums) > 1:
                for offset, num in enumerate(nums):
                    target_key = f"{prefix}_units_{start_idx + offset}"
                    existing = output.get(target_key)
                    if existing is None or (str(existing).strip() in {"0", "0.0", "0.00"} and str(num).strip() not in {"0", "0.0", "0.00"}):
                        output[target_key] = num
                return

        if value:
            existing = output.get(key)
            if existing is None or (str(existing).strip() in {"0", "0.0", "0.00"} and str(value).strip() not in {"0", "0.0", "0.00"}):
                output[key] = value

    @staticmethod
    def put_general_pairs(output, soup):
        """Read label/value pairs from both MDI and ordinary bill layouts."""
        known = set(GENERAL_LABELS)
        for row in soup.select("tr"):
            cells = [clean(x.get_text(" ")) for x in row.select("th, td") if clean(x.get_text(" "))]
            if len(cells) >= 2:
                label = cells[0].lower().rstrip(":")
                if label in known:
                    PITCBillScraper.put(output, cells[0], cells[-1])
        for node in soup.select(".grid-col-cell, .charges-bd-row, .right-section-cell, .ibn-bill-messages__cell"):
            bits = [clean(x) for x in node.stripped_strings if clean(x)]
            if len(bits) >= 2 and bits[0].lower().rstrip(":") in known:
                PITCBillScraper.put(output, bits[0], bits[-1])

    @staticmethod
    def put_general_text_pairs(output, soup):
        """Fallback for general batches whose labels are not table cells."""
        main = soup.select_one(".gbn-app-shell") or soup.body or soup
        if not main:
            return
        labels = set(GENERAL_LABELS)
        lines = [clean(x) for x in main.stripped_strings if clean(x)]
        for index, line in enumerate(lines):
            label = re.sub(r"[^a-z0-9]+", " ", line.lower()).strip()
            if label not in labels:
                continue
            key = GENERAL_LABELS[label]
            if is_dropped(key):
                continue
            if key in output:
                existing = str(output[key])
                if key == "san_load" and re.search(r"[A-Za-z]-?\d", existing):
                    del output[key]
                elif key == "tariff" and not re.search(r"[A-Za-z]-?\d", existing):
                    del output[key]
                else:
                    continue
            for candidate in lines[index + 1:index + 4]:
                candidate_key = re.sub(r"[^a-z0-9]+", " ", candidate.lower()).strip()
                if candidate_key in labels or not candidate:
                    continue
                if any(ord(ch) > 127 for ch in candidate) and not re.search(r"\d", candidate):
                    continue
                if key == "san_load" and not re.search(r"\d", candidate):
                    continue
                if key == "tariff" and not re.search(r"[A-Za-z]-?\d", candidate):
                    continue
                PITCBillScraper.put(output, key, candidate)
                break

    @staticmethod
    def put_meter_tables(output, soup):
        """Map both MDI and ordinary meter tables to stable reading headings."""
        header_aliases = {
            "meter": "meter", "meter no": "meter", "meter no.": "meter", "meter number": "meter",
            "present": "present", "present reading": "present", "present readings": "present",
            "mf": "mf", "units": "units", "unit": "units", "type": None,
            "previous": None, "previous reading": None, "previous readings": None,
        }
        for table in soup.select("table"):
            headers = [clean(x.get_text(" ")).lower() for x in table.select("thead th")]
            if not headers:
                headers = [clean(x.get_text(" ")).lower() for x in table.select("tr:first-child th, tr:first-child td")]
            mapped = [header_aliases.get(re.sub(r"\s+", " ", h).strip()) for h in headers]
            if not any(x in {"meter", "present", "mf", "units"} for x in mapped):
                continue
            rows = table.select("tbody tr") or table.select("tr")[1:]
            carry = None
            for row in rows:
                cells = row.select("th, td")
                if not cells:
                    continue
                if headers and len(cells) == len(headers) - 1 and carry is not None:
                    cells = [carry] + cells
                if cells and headers and headers[0] in {"meter", "meter no", "meter no."}:
                    carry = cells[0]
                row_text = " ".join(clean(x.get_text(" ")) for x in cells[:2]).lower()
                prefix = "kvarh" if "kvarh" in row_text or "kvar h" in row_text else "mdi" if re.search(r"\bmdi\b", row_text) else "kwh"
                for col_no, cell in enumerate(cells):
                    kind = mapped[col_no] if col_no < len(mapped) else None
                    if kind is None:
                        continue
                    for value_no, value in enumerate(text_lines(cell), 1):
                        self_value = re.sub(r"\b[IE]\s+(?=\d)", "", value, flags=re.I).strip()
                        PITCBillScraper.put(output, f"{prefix}_{kind}_{value_no}", self_value)

    @staticmethod
    def put_meter_text_sections(output, soup):
        """Fallback for general bills that render meter data as div/text blocks."""
        main = soup.select_one(".gbn-app-shell") or soup.body or soup
        lines = [clean(x) for x in main.stripped_strings if clean(x)]
        lowered = [re.sub(r"[^a-z0-9]+", " ", x.lower()).strip() for x in lines]
        try:
            start = next(i for i, value in enumerate(lowered) if value in {"meter info", "meter information"})
        except StopIteration:
            return
        end = next((i for i in range(start + 1, len(lines)) if lowered[i] in {"bill charges breakdown", "bill history"}), len(lines))
        section_lines = lines[start:end]
        section_keys = [re.sub(r"[^a-z0-9]+", " ", x.lower()).strip() for x in section_lines]
        if not any(value in {"meter no", "meter number"} for value in section_keys):
            return
        labels = {
            "meter no": "meter", "meter number": "meter", "meter": "meter",
            "present reading": "present", "present readings": "present",
            "present": "present", "mf": "mf", "units": "units", "unit": "units",
        }
        for index, label in enumerate(section_keys):
            kind = labels.get(label)
            if not kind:
                continue
            values = []
            for candidate in section_lines[index + 1:]:
                candidate_key = re.sub(r"[^a-z0-9]+", " ", candidate.lower()).strip()
                if candidate_key in labels or candidate_key in {"previous", "previous reading", "type", "meter info"}:
                    break
                if any(ord(ch) > 127 for ch in candidate) and not re.search(r"\d", candidate):
                    continue
                if kind == "meter" and not re.search(r"\d", candidate):
                    continue
                if kind != "meter" and not re.search(r"[-+]?\d", candidate):
                    continue
                values.append(re.sub(r"\b[IE]\s+(?=\d)", "", candidate, flags=re.I).strip())
            for value_no, value in enumerate(values, 1):
                PITCBillScraper.put(output, f"kwh_{kind}_{value_no}", value)

    @staticmethod
    def put_cum_mdi(output, soup):
        """Read previous and present cumulative MDI values."""
        # 1. Check explicit MDI meter footer cells if present
        for cell in soup.select(".mdi-meter-footer-cell"):
            lbl_el = cell.select_one(".en-lbl")
            val_el = cell.select_one(".mdi-meter-footer-val")
            if lbl_el and val_el:
                lbl = clean(lbl_el.get_text(" ")).upper()
                vals = [clean(x) for x in text_lines(val_el) if re.fullmatch(r"[-+]?(?:\d+(?:\.\d+)?|\.\d+)", clean(x))]
                if "PRV" in lbl and "CUM" in lbl:
                    for idx, val in enumerate(vals[:4], 1):
                        output[f"prv_cum_mdi_{idx}"] = val
                elif "PRS" in lbl and "CUM" in lbl:
                    for idx, val in enumerate(vals[:4], 1):
                        output[f"cum_mdi_{idx}"] = val

        # 2. Supplementary scan through stripped lines fallback
        main = soup.select_one(".gbn-app-shell") or soup.body or soup
        if not main:
            return
        lines = [clean(x) for x in main.stripped_strings if clean(x)]
        keys = [re.sub(r"[^a-z0-9]+", " ", x.lower()).strip() for x in lines]

        # PRV CUMM MDI
        prv_start = next((i for i, value in enumerate(keys) if value in {"prv cumm mdi", "previous cumm mdi", "previous cumulative mdi", "prv cum mdi"}), None)
        if prv_start is not None:
            values = []
            for candidate in lines[prv_start + 1:prv_start + 7]:
                if re.fullmatch(r"[-+]?(?:\d+(?:\.\d+)?|\.\d+)", candidate):
                    values.append(candidate)
                elif values:
                    break
            for value_no, value in enumerate(values[:4], 1):
                k = f"prv_cum_mdi_{value_no}"
                if k not in output:
                    output[k] = value

        # PRS CUMM MDI
        prs_start = next((i for i, value in enumerate(keys) if value in {"prs cumm mdi", "present cumm mdi", "present cumulative mdi", "prs cum mdi"}), None)
        if prs_start is not None:
            values = []
            for candidate in lines[prs_start + 1:prs_start + 7]:
                if re.fullmatch(r"[-+]?(?:\d+(?:\.\d+)?|\.\d+)", candidate):
                    values.append(candidate)
                elif values:
                    break
            for value_no, value in enumerate(values[:4], 1):
                k = f"cum_mdi_{value_no}"
                if k not in output:
                    output[k] = value

    @staticmethod
    def put_charges_breakdown(output, soup):
        """Extract Current Bill and Total FPA from charges breakdown."""
        # 1. Look for .charges-bd-row containers
        for row in soup.select(".charges-bd-row, tr"):
            row_text = clean(row.get_text(" ")).upper()

            # Current Bill (class=charges-bd-val)
            if "CURRENT BILL" in row_text:
                val_el = row.select_one(".charges-bd-val")
                if val_el:
                    v = clean(val_el.get_text(" "))
                    if v and re.search(r"\d", v):
                        output["Current Bill"] = v
                else:
                    bits = [clean(x) for x in row.stripped_strings if clean(x)]
                    if len(bits) >= 2 and re.search(r"\d", bits[-1]):
                        output["Current Bill"] = bits[-1]

            # Total FPA (class=charges-bd-val charges-bd-val--strong)
            if any(term in row_text for term in ("TOTAL F.P.A", "TOTAL FPA", "TOTAL F P A", "TOTAL FUEL PRICE")):
                val_el = (
                    row.select_one(".charges-bd-val.charges-bd-val--strong")
                    or row.select_one(".charges-bd-val--strong")
                    or row.select_one(".charges-bd-val")
                )
                if val_el:
                    v = clean(val_el.get_text(" "))
                    if v and re.search(r"\d", v):
                        output["Total FPA"] = v
                else:
                    bits = [clean(x) for x in row.stripped_strings if clean(x)]
                    if len(bits) >= 2 and re.search(r"\d", bits[-1]):
                        output["Total FPA"] = bits[-1]

        # 2. Target direct CSS classes if not yet populated
        if "Total FPA" not in output:
            for el in soup.select(".charges-bd-val--strong, .charges-bd-val.charges-bd-val--strong"):
                parent = el.find_parent(class_="charges-bd-row") or el.parent
                p_text = clean(parent.get_text(" ")).upper() if parent else ""
                if any(term in p_text for term in ("FPA", "F.P.A", "FUEL")):
                    v = clean(el.get_text(" "))
                    if v and re.search(r"\d", v):
                        output["Total FPA"] = v
                        break

        # 3. Fallback scan through stripped lines
        main = soup.select_one(".gbn-app-shell") or soup.body or soup
        if main and ("Current Bill" not in output or "Total FPA" not in output):
            lines = [clean(x) for x in main.stripped_strings if clean(x)]
            for idx, line in enumerate(lines):
                norm_line = re.sub(r"[^A-Z0-9]+", " ", line.upper()).strip()
                if "Current Bill" not in output and norm_line in {"CURRENT BILL", "CURRENT BILL AMOUNT"}:
                    for cand in lines[idx + 1:idx + 4]:
                        if re.search(r"[-+]?\d[\d,]*(?:\.\d+)?", cand):
                            output["Current Bill"] = cand
                            break
                if "Total FPA" not in output and ("TOTAL F P A" in norm_line or "TOTAL FPA" in norm_line):
                    for cand in lines[idx + 1:idx + 4]:
                        if re.search(r"[-+]?\d[\d,]*(?:\.\d+)?", cand):
                            output["Total FPA"] = cand
                            break

    def _fetch_url_with_retry(self, session, method, url, data=None, stop_event=None, log_callback=None):
        was_disconnected = False
        while True:
            if stop_event and stop_event.is_set():
                raise InterruptedError("Operation stopped by user")
            try:
                if method.lower() == "post":
                    resp = session.post(url, data=data, timeout=self.timeout)
                else:
                    resp = session.get(url, timeout=self.timeout)
                resp.raise_for_status()
                if was_disconnected and log_callback:
                    log_callback("Internet connection restored. Resuming bill processing...")
                return resp
            except (requests.exceptions.ConnectionError, requests.exceptions.Timeout, ConnectionResetError, OSError):
                was_disconnected = True
                if log_callback:
                    log_callback("Internet disconnected. Waiting for network connection to restore...")
                for _ in range(30):
                    if stop_event and stop_event.is_set():
                        raise InterruptedError("Operation stopped by user")
                    time.sleep(0.1)
            except requests.exceptions.HTTPError as exc:
                if exc.response is not None and 500 <= exc.response.status_code <= 504:
                    was_disconnected = True
                    if log_callback:
                        log_callback(f"PITC server busy (HTTP {exc.response.status_code}). Retrying in a moment...")
                    for _ in range(20):
                        if stop_event and stop_event.is_set():
                            raise InterruptedError("Operation stopped by user")
                        time.sleep(0.1)
                    continue
                raise

    def search(self, account, stop_event=None, log_callback=None):
        session = requests.Session()
        session.headers.update({"User-Agent": "Mozilla/5.0 (compatible; PITC-Bill-Exporter/1.0)"})
        page = self._fetch_url_with_retry(session, "get", URL, stop_event=stop_event, log_callback=log_callback)
        soup = BeautifulSoup(page.text, "lxml")
        form = soup.select_one("form#SubmitForm") or soup.find("form")
        if not form:
            raise RuntimeError("PITC search form was not found")
        payload = {i["name"]: i.get("value", "") for i in form.select("input[type=hidden][name]")}
        payload.update({"rbSearchByList": "refno", "searchTextBox": account, "ruCodeTextBox": "", "btnSearch": "Search"})
        result = self._fetch_url_with_retry(session, "post", URL, data=payload, stop_event=stop_event, log_callback=log_callback)
        if "CONSUMER DETAIL" not in result.text.upper() and "CONSUMER DETAILS" not in result.text.upper():
            raise RuntimeError(clean(BeautifulSoup(result.text, "lxml").get_text(" "))[:300] or "No bill returned")
        return BeautifulSoup(result.text, "lxml")

    def scrape_account(self, raw_account, existing_status=None, stop_event=None, log_callback=None):
        raw = str(raw_account).strip()
        if re.fullmatch(r"[0-9.]+e[+\-]?[0-9]+", raw, re.I):
            raw = format(Decimal(raw), "f")
        account = re.sub(r"\D", "", raw)
        if not re.fullmatch(r"\d{10,14}", account):
            return {"_account": account, "Remarks": "Bill Not Found", "scrape_status": "Bill Not Found"}
        try:
            soup = self.search(account, stop_event=stop_event, log_callback=log_callback)
            out = {"_account": account, "Remarks": "Processed", "scrape_status": "Processed"}
            self.put_general_pairs(out, soup)
            self.put_general_text_pairs(out, soup)
            for cell in soup.select(".grid-col-cell"):
                label, value = cell.select_one(".en-lbl"), cell.select_one(".val-space")
                if label and value:
                    self.put(out, label.get_text(" "), " | ".join(text_lines(value)))

            for node in soup.select(".charges-bd-row, .right-section-cell, .ibn-bill-messages__cell"):
                bits = [clean(x) for x in node.stripped_strings if clean(x)]
                if len(bits) >= 2:
                    self.put(out, bits[0], bits[-1])

            self.put_meter_tables(out, soup)
            self.put_meter_text_sections(out, soup)
            self.put_cum_mdi(out, soup)
            self.put_charges_breakdown(out, soup)

            # Final validation: ensure units fields don't have multiple numbers in a single key
            for prefix in ("kwh", "kvarh", "mdi"):
                for idx in range(1, 5):
                    k = f"{prefix}_units_{idx}"
                    if k in out and out[k]:
                        val_str = str(out[k])
                        nums = [clean(n) for n in re.findall(r"[-+]?(?:\d+(?:\.\d+)?|\.\d+)", val_str)]
                        if len(nums) > 1:
                            out[k] = nums[0]
                            next_k = f"{prefix}_units_{idx + 1}"
                            if next_k not in out or not out[next_k]:
                                out[next_k] = nums[1]

            def parse_num(val):
                if val is None:
                    return None
                cleaned = re.sub(r"[^\d.]", "", str(val))
                try:
                    return float(cleaned) if cleaned else None
                except ValueError:
                    return None

            def format_num(num, decimals=4):
                if num is None:
                    return None
                val = round(num, decimals)
                return int(val) if val.is_integer() else val

            # Check status:
            # 1. If P.Disc or T.Disc: skip subtraction, Actual MDI, and actual fixed charges.
            # 2. If Replaced: subtract (cum_mdi_2 - prv_cum_mdi_2) and (cum_mdi_4 - prv_cum_mdi_4).
            # 3. Normal / Other: subtract (cum_mdi_1 - prv_cum_mdi_1) and (cum_mdi_2 - prv_cum_mdi_2).
            status_val = out.get("status") or existing_status

            if is_disc_status(status_val):
                out["Diff_cum_mdi1"] = None
                out["diff_cum_mdi2"] = None
                out["Actual MDI"] = None
                out["Actual_Fix_Charges"] = None
            else:
                if is_replaced_status(status_val):
                    # Case Replaced: cumm mdi 2 - prv cumm mdi 2, and cumm mdi 4 - prv cumm mdi 4
                    prv2 = parse_num(out.get("prv_cum_mdi_2"))
                    cum2 = parse_num(out.get("cum_mdi_2"))
                    diff1 = format_num(abs(cum2 - prv2), 4) if (prv2 is not None and cum2 is not None) else None
                    if diff1 is not None:
                        out["Diff_cum_mdi1"] = diff1

                    prv4 = parse_num(out.get("prv_cum_mdi_4"))
                    cum4 = parse_num(out.get("cum_mdi_4"))
                    diff2 = format_num(abs(cum4 - prv4), 4) if (prv4 is not None and cum4 is not None) else None
                    if diff2 is not None:
                        out["diff_cum_mdi2"] = diff2
                else:
                    # Normal status: cumm mdi 1 - prv cumm mdi 1, and cumm mdi 2 - prv cumm mdi 2
                    prv1 = parse_num(out.get("prv_cum_mdi_1"))
                    cum1 = parse_num(out.get("cum_mdi_1"))
                    diff1 = format_num(abs(cum1 - prv1), 4) if (prv1 is not None and cum1 is not None) else None
                    if diff1 is not None:
                        out["Diff_cum_mdi1"] = diff1

                    prv2 = parse_num(out.get("prv_cum_mdi_2"))
                    cum2 = parse_num(out.get("cum_mdi_2"))
                    diff2 = format_num(abs(cum2 - prv2), 4) if (prv2 is not None and cum2 is not None) else None
                    if diff2 is not None:
                        out["diff_cum_mdi2"] = diff2

                # Actual MDI: max(Diff_cum_mdi1, diff_cum_mdi2) * mdi_mf_1
                diff_candidates = []
                for d in (out.get("Diff_cum_mdi1"), out.get("diff_cum_mdi2")):
                    if d is not None:
                        diff_candidates.append(float(d))

                if diff_candidates:
                    max_diff = max(diff_candidates)
                    mf = parse_num(out.get("mdi_mf_1"))
                    if mf is None or mf == 0:
                        mf = 1.0
                    actual_mdi = format_num(max_diff * mf, 4)
                    out["Actual MDI"] = actual_mdi

                    # Actual_Fix_Charges = Actual MDI x 1250
                    actual_fix = format_num(float(actual_mdi) * 1250, 2)
                    out["Actual_Fix_Charges"] = actual_fix

            # In_Bill_Fix_Charges (Estimated): mdi_units_1 multiplied by 1250
            mdi_u1 = parse_num(out.get("mdi_units_1") if "mdi_units_1" in out else out.get("mdi_units"))
            if mdi_u1 is not None:
                in_bill_fix = format_num(mdi_u1 * 1250, 2)
                out["In_Bill_Fix_Charges (Estimated)"] = in_bill_fix

            return out
        except InterruptedError:
            raise
        except Exception as exc:
            return {"_account": account, "Remarks": "Bill Not Found", "scrape_status": "Bill Not Found"}

    def update_workbook(self, path, progress=None, stop_event=None, log_callback=None):
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise RuntimeError("Please select an .xlsx or .ods workbook. The selected workbook is updated in place.")
        workbook = load_spreadsheet(path)
        sheet = workbook.active

        # 1. Read existing headers and existing rows into record dictionaries
        first = str(sheet.cell(1, 1).value or "").strip().lower().rstrip(".")
        has_heading = first in {"ac", "ac no", "account", "account no"}
        data_start_row = 2 if has_heading else 1

        old_max_col = sheet.max_column
        old_max_row = sheet.max_row

        existing_col_names = {}
        if has_heading:
            for col in range(1, old_max_col + 1):
                raw = sheet.cell(1, col).value
                if raw:
                    norm = "AC No." if col == 1 else normalized_key(raw)
                    if norm.lower() in ("scrape_status", "scrape status"):
                        norm = "Remarks"
                    existing_col_names[col] = norm

        row_records = []
        for r in range(data_start_row, old_max_row + 1):
            raw_ac = sheet.cell(r, 1).value
            if raw_ac is None or not str(raw_ac).strip():
                continue
            record = {}
            for col in range(1, old_max_col + 1):
                val = sheet.cell(r, col).value
                col_name = existing_col_names.get(col)
                if col_name and val is not None:
                    record[col_name] = val
            record["AC No."] = raw_ac
            row_records.append(record)

        total = len(row_records)
        all_keys = set()
        for col_name in existing_col_names.values():
            if not is_dropped(col_name) and col_name not in ("scrape_status", "scrape status"):
                all_keys.add(col_name)

        # 2. Process each account and merge results into record dictionary
        try:
            for done, record in enumerate(row_records, 1):
                if stop_event and stop_event.is_set():
                    break
                account = record.get("AC No.")
                existing_status = record.get("status")
                try:
                    result = self.scrape_account(account, existing_status=existing_status, stop_event=stop_event, log_callback=log_callback)
                except InterruptedError:
                    break
                record.update(result)

                # Double-check status condition on merged record
                effective_status = record.get("status") or existing_status
                if is_disc_status(effective_status):
                    record["Diff_cum_mdi1"] = None
                    record["diff_cum_mdi2"] = None
                    record["Actual MDI"] = None
                    record["Actual_Fix_Charges"] = None

                for k in result:
                    if not k.startswith("_") and not is_dropped(k) and k not in ("scrape_status", "scrape status"):
                        all_keys.add(k)
                if progress:
                    progress(done, total, record)

                if stop_event and stop_event.is_set():
                    break

                slept = 0.0
                delay_step = 0.1
                while slept < self.delay:
                    if stop_event and stop_event.is_set():
                        break
                    time.sleep(delay_step)
                    slept += delay_step
        finally:
            # 3. Build ordered final headers and save all progress made up to now
            all_keys.add("AC No.")
            all_keys.add("feeder")
            all_keys.add("adjustments")
            all_keys.add("Current Bill")
            all_keys.add("Total FPA")
            all_keys.add("Remarks")
            all_keys.discard("scrape_status")
            all_keys.discard("scrape status")

            final_headers = order_output_headers([k for k in all_keys if not is_dropped(k)])

            # 4. Write data strictly by column name - eliminating column shifting completely
            if sheet.max_column > len(final_headers):
                sheet.delete_cols(len(final_headers) + 1, sheet.max_column - len(final_headers))
            if sheet.max_row > len(row_records) + 1:
                sheet.delete_rows(len(row_records) + 2, sheet.max_row - (len(row_records) + 1))

            for col_idx, h in enumerate(final_headers, 1):
                sheet.cell(1, col_idx).value = h

            for row_idx, record in enumerate(row_records, 2):
                for col_idx, h in enumerate(final_headers, 1):
                    val = record.get(h)
                    if h == "Remarks":
                        raw_val = val if val is not None else record.get("scrape_status")
                        if raw_val is not None:
                            val_str = str(raw_val).strip().lower()
                            val = "Processed" if val_str in ("processed", "ok", "done") else "Bill Not Found"
                        else:
                            val = ""
                    elif val is not None and should_be_numeric(h):
                        val = excel_numeric_value(val)
                    sheet.cell(row_idx, col_idx).value = val if val is not None else ""

            workbook.save(path)


# System-adaptive font families
UI_FONT = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"
MONO_FONT = "Consolas" if sys.platform == "win32" else "DejaVu Sans Mono"

# Light-Dark Palette (Antigravity IDE / VS Code Dark+ theme)
BG_DARK = "#1e1e1e"            # Soft neutral IDE dark gray
CARD_BG = "#252526"            # IDE container / card panel background
BORDER_COLOR = "#383838"       # Subtle, soft border without harsh contrast
TEXT_WHITE = "#ffffff"         # Crisp white for titles & section headers (AGY IDE style)
TEXT_MAIN = "#cccccc"          # Comfortable silver/off-white text (low glare)
TEXT_MUTED = "#8b949e"         # Soft secondary text
TEXT_HINT = "#767676"          # Muted hint text
ACCENT_GREEN = "#238636"       # Soft matte green (GitHub/VS Code style)
ACCENT_GREEN_HOVER = "#2ea043"
ACCENT_RED = "#b83a3a"         # Soft matte red
ACCENT_RED_HOVER = "#c93b3b"
INPUT_BG = "#2d2d2d"           # Soft dark input field (not pitch black)
BTN_SEC_BG = "#3a3d41"         # Standard IDE secondary button
BTN_SEC_HOVER = "#45494e"      # Secondary button hover
BTN_DISABLED_BG = "#282828"    # Disabled button background
BTN_DISABLED_FG = "#666666"    # Disabled button text
LOG_BG = "#1e1e1e"             # Same as IDE code editor canvas
LOG_FG = "#d4d4d4"             # Clean readable terminal text

# Version & Secure License/Update server endpoint (keeps GitHub repo 100% private)
APP_VERSION = "1.3.1"
UPDATE_API_URL = f"{LICENSE_API_URL}/api/updates/latest"


def parse_version_tuple(v: str) -> tuple[int, ...]:
    """Parse a version string (e.g. 'v1.0.2') into an integer tuple for robust comparisons."""
    nums = [int(n) for n in re.findall(r"\d+", str(v))]
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums)


class UpdateInfo:
    """Metadata regarding an available release from the secure update server."""
    def __init__(self, version: str, download_url: str, file_name: str, file_size: int, release_notes: str = ""):
        self.version = version
        self.download_url = download_url
        self.file_name = file_name
        self.file_size = file_size
        self.release_notes = release_notes


def check_for_updates(current_version: str = APP_VERSION) -> UpdateInfo | None:
    """Query secure licensing & update API to check if a newer version is available.
    
    This calls our private Vercel backend, ensuring our GitHub repository
    and Python source code remain 100% private and protected.
    """
    try:
        headers = {
            "User-Agent": "BillScrapper-DesktopApp",
            "Accept": "application/json",
        }
        resp = requests.get(UPDATE_API_URL, headers=headers, timeout=8)
        if resp.status_code != 200:
            return None
        data = resp.json()
        target_version = data.get("version", "").strip().lstrip("v")
        download_url = data.get("download_url", "").strip()
        if not target_version or not download_url:
            return None

        # Compare versions
        if parse_version_tuple(target_version) <= parse_version_tuple(current_version):
            return None

        file_name = data.get("file_name", "BillScrapper-Setup.exe")
        return UpdateInfo(
            version=target_version,
            download_url=download_url,
            file_name=file_name,
            file_size=data.get("file_size", 0),
            release_notes=data.get("release_notes", ""),
        )
    except Exception:
        return None


def install_and_restart(new_file_path: str, root_window: tk.Tk):
    """Run downloaded Setup installer or replace executable and restart."""
    is_frozen = getattr(sys, "frozen", False)

    if not is_frozen:
        # Running from Python source (.py in development)
        msg = (
            f"Update downloaded successfully to:\n{new_file_path}\n\n"
            "Because you are running from Python source code, you can update via 'git pull' "
            "or run the downloaded Setup installer directly."
        )
        messagebox.showinfo("Update Downloaded", msg)
        try:
            if sys.platform == "win32":
                subprocess.Popen(["explorer.exe", f"/select,{new_file_path}"])
        except Exception:
            pass
        return

    # Running as compiled binary
    current_exe = os.path.abspath(sys.executable)
    pid = os.getpid()

    if sys.platform == "win32":
        is_setup_installer = "setup" in os.path.basename(new_file_path).lower()
        if not is_setup_installer and os.path.exists(new_file_path):
            try:
                with open(new_file_path, "rb") as f:
                    chunk = f.read(65536)
                    if b"Inno Setup" in chunk or b"InnoSetup" in chunk:
                        is_setup_installer = True
            except Exception:
                pass

        # Strip all PyInstaller, Python, and Tk environment variables to prevent child process conflicts
        clean_env = os.environ.copy()
        keys_to_remove = [
            k for k in clean_env
            if k.startswith('_PYI')
            or k.startswith('_MEI')
            or k.startswith('PYTHON')
            or k.startswith('TCL')
            or k.startswith('TK')
        ]
        for k in keys_to_remove:
            clean_env.pop(k, None)

        if hasattr(sys, '_MEIPASS'):
            meipass = sys._MEIPASS.lower()
            path_dirs = clean_env.get('PATH', '').split(os.pathsep)
            path_dirs = [p for p in path_dirs if p.lower().rstrip(r'\/') != meipass and '_mei' not in p.lower()]
            clean_env['PATH'] = os.pathsep.join(path_dirs)

        if is_setup_installer:
            # Clean Inno Setup silent install: closes app, updates files, and restarts
            try:
                DETACHED_FLAGS = 0x08000000 | 0x00000200
                subprocess.Popen(
                    [new_file_path, "/SILENT", "/CLOSEAPPLICATIONS", "/FORCECLOSEAPPLICATIONS", "/AUTOUPDATE=1"],
                    creationflags=DETACHED_FLAGS,
                    close_fds=True,
                    env=clean_env,
                )
            except Exception:
                subprocess.Popen([new_file_path, "/CLOSEAPPLICATIONS", "/FORCECLOSEAPPLICATIONS", "/AUTOUPDATE=1"], env=clean_env)
            try:
                root_window.destroy()
            except Exception:
                pass
            os._exit(0)

        # Fallback for standalone portable executable replacement
        bat_path = Path(tempfile.gettempdir()) / f"billscrapper_patch_{pid}.bat"
        bat_content = f"""@echo off
setlocal
set PID={pid}
set NEW_EXE="{new_file_path}"
set TARGET_EXE="{current_exe}"

:wait_pid
timeout /t 1 /nobreak >nul 2>&1
tasklist /FI "PID eq %PID%" 2>nul | findstr /I "%PID%" >nul
if not errorlevel 1 goto wait_pid

timeout /t 1 /nobreak >nul 2>&1

copy /y %NEW_EXE% %TARGET_EXE% >nul 2>&1
if errorlevel 1 (
    move /y %NEW_EXE% %TARGET_EXE% >nul 2>&1
)

start "" %TARGET_EXE%
del %NEW_EXE% >nul 2>&1
(goto) 2>nul & del "%~f0"
"""
        with open(bat_path, "w", encoding="utf-8") as f:
            f.write(bat_content)

        DETACHED_FLAGS = 0x08000000 | 0x00000200
        subprocess.Popen(
            ["cmd.exe", "/c", str(bat_path)],
            creationflags=DETACHED_FLAGS,
            close_fds=True,
            env=clean_env,
        )
        try:
            root_window.destroy()
        except Exception:
            pass
        sys.exit(0)

    else:
        # Linux / Unix
        os.chmod(new_file_path, 0o755)
        sh_path = Path(tempfile.gettempdir()) / f"billscrapper_patch_{pid}.sh"
        sh_content = f"""#!/bin/sh
PID={pid}
NEW_BIN="{new_file_path}"
TARGET_BIN="{current_exe}"

while kill -0 $PID 2>/dev/null; do
    sleep 0.5
done

mv -f "$NEW_BIN" "$TARGET_BIN"
chmod +x "$TARGET_BIN"
"$TARGET_BIN" &
rm -f "$0"
"""
        with open(sh_path, "w", encoding="utf-8") as f:
            f.write(sh_content)
        os.chmod(sh_path, 0o755)

        subprocess.Popen(["/bin/sh", str(sh_path)], close_fds=True)
        root_window.destroy()
        sys.exit(0)


def set_dark_title_bar(window: tk.Tk | tk.Toplevel):
    """Enable native Windows 10/11 immersive dark mode title bar."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        window.update_idletasks()
        hwnds = []
        try:
            hwnds.append(int(window.wm_frame(), 16))
        except Exception:
            pass
        try:
            parent_h = ctypes.windll.user32.GetParent(window.winfo_id())
            if parent_h:
                hwnds.append(parent_h)
        except Exception:
            pass
        hwnds.append(window.winfo_id())

        value = ctypes.c_int(1)
        for h in set(hwnds):
            if h:
                # DWMWA_USE_IMMERSIVE_DARK_MODE: 20 for Win 11 & Win 10 (20H1+), 19 for older Win 10
                if ctypes.windll.dwmapi.DwmSetWindowAttribute(h, 20, ctypes.byref(value), ctypes.sizeof(value)) != 0:
                    ctypes.windll.dwmapi.DwmSetWindowAttribute(h, 19, ctypes.byref(value), ctypes.sizeof(value))
    except Exception:
        pass


class UpdateDialog(tk.Toplevel):
    """Modal dialog displaying download progress and applying the update."""
    def __init__(self, parent: "BillScraperApp", info: UpdateInfo):
        super().__init__(parent)
        self.parent = parent
        self.info = info
        self.cancelled = False
        self.temp_file: str | None = None

        self.withdraw()
        self.title("Updating Bill Scrapper")
        self.configure(bg=BG_DARK)
        self.resizable(False, False)
        self.transient(parent)

        try:
            if sys.platform == "win32":
                icon_file = resource_path("assets/billscrapper.ico")
                if not os.path.exists(icon_file):
                    icon_file = resource_path("assets/mdiscrapper.ico")
                if os.path.exists(icon_file):
                    self.iconbitmap(icon_file)
            else:
                png_file = resource_path("assets/billscrapper-icon.png")
                if os.path.exists(png_file):
                    img = tk.PhotoImage(file=png_file)
                    self.iconphoto(False, img)
        except Exception:
            pass

        # Center over parent
        w, h = 420, 200
        px = parent.winfo_rootx() + (parent.winfo_width() // 2) - (w // 2)
        py = parent.winfo_rooty() + (parent.winfo_height() // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{max(0, px)}+{max(0, py)}")

        container = ttk.Frame(self, style="Root.TFrame", padding=16)
        container.pack(fill="both", expand=True)

        lbl_title = ttk.Label(
            container,
            text=f"Downloading Update ({self.info.version})...",
            style="Title.TLabel",
            font=(UI_FONT, 12, "bold"),
        )
        lbl_title.pack(anchor="w", pady=(0, 6))

        self.lbl_status = ttk.Label(
            container,
            text="Connecting to update server...",
            style="Root.TLabel",
        )
        self.lbl_status.pack(anchor="w", pady=(0, 8))

        self.prog = ttk.Progressbar(container, mode="determinate", style="Green.Horizontal.TProgressbar")
        self.prog.pack(fill="x", pady=(0, 10))

        self.lbl_details = ttk.Label(
            container,
            text="0% (0.0 MB / 0.0 MB)",
            style="Sub.TLabel",
        )
        self.lbl_details.pack(anchor="w", pady=(0, 12))

        btn_box = ttk.Frame(container, style="Root.TFrame")
        btn_box.pack(fill="x")

        self.cancel_btn = ttk.Button(btn_box, text="Cancel", command=self.cancel)
        self.cancel_btn.pack(side="right")

        self.queue = queue.Queue()
        self.protocol("WM_DELETE_WINDOW", self.cancel)

        set_dark_title_bar(self)
        self.update_idletasks()
        self.deiconify()
        self.grab_set()

        self._poll_queue()
        threading.Thread(target=self._download_worker, daemon=True).start()

    def _poll_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                kind = msg[0]
                if kind == "progress":
                    self._update_progress(*msg[1:])
                elif kind == "complete":
                    self._on_download_complete()
                elif kind == "error":
                    self._on_error(msg[1])
                elif kind == "cancelled":
                    if self.winfo_exists():
                        self.destroy()
        except queue.Empty:
            pass
        if self.winfo_exists():
            self.after(50, self._poll_queue)

    def cancel(self):
        self.cancelled = True
        self.lbl_status.configure(text="Cancelling download...")
        self.cancel_btn.configure(state="disabled")

    def _download_worker(self):
        try:
            fname = self.info.file_name or ("BillScrapper-Setup.exe" if sys.platform == "win32" else "BillScrapper")
            if sys.platform == "win32" and not fname.lower().endswith(".exe"):
                fname += ".exe"
            tpath = Path(tempfile.gettempdir()) / f"update_{int(time.time())}_{fname}"
            self.temp_file = str(tpath)

            headers = {"User-Agent": "BillScrapper-DesktopApp"}
            with requests.get(self.info.download_url, headers=headers, stream=True, timeout=30) as r:
                r.raise_for_status()
                total = int(r.headers.get("content-length", self.info.file_size or 0))
                downloaded = 0

                with open(tpath, "wb") as f:
                    for chunk in r.iter_content(chunk_size=65536):
                        if self.cancelled:
                            break
                        if chunk:
                            f.write(chunk)
                            downloaded += len(chunk)
                            if total > 0:
                                pct = int((downloaded / total) * 100)
                                d_mb = downloaded / (1024 * 1024)
                                t_mb = total / (1024 * 1024)
                                self.queue.put(("progress", pct, d_mb, t_mb))

            if self.cancelled:
                try:
                    if os.path.exists(tpath):
                        os.remove(tpath)
                except Exception:
                    pass
                self.queue.put(("cancelled",))
                return

            self.queue.put(("complete",))
        except Exception as exc:
            self.queue.put(("error", str(exc)))

    def _update_progress(self, pct, downloaded_mb, total_mb):
        if not self.winfo_exists():
            return
        self.prog.configure(maximum=100, value=pct)
        self.lbl_status.configure(text="Downloading update package...")
        self.lbl_details.configure(text=f"{pct}% ({downloaded_mb:.1f} MB / {total_mb:.1f} MB)")

    def _on_download_complete(self):
        if not self.winfo_exists():
            return
        self.prog.configure(maximum=100, value=100)
        self.lbl_status.configure(text="Download complete! Restarting to apply update...")
        self.lbl_details.configure(text="Please wait a moment...")
        self.cancel_btn.configure(state="disabled")
        self.after(800, self._apply_update)

    def _apply_update(self):
        if self.temp_file and os.path.exists(self.temp_file):
            install_and_restart(self.temp_file, self.parent)

    def _on_error(self, err_msg):
        if not self.winfo_exists():
            return
        messagebox.showerror("Update Error", f"Failed to download update:\n{err_msg}", parent=self)
        self.destroy()


# Global hardware device ID cache to make all subsequent queries instantaneous (0.00ms)
_CACHED_DEVICE_ID: str | None = None
_CACHED_CANDIDATE_IDS: list[str] | None = None


def get_device_id() -> str:
    """
    Returns a persistent, unique hardware identifier for this machine.
    Uses Windows Registry MachineGuid as primary (64-bit and 32-bit views),
    falling back to modern PowerShell CIM query for BIOS UUID (Windows 11 compatible,
    even when WMIC is deprecated/removed), legacy WMIC, physical MAC node,
    or Darwin/Linux machine UUID.
    Guaranteed never to crash or flash console windows.
    """
    global _CACHED_DEVICE_ID
    if _CACHED_DEVICE_ID:
        return _CACHED_DEVICE_ID

    device_id = _compute_device_id()
    _CACHED_DEVICE_ID = device_id
    return device_id


def _compute_device_id() -> str:
    if sys.platform == "win32":
        # 1. Windows Registry MachineGuid (both 64-bit and 32-bit registry views)
        for flag in [getattr(winreg, 'KEY_WOW64_64KEY', 0), getattr(winreg, 'KEY_WOW64_32KEY', 0), 0]:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography", 0, winreg.KEY_READ | flag) as key:
                    machine_guid, _ = winreg.QueryValueEx(key, "MachineGuid")
                    if machine_guid and len(str(machine_guid).strip()) > 10:
                        return str(machine_guid).strip().upper()
            except Exception:
                pass

        # 2. Modern PowerShell CIM query for BIOS UUID (works reliably on Windows 10 & 11)
        try:
            cmd = ["powershell", "-NoProfile", "-Command", "(Get-CimInstance Win32_ComputerSystemProduct).UUID"]
            creation_flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
            output = subprocess.check_output(cmd, creationflags=creation_flags, timeout=5).decode(errors="ignore").strip()
            if '-' in output and len(output) >= 30:
                return output.upper()
        except Exception:
            pass

        # 3. Legacy WMIC BIOS UUID (for older Windows versions)
        try:
            cmd = ["wmic", "csproduct", "get", "uuid"]
            creation_flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
            output = subprocess.check_output(cmd, creationflags=creation_flags, timeout=4).decode(errors="ignore").splitlines()
            for line in output:
                clean_l = line.strip()
                if '-' in clean_l and len(clean_l) >= 30:
                    return clean_l.upper()
        except Exception:
            pass

        # 4. Fallback to physical MAC address
        try:
            import uuid
            mac = uuid.getnode()
            if (mac >> 40) % 2 == 0:
                return f"MAC-{hex(mac)[2:].upper()}"
        except Exception:
            pass

        return "UNKNOWN-HWID-WIN"

    elif sys.platform == "darwin":
        try:
            out = subprocess.check_output(
                ["ioreg", "-rd1", "-c", "IOPlatformExpertDevice"],
                stderr=subprocess.DEVNULL,
                text=True,
                timeout=3,
            )
            m = re.search(r'"IOPlatformUUID"\s*=\s*"([^"]+)"', out)
            if m:
                return m.group(1).strip().upper()
        except Exception:
            pass

    else:
        # Linux / Unix machine-id
        for path in ("/etc/machine-id", "/var/lib/dbus/machine-id"):
            if os.path.exists(path):
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        val = f.read().strip()
                        if val:
                            return val.upper()
                except Exception:
                    pass

    # Generic fallback using MAC node
    try:
        import uuid
        mac = uuid.getnode()
        return f"FALLBACK-{hex(mac)[2:].upper()}"
    except Exception:
        pass

    return "UNKNOWN-DEVICE-ID"


def get_legacy_device_id() -> str:
    """Generate legacy SHA-256 fingerprint for grandfathered activations."""
    components = [sys.platform]
    if sys.platform == "win32":
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography") as key:
                guid, _ = winreg.QueryValueEx(key, "MachineGuid")
                if guid:
                    components.append(str(guid).strip().lower())
        except Exception:
            pass
        try:
            cmd = ["wmic", "csproduct", "get", "uuid"]
            creation_flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
            out = subprocess.check_output(cmd, creationflags=creation_flags, timeout=3).decode(errors="ignore")
            lines = [l.strip() for l in out.splitlines() if l.strip()]
            if len(lines) >= 2 and lines[1]:
                components.append(lines[1].strip().lower())
        except Exception:
            pass
    try:
        import uuid
        node = uuid.getnode()
        if (node >> 40) % 2 == 0:
            components.append(str(node))
    except Exception:
        pass
    raw = "|".join(components) if len(components) > 1 else f"fallback_{sys.platform}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def get_candidate_device_ids() -> list[str]:
    """Return all valid device IDs for this machine: [primary_id, legacy_hash]."""
    global _CACHED_CANDIDATE_IDS
    if _CACHED_CANDIDATE_IDS:
        return _CACHED_CANDIDATE_IDS

    primary = get_device_id()
    candidates = [primary]
    try:
        legacy = get_legacy_device_id()
        if legacy and legacy not in candidates:
            candidates.append(legacy)
    except Exception:
        pass

    _CACHED_CANDIDATE_IDS = candidates
    return candidates


def get_os_info() -> str:
    """Return human-readable OS and architecture details."""
    try:
        return f"{platform.system()} {platform.release()} ({platform.machine()})"
    except Exception:
        return sys.platform


def get_license_file_path() -> Path:
    """Return local path to license.json in the user's OS configuration folder."""
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
        app_dir = base / "BillScrapper"
    else:
        app_dir = Path.home() / ".billscrapper"
    app_dir.mkdir(parents=True, exist_ok=True)
    return app_dir / "license.json"


def load_local_license() -> dict:
    """Load cached license details from local disk."""
    path = get_license_file_path()
    if path.is_file():
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_local_license(data: dict):
    """Save active license details to local disk."""
    path = get_license_file_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


def verify_license_remote(serial_key: str, device_id: str, candidate_ids: list[str] | None = None) -> dict:
    """Check license status with Render Supabase API."""
    url = f"{LICENSE_API_URL.rstrip('/')}/api/verify"
    c_ids = candidate_ids or [device_id]
    payload = {
        "serial_key": serial_key.strip().upper(),
        "device_id": device_id,
        "candidate_ids": c_ids,
        "os_info": get_os_info(),
    }
    try:
        resp = requests.post(url, json=payload, timeout=12)
        try:
            data = resp.json()
        except Exception:
            data = {"valid": False, "status": "network_error", "message": f"Server HTTP {resp.status_code}"}

        if data.get("valid"):
            return data

        # If server returned limit_reached or not valid on primary ID, check legacy fallback ID
        if data.get("status") in ("limit_reached", "invalid") and len(c_ids) > 1:
            for alt_id in c_ids[1:]:
                alt_payload = {
                    "serial_key": serial_key.strip().upper(),
                    "device_id": alt_id,
                    "os_info": get_os_info(),
                }
                try:
                    alt_resp = requests.post(url, json=alt_payload, timeout=12)
                    alt_data = alt_resp.json()
                    if alt_data.get("valid"):
                        return alt_data
                except Exception:
                    pass

        return data
    except Exception as e:
        return {"valid": False, "status": "network_error", "message": f"Connection error: {e}"}


def activate_license_remote(serial_key: str, device_id: str, candidate_ids: list[str] | None = None) -> dict:
    """Activate license with Render Supabase API."""
    url = f"{LICENSE_API_URL.rstrip('/')}/api/activate"
    c_ids = candidate_ids or [device_id]
    payload = {
        "serial_key": serial_key.strip().upper(),
        "device_id": device_id,
        "candidate_ids": c_ids,
        "os_info": get_os_info(),
    }
    try:
        resp = requests.post(url, json=payload, timeout=12)
        try:
            return resp.json()
        except Exception:
            return {"valid": False, "status": "network_error", "message": f"Server HTTP {resp.status_code}"}
    except Exception as e:
        return {"valid": False, "status": "network_error", "message": f"Connection error: {e}"}


def format_expiry_date(iso_str: str | None) -> str:
    """Format an ISO timestamp cleanly, e.g. '14 Jan 2027 (89 days remaining)'."""
    if not iso_str:
        return "Not Activated"
    try:
        clean_ts = str(iso_str).replace("Z", "+00:00")
        dt = datetime.fromisoformat(clean_ts)
        now = datetime.now(timezone.utc)
        diff = dt - now
        date_str = dt.strftime("%d %b %Y")
        if diff.days > 0:
            return f"{date_str} ({diff.days} days remaining)"
        elif diff.total_seconds() > 0:
            return f"{date_str} (Expires today)"
        else:
            return f"{date_str} (Expired)"
    except Exception:
        return str(iso_str)[:10]


def format_plan_type(plan_code: str | None) -> str:
    """Convert machine plan code to human friendly label."""
    mapping = {
        "1_month": "1 Month (1 Device)",
        "3_months": "3 Months (Up to 2 Devices)",
        "6_months": "6 Months (Up to 2 Devices)",
        "1_year": "1 Year (Up to 2 Devices)",
    }
    if not plan_code:
        return "Not Activated"
    return mapping.get(str(plan_code).lower(), str(plan_code).replace("_", " ").title())


class LicenseGateDialog(tk.Toplevel):
    """Sleek VS Code Dark+ styled License Activation Gate & Subscription Dialog."""

    def __init__(self, parent: "BillScraperApp", initial_error: str = ""):
        super().__init__(parent)
        self.withdraw()
        self.parent = parent
        self.title("Bill Scrapper — License Activation")
        self.configure(bg=BG_DARK)
        self.resizable(False, False)
        self.transient(parent)

        self.activated = False
        self.license_info: dict | None = None
        self.device_id = get_device_id()
        self.queue = queue.Queue()

        # Exit application cleanly if closed without activation
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        w, h = 540, 620
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        px = max(0, (sw - w) // 2)
        py = max(0, (sh - h) // 2)
        self.geometry(f"{w}x{h}+{px}+{py}")

        try:
            if sys.platform == "win32":
                icon_file = resource_path("assets/billscrapper.ico")
                if not os.path.exists(icon_file):
                    icon_file = resource_path("assets/mdiscrapper.ico")
                if os.path.exists(icon_file):
                    self.iconbitmap(icon_file)
        except Exception:
            pass

        self._build_ui(initial_error)
        set_dark_title_bar(self)
        self.update_idletasks()
        self.deiconify()
        self.grab_set()

        self._poll_queue()

    def _poll_queue(self):
        try:
            while True:
                msg_type, payload = self.queue.get_nowait()
                if msg_type == "result":
                    self._on_activation_result(payload)
                elif msg_type == "error":
                    self._on_activation_error(payload)
        except queue.Empty:
            pass
        if self.winfo_exists():
            self.after(50, self._poll_queue)

    def _build_ui(self, initial_error: str):
        self.container = ttk.Frame(self, style="Root.TFrame", padding=(22, 16, 22, 16))
        self.container.pack(fill="both", expand=True)

        # Header
        top_bar = ttk.Frame(self.container, style="Root.TFrame")
        top_bar.pack(fill="x", pady=(0, 12))

        ttk.Label(top_bar, text="Bill Scrapper", style="Title.TLabel", font=(UI_FONT, 15, "bold")).pack(anchor="w")
        ttk.Label(top_bar, text="License Activation & Subscription Gate", style="Sub.TLabel").pack(anchor="w", pady=(2, 0))

        # Tab Navigation (Enter Key vs Buy)
        nav_frame = ttk.Frame(self.container, style="Root.TFrame")
        nav_frame.pack(fill="x", pady=(0, 12))

        self.btn_tab_activate = tk.Button(
            nav_frame,
            text="🔑  Enter Serial Key",
            font=(UI_FONT, 9, "bold"),
            relief="flat",
            cursor="hand2",
            padx=16,
            pady=6,
            command=self._show_activate_tab,
        )
        self.btn_tab_activate.pack(side="left", fill="x", expand=True, padx=(0, 4))

        self.btn_tab_buy = tk.Button(
            nav_frame,
            text="🛒  Buy Subscription",
            font=(UI_FONT, 9, "bold"),
            relief="flat",
            cursor="hand2",
            padx=16,
            pady=6,
            command=self._show_buy_tab,
        )
        self.btn_tab_buy.pack(side="left", fill="x", expand=True, padx=(4, 0))

        # Main Dynamic Content Area
        self.tab_container = ttk.Frame(self.container, style="Root.TFrame")
        self.tab_container.pack(fill="both", expand=True)

        self.tab_activate = ttk.Frame(self.tab_container, style="Root.TFrame")
        self.tab_buy = ttk.Frame(self.tab_container, style="Root.TFrame")

        self._build_activate_tab(initial_error)
        self._build_buy_tab()

        # Default to activate tab
        self._show_activate_tab()

    def _build_activate_tab(self, initial_error: str):
        # Hardware ID Card
        hw_card = ttk.Frame(self.tab_activate, style="Card.TFrame", padding=(14, 10))
        hw_card.pack(fill="x", pady=(0, 10))

        ttk.Label(hw_card, text="MACHINE HARDWARE ID", style="CardTitle.TLabel", font=(UI_FONT, 8, "bold")).pack(anchor="w")

        hw_row = ttk.Frame(hw_card, style="Card.TFrame")
        hw_row.pack(fill="x", pady=(4, 0))

        self.hw_entry = tk.Entry(
            hw_row,
            font=(MONO_FONT, 9, "bold"),
            bg=INPUT_BG,
            fg="#4ec9b0",
            insertbackground=TEXT_WHITE,
            relief="flat",
            justify="center",
            highlightthickness=1,
            highlightbackground=BORDER_COLOR,
            highlightcolor=ACCENT_GREEN,
        )
        self.hw_entry.insert(0, self.device_id)
        self.hw_entry.configure(state="readonly")
        self.hw_entry.pack(side="left", fill="x", expand=True, ipady=4, padx=(0, 8))

        self.copy_id_btn = tk.Button(
            hw_row,
            text="📋 Copy ID",
            bg=BTN_SEC_BG,
            fg=TEXT_MAIN,
            activebackground=BTN_SEC_HOVER,
            activeforeground=TEXT_WHITE,
            font=(UI_FONT, 8, "bold"),
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=3,
            command=self._copy_device_id,
        )
        self.copy_id_btn.pack(side="right")

        # Serial Key Input Card
        key_card = ttk.Frame(self.tab_activate, style="Card.TFrame", padding=(16, 14))
        key_card.pack(fill="x", pady=(0, 12))

        ttk.Label(key_card, text="ENTER SERIAL KEY", style="CardTitle.TLabel", font=(UI_FONT, 9, "bold")).pack(anchor="w")
        ttk.Label(key_card, text="Format: BS-XXXX-XXXX-XXXX (provided upon purchase)", style="Sub.TLabel", font=(UI_FONT, 8)).pack(anchor="w", pady=(2, 8))

        self.key_var = tk.StringVar()
        saved = load_local_license().get("serial_key", "")
        if saved:
            self.key_var.set(saved)

        self.key_entry = tk.Entry(
            key_card,
            textvariable=self.key_var,
            font=(MONO_FONT, 12, "bold"),
            bg=INPUT_BG,
            fg="#4ec9b0",
            insertbackground=TEXT_WHITE,
            relief="flat",
            justify="center",
            highlightthickness=1,
            highlightbackground=BORDER_COLOR,
            highlightcolor=ACCENT_GREEN,
        )
        self.key_entry.pack(fill="x", ipady=8, pady=(0, 10))
        self.key_entry.bind("<Return>", lambda e: self._do_activate())

        self.activate_btn = tk.Button(
            key_card,
            text="⚡  ACTIVATE LICENSE",
            bg=ACCENT_GREEN,
            fg=TEXT_WHITE,
            activebackground=ACCENT_GREEN_HOVER,
            activeforeground=TEXT_WHITE,
            font=(UI_FONT, 10, "bold"),
            relief="flat",
            cursor="hand2",
            pady=8,
            command=self._do_activate,
        )
        self.activate_btn.pack(fill="x")

        # Status Feedback Card
        self.status_card = ttk.Frame(self.tab_activate, style="Card.TFrame", padding=(14, 10))
        self.status_card.pack(fill="both", expand=True, pady=(0, 10))

        self.status_var = tk.StringVar()
        self.status_lbl = tk.Label(
            self.status_card,
            textvariable=self.status_var,
            font=(UI_FONT, 9),
            bg=CARD_BG,
            fg=TEXT_MUTED,
            wraplength=460,
            justify="left",
            anchor="nw",
        )
        self.status_lbl.pack(fill="both", expand=True)

        if initial_error:
            self._set_status(f"⚠️ {initial_error}", color="#f14c4c")
        else:
            self._set_status("Enter your serial key above and click 'Activate License'.", color=TEXT_MUTED)

        # Bottom Switcher Link
        help_row = ttk.Frame(self.tab_activate, style="Root.TFrame")
        help_row.pack(fill="x")

        ttk.Label(help_row, text="Don't have a serial key?", style="Sub.TLabel").pack(side="left")
        buy_link = tk.Button(
            help_row,
            text="View Subscription Plans & Buy →",
            bg=BG_DARK,
            fg="#58a6ff",
            activebackground=BG_DARK,
            activeforeground="#79c0ff",
            font=(UI_FONT, 9, "underline"),
            relief="flat",
            bd=0,
            cursor="hand2",
            command=self._show_buy_tab,
        )
        buy_link.pack(side="left", padx=(6, 0))

    def _build_buy_tab(self):
        # 1. Subscription Plans Card
        plans_card = ttk.Frame(self.tab_buy, style="Card.TFrame", padding=(16, 12))
        plans_card.pack(fill="x", pady=(0, 10))

        ttk.Label(plans_card, text="SUBSCRIPTION PLANS & PRICING", style="CardTitle.TLabel", font=(UI_FONT, 9, "bold")).pack(anchor="w", pady=(0, 8))

        plans_data = [
            ("1 Month", "Rs. 300", "1 Device", "#89d185"),
            ("3 Months", "Rs. 700", "Up to 2 Devices", "#79c0ff"),
            ("6 Months", "Rs. 1,200", "Up to 2 Devices", "#d2a8ff"),
            ("1 Year", "Rs. 2,000", "Up to 2 Devices", "#ffa657"),
        ]

        for plan_name, price, devices, pill_color in plans_data:
            p_row = tk.Frame(plans_card, bg=INPUT_BG, padx=10, pady=6, highlightthickness=1, highlightbackground=BORDER_COLOR)
            p_row.pack(fill="x", pady=2)

            tk.Label(p_row, text=plan_name, font=(UI_FONT, 9, "bold"), fg=TEXT_WHITE, bg=INPUT_BG).pack(side="left")
            tk.Label(p_row, text=f"•  {devices}", font=(UI_FONT, 8), fg=TEXT_MUTED, bg=INPUT_BG).pack(side="left", padx=(8, 0))
            tk.Label(p_row, text=price, font=(UI_FONT, 9, "bold"), fg=pill_color, bg=INPUT_BG).pack(side="right")

        # 2. Contact for Purchase Card
        contact_card = ttk.Frame(self.tab_buy, style="Card.TFrame", padding=(16, 12))
        contact_card.pack(fill="both", expand=True, pady=(0, 10))

        ttk.Label(contact_card, text="CONTACT FOR PURCHASE & ACTIVATION", style="CardTitle.TLabel", font=(UI_FONT, 9, "bold")).pack(anchor="w", pady=(0, 4))
        ttk.Label(contact_card, text="Send your Machine Hardware ID via WhatsApp or Email to receive your key.", style="Sub.TLabel", font=(UI_FONT, 8)).pack(anchor="w", pady=(0, 8))

        # WhatsApp Row
        wa_frame = tk.Frame(contact_card, bg=INPUT_BG, padx=10, pady=8, highlightthickness=1, highlightbackground=BORDER_COLOR)
        wa_frame.pack(fill="x", pady=(0, 6))

        wa_canvas = tk.Canvas(wa_frame, width=28, height=28, bg=INPUT_BG, highlightthickness=0)
        wa_canvas.pack(side="left", padx=(0, 8))
        wa_canvas.create_oval(2, 2, 26, 26, fill="#25D366", outline="")
        wa_canvas.create_polygon(3, 20, 2, 26, 8, 23, fill="#25D366", outline="")
        wa_canvas.create_text(14, 14, text="✆", fill="white", font=(UI_FONT, 12, "bold"))

        wa_text_box = tk.Frame(wa_frame, bg=INPUT_BG)
        wa_text_box.pack(side="left", fill="x", expand=True)
        tk.Label(wa_text_box, text="WhatsApp", font=(UI_FONT, 8, "bold"), fg="#25D366", bg=INPUT_BG).pack(anchor="w")
        tk.Label(wa_text_box, text="03339962474", font=(UI_FONT, 9, "bold"), fg=TEXT_WHITE, bg=INPUT_BG).pack(anchor="w")

        wa_msg = f"Hello,%20I%20would%20like%20to%20purchase%20Bill%20Scrapper%20license.%20My%20Hardware%20ID%20is:%20{self.device_id}"
        wa_chat_btn = tk.Button(
            wa_frame,
            text="💬 Chat",
            bg=ACCENT_GREEN,
            fg=TEXT_WHITE,
            activebackground=ACCENT_GREEN_HOVER,
            activeforeground=TEXT_WHITE,
            font=(UI_FONT, 8, "bold"),
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=3,
            command=lambda: webbrowser.open(f"https://wa.me/923339962474?text={wa_msg}"),
        )
        wa_chat_btn.pack(side="right", padx=(4, 0))

        wa_copy_btn = tk.Button(
            wa_frame,
            text="📋 Copy",
            bg=BTN_SEC_BG,
            fg=TEXT_MAIN,
            activebackground=BTN_SEC_HOVER,
            activeforeground=TEXT_WHITE,
            font=(UI_FONT, 8),
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=3,
            command=lambda: self._copy_text("03339962474", wa_copy_btn),
        )
        wa_copy_btn.pack(side="right")

        # Gmail Row
        gm_frame = tk.Frame(contact_card, bg=INPUT_BG, padx=10, pady=8, highlightthickness=1, highlightbackground=BORDER_COLOR)
        gm_frame.pack(fill="x")

        gm_canvas = tk.Canvas(gm_frame, width=28, height=28, bg=INPUT_BG, highlightthickness=0)
        gm_canvas.pack(side="left", padx=(0, 8))
        gm_canvas.create_oval(2, 2, 26, 26, fill="#EA4335", outline="")
        gm_canvas.create_text(14, 14, text="✉", fill="white", font=(UI_FONT, 12, "bold"))

        gm_text_box = tk.Frame(gm_frame, bg=INPUT_BG)
        gm_text_box.pack(side="left", fill="x", expand=True)
        tk.Label(gm_text_box, text="Gmail / Email", font=(UI_FONT, 8, "bold"), fg="#EA4335", bg=INPUT_BG).pack(anchor="w")
        tk.Label(gm_text_box, text="tahir.naqash4@gmail.com", font=(UI_FONT, 9, "bold"), fg=TEXT_WHITE, bg=INPUT_BG).pack(anchor="w")

        gm_subject = "Bill%20Scrapper%20License%20Purchase"
        gm_body = f"Hello,%0A%0AI%20would%20like%20to%20purchase%20a%20Bill%20Scrapper%20license.%0A%0AMy%20Hardware%20ID%20is:%20{self.device_id}%0A"
        gm_mail_btn = tk.Button(
            gm_frame,
            text="✉ Email",
            bg="#b83a3a",
            fg=TEXT_WHITE,
            activebackground="#c93b3b",
            activeforeground=TEXT_WHITE,
            font=(UI_FONT, 8, "bold"),
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=3,
            command=lambda: webbrowser.open(f"mailto:tahir.naqash4@gmail.com?subject={gm_subject}&body={gm_body}"),
        )
        gm_mail_btn.pack(side="right", padx=(4, 0))

        gm_copy_btn = tk.Button(
            gm_frame,
            text="📋 Copy",
            bg=BTN_SEC_BG,
            fg=TEXT_MAIN,
            activebackground=BTN_SEC_HOVER,
            activeforeground=TEXT_WHITE,
            font=(UI_FONT, 8),
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=3,
            command=lambda: self._copy_text("tahir.naqash4@gmail.com", gm_copy_btn),
        )
        gm_copy_btn.pack(side="right")

        # Bottom Switcher
        back_btn = tk.Button(
            self.tab_buy,
            text="←  Already Have a Serial Key? Return to Activation",
            bg=BTN_SEC_BG,
            fg=TEXT_MAIN,
            activebackground=BTN_SEC_HOVER,
            activeforeground=TEXT_WHITE,
            font=(UI_FONT, 9),
            relief="flat",
            cursor="hand2",
            pady=6,
            command=self._show_activate_tab,
        )
        back_btn.pack(fill="x")

    def _show_activate_tab(self):
        self.tab_buy.pack_forget()
        self.tab_activate.pack(fill="both", expand=True)
        self.btn_tab_activate.configure(bg=ACCENT_GREEN, fg=TEXT_WHITE, activebackground=ACCENT_GREEN_HOVER)
        self.btn_tab_buy.configure(bg=BTN_SEC_BG, fg=TEXT_MAIN, activebackground=BTN_SEC_HOVER)
        self.key_entry.focus_set()

    def _show_buy_tab(self):
        self.tab_activate.pack_forget()
        self.tab_buy.pack(fill="both", expand=True)
        self.btn_tab_buy.configure(bg=ACCENT_GREEN, fg=TEXT_WHITE, activebackground=ACCENT_GREEN_HOVER)
        self.btn_tab_activate.configure(bg=BTN_SEC_BG, fg=TEXT_MAIN, activebackground=BTN_SEC_HOVER)

    def _copy_text(self, text: str, btn: tk.Button):
        self.clipboard_clear()
        self.clipboard_append(text)
        old_txt = btn.cget("text")
        btn.configure(text="✅ Copied!", fg="#4ec9b0")
        self.after(2000, lambda: btn.configure(text=old_txt, fg=TEXT_MAIN) if self.winfo_exists() and btn.winfo_exists() else None)

    def _copy_device_id(self):
        self.clipboard_clear()
        self.clipboard_append(self.device_id)
        self.copy_id_btn.configure(text="✅ Copied!", fg="#4ec9b0")
        self.after(2000, lambda: self.copy_id_btn.configure(text="📋 Copy ID", fg=TEXT_MAIN) if self.winfo_exists() else None)

    def _set_status(self, text: str, color: str = TEXT_MUTED):
        self.status_var.set(text)
        self.status_lbl.configure(fg=color)

    def _do_activate(self):
        key = self.key_var.get().strip().upper()
        if not key or key == "BS-XXXX-XXXX-XXXX":
            self._set_status("Please enter your serial key to activate.", color="#f14c4c")
            return

        self.activate_btn.configure(state="disabled", text="Activating...", bg="#1a5c28")
        self._set_status("Connecting to license server...", color="#e5c07b")
        self.update_idletasks()

        def worker():
            try:
                res = activate_license_remote(key, self.device_id, candidate_ids=get_candidate_device_ids())
                self.queue.put(("result", res))
            except Exception as e:
                self.queue.put(("error", str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_activation_result(self, res: dict):
        if not self.winfo_exists():
            return
        self.activate_btn.configure(state="normal", text="⚡  ACTIVATE LICENSE", bg=ACCENT_GREEN)
        if res.get("valid"):
            save_local_license(res)
            self.activated = True
            self.license_info = res
            self._set_status("✅ License activated successfully! Launching Bill Scrapper...", color="#89d185")
            self.after(600, self._finish_activation)
        else:
            msg = res.get("message") or "License activation failed."
            self._set_status(f"❌ {msg}", color="#f14c4c")

    def _finish_activation(self):
        if self.winfo_exists():
            self.destroy()

    def _on_activation_error(self, err_msg: str):
        if not self.winfo_exists():
            return
        self.activate_btn.configure(state="normal", text="⚡  ACTIVATE LICENSE", bg=ACCENT_GREEN)
        self._set_status(f"❌ Connection error: {err_msg}", color="#f14c4c")

    def _on_close(self):
        self.activated = False
        if self.winfo_exists():
            self.destroy()


class AboutDialog(tk.Toplevel):
    """Sleek modal dialog showing application, active license, and developer information."""
    def __init__(self, parent: "BillScraperApp"):
        super().__init__(parent)
        self.withdraw()  # Keep hidden during configuration to prevent visual blink
        self.parent = parent
        self.title("About Bill Scrapper")
        self.configure(bg=BG_DARK)
        self.resizable(False, False)
        self.transient(parent)

        try:
            if sys.platform == "win32":
                icon_file = resource_path("assets/billscrapper.ico")
                if not os.path.exists(icon_file):
                    icon_file = resource_path("assets/mdiscrapper.ico")
                if os.path.exists(icon_file):
                    self.iconbitmap(icon_file)
            else:
                png_file = resource_path("assets/billscrapper-icon.png")
                if os.path.exists(png_file):
                    img = tk.PhotoImage(file=png_file)
                    self.iconphoto(False, img)
        except Exception:
            pass

        w, h = 450, 420
        px = parent.winfo_rootx() + (parent.winfo_width() // 2) - (w // 2)
        py = parent.winfo_rooty() + (parent.winfo_height() // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{max(0, px)}+{max(0, py)}")

        container = ttk.Frame(self, style="Root.TFrame", padding=(20, 16, 20, 16))
        container.pack(fill="both", expand=True)

        app_title = ttk.Label(container, text="Bill Scrapper", style="Title.TLabel", font=(UI_FONT, 14, "bold"))
        app_title.pack(anchor="w")

        app_ver = ttk.Label(container, text=f"Version {APP_VERSION}", style="Sub.TLabel")
        app_ver.pack(anchor="w", pady=(2, 8))

        # 1. License Information Card
        lic = getattr(parent, "active_license", None) or load_local_license()
        serial_key = lic.get("serial_key", "Not Activated")
        plan_str = format_plan_type(lic.get("plan_type"))
        expiry_str = format_expiry_date(lic.get("expires_at"))

        lic_card = ttk.Frame(container, style="Card.TFrame", padding=(14, 10))
        lic_card.pack(fill="x", pady=(0, 10))

        ttk.Label(lic_card, text="LICENSE INFORMATION", font=(UI_FONT, 8, "bold"), foreground=TEXT_MUTED, background=CARD_BG).pack(anchor="w", pady=(0, 4))

        key_row = ttk.Frame(lic_card, style="Card.TFrame")
        key_row.pack(fill="x", pady=(2, 4))

        ttk.Label(key_row, text="Key: ", font=(UI_FONT, 9, "bold"), foreground=TEXT_WHITE, background=CARD_BG).pack(side="left")
        ttk.Label(key_row, text=serial_key, font=(MONO_FONT, 9, "bold"), foreground="#4ec9b0", background=CARD_BG).pack(side="left")

        self.copy_key_btn = tk.Button(
            key_row,
            text="📋 Copy",
            bg=BTN_SEC_BG,
            fg=TEXT_MAIN,
            activebackground=BTN_SEC_HOVER,
            activeforeground=TEXT_WHITE,
            font=(UI_FONT, 8),
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=2,
            command=self._copy_serial_key,
        )
        self.copy_key_btn.pack(side="right")

        self.change_key_btn = tk.Button(
            key_row,
            text="🔑 Change Key",
            bg=BTN_SEC_BG,
            fg=TEXT_WHITE,
            activebackground=BTN_SEC_HOVER,
            activeforeground=TEXT_WHITE,
            font=(UI_FONT, 8, "bold"),
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=2,
            command=self._change_serial_key,
        )
        self.change_key_btn.pack(side="right", padx=(0, 4))

        hw_lbl = ttk.Label(
            lic_card,
            text=f"Hardware ID: {get_device_id()}",
            font=(MONO_FONT, 8),
            foreground=TEXT_MUTED,
            background=CARD_BG,
        )
        hw_lbl.pack(anchor="w", pady=(0, 2))

        plan_lbl = ttk.Label(
            lic_card,
            text=f"Plan: {plan_str}",
            font=(UI_FONT, 9),
            foreground=TEXT_MAIN,
            background=CARD_BG,
        )
        plan_lbl.pack(anchor="w", pady=(0, 2))

        exp_lbl = ttk.Label(
            lic_card,
            text=f"Status: {expiry_str}",
            font=(UI_FONT, 9),
            foreground="#89d185" if "remaining" in expiry_str or "today" in expiry_str else TEXT_MAIN,
            background=CARD_BG,
        )
        exp_lbl.pack(anchor="w")

        # 2. Developer Information Card
        info_card = ttk.Frame(container, style="Card.TFrame", padding=(14, 10))
        info_card.pack(fill="x", pady=(0, 12))

        dev_lbl = ttk.Label(
            info_card,
            text="Developed by TN Developers Co.",
            font=(UI_FONT, 9, "bold"),
            foreground=TEXT_WHITE,
            background=CARD_BG,
        )
        dev_lbl.pack(anchor="w", pady=(0, 4))

        email_lbl = ttk.Label(
            info_card,
            text="Email: tahir.naqash4@gmail.com",
            font=(UI_FONT, 9),
            foreground=TEXT_MAIN,
            background=CARD_BG,
        )
        email_lbl.pack(anchor="w", pady=(0, 2))

        contact_lbl = ttk.Label(
            info_card,
            text="Contact: 03339962474",
            font=(UI_FONT, 9),
            foreground=TEXT_MAIN,
            background=CARD_BG,
        )
        contact_lbl.pack(anchor="w")

        btn_box = ttk.Frame(container, style="Root.TFrame")
        btn_box.pack(fill="x")

        close_btn = ttk.Button(btn_box, text="Close", command=self.destroy)
        close_btn.pack(side="right")
        close_btn.focus_set()

        set_dark_title_bar(self)
        self.update_idletasks()
        self.deiconify()
        self.grab_set()

    def _copy_serial_key(self):
        lic = getattr(self.master, "active_license", None) or load_local_license()
        key = lic.get("serial_key", "")
        if key:
            self.clipboard_clear()
            self.clipboard_append(key)
            self.copy_key_btn.configure(text="✅ Copied!", fg="#4ec9b0")
            self.after(2000, lambda: self.copy_key_btn.configure(text="📋 Copy", fg=TEXT_MAIN) if self.winfo_exists() else None)

    def _change_serial_key(self):
        parent = self.parent
        self.destroy()
        gate = LicenseGateDialog(parent)
        parent.wait_window(gate)
        if gate.activated and gate.license_info:
            parent.active_license = gate.license_info
            messagebox.showinfo("License Activated", "Serial key updated successfully!", parent=parent)


class BillScraperApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw()  # Keep hidden immediately so no blank or unrendered window flashes
        self.title("Bill Scrapper")
        self.configure(bg=BG_DARK)
        self.resizable(True, True)
        self.geometry("640x600")
        self.minsize(580, 520)

        self._icon_img = None
        try:
            if sys.platform == "win32":
                icon_file = resource_path("assets/billscrapper.ico")
                if not os.path.exists(icon_file):
                    icon_file = resource_path("assets/mdiscrapper.ico")
                if os.path.exists(icon_file):
                    self.iconbitmap(icon_file)
            png_file = resource_path("assets/billscrapper-icon.png")
            if os.path.exists(png_file):
                self._icon_img = tk.PhotoImage(file=png_file)
                self.iconphoto(True, self._icon_img)
        except Exception:
            pass

        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self.style.configure(".", background=CARD_BG, foreground=TEXT_MAIN)
        self.style.configure("TFrame", background=BG_DARK)
        self.style.configure("Root.TFrame", background=BG_DARK)
        self.style.configure("Card.TFrame", background=CARD_BG)

        self.style.configure(
            "TLabelframe",
            background=CARD_BG,
            bordercolor=BORDER_COLOR,
            lightcolor=BORDER_COLOR,
            darkcolor=BORDER_COLOR,
            relief="solid",
        )
        self.style.configure(
            "TLabelframe.Label",
            background=CARD_BG,
            foreground=TEXT_WHITE,
            font=(UI_FONT, 9, "bold"),
        )

        self.style.configure("TLabel", background=CARD_BG, foreground=TEXT_MAIN, font=(UI_FONT, 9))
        self.style.configure("Root.TLabel", background=BG_DARK, foreground=TEXT_MAIN, font=(UI_FONT, 9))
        self.style.configure("Title.TLabel", background=BG_DARK, foreground=TEXT_WHITE, font=(UI_FONT, 16, "bold"))
        self.style.configure("Sub.TLabel", background=BG_DARK, foreground=TEXT_MUTED, font=(UI_FONT, 9))
        self.style.configure("Status.TLabel", background=BG_DARK, foreground=TEXT_MAIN, font=(UI_FONT, 9))
        self.style.configure("Muted.TLabel", background=CARD_BG, foreground=TEXT_MUTED, font=(UI_FONT, 8, "italic"))
        self.style.configure("StatTitle.TLabel", background=CARD_BG, foreground=TEXT_MUTED, font=(UI_FONT, 9))
        self.style.configure("StatVal.TLabel", background=CARD_BG, foreground=TEXT_WHITE, font=(UI_FONT, 9, "bold"))
        self.style.configure("StatOk.TLabel", background=CARD_BG, foreground="#89d185", font=(UI_FONT, 9, "bold"))
        self.style.configure("StatErr.TLabel", background=CARD_BG, foreground="#f14c4c", font=(UI_FONT, 9, "bold"))

        self.style.configure(
            "TEntry",
            fieldbackground=INPUT_BG,
            foreground=TEXT_WHITE,
            bordercolor=BORDER_COLOR,
            lightcolor=BORDER_COLOR,
            darkcolor=BORDER_COLOR,
            insertcolor=TEXT_WHITE,
        )
        self.style.configure(
            "TSpinbox",
            fieldbackground=INPUT_BG,
            foreground=TEXT_WHITE,
            bordercolor=BORDER_COLOR,
            lightcolor=BORDER_COLOR,
            darkcolor=BORDER_COLOR,
            arrowcolor=TEXT_MUTED,
            insertcolor=TEXT_WHITE,
        )

        self.style.configure(
            "TButton",
            background=BTN_SEC_BG,
            foreground=TEXT_MAIN,
            bordercolor=BORDER_COLOR,
            lightcolor=BORDER_COLOR,
            darkcolor=BORDER_COLOR,
            focuscolor="none",
            padding=(10, 4),
        )
        self.style.map(
            "TButton",
            background=[("active", BTN_SEC_HOVER), ("pressed", "#2d2d2d"), ("disabled", BTN_DISABLED_BG)],
            foreground=[("disabled", BTN_DISABLED_FG)],
        )

        self.style.configure(
            "Green.Horizontal.TProgressbar",
            troughcolor=INPUT_BG,
            background=ACCENT_GREEN,
            lightcolor=ACCENT_GREEN_HOVER,
            darkcolor="#1d6e2e",
            bordercolor=BORDER_COLOR,
        )
        self.style.configure(
            "Vertical.TScrollbar",
            background=BTN_SEC_BG,
            troughcolor=BG_DARK,
            bordercolor=BORDER_COLOR,
            arrowcolor=TEXT_MUTED,
        )
        self.style.map("Vertical.TScrollbar", background=[("active", BTN_SEC_HOVER)])

        self.events = queue.Queue()
        self.stop_event = threading.Event()
        self.input_var = tk.StringVar()
        self.delay_var = tk.DoubleVar(value=2.0)
        self.status_var = tk.StringVar(value="Ready. Choose an Excel (.xlsx) or LibreOffice (.ods) file containing AC numbers in Column A.")
        self.total_var = tk.StringVar(value="0")
        self.ok_var = tk.StringVar(value="0")
        self.err_var = tk.StringVar(value="0")
        self.success_count = 0
        self.error_count = 0
        self.latest_update: UpdateInfo | None = None
        self.update_banner: tk.Frame | None = None
        self.files_frame: ttk.LabelFrame | None = None

        self.active_license: dict | None = None

        if not self.ensure_licensed():
            try:
                if self.winfo_exists():
                    self.destroy()
            except Exception:
                pass
            return

        self.build_ui()
        set_dark_title_bar(self)
        self.update_idletasks()
        self.deiconify()
        self.after(100, self.drain_events)
        threading.Thread(target=self._bg_update_check, daemon=True).start()

    def ensure_licensed(self) -> bool:
        """Verify existing license or prompt user with the LicenseGateDialog."""
        local_lic = load_local_license()
        device_id = get_device_id()
        initial_error = ""

        if local_lic and local_lic.get("serial_key"):
            # Check local expiration timestamp for instant, latency-free startup (<50ms)
            exp_str = local_lic.get("expires_at")
            is_locally_valid = False
            if exp_str:
                try:
                    clean_ts = str(exp_str).replace("Z", "+00:00")
                    exp_dt = datetime.fromisoformat(clean_ts)
                    if exp_dt > datetime.now(timezone.utc):
                        is_locally_valid = True
                except Exception:
                    pass

            if is_locally_valid:
                # Fast path: instant UI display, verify remotely in background thread
                self.active_license = local_lic
                threading.Thread(target=self._bg_verify_license, daemon=True).start()
                return True

            # If local cache is expired or unverified, try remote verification directly
            try:
                candidate_ids = get_candidate_device_ids()
                res = verify_license_remote(local_lic["serial_key"], device_id, candidate_ids=candidate_ids)
                if res.get("valid"):
                    local_lic.update(res)
                    save_local_license(local_lic)
                    self.active_license = local_lic
                    return True
                else:
                    initial_error = res.get("message") or "License expired or invalid."
            except Exception as e:
                initial_error = f"License verification error: {e}"

        # If we reach here, license is missing, expired, or invalid
        gate = LicenseGateDialog(self, initial_error=initial_error)
        self.wait_window(gate)

        if gate.activated and gate.license_info:
            self.active_license = gate.license_info
            return True

        return False

    def _bg_verify_license(self):
        """Asynchronously validates the active license in the background without blocking the UI."""
        try:
            local_lic = load_local_license()
            if not local_lic or not local_lic.get("serial_key"):
                return
            key = local_lic["serial_key"]
            device_id = get_device_id()
            candidate_ids = get_candidate_device_ids()
            res = verify_license_remote(key, device_id, candidate_ids=candidate_ids)
            if res.get("valid"):
                local_lic.update(res)
                save_local_license(local_lic)
                self.active_license = local_lic
            else:
                msg = res.get("message") or "Your license is no longer valid or has expired."
                self.after(0, lambda: self._on_license_revoked(msg))
        except Exception:
            # Network issue / offline mode: tolerate silently while local expires_at is unexpired
            pass

    def _on_license_revoked(self, reason: str):
        messagebox.showerror("License Expired", reason, parent=self)
        gate = LicenseGateDialog(self, initial_error=reason)
        self.wait_window(gate)
        if gate.activated and gate.license_info:
            self.active_license = gate.license_info
        else:
            self.destroy()

    def build_ui(self):
        root = ttk.Frame(self, style="Root.TFrame", padding=(18, 14, 18, 14))
        root.pack(fill="both", expand=True)

        header_frame = ttk.Frame(root, style="Root.TFrame")
        header_frame.pack(fill="x", pady=(0, 10))

        header_top = ttk.Frame(header_frame, style="Root.TFrame")
        header_top.pack(fill="x")

        title_lbl = ttk.Label(header_top, text="Bill Scrapper", style="Title.TLabel")
        title_lbl.pack(side="left")

        ver_box = ttk.Frame(header_top, style="Root.TFrame")
        ver_box.pack(side="right", anchor="e")

        ttk.Label(ver_box, text=f"v{APP_VERSION}", style="Sub.TLabel").pack(side="left", padx=(0, 8))
        self.check_update_btn = tk.Button(
            ver_box, text="Check for Updates", command=self.manual_update_check,
            bg=BTN_SEC_BG, fg=TEXT_MAIN, activebackground=BTN_SEC_HOVER, activeforeground="white",
            font=(UI_FONT, 8), relief="flat", padx=8, pady=2, cursor="hand2", bd=0
        )
        self.check_update_btn.pack(side="left")

        sub_lbl = ttk.Label(header_frame, text="Automated Bill Scraping", style="Sub.TLabel")
        sub_lbl.pack(anchor="w", pady=(2, 0))

        # Notification banner for available updates (hidden by default)
        self.update_banner = tk.Frame(root, bg="#1a314c", bd=1, relief="solid", highlightthickness=1, highlightbackground="#1d70b8")

        banner_content = tk.Frame(self.update_banner, bg="#1a314c")
        banner_content.pack(fill="x", padx=10, pady=6)

        self.update_banner_lbl = tk.Label(
            banner_content,
            text="",
            bg="#1a314c", fg=TEXT_WHITE,
            font=(UI_FONT, 9, "bold")
        )
        self.update_banner_lbl.pack(side="left", anchor="w")

        btn_box = tk.Frame(banner_content, bg="#1a314c")
        btn_box.pack(side="right")

        self.update_now_btn = tk.Button(
            btn_box, text="⬇  Update Now", command=self.install_update,
            bg=ACCENT_GREEN, fg="white", activebackground=ACCENT_GREEN_HOVER, activeforeground="white",
            font=(UI_FONT, 9, "bold"), relief="flat", padx=10, pady=3, cursor="hand2", bd=0
        )
        self.update_now_btn.pack(side="left", padx=(0, 6))

        dismiss_btn = tk.Button(
            btn_box, text="✕", command=self.dismiss_update_banner,
            bg="#1a314c", fg=TEXT_MUTED, activebackground="#1a314c", activeforeground=TEXT_WHITE,
            font=(UI_FONT, 9), relief="flat", padx=6, pady=2, cursor="hand2", bd=0
        )
        dismiss_btn.pack(side="left")

        self.files_frame = ttk.LabelFrame(root, text=" Target Workbook ", padding=10)
        self.files_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(self.files_frame, text="Workbook (.xlsx / .ods):", font=(UI_FONT, 9, "bold")).grid(row=0, column=0, sticky="w", pady=4)
        entry = ttk.Entry(self.files_frame, textvariable=self.input_var, font=(UI_FONT, 9))
        entry.grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=4)

        browse_btn = ttk.Button(self.files_frame, text="Browse...", command=self.choose_input)
        browse_btn.grid(row=0, column=2, pady=4)
        self.files_frame.columnconfigure(1, weight=1)

        mid_frame = ttk.Frame(root, style="Root.TFrame")
        mid_frame.pack(fill="x", pady=(0, 10))

        settings_box = ttk.LabelFrame(mid_frame, text=" Settings ", padding=8)
        settings_box.pack(side="left", fill="both", expand=True, padx=(0, 6))

        ttk.Label(settings_box, text="Delay (seconds):").grid(row=0, column=0, sticky="w", pady=2)
        spin = ttk.Spinbox(settings_box, from_=0.5, to=30.0, increment=0.5, textvariable=self.delay_var, width=6)
        spin.grid(row=0, column=1, sticky="w", padx=(6, 4), pady=2)
        ttk.Label(settings_box, text="(Default 2.0s prevents PITC blocking)", style="Muted.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))

        stats_box = ttk.LabelFrame(mid_frame, text=" Progress Summary ", padding=8)
        stats_box.pack(side="right", fill="both", expand=True, padx=(6, 0))

        ttk.Label(stats_box, text="Total:", style="StatTitle.TLabel").grid(row=0, column=0, sticky="w", padx=4)
        ttk.Label(stats_box, textvariable=self.total_var, style="StatVal.TLabel").grid(row=0, column=1, sticky="w", padx=4)

        ttk.Label(stats_box, text="Processed:", style="StatTitle.TLabel").grid(row=0, column=2, sticky="w", padx=4)
        ttk.Label(stats_box, textvariable=self.ok_var, style="StatOk.TLabel").grid(row=0, column=3, sticky="w", padx=4)

        ttk.Label(stats_box, text="Errors:", style="StatTitle.TLabel").grid(row=0, column=4, sticky="w", padx=4)
        ttk.Label(stats_box, textvariable=self.err_var, style="StatErr.TLabel").grid(row=0, column=5, sticky="w", padx=4)

        self.progress = ttk.Progressbar(root, mode="determinate", style="Green.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(4, 6))

        status_lbl = ttk.Label(root, textvariable=self.status_var, style="Status.TLabel")
        status_lbl.pack(anchor="w", pady=(0, 6))

        log_frame = ttk.LabelFrame(root, text=" Live Execution Log ", padding=6)
        log_frame.pack(fill="both", expand=True, pady=(0, 10))

        self.log = tk.Text(
            log_frame, height=10, state="disabled", wrap="word",
            font=(MONO_FONT, 9), bg=LOG_BG, fg=LOG_FG,
            insertbackground=TEXT_WHITE, selectbackground="#264f78", selectforeground="#ffffff",
            relief="flat", bd=1, highlightthickness=1, highlightbackground=BORDER_COLOR, highlightcolor=BORDER_COLOR
        )
        self.log.tag_config("ok", foreground="#89d185")
        self.log.tag_config("err", foreground="#f14c4c")
        self.log.tag_config("info", foreground=TEXT_MAIN)
        self.log.tag_config("bold", foreground=TEXT_WHITE, font=(MONO_FONT, 9, "bold"))

        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        btn_bar = ttk.Frame(root, style="Root.TFrame")
        btn_bar.pack(fill="x")

        self.start_button = tk.Button(
            btn_bar, text="▶  Start Scraping", command=self.start,
            bg=ACCENT_GREEN, fg="white", activebackground=ACCENT_GREEN_HOVER, activeforeground="white",
            font=(UI_FONT, 10, "bold"), relief="flat", padx=16, pady=5, cursor="hand2", bd=0
        )
        self.start_button.pack(side="left")

        self.cancel_button = tk.Button(
            btn_bar, text="⏹  Stop", command=self.cancel,
            bg=BTN_DISABLED_BG, fg="white", activebackground=ACCENT_RED_HOVER, activeforeground="white",
            font=(UI_FONT, 10, "bold"), relief="flat", padx=14, pady=5, state="disabled", cursor="hand2", bd=0,
            disabledforeground=BTN_DISABLED_FG
        )
        self.cancel_button.pack(side="left", padx=(10, 0))

        self.about_btn = ttk.Button(btn_bar, text="About", command=self.show_about)
        self.about_btn.pack(side="right")

        self.open_btn = ttk.Button(btn_bar, text="Open File Folder", command=self.open_output_folder)
        self.open_btn.pack(side="right", padx=(0, 8))

    def choose_input(self):
        path = filedialog.askopenfilename(
            title="Choose Account Workbook",
            filetypes=[
                ("Spreadsheet Files", "*.xlsx *.ods"),
                ("Excel Workbook", "*.xlsx"),
                ("LibreOffice Calc (.ods)", "*.ods"),
            ],
        )
        if path:
            self.input_var.set(path)
            self.status_var.set(f"Selected: {Path(path).name}")

    def show_about(self):
        AboutDialog(self)

    def open_output_folder(self):
        val = self.input_var.get().strip()
        if val and Path(val).parent.exists():
            folder = str(Path(val).parent.resolve())
            try:
                if sys.platform == "win32":
                    os.startfile(folder)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", folder])
                else:
                    subprocess.Popen(["xdg-open", folder])
            except Exception as exc:
                messagebox.showerror("Error", f"Failed to open folder: {exc}")
        else:
            messagebox.showinfo("Select File", "Please select a workbook file first.")

    def write_log(self, text, tag=None):
        self.log.configure(state="normal")
        if tag:
            self.log.insert("end", text + "\n", tag)
        else:
            self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def start(self):
        path = Path(self.input_var.get().strip())
        if not path.is_file() or path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            messagebox.showerror("Input Required", "Choose an existing .xlsx or .ods workbook first.")
            return

        if not messagebox.askyesno(
            "Update Workbook",
            "The workbook will be updated in place with bill data, separate unit columns, cumulative MDI values, differences, Actual MDI, and fixed charges calculations.\n\nContinue?"
        ):
            return

        try:
            delay = max(0.1, float(self.delay_var.get()))
        except (TypeError, ValueError):
            messagebox.showerror("Invalid Delay", "Delay must be a positive number.")
            return

        self.stop_event.clear()
        self.progress.configure(value=0)
        self.success_count = 0
        self.error_count = 0
        self.ok_var.set("0")
        self.err_var.set("0")
        self.total_var.set("0")

        self.start_button.configure(state="disabled", bg=BTN_DISABLED_BG)
        self.cancel_button.configure(state="normal", bg=ACCENT_RED)
        self.write_log(f"--- Starting Scraping on: {path.name} ---", "info")
        self.write_log(f"Request delay: {delay}s | Calculating Actual MDI (max diff x MF) & Actual Fix Charges...", "info")

        threading.Thread(target=self.run, args=(path, delay), daemon=True).start()

    def run(self, path, delay):
        try:
            def progress(done, total, row):
                raw_status = row.get("Remarks") or row.get("scrape_status", "")
                if str(raw_status).strip().lower() in ("ok", "done", "processed"):
                    status_text = "Processed"
                else:
                    status_text = "Bill Not Found"
                self.events.put(("progress", done, total, row.get("_account", ""), status_text))

            def log_callback(msg):
                self.events.put(("log", msg))

            PITCBillScraper(delay=delay).update_workbook(path, progress, self.stop_event, log_callback=log_callback)
            self.events.put(("done", str(path), self.stop_event.is_set()))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def cancel(self):
        self.stop_event.set()
        self.status_var.set("Stopping... Saving all processed records...")
        self.cancel_button.configure(state="disabled", bg=BTN_DISABLED_BG)
        self.write_log("⏹ Stop requested. Saving all progress processed so far...", "err")

    def show_update_banner(self, info: UpdateInfo):
        self.latest_update = info
        self.update_banner_lbl.configure(
            text=f"🚀 Update Available: Version {info.version} is ready to install!"
        )
        if self.files_frame and self.files_frame.winfo_exists():
            self.update_banner.pack(fill="x", pady=(0, 10), before=self.files_frame)
        else:
            self.update_banner.pack(fill="x", pady=(0, 10))

    def dismiss_update_banner(self):
        if self.update_banner:
            self.update_banner.pack_forget()

    def install_update(self):
        if not self.latest_update:
            return
        if self.cancel_button["state"] == "normal":
            messagebox.showwarning(
                "Task Running",
                "Please stop or wait for the current scraping process to complete before updating.",
                parent=self
            )
            return
        UpdateDialog(self, self.latest_update)

    def _bg_update_check(self, manual: bool = False):
        try:
            info = check_for_updates(APP_VERSION)
            if info:
                self.events.put(("update_available", info))
            elif manual:
                self.events.put(("update_not_found",))
        except Exception as exc:
            if manual:
                self.events.put(("update_check_error", str(exc)))

    def manual_update_check(self):
        self.status_var.set("Checking for updates...")
        self.write_log(f"Checking for newer releases (current: v{APP_VERSION})...", "info")
        threading.Thread(target=self._bg_update_check, args=(True,), daemon=True).start()

    def drain_events(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "log":
                    _, msg = event
                    self.write_log(msg, "info")

                elif event[0] == "progress":
                    _, done, total, account, status = event
                    self.progress.configure(maximum=max(total, 1), value=done)
                    self.total_var.set(str(total))
                    display_status = "Processed" if str(status).strip().lower() in ("ok", "done", "processed") else "Bill Not Found"
                    self.status_var.set(f"Processed {done} of {total} accounts...")

                    if display_status == "Processed":
                        self.success_count += 1
                        self.ok_var.set(str(self.success_count))
                        self.write_log(f"[{done}/{total}] {account}: Processed", "ok")
                    else:
                        self.error_count += 1
                        self.err_var.set(str(self.error_count))
                        self.write_log(f"[{done}/{total}] {account}: Bill Not Found", "err")

                elif event[0] == "done":
                    _, path, cancelled = event
                    self.start_button.configure(state="normal", bg=ACCENT_GREEN)
                    self.cancel_button.configure(state="disabled", bg=BTN_DISABLED_BG)
                    if cancelled:
                        self.status_var.set("Stopped (All progress saved)")
                        self.write_log(
                            f"✔ Stopped. All progress up to the last bill has been saved to: {path}",
                            "err"
                        )
                        messagebox.showinfo(
                            "Processing Stopped",
                            f"Processing was stopped. All progress up to the last bill has been safely saved to:\n\n{path}"
                        )
                    else:
                        self.status_var.set("Completed successfully!")
                        self.write_log(
                            f"=== All Done! Successfully updated: {path} ===",
                            "info"
                        )
                        if not cancelled:
                            messagebox.showinfo(
                                "Processing Finished",
                                f"Workbook updated successfully with separated unit columns, Actual MDI, and Fixed Charges calculations:\n\n{path}"
                            )

                elif event[0] == "error":
                    self.start_button.configure(state="normal", bg=ACCENT_GREEN)
                    self.cancel_button.configure(state="disabled", bg=BTN_DISABLED_BG)
                    self.status_var.set("Failed with error.")
                    self.write_log("ERROR: " + event[1], "err")
                    messagebox.showerror("Export Failed", event[1])

                elif event[0] == "update_available":
                    _, info = event
                    self.show_update_banner(info)
                    self.write_log(f"🔔 Update available: {info.version} (current: v{APP_VERSION})", "info")
                    self.status_var.set(f"Update available: {info.version}. Click 'Update Now' above.")

                elif event[0] == "update_not_found":
                    self.status_var.set(f"Bill Scrapper is up to date (v{APP_VERSION}).")
                    messagebox.showinfo("Update Check", f"You are on the latest version of Bill Scrapper (v{APP_VERSION}).")

                elif event[0] == "update_check_error":
                    self.status_var.set("Update check failed.")
                    messagebox.showwarning("Update Check", f"Could not check for updates:\n{event[1]}")
        except queue.Empty:
            pass
        self.after(100, self.drain_events)


def start_app():
    app = BillScraperApp()
    try:
        if getattr(app, "active_license", None) and app.winfo_exists():
            app.mainloop()
    except (tk.TclError, SystemExit):
        pass


if __name__ == "__main__":
    start_app()
