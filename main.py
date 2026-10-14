"""Desktop Bill Scrapper.

The selected .xlsx or .ods workbook is updated in place using its existing headings.
Excluded fields are removed before requests are made, and returned values are
written strictly below their matching headings without column shifting.
"""

from __future__ import annotations

import io
import os
import queue
import re
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import zipfile
from decimal import Decimal
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import requests
from bs4 import BeautifulSoup
from lxml import etree
from openpyxl import load_workbook

URL = "https://bill.pitc.com.pk/pescobill"


def resource_path(relative_path):
    """Resolve bundled resources in both source runs and PyInstaller builds."""
    bundle_root = getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)
    return str(Path(bundle_root) / relative_path)

# Fields that are intentionally never requested/exported.
DROP_LETTERS = {"B", "C", "E", "H", "K", "L", "M", "N", "P", "Q", "S"}
DROP_RANGES = ((24, 50), (28, 55), (36, 64), (67, 67), (98, 214), (217, 244))  # X:AX, AB:BC, AJ:BL, BO, CT:HF and HI:IJ
DROP_EXACT = {
    "reference_no", "consumer_id", "sub_division", "tariff_cat", "transformer",
    "energy_charges", "taxes", "current_bill", "arrears", "installment",
    "tracking_id",
    "total_electricity_charges", "net_electricity_charges", "total_fpa", "subsidies", "category", "26",
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
    "current bill": "current_bill",
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
        "AC No.", "tariff", "san_load", "status", "adjustments", "grand_total",
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

    def search(self, account):
        session = requests.Session()
        session.headers.update({"User-Agent": "Mozilla/5.0 (compatible; PITC-Bill-Exporter/1.0)"})
        page = session.get(URL, timeout=self.timeout)
        page.raise_for_status()
        soup = BeautifulSoup(page.text, "lxml")
        form = soup.select_one("form#SubmitForm") or soup.find("form")
        if not form:
            raise RuntimeError("PITC search form was not found")
        payload = {i["name"]: i.get("value", "") for i in form.select("input[type=hidden][name]")}
        payload.update({"rbSearchByList": "refno", "searchTextBox": account, "ruCodeTextBox": "", "btnSearch": "Search"})
        result = session.post(URL, data=payload, timeout=self.timeout)
        result.raise_for_status()
        if "CONSUMER DETAIL" not in result.text.upper() and "CONSUMER DETAILS" not in result.text.upper():
            raise RuntimeError(clean(BeautifulSoup(result.text, "lxml").get_text(" "))[:300] or "No bill returned")
        return BeautifulSoup(result.text, "lxml")

    def scrape_account(self, raw_account, existing_status=None):
        raw = str(raw_account).strip()
        if re.fullmatch(r"[0-9.]+e[+\-]?[0-9]+", raw, re.I):
            raw = format(Decimal(raw), "f")
        account = re.sub(r"\D", "", raw)
        if not re.fullmatch(r"\d{10,14}", account):
            return {"_account": account, "Remarks": "Bill not Found", "scrape_status": "invalid account number"}
        try:
            soup = self.search(account)
            out = {"_account": account, "Remarks": "Done", "scrape_status": "ok"}
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
        except Exception as exc:
            return {"_account": account, "Remarks": "Bill not Found", "scrape_status": f"error: {exc}"}

    def update_workbook(self, path, progress=None, stop_event=None):
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
            if not is_dropped(col_name):
                all_keys.add(col_name)

        # 2. Process each account and merge results into record dictionary
        for done, record in enumerate(row_records, 1):
            if stop_event and stop_event.is_set():
                break
            account = record.get("AC No.")
            existing_status = record.get("status")
            result = self.scrape_account(account, existing_status=existing_status)
            record.update(result)

            # Double-check status condition on merged record
            effective_status = record.get("status") or existing_status
            if is_disc_status(effective_status):
                record["Diff_cum_mdi1"] = None
                record["diff_cum_mdi2"] = None
                record["Actual MDI"] = None
                record["Actual_Fix_Charges"] = None

            for k in result:
                if not k.startswith("_") and not is_dropped(k):
                    all_keys.add(k)
            if progress:
                progress(done, total, result)
            time.sleep(self.delay)

        # 3. Build ordered final headers
        all_keys.add("AC No.")
        all_keys.add("feeder")
        all_keys.add("adjustments")
        all_keys.add("Remarks")

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
                    val_str = str(raw_val).strip().lower() if raw_val is not None else ""
                    val = "Done" if val_str in ("done", "ok", "processed") else "Bill not Found"
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

# Version & GitHub repository details for automated updates
APP_VERSION = "1.0.0"
GITHUB_REPO = "naqasht128/BillScrapper"
GITHUB_API_LATEST = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"


def parse_version_tuple(v: str) -> tuple[int, ...]:
    """Parse a version string (e.g. 'v1.0.2') into an integer tuple for robust comparisons."""
    nums = [int(n) for n in re.findall(r"\d+", str(v))]
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums)


class UpdateInfo:
    """Metadata regarding an available release from GitHub."""
    def __init__(self, version: str, download_url: str, file_name: str, file_size: int, release_notes: str = ""):
        self.version = version
        self.download_url = download_url
        self.file_name = file_name
        self.file_size = file_size
        self.release_notes = release_notes


def check_for_updates(current_version: str = APP_VERSION) -> UpdateInfo | None:
    """Query GitHub Releases API to check if a newer version is available."""
    try:
        headers = {
            "User-Agent": "BillScrapper-DesktopApp",
            "Accept": "application/vnd.github.v3+json",
        }
        resp = requests.get(GITHUB_API_LATEST, headers=headers, timeout=6)
        if resp.status_code != 200:
            return None
        data = resp.json()
        tag = data.get("tag_name", "").strip()
        if not tag:
            return None

        # Compare versions
        if parse_version_tuple(tag) <= parse_version_tuple(current_version):
            return None

        # Determine target executable asset for current OS
        assets = data.get("assets", [])
        is_win = sys.platform == "win32"
        target_asset = None
        for asset in assets:
            name = asset.get("name", "")
            if is_win and name.lower().endswith(".exe"):
                target_asset = asset
                break
            elif not is_win and (name == "BillScrapper" or not name.lower().endswith(".exe")):
                target_asset = asset
                break

        if not target_asset or not target_asset.get("browser_download_url"):
            return None

        return UpdateInfo(
            version=tag,
            download_url=target_asset["browser_download_url"],
            file_name=target_asset.get("name", "BillScrapper"),
            file_size=target_asset.get("size", 0),
            release_notes=data.get("body", ""),
        )
    except Exception:
        return None


def install_and_restart(new_file_path: str, root_window: tk.Tk):
    """Replace running executable with new_file_path and relaunch."""
    is_frozen = getattr(sys, "frozen", False)

    if not is_frozen:
        # Running from Python source (.py)
        msg = (
            f"Update downloaded successfully to:\n{new_file_path}\n\n"
            "Because you are running from Python source code, you can update via 'git pull' "
            "or run the downloaded standalone executable directly."
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
        )
        root_window.destroy()
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


class UpdateDialog(tk.Toplevel):
    """Modal dialog displaying download progress and applying the update."""
    def __init__(self, parent: "BillScraperApp", info: UpdateInfo):
        super().__init__(parent)
        self.parent = parent
        self.info = info
        self.cancelled = False
        self.temp_file: str | None = None

        self.title("Updating Bill Scrapper")
        self.configure(bg=BG_DARK)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

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

        self.protocol("WM_DELETE_WINDOW", self.cancel)
        threading.Thread(target=self._download_worker, daemon=True).start()

    def cancel(self):
        self.cancelled = True
        self.lbl_status.configure(text="Cancelling download...")
        self.cancel_btn.configure(state="disabled")

    def _download_worker(self):
        try:
            ext = ".exe" if sys.platform == "win32" else ""
            tpath = Path(tempfile.gettempdir()) / f"BillScrapper_update_{self.info.version}_{int(time.time())}{ext}"
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
                                self.after(0, self._update_progress, pct, d_mb, t_mb)

            if self.cancelled:
                try:
                    if os.path.exists(tpath):
                        os.remove(tpath)
                except Exception:
                    pass
                self.after(0, self.destroy)
                return

            self.after(0, self._on_download_complete)
        except Exception as exc:
            self.after(0, self._on_error, str(exc))

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


class AboutDialog(tk.Toplevel):
    """Sleek modal dialog showing application and developer information."""
    def __init__(self, parent: "BillScraperApp"):
        super().__init__(parent)
        self.title("About Bill Scrapper")
        self.configure(bg=BG_DARK)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        w, h = 380, 240
        px = parent.winfo_rootx() + (parent.winfo_width() // 2) - (w // 2)
        py = parent.winfo_rooty() + (parent.winfo_height() // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{max(0, px)}+{max(0, py)}")

        container = ttk.Frame(self, style="Root.TFrame", padding=(20, 18, 20, 18))
        container.pack(fill="both", expand=True)

        app_title = ttk.Label(container, text="Bill Scrapper", style="Title.TLabel", font=(UI_FONT, 14, "bold"))
        app_title.pack(anchor="w")

        app_ver = ttk.Label(container, text=f"Version {APP_VERSION}", style="Sub.TLabel")
        app_ver.pack(anchor="w", pady=(2, 10))

        # Card box for developer info
        info_card = ttk.Frame(container, style="Card.TFrame", padding=(14, 12))
        info_card.pack(fill="x", pady=(0, 14))

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


class BillScraperApp(tk.Tk):
    def __init__(self):
        super().__init__()
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
            else:
                png_file = resource_path("assets/billscrapper-icon.png")
                if os.path.exists(png_file):
                    self._icon_img = tk.PhotoImage(file=png_file)
                    self.iconphoto(True, self._icon_img)
        except Exception:
            pass

        # Windows 11 natively supports DWM immersive dark mode for Win32 apps.
        # On Windows 10, DWMWA_USE_IMMERSIVE_DARK_MODE causes a white title bar with
        # invisible caption buttons and black box text artifacts, so we restrict it to build >= 22000.
        if sys.platform == "win32":
            try:
                if sys.getwindowsversion().build >= 22000:
                    import ctypes
                    self.update()
                    hwnd = ctypes.windll.user32.GetParent(self.winfo_id()) or self.winfo_id()
                    val = ctypes.c_int(2)
                    ctypes.windll.dwmapi.DwmSetWindowAttribute(
                        hwnd, 20, ctypes.byref(val), ctypes.sizeof(val)
                    )
                    SWP_FLAGS = 0x0001 | 0x0002 | 0x0004 | 0x0020
                    ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, SWP_FLAGS)
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

        self.build_ui()
        self.after(100, self.drain_events)
        threading.Thread(target=self._bg_update_check, daemon=True).start()

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

        sub_lbl = ttk.Label(header_frame, text="Automated Bill Scraping & Fixed Charges Audit", style="Sub.TLabel")
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
                self.events.put(("progress", done, total, row.get("_account", ""), row.get("scrape_status", row.get("Remarks", ""))))

            PITCBillScraper(delay=delay).update_workbook(path, progress, self.stop_event)
            self.events.put(("done", str(path), self.stop_event.is_set()))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def cancel(self):
        self.stop_event.set()
        self.status_var.set("Finishing current bill, then stopping...")
        self.cancel_button.configure(state="disabled", bg=BTN_DISABLED_BG)
        self.write_log("Stop requested. Waiting for current account to finish...", "err")

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
                if event[0] == "progress":
                    _, done, total, account, status = event
                    self.progress.configure(maximum=max(total, 1), value=done)
                    self.total_var.set(str(total))
                    self.status_var.set(f"Processed {done} of {total} accounts...")

                    if status == "ok" or status == "Done" or status == "Processed":
                        self.success_count += 1
                        self.ok_var.set(str(self.success_count))
                        self.write_log(f"[{done}/{total}] {account}: Processed", "ok")
                    else:
                        self.error_count += 1
                        self.err_var.set(str(self.error_count))
                        if not status or status in ("invalid account number", "Bill not Found") or "no bill returned" in str(status).lower():
                            self.write_log(f"[{done}/{total}] {account}: Bill not Found", "err")
                        else:
                            self.write_log(f"[{done}/{total}] {account}: Bill not Found ({status})", "err")

                elif event[0] == "done":
                    _, path, cancelled = event
                    self.start_button.configure(state="normal", bg=ACCENT_GREEN)
                    self.cancel_button.configure(state="disabled", bg=BTN_DISABLED_BG)
                    self.status_var.set("Stopped by user." if cancelled else "Completed successfully!")
                    self.write_log(
                        f"=== {'Scraping Stopped.' if cancelled else 'All Done! Successfully updated:'} {path} ===",
                        "info" if not cancelled else "err"
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


if __name__ == "__main__":
    BillScraperApp().mainloop()
