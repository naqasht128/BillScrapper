"""Desktop PITC Bill Scraper & MDI Analyzer.

The selected .xlsx workbook is updated in place using its existing headings.
Excluded fields are removed before requests are made, and returned values are
written strictly below their matching headings without column shifting.
"""

from __future__ import annotations

import os
import queue
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
from decimal import Decimal
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import requests
from bs4 import BeautifulSoup
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
    ordered += [name for name in headers if name not in ordered and name not in {"feeder", "scrape_status"}]
    if "feeder" in headers:
        ordered.append("feeder")
    if "scrape_status" in headers:
        ordered.append("scrape_status")
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
            return {"_account": account, "scrape_status": "invalid account number"}
        try:
            soup = self.search(account)
            out = {"_account": account, "scrape_status": "ok"}
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
            return {"_account": account, "scrape_status": f"error: {exc}"}

    def update_workbook(self, path, progress=None, stop_event=None):
        if path.suffix.lower() != ".xlsx":
            raise RuntimeError("Please select an .xlsx workbook. The selected workbook is updated in place.")
        workbook = load_workbook(path)
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
        all_keys.add("scrape_status")

        final_headers = order_output_headers([k for k in all_keys if not is_dropped(k)])

        # 4. Write data strictly by column name - eliminating column shifting completely
        if sheet.max_column > len(final_headers):
            sheet.delete_cols(len(final_headers) + 1, sheet.max_column - len(final_headers))

        for col_idx, h in enumerate(final_headers, 1):
            sheet.cell(1, col_idx).value = h

        for row_idx, record in enumerate(row_records, 2):
            for col_idx, h in enumerate(final_headers, 1):
                val = record.get(h)
                if val is not None and should_be_numeric(h):
                    val = excel_numeric_value(val)
                sheet.cell(row_idx, col_idx).value = val if val is not None else ""

        workbook.save(path)


class BillScraperApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PITC Bill Scrapper & MDI Analyzer")
        try:
            self.iconbitmap(resource_path("assets/mdiscrapper.ico"))
        except tk.TclError:
            pass
        self.geometry("640x600")
        self.minsize(580, 520)

        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass
        self.style.configure(
            "Green.Horizontal.TProgressbar",
            troughcolor="#e2e8f0",
            background="#16a34a",
            lightcolor="#16a34a",
            darkcolor="#15803d",
            bordercolor="#cbd5e1",
        )

        self.events = queue.Queue()
        self.stop_event = threading.Event()
        self.input_var = tk.StringVar()
        self.delay_var = tk.DoubleVar(value=2.0)
        self.status_var = tk.StringVar(value="Ready. Choose an Excel (.xlsx) file containing AC numbers in Column A.")
        self.total_var = tk.StringVar(value="0")
        self.ok_var = tk.StringVar(value="0")
        self.err_var = tk.StringVar(value="0")
        self.success_count = 0
        self.error_count = 0

        self.build_ui()
        self.after(100, self.drain_events)

    def build_ui(self):
        root = ttk.Frame(self, padding=(18, 14, 18, 14))
        root.pack(fill="both", expand=True)

        header_frame = ttk.Frame(root)
        header_frame.pack(fill="x", pady=(0, 12))

        title_lbl = ttk.Label(header_frame, text="PITC Bill Scrapper & MDI Analyzer", font=("Segoe UI", 16, "bold"), foreground="#1e3a8a")
        title_lbl.pack(anchor="w")

        sub_lbl = ttk.Label(header_frame, text="Developed by Tahir Naqash | Automated Bill Scraping & Fixed Charges Audit", font=("Segoe UI", 9), foreground="#64748b")
        sub_lbl.pack(anchor="w", pady=(2, 0))

        files_frame = ttk.LabelFrame(root, text=" Target Workbook ", padding=10)
        files_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(files_frame, text="Excel File (.xlsx):", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=4)
        entry = ttk.Entry(files_frame, textvariable=self.input_var, font=("Segoe UI", 9))
        entry.grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=4)

        browse_btn = ttk.Button(files_frame, text="Browse...", command=self.choose_input)
        browse_btn.grid(row=0, column=2, pady=4)
        files_frame.columnconfigure(1, weight=1)

        mid_frame = ttk.Frame(root)
        mid_frame.pack(fill="x", pady=(0, 10))

        settings_box = ttk.LabelFrame(mid_frame, text=" Settings ", padding=8)
        settings_box.pack(side="left", fill="both", expand=True, padx=(0, 6))

        ttk.Label(settings_box, text="Delay (seconds):", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=2)
        spin = ttk.Spinbox(settings_box, from_=0.5, to=30.0, increment=0.5, textvariable=self.delay_var, width=6)
        spin.grid(row=0, column=1, sticky="w", padx=(6, 4), pady=2)
        ttk.Label(settings_box, text="(Default 2.0s prevents PITC blocking)", font=("Segoe UI", 8, "italic"), foreground="#64748b").grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))

        stats_box = ttk.LabelFrame(mid_frame, text=" Progress Summary ", padding=8)
        stats_box.pack(side="right", fill="both", expand=True, padx=(6, 0))

        ttk.Label(stats_box, text="Total:").grid(row=0, column=0, sticky="w", padx=4)
        ttk.Label(stats_box, textvariable=self.total_var, font=("Segoe UI", 9, "bold")).grid(row=0, column=1, sticky="w", padx=4)

        ttk.Label(stats_box, text="OK:").grid(row=0, column=2, sticky="w", padx=4)
        ttk.Label(stats_box, textvariable=self.ok_var, font=("Segoe UI", 9, "bold"), foreground="#15803d").grid(row=0, column=3, sticky="w", padx=4)

        ttk.Label(stats_box, text="Errors:").grid(row=0, column=4, sticky="w", padx=4)
        ttk.Label(stats_box, textvariable=self.err_var, font=("Segoe UI", 9, "bold"), foreground="#b91c1c").grid(row=0, column=5, sticky="w", padx=4)

        self.progress = ttk.Progressbar(root, mode="determinate", style="Green.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(4, 6))

        status_lbl = ttk.Label(root, textvariable=self.status_var, font=("Segoe UI", 9), foreground="#0f172a")
        status_lbl.pack(anchor="w", pady=(0, 6))

        log_frame = ttk.LabelFrame(root, text=" Live Execution Log ", padding=6)
        log_frame.pack(fill="both", expand=True, pady=(0, 10))

        self.log = tk.Text(log_frame, height=10, state="disabled", wrap="word", font=("Consolas", 9), bg="#f8fafc", fg="#0f172a")
        self.log.tag_config("ok", foreground="#16a34a")
        self.log.tag_config("err", foreground="#dc2626")
        self.log.tag_config("info", foreground="#2563eb")
        self.log.tag_config("bold", font=("Consolas", 9, "bold"))

        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        btn_bar = ttk.Frame(root)
        btn_bar.pack(fill="x")

        self.start_button = tk.Button(
            btn_bar, text="▶  Start Scraping", command=self.start,
            bg="#16a34a", fg="white", activebackground="#15803d", activeforeground="white",
            font=("Segoe UI", 10, "bold"), relief="raised", padx=16, pady=4, cursor="hand2"
        )
        self.start_button.pack(side="left")

        self.cancel_button = tk.Button(
            btn_bar, text="⏹  Stop", command=self.cancel,
            bg="#dc2626", fg="white", activebackground="#b91c1c", activeforeground="white",
            font=("Segoe UI", 10, "bold"), relief="raised", padx=14, pady=4, state="disabled", cursor="hand2"
        )
        self.cancel_button.pack(side="left", padx=(10, 0))

        self.open_btn = ttk.Button(btn_bar, text="Open File Folder", command=self.open_output_folder)
        self.open_btn.pack(side="right")

    def choose_input(self):
        path = filedialog.askopenfilename(title="Choose Account Workbook", filetypes=[("Excel Workbook", "*.xlsx")])
        if path:
            self.input_var.set(path)
            self.status_var.set(f"Selected: {Path(path).name}")

    def open_output_folder(self):
        val = self.input_var.get().strip()
        if val and Path(val).parent.exists():
            folder = str(Path(val).parent.resolve())
            if sys.platform == "win32":
                os.startfile(folder)
            else:
                subprocess.Popen(["explorer", folder])
        else:
            messagebox.showinfo("Select File", "Please select an Excel file first.")

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
        if not path.is_file():
            messagebox.showerror("Input Required", "Choose an existing .xlsx workbook first.")
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

        self.start_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.write_log(f"--- Starting Scraping on: {path.name} ---", "info")
        self.write_log(f"Request delay: {delay}s | Calculating Actual MDI (max diff x MF) & Actual Fix Charges...", "info")

        threading.Thread(target=self.run, args=(path, delay), daemon=True).start()

    def run(self, path, delay):
        try:
            def progress(done, total, row):
                self.events.put(("progress", done, total, row.get("_account", ""), row.get("scrape_status", "")))

            PITCBillScraper(delay=delay).update_workbook(path, progress, self.stop_event)
            self.events.put(("done", str(path), self.stop_event.is_set()))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def cancel(self):
        self.stop_event.set()
        self.status_var.set("Finishing current bill, then stopping...")
        self.cancel_button.configure(state="disabled")
        self.write_log("Stop requested. Waiting for current account to finish...", "err")

    def drain_events(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "progress":
                    _, done, total, account, status = event
                    self.progress.configure(maximum=max(total, 1), value=done)
                    self.total_var.set(str(total))
                    self.status_var.set(f"Processed {done} of {total} accounts...")

                    if status == "ok":
                        self.success_count += 1
                        self.ok_var.set(str(self.success_count))
                        self.write_log(f"[{done}/{total}] {account}: OK", "ok")
                    else:
                        self.error_count += 1
                        self.err_var.set(str(self.error_count))
                        self.write_log(f"[{done}/{total}] {account}: {status}", "err")

                elif event[0] == "done":
                    _, path, cancelled = event
                    self.start_button.configure(state="normal")
                    self.cancel_button.configure(state="disabled")
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
                    self.start_button.configure(state="normal")
                    self.cancel_button.configure(state="disabled")
                    self.status_var.set("Failed with error.")
                    self.write_log("ERROR: " + event[1], "err")
                    messagebox.showerror("Export Failed", event[1])
        except queue.Empty:
            pass
        self.after(100, self.drain_events)


if __name__ == "__main__":
    BillScraperApp().mainloop()
