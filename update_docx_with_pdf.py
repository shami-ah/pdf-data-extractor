#!/usr/bin/env python3
"""
Enhanced NHVAS PDF to DOCX JSON Merger
Comprehensive extraction and mapping from PDF to DOCX structure
(keep pipeline intact; fix spacing, operator info mapping, vehicle-reg header mapping, date fallback)
"""
import json
import re
import sys
from pathlib import Path
from typing import Dict, List, Any, Optional
from collections import OrderedDict  # <-- add this


def _nz(x):
    return x if isinstance(x, str) and x.strip() else ""

SUMMARY_SECTIONS = {
    "MAINTENANCE MANAGEMENT": "Maintenance Management Summary",
    "MASS MANAGEMENT": "Mass Management Summary",
    "FATIGUE MANAGEMENT": "Fatigue Management Summary",
}

# ───────────────────────────── helpers: text cleanup & label matching ─────────────────────────────
def _canon_header(s: str) -> str:
    if not s: return ""
    s = re.sub(r"\s+", " ", str(s)).strip().lower()
    s = s.replace("–", "-").replace("—", "-")
    s = re.sub(r"[/]+", " / ", s)
    s = re.sub(r"[^a-z0-9#/ ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


# Header aliases -> internal keys we already use later during mapping
_VEH_HEADER_ALIASES = {
    # common
    "registration number": "registration",
    "reg no": "registration",
    "reg.#": "registration",
    "no.": "no",
    "no": "no",

    # maintenance table
    "roadworthiness certificates": "roadworthiness",
    "maintenance records": "maintenance_records",
    "daily checks": "daily_checks",
    "fault recording reporting": "fault_recording",
    "fault recording / reporting": "fault_recording",
    "fault repair": "fault_repair",

    # mass table
    "sub contractor": "sub_contractor",
    "sub-contractor": "sub_contractor",
    "sub contracted vehicles statement of compliance": "sub_comp",
    "sub-contracted vehicles statement of compliance": "sub_comp",
    "weight verification records": "weight_verification",
    "rfs suspension certification #": "rfs_certification",
    "rfs suspension certification number": "rfs_certification",
    "suspension system maintenance": "suspension_maintenance",
    "trip records": "trip_records",
    "fault recording reporting on suspension system": "fault_reporting_suspension",
    "fault recording / reporting on suspension system": "fault_reporting_suspension",
}

# --- helpers ---
def build_vehicle_sections(extracted: dict) -> dict:
    """Build arrays for Maintenance and Mass tables. Maintenance uses recorded rows to include ALL entries."""
    maint = {
        "Registration Number": [],
        "Roadworthiness Certificates": [],
        "Maintenance Records": [],
        "Daily Checks": [],
        "Fault Recording/ Reporting": [],
        "Fault Repair": [],
    }
    mass = {
        "Registration Number": [],
        "Weight Verification Records": [],
        "RFS Suspension Certification #": [],
        "Suspension System Maintenance": [],
        "Trip Records": [],
        "Fault Recording/ Reporting on Suspension System": [],
    }

    # Prefer authoritative maintenance rows captured during parsing (spans all pages)
    if extracted.get("_maint_rows"):
        for row in extracted["_maint_rows"]:
            maint["Registration Number"].append(_smart_space(row.get("registration", "")))
            maint["Roadworthiness Certificates"].append(_nz(row.get("roadworthiness", "")))
            maint["Maintenance Records"].append(_nz(row.get("maintenance_records", "")))
            maint["Daily Checks"].append(_nz(row.get("daily_checks", "")))
            maint["Fault Recording/ Reporting"].append(_nz(row.get("fault_recording", "")))
            maint["Fault Repair"].append(_nz(row.get("fault_repair", "")))
    else:
        # Fallback to vehicles map (older behavior)
        for v in extracted.get("vehicles", []) or []:
            if not v.get("registration"): continue
            if v.get("seen_in_maintenance") or any(v.get(k) for k in ["roadworthiness","maintenance_records","daily_checks","fault_recording","fault_repair"]):
                rw = _nz(v.get("roadworthiness", "")); mr = _nz(v.get("maintenance_records", "")); dc = _nz(v.get("daily_checks", ""))
                fr = _nz(v.get("fault_recording", "")); rp = _nz(v.get("fault_repair", ""))
                if not mr and dc: mr = dc
                if not rp and fr: rp = fr
                if not fr and rp: fr = rp
                maint["Registration Number"].append(_smart_space(v["registration"]))
                maint["Roadworthiness Certificates"].append(rw)
                maint["Maintenance Records"].append(mr)
                maint["Daily Checks"].append(dc)
                maint["Fault Recording/ Reporting"].append(fr)
                maint["Fault Repair"].append(rp)

    # Mass stays as-is (from vehicles)
    for v in extracted.get("vehicles", []) or []:
        if not v.get("registration"): continue
        if v.get("seen_in_mass") or any(v.get(k) for k in ["weight_verification","rfs_certification","suspension_maintenance","trip_records","fault_reporting_suspension"]):
            mass["Registration Number"].append(_smart_space(v["registration"]))
            mass["Weight Verification Records"].append(_nz(v.get("weight_verification", "")))
            mass["RFS Suspension Certification #"].append(_nz(v.get("rfs_certification", "")))
            mass["Suspension System Maintenance"].append(_nz(v.get("suspension_maintenance", "")))
            mass["Trip Records"].append(_nz(v.get("trip_records", "")))
            mass["Fault Recording/ Reporting on Suspension System"].append(_nz(v.get("fault_reporting_suspension", "")))

    return {
        "Vehicle Registration Numbers Maintenance": maint,
        "Vehicle Registration Numbers Mass": mass,
    }


def _map_header_indices(headers: list[str]) -> dict:
    """Return {internal_key: column_index} by matching/aliasing header text."""
    idx = {}
    for i, h in enumerate(headers or []):
        ch = _canon_header(h)
        # try direct alias
        if ch in _VEH_HEADER_ALIASES:
            idx[_VEH_HEADER_ALIASES[ch]] = i
            continue
        # relax a little for 'registration number' variants
        if "registration" in ch and "number" in ch:
            idx["registration"] = i
            continue
        if "roadworthiness" in ch:
            idx["roadworthiness"] = i
            continue
        if "maintenance" in ch and "records" in ch:
            idx["maintenance_records"] = i
            continue
        if "daily" in ch and "check" in ch:
            idx["daily_checks"] = i
            continue
        if "fault" in ch and "record" in ch and "suspension" not in ch:
            # maintenance fault-recording column
            if "repair" in ch:
                idx["fault_repair"] = i
            else:
                idx["fault_recording"] = i
            continue
        if "weight" in ch and "verification" in ch:
            idx["weight_verification"] = i
            continue
        if "rfs" in ch and "certification" in ch:
            idx["rfs_certification"] = i
            continue
        if "suspension" in ch and "maintenance" in ch:
            idx["suspension_maintenance"] = i
            continue
        if "trip" in ch and "record" in ch:
            idx["trip_records"] = i
            continue
        if "fault" in ch and "report" in ch and "suspension" in ch:
            idx["fault_reporting_suspension"] = i
            continue
    return idx

def _canon(s: str) -> str:
    if not s: return ""
    s = re.sub(r"\s+", " ", str(s)).strip().lower()
    s = re.sub(r"[^a-z0-9#]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def _smart_space(s: str) -> str:
    if not s: return s
    s = str(s)

    # Insert spaces at typical OCR glue points
    s = re.sub(r'([a-z])([A-Z])', r'\1 \2', s)
    s = re.sub(r'([A-Za-z])(\d)', r'\1 \2', s)
    s = re.sub(r'(\d)([A-Za-z])', r'\1 \2', s)
    s = re.sub(r'([A-Z]{2,})(\d)', r'\1 \2', s)

    # Fix common glued tokens
    s = s.replace("POBox", "PO Box")

    # Compact ordinals back together: "9 th" -> "9th", but preserve a space after the ordinal if followed by a word
    s = re.sub(r'\b(\d+)\s*(st|nd|rd|th)\b', r'\1\2', s)

    s = re.sub(r"\s+", " ", s).strip()
    return s

def looks_like_plate(s: str) -> bool:
    if not s: return False
    t = re.sub(r"[\s-]", "", str(s).upper())
    if not (5 <= len(t) <= 8): return False
    if not re.fullmatch(r"[A-Z0-9]+", t): return False
    if sum(c.isalpha() for c in t) < 2: return False
    if sum(c.isdigit() for c in t) < 2: return False
    if t in {"ENTRY","YES","NO","N/A","NA"}: return False
    return True

def is_dateish(s: str) -> bool:
    if not s: return False
    s = _smart_space(s)
    # tokens like 03/22, 20/02/2023, 01.02.21, 2023-02-20
    return bool(re.search(r"\b\d{1,4}(?:[./-]\d{1,2}){1,2}\b", s))

def extract_date_tokens(s: str) -> list[str]:
    if not s: return []
    s = _smart_space(s)
    return re.findall(r"\b\d{1,4}(?:[./-]\d{1,2}){1,2}\b", s)


def _clean_list(vals: List[str]) -> List[str]:
    out = []
    for v in vals:
        v = _smart_space(v)
        if v:
            out.append(v)
    return out

def _looks_like_manual_value(s: str) -> bool:
    if not s: return False
    s = s.strip()
    # reject pure digits (e.g., "51902") and very short tokens
    if re.fullmatch(r"\d{3,}", s): 
        return False
    # accept if it has any letters or typical version hints
    return bool(re.search(r"[A-Za-z]", s))

def _looks_like_company(s: str) -> bool:
    """Very light validation to avoid capturing labels as values."""
    if not s: return False
    s = _smart_space(s)
    # at least two words containing letters (e.g., "Kangaroo Transport")
    return bool(re.search(r"[A-Za-z]{2,}\s+[A-Za-z&]{2,}", s))

# ───────────────────────────── label index (non-summary only; no values) ─────────────────────────────
LABEL_INDEX: Dict[str, Dict[str, Dict[str, Any]]] = {
    "Audit Information": {
        "Date of Audit": {"alts": ["Date of Audit"]},
        "Location of audit": {"alts": ["Location of audit", "Location"]},
        "Auditor name": {"alts": ["Auditor name", "Auditor"]},
        "Audit Matrix Identifier (Name or Number)": {"alts": ["Audit Matrix Identifier (Name or Number)", "Audit Matrix Identifier"]},
        "Auditor Exemplar Global Reg No.": {"alts": ["Auditor Exemplar Global Reg No."]},
        "NHVR Auditor Registration Number": {"alts": ["NHVR Auditor Registration Number"]},
        "expiry Date:": {"alts": ["expiry Date:", "Expiry Date:"]},
    },
    "Operator Information": {
        "Operator name (Legal entity)": {"alts": ["Operator name (Legal entity)", "Operator's Name (legal entity)"]},
        "NHVAS Accreditation No. (If applicable)": {"alts": ["NHVAS Accreditation No. (If applicable)", "NHVAS Accreditation No."]},
        "Registered trading name/s": {"alts": ["Registered trading name/s", "Trading name/s"]},
        "Australian Company Number": {"alts": ["Australian Company Number", "ACN"]},
        "NHVAS Manual (Policies and Procedures) developed by": {"alts": [
            "NHVAS Manual (Policies and Procedures) developed by",
            "NHVAS Manual developed by",
            "Manual developed by"
        ]},
    },
    "Operator contact details": {
        "Operator business address": {"alts": ["Operator business address", "Business address"]},
        "Operator Postal address": {"alts": ["Operator Postal address", "Postal address"]},
        "Email address": {"alts": ["Email address", "Email"]},
        "Operator Telephone Number": {"alts": ["Operator Telephone Number", "Telephone", "Phone"]},
    },
    "Attendance List (Names and Position Titles)": {
        "Attendance List (Names and Position Titles)": {"alts": ["Attendance List (Names and Position Titles)", "Attendance List"]},
    },
    "Nature of the Operators Business (Summary)": {
        "Nature of the Operators Business (Summary):": {"alts": ["Nature of the Operators Business (Summary):"]},
    },
    "Accreditation Vehicle Summary": {
        "Number of powered vehicles": {"alts": ["Number of powered vehicles"]},
        "Number of trailing vehicles": {"alts": ["Number of trailing vehicles"]},
    },
    "Accreditation Driver Summary": {
        "Number of drivers in BFM": {"alts": ["Number of drivers in BFM"]},
        "Number of drivers in AFM": {"alts": ["Number of drivers in AFM"]},
    },
    "Vehicle Registration Numbers Maintenance": {
        "No.": {"alts": ["No.", "No"]},
        "Registration Number": {"alts": ["Registration Number", "Registration"]},
        "Roadworthiness Certificates": {"alts": ["Roadworthiness Certificates", "Roadworthiness"]},
        "Maintenance Records": {"alts": ["Maintenance Records"]},
        "Daily Checks": {"alts": ["Daily Checks", "Daily Check"]},
        "Fault Recording/ Reporting": {"alts": ["Fault Recording/ Reporting", "Fault Recording / Reporting"]},
        "Fault Repair": {"alts": ["Fault Repair"]},
    },
    "Vehicle Registration Numbers Mass": {
        "No.": {"alts": ["No.", "No"]},
        "Registration Number": {"alts": ["Registration Number", "Registration"]},
        "Sub contractor": {"alts": ["Sub contractor", "Sub-contractor"]},
        "Sub-contracted Vehicles Statement of Compliance": {"alts": ["Sub-contracted Vehicles Statement of Compliance"]},
        "Weight Verification Records": {"alts": ["Weight Verification Records"]},
        "RFS Suspension Certification #": {"alts": ["RFS Suspension Certification #", "RFS Suspension Certification Number"]},
        "Suspension System Maintenance": {"alts": ["Suspension System Maintenance"]},
        "Trip Records": {"alts": ["Trip Records"]},
        "Fault Recording/ Reporting on Suspension System": {"alts": ["Fault Recording/ Reporting on Suspension System"]},
    },
    "Driver / Scheduler Records Examined": {
        "No.": {"alts": ["No.", "No"]},
        "Driver / Scheduler Name": {"alts": ["Driver / Scheduler Name"]},
        "Driver TLIF Course # Completed": {"alts": ["Driver TLIF Course # Completed"]},
        "Scheduler TLIF Course # Completed": {"alts": ["Scheduler TLIF Course # Completed"]},
        "Medical Certificates (Current Yes/No) Date of expiry": {"alts": ["Medical Certificates (Current Yes/No) Date of expiry"]},
        "Roster / Schedule / Safe Driving Plan (Date Range)": {"alts": ["Roster / Schedule / Safe Driving Plan (Date Range)"]},
        "Fit for Duty Statement Completed (Yes/No)": {"alts": ["Fit for Duty Statement Completed (Yes/No)"]},
        "Work Diary Pages (Page Numbers) Electronic Work Diary Records (Date Range)": {"alts": ["Work Diary Pages (Page Numbers) Electronic Work Diary Records (Date Range)"]},
    },
    "NHVAS Approved Auditor Declaration": {
        "Print Name": {"alts": ["Print Name"]},
        "NHVR or Exemplar Global Auditor Registration Number": {"alts": ["NHVR or Exemplar Global Auditor Registration Number"]},
    },
    "Audit Declaration dates": {
        "Audit was conducted on": {"alts": ["Audit was conducted on"]},
        "Unconditional CARs closed out on:": {"alts": ["Unconditional CARs closed out on:"]},
        "Conditional CARs to be closed out by:": {"alts": ["Conditional CARs to be closed out by:"]},
    },
    "Print accreditation name": {
        "(print accreditation name)": {"alts": ["(print accreditation name)"]},
    },
    "Operator Declaration": {
        "Print Name": {"alts": ["Print Name"]},
        "Position Title": {"alts": ["Position Title"]},
    },
}

class NHVASMerger:
    def __init__(self):
        self.debug_mode = True
        self._vehicle_by_reg = OrderedDict()

    def log_debug(self, msg: str):
        if self.debug_mode:
            print(f"🔍 {msg}")

    def normalize_std_label(self, label: str) -> str:
        if not label: return ""
        base = re.sub(r"\([^)]*\)", "", label)
        base = re.sub(r"\s+", " ", base).strip()
        m = re.match(r"^(Std\s*\d+\.\s*[^:]+?)\s*$", base, flags=re.IGNORECASE)
        return m.group(1).strip() if m else base

    def _pick_nearby(self, row, anchor_idx: int | None, want: str = "plate", window: int = 3) -> str:
        """Return the best cell for a field by looking at the anchor index and nearby columns.
        want ∈ {"plate","date","rf","yn"}"""
        def cell(i):
            if i is None or i < 0 or i >= len(row): return ""
            v = row[i]
            return v.strip() if isinstance(v, str) else str(v).strip()

        # 1) try the anchor cell
        cand = cell(anchor_idx)
        if want == "plate" and looks_like_plate(cand): return _smart_space(cand)
        if want == "date"  and is_dateish(cand):      return _smart_space(cand)
        if want == "rf"    and re.search(r"\bRF\s*\d+\b", cand, re.I): return _smart_space(re.search(r"\bRF\s*\d+\b", cand, re.I).group(0))
        if want == "yn"    and cand.strip().lower() in {"yes","no"}:   return cand.strip().title()

        # 2) scan a window around the anchor
        if anchor_idx is not None:
            for offset in range(1, window+1):
                for i in (anchor_idx - offset, anchor_idx + offset):
                    c = cell(i)
                    if not c: continue
                    if want == "plate" and looks_like_plate(c): return _smart_space(c)
                    if want == "date"  and is_dateish(c):      return _smart_space(c)
                    if want == "rf":
                        m = re.search(r"\bRF\s*\d+\b", c, re.I)
                        if m: return _smart_space(m.group(0))
                    if want == "yn" and c.strip().lower() in {"yes","no"}: return c.strip().title()

        # 3) last resort: scan whole row
        joined = " ".join(str(c or "") for c in row)
        if want == "plate":
            for tok in joined.split():
                if looks_like_plate(tok): return _smart_space(tok)
        if want == "date":
            tok = extract_date_tokens(joined)
            return tok[0] if tok else ""
        if want == "rf":
            m = re.search(r"\bRF\s*\d+\b", joined, re.I)
            return _smart_space(m.group(0)) if m else ""
        if want == "yn":
            j = f" {joined.lower()} "
            if " yes " in j: return "Yes"
            if " no "  in j: return "No"
        return ""


    def _force_fill_maintenance_from_tables(self, pdf_data: Dict, merged: Dict) -> None:
        """Overwrite Maintenance arrays by scanning ALL maintenance tables across pages."""
        maint = merged.get("Vehicle Registration Numbers Maintenance")
        if not isinstance(maint, dict):
            return

        tables = (pdf_data.get("extracted_data") or {}).get("all_tables") or []
        regs, rw, mr, dc, fr, rp = [], [], [], [], [], []

        for t in tables:
            hdrs = [_canon_header(h or "") for h in t.get("headers") or []]
            if not hdrs:
                continue
            # detect a maintenance table
            txt = " ".join(hdrs)
            if ("registration" not in txt) or not any(
                k in txt for k in ["maintenance records", "daily", "fault recording", "fault repair", "roadworthiness"]
            ):
                continue

            def fidx(pred):
                for i, h in enumerate(hdrs):
                    if pred(h):
                        return i
                return None

            reg_i   = fidx(lambda h: "registration" in h)
            rw_i    = fidx(lambda h: "roadworthiness" in h)
            mr_i    = fidx(lambda h: "maintenance" in h and "record" in h)
            dc_i    = fidx(lambda h: "daily" in h and "check" in h)
            fr_i    = fidx(lambda h: "fault" in h and "record" in h and "suspension" not in h)
            rp_i    = fidx(lambda h: "fault" in h and "repair" in h)

            for r in t.get("data") or []:
                def cell(i):
                    if i is None or i >= len(r): return ""
                    v = r[i]
                    return v.strip() if isinstance(v, str) else str(v).strip()

                plate = _smart_space(cell(reg_i))
                if not plate or not looks_like_plate(plate):
                    continue

                v_rw = _nz(cell(rw_i))
                v_mr = _nz(cell(mr_i))
                v_dc = _nz(cell(dc_i))
                v_fr = _nz(cell(fr_i))
                v_rp = _nz(cell(rp_i))

                # sensible fallbacks
                if not v_mr and v_dc: v_mr = v_dc
                if not v_rp and v_fr: v_rp = v_fr
                if not v_fr and v_rp: v_fr = v_rp

                regs.append(plate); rw.append(v_rw); mr.append(v_mr)
                dc.append(v_dc);    fr.append(v_fr); rp.append(v_rp)

        if regs:  # overwrite arrays only if we found rows
            maint["Registration Number"] = regs
            maint["Roadworthiness Certificates"] = rw
            maint["Maintenance Records"] = mr
            maint["Daily Checks"] = dc
            maint["Fault Recording/ Reporting"] = fr
            maint["Fault Repair"] = rp

    def _collapse_multiline_headers(self, headers: List[str], data_rows: List[List[str]]):
        """
        Merge header continuation rows (when first data rows are not numeric '1.', '2.', …)
        into the main headers, then return (merged_headers, remaining_data_rows).
        """
        merged = [_smart_space(h or "") for h in (headers or [])]
        consumed = 0
        header_frags: List[List[str]] = []

        # Collect up to 5 leading rows that look like header fragments
        for r in data_rows[:5]:
            first = (str(r[0]).strip() if r else "")
            if re.match(r"^\d+\.?$", first):
                break  # real data starts
            consumed += 1
            header_frags.append(r)

        # Merge every collected fragment row into merged
        for frag in header_frags:
            for i, cell in enumerate(frag):
                cell_txt = _smart_space(str(cell or "").strip())
                if not cell_txt:
                    continue
                if i >= len(merged):
                    merged.append(cell_txt)
                else:
                    merged[i] = (merged[i] + " " + cell_txt).strip()

        return merged, data_rows[consumed:]

    def _first_attendance_name_title(self, att_list: List[str]) -> Optional[tuple[str, str]]:
        """Return (print_name, position_title) from the first 'Name - Title' in attendance."""
        if not att_list:
            return None
        # First "Name - Title", stop before next "Name -"
        pat = re.compile(
            r'([A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,3})\s*-\s*(.*?)(?=(?:\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,3}\s*-\s*)|$)'
        )
        for item in att_list:
            s = _smart_space(str(item))
            m = pat.search(s)
            if m:
                name = _smart_space(m.group(1))
                title = _smart_space(m.group(2))
                return name, title
        return None

    
    # ───────────────────────────── summary tables (unchanged logic) ─────────────────────────────
    def build_summary_maps(self, pdf_json: dict) -> dict:
        out = {v: {} for v in SUMMARY_SECTIONS.values()}
        try:
            tables = pdf_json["extracted_data"]["all_tables"]
        except Exception:
            return out

        for t in tables:
            headers = [re.sub(r"\s+", " ", (h or "")).strip().upper() for h in t.get("headers", [])]
            if "DETAILS" not in headers:
                continue
            section_key_raw = next((h for h in headers if h in SUMMARY_SECTIONS), None)
            if not section_key_raw:
                continue
            section_name = SUMMARY_SECTIONS[section_key_raw]
            for row in t.get("data", []):
                if not row: continue
                left = str(row[0]) if len(row) >= 1 else ""
                right = str(row[1]) if len(row) >= 2 else ""
                left_norm = self.normalize_std_label(left)
                if left_norm and right:
                    prev = out[section_name].get(left_norm, "")
                    merged_text = (prev + " " + right).strip() if prev else right.strip()
                    out[section_name][left_norm] = merged_text

        for sec in out:
            out[sec] = {k: [_smart_space(v)] for k, v in out[sec].items() if v}
        return out

    # ───────────────────────────── NEW: find cell by label in tables ─────────────────────────────
    def _find_table_value(self, tables: List[Dict], label_variants: List[str]) -> Optional[str]:
        targets = {_canon(v) for v in label_variants}
        for t in tables:
            data = t.get("data", [])
            if not data: continue
            for row in data:
                if not row: continue
                key = _canon(str(row[0]))
                if key in targets:
                    vals = [str(c).strip() for c in row[1:] if str(c).strip()]
                    if vals:
                        return _smart_space(" ".join(vals))
        return None

    # ───────────────────────────── comprehensive extraction (minimal changes) ─────────────────────────────
    def extract_from_pdf_comprehensive(self, pdf_data: Dict) -> Dict[str, Any]:
        self._vehicle_by_reg.clear()
        extracted = {}
        extracted_data = pdf_data.get("extracted_data", {})
        tables = extracted_data.get("all_tables", [])

        # Capture "Audit was conducted on" from tables; ignore placeholder "Date"
        awd = self._find_table_value(
            tables,
            LABEL_INDEX["Audit Declaration dates"]["Audit was conducted on"]["alts"]
        )
        if awd:
            awd = _smart_space(awd)
            if re.search(r"\d", awd) and not re.fullmatch(r"date", awd, re.I):
                extracted["audit_conducted_date"] = awd



        # 1) Audit Information (table first)
        audit_info = extracted_data.get("audit_information", {})
        if audit_info:
            extracted["audit_info"] = {
                "date_of_audit": _smart_space(audit_info.get("DateofAudit", "")),
                "location": _smart_space(audit_info.get("Locationofaudit", "")),
                "auditor_name": _smart_space(audit_info.get("Auditorname", "")),
                "matrix_id": _smart_space(audit_info.get("AuditMatrixIdentifier (Name or Number)", "")),
            }
        # If missing, try generic table lookup
        for label, meta in LABEL_INDEX.get("Audit Information", {}).items():
            if label == "expiry Date:":  # not used in your DOCX example
                continue
            val = self._find_table_value(tables, meta.get("alts", [label]))
            if val:
                extracted.setdefault("audit_info", {})
                if _canon(label) == _canon("Date of Audit"): extracted["audit_info"]["date_of_audit"] = val
                elif _canon(label) == _canon("Location of audit"): extracted["audit_info"]["location"] = val
                elif _canon(label) == _canon("Auditor name"): extracted["audit_info"]["auditor_name"] = val
                elif _canon(label) == _canon("Audit Matrix Identifier (Name or Number)"): extracted["audit_info"]["matrix_id"] = val

        # 2) Operator Information (prefer table rows)
        operator_info = extracted_data.get("operator_information", {})
        if operator_info:
            extracted["operator_info"] = {
                "name": "",
                "trading_name": _smart_space(operator_info.get("trading_name", "")),
                "acn": _smart_space(operator_info.get("company_number", "")),
                "manual": _smart_space(operator_info.get("nhvas_accreditation", "")),
                "business_address": _smart_space(operator_info.get("business_address", "")),
                "postal_address": _smart_space(operator_info.get("postal_address", "")),
                "email": operator_info.get("email", ""),
                "phone": _smart_space(operator_info.get("phone", "")),
            }

        # Fill operator info via table lookup
        for label, meta in LABEL_INDEX.get("Operator Information", {}).items():
            val = self._find_table_value(tables, meta.get("alts", [label]))
            if not val: continue
            if _canon(label) == _canon("Operator name (Legal entity)") and _looks_like_company(val):
                extracted.setdefault("operator_info", {})
                extracted["operator_info"]["name"] = val
            elif _canon(label) == _canon("Registered trading name/s"):
                extracted.setdefault("operator_info", {})
                extracted["operator_info"]["trading_name"] = val
            elif _canon(label) == _canon("Australian Company Number"):
                extracted.setdefault("operator_info", {})
                extracted["operator_info"]["acn"] = val
            elif _canon(label) == _canon("NHVAS Manual (Policies and Procedures) developed by"):
                extracted.setdefault("operator_info", {})
                if _looks_like_manual_value(val):
                    extracted["operator_info"]["manual"] = val

        # 3) Generic table parsing (unchanged logic for other sections)
        self._extract_table_data(tables, extracted)

        # 4) Text parsing (kept, but spacing applied)
        self._extract_text_content(extracted_data.get("all_text_content", []), extracted)
        # Vehicle tables sometimes fail to land in all_tables; parse from text as a fallback
        self._extract_vehicle_tables_from_text(extracted_data.get("all_text_content", []), extracted)

        # 5) Vehicle/Driver data (kept)
        self._extract_vehicle_driver_data(extracted_data, extracted)

        # 6) Detailed mgmt (kept)
        self._extract_detailed_management_data(extracted_data, extracted)

        return extracted

    # ───────────────────────────── table classifiers ─────────────────────────────
    # replace your _extract_table_data with this version
    def _extract_table_data(self, tables: List[Dict], extracted: Dict):
        for table in tables:
            headers   = table.get("headers", []) or []
            data_rows = table.get("data", []) or []
            if not data_rows:
                continue

            page_num = table.get("page", 0)
            self.log_debug(f"Processing table on page {page_num} with headers: {headers[:3]}...")

            # 🔧 NEW: collapse possible multi-line headers once up front
            collapsed_headers, collapsed_rows = self._collapse_multiline_headers(headers, data_rows)

            # 🔧 Try vehicle tables FIRST using either raw or collapsed headers
            if self._is_vehicle_registration_table(headers) or self._is_vehicle_registration_table(collapsed_headers):
                # always extract with the collapsed header/rows so we see "Registration Number", etc.
                self._extract_vehicle_registration_table(collapsed_headers, collapsed_rows, extracted, page_num)
                continue

            # the rest keep their existing order/logic (use the original headers/rows)
            if self._is_audit_info_table(headers):
                self._extract_audit_info_table(data_rows, extracted)
            elif self._is_operator_info_table(headers):
                self._extract_operator_info_table(data_rows, extracted)
            elif self._is_attendance_table(headers):
                self._extract_attendance_table(data_rows, extracted)
            elif self._is_vehicle_summary_table(headers):
                self._extract_vehicle_summary_table(data_rows, extracted)
            elif self._is_driver_table(headers):
                self._extract_driver_table(headers, data_rows, extracted)
            elif self._is_management_compliance_table(headers):
                self._extract_management_table(data_rows, extracted, headers)


    def _is_audit_info_table(self, headers: List[str]) -> bool:
        txt = " ".join(str(h) for h in headers).lower()
        return any(t in txt for t in ["audit", "date", "location", "auditor"])

    def _is_operator_info_table(self, headers: List[str]) -> bool:
        txt = " ".join(str(h) for h in headers).lower()
        return any(t in txt for t in ["operator", "company", "trading", "address"])

    def _is_attendance_table(self, headers: List[str]) -> bool:
        txt = " ".join(str(h) for h in headers).lower()
        return "attendance" in txt

    def _is_vehicle_summary_table(self, headers: List[str]) -> bool:
        txt = " ".join(str(h) for h in headers).lower()
        return any(t in txt for t in ["powered vehicles", "trailing vehicles", "drivers in bfm"])

    def _is_vehicle_registration_table(self, headers: List[str]) -> bool:
        if not headers: return False
        ch = [_canon_header(h) for h in headers]
        has_reg = any(
            ("registration" in h) or re.search(r"\breg(?:istration)?\b", h) or ("reg" in h and "no" in h)
            for h in ch
        )
        others = ["roadworthiness","maintenance records","daily checks","fault recording","fault repair",
                "sub contractor","sub-contractor","weight verification","rfs suspension","suspension system maintenance",
                "trip records","fault recording reporting on suspension system","fault reporting suspension"]
        has_signal = any(any(tok in h for tok in others) for h in ch)
        return has_reg and has_signal

    def _is_driver_table(self, headers: List[str]) -> bool:
        txt = " ".join(str(h) for h in headers).lower()
        return any(t in txt for t in ["driver", "scheduler", "tlif", "medical"])

    def _is_management_compliance_table(self, headers: List[str]) -> bool:
        txt = " ".join(str(h) for h in headers).lower()
        return any(t in txt for t in ["maintenance management", "mass management", "fatigue management"])

    def _extract_vehicle_tables_from_text(self, text_pages: List[Dict], extracted: Dict):
        # flatten text
        lines = []
        for p in text_pages or []:
            for ln in re.split(r"\s*\n\s*", p.get("text", "")):
                ln = _smart_space(ln)
                if ln: lines.append(ln)

        maint_rows, mass_rows = [], []
        rf_pat = re.compile(r"\bRF\s*\d+\b", re.IGNORECASE)

        for ln in lines:
            # find first token that looks like a rego
            tokens = ln.split()
            reg = next((t for t in tokens if looks_like_plate(t)), None)
            if not reg: 
                continue

            # everything after the reg on that line
            tail = _smart_space(ln.split(reg, 1)[1]) if reg in ln else ""
            dates = extract_date_tokens(tail)
            has_rf = bool(rf_pat.search(ln)) or "suspension" in ln.lower()

            if has_rf:
                rfs = (rf_pat.search(ln).group(0).upper().replace(" ", "") if rf_pat.search(ln) else "")
                wv = dates[0] if len(dates) > 0 else ""
                rest = dates[1:]
                mass_rows.append({
                    "registration": reg,
                    "sub_contractor": "Yes" if " yes " in f" {ln.lower()} " else ("No" if " no " in f" {ln.lower()} " else ""),
                    "sub_comp":      "Yes" if " yes " in f" {ln.lower()} " else ("No" if " no " in f" {ln.lower()} " else ""),
                    "weight_verification": wv,
                    "rfs_certification": rfs or ("N/A" if "n/a" in ln.lower() else ""),
                    "suspension_maintenance": rest[0] if len(rest) > 0 else "",
                    "trip_records":            rest[1] if len(rest) > 1 else "",
                    "fault_reporting_suspension": rest[2] if len(rest) > 2 else "",
                })
            else:
                # map first 5 date-like tokens in sensible order; fallbacks keep table consistent
                rw = dates[0] if len(dates) > 0 else ""
                mr = dates[1] if len(dates) > 1 else ""
                dc = dates[2] if len(dates) > 2 else ""
                fr = dates[3] if len(dates) > 3 else ""
                rp = dates[4] if len(dates) > 4 else ""
                maint_rows.append({
                    "registration": reg,
                    "roadworthiness": rw,
                    "maintenance_records": mr or dc,
                    "daily_checks": dc,
                    "fault_recording": fr or rp,
                    "fault_repair": rp or fr,
                })

            # ... after building maint_rows and mass_rows ...
        vlist = extracted.setdefault("vehicles", [])  # ensure it always exists

        if maint_rows or mass_rows:
            for r in maint_rows:
                r["section"] = "maintenance"
                vlist.append(r)
            for r in mass_rows:
                r["section"] = "mass"
                vlist.append(r)
            self.log_debug(f"Vehicle rows (text fallback): maint={len(maint_rows)} mass={len(mass_rows)} total={len(vlist)}")
        else:
            self.log_debug("Vehicle rows (text fallback): none detected.")


    # ───────────────────────────── simple extractors (spacing applied) ─────────────────────────────
    def _extract_audit_info_table(self, data_rows: List[List], extracted: Dict):
        ai = extracted.setdefault("audit_info", {})
        for row in data_rows:
            if len(row) < 2: continue
            key = _canon(row[0])
            val = _smart_space(" ".join(str(c).strip() for c in row[1:] if str(c).strip()))
            if not val: continue
            if "date" in key and "audit" in key: ai["date_of_audit"] = val
            elif "location" in key: ai["location"] = val
            elif "auditor" in key and "name" in key: ai["auditor_name"] = val
            elif "matrix" in key: ai["matrix_id"] = val

    def _extract_operator_info_table(self, data_rows: List[List], extracted: Dict):
        oi = extracted.setdefault("operator_info", {})
        for row in data_rows:
            if len(row) < 2: continue
            key = _canon(row[0])
            val = _smart_space(" ".join(str(c).strip() for c in row[1:] if str(c).strip()))
            if not val: continue
            if "operator" in key and "name" in key and _looks_like_company(val): oi["name"] = val
            elif "trading" in key: oi["trading_name"] = val
            elif "australian" in key and "company" in key: oi["acn"] = val
            elif "business" in key and "address" in key: oi["business_address"] = val
            elif "postal" in key and "address" in key: oi["postal_address"] = val
            elif "email" in key: oi["email"] = val
            elif "telephone" in key or "phone" in key: oi["phone"] = val
            elif "manual" in key or ("nhvas" in key and "manual" in key) or "developed" in key:
                if _looks_like_manual_value(val):
                    oi["manual"] = val

    def _extract_attendance_table(self, data_rows: List[List], extracted: Dict):
        lst = []
        for row in data_rows:
            if not row: continue
            cells = [str(c).strip() for c in row if str(c).strip()]
            if not cells: continue
            lst.append(_smart_space(" ".join(cells)))
        if lst:
            extracted["attendance"] = lst

    def _extract_vehicle_summary_table(self, data_rows: List[List], extracted: Dict):
        vs = extracted.setdefault("vehicle_summary", {})
        for row in data_rows:
            if len(row) < 2: continue
            key = _canon(row[0])
            value = ""
            for c in row[1:]:
                if str(c).strip():
                    value = _smart_space(str(c).strip()); break
            if not value: continue
            if "powered" in key and "vehicle" in key: vs["powered_vehicles"] = value
            elif "trailing" in key and "vehicle" in key: vs["trailing_vehicles"] = value
            elif "drivers" in key and "bfm" in key: vs["drivers_bfm"] = value
            elif "drivers" in key and "afm" in key: vs["drivers_afm"] = value

    # ▶▶ REPLACED: column mapping by headers
    def _extract_vehicle_registration_table(self, headers, rows, extracted, page_num):
        ch    = [_canon_header(h) for h in (headers or [])]
        alias = _map_header_indices(headers or [])

        # header indices (may be misaligned vs data; that's OK, we’ll search near them)
        def idx_of(*needles):
            for i, h in enumerate(ch):
                if all(n in h for n in needles): return i
            return None

        reg_i   = alias.get("registration") or idx_of("registration number") or idx_of("registration") or idx_of("reg","no")
        rw_i    = alias.get("roadworthiness") or idx_of("roadworthiness")
        maint_i = alias.get("maintenance_records") or idx_of("maintenance","records")
        daily_i = alias.get("daily_checks") or idx_of("daily","check")
        fr_i    = alias.get("fault_recording") or idx_of("fault","recording")
        rep_i   = alias.get("fault_repair")    or idx_of("fault","repair")

        weight_i = alias.get("weight_verification") or idx_of("weight","verification")
        rfs_i    = alias.get("rfs_certification")   or idx_of("rfs","certification")
        susp_i   = alias.get("suspension_maintenance") or idx_of("suspension","maintenance")
        trip_i   = alias.get("trip_records") or idx_of("trip","records")
        frs_i    = alias.get("fault_reporting_suspension") or idx_of("fault","reporting","suspension")

        # classify table type by header signals
        is_maint = any("roadworthiness" in h or "maintenance records" in h or ("daily" in h and "check" in h) or "fault repair" in h for h in ch)
        is_mass  = any("weight verification" in h or "rfs" in h or "suspension system" in h or "trip records" in h or "reporting on suspension" in h for h in ch)

        maint_rows = extracted.setdefault("_maint_rows", []) if is_maint else None
        added = 0

        for r in rows or []:
            # tolerant plate pick (handles misaligned columns)
            reg = self._pick_nearby(r, reg_i, "plate", window=4)
            if not reg or not looks_like_plate(reg):
                continue

            # collect values using tolerant picks
            if is_maint:
                rw  = self._pick_nearby(r, rw_i,    "date", window=4)
                mr  = self._pick_nearby(r, maint_i, "date", window=4)
                dc  = self._pick_nearby(r, daily_i, "date", window=4)
                fr  = self._pick_nearby(r, fr_i,    "date", window=4)
                rep = self._pick_nearby(r, rep_i,   "date", window=4)

                # sensible fallbacks
                if not mr and dc: mr = dc
                if not rep and fr: rep = fr
                if not fr and rep: fr = rep

            else:  # mass or mixed
                wv  = self._pick_nearby(r, weight_i, "date", window=4)
                rfs = self._pick_nearby(r, rfs_i,    "rf",   window=5)
                sm  = self._pick_nearby(r, susp_i,   "date", window=4)
                tr  = self._pick_nearby(r, trip_i,   "date", window=4)
                frs = self._pick_nearby(r, frs_i,    "date", window=4)
                yn1 = self._pick_nearby(r, idx_of("sub","contractor"), "yn", window=3) or ""
                yn2 = self._pick_nearby(r, idx_of("sub contracted vehicles statement of compliance"), "yn", window=3) or yn1

            # merge into vehicle map
            v = self._vehicle_by_reg.get(reg)
            if v is None:
                v = {"registration": reg}
                self._vehicle_by_reg[reg] = v
                added += 1

            if is_maint:
                v["seen_in_maintenance"] = True
                if rw:  v.setdefault("roadworthiness", rw)
                if mr:  v.setdefault("maintenance_records", mr)
                if dc:  v.setdefault("daily_checks", dc)
                if fr:  v.setdefault("fault_recording", fr)
                if rep: v.setdefault("fault_repair", rep)

                if maint_rows is not None:
                    maint_rows.append({
                        "registration": reg,
                        "roadworthiness": rw,
                        "maintenance_records": mr or dc,
                        "daily_checks": dc,
                        "fault_recording": fr or rep,
                        "fault_repair": rep or fr,
                    })
            else:
                v["seen_in_mass"] = True
                if yn1: v.setdefault("sub_contractor", yn1)
                if yn2: v.setdefault("sub_comp", yn2)
                if wv:  v.setdefault("weight_verification", wv)
                if rfs: v.setdefault("rfs_certification", _smart_space(rfs).upper().replace(" ", ""))
                if sm:  v.setdefault("suspension_maintenance", sm)
                if tr:  v.setdefault("trip_records", tr)
                if frs: v.setdefault("fault_reporting_suspension", frs)

        extracted["vehicles"] = list(self._vehicle_by_reg.values())
        return added

    def _extract_driver_table(self, headers: List[str], data_rows: List[List], extracted: Dict):
        """Header-driven extraction for Driver / Scheduler Records."""
        drivers = []
        ch = [_canon_header(h) for h in headers or []]

        # helpers
        def find_col(needles: list[str]) -> Optional[int]:
            for i, h in enumerate(ch):
                if any(n in h for n in needles):
                    return i
            return None

        def find_col_rx(patterns: list[str]) -> Optional[int]:
            for i, h in enumerate(ch):
                if any(re.search(p, h) for p in patterns):
                    return i
            return None

        name_idx   = find_col_rx([r"\bdriver\s*/\s*scheduler\s*name\b",
                              r"\bdriver\s+name\b", r"\bscheduler\s+name\b", r"\bname\b"])
        tlif_d_idx = find_col(["driver tlif"])
        tlif_s_idx = find_col(["scheduler tlif"])
        medical_idx= find_col(["medical", "expiry"])
        roster_idx = find_col_rx([r"\broster\b", r"\bsafe\s+driving\s+plan\b", r"\bschedule\b(?!r\b)"])
        fit_idx    = find_col(["fit for duty"])
        diary_idx  = find_col(["work diary", "electronic work diary", "page numbers"])

        for row in data_rows:
            if not row:
                continue

            name = None
            if name_idx is not None and name_idx < len(row):
                name = _smart_space(str(row[name_idx]).strip())
            if not name:
                continue

            d = {"name": name}

            if tlif_d_idx is not None and tlif_d_idx < len(row):
                d["driver_tlif"] = _smart_space(str(row[tlif_d_idx]).strip())
            if tlif_s_idx is not None and tlif_s_idx < len(row):
                d["scheduler_tlif"] = _smart_space(str(row[tlif_s_idx]).strip())
            if medical_idx is not None and medical_idx < len(row):
                d["medical_expiry"] = _smart_space(str(row[medical_idx]).strip())

            # Roster/Schedule/SDP: prefer the detected column; accept only date/range-like, not the name
            if roster_idx is not None and roster_idx < len(row):
                raw_roster = _smart_space(str(row[roster_idx]).strip())
                if raw_roster and re.search(r"[0-9/–-]", raw_roster) and raw_roster.lower() != name.lower():
                    d["roster_schedule"] = raw_roster

            # Fallback: scan the row for the first date/range-like cell that's not the name cell
            if "roster_schedule" not in d:
                for j, cell in enumerate(row):
                    if j == name_idx:
                        continue
                    s = _smart_space(str(cell).strip())
                if s and re.search(r"[0-9/–-]", s) and s.lower() != name.lower():
                    d["roster_schedule"] = s
                    break

            if fit_idx is not None and fit_idx < len(row):
                d["fit_for_duty"] = _smart_space(str(row[fit_idx]).strip())
            if diary_idx is not None and diary_idx < len(row):
                d["work_diary"] = _smart_space(str(row[diary_idx]).strip())

            drivers.append(d)

        if drivers:
            extracted["drivers_detailed"] = drivers
            self.log_debug(f"Driver rows extracted (header-based): {len(drivers)}")


    def _extract_management_table(self, data_rows: List[List], extracted: Dict, headers: List[str]):
        txt = " ".join(str(h) for h in headers).lower()
        comp = {}
        for row in data_rows:
            if len(row) < 2: continue
            std = str(row[0]).strip()
            val = _smart_space(str(row[1]).strip())
            if std.startswith("Std") and val:
                comp[std] = val
        if comp:
            if "maintenance" in txt: extracted["maintenance_compliance"] = comp
            elif "mass" in txt: extracted["mass_compliance"] = comp
            elif "fatigue" in txt: extracted["fatigue_compliance"] = comp

    def _extract_text_content(self, text_pages: List[Dict], extracted: Dict):
        all_text = " ".join(page.get("text", "") for page in text_pages)
        all_text = _smart_space(all_text)

        # business summary
        patt = [
            r"Nature of the Operators? Business.*?:\s*(.*?)(?:Accreditation Number|Expiry Date|$)",
            r"Nature of.*?Business.*?Summary.*?:\s*(.*?)(?:Accreditation|$)"
        ]
        for p in patt:
            m = re.search(p, all_text, re.IGNORECASE | re.DOTALL)
            if m:
                txt = re.sub(r'\s+', ' ', m.group(1).strip())
                txt = re.sub(r'\s*(Accreditation Number.*|Expiry Date.*)', '', txt, flags=re.IGNORECASE)
                if len(txt) > 50:
                    extracted["business_summary"] = txt
                    break

        # audit conducted date
        for p in [
            r"Audit was conducted on\s+([0-9]+(?:st|nd|rd|th)?\s+[A-Za-z]+\s+\d{4})",
            r"DATE\s+([0-9]+(?:st|nd|rd|th)?\s+[A-Za-z]+\s+\d{4})",
            r"AUDITOR SIGNATURE\s+DATE\s+([0-9]+(?:st|nd|rd|th)?\s+[A-Za-z]+\s+\d{4})"
        ]:
            m = re.search(p, all_text, re.IGNORECASE)
            if m:
                extracted["audit_conducted_date"] = _smart_space(m.group(1).strip())
                break

        # print accreditation name
        for p in [
            r"\(print accreditation name\)\s*([A-Za-z0-9\s&().,'/\-]+?)(?:\s+DOES|\s+does|\n|$)",
            r"print accreditation name.*?\n\s*([A-Za-z0-9\s&().,'/\-]+?)(?:\s+DOES|\s+does|\n|$)"
        ]:
            m = re.search(p, all_text, re.IGNORECASE)
            if m:
                extracted["print_accreditation_name"] = _smart_space(m.group(1).strip())
                break

        # numbers in text (optional)
        for p in [
            r"Number of powered vehicles\s+(\d+)",
            r"powered vehicles\s+(\d+)",
            r"Number of trailing vehicles\s+(\d+)",
            r"trailing vehicles\s+(\d+)",
            r"Number of drivers in BFM\s+(\d+)",
            r"drivers in BFM\s+(\d+)"
        ]:
            m = re.search(p, all_text, re.IGNORECASE)
            if m:
                val = m.group(1)
                if "powered" in p: extracted.setdefault("vehicle_summary", {})["powered_vehicles"] = val
                elif "trailing" in p: extracted.setdefault("vehicle_summary", {})["trailing_vehicles"] = val
                elif "bfm" in p.lower(): extracted.setdefault("vehicle_summary", {})["drivers_bfm"] = val

    def _extract_detailed_management_data(self, extracted_data: Dict, extracted: Dict):
        all_tables = extracted_data.get("all_tables", [])
        for table in all_tables:
            headers = table.get("headers", [])
            data_rows = table.get("data", [])
            page_num = table.get("page", 0)
            if self._has_details_column(headers):
                section = self._identify_management_section(headers)
                if section:
                    self._extract_management_details(data_rows, extracted, section)
            elif 6 <= page_num <= 15:
                self._extract_summary_by_content(data_rows, headers, extracted, page_num)

    def _extract_summary_by_content(self, data_rows: List[List], headers: List[str], extracted: Dict, page_num: int):
        section_type = "maintenance" if 6 <= page_num <= 9 else "mass" if 10 <= page_num <= 12 else "fatigue" if 13 <= page_num <= 15 else None
        if not section_type: return
        details_key = f"{section_type}_summary_details"
        extracted[details_key] = {}
        for row in data_rows:
            if len(row) < 2: continue
            standard = str(row[0]).strip()
            details = _smart_space(str(row[1]).strip())
            if standard.startswith("Std") and details and len(details) > 10:
                m = re.search(r"Std\s+(\d+)\.\s*([^(]+)", standard)
                if m:
                    key = f"Std {m.group(1)}. {m.group(2).strip()}"
                    extracted[details_key][key] = details

    def _has_details_column(self, headers: List[str]) -> bool:
        return "details" in " ".join(str(h) for h in headers).lower()

    def _identify_management_section(self, headers: List[str]) -> Optional[str]:
        txt = " ".join(str(h) for h in headers).lower()
        if "maintenance" in txt: return "maintenance"
        if "mass" in txt: return "mass"
        if "fatigue" in txt: return "fatigue"
        return None

    def _extract_management_details(self, data_rows: List[List], extracted: Dict, section: str):
        details_key = f"{section}_details"
        extracted[details_key] = {}
        for row in data_rows:
            if len(row) < 2: continue
            standard = str(row[0]).strip()
            details = _smart_space(str(row[1]).strip())
            if standard.startswith("Std") and details and details != "V" and len(details) > 10:
                m = re.search(r"Std\s+\d+\.\s*([^(]+)", standard)
                if m:
                    extracted[details_key][m.group(1).strip()] = details

    def _extract_vehicle_driver_data(self, extracted_data: Dict, extracted: Dict):
        vehicle_regs = extracted_data.get("vehicle_registrations", [])
        if vehicle_regs:
            extracted["vehicle_registrations"] = vehicle_regs
        driver_records = extracted_data.get("driver_records", [])
        if driver_records:
            extracted["driver_records"] = driver_records

    # Add this method inside your NHVASMerger class, with proper indentation
    # Place it after the _extract_vehicle_driver_data method

    def map_vehicle_registration_arrays(self, pdf_extracted: Dict, merged: Dict):
        """Extract and map vehicle registration data (Maintenance + Mass) to DOCX arrays."""
        vehicles_src = []

        # Prefer rows we parsed ourselves (header-based). Fall back to curated list if present.
        if "vehicles" in pdf_extracted and isinstance(pdf_extracted["vehicles"], list):
            vehicles_src = pdf_extracted["vehicles"]
        elif "vehicle_registrations" in pdf_extracted and isinstance(pdf_extracted["vehicle_registrations"], list):
            # Normalize curated structure (list of dicts with keys like 'registration_number', etc.)
            for row in pdf_extracted["vehicle_registrations"]:
                if not isinstance(row, dict):
                    continue
                v = {
                "registration": _smart_space(row.get("registration_number") or row.get("registration") or ""),
                # Maintenance table columns (names as seen in curated JSON)
                "roadworthiness": _smart_space(row.get("roadworthiness_certificates", "")),
                "maintenance_records": _smart_space(row.get("maintenance_records", "")),
                "daily_checks": _smart_space(row.get("daily_checks", "")),
                "fault_recording": _smart_space(row.get("fault_recording_reporting", "")),
                "fault_repair": _smart_space(row.get("fault_repair", "")),
                # Mass table columns (in case the curated list ever includes them)
                "sub_contractor": _smart_space(row.get("sub_contractor", "")),
                "sub_comp": _smart_space(row.get("sub_contracted_vehicles_statement_of_compliance", "")),
                "weight_verification": _smart_space(row.get("weight_verification_records", "")),
                "rfs_certification": _smart_space(row.get("rfs_suspension_certification", row.get("rfs_suspension_certification_#", ""))),
                "suspension_maintenance": _smart_space(row.get("suspension_system_maintenance", "")),
                "trip_records": _smart_space(row.get("trip_records", "")),
                "fault_reporting_suspension": _smart_space(row.get("fault_recording_reporting_on_suspension_system", "")),
            }
            if v["registration"]:
                vehicles_src.append(v)

        if not vehicles_src:
            return  # nothing to map

        # Build column arrays
        regs = []
        roadworthiness = []
        maint_records = []
        daily_checks = []
        fault_recording = []
        fault_repair = []

        sub_contractors = []
        weight_verification = []
        rfs_certification = []
        suspension_maintenance = []
        trip_records = []
        fault_reporting_suspension = []

        for v in vehicles_src:
            reg = _smart_space(v.get("registration", "")).strip()
            if not reg:
                continue
            regs.append(reg)

        roadworthiness.append(_smart_space(v.get("roadworthiness", "")).strip())
        maint_records.append(_smart_space(v.get("maintenance_records", "")).strip())
        daily_checks.append(_smart_space(v.get("daily_checks", "")).strip())
        fault_recording.append(_smart_space(v.get("fault_recording", "")).strip())
        fault_repair.append(_smart_space(v.get("fault_repair", "")).strip())

        sub_contractors.append(_smart_space(v.get("sub_contractor", "")).strip())
        weight_verification.append(_smart_space(v.get("weight_verification", "")).strip())
        rfs_certification.append(_smart_space(v.get("rfs_certification", "")).strip())
        suspension_maintenance.append(_smart_space(v.get("suspension_maintenance", "")).strip())
        trip_records.append(_smart_space(v.get("trip_records", "")).strip())
        fault_reporting_suspension.append(_smart_space(v.get("fault_reporting_suspension", "")).strip())

        # Update Maintenance table arrays (if present in template)
        if "Vehicle Registration Numbers Maintenance" in merged and regs:
            m = merged["Vehicle Registration Numbers Maintenance"]
            m["Registration Number"] = regs
            m["Roadworthiness Certificates"] = roadworthiness
            m["Maintenance Records"] = maint_records
            m["Daily Checks"] = daily_checks
            m["Fault Recording/ Reporting"] = fault_recording
            m["Fault Repair"] = fault_repair

        # Update Mass table arrays (if present in template)
        if "Vehicle Registration Numbers Mass" in merged and regs:
            ms = merged["Vehicle Registration Numbers Mass"]
            ms["Registration Number"] = regs
            ms["Sub contractor"] = sub_contractors
            ms["Weight Verification Records"] = weight_verification
            ms["RFS Suspension Certification #"] = rfs_certification
            ms["Suspension System Maintenance"] = suspension_maintenance
            ms["Trip Records"] = trip_records
            ms["Fault Recording/ Reporting on Suspension System"] = fault_reporting_suspension

        self.log_debug(f"Updated vehicle registration arrays for {len(regs)} vehicles")
    # ───────────────────────────── map to DOCX (apply spacing + safe fallbacks) ─────────────────────────────
    def map_to_docx_structure(self, pdf_extracted: Dict, docx_data: Dict, pdf_data: Dict) -> Dict:
        merged = json.loads(json.dumps(docx_data))

        # Audit Information
        if "audit_info" in pdf_extracted and "Audit Information" in merged:
            ai = pdf_extracted["audit_info"]
            if ai.get("date_of_audit"):
                merged["Audit Information"]["Date of Audit"] = [_smart_space(ai["date_of_audit"])]
            if ai.get("location"):
                merged["Audit Information"]["Location of audit"] = [_smart_space(ai["location"])]
            if ai.get("auditor_name"):
                merged["Audit Information"]["Auditor name"] = [_smart_space(ai["auditor_name"])]
            if ai.get("matrix_id"):
                merged["Audit Information"]["Audit Matrix Identifier (Name or Number)"] = [_smart_space(ai["matrix_id"])]

        # Operator Information
        if "operator_info" in pdf_extracted and "Operator Information" in merged:
            op = pdf_extracted["operator_info"]
            if op.get("name") and _looks_like_company(op["name"]):
                merged["Operator Information"]["Operator name (Legal entity)"] = [_smart_space(op["name"])]
            if op.get("trading_name"):
                merged["Operator Information"]["Registered trading name/s"] = [_smart_space(op["trading_name"])]
            if op.get("acn"):
                merged["Operator Information"]["Australian Company Number"] = [_smart_space(op["acn"])]
            if op.get("manual"):
                merged["Operator Information"]["NHVAS Manual (Policies and Procedures) developed by"] = [_smart_space(op["manual"])]

        # Contact details
        if "operator_info" in pdf_extracted and "Operator contact details" in merged:
            op = pdf_extracted["operator_info"]
            if op.get("business_address"):
                merged["Operator contact details"]["Operator business address"] = [_smart_space(op["business_address"])]
            if op.get("postal_address"):
                merged["Operator contact details"]["Operator Postal address"] = [_smart_space(op["postal_address"])]
            if op.get("email"):
                merged["Operator contact details"]["Email address"] = [op["email"]]
            if op.get("phone"):
                merged["Operator contact details"]["Operator Telephone Number"] = [_smart_space(op["phone"])]

        # Attendance
        if "attendance" in pdf_extracted and "Attendance List (Names and Position Titles)" in merged:
            merged["Attendance List (Names and Position Titles)"]["Attendance List (Names and Position Titles)"] = _clean_list(pdf_extracted["attendance"])

        # Business summary
        if "business_summary" in pdf_extracted and "Nature of the Operators Business (Summary)" in merged:
            merged["Nature of the Operators Business (Summary)"]["Nature of the Operators Business (Summary):"] = [_smart_space(pdf_extracted["business_summary"])]

        # Vehicle summary
        if "vehicle_summary" in pdf_extracted:
            vs = pdf_extracted["vehicle_summary"]
            if "Accreditation Vehicle Summary" in merged:
                if vs.get("powered_vehicles"):
                    merged["Accreditation Vehicle Summary"]["Number of powered vehicles"] = [vs["powered_vehicles"]]
                if vs.get("trailing_vehicles"):
                    merged["Accreditation Vehicle Summary"]["Number of trailing vehicles"] = [vs["trailing_vehicles"]]
            if "Accreditation Driver Summary" in merged:
                if vs.get("drivers_bfm"):
                    merged["Accreditation Driver Summary"]["Number of drivers in BFM"] = [vs["drivers_bfm"]]
                if vs.get("drivers_afm"):
                    merged["Accreditation Driver Summary"]["Number of drivers in AFM"] = [vs["drivers_afm"]]

        # Summary sections (unchanged behavior)
        summary_maps = self.build_summary_maps(pdf_data)
        for section_name, std_map in summary_maps.items():
            if section_name in merged and std_map:
                for detail_key, details_list in std_map.items():
                    if detail_key in merged[section_name]:
                        merged[section_name][detail_key] = details_list
                        continue
                    for docx_key in list(merged[section_name].keys()):
                        m1 = re.search(r"Std\s+(\d+)", detail_key)
                        m2 = re.search(r"Std\s+(\d+)", docx_key)
                        if m1 and m2 and m1.group(1) == m2.group(1):
                            merged[section_name][docx_key] = details_list
                            break

        # Vehicle registration arrays via consolidated builder
        sections = build_vehicle_sections(pdf_extracted)
        if "Vehicle Registration Numbers Maintenance" in merged:
            merged["Vehicle Registration Numbers Maintenance"].update(
                sections["Vehicle Registration Numbers Maintenance"]
            )
        if "Vehicle Registration Numbers Mass" in merged:
            merged["Vehicle Registration Numbers Mass"].update(
                sections["Vehicle Registration Numbers Mass"]
            )


        # replace the whole Drivers/Scheduler block with:
        if "drivers_detailed" in pdf_extracted and "Driver / Scheduler Records Examined" in merged:
            drivers = pdf_extracted["drivers_detailed"]

            def _looks_like_range(s):
                return bool(re.search(r"[0-9]{1,2}[/-]", s or ""))

            merged["Driver / Scheduler Records Examined"]["Roster / Schedule / Safe Driving Plan (Date Range)"] = [d.get("roster_schedule","") for d in drivers]
            merged["Driver / Scheduler Records Examined"]["Fit for Duty Statement Completed (Yes/No)"]          = [d.get("fit_for_duty","") for d in drivers]
            merged["Driver / Scheduler Records Examined"]["Work Diary Pages (Page Numbers) Electronic Work Diary Records (Date Range)"] = [d.get("work_diary","") for d in drivers]


        # --- Print accreditation name (robust, no UnboundLocalError) ---
        if "Print accreditation name" in merged:
            acc_name = ""  # init
            acc_name = _smart_space(pdf_extracted.get("print_accreditation_name") or "")
            if not acc_name:
                oi = pdf_extracted.get("operator_info") or {}
                acc_name = _smart_space(oi.get("name") or "") or _smart_space(oi.get("trading_name") or "")
            if acc_name:
                merged["Print accreditation name"]["(print accreditation name)"] = [acc_name]

        # Audit Declaration dates: prefer explicit extracted date; fallback to audit_info; ignore literal "Date"
        if "Audit Declaration dates" in merged:
            def _real_date(s: Optional[str]) -> bool:
                return bool(s and re.search(r"\d", s) and not re.fullmatch(r"date", s.strip(), re.I))

            val = pdf_extracted.get("audit_conducted_date")
            if not _real_date(val):
                val = (pdf_extracted.get("audit_info", {}) or {}).get("date_of_audit")

            if _real_date(val):
                merged["Audit Declaration dates"]["Audit was conducted on"] = [_smart_space(val)]


        # Operator Declaration: page 22 image missing → derive from first Attendance "Name - Title"
        if "Operator Declaration" in merged:
            # If an explicit operator declaration exists, use it
            if "operator_declaration" in pdf_extracted:
                od = pdf_extracted["operator_declaration"]
                pn = _smart_space(od.get("print_name", ""))
                pt = _smart_space(od.get("position_title", ""))
                if pn:
                    merged["Operator Declaration"]["Print Name"] = [pn]
                if pt:
                    merged["Operator Declaration"]["Position Title"] = [pt]
            else:
                # Fallback: first "Name - Title" from Attendance
                nt = self._first_attendance_name_title(pdf_extracted.get("attendance", []))
                if nt:
                    merged["Operator Declaration"]["Print Name"] = [nt[0]]
                    merged["Operator Declaration"]["Position Title"] = [nt[1]]


        # Paragraphs: fill company name for the 3 management headings; set the 2 dates
        if "paragraphs" in merged:
            paras = merged["paragraphs"]

            audit_date = (
                pdf_extracted.get("audit_conducted_date")
                or pdf_extracted.get("audit_info", {}).get("date_of_audit")
            )

            # Prefer accreditation name, else operator legal name, else trading name
            company_name = (
                _smart_space(pdf_extracted.get("print_accreditation_name") or "")
                or _smart_space(pdf_extracted.get("operator_info", {}).get("name") or "")
                or _smart_space(pdf_extracted.get("operator_info", {}).get("trading_name") or "")
            )

            # Update the three layered headings
            for key in ("MAINTENANCE MANAGEMENT", "MASS MANAGEMENT", "FATIGUE MANAGEMENT"):
                if key in paras and company_name:
                    paras[key] = [company_name]

            # Second-last page: date under page heading
            if "NHVAS APPROVED AUDITOR DECLARATION" in paras and audit_date:
                paras["NHVAS APPROVED AUDITOR DECLARATION"] = [_smart_space(audit_date)]

            # Last page: date under long acknowledgement paragraph
            ack_key = ("I hereby acknowledge and agree with the findings detailed in this NHVAS Audit Summary Report. "
                    "I have read and understand the conditions applicable to the Scheme, including the NHVAS Business Rules and Standards.")
            if ack_key in paras and audit_date:
                paras[ack_key] = [_smart_space(audit_date)]

        self._force_fill_maintenance_from_tables(pdf_data, merged)
        return merged

    # ───────────────────────────── merge & CLI (unchanged) ─────────────────────────────
    def merge_pdf_to_docx(self, docx_data: Dict, pdf_data: Dict) -> Dict:
        self.log_debug("Starting comprehensive PDF extraction...")
        pdf_extracted = self.extract_from_pdf_comprehensive(pdf_data)
        self.log_debug(f"Extracted PDF data keys: {list(pdf_extracted.keys())}")

        self.log_debug("Mapping to DOCX structure...")
        merged_data = self.map_to_docx_structure(pdf_extracted, docx_data, pdf_data)

        for section_name, section_data in docx_data.items():
            if isinstance(section_data, dict):
                for label in section_data:
                    if (section_name in merged_data and 
                        label in merged_data[section_name] and 
                        merged_data[section_name][label] != docx_data[section_name][label]):
                        print(f"✓ Updated {section_name}.{label}: {merged_data[section_name][label]}")
        return merged_data

    def process_files(self, docx_file: str, pdf_file: str, output_file: str):
        try:
            print(f"Loading DOCX JSON from: {docx_file}")
            with open(docx_file, 'r', encoding='utf-8') as f:
                docx_data = json.load(f)
            print(f"Loading PDF JSON from: {pdf_file}")
            with open(pdf_file, 'r', encoding='utf-8') as f:
                pdf_data = json.load(f)

            print("Merging PDF data into DOCX structure...")
            merged_data = self.merge_pdf_to_docx(docx_data, pdf_data)

            print(f"Saving merged data to: {output_file}")
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(merged_data, f, indent=2, ensure_ascii=False)

            print("✅ Merge completed successfully!")
            return merged_data
        except Exception as e:
            print(f"❌ Error processing files: {str(e)}")
            import traceback
            traceback.print_exc()
            raise

def main():
    if len(sys.argv) != 4:
        print("Usage: python nhvas_merger.py <docx_json_file> <pdf_json_file> <output_file>")
        print("Example: python nhvas_merger.py docx_template.json pdf_extracted.json merged_output.json")
        sys.exit(1)

    docx_file = sys.argv[1]
    pdf_file = sys.argv[2]
    output_file = sys.argv[3]

    for file_path in [docx_file, pdf_file]:
        if not Path(file_path).exists():
            print(f"❌ File not found: {file_path}")
            sys.exit(1)

    merger = NHVASMerger()
    merger.process_files(docx_file, pdf_file, output_file)

if __name__ == "__main__":
    main()