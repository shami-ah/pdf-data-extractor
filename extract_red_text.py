#!/usr/bin/env python3
import re, json, sys
from docx import Document
from docx.oxml.ns import qn

def is_red_font(run) -> bool:
    """Return True if this run is coloured red-ish."""
    col = run.font.color
    if col and col.rgb:
        r,g,b = col.rgb[0], col.rgb[1], col.rgb[2]
        if r>150 and g<100 and b<100 and (r-g)>30 and (r-b)>30:
            return True
    # fallback: raw <w:color w:val="XXXXXX"/>
    rPr = getattr(run._element, "rPr", None)
    if rPr is not None:
        clr = rPr.find(qn('w:color'))
        if clr is not None:
            val = clr.get(qn('w:val'))
            if re.fullmatch(r"[0-9A-Fa-f]{6}", val):
                rr,gg,bb = int(val[:2],16), int(val[2:4],16), int(val[4:],16)
                if rr>150 and gg<100 and bb<100 and (rr-gg)>30 and (rr-bb)>30:
                    return True
    return False

# ─────────────────────────────────────────────────────────────────────────────
# Your template, mapped 1:1 to doc.tables[0..18]
MASTER_TABLES = [
  # Table 0: Tick as appropriate (Mass, Maintenance, etc.)
  {
    "name": "Tick as appropriate",
    "labels_on_row1": True,
    "labels": [
      "Mass", "Maintenance", "Basic Fatigue", "Advanced Fatigue",
      "Entry Audit", "Initial Compliance Audit", "Compliance Audit",
      "Spot Check", "Triggered Audit"
    ]
  },

  # Table 1: Audit Information
  {
    "name": "Audit Information",
    "labels_on_left": True,
    "labels": [
      "Date of Audit",
      "Location of audit",
      "Auditor name",
      "Audit Matrix Identifier (Name or Number)",  # Corrected full label
      "Auditor Exemplar Global Reg No.",
      "expiry Date:",
      "NHVR Auditor Registration Number",
      "expiry Date:"  # Note: Duplicate label, might need special handling
    ]
  },

  # Table 2: Operator Information (including contact details)
  {
    "name": "Operator Information",
    "labels_on_left": True,
    "skip_rows": ["Operator contact details", ""],  # Skip subheading and blank rows
    "labels": [
      "Operator name (Legal entity)",
      "NHVAS Accreditation No. (If applicable)",
      "Registered trading name/s",
      "Australian Company Number",
      "NHVAS Manual (Policies and Procedures) developed by",
      "Operator business address",
      "Operator Postal address",
      "Email address",
      "Operator Telephone Number"
    ]
  },

  # Table 3: Attendance List
  {
    "name": "Attendance List (Names and Position Titles)",
    "labels": ["Attendance List (Names and Position Titles)"]
  },

  # Table 4: Nature of the Operators Business
  {
    "name": "Nature of the Operators Business (Summary)",
    "labels": [
      "Nature of the Operators Business (Summary)",
      "Accreditation Number:",
      "Expiry Date:"
    ]
  },

  # Table 5: Accreditation Vehicle Summary
  {
    "name": "Accreditation Vehicle Summary",
    "labels_on_left": True,
    "labels": [
      "Number of powered vehicles",
      "Number of trailing vehicles"
    ]
  },

  # Table 6: Accreditation Driver Summary
  {
    "name": "Accreditation Driver Summary",
    "labels_on_left": True,
    "labels": [
      "Number of drivers in BFM",
      "Number of drivers in AFM"
    ]
  },

  # Table 7: Compliance Codes
  {
    "name": "Compliance Codes",
    "labels_on_row1": True,
    "labels": ["V", "SFI", "NA", "NC", "NAP"]
  },

  # Table 8: Corrective Action Request Identification
  {
    "name": "Corrective Action Request Identification",
    "labels_on_row1": True,
    "labels": ["Title", "Abbreviation", "Description"]
  },

  # Table 9: MASS MANAGEMENT (Standards 1-8)
  {
    "name": "MASS MANAGEMENT",
    "labels_on_left": True,
    "labels": [
      "Std 1. Responsibilities",
      "Std 2. Vehicle Control",
      "Std 3. Vehicle Use",
      "Std 4. Records and Documentation",
      "Std 5. Verification",
      "Std 6. Internal Review",
      "Std 7. Training and Education",
      "Std 8. Maintenance of Suspension"
    ]
  },

  # Table 10: Mass Management Summary of Audit findings (Standards 1-8)
  {
    "name": "Mass Management Summary of Audit findings",
    "labels_on_left": True,
    "labels": [
      "Std 1. Responsibilities",
      "Std 2. Vehicle Control",
      "Std 3. Vehicle Use",
      "Std 4. Records and Documentation",
      "Std 5. Verification",
      "Std 6. Internal Review",
      "Std 7. Training and Education",
      "Std 8. Maintenance of Suspension"
    ]
  },

  # Table 11: Vehicle Registration Numbers of Records Examined
  {
    "name": "Vehicle Registration Numbers of Records Examined",
    "labels_on_row1": True,
    "labels": [
      "No.", "Registration Number",
      "Sub-contractor (Yes/No)",
      "Sub-contracted Vehicles Statement of Compliance (Yes/No)",
      "Weight Verification Records (Date Range)",
      "RFS Suspension Certification # (N/A if not applicable)",
      "Suspension System Maintenance (Date Range)",
      "Trip Records (Date Range)",
      "Fault Recording/ Reporting on Suspension System (Date Range)"
    ]
  },

  # Table 12: Operator's Name (legal entity) - Signature block
  {
    "name": "Operator’s Name (legal entity)",
    "labels": ["Operator’s Name (legal entity)"]
  },

  # Table 13: Non-conformance type
  {
    "name": "Non-conformance type (please tick)",
    "labels": ["Un-conditional", "Conditional"]
  },

  # Table 14: Non-conformance Information
  {
    "name": "Non-conformance Information",
    "labels_on_row1": True,
    "labels": [
      "Non-conformance agreed close out date",
      "Module and Standard",
      "Corrective Action Request (CAR) Number"
    ]
  },

  # Table 15: Non-conformance and action taken
  {
    "name": "Non-conformance and action taken",
    "labels_on_row1": True,
    "labels": [
      "Observed Non-conformance:",
      "Corrective Action taken or to be taken by operator:",
      "Operator or Representative Signature", "Position", "Date"
    ]
  },

  # Table 16: Print Name / Auditor Reg Number
  {
    "name": "Print Name / Auditor Reg Number",
    "labels_on_row1": True,
    "labels": [
      "Print Name",
      "NHVR or Exemplar Global Auditor Registration Number"
    ]
  },

  # Table 17: Audit Declaration
  {
    "name": "Audit Declaration",
    "labels_on_left": True,
    "labels": [
      "Audit was conducted on",
      "Unconditional CARs closed out on:",
      "Conditional CARs to be closed out by:"
    ]
  },

  # Table 18: print accreditation name
  {
    "name": "print accreditation name",
    "labels": ["print accreditation name"]
  },

  # Table 19: Operator Declaration
  {
    "name": "Operator Declaration",
    "labels_on_row1": True,
    "labels": ["Print Name", "Position Title"]
  }
]

def extract_red_text(path):
    doc = Document(path)

    # debug print
    print(f"Found {len(doc.tables)} tables:")
    for i,t in enumerate(doc.tables):
        print(f"  Table#{i}: “{t.rows[0].cells[0].text.strip()[:30]}…”")
    print()

    out = {}
    for ti, spec in enumerate(MASTER_TABLES):
        if ti >= len(doc.tables):
            break
        tbl = doc.tables[ti]
        name = spec["name"]

        # prepare container & dedupe sets
        collected = {lbl:[] for lbl in spec["labels"]}
        seen = {lbl:set() for lbl in spec["labels"]}

        # choose orientation
        if spec.get("labels_on_row1"):
            headers   = spec["labels"]
            rows      = tbl.rows[1:]
            col_mode  = True
        elif spec.get("labels_on_left"):
            headers   = spec["labels"]
            # skip any unwanted header/subheading rows
            rows = [
                row for row in tbl.rows[1:]
                if row.cells[0].text.strip() not in spec.get("skip_rows",[])
            ]
            col_mode = False
        else:
            headers   = [name]
            rows      = tbl.rows
            col_mode  = None

        # scan each cell
        for ri,row in enumerate(rows):
            for ci,cell in enumerate(row.cells):
                red = "".join(
                    run.text for p in cell.paragraphs for run in p.runs
                    if is_red_font(run)
                ).strip()
                if not red: continue

                # assign label
                if   col_mode is True:
                    lbl = headers[ci] if ci<len(headers) else name
                elif col_mode is False:
                    lbl = headers[ri] if ri<len(headers) else name
                else:
                    lbl = name

                # dedupe & collect
                if red not in seen[lbl]:
                    seen[lbl].add(red)
                    collected[lbl].append(red)

        # only keep non-empty labels
        filtered = {l:collected[l] for l in collected if collected[l]}
        if filtered:
            out[name] = filtered

    # paragraphs
    paras = {}
    for i,para in enumerate(doc.paragraphs):
        red = "".join(r.text for r in para.runs if is_red_font(r)).strip()
        if not red: continue
        # find nearest non-red above
        lab = None
        for j in range(i-1,-1,-1):
            if any(is_red_font(r) for r in doc.paragraphs[j].runs):
                continue
            txt = doc.paragraphs[j].text.strip()
            if txt:
                lab = txt; break
        key = lab or "(para)"
        paras.setdefault(key,[]).append(red)

    if paras:
        out["paragraphs"] = paras
    return out

if __name__=="__main__":
    fn = sys.argv[1] if len(sys.argv)>1 else "test.docx"
    word_data = extract_red_text(fn)

    # --- STORE TO JSON for later reuse ---
    with open('word_red_data.json', 'w', encoding='utf-8') as f:
        json.dump(word_data, f, indent=2, ensure_ascii=False)
    # ----------------------------------------

    # still print to console for immediate feedback
    print(json.dumps(word_data, indent=2, ensure_ascii=False))
