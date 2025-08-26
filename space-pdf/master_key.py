"""
Improved Master Key for NHVAS Audit extraction:
- TABLE_SCHEMAS: Enhanced definitions with better matching criteria for Summary vs Basic tables
- HEADING_PATTERNS: Improved regex patterns for main/sub headings
- PARAGRAPH_PATTERNS: Enhanced patterns for key narrative sections
"""

# 1. Enhanced table schemas with better matching logic
TABLE_SCHEMAS = {
    "Tick as appropriate": {
        "headings": [
            {"level": 1, "text": "NHVAS Audit Summary Report"},
        ],
        "orientation": "left",
        "labels": [
            "Mass",
            "Entry Audit", 
            "Maintenance",
            "Initial Compliance Audit",
            "Basic Fatigue",
            "Compliance Audit",
            "Advanced Fatigue",
            "Spot Check",
            "Triggered Audit"
        ],
        "priority": 90  # High priority for direct match
    },
    "Audit Information": {
        "orientation": "left",
        "labels": [
            "Date of Audit",
            "Location of audit", 
            "Auditor name",
            "Audit Matrix Identifier (Name or Number)",
            "Auditor Exemplar Global Reg No.",
            "expiry Date:",
            "NHVR Auditor Registration Number",
            "expiry Date:"
        ],
        "priority": 80
    },
    "Operator Information": {
        "headings": [
            {"level": 1, "text": "Operator Information"}
        ],
        "orientation": "left",
        "labels": [
            "Operator name (Legal entity)",
            "NHVAS Accreditation No. (If applicable)",
            "Registered trading name/s", 
            "Australian Company Number",
            "NHVAS Manual (Policies and Procedures) developed by"
        ],
        "priority": 85
    },
    "Operator contact details": {
        "orientation": "left",
        "labels": [
            "Operator business address",
            "Operator Postal address",
            "Email address", 
            "Operator Telephone Number"
        ],
        "priority": 75,
        "context_keywords": ["contact", "address", "email", "telephone"]
    },
    "Attendance List (Names and Position Titles)": {
        "headings": [
            {"level": 1, "text": "NHVAS Audit Summary Report"}
        ],
        "orientation": "row1",
        "labels": ["Attendance List (Names and Position Titles)"],
        "priority": 90
    },
    "Nature of the Operators Business (Summary)": {
        "orientation": "row1",
        "labels": ["Nature of the Operators Business (Summary):"],
        "split_labels": ["Accreditation Number:", "Expiry Date:"],
        "priority": 85
    },
    "Accreditation Vehicle Summary": {
        "orientation": "left",
        "labels": ["Number of powered vehicles", "Number of trailing vehicles"],
        "priority": 80
    },
    "Accreditation Driver Summary": {
        "orientation": "left", 
        "labels": ["Number of drivers in BFM", "Number of drivers in AFM"],
        "priority": 80
    },
    "Compliance Codes": {
        "orientation": "left",
        "labels": ["V", "NC", "TNC", "SFI", "NAP", "NA"],
        "priority": 70,
        "context_exclusions": ["MASS MANAGEMENT", "MAINTENANCE MANAGEMENT", "FATIGUE MANAGEMENT"]
    },
    "Corrective Action Request Identification": {
        "orientation": "row1",
        "labels": ["Title", "Abbreviation", "Description"],
        "priority": 80
    },
    
    # 🎯 BASIC MANAGEMENT SCHEMAS (Compliance Tables - Lower Priority)
    "Maintenance Management": {
        "headings": [
            {"level": 1, "text": "NHVAS AUDIT SUMMARY REPORT"}
        ],
        "orientation": "left",
        "labels": [
            "Std 1. Daily Check",
            "Std 2. Fault Recording and Reporting", 
            "Std 3. Fault Repair",
            "Std 4. Maintenance Schedules and Methods",
            "Std 5. Records and Documentation",
            "Std 6. Responsibilities",
            "Std 7. Internal Review",
            "Std 8. Training and Education"
        ],
        "priority": 60,
        "context_keywords": ["maintenance"],
        "context_exclusions": ["summary", "details", "audit findings"]  # Exclude Summary tables
    },
    "Mass Management": {
        "headings": [
            {"level": 1, "text": "NHVAS AUDIT SUMMARY REPORT"}
        ],
        "orientation": "left",
        "labels": [
            "Std 1. Responsibilities",
            "Std 2. Vehicle Control",
            "Std 3. Vehicle Use", 
            "Std 4. Records and Documentation",
            "Std 5. Verification",
            "Std 6. Internal Review",
            "Std 7. Training and Education",
            "Std 8. Maintenance of Suspension"
        ],
        "priority": 60,
        "context_keywords": ["mass"],
        "context_exclusions": ["summary", "details", "audit findings"]  # Exclude Summary tables
    },
    "Fatigue Management": {
        "headings": [
            {"level": 1, "text": "NHVAS AUDIT SUMMARY REPORT"}
        ],
        "orientation": "left",
        "labels": [
            "Std 1. Scheduling and Rostering",
            "Std 2. Health and wellbeing for performed duty",
            "Std 3. Training and Education",
            "Std 4. Responsibilities and management practices", 
            "Std 5. Internal Review",
            "Std 6. Records and Documentation",
            "Std 7. Workplace conditions"
        ],
        "priority": 60,
        "context_keywords": ["fatigue"],
        "context_exclusions": ["summary", "details", "audit findings"]  # Exclude Summary tables
    },
    
    # 🎯 SUMMARY MANAGEMENT SCHEMAS (Detailed Tables with DETAILS column - Higher Priority)
    "Maintenance Management Summary": {
        "headings": [
            {"level": 1, "text": "Audit Observations and Comments"},
            {"level": 2, "text": "Maintenance Management Summary of Audit findings"}
        ],
        "orientation": "left",
        "columns": ["MAINTENANCE MANAGEMENT", "DETAILS"],
        "labels": [
            "Std 1. Daily Check", 
            "Std 2. Fault Recording and Reporting",
            "Std 3. Fault Repair", 
            "Std 4. Maintenance Schedules and Methods",
            "Std 5. Records and Documentation", 
            "Std 6. Responsibilities",
            "Std 7. Internal Review", 
            "Std 8. Training and Education"
        ],
        "priority": 85,  # Higher priority than basic Maintenance Management
        "context_keywords": ["maintenance", "summary", "details", "audit findings"]
    },
    "Mass Management Summary": {
        "headings": [
            {"level": 1, "text": "Mass Management Summary of Audit findings"}
        ],
        "orientation": "left",
        "columns": ["MASS MANAGEMENT", "DETAILS"],
        "labels": [
            "Std 1. Responsibilities",
            "Std 2. Vehicle Control", 
            "Std 3. Vehicle Use",
            "Std 4. Records and Documentation",
            "Std 5. Verification",
            "Std 6. Internal Review",
            "Std 7. Training and Education",
            "Std 8. Maintenance of Suspension"
        ],
        "priority": 85,  # Higher priority than basic Mass Management
        "context_keywords": ["mass", "summary", "details", "audit findings"]
    },
    "Fatigue Management Summary": {
        "headings": [
            {"level": 1, "text": "Fatigue Management Summary of Audit findings"}
        ],
        "orientation": "left",
        "columns": ["FATIGUE MANAGEMENT", "DETAILS"],
        "labels": [
            "Std 1. Scheduling and Rostering",
            "Std 2. Health and wellbeing for performed duty",
            "Std 3. Training and Education",
            "Std 4. Responsibilities and management practices",
            "Std 5. Internal Review", 
            "Std 6. Records and Documentation",
            "Std 7. Workplace conditions"
        ],
        "priority": 85,  # Higher priority than basic Fatigue Management
        "context_keywords": ["fatigue", "summary", "details", "audit findings"]
    },
    
    # Vehicle Registration Tables
    "Vehicle Registration Numbers Mass": {
    "headings": [
        {"level": 1, "text": "Vehicle Registration Numbers of Records Examined"},
        {"level": 2, "text": "MASS MANAGEMENT"}
    ],
    "orientation": "row1", 
    "labels": [
        "No.", "Registration Number", "Sub contractor",
        "Sub-contracted Vehicles Statement of Compliance",
        "Weight Verification Records",
        "RFS Suspension Certification #",
        "Suspension System Maintenance", "Trip Records",
        "Fault Recording/ Reporting on Suspension System"
    ],
    "priority": 90,  # Higher priority
    "context_keywords": ["mass", "vehicle registration", "rfs suspension", "weight verification"],
    "context_exclusions": ["maintenance", "roadworthiness", "daily checks"]  # Exclude maintenance-specific terms
},
"Vehicle Registration Numbers Maintenance": {
    "headings": [
        {"level": 1, "text": "Vehicle Registration Numbers of Records Examined"},
        {"level": 2, "text": "Maintenance Management"}
    ],
    "orientation": "row1",
    "labels": [
        "No.", "Registration Number", "Roadworthiness Certificates",
        "Maintenance Records", "Daily Checks",
        "Fault Recording/ Reporting", "Fault Repair"
    ],
    "priority": 85,  # Lower priority
    "context_keywords": ["maintenance", "vehicle registration", "roadworthiness", "daily checks"],
    "context_exclusions": ["mass", "rfs suspension", "weight verification"]  # Exclude mass-specific terms
},
    "Driver / Scheduler Records Examined": {
        "headings": [
            {"level": 1, "text": "Driver / Scheduler Records Examined"},
            {"level": 2, "text": "FATIGUE MANAGEMENT"},
        ],
        "orientation": "row1",
        "labels": [
            "No.",
            "Driver / Scheduler Name", 
            "Driver TLIF Course # Completed",
            "Scheduler TLIF Course # Completed",
            "Medical Certificates (Current Yes/No) Date of expiry",
            "Roster / Schedule / Safe Driving Plan (Date Range)",
            "Fit for Duty Statement Completed (Yes/No)",
            "Work Diary Pages (Page Numbers) Electronic Work Diary Records (Date Range)"
        ],
        "priority": 80,
        "context_keywords": ["driver", "scheduler", "fatigue"]
    },
    
    # Other Tables
    "Operator's Name (legal entity)": {
        "headings": [
            {"level": 1, "text": "CORRECTIVE ACTION REQUEST (CAR)"}
        ],
        "orientation": "left",
        "labels": ["Operator's Name (legal entity)"],
        "priority": 85
    },
    "Non-conformance and CAR details": {
        "orientation": "left",
        "labels": [
            "Non-conformance agreed close out date",
            "Module and Standard",
            "Corrective Action Request (CAR) Number",
            "Observed Non-conformance:",
            "Corrective Action taken or to be taken by operator:",
            "Operator or Representative Signature",
            "Position",
            "Date",
            "Comments:",
            "Auditor signature", 
            "Date"
        ],
        "priority": 75,
        "context_keywords": ["non-conformance", "corrective action"]
    },
    "NHVAS Approved Auditor Declaration": {
        "headings": [
            {"level": 1, "text": "NHVAS APPROVED AUDITOR DECLARATION"}
        ],
        "orientation": "row1",
        "labels": ["Print Name", "NHVR or Exemplar Global Auditor Registration Number"],
        "priority": 90,
        "context_keywords": ["auditor declaration", "NHVR"],
        "context_exclusions": ["manager", "operator declaration"]
    },
    "Audit Declaration dates": {
        "headings": [
            {"level": 1, "text": "Audit Declaration dates"}
        ],
        "orientation": "left",
        "labels": [
            "Audit was conducted on",
            "Unconditional CARs closed out on:",
            "Conditional CARs to be closed out by:"
        ],
        "priority": 80
    },
    "Print accreditation name": {
        "headings": [
            {"level": 1, "text": "(print accreditation name)"}
        ],
        "orientation": "left", 
        "labels": ["(print accreditation name)"],
        "priority": 85
    },
    "Operator Declaration": {
        "headings": [
            {"level": 1, "text": "Operator Declaration"}
        ],
        "orientation": "row1",
        "labels": ["Print Name", "Position Title"],
        "priority": 90,
        "context_keywords": ["operator declaration", "manager"],
        "context_exclusions": ["auditor", "nhvas approved"]
    }
}

# 2. Enhanced heading detection patterns
HEADING_PATTERNS = {
    "main": [
        r"NHVAS\s+Audit\s+Summary\s+Report",
        r"NATIONAL\s+HEAVY\s+VEHICLE\s+ACCREDITATION\s+AUDIT\s+SUMMARY\s+REPORT",
        r"NHVAS\s+AUDIT\s+SUMMARY\s+REPORT"
    ],
    "sub": [
        r"AUDIT\s+OBSERVATIONS\s+AND\s+COMMENTS",
        r"MAINTENANCE\s+MANAGEMENT", 
        r"MASS\s+MANAGEMENT",
        r"FATIGUE\s+MANAGEMENT",
        r"Fatigue\s+Management\s+Summary\s+of\s+Audit\s+findings",
        r"MAINTENANCE\s+MANAGEMENT\s+SUMMARY\s+OF\s+AUDIT\s+FINDINGS",
        r"MASS\s+MANAGEMENT\s+SUMMARY\s+OF\s+AUDIT\s+FINDINGS",
        r"Vehicle\s+Registration\s+Numbers\s+of\s+Records\s+Examined",
        r"CORRECTIVE\s+ACTION\s+REQUEST\s+\(CAR\)",
        r"NHVAS\s+APPROVED\s+AUDITOR\s+DECLARATION",
        r"Operator\s+Declaration",
        r"Operator\s+Information"
    ]
}

# 3. Enhanced paragraph patterns for key narrative sections
PARAGRAPH_PATTERNS = {
    "findings_summary": r"Provide a summary of findings based on the evidence gathered during the audit\.",
    "declaration_text": r"I hereby acknowledge and agree with the findings.*",
    "introductory_note": r"This audit assesses the.*",
    "date_line": r"^\s*\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+\d{4}\s*$|^Date$"
}
