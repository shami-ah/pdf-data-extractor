#!/usr/bin/env python3
"""
Fixed PDF Data Extractor - Addresses key issues in comprehensive_extract.py
Key fixes:
1. Better table extraction and cleaning
2. Improved key-value pair extraction
3. More robust text processing
4. Enhanced vehicle registration extraction
5. Better date/number pattern recognition
"""

import json
import re
import pandas as pd
from typing import Dict, List, Any, Optional
import logging
from pathlib import Path
import sys
from datetime import datetime

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("fixed_pdf_extractor")

class FixedPDFExtractor:
    def __init__(self):
        logger.info("🚀 Initializing Fixed PDF Extractor")
        
    def extract_everything(self, pdf_path: str) -> Dict[str, Any]:
        if not HAS_PDFPLUMBER:
            raise RuntimeError("pdfplumber is required. Install with: pip install pdfplumber")

        logger.info(f"📖 Processing PDF: {pdf_path}")
        result = {
            "document_info": {
                "filename": Path(pdf_path).name,
                "total_pages": 0,
                "extraction_timestamp": datetime.now().isoformat()
            },
            "extracted_data": {
                "all_text_content": [],
                "all_tables": [],
                "key_value_pairs": {},
                "audit_information": {},
                "operator_information": {},
                "vehicle_registrations": [],
                "driver_records": [],
                "compliance_summary": {},
                "dates_and_numbers": {}
            }
        }

        all_text_blocks, all_tables = [], []

        with pdfplumber.open(pdf_path) as pdf:
            result["document_info"]["total_pages"] = len(pdf.pages)
            
            for page_num, page in enumerate(pdf.pages, 1):
                logger.info(f"📄 Processing page {page_num}")
                
                # Extract text with better handling
                page_text = self._extract_page_text(page)
                if page_text:
                    all_text_blocks.append({
                        "page": page_num, 
                        "text": page_text,
                        "word_count": len(page_text.split())
                    })

                # Extract tables with improved cleaning
                tables = self._extract_page_tables(page, page_num)
                all_tables.extend(tables)

        result["extracted_data"]["all_text_content"] = all_text_blocks
        result["extracted_data"]["all_tables"] = all_tables

        # Process extracted data with improved methods
        combined_text = "\n\n".join(b["text"] for b in all_text_blocks)
        
        result["extracted_data"]["key_value_pairs"] = self._extract_key_value_pairs_improved(combined_text)
        result["extracted_data"]["audit_information"] = self._extract_audit_info(combined_text, all_tables)
        result["extracted_data"]["operator_information"] = self._extract_operator_info(combined_text, all_tables)
        result["extracted_data"]["vehicle_registrations"] = self._extract_vehicle_registrations(all_tables)
        result["extracted_data"]["driver_records"] = self._extract_driver_records(all_tables)
        result["extracted_data"]["compliance_summary"] = self._extract_compliance_summary(combined_text, all_tables)
        result["extracted_data"]["dates_and_numbers"] = self._extract_dates_and_numbers_improved(combined_text)

        # Generate summary
        result["extraction_summary"] = {
            "text_blocks_found": len(all_text_blocks),
            "tables_found": len(all_tables),
            "key_value_pairs_found": len(result["extracted_data"]["key_value_pairs"]),
            "vehicle_registrations_found": len(result["extracted_data"]["vehicle_registrations"]),
            "driver_records_found": len(result["extracted_data"]["driver_records"]),
            "total_characters": len(combined_text),
            "processing_timestamp": datetime.now().isoformat()
        }

        logger.info("✅ Extraction completed!")
        return result

    def _extract_page_text(self, page) -> Optional[str]:
        """Extract text from page with better handling"""
        try:
            text = page.extract_text()
            if text:
                # Clean up text
                text = re.sub(r'[ \t]+', ' ', text.strip())
                text = re.sub(r'\n\s*\n', '\n', text)
                return text
        except Exception as e:
            logger.warning(f"Failed to extract text from page: {e}")
        return None

    def _extract_page_tables(self, page, page_num: int) -> List[Dict]:
        """Extract tables with improved processing"""
        tables = []
        try:
            raw_tables = page.extract_tables()
            if raw_tables:
                for table_idx, table in enumerate(raw_tables):
                    cleaned_table = self._clean_table_improved(table)
                    if cleaned_table and len(cleaned_table) > 0:
                        tables.append({
                            "page": page_num,
                            "table_index": table_idx + 1,
                            "headers": cleaned_table[0] if cleaned_table else [],
                            "data": cleaned_table[1:] if len(cleaned_table) > 1 else [],
                            "raw_data": cleaned_table,
                            "row_count": len(cleaned_table) - 1 if len(cleaned_table) > 1 else 0,
                            "column_count": len(cleaned_table[0]) if cleaned_table else 0
                        })
        except Exception as e:
            logger.warning(f"Failed to extract tables from page {page_num}: {e}")
        
        return tables

    def _clean_table_improved(self, table: List[List]) -> List[List[str]]:
        """Improved table cleaning with better cell processing"""
        if not table:
            return []
        
        cleaned = []
        for row in table:
            cleaned_row = []
            for cell in row:
                if cell is None:
                    cleaned_cell = ""
                else:
                    cleaned_cell = str(cell).strip()
                    cleaned_cell = re.sub(r'\s+', ' ', cleaned_cell)
                    cleaned_cell = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', cleaned_cell)
                cleaned_row.append(cleaned_cell)
            if any(cell.strip() for cell in cleaned_row):
                cleaned.append(cleaned_row)
        
        # Optional: collapse single-column tables of empty strings
        if cleaned and all(len(r) == len(cleaned[0]) for r in cleaned):
            return cleaned
        return cleaned

    def _extract_key_value_pairs_improved(self, text: str) -> Dict[str, str]:
        """Improved key-value pair extraction with better cleaning"""
        pairs: Dict[str, str] = {}

        # Normalize text a bit for regex stability
        t = text.replace('\r', '\n')

        # Pattern 1: colon-separated pairs (key: value)
        pattern1 = re.compile(
            r'([A-Za-z][\w\s()/\-.]{2,80}?):\s*([^\n\r:][^\n\r]*)'
        )
        for key, val in pattern1.findall(t):
            k = key.strip()
            v = val.strip()
            # Filter junk: very long values, pure separators, or obvious headers
            if not v or len(v) > 200:
                continue
            if re.fullmatch(r'[-_/\.]+', v):
                continue
            # Avoid capturing the next key as value by trimming trailing key-like tokens
            v = re.sub(r'\s+[A-Z][\w\s()/\-.]{2,40}:$', '', v).strip()
            # Skip values that are just long digit runs (likely id lists without meaning)
            if re.fullmatch(r'\d{6,}', v):
                continue
            pairs[k] = v

        # Pattern 2: inline “Key – Value” or “Key — Value”
        pattern2 = re.compile(r'([A-Za-z][\w\s()/\-.]{2,80}?)\s*[–—-]\s*([^\n\r]+)')
        for key, val in pattern2.findall(t):
            k = key.strip()
            v = val.strip()
            if v and len(v) <= 200 and not re.fullmatch(r'\d{6,}', v):
                pairs.setdefault(k, v)

        return pairs

    def _extract_audit_info(self, text: str, tables: List[Dict]) -> Dict[str, Any]:
        """Extract audit-specific information with better filtering"""
        audit_info: Dict[str, Any] = {}
        
        # Prefer tables
        for table in tables:
            headers = [str(h).lower() for h in table.get("headers", [])]
            joined = ' '.join(headers)
            if "audit information" in joined or "auditinformation" in joined:
                data = table.get("data", [])
                for row in data:
                    if len(row) >= 2 and row[0] and row[1]:
                        key = str(row[0]).strip()
                        value = str(row[1]).strip()
                        # Skip numbered list rows (e.g., "1.", "2)")
                        if re.match(r'^\s*\d+\s*[.)]\s*$', key):
                            continue
                        if key and value:
                            audit_info[key] = value

        # Backup from text
        candidates = {
            "Date of Audit": r'Date\s+of\s+Audit[:\s]*([^\n\r]+)',
            "Location of audit": r'Location\s+of\s+audit[:\s]*([^\n\r]+)',
            "Auditor name": r'Auditor\s+name[:\s]*([^\n\r]+)',
            "Audit Matrix Identifier (Name or Number)": r'Audit\s+Matrix\s+Identifier.*?[:\s]*([^\n\r]+)',
        }
        for k, pat in candidates.items():
            if k not in audit_info:
                m = re.search(pat, text, re.IGNORECASE)
                if m:
                    audit_info[k] = m.group(1).strip()
        
        return audit_info

    def _extract_operator_info(self, text: str, tables: List[Dict]) -> Dict[str, Any]:
        """Extract operator information with better table parsing"""
        operator_info: Dict[str, Any] = {}
        
        # Look for operator information in tables first
        for table in tables:
            headers = [str(h).lower() for h in table.get("headers", [])]
            if ("operatorinformation" in ' '.join(headers) or 
                "operator information" in ' '.join(headers) or
                "operatorcontactdetails" in ' '.join(headers)):
                
                data = table.get("data", [])
                for row in data:
                    if len(row) >= 2 and row[0] and row[1]:
                        key = str(row[0]).strip()
                        value = str(row[1]).strip()
                        if key and value:
                            # Clean up key names
                            kl = key.lower()
                            if "operator name" in kl:
                                operator_info["operator_name"] = value
                            elif "trading name" in kl:
                                operator_info["trading_name"] = value
                            elif "company number" in kl:
                                if len(row) > 2:
                                    company_parts = [str(r).strip() for r in row[1:] if str(r).strip()]
                                    operator_info["company_number"] = "".join(company_parts)
                                else:
                                    operator_info["company_number"] = value
                            elif "business address" in kl:
                                operator_info["business_address"] = value
                            elif "postal address" in kl:
                                operator_info["postal_address"] = value
                            elif "email" in kl:
                                operator_info["email"] = value
                            elif "telephone" in kl or "phone" in kl:
                                operator_info["phone"] = value
                            elif "nhvas accreditation" in kl:
                                operator_info["nhvas_accreditation"] = value
                            elif "nhvas manual" in kl:
                                operator_info["nhvas_manual"] = value
        
        # Extract from text patterns as backup
        patterns = {
            'operator_name': r'Operator\s*name[:\s\(]*([^\n\r\)]+?)(?=\s*NHVAS|\s*Registered|$)',
            'trading_name': r'Registered\s*trading\s*name[:\s\/]*([^\n\r]+?)(?=\s*Australian|$)', 
            'company_number': r'Australian\s*Company\s*Number[:\s]*([0-9\s]+?)(?=\s*NHVAS|$)',
            'business_address': r'Operator\s*business\s*address[:\s]*([^\n\r]+?)(?=\s*Operator\s*Postal|$)',
            'postal_address': r'Operator\s*Postal\s*address[:\s]*([^\n\r]+?)(?=\s*Email|$)',
            'email': r'Email\s*address[:\s]*([^\s\n\r]+)',
            'phone': r'Operator\s*Telephone\s*Number[:\s]*([^\s\n\r]+)',
            'nhvas_accreditation': r'NHVAS\s*Accreditation\s*No\.[:\s\(]*([^\n\r\)]+)',
        }
        
        for key, pattern in patterns.items():
            if key not in operator_info:  # Only use text if not found in tables
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    value = match.group(1).strip()
                    if value and len(value) < 200:
                        if key == 'company_number':
                            value = re.sub(r'\s+', '', value)
                        operator_info[key] = value
        
        return operator_info

    def _extract_vehicle_registrations(self, tables: List[Dict]) -> List[Dict]:
        """Extract vehicle registration information from tables"""
        vehicles: List[Dict[str, Any]] = []
        
        for table in tables:
            headers = [str(h).lower() for h in table.get("headers", [])]
            
            # Look for vehicle registration tables
            if any(keyword in ' '.join(headers) for keyword in ['registration', 'vehicle', 'number']):
                reg_col = None
                for i, header in enumerate(headers):
                    if 'registration' in header and 'number' in header:
                        reg_col = i
                        break
                
                if reg_col is not None:
                    data = table.get("data", [])
                    for row in data:
                        if len(row) > reg_col and row[reg_col]:
                            reg_num = str(row[reg_col]).strip()
                            # Validate registration format (letters/numbers)
                            if re.match(r'^[A-Z]{1,3}\s*\d{1,3}\s*[A-Z]{0,3}$', reg_num):
                                vehicle_info = {"registration_number": reg_num}
                                
                                # Add other columns as additional info
                                for i, header in enumerate(table.get("headers", [])):
                                    if i < len(row) and i != reg_col:
                                        vehicle_info[str(header)] = str(row[i]).strip()
                                
                                vehicles.append(vehicle_info)
        
        return vehicles

    def _extract_driver_records(self, tables: List[Dict]) -> List[Dict]:
        """Extract driver records from tables"""
        drivers: List[Dict[str, Any]] = []
        
        for table in tables:
            headers = [str(h).lower() for h in table.get("headers", [])]
            
            # Look for driver/scheduler tables
            if any(keyword in ' '.join(headers) for keyword in ['driver', 'scheduler', 'name']):
                name_col = None
                for i, header in enumerate(headers):
                    if 'name' in header:
                        name_col = i
                        break
                
                if name_col is not None:
                    data = table.get("data", [])
                    for row in data:
                        if len(row) > name_col and row[name_col]:
                            name = str(row[name_col]).strip()
                            # Basic name validation
                            if re.match(r'^[A-Za-z\s]{2,}$', name) and len(name.split()) >= 2:
                                driver_info = {"name": name}
                                
                                # Add other columns
                                for i, header in enumerate(table.get("headers", [])):
                                    if i < len(row) and i != name_col:
                                        driver_info[str(header)] = str(row[i]).strip()
                                
                                drivers.append(driver_info)
        
        return drivers

    def _extract_compliance_summary(self, text: str, tables: List[Dict]) -> Dict[str, Any]:
        """Extract compliance information"""
        compliance = {
            "standards_compliance": {},
            "compliance_codes": {},
            "audit_results": []
        }
        
        # Look for compliance tables
        for table in tables:
            headers = [str(h).lower() for h in table.get("headers", [])]
            
            if any(keyword in ' '.join(headers) for keyword in ['compliance', 'standard', 'requirement']):
                data = table.get("data", [])
                for row in data:
                    if len(row) >= 2:
                        standard = str(row[0]).strip()
                        code = str(row[1]).strip()
                        if standard.startswith('Std') and code in ['V', 'NC', 'SFI', 'NAP', 'NA']:
                            compliance["standards_compliance"][standard] = code
        
        # Extract compliance codes definitions
        code_patterns = {
            'V': r'\bV\b\s+([^\n\r]+)',
            'NC': r'\bNC\b\s+([^\n\r]+)',
            'SFI': r'\bSFI\b\s+([^\n\r]+)',
            'NAP': r'\bNAP\b\s+([^\n\r]+)',
            'NA': r'\bNA\b\s+([^\n\r]+)',
        }
        
        for code, pattern in code_patterns.items():
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                compliance["compliance_codes"][code] = match.group(1).strip()
        
        return compliance

    def _extract_dates_and_numbers_improved(self, text: str) -> Dict[str, Any]:
        """Improved date and number extraction"""
        result = {
            "dates": [],
            "registration_numbers": [],
            "phone_numbers": [],
            "email_addresses": [],
            "reference_numbers": []
        }
        
        # Date patterns
        date_patterns = [
            r'\b(\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+\d{4})\b',
            r'\b(\d{1,2}/\d{1,2}/\d{4})\b',
            r'\b(\d{1,2}-\d{1,2}-\d{4})\b',
            r'\b(\d{1,2}\.\d{1,2}\.\d{4})\b',
        ]
        for pattern in date_patterns:
            result["dates"].extend(re.findall(pattern, text))
        
        # Registration numbers (Australian format-ish)
        reg_pattern = r'\b([A-Z]{1,3}\s*\d{1,3}\s*[A-Z]{0,3})\b'
        result["registration_numbers"] = list(set(re.findall(reg_pattern, text)))
        
        # Phone numbers (AU)
        phone_pattern = r'\b((?:\+61|0)[2-9]\s?\d{4}\s?\d{4})\b'
        result["phone_numbers"] = list(set(re.findall(phone_pattern, text)))
        
        # Email addresses
        email_pattern = r'\b([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})\b'
        result["email_addresses"] = list(set(re.findall(email_pattern, text)))
        
        # Reference numbers
        ref_patterns = [
            (r'RF(?:S)?\s*#?\s*(\d+)', 'RFS_Certifications'),
            (r'NHVAS\s+Accreditation\s+No\.?\s*(\d+)', 'NHVAS_Numbers'),
            (r'Registration\s+Number\s*#?\s*(\d+)', 'Registration_Numbers'),
        ]
        for pattern, key in ref_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                result["reference_numbers"].extend([f"{key}: {m}" for m in matches])
        
        return result

    @staticmethod
    def save_results(results: Dict[str, Any], output_path: str):
        """Save results to JSON file"""
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(results, f, indent=2, ensure_ascii=False)
            logger.info(f"💾 Results saved to {output_path}")
        except Exception as e:
            logger.error(f"Failed to save results: {e}")

    @staticmethod
    def export_to_excel(results: Dict[str, Any], excel_path: str):
        """Export results to Excel with improved formatting"""
        try:
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                # Summary sheet
                summary_data = []
                extraction_summary = results.get("extraction_summary", {})
                for key, value in extraction_summary.items():
                    summary_data.append({"Metric": key.replace("_", " ").title(), "Value": value})
                pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
                
                # Key-value pairs
                kv_pairs = results.get("extracted_data", {}).get("key_value_pairs", {})
                if kv_pairs:
                    kv_df = pd.DataFrame(list(kv_pairs.items()), columns=['Key', 'Value'])
                    kv_df.to_excel(writer, sheet_name='Key_Value_Pairs', index=False)
                
                # Vehicle registrations
                vehicles = results.get("extracted_data", {}).get("vehicle_registrations", [])
                if vehicles:
                    pd.DataFrame(vehicles).to_excel(writer, sheet_name='Vehicle_Registrations', index=False)
                
                # Driver records
                drivers = results.get("extracted_data", {}).get("driver_records", [])
                if drivers:
                    pd.DataFrame(drivers).to_excel(writer, sheet_name='Driver_Records', index=False)
                
                # Compliance summary
                compliance = results.get("extracted_data", {}).get("compliance_summary", {})
                if compliance.get("standards_compliance"):
                    comp_df = pd.DataFrame(list(compliance["standards_compliance"].items()), 
                                           columns=['Standard', 'Compliance_Code'])
                    comp_df.to_excel(writer, sheet_name='Compliance_Standards', index=False)
                
                logger.info(f"📊 Results exported to Excel: {excel_path}")
        except Exception as e:
            logger.error(f"Failed to export to Excel: {e}")

def main():
    if len(sys.argv) < 2:
        print("Usage: python fixed_pdf_extractor.py <pdf_path>")
        sys.exit(1)
    
    pdf_path = Path(sys.argv[1])
    if not pdf_path.exists():
        print(f"❌ PDF not found: {pdf_path}")
        sys.exit(1)
    
    print("🚀 Fixed PDF Data Extractor")
    print("=" * 50)
    
    extractor = FixedPDFExtractor()
    results = extractor.extract_everything(str(pdf_path))
    
    base = pdf_path.stem
    output_dir = pdf_path.parent
    
    # Save outputs
    json_path = output_dir / f"{base}_comprehensive_data.json"
    excel_path = output_dir / f"{base}_fixed_extraction.xlsx"
    
    FixedPDFExtractor.save_results(results, str(json_path))
    FixedPDFExtractor.export_to_excel(results, str(excel_path))
    
    print("\n💾 OUTPUT FILES:")
    print(f"   📄 JSON Data: {json_path}")
    print(f"   📊 Excel Data: {excel_path}")
    print(f"\n✨ FIXED EXTRACTION COMPLETE!")

if __name__ == "__main__":
    main()
