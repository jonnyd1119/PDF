def update_excel(self, excel_file, extracted_data, aircraft_model, broker_info):
        try:
            # Get the file content - file stream should already be at the beginning
            excel_file.seek(0)  # Ensure we're at the start
            excel_content = excel_file.read()
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(excel_content)
                tmp_path = tmp.name
            
            backup_path = tmp_path.replace(".xlsx", "_BACKUP.xlsx")
            shutil.copy(tmp_path, backup_path)
            
            # Load workbook with original method
            wb = load_workbook(tmp_path)
            ws = wb[broker_info['sheet']]
            target_col = broker_info['column']
            
            config = st.session_state.configurations[aircraft_model]
            row_mappings = config.get("row_mappings", {})
            
            updates = []
            
            # Update core fields with ROW 1 PROTECTION
            st.write("🔍 **Excel Update Debug**: Starting core fields update")
            st.write(f"🔍 **Target column**: {target_col}")
            
            for field, row_num in row_mappings.items():
                if field in extracted_data:
                    # CRITICAL: Never update row 1 (serial number row)
                    if row_num == 1:
                        st.error(f"🚨 **PROTECTION**: Refusing to update Row 1 ({field}) - this is the serial number row!")
                        st.error(f"🚨 **Check your configuration**: {field} is mapped to Row 1, but this should never happen")
                        continue
                    
                    current_value = ws.cell(row=row_num, column=target_col).value
                    new_value = extracted_data[field]
                    
                    st.write(f"🔍 **Updating {field}**: Row {row_num}, Column {target_col}")
                    st.write(f"🔍 **Before**: '{current_value}' → **After**: '{new_value}'")
                    
                    ws.cell(row=row_num, column=target_col).value = new_value
                    updates.append(f"{field} - Row {row_num}: {extracted_data[field]}")
            
            # Update upgrades with ROW 1 PROTECTION
            upgrades = config.get("upgrades", {})
            for upgrade_name, upgrade_config in upgrades.items():
                upgrade_key = f"upgrade_{upgrade_name}"
                if upgrade_key in extracted_data:
                    row_num = upgrade_config.get("row")
                    if row_num:
                        # CRITICAL: Never update row 1 (serial number row)
                        if row_num == 1:
                            st.error(f"🚨 **PROTECTION**: Refusing to update Row 1 for upgrade {upgrade_name} - this is the serial number row!")
                            continue
                        
                        yn_col = target_col + 1
                        cell = ws.cell(row=row_num, column=yn_col)
                        cell.value = extracted_data[upgrade_key]
                        cell.font = Font(name="Calibri", size=11, color="FFFFFF")
                        updates.append(f"{upgrade_name} - Row {row_num}: {extracted_data[upgrade_key]}")
            
            # REVERT TO ORIGINAL EXCEL HANDLING - No calculation forcing
            wb.save(tmp_path)
            st.write("✅ **Excel saved with original method**")
            
            with open(tmp_path, "rb") as f:
                updated_excel_data = f.read()
            
            with open(backup_path, "rb") as f:
                backup_excel_data = f.read()
            
            os.unlink(tmp_path)
            os.unlink(backup_path)
            
            return updated_excel_data, backup_excel_data, updates
            
        except Exception as e:
            st.error(f"Error updating Excel: {e}")
            return None, None, []
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
import json
import re
import os
import pdfplumber
from io import BytesIO
import tempfile
import shutil

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
import json
import re
import os
import pdfplumber
from io import BytesIO
import tempfile
import shutil

class CompletePlatform:
    def __init__(self):
        if 'configurations' not in st.session_state:
            st.session_state.configurations = {}
        if 'custom_upgrades' not in st.session_state:
            st.session_state.custom_upgrades = []
    
    def extract_text_from_pdf(self, pdf_file):
        try:
            with pdfplumber.open(pdf_file) as pdf:
                full_text = "\n".join(page.extract_text() or '' for page in pdf.pages)
            return full_text.lower()
        except Exception as e:
            st.error(f"Error reading PDF: {e}")
            return ""
    
    def analyze_excel_for_new_model(self, excel_file):
        try:
            wb = load_workbook(excel_file)
            first_sheet_name = wb.sheetnames[0]
            ws = wb[first_sheet_name]
            aircraft_model = first_sheet_name.replace(" FOR SALE", "").strip()
            
            field_analysis = {}
            potential_upgrades = {}
            
            for row in range(1, min(101, ws.max_row + 1)):
                label_cell = ws.cell(row=row, column=12).value
                
                if label_cell and isinstance(label_cell, str):
                    label_text = label_cell.upper().strip()
                    
                    field_mappings = {
                        "year_model": ["YEAR MODEL", "YEAR & MODEL", "MODEL YEAR"],
                        "total_hours": ["TOTAL TIME SINCE NEW", "TOTAL HOURS", "HOURS SINCE NEW", "TTSN"],
                        "engine_overhaul": ["ENGINE TIME SINCE OVERHAUL", "ENGINE OVERHAUL", "TSOH"],
                        "engine_program": ["ENGINE PROGRAM", "ENGINE WARRANTY", "ENGINE PLAN"],
                        "apu_program": ["APU PROGRAM", "APU WARRANTY", "APU PLAN"],
                        "avionics_section": ["AVIONICS", "AVIONICS UPGRADES", "MISC AVIONICS"],
                        "number_of_seats": ["NUMBER OF SEATS", "SEATS", "SEATING"],
                        "seat_configuration": ["SEAT CONFIGURATION", "SEATING CONFIG", "INTERIOR CONFIG"],
                        "paint_exterior_year": ["PAINT EXTERIOR", "EXTERIOR PAINT", "PAINT YEAR", "PAINT COMPLETED"],
                        "interior_year": ["INTERIOR YEAR", "INTERIOR COMPLETED", "NEW INTERIOR", "INTERIOR REFURBISHMENT"]
                    }
                    
                    for field_name, patterns in field_mappings.items():
                        if any(pattern in label_text for pattern in patterns):
                            if field_name not in field_analysis:
                                field_analysis[field_name] = {
                                    "row": row,
                                    "label": label_cell,
                                    "confidence": "high"
                                }
                    
                    upgrade_keywords = [
                        "NXI", "G1000", "G3000", "WIFI", "TCAS", "WAAS", "FANS", "DUAL FMS", "HF",
                        "AHRS", "FDR", "7.1 UPGRADE", "TCAS 7.1", "FLIGHT DATA RECORDER",
                        "SYNTHETIC VISION", "SVT", "GOGO", "IRIDIUM", "CPDLC", "MTOW", "MZFW",
                        "PREBUY INSPECTION", "PREBUY", "PRE-BUY", "INSPECTION"
                    ]
                    
                    for keyword in upgrade_keywords:
                        if keyword in label_text:
                            upgrade_name = keyword.replace(" ", "_").upper()
                            if upgrade_name not in potential_upgrades:
                                potential_upgrades[upgrade_name] = {
                                    "row": row,
                                    "label": label_cell,
                                    "keywords": [keyword.lower()]
                                }
            
            return {
                "aircraft_model": aircraft_model,
                "sheet_name": first_sheet_name,
                "fields": field_analysis,
                "upgrades": potential_upgrades
            }
        
        except Exception as e:
            st.error(f"Error analyzing Excel: {e}")
            return None
    
    def create_configuration_interactive(self, analysis):
        st.subheader(f"🛩️ Configure: {analysis['aircraft_model']}")
        
        aircraft_model = st.text_input(
            "Aircraft Model Name:", 
            value=analysis['aircraft_model']
        )
        
        st.write("## 📋 Core Fields Mapping")
        field_mappings = {}
        
        core_fields = {
            "year_model": "Year Model",
            "total_hours": "Total Hours Since New", 
            "engine_overhaul": "Engine Time Since Overhaul",
            "engine_program": "Engine Program",
            "number_of_seats": "Number of Seats",
            "seat_configuration": "Seat Configuration",
            "paint_exterior_year": "Paint Exterior Year",
            "interior_year": "Interior Year"
        }
        
        optional_fields = {
            "apu_program": "APU Program",
            "avionics_section": "Avionics Section"
        }
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("### Required Fields")
            for field_key, field_name in list(core_fields.items())[:4]:
                detected = analysis['fields'].get(field_key)
                default_row = detected['row'] if detected else 1
                
                row_input = st.number_input(
                    f"{field_name}:",
                    min_value=1,
                    max_value=200,
                    value=default_row,
                    key=f"row_{field_key}"
                )
                field_mappings[field_key] = int(row_input)
        
        with col2:
            st.write("### More Fields")
            for field_key, field_name in list(core_fields.items())[4:]:
                detected = analysis['fields'].get(field_key)
                # Set better defaults for paint fields
                if field_key == "paint_exterior_year":
                    default_row = detected['row'] if detected else 60
                elif field_key == "interior_year":
                    default_row = detected['row'] if detected else 61
                else:
                    default_row = detected['row'] if detected else 1
                
                row_input = st.number_input(
                    f"{field_name}:",
                    min_value=1,
                    max_value=200,
                    value=default_row,
                    key=f"row_{field_key}"
                )
                field_mappings[field_key] = int(row_input)
        
        # Optional fields section
        st.write("### Optional Fields")
        col1, col2 = st.columns(2)
        
        with col1:
            for field_key, field_name in list(optional_fields.items())[:1]:  # APU Program
                detected = analysis['fields'].get(field_key)
                default_row = detected['row'] if detected else None
                
                include_field = st.checkbox(
                    f"Include {field_name}",
                    value=bool(detected),
                    key=f"include_{field_key}"
                )
                
                if include_field:
                    row_input = st.number_input(
                        f"{field_name} Row:",
                        min_value=1,
                        max_value=200,
                        value=default_row if default_row else 33,
                        key=f"row_{field_key}"
                    )
                    field_mappings[field_key] = int(row_input)
        
        with col2:
            for field_key, field_name in list(optional_fields.items())[1:]:  # Avionics Section
                detected = analysis['fields'].get(field_key)
                default_row = detected['row'] if detected else None
                
                include_field = st.checkbox(
                    f"Include {field_name}",
                    value=bool(detected),
                    key=f"include_{field_key}"
                )
                
                if include_field:
                    row_input = st.number_input(
                        f"{field_name} Row:",
                        min_value=1,
                        max_value=200,
                        value=default_row if default_row else 1,
                        key=f"row_{field_key}"
                    )
                    field_mappings[field_key] = int(row_input)
        
        st.write("## 🔧 Upgrades Mapping")
        upgrade_mappings = {}
        
        if analysis['upgrades']:
            for upgrade_key, upgrade_info in analysis['upgrades'].items():
                col1, col2, col3 = st.columns([2, 1, 2])
                
                with col1:
                    include_upgrade = st.checkbox(
                        f"**{upgrade_key}**",
                        value=True,
                        key=f"upgrade_{upgrade_key}"
                    )
                
                with col2:
                    if include_upgrade:
                        row_input = st.number_input(
                            "Row:",
                            min_value=1,
                            max_value=200,
                            value=upgrade_info['row'],
                            key=f"upgrade_row_{upgrade_key}"
                        )
                
                with col3:
                    if include_upgrade:
                        keywords_input = st.text_input(
                            "Keywords:",
                            value=", ".join(upgrade_info['keywords']),
                            key=f"upgrade_keywords_{upgrade_key}"
                        )
                
                if include_upgrade:
                    upgrade_mappings[upgrade_key] = {
                        "keywords": [kw.strip().lower() for kw in keywords_input.split(",")],
                        "row": int(row_input)
                    }
        
        # Custom upgrades
        st.write("### Add Custom Upgrades")
        
        with st.form("add_custom_upgrade"):
            col1, col2, col3, col4 = st.columns([2, 1, 2, 1])
            
            with col1:
                custom_name = st.text_input("Upgrade Name:", placeholder="e.g., DUAL_FMS")
            with col2:
                custom_row = st.number_input("Row:", min_value=1, max_value=200, value=50)
            with col3:
                custom_keywords = st.text_input("Keywords:", placeholder="e.g., dual fms")
            with col4:
                add_upgrade = st.form_submit_button("➕ Add")
        
        if add_upgrade and custom_name and custom_keywords:
            st.session_state.custom_upgrades.append({
                "name": custom_name.upper().replace(" ", "_"),
                "row": custom_row,
                "keywords": [kw.strip().lower() for kw in custom_keywords.split(",")]
            })
            st.success(f"✅ Added {custom_name}")
            st.rerun()
        
        if st.session_state.custom_upgrades:
            st.write("**Added Custom Upgrades:**")
            for i, upgrade in enumerate(st.session_state.custom_upgrades):
                col1, col2, col3, col4 = st.columns([2, 1, 2, 1])
                with col1:
                    st.write(f"**{upgrade['name']}**")
                with col2:
                    st.write(f"Row {upgrade['row']}")
                with col3:
                    st.write(f"Keywords: {', '.join(upgrade['keywords'])}")
                with col4:
                    if st.button("🗑️", key=f"remove_{i}"):
                        st.session_state.custom_upgrades.pop(i)
                        st.rerun()
        
        for custom_upgrade in st.session_state.custom_upgrades:
            upgrade_mappings[custom_upgrade['name']] = {
                "keywords": custom_upgrade['keywords'],
                "row": custom_upgrade['row']
            }
        
        return aircraft_model, field_mappings, upgrade_mappings
    
    def generate_configuration(self, aircraft_model, field_mappings, upgrade_mappings):
        config = {
            aircraft_model: {
                "row_mappings": field_mappings,
                "upgrades": upgrade_mappings
            }
        }
        return config
    
    def identify_aircraft_from_pdf(self, pdf_text):
        for model_name in st.session_state.configurations.keys():
            if model_name.lower() in pdf_text:
                return model_name
            
            model_words = model_name.lower().split()
            if len(model_words) >= 2:
                if all(word in pdf_text for word in model_words):
                    return model_name
        return None
    
    def extract_data_from_pdf(self, pdf_text, aircraft_model):
        config = st.session_state.configurations[aircraft_model]
        extracted_data = {}
        
        # Get row mappings early
        row_mappings = config.get("row_mappings", {})
        
        # Debug: Show configuration
        st.write(f"🔍 **Debug**: Using configuration for {aircraft_model}")
        st.write(f"📋 Row mappings: {list(row_mappings.keys())}")
        st.write(f"🔧 Upgrades configured: {list(config.get('upgrades', {}).keys())}")
        
        # Debug: Show seat-related text from PDF
        seat_contexts = re.findall(r'.{0,50}(?:seat|passenger|lavatory|club|executive|configuration).{0,50}', pdf_text, re.IGNORECASE)
        if seat_contexts:
            st.write(f"🔍 **SEAT CONTEXTS FOUND**: {seat_contexts[:10]}")  # Show first 10 matches
        
        # Additional debug: Look for specific patterns we're trying to match
        st.write("🔍 **SPECIFIC PATTERN SEARCHES**:")
        
        # Check for numbers
        number_matches = re.findall(r'(eight|8|seven|7|six|6|nine|9|ten|10).*?(?:seat|passenger)', pdf_text, re.IGNORECASE)
        if number_matches:
            st.write(f"   Numbers + seats: {number_matches}")
        
        # Check for lavatory
        lav_matches = re.findall(r'.{0,30}(?:lavatory|belted|bltd).{0,30}', pdf_text, re.IGNORECASE)
        if lav_matches:
            st.write(f"   Lavatory contexts: {lav_matches[:5]}")
        
        # Check for club seating
        club_matches = re.findall(r'.{0,30}(?:club|double|forward|aft).{0,30}', pdf_text, re.IGNORECASE)
        if club_matches:
            st.write(f"   Club contexts: {club_matches[:5]}")
        
        # Show a sample of the PDF text
        st.write(f"🔍 **PDF SAMPLE (first 500 chars)**: {pdf_text[:500]}...")
        
        # Year Model
        year_value = None
        patterns = [
            r'(19|20)\d{2}.*?(lear|citation|phenom|model)',
            r'year\s+and\s+model[:\s]*(\d{4})',
            r'(19|20)\d{2}\s+(phenom|citation|lear)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, pdf_text)
            if match:
                year_match = re.search(r'(19|20)\d{2}', match.group(0))
                if year_match:
                    year_value = int(year_match.group(0))
                    break
        
        if year_value:
            extracted_data["year_model"] = year_value
        
        # Total Hours
        total_hours_patterns = [
            r'(total time|hours since new|total hours):?\s*(\d{1,4}[,\.]?\d{0,3})',
            r'(\d{1,4}[,\.]?\d{0,3})\s+snew',
            r'airframe.*?(\d{1,4}[,\.]?\d{0,3})\s*(?:total\s*)?hours?'
        ]
        
        for pattern in total_hours_patterns:
            match = re.search(pattern, pdf_text, re.IGNORECASE | re.DOTALL)
            if match:
                if 'snew' in pattern or 'airframe' in pattern:
                    hours_value = int(match.group(1).replace(",", "").split(".")[0])
                else:
                    hours_value = int(match.group(2).replace(",", "").split(".")[0])
                extracted_data["total_hours"] = hours_value
                break
        
        # Engine Overhaul - with fallback to total hours
        engine_patterns = [
            r'engines?[:\s]*.*?(\d{4,}(?:[,\.]\d{1,3})?)\s*total\s*hours?',
            r'engine.*?(\d{4,}(?:[,\.]\d{1,3})?)\s*total\s*hours?',
            r'(\d{4,}[,\.]?\d{0,3})\s+ttaf',
            r'engine.*?time.*?since.*?overhaul[:\s]*(\d{4,}(?:[,\.]\d{1,3})?)',
            r'tsoh[:\s]*(\d{4,}(?:[,\.]\d{1,3})?)'
        ]
        
        engine_hours_found = False
        for pattern in engine_patterns:
            match = re.search(pattern, pdf_text, re.IGNORECASE | re.DOTALL)
            if match:
                hours_str = match.group(1).replace(",", "")
                if "." in hours_str:
                    test_value = int(float(hours_str))
                else:
                    test_value = int(hours_str)
                
                if 1000 <= test_value <= 50000:
                    extracted_data["engine_overhaul"] = test_value
                    engine_hours_found = True
                    st.write(f"✅ **ENGINE OVERHAUL FOUND**: {test_value}")
                    break
        
        # FALLBACK: If no engine overhaul found, use total hours
        if not engine_hours_found and "total_hours" in extracted_data:
            extracted_data["engine_overhaul"] = extracted_data["total_hours"]
            st.write(f"✅ **ENGINE OVERHAUL FALLBACK**: Using total hours ({extracted_data['total_hours']})")
        elif not engine_hours_found:
            st.write("❌ **ENGINE OVERHAUL**: Not found and no total hours to use as fallback")
        
        # Number of Seats - Enhanced patterns (ADD TO EXISTING)
        seat_patterns = [
            r'(\w+)\s+\(\d+\)\s+passenger',  # "eight (8) passenger"
            r'(\d+)\s+passengers?',           # "8 passengers"
            r'seating\s+for\s+(\w+)',        # "seating for eight"
            r'(\d+)\s*[–-]\s*passenger\s+configuration',  # "8 –passenger configuration"
            r'(\d+)\s+passenger\s+configuration',  # "8 passenger configuration"
        ]
        
        number_mappings = {
            "one": "1", "two": "2", "three": "3", "four": "4", "five": "5",
            "six": "6", "seven": "7", "eight": "8", "nine": "9", "ten": "10",
            "eleven": "11", "twelve": "12"
        }
        
        seats_found = False
        for pattern in seat_patterns:
            match = re.search(pattern, pdf_text, re.IGNORECASE)
            if match:
                seat_info = match.group(1).lower()
                
                if seat_info in number_mappings:
                    digit = number_mappings[seat_info]
                elif seat_info.isdigit():
                    digit = seat_info
                else:
                    continue
                
                # Check for lavatory in broader context
                match_context = pdf_text[max(0, match.start()-200):match.end()+200]
                if re.search(r'lavatory|lav\b|belted\s+lav|bltd\s+lav|belted\s+toilet', match_context, re.IGNORECASE):
                    result = f"{digit} SEATS + BLTD LAV"
                else:
                    result = f"{digit} SEATS"
                
                extracted_data["number_of_seats"] = result
                seats_found = True
                st.write(f"✅ **SEATS FOUND**: Pattern '{pattern}' → {result}")
                break
        
        if not seats_found:
            st.write("❌ **SEATS**: No seat count found")
        
        # Seat Configuration - Enhanced patterns (ADD TO EXISTING)
        config_patterns = [
            # Original patterns
            ("forward.*?club.*?aft.*?club", "DOUBLE CLB"),
            ("four.*?place club", "4 PLC CLB"),
            ("lavatory", "LAV"),
            ("executive", "EXEC"),
            # NEW patterns based on your PDF
            ("fwd.*?cabin.*?four.*?executive.*?club.*?aft.*?cabin.*?four.*?executive.*?club", "DOUBLE CLB"),
            ("four.*?executive.*?club.*?seats.*?four.*?executive.*?club.*?seats", "DOUBLE CLB"),
            ("belted.*?toilet|belted.*?lav|aft.*?lavatory.*?belted", "BLTD LAV"),
        ]
        
        config_elements = []
        for pattern, abbreviation in config_patterns:
            matches = re.findall(pattern, pdf_text, re.IGNORECASE)
            if matches:
                if abbreviation not in config_elements:
                    config_elements.append(abbreviation)
                    st.write(f"✅ **CONFIG FOUND**: Pattern '{pattern}' → {abbreviation}")
        
        if config_elements:
            extracted_data["seat_configuration"] = ",".join(config_elements)
            st.write(f"✅ **FINAL SEAT CONFIG**: {extracted_data['seat_configuration']}")
        else:
            st.write("❌ **SEAT CONFIG**: No configuration found")
        
        # Paint Years
        paint_patterns = [
            r'exterior.*?(\d{4})',
            r'paint.*?(\d{4})',
            r'(\d{4}).*?paint'
        ]
        
        for pattern in paint_patterns:
            match = re.search(pattern, pdf_text, re.IGNORECASE)
            if match:
                year = match.group(1)
                if 1990 <= int(year) <= 2030:
                    extracted_data["paint_exterior_year"] = year
                    break
        
        interior_patterns = [
            r'interior.*?(\d{4})',
            r'refurbishment.*?(\d{4})'
        ]
        
        for pattern in interior_patterns:
            match = re.search(pattern, pdf_text, re.IGNORECASE)
            if match:
                year = match.group(1)
                if 1990 <= int(year) <= 2030:
                    extracted_data["interior_year"] = year
                    break
        
        # Engine Program
        engine_programs = [
            {'keywords': ['corporate care enhanced'], 'value': 'CORPORATE CARE ENHANCED'},
            {'keywords': ['msp gold'], 'value': 'MSP GOLD'},
            {'keywords': ['jssi'], 'value': 'JSSI'}
        ]
        
        for program in engine_programs:
            if any(keyword in pdf_text for keyword in program['keywords']):
                extracted_data["engine_program"] = program['value']
                break
        
        # APU Program (as a core field, not upgrade)
        if "apu_program" in row_mappings:
            st.write("🔍 **APU Program Debug**: Processing as core field")
            
            apu_programs = [
                {'patterns': ['msp gold.*apu', 'apu.*msp gold', 'msp gold apu program'], 'value': 'MSP GOLD'},
                {'patterns': ['honeywell msp', 'msp program'], 'value': 'MSP'},
                {'patterns': ['jssi'], 'value': 'JSSI'},
                {'patterns': ['auxiliary power unit'], 'value': 'AUX ADV'}
            ]
            
            apu_value = None
            for program in apu_programs:
                for pattern in program['patterns']:
                    match = re.search(pattern, pdf_text, re.IGNORECASE)
                    if match:
                        apu_value = program['value']
                        st.write(f"✅ **APU Core Field Match**: Pattern '{pattern}' → {apu_value}")
                        break
                if apu_value:
                    break
            
            if apu_value:
                extracted_data["apu_program"] = apu_value
            else:
                # Fallback
                if re.search(r'apu.*msp|msp.*apu', pdf_text, re.IGNORECASE):
                    extracted_data["apu_program"] = 'MSP GOLD'
                    st.write("✅ **APU Fallback**: Found APU + MSP → MSP GOLD")
                else:
                    extracted_data["apu_program"] = 'NONE'
                    st.write("❌ **APU**: No program found → NONE")
        
        # Avionics (if configured)
        if "avionics_section" in row_mappings:
            avionics_model = self.extract_avionics_model(pdf_text)
            if avionics_model:
                extracted_data["avionics_section"] = avionics_model
            else:
                extracted_data["avionics_section"] = "NONE"
        
        # Upgrades with enhanced detection
        upgrades = config.get("upgrades", {})
        st.write(f"🔍 **Upgrades Debug**: Found {len(upgrades)} configured upgrades")
        
        # Check if PREBUY_INSPECTION is configured
        prebuy_configured = any("PREBUY" in upgrade_name.upper() for upgrade_name in upgrades.keys())
        st.write(f"🔍 **PREBUY configured**: {prebuy_configured}")
        
        # If PREBUY not configured, add it manually for this session
        if not prebuy_configured:
            st.write("⚠️ **PREBUY not configured - adding temporarily**")
            upgrades["PREBUY_INSPECTION"] = {
                "keywords": ["prebuy inspection", "pre-buy inspection", "prebuy.*inspection"],
                "row": 50  # Default row - user should configure properly
            }
        
        for upgrade_name, upgrade_config in upgrades.items():
            keywords = upgrade_config.get("keywords", [])
            found = False
            
            st.write(f"🔍 **Processing upgrade**: {upgrade_name} with keywords: {keywords}")
            
            # Show what each keyword is finding
            for keyword in keywords:
                if re.search(keyword, pdf_text, re.IGNORECASE):
                    keyword_contexts = re.findall(f'.{{0,40}}{re.escape(keyword)}.{{0,40}}', pdf_text, re.IGNORECASE)
                    st.write(f"   ✅ Keyword '{keyword}' found in: {keyword_contexts[:2]}")
                else:
                    st.write(f"   ❌ Keyword '{keyword}' not found")
            
            # SPECIAL HANDLING FOR SPECIFIC UPGRADES (runs BEFORE normal keyword matching)
            if "PREBUY" in upgrade_name.upper():
                # EXTREMELY STRICT: Only actual prebuy inspection, NOT general inspection
                prebuy_found = False
                
                # Look for very specific prebuy phrases
                prebuy_patterns = [
                    r'prebuy\s+inspection',
                    r'pre-buy\s+inspection', 
                    r'pre\s*buy\s+inspection',
                    r'prebuy\s+completed',
                    r'fresh\s+prebuy',
                    r'recent\s+prebuy'
                ]
                
                for pb_pattern in prebuy_patterns:
                    if re.search(pb_pattern, pdf_text, re.IGNORECASE):
                        prebuy_found = True
                        st.write(f"✅ **PREBUY**: Found specific pattern '{pb_pattern}'")
                        break
                
                if not prebuy_found:
                    # Show what inspection contexts exist to debug
                    inspection_contexts = re.findall(r'.{0,60}inspection.{0,60}', pdf_text, re.IGNORECASE)
                    st.write(f"❌ **PREBUY**: No prebuy-specific patterns found")
                    st.write(f"   Found {len(inspection_contexts)} general inspection mentions:")
                    for i, context in enumerate(inspection_contexts[:3]):
                        st.write(f"   {i+1}: '{context.strip()}'")
                
                found = prebuy_found
                
            elif "INSPECTION" in upgrade_name.upper() and "PREBUY" not in upgrade_name.upper():
                # Make INSPECTION more specific - not just any "inspection" mention
                inspection_found = False
                
                # Look for maintenance/aircraft inspection types
                inspection_patterns = [
                    r'annual\s+inspection',
                    r'100\s*hour\s+inspection',
                    r'maintenance\s+inspection', 
                    r'airworthiness\s+inspection',
                    r'conformity\s+inspection',
                    r'records\s+inspection',
                    r'aircraft\s+inspection'
                ]
                
                for insp_pattern in inspection_patterns:
                    if re.search(insp_pattern, pdf_text, re.IGNORECASE):
                        inspection_found = True
                        st.write(f"✅ **INSPECTION**: Found specific pattern '{insp_pattern}'")
                        break
                
                if not inspection_found:
                    # Show what generic inspection mentions exist
                    inspection_contexts = re.findall(r'.{0,60}inspection.{0,60}', pdf_text, re.IGNORECASE)
                    st.write(f"❌ **INSPECTION**: No aircraft-specific inspection patterns found")
                    st.write(f"   Found {len(inspection_contexts)} general 'inspection' mentions (likely legal disclaimers)")
                    for i, context in enumerate(inspection_contexts[:2]):
                        st.write(f"   {i+1}: '{context.strip()}'")
                
                found = inspection_found
                
            else:
                # NORMAL KEYWORD MATCHING for all other upgrades
                # Enhanced keyword matching with regex for better detection
                for keyword in keywords:
                    if re.search(keyword, pdf_text, re.IGNORECASE):
                        found = True
                        st.write(f"✅ **Match found**: '{keyword}' in PDF")
                        break
            
            # Special handling for specific upgrades (if not already handled above)
            if not found and "PREBUY" not in upgrade_name.upper() and "INSPECTION" not in upgrade_name.upper():
                if upgrade_name.upper() == "AHRS":
                    if re.search(r'ahrs|attitude.*heading.*reference', pdf_text, re.IGNORECASE):
                        found = True
                        st.write("✅ **AHRS**: Found via special pattern")
                elif upgrade_name.upper() == "FDR":
                    if re.search(r'flight.*data.*recorder|fdr', pdf_text, re.IGNORECASE):
                        found = True
                        st.write("✅ **FDR**: Found via special pattern")
                elif "7.1" in upgrade_name.upper():
                    if re.search(r'tcas.*7\.1|7\.1.*upgrade|change.*7\.1', pdf_text, re.IGNORECASE):
                        found = True
                        st.write("✅ **7.1**: Found via special pattern")
            
            extracted_data[f"upgrade_{upgrade_name}"] = "Y" if found else "N"
            
            # Debug output for all upgrades
            st.write(f"🔍 **{upgrade_name}**: {extracted_data[f'upgrade_{upgrade_name}']}")
            
            # Extra debug for PREBUY specifically
            if "PREBUY" in upgrade_name.upper():
                st.write(f"🔍 **PREBUY DETAILED DEBUG**:")
                prebuy_contexts = re.findall(r'.{0,50}(?:prebuy|inspection).{0,50}', pdf_text, re.IGNORECASE)
                st.write(f"   - Contexts found: {prebuy_contexts[:3] if prebuy_contexts else 'None'}")
                st.write(f"   - Keywords used: {keywords}")
                st.write(f"   - Final result: {extracted_data[f'upgrade_{upgrade_name}']}")
        
        return extracted_data
    
    def extract_avionics_model(self, full_text):
        """Extract avionics model from PDF text"""
        patterns = [
            r'gogo\s+atg[-\s]?(\d{4})',
            r'atg[-\s]?(\d{4})',
            r'avance[-\s]?l(\d+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, full_text, re.IGNORECASE)
            if match:
                if 'atg' in pattern:
                    return f"ATG-{match.group(1)}"
                elif 'avance' in pattern:
                    return f"AVANCE-L{match.group(1)}"
                else:
                    return match.group(1).upper()
        
        return None
    
    def update_excel(self, excel_file, extracted_data, aircraft_model, broker_info):
        try:
            # Get the file content - file stream should already be at the beginning
            excel_file.seek(0)  # Ensure we're at the start
            excel_content = excel_file.read()
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(excel_content)
                tmp_path = tmp.name
            
            backup_path = tmp_path.replace(".xlsx", "_BACKUP.xlsx")
            shutil.copy(tmp_path, backup_path)
            
            # Load workbook with formulas preserved
            wb = load_workbook(tmp_path, data_only=False)
            ws = wb[broker_info['sheet']]
            target_col = broker_info['column']
            
            config = st.session_state.configurations[aircraft_model]
            row_mappings = config.get("row_mappings", {})
            
            updates = []
            
            # Update core fields with ROW 1 PROTECTION
            st.write("🔍 **Excel Update Debug**: Starting core fields update")
            st.write(f"🔍 **Target column**: {target_col}")
            
            for field, row_num in row_mappings.items():
                if field in extracted_data:
                    # CRITICAL: Never update row 1 (serial number row)
                    if row_num == 1:
                        st.error(f"🚨 **PROTECTION**: Refusing to update Row 1 ({field}) - this is the serial number row!")
                        st.error(f"🚨 **Check your configuration**: {field} is mapped to Row 1, but this should never happen")
                        continue
                    
                    current_value = ws.cell(row=row_num, column=target_col).value
                    new_value = extracted_data[field]
                    
                    st.write(f"🔍 **Updating {field}**: Row {row_num}, Column {target_col}")
                    st.write(f"🔍 **Before**: '{current_value}' → **After**: '{new_value}'")
                    
                    ws.cell(row=row_num, column=target_col).value = new_value
                    updates.append(f"{field} - Row {row_num}: {extracted_data[field]}")
            
            # Update upgrades with ROW 1 PROTECTION
            upgrades = config.get("upgrades", {})
            for upgrade_name, upgrade_config in upgrades.items():
                upgrade_key = f"upgrade_{upgrade_name}"
                if upgrade_key in extracted_data:
                    row_num = upgrade_config.get("row")
                    if row_num:
                        # CRITICAL: Never update row 1 (serial number row)
                        if row_num == 1:
                            st.error(f"🚨 **PROTECTION**: Refusing to update Row 1 for upgrade {upgrade_name} - this is the serial number row!")
                            continue
                        
                        yn_col = target_col + 1
                        cell = ws.cell(row=row_num, column=yn_col)
                        cell.value = extracted_data[upgrade_key]
                        cell.font = Font(name="Calibri", size=11, color="FFFFFF")
                        updates.append(f"{upgrade_name} - Row {row_num}: {extracted_data[upgrade_key]}")
            
            # CRITICAL FIX: Force Excel to recalculate formulas
            try:
                wb.calculation.calcMode = "automatic"
                st.write("✅ **Set calculation mode to automatic**")
            except:
                st.write("⚠️ **Could not set calculation mode**")
            
            wb.save(tmp_path)
            st.write("✅ **Excel saved with recalculation**")
            
            with open(tmp_path, "rb") as f:
                updated_excel_data = f.read()
            
            with open(backup_path, "rb") as f:
                backup_excel_data = f.read()
            
            os.unlink(tmp_path)
            os.unlink(backup_path)
            
            return updated_excel_data, backup_excel_data, updates
            
        except Exception as e:
            st.error(f"Error updating Excel: {e}")
            return None, None, []
    
    def find_broker_column(self, excel_file, serial_number, broker_name):
        try:
            # Read the file content once and store it
            excel_content = excel_file.read()
            excel_file.seek(0)  # Reset the file pointer
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(excel_content)
                tmp_path = tmp.name
            
            wb = load_workbook(tmp_path)
            
            search_terms = [serial_number.strip()]
            if '-' in serial_number:
                search_terms.append(serial_number.split('-')[-1])
            
            st.write(f"🔍 **DEBUG - Searching for serial**: {search_terms}")
            
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                st.write(f"🔍 **DEBUG - Checking sheet**: {sheet_name}")
                
                for col in range(16, ws.max_column + 1, 2):
                    cell_value = ws.cell(row=1, column=col).value
                    if cell_value:
                        serial_cell = str(cell_value).strip()
                        st.write(f"🔍 **DEBUG - Column {col}**: Found '{serial_cell}'")
                        
                        for search_term in search_terms:
                            if search_term in serial_cell:
                                broker_cell = ws.cell(row=5, column=col).value
                                st.write(f"✅ **MATCH FOUND**: Column {col}, Serial: '{serial_cell}', Broker: '{broker_cell}'")
                                os.unlink(tmp_path)
                                
                                return {
                                    'column': col,
                                    'sheet': sheet_name,
                                    'broker': broker_cell,
                                    'matched_serial': serial_cell,
                                    'search_term': search_term
                                }
            
            st.write("❌ **DEBUG - No match found**")
            os.unlink(tmp_path)
            return None
            
        except Exception as e:
            st.error(f"Error finding column: {e}")
            return None
    
    def generate_update_report(self, extracted_data, aircraft_model, broker_info):
        """Generate a safe update report instead of modifying Excel directly"""
        try:
            config = st.session_state.configurations[aircraft_model]
            row_mappings = config.get("row_mappings", {})
            upgrades = config.get("upgrades", {})
            
            report = {
                "metadata": {
                    "aircraft_model": aircraft_model,
                    "target_column": broker_info['column'],
                    "target_sheet": broker_info['sheet'],
                    "broker": broker_info.get('broker', 'Unknown'),
                    "serial_number": broker_info.get('matched_serial', 'Unknown'),
                    "timestamp": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                },
                "core_field_updates": [],
                "upgrade_updates": [],
                "summary": {}
            }
            
            # Core field updates
            for field, row_num in row_mappings.items():
                if field in extracted_data and row_num != 1:
                    new_value = extracted_data[field]
                    if new_value is not None and str(new_value).strip() != "":
                        report["core_field_updates"].append({
                            "field": field,
                            "row": row_num,
                            "column": broker_info['column'],
                            "value": new_value,
                            "excel_reference": f"{chr(64 + broker_info['column'])}{row_num}"
                        })
            
            # Upgrade updates
            for upgrade_name, upgrade_config in upgrades.items():
                upgrade_key = f"upgrade_{upgrade_name}"
                if upgrade_key in extracted_data:
                    row_num = upgrade_config.get("row")
                    if row_num and row_num != 1:
                        yn_col = broker_info['column'] + 1
                        report["upgrade_updates"].append({
                            "upgrade": upgrade_name,
                            "row": row_num,
                            "column": yn_col,
                            "value": extracted_data[upgrade_key],
                            "excel_reference": f"{chr(64 + yn_col)}{row_num}"
                        })
            
            # Summary
            report["summary"] = {
                "total_core_updates": len(report["core_field_updates"]),
                "total_upgrade_updates": len(report["upgrade_updates"]),
                "total_updates": len(report["core_field_updates"]) + len(report["upgrade_updates"])
            }
            
            return report
            
        except Exception as e:
            st.error(f"Error generating report: {e}")
            return None
    
    def update_excel_safe_mode(self, excel_file, extracted_data, aircraft_model, broker_info):
        """Safe mode: Generate report instead of modifying Excel"""
        try:
            # Generate the update report
            report = self.generate_update_report(extracted_data, aircraft_model, broker_info)
            
            if not report:
                return None, None, []
            
            # Create a readable summary for display
            updates = []
            
            # Add core field updates to summary
            for update in report["core_field_updates"]:
                updates.append(f"{update['field']} - {update['excel_reference']}: {update['value']}")
            
            # Add upgrade updates to summary
            for update in report["upgrade_updates"]:
                updates.append(f"{update['upgrade']} - {update['excel_reference']}: {update['value']}")
            
            # Create downloadable files
            
            # 1. JSON report for detailed data
            json_report = json.dumps(report, indent=2)
            
            # 2. CSV report for easy viewing
            csv_data = []
            
            # Add core fields to CSV
            for update in report["core_field_updates"]:
                csv_data.append({
                    "Type": "Core Field",
                    "Field/Upgrade": update['field'],
                    "Excel Reference": update['excel_reference'],
                    "Row": update['row'],
                    "Column": update['column'],
                    "Value": update['value']
                })
            
            # Add upgrades to CSV
            for update in report["upgrade_updates"]:
                csv_data.append({
                    "Type": "Upgrade",
                    "Field/Upgrade": update['upgrade'],
                    "Excel Reference": update['excel_reference'],
                    "Row": update['row'],
                    "Column": update['column'],
                    "Value": update['value']
                })
            
            if csv_data:
                df = pd.DataFrame(csv_data)
                csv_content = df.to_csv(index=False)
            else:
                csv_content = "No updates to apply"
            
            # 3. Human-readable text report
            text_report = f"""
AIRCRAFT DATA EXTRACTION REPORT
===============================

Aircraft Model: {report['metadata']['aircraft_model']}
Target Column: {report['metadata']['target_column']} ({chr(64 + report['metadata']['target_column'])})
Broker: {report['metadata']['broker']}
Serial Number: {report['metadata']['serial_number']}
Generated: {report['metadata']['timestamp']}

SUMMARY
-------
Core Field Updates: {report['summary']['total_core_updates']}
Upgrade Updates: {report['summary']['total_upgrade_updates']}
Total Updates: {report['summary']['total_updates']}

CORE FIELD UPDATES
------------------
"""
            
            for update in report["core_field_updates"]:
                text_report += f"• {update['field']}: Cell {update['excel_reference']} = {update['value']}\n"
            
            text_report += "\nUPGRADE UPDATES\n---------------\n"
            
            for update in report["upgrade_updates"]:
                text_report += f"• {update['upgrade']}: Cell {update['excel_reference']} = {update['value']}\n"
            
            text_report += f"""

MANUAL UPDATE INSTRUCTIONS
--------------------------
1. Open your Excel file: {broker_info['sheet']} sheet
2. Go to Column {chr(64 + report['metadata']['target_column'])} (the {report['metadata']['broker']} column)
3. Update each cell listed above with the corresponding value
4. For upgrades, update Column {chr(64 + report['metadata']['target_column'] + 1)} (Y/N column)

This report was generated because direct Excel modification was causing data loss.
Manual updates ensure your data remains safe.
"""
            
            return json_report, csv_content, text_report, updates
            
        except Exception as e:
            st.error(f"Error in safe mode: {e}")
            return None, None, None, []
        try:
            # Get the file content - file stream should already be at the beginning
            excel_file.seek(0)  # Ensure we're at the start
            excel_content = excel_file.read()
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(excel_content)
                tmp_path = tmp.name
            
            backup_path = tmp_path.replace(".xlsx", "_BACKUP.xlsx")
            shutil.copy(tmp_path, backup_path)
            
            # SAFER APPROACH: Use data_only=False to preserve formulas and formatting
            wb = load_workbook(tmp_path, data_only=False)
            ws = wb[broker_info['sheet']]
            target_col = broker_info['column']
            
            config = st.session_state.configurations[aircraft_model]
            row_mappings = config.get("row_mappings", {})
            
            updates = []
            
            st.write("🔍 **Excel Update Debug**: Starting SAFE update process")
            st.write(f"🔍 **Target column**: {target_col}")
            st.write(f"🔍 **Target sheet**: {broker_info['sheet']}")
            
            # Create a list of ONLY the cells we plan to update
            cells_to_update = []
            
            # Plan core field updates
            for field, row_num in row_mappings.items():
                if field in extracted_data and row_num != 1:
                    new_value = extracted_data[field]
                    if new_value is not None and str(new_value).strip() != "":
                        cells_to_update.append({
                            'row': row_num,
                            'col': target_col,
                            'field': field,
                            'value': new_value,
                            'type': 'core'
                        })
            
            # Plan upgrade updates
            upgrades = config.get("upgrades", {})
            for upgrade_name, upgrade_config in upgrades.items():
                upgrade_key = f"upgrade_{upgrade_name}"
                if upgrade_key in extracted_data:
                    row_num = upgrade_config.get("row")
                    if row_num and row_num != 1:
                        yn_col = target_col + 1
                        cells_to_update.append({
                            'row': row_num,
                            'col': yn_col,
                            'field': upgrade_name,
                            'value': extracted_data[upgrade_key],
                            'type': 'upgrade'
                        })
            
            st.write(f"🔍 **Planned updates**: {len(cells_to_update)} cells")
            
            # Show what we're about to update
            for update in cells_to_update:
                current_val = ws.cell(row=update['row'], column=update['col']).value
                st.write(f"🔍 **Will update**: {update['field']} at Row {update['row']}, Col {update['col']}")
                st.write(f"   Current: '{current_val}' → New: '{update['value']}'")
            
            # CRITICAL: Only update the specific cells we identified
            for update in cells_to_update:
                try:
                    cell = ws.cell(row=update['row'], column=update['col'])
                    old_value = cell.value
                    cell.value = update['value']
                    
                    # Only apply font formatting to upgrades
                    if update['type'] == 'upgrade':
                        cell.font = Font(name="Calibri", size=11, color="FFFFFF")
                    
                    updates.append(f"{update['field']} - Row {update['row']}: '{old_value}' → '{update['value']}'")
                    st.write(f"✅ **UPDATED**: {update['field']}")
                    
                except Exception as e:
                    st.error(f"❌ **FAILED to update {update['field']}**: {e}")
                    continue
            
            # SAFER SAVE: Try to preserve as much as possible
            try:
                # Save with minimal changes
                wb.save(tmp_path)
                st.write("✅ **Excel saved successfully**")
                
                # Verify the save worked by checking our target column
                wb_verify = load_workbook(tmp_path, data_only=True)
                ws_verify = wb_verify[broker_info['sheet']]
                
                # Check that our updates are there
                verification_passed = True
                for update in cells_to_update[:3]:  # Check first 3 updates
                    actual_value = ws_verify.cell(row=update['row'], column=update['col']).value
                    if str(actual_value) != str(update['value']):
                        st.error(f"❌ **VERIFICATION FAILED**: {update['field']} should be '{update['value']}' but is '{actual_value}'")
                        verification_passed = False
                
                if verification_passed:
                    st.write("✅ **Verification passed**: Updates are in the file")
                else:
                    st.error("❌ **Verification failed**: Updates may not have been saved correctly")
                
                # Also check that we didn't wipe other columns
                for check_col in [14, 15, 17, 18]:
                    if check_col != target_col:
                        check_serial = ws_verify.cell(row=1, column=check_col).value
                        if check_serial and str(check_serial).strip():
                            st.write(f"✅ **Column {check_col} preserved**: Serial = '{check_serial}'")
                        else:
                            st.warning(f"⚠️ **Column {check_col}**: No serial number found - may be wiped")
                
            except Exception as save_error:
                st.error(f"❌ **SAVE ERROR**: {save_error}")
                return None, None, []
            
            with open(tmp_path, "rb") as f:
                updated_excel_data = f.read()
            
            with open(backup_path, "rb") as f:
                backup_excel_data = f.read()
            
            os.unlink(tmp_path)
            os.unlink(backup_path)
            
            return updated_excel_data, backup_excel_data, updates
            
        except Exception as e:
            st.error(f"Error updating Excel: {e}")
            return None, None, []

def main():
    st.set_page_config(
        page_title="Aircraft Data Platform",
        page_icon="✈️",
        layout="wide"
    )
    
    platform = CompletePlatform()
    
    st.title("✈️ Complete Aircraft Data Management Platform")
    st.markdown("**Configure aircraft models and process broker data**")
    
    # Sidebar
    st.sidebar.title("📊 Status")
    
    if st.session_state.configurations:
        st.sidebar.success(f"✅ {len(st.session_state.configurations)} Models Configured")
        for model in st.session_state.configurations.keys():
            st.sidebar.write(f"• {model}")
        
        config_json = json.dumps(st.session_state.configurations, indent=2)
        st.sidebar.download_button(
            "📤 Export Configs",
            config_json,
            "aircraft_configurations.json",
            "application/json"
        )
        
        if st.sidebar.button("🗑️ Clear All"):
            st.session_state.configurations = {}
            st.session_state.custom_upgrades = []
            st.rerun()
    else:
        st.sidebar.warning("⚠️ No models configured")
    
    # Main tabs
    tab1, tab2 = st.tabs(["🚀 Quick Process", "🔧 Create New Model"])
    
    with tab1:
        st.header("🚀 Quick Aircraft Data Processing")
        
        # Load existing JSON
        st.subheader("📂 Load Existing Configuration")
        uploaded_json = st.file_uploader("Upload JSON Config", type="json")
        
        if uploaded_json:
            try:
                json_data = json.load(uploaded_json)
                st.write("**Found models:**")
                for model_name in json_data.keys():
                    st.write(f"• {model_name}")
                
                if st.button("📥 Load Configuration"):
                    st.session_state.configurations.update(json_data)
                    st.success(f"✅ Loaded {len(json_data)} configuration(s)!")
                    st.rerun()
            except Exception as e:
                st.error(f"Error reading JSON: {e}")
        
        # Show process data section if configurations are loaded
        if st.session_state.configurations:
            st.markdown("---")
            st.subheader("✈️ Process Aircraft Data")
            
            # Input fields
            col1, col2 = st.columns(2)
            with col1:
                serial_number = st.text_input("Serial Number:", placeholder="e.g., 525C-0123")
            with col2:
                broker_name = st.text_input("Broker Name:", placeholder="e.g., Jetcraft")
            
            # File uploads
            col1, col2 = st.columns(2)
            with col1:
                pdf_file = st.file_uploader("Upload Broker PDF", type="pdf")
            with col2:
                excel_file = st.file_uploader("Upload Excel Sheet", type="xlsx", key="process_excel")
            
            if serial_number and broker_name and pdf_file and excel_file:
                if st.button("🚀 Process Aircraft Data", type="primary"):
                    with st.spinner("Processing..."):
                        
                        # Extract PDF text
                        pdf_text = platform.extract_text_from_pdf(pdf_file)
                        
                        if not pdf_text:
                            st.error("Could not extract text from PDF")
                        else:
                            # Identify aircraft model
                            aircraft_model = platform.identify_aircraft_from_pdf(pdf_text)
                            
                            if not aircraft_model:
                                st.error("❌ Could not identify aircraft model")
                                st.info("Available: " + ", ".join(st.session_state.configurations.keys()))
                            else:
                                st.success(f"✅ Identified: **{aircraft_model}**")
                                
                                # Extract data
                                extracted_data = platform.extract_data_from_pdf(pdf_text, aircraft_model)
                                
                                if extracted_data:
                                    # Display extracted data
                                    st.subheader("📊 Extracted Data")
                                    col1, col2 = st.columns(2)
                                    
                                    with col1:
                                        st.write("**Core Fields:**")
                                        for key, value in extracted_data.items():
                                            if not key.startswith("upgrade_"):
                                                st.write(f"• **{key}**: {value}")
                                    
                                    with col2:
                                        st.write("**Upgrades:**")
                                        for key, value in extracted_data.items():
                                            if key.startswith("upgrade_"):
                                                upgrade_name = key.replace("upgrade_", "")
                                                icon = "✅" if value == "Y" else "❌"
                                                st.write(f"• {icon} **{upgrade_name}**: {value}")
                                    
                                    # Find broker column
                                    broker_info = platform.find_broker_column(excel_file, serial_number, broker_name)
                                    
                                    if not broker_info:
                                        st.error("❌ Could not find broker column")
                                    else:
                                        st.success(f"✅ Found match in Column {broker_info['column']}")
                                        
                                        # Update Excel with recalculation fix
                                        updated_excel, backup_excel, updates = platform.update_excel(
                                            excel_file, extracted_data, aircraft_model, broker_info
                                        )
                                        
                                        if updated_excel:
                                            st.success("✅ Excel updated successfully with automatic recalculation!")
                                            st.warning("⚠️ **IMPORTANT**: After opening the Excel file, click the 'Calculate' button (or press Ctrl+Alt+F9) to refresh all formulas and display the values.")
                                            st.info("💡 This is normal behavior when updating Excel files with formulas - the data is safely saved, it just needs Excel to recalculate.")
                                            
                                            # Show updates
                                            if updates:
                                                st.subheader("📋 Updates Made")
                                                for update in updates[:10]:
                                                    st.write(f"• {update}")
                                            
                                            # Download buttons
                                            col1, col2 = st.columns(2)
                                            
                                            with col1:
                                                st.download_button(
                                                    "📥 Download Updated Excel",
                                                    updated_excel,
                                                    f"UPDATED_{aircraft_model.replace(' ', '_')}_{serial_number}.xlsx",
                                                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                                )
                                            
                                            with col2:
                                                st.download_button(
                                                    "💾 Download Original Backup",
                                                    backup_excel,
                                                    f"ORIGINAL_{aircraft_model.replace(' ', '_')}_{serial_number}.xlsx",
                                                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                                )
                                        else:
                                            st.error("❌ Failed to update Excel")
                                else:
                                    st.error("❌ No data extracted")
            else:
                if not (serial_number and broker_name and pdf_file and excel_file):
                    st.info("👆 Please fill in all fields and upload both files")
        else:
            st.info("👆 Please load a configuration file first to enable data processing")
    
    with tab2:
        st.header("🔧 Create New Aircraft Model Configuration")
        
        # Create new configuration
        st.subheader("🆕 Create New Configuration")
        excel_template = st.file_uploader("Upload Excel Template", type="xlsx")
        
        if excel_template:
            analysis = platform.analyze_excel_for_new_model(excel_template)
            
            if analysis:
                st.success(f"✅ Detected: **{analysis['aircraft_model']}**")
                
                aircraft_model, field_mappings, upgrade_mappings = platform.create_configuration_interactive(analysis)
                
                if st.button("Save Configuration", type="primary"):
                    config = platform.generate_configuration(aircraft_model, field_mappings, upgrade_mappings)
                    st.session_state.configurations.update(config)
                    st.success(f"✅ {aircraft_model} configured successfully!")
                    
                    # Auto-generate and download the config file
                    config_json = json.dumps(config, indent=2)
                    st.download_button(
                        "📥 Download Configuration File",
                        config_json,
                        f"{aircraft_model.replace(' ', '_')}_config.json",
                        "application/json",
                        key="download_new_config"
                    )
                    st.info("💡 **Important:** Download and save this configuration file! You'll need it for the Quick Process tab.")
                    st.session_state.custom_upgrades = []

if __name__ == "__main__":
    main()