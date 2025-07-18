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
import io

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
                        "PREBUY INSPECTION", "PREBUY", "PRE-BUY", "INSPECTION", "DUAL UNS-1ESPW",
                        "APU", "AUXILIARY POWER UNIT"
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
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("### Core Fields")
            for field_key, field_name in list(core_fields.items())[:4]:
                detected = analysis['fields'].get(field_key)
                default_row = detected['row'] if detected else 20
                
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
                if field_key == "paint_exterior_year":
                    default_row = detected['row'] if detected else 60
                elif field_key == "interior_year":
                    default_row = detected['row'] if detected else 61
                else:
                    default_row = detected['row'] if detected else 30
                
                row_input = st.number_input(
                    f"{field_name}:",
                    min_value=1,
                    max_value=200,
                    value=default_row,
                    key=f"row_{field_key}"
                )
                field_mappings[field_key] = int(row_input)
        
        st.write("## 🔧 Upgrades Mapping")
        
        # Display detected Y/N fields
        if analysis.get('yn_fields_detected', 0) > 0:
            st.info(f"✅ Detected {analysis['yn_fields_detected']} Y/N upgrade fields from existing brokers!")
        
        if analysis.get('item_upgrades_detected', 0) > 0:
            st.info(f"✅ Detected {analysis['item_upgrades_detected']} optional upgrades from ITEM column with default 'N' values!")
        
        upgrade_mappings = {}
        
        # Initialize custom upgrades in session state
        if 'temp_custom_upgrades' not in st.session_state:
            st.session_state.temp_custom_upgrades = []
        
        # Add custom upgrade input section
        st.write("### Add Custom Upgrades")
        
        # Add multiple custom upgrades
        num_custom = st.number_input("Number of custom upgrades to add:", min_value=0, max_value=10, value=1)
        
        for i in range(num_custom):
            st.write(f"#### Custom Upgrade {i+1}")
            col1, col2, col3 = st.columns([2, 1, 2])
            with col1:
                custom_name = st.text_input(f"Upgrade Name {i+1}:", placeholder="e.g., GARMIN G-5000", key=f"custom_name_{i}")
            with col2:
                custom_row = st.number_input(f"Row {i+1}:", min_value=1, max_value=200, value=50, key=f"custom_row_{i}")
            with col3:
                custom_keywords = st.text_input(f"Keywords {i+1}:", placeholder="e.g., garmin g-5000, g5000", key=f"custom_keywords_{i}")
            
            if custom_name:
                upgrade_name = custom_name.upper().replace(" ", "_")
                if upgrade_name not in analysis['upgrades']:
                    analysis['upgrades'][upgrade_name] = {
                        "row": custom_row,
                        "label": custom_name,
                        "keywords": [kw.strip().lower() for kw in custom_keywords.split(",") if kw.strip()]
                    }
        
        st.write("### Detected Upgrades")
        if analysis['upgrades']:
            for upgrade_key, upgrade_info in analysis['upgrades'].items():
                col1, col2, col3 = st.columns([2, 1, 2])
                
                with col1:
                    include_upgrade = st.checkbox(
                        f"**{upgrade_key}**" + 
                        (" (Y/N field)" if upgrade_info.get('is_yn') else "") +
                        (" [from ITEM column]" if upgrade_info.get('from_item_column') else ""),
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
        st.write(f"🔍 **Aircraft Model Debug**: Looking for matches in PDF")
        st.write(f"📋 **Configured models**: {list(st.session_state.configurations.keys())}")
        
        for model_name in st.session_state.configurations.keys():
            
            if model_name.lower() in pdf_text:
                st.write(f"✅ **Direct match found**: {model_name}")
                return model_name
            
            model_words = model_name.lower().split()
            
            if "excel" in model_name.lower():
                if "citation" in pdf_text and "excel" in pdf_text:
                    st.write(f"✅ **Citation Excel match found**: {model_name}")
                    return model_name
                elif "excel" in pdf_text:
                    st.write(f"✅ **Excel match found**: {model_name}")
                    return model_name
            
            if len(model_words) >= 2:
                matches = []
                for word in model_words:
                    if word not in ["-", "master", "for", "sale"] and len(word) > 2:
                        if word in pdf_text:
                            matches.append(word)
                
                important_words = [w for w in model_words if w not in ["-", "master", "for", "sale"] and len(w) > 2]
                if len(matches) >= len(important_words) * 0.6:
                    st.write(f"✅ **Partial match found**: {model_name} (matched {len(matches)}/{len(important_words)} words)")
                    return model_name
        
        st.write("❌ **No aircraft model matches found**")
        return None
    
    def extract_data_from_pdf(self, pdf_text, aircraft_model):
        config = st.session_state.configurations[aircraft_model]
        extracted_data = {}
        row_mappings = config.get("row_mappings", {})
        
        st.write("🔍 **Starting PDF data extraction...**")
        st.write(f"📋 **Configured fields**: {list(row_mappings.keys())}")
        
        # Extract Serial Number
        serial_patterns = [
            r'serial\s+number[:\s]*([A-Z0-9\-]+)',
            r'sn[:\s]*([A-Z0-9\-]+)',
            r's/n[:\s]*([A-Z0-9\-]+)', 
            r'serial[:\s]*([A-Z0-9\-]+)',
            r'(\d{2,4}[A-Z]*\-?\d{3,4})',
            r'airframe[:\s]*([A-Z0-9\-]+)',
            r'aircraft[:\s]*([A-Z0-9\-]+)',
            r'msn[:\s]*([A-Z0-9\-]+)'
        ]
        
        for pattern in serial_patterns:
            match = re.search(pattern, pdf_text, re.IGNORECASE)
            if match:
                serial_candidate = match.group(1).strip()
                if re.search(r'\d', serial_candidate) and len(serial_candidate) >= 3:
                    extracted_data["serial_number"] = serial_candidate
                    st.write(f"✅ **SERIAL NUMBER FOUND**: {serial_candidate}")
                    break
        
        # Extract other data
        year_patterns = [r'(19|20)\d{2}.*?(lear|citation|phenom|model)']
        for pattern in year_patterns:
            match = re.search(pattern, pdf_text)
            if match:
                year_match = re.search(r'(19|20)\d{2}', match.group(0))
                if year_match:
                    extracted_data["year_model"] = int(year_match.group(0))
                    break
        
        # Total Hours - Improved to handle various formats
        total_hours_patterns = [
            r'(\d{1,2}[,\.]?\d{3})\s*airframe\s*hours\s*since\s*new',
            r'airframe\s*hours\s*since\s*new[:\s]*(\d{1,2}[,\.]?\d{3})',
            r'total\s+time\s+since\s+new[:\s]*(\d{1,2}[,\.]?\d{3})',
            r'(\d{1,2}[,\.]?\d{3})\s*hours\s*since\s*new',
            r'hours\s*since\s*new[:\s]*(\d{1,2}[,\.]?\d{3})',
            r'hours\s*/?\s*new[:\s]*(\d{1,2}[,\.]?\d{3})',
            r'ttsn[:\s]*(\d{1,2}[,\.]?\d{3})',
            r'total\s+hours[:\s]*(\d{1,2}[,\.]?\d{3})'
        ]
        
        for pattern in total_hours_patterns:
            match = re.search(pattern, pdf_text, re.IGNORECASE)
            if match:
                hours_str = match.group(1).replace(",", "").replace(".", "")
                hours_value = int(hours_str)
                extracted_data["total_hours"] = hours_value
                st.write(f"✅ **TOTAL HOURS FOUND**: {hours_value} (pattern: {pattern})")
                break
        
        # Engine Overhaul - Handle format like "985/1227 Engine Hours Since Overhaul"
        engine_patterns = [
            r'(\d{3,4})/(\d{1,2}[,\.]?\d{3})\s*engine\s*hours\s*since\s*overhaul',  # Matches 985/1227 pattern
            r'engine\s*hours\s*since\s*overhaul[:\s]*(\d{3,4})/(\d{1,2}[,\.]?\d{3})',
            r'(\d{1,2}[,\.]?\d{3})\s*hours\s*since\s*overhaul',
            r'hours\s*since\s*overhaul[:\s]*(\d{1,2}[,\.]?\d{3})',
            r'time\s*since\s*overhaul[:\s]*(\d{1,2}[,\.]?\d{3})',
            r'overhaul\s*hours[:\s]*(\d{1,2}[,\.]?\d{3})',
            r'engines?\s*[\s\S]{0,200}?(\d{1,2}[,\.]?\d{3})\s*hours',
            r'engine\s+time\s+since\s+overhaul[:\s]*(\d{1,2}[,\.]?\d{3})',
            r'tsoh[:\s]*(\d{1,2}[,\.]?\d{3})'
        ]
        
        engine_found = False
        
        # Check for the special format first (985/1227)
        for pattern in engine_patterns[:2]:
            match = re.search(pattern, pdf_text, re.IGNORECASE)
            if match:
                # For patterns with two groups, take the second (larger) number
                if match.lastindex == 2:
                    hours_str = match.group(2).replace(",", "").replace(".", "")
                    hours_value = int(hours_str)
                    extracted_data["engine_overhaul"] = hours_value
                    st.write(f"✅ **ENGINE OVERHAUL FOUND (from dual format)**: {hours_value}")
                    engine_found = True
                    break
        
        # If not found in dual format, try other patterns
        if not engine_found:
            # Try to find an ENGINES section
            engines_section_match = re.search(r'engines?\s*[:\-\s]*([\s\S]{0,500}?)(?=\n\n|\navionics|\ninterior|\nexterior|$)', pdf_text, re.IGNORECASE)
            
            if engines_section_match:
                engines_text = engines_section_match.group(1)
                # Look for hours in the engines section
                hours_match = re.search(r'(\d{1,2}[,\.]?\d{3})\s*hours', engines_text, re.IGNORECASE)
                if hours_match:
                    hours_str = hours_match.group(1).replace(",", "").replace(".", "")
                    hours_value = int(hours_str)
                    extracted_data["engine_overhaul"] = hours_value
                    st.write(f"✅ **ENGINE OVERHAUL FOUND IN ENGINES SECTION**: {hours_value}")
                    engine_found = True
            
            # Try remaining patterns
            if not engine_found:
                for pattern in engine_patterns[2:]:
                    match = re.search(pattern, pdf_text, re.IGNORECASE)
                    if match:
                        hours_str = match.group(1).replace(",", "").replace(".", "")
                        hours_value = int(hours_str)
                        extracted_data["engine_overhaul"] = hours_value
                        st.write(f"✅ **ENGINE OVERHAUL FOUND**: {hours_value}")
                        engine_found = True
                        break
        
        # If no engine overhaul found, default to total hours
        if not engine_found and "total_hours" in extracted_data:
            extracted_data["engine_overhaul"] = extracted_data["total_hours"]
            st.write(f"⚠️ **ENGINE OVERHAUL defaulted to total hours**: {extracted_data['total_hours']}")
        
        # Number of Seats
        seat_patterns = [r'(\w+)\s+\(\d+\)\s+passenger', r'(\d+)\s+passengers?']
        number_mappings = {"one": "1", "two": "2", "three": "3", "four": "4", "five": "5", "six": "6", "seven": "7", "eight": "8", "nine": "9", "ten": "10"}
        
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
                
                match_context = pdf_text[max(0, match.start()-200):match.end()+200]
                if re.search(r'lavatory|lav\b|belted\s+lav', match_context, re.IGNORECASE):
                    result = f"{digit} SEATS + BLTD LAV"
                else:
                    result = f"{digit} SEATS"
                
                extracted_data["number_of_seats"] = result
                break
        
        # Seat Configuration - NEW LOGIC
        seat_config_patterns = [
            r'forward.*?two.*?place.*?divan',
            r'center.*?4.*?place.*?club',
            r'two.*?individual.*?forward.*?facing.*?seats',
            r'belted.*?lav'
        ]
        
        seat_config_parts = []
        
        # Look for specific seating configurations
        if re.search(r'forward.*?two.*?place.*?divan', pdf_text, re.IGNORECASE):
            seat_config_parts.append("FWD 2 PLC DIV")
        
        if re.search(r'(center|mid).*?4.*?place.*?club', pdf_text, re.IGNORECASE):
            seat_config_parts.append("MID 4 PLC CLB")
        
        if re.search(r'two.*?individual.*?forward.*?facing.*?seats', pdf_text, re.IGNORECASE):
            seat_config_parts.append("AFT DUAL FWD FCING SEATS")
        elif re.search(r'individual.*?seats', pdf_text, re.IGNORECASE):
            seat_config_parts.append("INDIV SEATS")
        
        if re.search(r'belted.*?lav', pdf_text, re.IGNORECASE):
            seat_config_parts.append("BLTD LAV")
        
        if seat_config_parts:
            extracted_data["seat_configuration"] = ",".join(seat_config_parts)
            st.write(f"✅ **SEAT CONFIGURATION FOUND**: {extracted_data['seat_configuration']}")
        
        # Paint Exterior Year - IMPROVED LOGIC
        paint_patterns = [
            r'painted\s+in\s+(20\d{2})',
            r'paint.*?(20\d{2})',
            r'exterior.*?paint.*?(20\d{2})',
            r'(20\d{2}).*?exterior.*?paint'
        ]
        
        for pattern in paint_patterns:
            match = re.search(pattern, pdf_text, re.IGNORECASE)
            if match:
                year = match.group(1)
                extracted_data["paint_exterior_year"] = year
                st.write(f"✅ **PAINT EXTERIOR YEAR FOUND**: {year}")
                break
        
        # Interior Year - IMPROVED LOGIC
        interior_patterns = [
            r'interior.*?refurb.*?(20\d{2})',
            r'interior.*?(20\d{2})',
            r'(20\d{2}).*?interior.*?refurb',
            r'new.*?interior.*?(20\d{2})'
        ]
        
        for pattern in interior_patterns:
            match = re.search(pattern, pdf_text, re.IGNORECASE)
            if match:
                year = match.group(1)
                extracted_data["interior_year"] = year
                st.write(f"✅ **INTERIOR YEAR FOUND**: {year}")
                break
        
        # Engine Programs
        if "jssi" in pdf_text:
            extracted_data["engine_program"] = "JSSI"
        elif "msp gold" in pdf_text:
            extracted_data["engine_program"] = "MSP GOLD"
        

        
        # APU Program - Always include this field as Y/N
        apu_keywords = ["apu", "auxiliary power unit"]
        apu_found = False
        
        for keyword in apu_keywords:
            if keyword in pdf_text:
                apu_found = True
                break
        
        # Always set APU field - Y if found, N if not
        if "apu_program" in row_mappings:
            if apu_found and "jssi" in pdf_text:
                extracted_data["apu_program"] = "JSSI"
            elif apu_found and "msp" in pdf_text:
                extracted_data["apu_program"] = "MSP"
            else:
                extracted_data["apu_program"] = "NONE"
        
        # Upgrades
        upgrades = config.get("upgrades", {})
        
        # Check for APU as a Y/N upgrade field
        for upgrade_name, upgrade_config in upgrades.items():
            if "APU" in upgrade_name.upper():
                # This is an APU Y/N field
                extracted_data[f"upgrade_{upgrade_name}"] = "Y" if apu_found else "N"
                st.write(f"✅ **APU Y/N FIELD**: {extracted_data[f'upgrade_{upgrade_name}']}")
        
        # Process all other upgrades
        for upgrade_name, upgrade_config in upgrades.items():
            if not upgrade_name.startswith("upgrade_"):  # Skip if already processed
                upgrade_key = f"upgrade_{upgrade_name}"
                if upgrade_key not in extracted_data:  # Only process if not already set
                    keywords = upgrade_config.get("keywords", [])
                    found = False
                    
                    # Check each keyword
                    for keyword in keywords:
                        if keyword and re.search(re.escape(keyword), pdf_text, re.IGNORECASE):
                            found = True
                            st.write(f"✅ **UPGRADE MATCH**: {upgrade_name} found with keyword '{keyword}'")
                            break
                    
                    extracted_data[upgrade_key] = "Y" if found else "N"
        
        return extracted_data
    
    def find_broker_row(self, ws):
        for row in range(1, min(50, ws.max_row + 1)):
            label_cell = ws.cell(row=row, column=12).value
            if label_cell and isinstance(label_cell, str):
                if "broker" in label_cell.lower():
                    st.write(f"✅ **Found BROKER row**: {row} ('{label_cell}')")
                    return row
        st.write("⚠️ **BROKER row not found, using default row 5**")
        return 5
    
    def shift_formulas_in_cell(self, cell, shift_amount):
        if not cell.value or not isinstance(cell.value, str) or not cell.value.startswith('='):
            return
        
        formula = cell.value
        
        def shift_column_ref(match):
            col_ref = match.group(1)
            row_ref = match.group(2)
            
            col_num = 0
            for char in col_ref:
                col_num = col_num * 26 + (ord(char) - ord('A') + 1)
            
            new_col_num = col_num + shift_amount
            
            new_col_ref = ""
            while new_col_num > 0:
                new_col_num -= 1
                new_col_ref = chr(new_col_num % 26 + ord('A')) + new_col_ref
                new_col_num //= 26
            
            return f"{new_col_ref}{row_ref}"
        
        pattern = r'([A-Z]+)(\d+)'
        new_formula = re.sub(pattern, shift_column_ref, formula)
        
        if new_formula != formula:
            cell.value = new_formula
    
    def find_insertion_point(self, excel_file, serial_number):
        try:
            excel_content = excel_file.read()
            excel_file.seek(0)
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(excel_content)
                tmp_path = tmp.name
            
            wb = load_workbook(tmp_path)
            
            try:
                if '-' in serial_number:
                    new_serial_num = int(serial_number.split('-')[-1])
                    display_serial = serial_number.split('-')[-1]
                else:
                    new_serial_num = int(serial_number[-4:]) if len(serial_number) >= 4 else int(serial_number)
                    display_serial = serial_number[-4:] if len(serial_number) >= 4 else serial_number
            except:
                digits = ''.join(filter(str.isdigit, serial_number))
                new_serial_num = int(digits[-4:])
                display_serial = digits[-4:]
            
            st.write(f"🔍 **Looking for insertion point for serial**: {new_serial_num} (display: {display_serial})")
            
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                
                serial_positions = []
                for col in range(16, ws.max_column + 1, 2):
                    cell_value = ws.cell(row=1, column=col).value
                    if cell_value and str(cell_value).strip():
                        try:
                            serial_str = str(cell_value).strip()
                            if serial_str.isdigit():
                                existing_serial = int(serial_str)
                            else:
                                if '-' in serial_str:
                                    existing_serial = int(serial_str.split('-')[-1])
                                else:
                                    digits = ''.join(filter(str.isdigit, serial_str))
                                    existing_serial = int(digits[-4:]) if len(digits) >= 4 else int(digits)
                            
                            serial_positions.append({
                                'column': col,
                                'serial': existing_serial,
                                'original': serial_str
                            })
                        except:
                            continue
                
                serial_positions.sort(key=lambda x: x['serial'])
                
                insert_col = None
                for pos in serial_positions:
                    if new_serial_num < pos['serial']:
                        insert_col = pos['column']
                        break
                
                if insert_col is None and serial_positions:
                    insert_col = serial_positions[-1]['column'] + 2
                elif insert_col is None:
                    insert_col = 16
                
                st.write(f"✅ **Insertion point found**: Column {insert_col}")
                st.write(f"📋 **Current serials**: {[p['original'] for p in serial_positions]}")
                
                os.unlink(tmp_path)
                return {
                    'column': insert_col,
                    'sheet': sheet_name,
                    'serial_positions': serial_positions,
                    'display_serial': display_serial
                }
            
            os.unlink(tmp_path)
            return None
            
        except Exception as e:
            st.error(f"Error finding insertion point: {e}")
            return None
    
    def find_broker_column(self, excel_file, serial_number, broker_name):
        try:
            excel_content = excel_file.read()
            excel_file.seek(0)
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(excel_content)
                tmp_path = tmp.name
            
            wb = load_workbook(tmp_path)
            
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                
                for col in range(16, ws.max_column + 1, 2):
                    cell_value = ws.cell(row=1, column=col).value
                    if cell_value and str(cell_value).strip() == serial_number:
                        st.write(f"✅ **Serial number {serial_number} found in existing column {col}**")
                        os.unlink(tmp_path)
                        return {
                            'column': col,
                            'sheet': sheet_name,
                            'broker': broker_name,
                            'matched_serial': serial_number,
                            'mode': 'update'
                        }
                
                insertion_info = self.find_insertion_point(excel_file, serial_number)
                if insertion_info:
                    insertion_info['broker'] = broker_name
                    insertion_info['matched_serial'] = serial_number
                    insertion_info['mode'] = 'insert'
                    os.unlink(tmp_path)
                    return insertion_info
            
            os.unlink(tmp_path)
            return None
            
        except Exception as e:
            st.error(f"Error finding broker column: {e}")
            return None
    
    def insert_new_row(self, excel_file, extracted_data, aircraft_model, insertion_info, serial_number):
        try:
            excel_file.seek(0)
            excel_content = excel_file.read()
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(excel_content)
                tmp_path = tmp.name
            
            backup_path = tmp_path.replace(".xlsx", "_BACKUP.xlsx")
            shutil.copy(tmp_path, backup_path)
            
            wb = load_workbook(tmp_path)
            ws = wb[insertion_info['sheet']]
            target_col = insertion_info['column']
            
            config = st.session_state.configurations[aircraft_model]
            row_mappings = config.get("row_mappings", {})
            updates = []
            
            st.write("🔍 **Row Insertion Debug**: Starting new row insertion")
            st.write(f"🔍 **Target column**: {target_col}")
            
            # Step 1: Shift existing columns with formula updating
            if insertion_info['serial_positions']:
                columns_to_shift = []
                for pos in insertion_info['serial_positions']:
                    if pos['column'] >= target_col:
                        columns_to_shift.append(pos['column'])
                
                columns_to_shift.sort(reverse=True)
                
                for source_col in columns_to_shift:
                    dest_col = source_col + 2
                    st.write(f"📋 **Shifting column {source_col} to {dest_col}**")
                    
                    for row in range(1, ws.max_row + 1):
                        source_cell = ws.cell(row=row, column=source_col)
                        dest_cell = ws.cell(row=row, column=dest_col)
                        
                        if source_cell.value is not None:
                            dest_cell.value = source_cell.value
                            
                            if isinstance(source_cell.value, str) and source_cell.value.startswith('='):
                                self.shift_formulas_in_cell(dest_cell, 2)
                            
                            try:
                                if source_cell.font:
                                    dest_cell.font = Font(
                                        name=source_cell.font.name,
                                        size=source_cell.font.size,
                                        bold=source_cell.font.bold,
                                        italic=source_cell.font.italic,
                                        color=source_cell.font.color
                                    )
                            except:
                                pass
                        
                        source_yn_cell = ws.cell(row=row, column=source_col + 1)
                        dest_yn_cell = ws.cell(row=row, column=dest_col + 1)
                        
                        if source_yn_cell.value is not None:
                            dest_yn_cell.value = source_yn_cell.value
                            
                            if isinstance(source_yn_cell.value, str) and source_yn_cell.value.startswith('='):
                                self.shift_formulas_in_cell(dest_yn_cell, 2)
                            
                            try:
                                if source_yn_cell.font:
                                    dest_yn_cell.font = Font(
                                        name=source_yn_cell.font.name,
                                        size=source_yn_cell.font.size,
                                        bold=source_yn_cell.font.bold,
                                        italic=source_yn_cell.font.italic,
                                        color=source_yn_cell.font.color
                                    )
                            except:
                                pass
                        
                        source_cell.value = None
                        source_yn_cell.value = None
            
            # Step 2: Add serial number (last 4 digits only) and broker name
            display_serial = insertion_info.get('display_serial', serial_number[-4:])
            
            ws.cell(row=1, column=target_col).value = display_serial
            updates.append(f"Serial Number - Row 1: {display_serial}")
            
            broker_row = self.find_broker_row(ws)
            if "broker_name" in extracted_data:
                broker_cell = ws.cell(row=broker_row, column=target_col)
                broker_cell.value = extracted_data["broker_name"].upper()  # Convert to uppercase
                updates.append(f"Broker Name - Row {broker_row}: {extracted_data['broker_name'].upper()}")
            else:
                broker_cell = ws.cell(row=broker_row, column=target_col)
                broker_cell.value = insertion_info.get('broker', 'Unknown Broker').upper()  # Convert to uppercase
                updates.append(f"Broker Name - Row {broker_row}: {insertion_info.get('broker', 'Unknown Broker').upper()}")
            
            # Step 3: Add extracted data
            for field, row_num in row_mappings.items():
                if field in extracted_data and row_num != 1 and row_num != broker_row:
                    new_value = extracted_data[field]
                    ws.cell(row=row_num, column=target_col).value = new_value
                    updates.append(f"{field} - Row {row_num}: {new_value}")
            
            # Step 4: Add upgrade data
            upgrades = config.get("upgrades", {})
            for upgrade_name, upgrade_config in upgrades.items():
                upgrade_key = f"upgrade_{upgrade_name}"
                if upgrade_key in extracted_data:
                    row_num = upgrade_config.get("row")
                    if row_num and row_num != 1:
                        yn_col = target_col + 1
                        cell = ws.cell(row=row_num, column=yn_col)
                        cell.value = extracted_data[upgrade_key]
                        try:
                            cell.font = Font(name="Calibri", size=11, color="FFFFFF")
                        except:
                            cell.font = Font(name="Calibri", size=11)
                        updates.append(f"{upgrade_name} - Row {row_num}: {extracted_data[upgrade_key]}")
            
            try:
                wb.calculation.calcMode = "automatic"
                st.write("✅ **Set calculation mode to automatic**")
            except:
                st.write("⚠️ **Could not set calculation mode**")
            
            wb.save(tmp_path)
            st.write("✅ **New row inserted successfully**")
            
            with open(tmp_path, "rb") as f:
                updated_excel_data = f.read()
            
            with open(backup_path, "rb") as f:
                backup_excel_data = f.read()
            
            os.unlink(tmp_path)
            os.unlink(backup_path)
            
            return updated_excel_data, backup_excel_data, updates
            
        except Exception as e:
            st.error(f"Error inserting new row: {e}")
            return None, None, []
    
    def update_excel(self, excel_file, extracted_data, aircraft_model, broker_info):
        try:
            # Add broker name to extracted data
            extracted_data["broker_name"] = broker_info['broker']
            
            if broker_info['mode'] == 'insert':
                return self.insert_new_row(excel_file, extracted_data, aircraft_model, broker_info, broker_info['matched_serial'])
            else:
                return self.update_existing_row(excel_file, extracted_data, aircraft_model, broker_info)
        except Exception as e:
            st.error(f"Error in update_excel: {e}")
            return None, None, []
    
    def update_existing_row(self, excel_file, extracted_data, aircraft_model, broker_info):
        try:
            excel_file.seek(0)
            excel_content = excel_file.read()
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(excel_content)
                tmp_path = tmp.name
            
            backup_path = tmp_path.replace(".xlsx", "_BACKUP.xlsx")
            shutil.copy(tmp_path, backup_path)
            
            wb = load_workbook(tmp_path)
            ws = wb[broker_info['sheet']]
            target_col = broker_info['column']
            
            config = st.session_state.configurations[aircraft_model]
            row_mappings = config.get("row_mappings", {})
            updates = []
            
            # Update broker name (convert to uppercase)
            broker_row = self.find_broker_row(ws)
            if "broker_name" in extracted_data:
                broker_cell = ws.cell(row=broker_row, column=target_col)
                broker_cell.value = extracted_data["broker_name"].upper()
                updates.append(f"Broker Name - Row {broker_row}: {extracted_data['broker_name'].upper()}")
            
            for field, row_num in row_mappings.items():
                if field in extracted_data:
                    if row_num == 1:
                        st.error(f"🚨 **PROTECTION**: Refusing to update Row 1 ({field}) - this is the serial number row!")
                        continue
                    
                    current_value = ws.cell(row=row_num, column=target_col).value
                    new_value = extracted_data[field]
                    
                    ws.cell(row=row_num, column=target_col).value = new_value
                    updates.append(f"{field} - Row {row_num}: {extracted_data[field]}")
            
            upgrades = config.get("upgrades", {})
            for upgrade_name, upgrade_config in upgrades.items():
                upgrade_key = f"upgrade_{upgrade_name}"
                if upgrade_key in extracted_data:
                    row_num = upgrade_config.get("row")
                    if row_num and row_num != 1:
                        yn_col = target_col + 1
                        cell = ws.cell(row=row_num, column=yn_col)
                        cell.value = extracted_data[upgrade_key]
                        try:
                            cell.font = Font(name="Calibri", size=11, color="FFFFFF")
                        except:
                            cell.font = Font(name="Calibri", size=11)
                        updates.append(f"{upgrade_name} - Row {row_num}: {extracted_data[upgrade_key]}")
            
            wb.save(tmp_path)
            
            with open(tmp_path, "rb") as f:
                updated_excel_data = f.read()
            
            with open(backup_path, "rb") as f:
                backup_excel_data = f.read()
            
            os.unlink(tmp_path)
            os.unlink(backup_path)
            
            return updated_excel_data, backup_excel_data, updates
            
        except Exception as e:
            st.error(f"Error updating existing row: {e}")
            return None, None, []

def main():
    st.set_page_config(page_title="Aircraft Data Platform", page_icon="✈️", layout="wide")
    
    platform = CompletePlatform()
    
    st.title("✈️ Complete Aircraft Data Management Platform")
    st.markdown("**Configure aircraft models and process broker data**")
    
    # Sidebar
    st.sidebar.title("📊 Status")
    
    if st.session_state.get('configurations'):
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
        
        if st.session_state.get('configurations'):
            st.markdown("---")
            st.subheader("✈️ Process Aircraft Data")
            
            # Choose between single or multiple PDF mode
            process_mode = st.radio("Processing Mode:", ["Single PDF", "Multiple PDFs"], horizontal=True)
            
            if process_mode == "Single PDF":
                col1, col2 = st.columns(2)
                with col1:
                    serial_number = st.text_input("Serial Number:", placeholder="e.g., 0028", key="single_serial")
                with col2:
                    broker_name = st.text_input("Broker Name:", placeholder="e.g., FlyAlliance", key="single_broker")
                
                col1, col2 = st.columns(2)
                with col1:
                    pdf_file = st.file_uploader("Upload Broker PDF", type="pdf", key="single_pdf")
                with col2:
                    excel_file = st.file_uploader("Upload Excel Sheet", type="xlsx", key="single_excel")
                
                if pdf_file:
                    pdf_details = [{"pdf": pdf_file, "serial": serial_number, "broker": broker_name}]
                else:
                    pdf_details = []
            
            else:  # Multiple PDFs mode
                excel_file = st.file_uploader("Upload Excel Sheet", type="xlsx", key="multi_excel")
                
                st.write("### Add PDF Details")
                
                # Initialize session state for PDF details if not exists
                if 'pdf_entries' not in st.session_state:
                    st.session_state.pdf_entries = []
                
                # Add new PDF entry
                with st.expander("➕ Add New PDF Entry", expanded=True):
                    col1, col2, col3 = st.columns([2, 1, 1])
                    with col1:
                        new_pdf = st.file_uploader("PDF File:", type="pdf", key="new_pdf")
                    with col2:
                        new_serial = st.text_input("Serial Number:", key="new_serial")
                    with col3:
                        new_broker = st.text_input("Broker Name:", key="new_broker")
                    
                    if st.button("➕ Add PDF", type="secondary"):
                        if new_pdf and new_serial and new_broker:
                            st.session_state.pdf_entries.append({
                                "pdf": new_pdf,
                                "serial": new_serial,
                                "broker": new_broker,
                                "name": new_pdf.name
                            })
                            st.success(f"Added {new_pdf.name}")
                            st.rerun()
                        else:
                            st.error("Please fill all fields before adding")
                
                # Display current PDF entries
                if st.session_state.pdf_entries:
                    st.write("### Current PDF Queue:")
                    for idx, entry in enumerate(st.session_state.pdf_entries):
                        col1, col2, col3, col4 = st.columns([3, 2, 2, 1])
                        with col1:
                            st.write(f"📄 {entry['name']}")
                        with col2:
                            st.write(f"Serial: {entry['serial']}")
                        with col3:
                            st.write(f"Broker: {entry['broker']}")
                        with col4:
                            if st.button("❌", key=f"remove_{idx}"):
                                st.session_state.pdf_entries.pop(idx)
                                st.rerun()
                    
                    # Clear all button
                    if st.button("🗑️ Clear All PDFs"):
                        st.session_state.pdf_entries = []
                        st.rerun()
                
                pdf_details = st.session_state.pdf_entries
            
            if pdf_details and excel_file and all(d["serial"] and d["broker"] for d in pdf_details):
                if st.button("🚀 Process Aircraft Data", type="primary", key="process_btn"):
                    results = []
                    current_excel = excel_file
                    
                    for idx, detail in enumerate(pdf_details):
                        st.write(f"\n### Processing PDF {idx + 1} of {len(pdf_details)}: {detail.get('name', detail['pdf'].name)}")
                        
                        with st.spinner(f"Processing {detail.get('name', detail['pdf'].name)}..."):
                            pdf_text = platform.extract_text_from_pdf(detail['pdf'])
                            
                            if not pdf_text:
                                st.error(f"Could not extract text from {detail.get('name', detail['pdf'].name)}")
                                continue
                            
                            aircraft_model = platform.identify_aircraft_from_pdf(pdf_text)
                            
                            if not aircraft_model:
                                st.error(f"❌ Could not identify aircraft model in {detail.get('name', detail['pdf'].name)}")
                                st.info("Available: " + ", ".join(st.session_state.configurations.keys()))
                                continue
                            
                            st.success(f"✅ Identified: **{aircraft_model}**")
                            
                            extracted_data = platform.extract_data_from_pdf(pdf_text, aircraft_model)
                            
                            if extracted_data:
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
                                
                                broker_info = platform.find_broker_column(current_excel, detail["serial"], detail["broker"])
                                
                                if not broker_info:
                                    st.error(f"❌ Could not find broker column or insertion point for {detail['serial']}")
                                    continue
                                
                                if broker_info['mode'] == 'update':
                                    st.success(f"✅ Found existing entry in Column {broker_info['column']} - will update")
                                else:
                                    st.success(f"✅ Will insert new row at Column {broker_info['column']}")
                                
                                updated_excel, backup_excel, updates = platform.update_excel(
                                    current_excel, extracted_data, aircraft_model, broker_info
                                )
                                
                                if updated_excel:
                                    mode_text = "updated" if broker_info['mode'] == 'update' else "inserted"
                                    st.success(f"✅ Excel {mode_text} successfully for {detail['serial']}!")
                                    
                                    # Save the result
                                    results.append({
                                        "serial": detail["serial"],
                                        "broker": detail["broker"],
                                        "pdf_name": detail.get('name', detail['pdf'].name),
                                        "updates": updates,
                                        "mode": broker_info['mode']
                                    })
                                    
                                    # Use the updated Excel for the next iteration
                                    if idx < len(pdf_details) - 1:
                                        # Create a file-like object for the next iteration
                                        current_excel = io.BytesIO(updated_excel)
                                        current_excel.name = "temp.xlsx"
                                else:
                                    st.error(f"❌ Failed to update Excel for {detail['serial']}")
                    
                    # Show summary and download
                    if results:
                        st.write("\n## 📊 Processing Summary")
                        for result in results:
                            st.write(f"• **{result['serial']}** ({result['broker']}): {result['mode']} - {len(result['updates'])} fields updated")
                        
                        st.write("\n## 📥 Download Results")
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.download_button(
                                "📥 Download Updated Excel",
                                updated_excel,
                                f"UPDATED_MULTIPLE_{aircraft_model.replace(' ', '_')}.xlsx",
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        
                        with col2:
                            st.download_button(
                                "💾 Download Original Backup",
                                backup_excel,
                                f"ORIGINAL_BACKUP_{aircraft_model.replace(' ', '_')}.xlsx",
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    
                    # Clear PDF entries after successful processing
                    if process_mode == "Multiple PDFs" and results:
                        st.session_state.pdf_entries = []
            else:
                if not excel_file:
                    st.info("👆 Please upload an Excel file")
                elif not pdf_details:
                    st.info("👆 Please add at least one PDF with details")
                else:
                    st.info("👆 Please fill in all serial numbers and broker names")
        else:
            st.info("👆 Please load a configuration file first to enable data processing")
    
    with tab2:
        st.header("🔧 Create New Aircraft Model Configuration")
        
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
                    
                    # Clear temporary custom upgrades
                    if 'temp_custom_upgrades' in st.session_state:
                        del st.session_state.temp_custom_upgrades
                    
                    config_json = json.dumps(config, indent=2)
                    st.download_button(
                        "📥 Download Configuration File",
                        config_json,
                        f"{aircraft_model.replace(' ', '_')}_config.json",
                        "application/json",
                        key="download_new_config"
                    )
                    st.info("💡 **Important:** Download and save this configuration file! You'll need it for the Quick Process tab.")
                    if 'custom_upgrades' in st.session_state:
                        st.session_state.custom_upgrades = []

if __name__ == "__main__":
    main()
