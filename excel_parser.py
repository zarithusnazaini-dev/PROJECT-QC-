import os
import re
import openpyxl

def detect_excel_format(file_path_or_bytes):
    """
    Detect whether the Excel file is Format Baru (New) or Format Lama (Old).
    Returns 'new' or 'old'.
    """
    import io
    if isinstance(file_path_or_bytes, (bytes, bytearray)):
        file_path_or_bytes = io.BytesIO(file_path_or_bytes)
    wb = openpyxl.load_workbook(file_path_or_bytes, data_only=True)
    all_text = ""
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=True):
            row_str = " ".join([str(val) for val in row if val is not None]).upper()
            all_text += " " + row_str
            
    wb.close()
    
    if "MODULE PANEL QUANTITY" in all_text or "SECTION 1 : PERSONAL INFORMATION" in all_text or "SECTION 3 : TNB DETAILS" in all_text:
        return "new"
    elif "SECTION 1:  PROJECT INFORMATION" in all_text or "APPLY CCC FOR" in all_text or "DOC. NO" in all_text or "CLIENT/COMPANY NAME" in all_text:
        return "old"
    
    if "PERSONAL INFORMATION" in all_text:
        return "new"
    return "old"

def parse_new_format(file_path_or_bytes):
    """
    Extract data from New Project Form Template.
    """
    import io
    if isinstance(file_path_or_bytes, (bytes, bytearray)):
        file_path_or_bytes = io.BytesIO(file_path_or_bytes)
    wb = openpyxl.load_workbook(file_path_or_bytes, data_only=True)
    sheet = wb.active
    
    data = {
        "format": "Format Baru",
        "client_name": "",
        "ic": "",
        "address": "",
        "postcode": "",
        "town": "",
        "state": "",
        "mobile": "",
        "email": "",
        "kwac": 0.0,
        "kwp": 0.0,
        "module_brand": "",
        "module_type": "",
        "module_model": "",
        "module_power_w": 0,
        "panel_qty": 0,
        "inverter_brand": "",
        "inverter_model": "",
        "inverter_kwac": 0.0,
        "inverter_qty": 0,
        "battery_brand": "",
        "battery_model": "",
        "battery_ah": 0,
        "battery_qty": 0,
        "roof_type": "",
        "tnb_meter": "",
        "wireman_name": "",
        "raw_text": ""
    }
    
    kv_map = {}
    for row in sheet.iter_rows(values_only=True):
        clean_row = [str(c).strip() for c in row if c is not None and str(c).strip() != ""]
        if clean_row and len(clean_row) >= 2:
            key = re.sub(r"\s+", " ", clean_row[0].upper().strip())
            val = clean_row[1]
            kv_map[key] = val
            # Also handle key-value pairs that might be in column B and C
            if len(clean_row) >= 4:
                k2 = re.sub(r"\s+", " ", clean_row[2].upper().strip())
                v2 = clean_row[3]
                kv_map[k2] = v2
                
    # Client Name
    data["client_name"] = str(kv_map.get("NAME", "")).strip()
    data["ic"] = str(kv_map.get("IC", "")).strip()
    data["email"] = str(kv_map.get("EMAIL", "")).strip()
    data["mobile"] = str(kv_map.get("MOBILE NO", kv_map.get("MOBILE NUMBER", ""))).strip()
    
    # Address
    address_premis = str(kv_map.get("ADDRESS (PREMIS)", kv_map.get("ADDRESS", ""))).strip()
    postcode = str(kv_map.get("POSTCODE", "")).strip()
    town = str(kv_map.get("TOWN", "")).strip()
    state = str(kv_map.get("STATE", "")).strip()
    data["address"] = f"{address_premis}, {postcode} {town}, {state}".strip(", ")
    data["postcode"] = postcode
    data["town"] = town
    data["state"] = state
    
    # Capacity
    try:
        data["kwac"] = float(kv_map.get("KWAC", 0) or 0)
    except:
        data["kwac"] = 0.0
        
    try:
        data["kwp"] = float(kv_map.get("KWP", 0) or 0)
    except:
        data["kwp"] = 0.0
        
    # Module / Panel
    data["module_brand"] = str(kv_map.get("MODULE BRAND", "")).strip()
    data["module_type"] = str(kv_map.get("MODULE TYPE", "")).strip()
    data["module_model"] = str(kv_map.get("MODULE PANEL", kv_map.get("MODULE MODEL", ""))).strip()
    try:
        data["module_power_w"] = int(float(kv_map.get("MODULE POWER (W)", kv_map.get("MODULE POWER", 0)) or 0))
    except:
        data["module_power_w"] = 0
        
    try:
        data["panel_qty"] = int(float(kv_map.get("MODULE PANEL QUANTITY", kv_map.get("PANEL QUANTITY", kv_map.get("PANEL QTY", 0))) or 0))
    except:
        data["panel_qty"] = 0
        
    # Inverter
    data["inverter_brand"] = str(kv_map.get("INVERTER BRAND", "")).strip()
    data["inverter_model"] = str(kv_map.get("INVERTER MODEL", "")).strip()
    try:
        data["inverter_kwac"] = float(kv_map.get("INVERTER KWAC", 0) or 0)
    except:
        data["inverter_kwac"] = 0.0
        
    try:
        data["inverter_qty"] = int(float(kv_map.get("INVERTER QUANTITY", 0) or 0))
    except:
        data["inverter_qty"] = 0
        
    # Battery
    data["battery_brand"] = str(kv_map.get("BATTERY BRAND", "")).strip()
    data["battery_model"] = str(kv_map.get("BATTERY  MODEL", kv_map.get("BATTERY MODEL", ""))).strip()
    try:
        data["battery_ah"] = int(float(kv_map.get("BATTERY AH", 0) or 0))
    except:
        data["battery_ah"] = 0
        
    try:
        data["battery_qty"] = int(float(kv_map.get("BATTERY QUANTITY", kv_map.get("BATTERY QTY", kv_map.get("BATTERY CAPACITY (16KWH)", 0))) or 0))
    except:
        data["battery_qty"] = 0
        
    # Office Use / Technical Specs
    data["roof_type"] = str(kv_map.get("ROOF TYPE", "")).strip()
    data["tnb_meter"] = str(kv_map.get("TNB METER", "")).strip()
    data["wireman_name"] = str(kv_map.get("WIREMAN NAME", "")).strip()
    
    wb.close()
    return data

def parse_old_format(file_path_or_bytes):
    """
    Extract data from Old Project Form Template.
    """
    import io
    if isinstance(file_path_or_bytes, (bytes, bytearray)):
        file_path_or_bytes = io.BytesIO(file_path_or_bytes)
    wb = openpyxl.load_workbook(file_path_or_bytes, data_only=True)
    sheet = wb.active
    
    data = {
        "format": "Format Lama",
        "client_name": "",
        "ic": "",
        "address": "",
        "postcode": "",
        "town": "",
        "state": "",
        "mobile": "",
        "email": "",
        "kwac": 0.0,
        "kwp": 0.0,
        "module_brand": "",
        "module_type": "",
        "module_model": "",
        "module_power_w": 0,
        "panel_qty": 0,
        "inverter_brand": "",
        "inverter_model": "",
        "inverter_kwac": 0.0,
        "inverter_qty": 1,
        "battery_brand": "",
        "battery_model": "",
        "battery_ah": 0,
        "battery_qty": 0,
        "roof_type": "",
        "tnb_meter": "",
        "wireman_name": "",
        "raw_text": ""
    }
    
    all_text_lines = []
    for row in sheet.iter_rows(values_only=True):
        line = " | ".join([str(c).strip() for c in row if c is not None and str(c).strip() != ""])
        if line:
            all_text_lines.append(line)
            
    full_text = "\n".join(all_text_lines)
    data["raw_text"] = full_text
    
    # 1. Client Name: e.g. Client/Company Name | CHU CHIH YONG or Liang Yee Mun
    m_name = re.search(r"(?:Client/Company Name|Contact Person\'s Name)[^\|\n]*\|\s*([^\|\n]+)", full_text, re.I)
    if m_name:
        data["client_name"] = m_name.group(1).strip()
            
    # 2. Address / Location
    m_addr = re.search(r"Site Location\s*/\s*GPS[^\|\n]*\|\s*([^\|\n]+)", full_text, re.I)
    if m_addr:
        data["address"] = m_addr.group(1).strip().strip('"')
            
    # 3. Roof Type
    m_roof = re.search(r"Installation Type[^\|\n]*\|\s*([^\|\n]+)", full_text, re.I)
    if m_roof:
        data["roof_type"] = m_roof.group(1).strip()
    elif "Roof Tiles" in full_text or "Roof Tile" in full_text:
        data["roof_type"] = "Roof Tile"
    elif "Metal Deck" in full_text:
        data["roof_type"] = "Metal Deck"
        
    # 4. Preferred Brands
    m_pvm = re.search(r"Prefered PV Module\'s Brand[^\|\n]*\|\s*([^\|\n]+)", full_text, re.I)
    if m_pvm:
        data["module_brand"] = m_pvm.group(1).strip()
        
    m_inv = re.search(r"Prefered Inverter\'s Brand[^\|\n]*\|\s*([^\|\n]+)", full_text, re.I)
    if m_inv:
        data["inverter_brand"] = m_inv.group(1).strip()
        
    # 5. TNB Meter / Phase
    if re.search(r"TNB Meter[^\|\n]*\|\s*Three Phase", full_text, re.I):
        data["tnb_meter"] = "Three Phase"
    elif re.search(r"TNB Meter[^\|\n]*\|\s*Single Phase", full_text, re.I):
        data["tnb_meter"] = "Single Phase"
    elif "0.4" in full_text and "kV" in full_text:
        data["tnb_meter"] = "Three Phase"
    elif "0.23" in full_text and "kV" in full_text:
        data["tnb_meter"] = "Single Phase"
        
    # 6. Capacity from Section 6 (Installed Capacity)
    m_cap = re.search(r"Installed Cap(?:ca|ac)ity\s*\(kWp\)[^\|\n]*\|\s*([\d\.]+)", full_text, re.I)
    if m_cap:
        try:
            data["kwp"] = float(m_cap.group(1))
        except:
            pass
            
    # 7. Search Remark or Entire Sheet for Panel Qty, Wattage, and kWp
    # Remark might say: "APPLY CCC FOR 25.4KWP 40 PANELS X 635W" or "14 PCS X 630W" or "8.82 kWp"
    m_remark = re.search(r"Remark[^\n]*\s*\|\s*([^\n]+)", full_text, re.I)
    search_scope = (m_remark.group(1).upper() if m_remark else "") + " " + full_text.upper()
    
    # Check for panel quantity:
    # Pattern A: 40 PANELS X 635W or 14 PCS X 630W or 14 X 630W
    m_pw = re.search(r"(\d+)\s*(?:PANELS?|PCS|NOS)?\s*(?:X|\*)\s*(\d{3,4})\s*W?", search_scope)
    if m_pw:
        data["panel_qty"] = int(m_pw.group(1))
        data["module_power_w"] = int(m_pw.group(2))
    else:
        # Pattern B: 40 PANELS or 14 PCS
        m_qty_only = re.search(r"(\d+)\s*(?:PANELS?|PCS|NOS)\b", search_scope)
        if m_qty_only and int(m_qty_only.group(1)) < 200:
            data["panel_qty"] = int(m_qty_only.group(1))
            
    # Check for kWp
    m_kwp = re.search(r"([\d\.]+)\s*KWP", search_scope)
    if m_kwp and data["kwp"] == 0.0:
        data["kwp"] = float(m_kwp.group(1))
        
    # If kWp and panel wattage known, deduce panel qty if still 0
    if data["panel_qty"] == 0 and data["kwp"] > 0 and data["module_power_w"] > 0:
        data["panel_qty"] = round((data["kwp"] * 1000.0) / data["module_power_w"])
    elif data["kwp"] == 0.0 and data["panel_qty"] > 0 and data["module_power_w"] > 0:
        data["kwp"] = round((data["panel_qty"] * data["module_power_w"]) / 1000.0, 2)
        
    wb.close()
    return data

def is_pdf_file(file_path_or_bytes, filename=None):
    """
    Checks if the input file or bytes is a PDF.
    """
    if filename and filename.lower().endswith(".pdf"):
        return True
    if isinstance(file_path_or_bytes, (bytes, bytearray)):
        return file_path_or_bytes.startswith(b"%PDF")
    if hasattr(file_path_or_bytes, "getvalue"):
        try:
            return file_path_or_bytes.getvalue().startswith(b"%PDF")
        except Exception:
            pass
    if hasattr(file_path_or_bytes, "read") and hasattr(file_path_or_bytes, "seek"):
        try:
            pos = file_path_or_bytes.tell()
            header = file_path_or_bytes.read(5)
            file_path_or_bytes.seek(pos)
            return header.startswith(b"%PDF")
        except Exception:
            pass
    if isinstance(file_path_or_bytes, str) and os.path.exists(file_path_or_bytes):
        return file_path_or_bytes.lower().endswith(".pdf")
    return False

def parse_project_pdf(file_path_or_bytes):
    """
    Extract data directly from a Project Form PDF (supports both Format Baru and Format Lama).
    Runs instantly (<0.05s) on Windows or Linux cloud environments.
    """
    import io
    import pdfplumber

    if isinstance(file_path_or_bytes, (bytes, bytearray)):
        pdf_source = io.BytesIO(file_path_or_bytes)
    elif hasattr(file_path_or_bytes, "getvalue"):
        pdf_source = io.BytesIO(file_path_or_bytes.getvalue())
    else:
        pdf_source = file_path_or_bytes

    with pdfplumber.open(pdf_source) as pdf:
        all_text = ""
        all_table_rows = []
        for p in pdf.pages:
            all_text += " " + (p.extract_text() or "")
            for tbl in p.extract_tables() or []:
                all_table_rows.extend(tbl)

    all_text_upper = all_text.upper()
    is_new = ("MODULE PANEL QUANTITY" in all_text_upper or 
              "SECTION 1 : PERSONAL INFORMATION" in all_text_upper or 
              "SECTION 3 : TNB DETAILS" in all_text_upper or
              "PERSONAL INFORMATION" in all_text_upper)

    if is_new:
        data = {
            "format": "Format Baru", "client_name": "", "ic": "", "address": "", "postcode": "", "town": "", "state": "",
            "mobile": "", "email": "", "kwac": 0.0, "kwp": 0.0, "module_brand": "", "module_type": "", "module_model": "",
            "module_power_w": 0, "panel_qty": 0, "inverter_brand": "", "inverter_model": "", "inverter_kwac": 0.0,
            "inverter_qty": 0, "battery_brand": "", "battery_model": "", "battery_ah": 0, "battery_qty": 0,
            "roof_type": "", "tnb_meter": "", "wireman_name": "", "raw_text": all_text
        }
        kv_map = {}
        curr_section = ""
        for row in all_table_rows:
            clean_row = [str(c).strip() for c in row if c is not None and str(c).strip() != ""]
            if not clean_row:
                continue
            
            if "SECTION" in clean_row[0].upper():
                curr_section = clean_row[0].upper()
                
            if len(clean_row) >= 2:
                key = re.sub(r"\s+", " ", clean_row[0].upper().strip())
                val = clean_row[1]
                if "SECTION 2" in curr_section:
                    kv_map["KIN_" + key] = val
                else:
                    if key not in kv_map:
                        kv_map[key] = val
                        
            if len(clean_row) >= 4:
                k2 = re.sub(r"\s+", " ", clean_row[2].upper().strip())
                v2 = clean_row[3]
                if "SECTION 2" in curr_section:
                    kv_map["KIN_" + k2] = v2
                elif k2 not in kv_map:
                    kv_map[k2] = v2
                    
        data["client_name"] = str(kv_map.get("NAME", "")).strip()
        data["ic"] = str(kv_map.get("IC", "")).strip()
        data["email"] = str(kv_map.get("EMAIL", "")).strip()
        data["mobile"] = str(kv_map.get("MOBILE NO", kv_map.get("MOBILE NUMBER", ""))).strip()
        
        address_premis = str(kv_map.get("ADDRESS (PREMIS)", kv_map.get("ADDRESS", ""))).strip()
        postcode = str(kv_map.get("POSTCODE", "")).strip()
        town = str(kv_map.get("TOWN", "")).strip()
        state = str(kv_map.get("STATE", "")).strip()
        data["address"] = f"{address_premis}, {postcode} {town}, {state}".strip(", ")
        data["postcode"] = postcode
        data["town"] = town
        data["state"] = state
        
        try: data["kwac"] = float(kv_map.get("KWAC", 0) or 0)
        except: data["kwac"] = 0.0
        try: data["kwp"] = float(kv_map.get("KWP", 0) or 0)
        except: data["kwp"] = 0.0
        
        data["module_brand"] = str(kv_map.get("MODULE BRAND", "")).strip()
        data["module_type"] = str(kv_map.get("MODULE TYPE", "")).strip()
        data["module_model"] = str(kv_map.get("MODULE PANEL", kv_map.get("MODULE MODEL", ""))).strip()
        try: data["module_power_w"] = int(float(kv_map.get("MODULE POWER (W)", kv_map.get("MODULE POWER", 0)) or 0))
        except: data["module_power_w"] = 0
        try: data["panel_qty"] = int(float(kv_map.get("MODULE PANEL QUANTITY", kv_map.get("PANEL QUANTITY", kv_map.get("PANEL QTY", 0))) or 0))
        except: data["panel_qty"] = 0
        
        data["inverter_brand"] = str(kv_map.get("INVERTER BRAND", "")).strip()
        data["inverter_model"] = str(kv_map.get("INVERTER MODEL", "")).strip()
        try: data["inverter_kwac"] = float(kv_map.get("INVERTER KWAC", 0) or 0)
        except: data["inverter_kwac"] = 0.0
        try: data["inverter_qty"] = int(float(kv_map.get("INVERTER QUANTITY", 0) or 0))
        except: data["inverter_qty"] = 0
        
        data["battery_brand"] = str(kv_map.get("BATTERY BRAND", "")).strip()
        data["battery_model"] = str(kv_map.get("BATTERY  MODEL", kv_map.get("BATTERY MODEL", ""))).strip()
        try: data["battery_ah"] = int(float(kv_map.get("BATTERY AH", 0) or 0))
        except: data["battery_ah"] = 0
        try: data["battery_qty"] = int(float(kv_map.get("BATTERY QUANTITY", kv_map.get("BATTERY QTY", kv_map.get("BATTERY CAPACITY (16KWH)", 0))) or 0))
        except: data["battery_qty"] = 0
        
        data["roof_type"] = str(kv_map.get("ROOF TYPE", "")).strip()
        data["tnb_meter"] = str(kv_map.get("TNB METER", "")).strip()
        data["wireman_name"] = str(kv_map.get("WIREMAN NAME", "")).strip()
        return data
    else:
        # Old format
        all_text_lines = []
        for row in all_table_rows:
            line = " | ".join([str(c).strip() for c in row if c is not None and str(c).strip() != ""])
            if line:
                all_text_lines.append(line)
        full_text = "\n".join(all_text_lines)
        data = {
            "format": "Format Lama", "client_name": "", "ic": "", "address": "", "postcode": "", "town": "", "state": "",
            "mobile": "", "email": "", "kwac": 0.0, "kwp": 0.0, "module_brand": "", "module_type": "", "module_model": "",
            "module_power_w": 0, "panel_qty": 0, "inverter_brand": "", "inverter_model": "", "inverter_kwac": 0.0,
            "inverter_qty": 1, "battery_brand": "", "battery_model": "", "battery_ah": 0, "battery_qty": 0,
            "roof_type": "", "tnb_meter": "", "wireman_name": "", "raw_text": full_text
        }
        m_name = re.search(r"(?:Client/Company Nam(?:e)?|Contact Person\'s Name)[^\|\n]*\|\s*(?:e\s+)?([^\|\n]+)", full_text, re.I)
        if m_name:
            data["client_name"] = m_name.group(1).strip()
        m_addr = re.search(r"Site Location\s*/\s*GPS[^\|\n]*\|\s*([^\|\n]+)", full_text, re.I)
        if m_addr:
            data["address"] = m_addr.group(1).strip().strip('"')
        m_roof = re.search(r"Installation Type[^\|\n]*\|\s*([^\|\n]+)", full_text, re.I)
        if m_roof:
            data["roof_type"] = m_roof.group(1).strip()
        elif "Roof Tiles" in full_text or "Roof Tile" in full_text:
            data["roof_type"] = "Roof Tile"
        elif "Metal Deck" in full_text:
            data["roof_type"] = "Metal Deck"
            
        m_pvm = re.search(r"Prefered PV Module\'s Brand[^\|\n]*\|\s*([^\|\n]+)", full_text, re.I)
        if m_pvm:
            data["module_brand"] = m_pvm.group(1).strip()
        m_inv = re.search(r"Prefered Inverter\'s Brand[^\|\n]*\|\s*([^\|\n]+)", full_text, re.I)
        if m_inv:
            data["inverter_brand"] = m_inv.group(1).strip()
        if re.search(r"TNB Meter[^\|\n]*\|\s*Three Phase", full_text, re.I):
            data["tnb_meter"] = "Three Phase"
        elif re.search(r"TNB Meter[^\|\n]*\|\s*Single Phase", full_text, re.I):
            data["tnb_meter"] = "Single Phase"
        elif "0.4" in full_text and "kV" in full_text:
            data["tnb_meter"] = "Three Phase"
        elif "0.23" in full_text and "kV" in full_text:
            data["tnb_meter"] = "Single Phase"
            
        m_cap = re.search(r"Installed Cap(?:ca|ac)ity\s*\(kWp\)[^\|\n]*\|\s*([\d\.]+)", full_text, re.I)
        if m_cap:
            try:
                data["kwp"] = float(m_cap.group(1))
            except:
                pass
                
        search_scope = all_text_upper + " " + full_text.upper()
        m_pw = re.search(r"(\d+)\s*(?:PANELS?|PCS|NOS)?\s*(?:X|\*)\s*(\d{3,4})\s*W?", search_scope)
        if m_pw:
            data["panel_qty"] = int(m_pw.group(1))
            data["module_power_w"] = int(m_pw.group(2))
        else:
            m_qty_only = re.search(r"(\d+)\s*(?:PANELS?|PCS|NOS)\b", search_scope)
            if m_qty_only and int(m_qty_only.group(1)) < 200:
                data["panel_qty"] = int(m_qty_only.group(1))
                
        m_kwp = re.search(r"([\d\.]+)\s*KWP", search_scope)
        if m_kwp and data["kwp"] == 0.0:
            data["kwp"] = float(m_kwp.group(1))
            
        if data["panel_qty"] == 0 and data["kwp"] > 0 and data["module_power_w"] > 0:
            data["panel_qty"] = round((data["kwp"] * 1000.0) / data["module_power_w"])
        elif data["kwp"] == 0.0 and data["panel_qty"] > 0 and data["module_power_w"] > 0:
            data["kwp"] = round((data["panel_qty"] * data["module_power_w"]) / 1000.0, 2)
            
        return data

def parse_project_excel(file_path_or_bytes):
    """
    Main entry point to parse either New or Old Project Form (supports PDF or Excel).
    """
    if is_pdf_file(file_path_or_bytes):
        return parse_project_pdf(file_path_or_bytes)
    import io
    if isinstance(file_path_or_bytes, (bytes, bytearray)):
        file_path_or_bytes = io.BytesIO(file_path_or_bytes)
    fmt = detect_excel_format(file_path_or_bytes)
    if hasattr(file_path_or_bytes, "seek"):
        file_path_or_bytes.seek(0)
    if fmt == "new":
        return parse_new_format(file_path_or_bytes)
    else:
        return parse_old_format(file_path_or_bytes)

def parse_project_form(file_path_or_bytes, filename=None):
    """
    Universal entry point: parses Project Form whether uploaded as PDF or Excel.
    """
    if is_pdf_file(file_path_or_bytes, filename=filename):
        return parse_project_pdf(file_path_or_bytes)
    return parse_project_excel(file_path_or_bytes)
