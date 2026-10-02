import re
import pymupdf as fitz
import pdfplumber

def extract_client_and_address(doc, hint_client_name=None):
    """
    Robustly extract Client Name and Installation Address from DWG Title Block.
    Guarantees that contractor/consultant names (Northern Solar, North Consult, etc.)
    and CAD labels (SIZE, SYSTEM INFORMATION, etc.) are never mistaken for the client.
    Prioritizes full client address located directly below client name in the CLIENT title block.
    """
    page = doc[0]
    blocks = page.get_text("blocks")
    page_w = page.rect.width
    page_h = page.rect.height
    
    # Blacklisted terms that belong to contractor, consultant, or standard title block labels
    # Note: City names like KUALA LUMPUR, PETALING JAYA, and premise tags like UNIT are NOT blacklisted here
    # so that genuine client addresses are never discarded.
    blacklist = [
        "NORTHERN SOLAR", "NORTH CONSULT", "MENARA LAGENDA", "ENGINEERING",
        "CONTRACTOR", "CONSULTANT", "CONSTRUCTION DRAWING", "PROJECT TITLE",
        "DRAWING TITLE", "DESIGNED BY", "CHECKED BY", "SCALE", "DRW NO", "REV",
        "DATE", "CLIENT TO ACKNOWLEDGE", "INSTALLATION LOCATION", "MAP VIEW",
        "ELEVATION VIEW", "FRONT VIEW", "KEY PLAN", "REMARKS", "SS20/27",
        "OLD KLANG RD", "SIZE", "SYSTEM INFORMATION", "STRING CONFIGURATION",
        "STRING", "TOTAL PANELS", "PANEL TYPE", "PANEL", "INVERTER", "BATTERY",
        "SI UNIT", "SPECIFICATION", "MOUNTING", "STRUCTURE", "CABLE",
        "EARTHING", "GROUNDING", "SCHEMATIC", "SLD", "LAYOUT", "SINGLE LINE",
        "DWG", "A3", "N.T.S", "NTS", "APPROVED", "VERIFIED"
    ]
    
    def is_blacklisted(txt):
        u = txt.upper().strip()
        if not u:
            return True
        if u in blacklist:
            return True
        return any(b in u for b in blacklist if len(b) > 3)
        
    client_name = ""
    installation_address = ""
    p1_text = page.get_text("text")
    
    # 1. First, check if hint_client_name is provided or exists on Page 1
    # Priority: The block containing Client Name contains the authentic address directly below it!
    if hint_client_name and len(hint_client_name.strip()) > 3:
        h_u = hint_client_name.strip().upper()
        for b in blocks:
            btext = b[4].strip()
            if h_u in btext.upper():
                lines = [l.strip() for l in btext.split("\n") if l.strip()]
                for idx, line in enumerate(lines):
                    if h_u in line.upper():
                        client_name = line.split(",")[0].strip()
                        # Extract lines following client name as address
                        addr_candidates = []
                        for sub_line in lines[idx+1:]:
                            if not is_blacklisted(sub_line) and not any(h in sub_line.upper() for h in ["DRAWING TITLE", "PROJECT TITLE", "REV:", "SCALE:"]):
                                addr_candidates.append(sub_line)
                        if addr_candidates:
                            candidate_addr = " ".join(addr_candidates)
                            candidate_addr = re.sub(r"\s+", " ", candidate_addr).strip(", ")
                            if len(candidate_addr) > 5:
                                installation_address = candidate_addr
                        break
                if client_name:
                    break
        if not client_name and h_u in p1_text.upper():
            client_name = hint_client_name.strip()

    # 2. Extract Client Name & Address from CLIENT: title block
    if not client_name or not installation_address:
        client_header_box = None
        contractor_header_box = None
        
        for b in blocks:
            x0, y0, x1, y1, text = b[0], b[1], b[2], b[3], b[4].strip()
            # Direct match with CLIENT: header
            m_client_block = re.search(r"CLIENT\s*:?\s*\n?([\s\S]*?)(?:CONTRACTOR|CONSULTANT|PROJECT TITLE|DRAWING TITLE|$)", text, re.I)
            if m_client_block and "NORTHERN SOLAR" not in text.upper():
                c_section = m_client_block.group(1).strip()
                c_lines = [l.strip() for l in c_section.split("\n") if l.strip() and not is_blacklisted(l)]
                if c_lines:
                    if not client_name and len(c_lines[0]) > 3:
                        client_name = c_lines[0].split(",")[0].strip()
                    if not installation_address and len(c_lines) > 1:
                        addr_str = " ".join(c_lines[1:])
                        addr_str = re.sub(r"\s+", " ", addr_str).strip(", ")
                        if len(addr_str) > 5:
                            installation_address = addr_str
            
            if re.search(r"^CLIENT\s*:?$", text, re.I):
                client_header_box = (x0, y0, x1, y1)
            elif "CONTRACTOR" in text.upper():
                contractor_header_box = (x0, y0, x1, y1)

        # Look for block situated vertically below CLIENT: header
        if client_header_box and (not client_name or not installation_address):
            cy_bottom = client_header_box[3]
            cx_left = client_header_box[0]
            py_top = contractor_header_box[1] if contractor_header_box else cy_bottom + 120
            
            candidates = []
            for b in blocks:
                bx0, by0, bx1, by1, btext = b[0], b[1], b[2], b[3], b[4].strip()
                if cy_bottom - 5 <= by0 <= py_top + 15 and abs(bx0 - cx_left) < 180 and bx0 >= page_w * 0.40:
                    if not is_blacklisted(btext) and btext.upper() != "CLIENT:":
                        candidates.append((by0, btext))
                        
            if candidates:
                candidates.sort(key=lambda x: x[0])
                for _, btext in candidates:
                    lines = [l.strip() for l in btext.split("\n") if l.strip() and not is_blacklisted(l)]
                    if lines:
                        if not client_name:
                            client_name = lines[0].split(",")[0].strip()
                        if not installation_address and len(lines) > 1:
                            addr_str = " ".join(lines[1:])
                            addr_str = re.sub(r"\s+", " ", addr_str).strip(", ")
                            if len(addr_str) > 5:
                                installation_address = addr_str
                    if client_name and installation_address:
                        break

    # 3. Fallback: Search for address below client name across adjacent blocks
    if client_name and not installation_address:
        for b in blocks:
            btext = b[4].strip()
            if client_name.upper() in btext.upper():
                after_idx = btext.upper().find(client_name.upper()) + len(client_name)
                after_text = btext[after_idx:].strip(", \n\r")
                clean_lines = [l.strip() for l in after_text.split("\n") if l.strip() and not is_blacklisted(l)]
                if clean_lines:
                    candidate = " ".join(clean_lines)
                    candidate = re.sub(r"\s+", " ", candidate).strip(", ")
                    if len(candidate) > 5:
                        installation_address = candidate
                        break

    # 4. Fallback: Extract Address from PROJECT TITLE: ... AT <ADDRESS>
    if not installation_address:
        m_proj = re.search(r"PROJECT TITLE:[\s\S]*?(?:ATAP|ON-GRID|AT)?\s+AT\s+([\s\S]*?)(?:DRAWING TITLE|This drawing|REV:|$)", p1_text, re.I)
        if m_proj:
            raw_addr = m_proj.group(1).strip()
            clean_lines = [l.strip() for l in raw_addr.split("\n") if l.strip() and not is_blacklisted(l)]
            if clean_lines:
                candidate_addr = ", ".join(clean_lines)
                candidate_addr = re.split(r"This drawing|Contractors must|DRAWING TITLE", candidate_addr, flags=re.I)[0].strip(", ")
                if len(candidate_addr) > 5:
                    installation_address = candidate_addr

    # Final cleanup
    client_name = client_name.strip(", ")
    installation_address = installation_address.strip(", ")
    
    return client_name, installation_address

def extract_pdf_data(pdf_path_or_bytes, hint_client_name=None):
    """
    Extract structured data from the DWG Construction Drawing PDF.
    Pages:
      Page 1: PV Layout & System Info
      Page 2: Mounting Structure BOM
      Page 3: SLD (Single Line Diagram)
      Page 4: SLD - PVMSB
    """
    if isinstance(pdf_path_or_bytes, bytes):
        doc = fitz.open(stream=pdf_path_or_bytes, filetype="pdf")
    else:
        doc = fitz.open(pdf_path_or_bytes)
        
    num_pages = len(doc)
    
    extracted = {
        "num_pages": num_pages,
        "client_name": "",
        "project_title": "",
        "address": "",
        "page1": {},
        "page2": {},
        "page3": {},
        "page4": {},
        "all_pages_text": []
    }
    
    for i in range(num_pages):
        page_text = doc[i].get_text("text")
        extracted["all_pages_text"].append(page_text)
        
    # Extract Client & Address using specialized geometric & text parser
    if num_pages >= 1:
        extracted["client_name"], extracted["address"] = extract_client_and_address(doc, hint_client_name=hint_client_name)

    # ==========================
    # PAGE 1: PV Layout & System Info
    # ==========================
    if num_pages >= 1:
        p1_text = extracted["all_pages_text"][0]
        p1_data = {
            "panel_model": "",
            "panel_type": "",
            "total_panels": 0,
            "inverter": "",
            "battery": "",
            "string_config": [],
            "kwp": 0.0,
            "kwac": 0.0,
            "strings_count": 0,
            "jumper_lines_count": 0
        }
        
        # Panel Model
        m_panel = re.search(r"PANEL\s+([A-Z0-9\-_ ]+W)", p1_text, re.I)
        if m_panel:
            p1_data["panel_model"] = m_panel.group(1).strip()
            
        m_ptype = re.search(r"PANEL TYPE\s+([A-Z]+)", p1_text, re.I)
        if m_ptype:
            p1_data["panel_type"] = m_ptype.group(1).strip()
            
        # Total Panels
        m_tot_p = re.search(r"TOTAL PANELS\s+(\d+)\s*PCS", p1_text, re.I)
        if m_tot_p:
            p1_data["total_panels"] = int(m_tot_p.group(1))
            
        # Inverter
        m_inv = re.search(r"INVERTER\s+([^\n]+(?:HOYMILES|HUAWEI|SOLIS|SUNGROW|GROWATT)[^\n]*)", p1_text, re.I)
        if m_inv:
            p1_data["inverter"] = m_inv.group(1).strip()
        else:
            m_inv2 = re.search(r"INVERTER\s+(\d+\s*x\s*[A-Z0-9\.\-_ ]+)", p1_text, re.I)
            if m_inv2:
                p1_data["inverter"] = m_inv2.group(1).strip()
                
        # Battery
        p1_data["battery_qty"] = 0
        m_bat = re.search(r"BATTERY\s+([^\n]+(?:HOYMILES|LB-|LUNA|BYD|GROWATT)[^\n]*)", p1_text, re.I)
        if m_bat:
            p1_data["battery"] = m_bat.group(1).strip()
        else:
            m_bat2 = re.search(r"BATTERY\s+([\d\.]+\s*x\s*[A-Z0-9\.\-_ ]+)", p1_text, re.I)
            if m_bat2:
                p1_data["battery"] = m_bat2.group(1).strip()
                
        m_bat_qty = re.search(r"BATTERY\s+([\d\.]+)\s*x", p1_text, re.I)
        if m_bat_qty:
            try:
                p1_data["battery_qty"] = int(float(m_bat_qty.group(1)))
            except:
                p1_data["battery_qty"] = 1
        elif p1_data["battery"]:
            p1_data["battery_qty"] = 1
                
        # String Config (e.g. 08 PCS X 1 STG, 05 PCS X 1 STG)
        str_matches = re.findall(r"(\d+)\s*PCS\s*X\s*(\d+)\s*STG", p1_text, re.I)
        if str_matches:
            for count, stg in str_matches:
                p1_data["string_config"].append(f"{count} PCS X {stg} STG")
            p1_data["strings_count"] = sum([int(s) for _, s in str_matches])
        else:
            layout_strings = re.findall(r"(?:INV\s*\d*\s*-\s*STR\s*\d*|\bSTR\s*\d+)\s*(?:\(\s*(\d+)\s*\)|(\d+))", p1_text)
            if layout_strings:
                p1_data["strings_count"] = len(layout_strings)
                
        # Capacity Size: 11.34kWp/ 10.0 kWac or 8.82 kWp
        m_sz = re.search(r"SIZE\s+([\d\.]+)\s*kWp\s*/\s*([\d\.]+)\s*kWac", p1_text, re.I)
        if m_sz:
            p1_data["kwp"] = float(m_sz.group(1))
            p1_data["kwac"] = float(m_sz.group(2))
        else:
            m_kwp = re.search(r"([\d\.]+)\s*kWp", p1_text, re.I)
            if m_kwp:
                p1_data["kwp"] = float(m_kwp.group(1))
            m_kwac = re.search(r"([\d\.]+)\s*kWac", p1_text, re.I)
            if m_kwac:
                p1_data["kwac"] = float(m_kwac.group(1))
                
        # Jumper Line check (cyan/green graphics)
        page1_obj = doc[0]
        drawings = page1_obj.get_drawings()
        cyan_count = 0
        for d in drawings:
            color = d.get("color")
            if color and len(color) == 3:
                r, g, b = color
                if r < 0.3 and g > 0.6 and b > 0.5:
                    cyan_count += 1
                elif r < 0.3 and g > 0.6 and b < 0.4:
                    cyan_count += 1
                    
        p1_data["jumper_lines_count"] = min(cyan_count // 2, 4) if cyan_count > 0 else 0
        extracted["page1"] = p1_data

    # ==========================
    # PAGE 2: Mounting Structure BOM
    # ==========================
    if num_pages >= 2:
        p2_text = extracted["all_pages_text"][1]
        p2_data = {
            "bom_items": {},
            "end_clamp": 0,
            "mid_clamp": 0,
            "tile_hook": 0,
            "cliplock": 0,
            "grounding_clip": 0,
            "grounding_lug": 0,
            "mc4_connector": 0,
            "r_railing": 0,
            "splice": 0,
            "u_rail": 0,
            "l_foot": 0,
            "shingle": 0,
            "ballast": 0,
            "roof_attachment_name": "",
            "roof_attachment_qty": 0,
            "tile_hook_count_symbols": 0
        }
        
        patterns = {
            "end_clamp": r"End\s+Clamp\s*[\n\|]?\s*(\d+)",
            "mid_clamp": r"Mid\s+Clamp\s*[\n\|]?\s*(\d+)",
            "tile_hook": r"Tile\s+Hook\s*[\n\|]?\s*(\d+)",
            "cliplock": r"(?:Universal\s+Cliplock|Cliplock)[^\d\n]*\s*(\d+)",
            "grounding_clip": r"Grounding\s+Clip\s*[\n\|]?\s*(\d+)",
            "grounding_lug": r"Grounding\s+Lug\s*[\n\|]?\s*(\d+)",
            "mc4_connector": r"MC4\s+Connector\s*[\n\|]?\s*(\d+)",
            "r_railing": r"R-Railing\s*[\n\|]?\s*(\d+)",
            "splice": r"Splice\s*[\n\|]?\s*(\d+)",
            "u_rail": r"U-Rail\s*[\n\|]?\s*(\d+)",
            "l_foot": r"L-Foot\s*[\n\|]?\s*(\d+)",
            "shingle": r"(?:Shingle\s+Plate|Shingle)[^\d\n]*\s*(\d+)",
            "ballast": r"(?:Concrete\s+Ballast|Ballast)[^\d\n]*\s*(\d+)"
        }
        
        for key, pat in patterns.items():
            m = re.search(pat, p2_text, re.I)
            if m:
                p2_data[key] = int(m.group(1))
                p2_data["bom_items"][key] = int(m.group(1))
                
        # Count explicit 'TH' occurrences on Page 2 drawing
        th_symbols = re.findall(r"\bTH\b", p2_text)
        if len(th_symbols) > 2:
            p2_data["tile_hook_count_symbols"] = len(th_symbols) - 1

        # Full table extraction from Page 2 System Information table
        raw_bom_table = []
        try:
            tabs = doc[1].find_tables()
            for t in tabs.tables:
                ext = t.extract()
                has_sys_info = any("System Information" in str(cell) for row in ext for cell in row if cell)
                has_splice = any("Splice" in str(cell) for row in ext for cell in row if cell)
                has_end_clamp = any("End Clamp" in str(cell) for row in ext for cell in row if cell)
                
                if has_sys_info or has_splice or has_end_clamp:
                    for row in ext:
                        if not row or len(row) < 2:
                            continue
                        cells = [str(c).strip() if c is not None else "" for c in row]
                        item_num = ""
                        item_name = ""
                        item_qty = 0
                        
                        if cells[0].isdigit():
                            item_num = cells[0]
                            item_name = cells[1]
                            for c in cells[2:]:
                                m_q = re.search(r"(\d+)", c)
                                if m_q:
                                    item_qty = int(m_q.group(1))
                                    break
                        elif len(cells) >= 3 and cells[1].isdigit():
                            item_num = cells[1]
                            item_name = cells[2]
                            for c in cells[3:]:
                                m_q = re.search(r"(\d+)", c)
                                if m_q:
                                    item_qty = int(m_q.group(1))
                                    break
                                    
                        if item_num and item_name and not any(h in item_name.upper() for h in ["INFO", "DESCRIPTION", "ITEM"]):
                            raw_bom_table.append({
                                "item_no": item_num,
                                "name": item_name,
                                "qty": item_qty
                            })
        except Exception:
            pass

        # Fallback regex for raw_bom_table if table finder was empty
        if not raw_bom_table:
            for line in p2_text.split("\n"):
                m_row = re.match(r"^(\d+)\s+([A-Za-z0-9\-\s\(\)]+?)\s+(\d+)$", line.strip())
                if m_row:
                    num, name, qty = m_row.group(1), m_row.group(2).strip(), int(m_row.group(3))
                    if not any(h in name.upper() for h in ["INFO", "ITEM", "PAGE", "TOTAL"]):
                        raw_bom_table.append({
                            "item_no": num,
                            "name": name,
                            "qty": qty
                        })

        p2_data["raw_bom_table"] = raw_bom_table
        for row in raw_bom_table:
            num = row["item_no"]
            name = row["name"]
            qty = row["qty"]
            u_name = name.upper()
            
            p2_data["bom_items"][name] = qty
            if num == "1" or "END CLAMP" in u_name:
                p2_data["end_clamp"] = qty
            elif num == "2" or "MID CLAMP" in u_name:
                p2_data["mid_clamp"] = qty
            elif num == "3":
                p2_data["roof_attachment_name"] = name
                p2_data["roof_attachment_qty"] = qty
                if "TILE" in u_name: p2_data["tile_hook"] = qty
                elif "U-RAIL" in u_name or "URAIL" in u_name: p2_data["u_rail"] = qty
                elif "CLIPLOCK" in u_name: p2_data["cliplock"] = qty
                elif "L-FOOT" in u_name or "L FOOT" in u_name: p2_data["l_foot"] = qty
                elif "SHINGLE" in u_name: p2_data["shingle"] = qty
                elif "BALLAST" in u_name: p2_data["ballast"] = qty
            elif num == "4" or "GROUNDING CLIP" in u_name:
                p2_data["grounding_clip"] = qty
            elif num == "5" or "GROUNDING LUG" in u_name:
                p2_data["grounding_lug"] = qty
            elif num == "6" or "MC4" in u_name:
                p2_data["mc4_connector"] = qty
            elif num == "7" or "RAILING" in u_name:
                p2_data["r_railing"] = qty
            elif num == "8" or "SPLICE" in u_name:
                p2_data["splice"] = qty
            else:
                if "L-FOOT" in u_name or "L FOOT" in u_name:
                    p2_data["l_foot"] = qty
            
        extracted["page2"] = p2_data

    # ==========================
    # PAGE 3: SLD (Single Line Diagram)
    # ==========================
    if num_pages >= 3:
        p3_text = extracted["all_pages_text"][2]
        p3_data = {
            "phase": "Unknown",
            "mppt_count": 0,
            "inverter_model": "",
            "battery_model": "",
            "kwp": 0.0
        }
        
        has_r = bool(re.search(r"\bR\b", p3_text))
        has_y = bool(re.search(r"\bY\b", p3_text))
        has_b = bool(re.search(r"\bB\b", p3_text))
        has_n = bool(re.search(r"\bN\b", p3_text))
        
        if has_r and has_y and has_b and has_n:
            p3_data["phase"] = "Three Phase"
        elif "THREE PHASE" in p3_text.upper():
            p3_data["phase"] = "Three Phase"
        elif "SINGLE PHASE" in p3_text.upper():
            p3_data["phase"] = "Single Phase"
        elif has_n and bool(re.search(r"\bL\b", p3_text)):
            p3_data["phase"] = "Single Phase"
            
        mppts = re.findall(r"MPPT\s*\d+", p3_text, re.I)
        p3_data["mppt_count"] = len(set(mppts))
        
        m_inv = re.search(r"INVERTER[^\n]*\s+(\d+\s*x\s*[A-Z0-9\.\-_ ]+)", p3_text, re.I)
        if m_inv:
            p3_data["inverter_model"] = m_inv.group(1).strip()
            
        m_kwp = re.search(r"([\d\.]+)\s*kWp", p3_text, re.I)
        if m_kwp:
            p3_data["kwp"] = float(m_kwp.group(1))
            
        extracted["page3"] = p3_data

    # ==========================
    # PAGE 4: SLD - PVMSB
    # ==========================
    if num_pages >= 4:
        p4_text = extracted["all_pages_text"][3]
        p4_data = {
            "inverter_kw": 0.0,
            "inverter_amp": 0.0,
            "inverter_mcb": "",
            "mcb_poles": "",
            "pv_meter": "",
            "battery_ah": 0,
            "battery_mccb": "",
            "ac_cable": ""
        }
        
        m_ikw = re.search(r"INVERTER\s*[\n\r]*\s*([\d\.]+)\s*kW\s*(?:\(([\d\.]+)\s*A\))?", p4_text, re.I)
        if m_ikw:
            p4_data["inverter_kw"] = float(m_ikw.group(1))
            if m_ikw.group(2):
                p4_data["inverter_amp"] = float(m_ikw.group(2))
                
        # Inverter MCB (e.g. "32A 3P 10kA MCB" or "25A 1P MCB")
        m_mcb = re.search(r"(\d+A[\s\n]+(?:3P|1P)[\s\n]+(?:10kA[\s\n]+)?MCB)", p4_text, re.I)
        if m_mcb:
            p4_data["inverter_mcb"] = re.sub(r"\s+", " ", m_mcb.group(1)).strip()
            if "3P" in p4_data["inverter_mcb"].upper():
                p4_data["mcb_poles"] = "3P (Three Phase)"
            elif "1P" in p4_data["inverter_mcb"].upper():
                p4_data["mcb_poles"] = "1P (Single Phase)"
        else:
            m_mcb2 = re.search(r"(\d+A)\s+(\d+P)\s+MCB", p4_text, re.I)
            if m_mcb2:
                p4_data["inverter_mcb"] = f"{m_mcb2.group(1)} {m_mcb2.group(2)} MCB"
                p4_data["mcb_poles"] = m_mcb2.group(2)
            else:
                m_mcb3 = re.search(r"(\d+A\s*(?:3P|1P)\s*MCB)", p4_text, re.I)
                if m_mcb3:
                    p4_data["inverter_mcb"] = m_mcb3.group(1).strip()
                    if "3P" in p4_data["inverter_mcb"].upper():
                        p4_data["mcb_poles"] = "3P (Three Phase)"
                    elif "1P" in p4_data["inverter_mcb"].upper():
                        p4_data["mcb_poles"] = "1P (Single Phase)"
                
        # PV Meter & CT Rating (e.g. "32VA CL1 3CTS" or "25VA CL1 3CTS" or "32VA")
        m_meter = re.search(r"(\d+VA[\s\n]+CL\d+[\s\n]+\d*CTs?)", p4_text, re.I)
        if m_meter:
            p4_data["pv_meter"] = re.sub(r"\s+", " ", m_meter.group(1)).strip()
        else:
            m_meter2 = re.search(r"(\d+VA)", p4_text, re.I)
            if m_meter2:
                p4_data["pv_meter"] = f"{m_meter2.group(1)} CL1 3CTS"
            else:
                m_meter3 = re.search(r"(\d+/\d+A[^\n]*\d*CTs?)", p4_text, re.I)
                if m_meter3:
                    p4_data["pv_meter"] = m_meter3.group(1).strip()
            
        m_bah = re.search(r"BATTERY[^\d]*(\d+)\s*Ah", p4_text, re.I)
        if m_bah:
            p4_data["battery_ah"] = int(m_bah.group(1))
            
        m_bmccb = re.search(r"(\d+A\s*2P\s*MCCB)", p4_text, re.I)
        if m_bmccb:
            p4_data["battery_mccb"] = m_bmccb.group(1).strip()
            
        m_cbl = re.search(r"(\d+\s*X\s*1C\s*\d+MM²[^\n]*\(RYBN\))", p4_text, re.I)
        if m_cbl:
            p4_data["ac_cable"] = m_cbl.group(1).strip()
            
        extracted["page4"] = p4_data

    doc.close()
    return extracted

def render_pdf_page_image(pdf_bytes, page_num=0, dpi=150):
    """Renders a specific page of a PDF as PNG bytes at the given DPI."""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if 0 <= page_num < len(doc):
            pix = doc[page_num].get_pixmap(dpi=dpi)
            img_bytes = pix.tobytes("png")
            doc.close()
            return img_bytes
        doc.close()
    except Exception:
        pass
    return None

def extract_dwg_system_info_image(pdf_bytes, dpi=200):
    """
    Extracts the high-resolution crop of the SYSTEM INFORMATION table
    from Page 1 of the DWG PDF.
    """
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page = doc[0]
        
        # 1. Try find_tables
        tables = page.find_tables()
        for t in tables.tables:
            text = " ".join([str(c) for row in t.extract() for c in row if c])
            if "SYSTEM INFORMATION" in text.upper():
                bbox = fitz.Rect(t.bbox)
                pad_rect = fitz.Rect(bbox.x0 - 5, bbox.y0 - 5, bbox.x1 + 5, bbox.y1 + 5)
                pad_rect.intersect(page.rect)
                pix = page.get_pixmap(clip=pad_rect, dpi=dpi)
                b = pix.tobytes("png")
                doc.close()
                return b
                
        # 2. Try search_for
        matches = page.search_for("SYSTEM INFORMATION")
        if matches:
            r = matches[0]
            clip_rect = fitz.Rect(r.x0 - 15, r.y0 - 15, r.x0 + 195, r.y0 + 225)
            clip_rect.intersect(page.rect)
            pix = page.get_pixmap(clip=clip_rect, dpi=dpi)
            b = pix.tobytes("png")
            doc.close()
            return b
            
        doc.close()
    except Exception:
        pass
    return None

def extract_dwg_pv_layout_image(pdf_bytes, dpi=200):
    """
    Extracts the high-resolution crop of the Roof PV Layout (Front View)
    with string routing and panel modules from Page 1 of the DWG PDF.
    """
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page = doc[0]
        
        str_runs = []
        for b in page.get_text("words"):
            if any(k in b[4] for k in ["STR", "INV", "FALL", "mm", "FRONT", "VIEW"]):
                if 200 <= b[1] <= 740:
                    str_runs.append(b)
                    
        if str_runs:
            min_x = max(10, min(b[0] for b in str_runs) - 30)
            min_y = max(10, min(b[1] for b in str_runs) - 20)
            max_x = min(page.rect.width - 10, max(b[2] for b in str_runs) + 30)
            max_y = min(730, max(b[3] for b in str_runs) + 30)
            crop_rect = fitz.Rect(min_x, min_y, max_x, max_y)
            pix = page.get_pixmap(clip=crop_rect, dpi=dpi)
            b = pix.tobytes("png")
            doc.close()
            return b
            
        # Fallback: crop the central roof drawing area
        crop_rect = fitz.Rect(30, page.rect.height * 0.25, page.rect.width * 0.75, page.rect.height * 0.85)
        pix = page.get_pixmap(clip=crop_rect, dpi=dpi)
        b = pix.tobytes("png")
        doc.close()
        return b
    except Exception:
        pass
    return None

def extract_dwg_page2_table_image(pdf_bytes, dpi=200):
    """
    Extracts a focused high-resolution zoom of both the System Information table
    and the separate Cable Information table on Page 2 of the DWG PDF.
    """
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if len(doc) < 2:
            doc.close()
            return None
        page = doc[1]
        tabs = page.find_tables()
        for t in tabs.tables:
            if (t.bbox[2] - t.bbox[0]) < page.rect.width * 0.4 and (t.bbox[3] - t.bbox[1]) < page.rect.height * 0.7:
                bbox = fitz.Rect(t.bbox[0] - 8, t.bbox[1] - 8, t.bbox[2] + 8, t.bbox[3] + 125)
                bbox.intersect(page.rect)
                pix = page.get_pixmap(clip=bbox, dpi=dpi)
                b = pix.tobytes("png")
                doc.close()
                return b
        # Fallback
        clip_rect = fitz.Rect(20, 20, 220, 420)
        clip_rect.intersect(page.rect)
        pix = page.get_pixmap(clip=clip_rect, dpi=dpi)
        b = pix.tobytes("png")
        doc.close()
        return b
    except Exception:
        pass
    return None


