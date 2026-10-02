import re
import math

# Inverter Technical Lookup Matrix according to Northern Solar standard
INVERTER_SPEC_RULES = {
    "HIS-5": {
        "tag": "HIS-5L-G3",
        "name": "HOYMILES HIS-5.0L-G3",
        "current_a": 25.0,
        "mcb": "32A 1P 10kA MCB",
        "poles": "1P",
        "phase": "Single Phase",
        "meter": "32VA CL1 3CTS"
    },
    "HIT-8": {
        "tag": "HIT-8L-G3",
        "name": "HOYMILES HIT-8.0L-G3",
        "current_a": 13.3,
        "mcb": "32A 3P 10kA MCB",
        "poles": "3P",
        "phase": "Three Phase",
        "meter": "32VA CL1 3CTS"
    },
    "HIT-10": {
        "tag": "HIT-10L-G3",
        "name": "HOYMILES HIT-10.0L-G3",
        "current_a": 16.7,
        "mcb": "32A 3P 10kA MCB",
        "poles": "3P",
        "phase": "Three Phase",
        "meter": "32VA CL1 3CTS"
    },
    "HIT-12": {
        "tag": "HIT-12L-G3",
        "name": "HOYMILES HIT-12.0L-G3",
        "current_a": 20.0,
        "mcb": "32A 3P 10kA MCB",
        "poles": "3P",
        "phase": "Three Phase",
        "meter": "32VA CL1 3CTS"
    },
    "HIT-15": {
        "tag": "HIT-15L-G3",
        "name": "HOYMILES HIT-15.0L-G3",
        "current_a": 25.0,
        "mcb": "32A 3P 10kA MCB",
        "poles": "3P",
        "phase": "Three Phase",
        "meter": "32VA CL1 3CTS"
    },
    "HIT-20": {
        "tag": "HIT-20L-G3",
        "name": "HOYMILES HIT-20.0L-G3",
        "current_a": 33.3,
        "mcb": "40A 3P 10kA MCB",
        "poles": "3P",
        "phase": "Three Phase",
        "meter": "40VA CL1 3CTS"
    }
}

def normalize_text(text):
    if not text:
        return ""
    t = text.upper()
    t = re.sub(r"[,\.\-@/\\_\(\)]", " ", t)
    t = re.sub(r"\bBIN\b", "B", t)
    t = re.sub(r"\bBINTI\b", "BT", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t

def check_address_similarity(addr1, addr2):
    n1 = normalize_text(addr1)
    n2 = normalize_text(addr2)
    
    if not n1 or not n2:
        return False, "Address information incomplete for comparison."
        
    digits1 = set(re.findall(r"\b\d+\b", n1))
    digits2 = set(re.findall(r"\b\d+\b", n2))
    
    common_digits = digits1.intersection(digits2)
    if digits1 and digits2 and not common_digits:
        return False, f"Premise number/postcode mismatch: [{', '.join(digits1)}] vs [{', '.join(digits2)}]"
        
    words1 = set([w for w in n1.split() if len(w) > 2])
    words2 = set([w for w in n2.split() if len(w) > 2])
    
    if not words1 or not words2:
        return True, "Address matches."
        
    overlap = words1.intersection(words2)
    ratio = len(overlap) / max(len(words1), 1)
    
    if ratio >= 0.35:
        return True, f"Address matches ({int(ratio*100)}% keyword match)."
    else:
        return False, f"Address text mismatch ({int(ratio*100)}% keyword match)."

def extract_panel_wattage(model_str):
    if not model_str:
        return 0.0
    m = re.search(r"\b(\d{3,4})\s*W?\b", model_str)
    if m:
        w = int(m.group(1))
        if 300 <= w <= 800:
            return w / 1000.0
    return 0.0

def match_inverter_spec(inverter_str):
    if not inverter_str:
        return None
    inv_u = inverter_str.upper()
    for key, spec in INVERTER_SPEC_RULES.items():
        if key in inv_u or key.replace("-", "") in inv_u:
            return spec
    for key, spec in INVERTER_SPEC_RULES.items():
        sub = key.split("-")[1]
        if f"{sub}L" in inv_u or f"{sub}.0L" in inv_u:
            return spec
    return None

def run_qc_check(excel_data, pdf_data, manual_jumpers=None, roof_name_override=None, table_overrides=None):
    """
    Runs full QC check and returns:
      1. page1_results: Client Info & Page 1 (PV Layout)
      2. page2_bom_rows: Detailed Mounting Structure comparison table
      3. page3_4_results: SLD, Phase, Inverter Current, Breaker & PV Meter
    """
    page1_results = []
    
    # ----------------------------------------------------
    # 1. CLIENT & SITE DETAILS
    # ----------------------------------------------------
    ex_name = excel_data.get("client_name", "")
    pdf_name = pdf_data.get("client_name", "")
    
    # Safety check: if pdf_name did not match or was extracted as a CAD label,
    # but the DWG drawing text actually contains the client name from the Project Form:
    p1_dwg_text = ""
    if pdf_data.get("all_pages_text"):
        p1_dwg_text = pdf_data["all_pages_text"][0]
        
    n_ex_name = normalize_text(ex_name)
    n_pdf_name = normalize_text(pdf_name)
    
    if n_ex_name and (n_ex_name in normalize_text(p1_dwg_text)):
        pdf_name = ex_name
        n_pdf_name = n_ex_name
    
    if n_ex_name and n_pdf_name:
        is_name_match = (n_ex_name in n_pdf_name) or (n_pdf_name in n_ex_name) or (n_ex_name == n_pdf_name)
        status_name = "PASS" if is_name_match else "FAIL"
        note_name = "Client name matches." if is_name_match else f"Name mismatch: Excel '{ex_name}' vs DWG '{pdf_name}'!"
    elif not n_ex_name and not n_pdf_name:
        status_name = "WARNING"
        note_name = "No client name detected in both files."
    else:
        status_name = "WARNING"
        note_name = "Client name detected in only one document."
        
    page1_results.append({
        "category": "Client Information",
        "item": "Client Name",
        "excel_val": ex_name or "No Data",
        "dwg_val": pdf_name or "No Data",
        "status": status_name,
        "notes": note_name
    })
    
    # Address
    ex_addr = excel_data.get("address", "")
    pdf_addr = pdf_data.get("address", "")
    
    if ex_addr and pdf_addr:
        is_addr_match, addr_note = check_address_similarity(ex_addr, pdf_addr)
        status_addr = "PASS" if is_addr_match else "FAIL"
    else:
        status_addr = "WARNING"
        addr_note = "Incomplete address information in one of the files."
        
    page1_results.append({
        "category": "Client Information",
        "item": "Installation Address",
        "excel_val": ex_addr or "No Data",
        "dwg_val": pdf_addr or "No Data",
        "status": status_addr,
        "notes": addr_note
    })

    # ----------------------------------------------------
    # 2. PAGE 1: PV LAYOUT & SYSTEM SPECIFICATIONS
    # ----------------------------------------------------
    p1 = pdf_data.get("page1", {})
    
    dwg_pmodel = p1.get("panel_model", "")
    ex_pmodel = f"{excel_data.get('module_brand', '')} {excel_data.get('module_model', '')}".strip()
    
    unit_panel_kw = extract_panel_wattage(dwg_pmodel) or extract_panel_wattage(ex_pmodel) or 0.630
    
    ex_kwp = excel_data.get("kwp", 0.0)
    dwg_kwp = p1.get("kwp", 0.0)
    
    ex_qty = excel_data.get("panel_qty", 0)
    dwg_qty = p1.get("total_panels", 0)
    
    formula_used_note = ""
    if ex_qty == 0 and ex_kwp > 0 and unit_panel_kw > 0:
        ex_qty = round(ex_kwp / unit_panel_kw)
        formula_used_note = f" (Formula: {ex_kwp:.2f}kWp ÷ {unit_panel_kw:.3f})"
    elif ex_qty == 0 and dwg_kwp > 0 and unit_panel_kw > 0:
        ex_qty = round(dwg_kwp / unit_panel_kw)
        formula_used_note = f" (Formula: {dwg_kwp:.2f}kWp ÷ {unit_panel_kw:.3f})"
        
    if ex_qty > 0 and dwg_qty > 0:
        qty_match = (ex_qty == dwg_qty)
        status_qty = "PASS" if qty_match else "FAIL"
        note_qty = f"Panel quantity matches ({ex_qty} PCS)." if qty_match else f"MISMATCH: Excel requests {ex_qty} panels vs DWG drawings {dwg_qty} panels!"
    else:
        status_qty = "WARNING"
        note_qty = "Unable to determine panel quantity."
        
    page1_results.append({
        "category": "Page 1: PV Layout",
        "item": "Solar Panel Quantity",
        "excel_val": f"{ex_qty} PCS{formula_used_note}" if ex_qty else "No Data",
        "dwg_val": f"{dwg_qty} PCS" if dwg_qty else "No Data",
        "status": status_qty,
        "notes": note_qty
    })
    
    # System Capacity kWp
    if ex_kwp > 0 and dwg_kwp > 0:
        kwp_match = abs(ex_kwp - dwg_kwp) < 0.1
        status_kwp = "PASS" if kwp_match else "FAIL"
        note_kwp = "System kWp capacity matches." if kwp_match else f"Capacity difference: Excel {ex_kwp:.2f} kWp vs DWG {dwg_kwp:.2f} kWp."
    elif ex_kwp == 0 and dwg_kwp > 0:
        status_kwp = "INFO"
        note_kwp = f"DWG capacity is {dwg_kwp:.2f} kWp."
    else:
        status_kwp = "WARNING"
        note_kwp = "kWp capacity not specified."
        
    page1_results.append({
        "category": "Page 1: PV Layout",
        "item": "System Capacity (kWp)",
        "excel_val": f"{ex_kwp:.2f} kWp" if ex_kwp else "No Data",
        "dwg_val": f"{dwg_kwp:.2f} kWp" if dwg_kwp else "No Data",
        "status": status_kwp,
        "notes": note_kwp
    })
    
    # Panel Model
    if ex_pmodel and dwg_pmodel:
        n_ex_p = normalize_text(ex_pmodel)
        n_dwg_p = normalize_text(dwg_pmodel)
        pmodel_match = (n_ex_p in n_dwg_p) or (n_dwg_p in n_ex_p) or bool(set(n_ex_p.split()).intersection(set(n_dwg_p.split())))
        status_pmodel = "PASS" if pmodel_match else "WARNING"
        note_pmodel = "PV Module model matches." if pmodel_match else "Module model may differ between Excel and DWG."
    elif not ex_pmodel and dwg_pmodel:
        status_pmodel = "INFO"
        note_pmodel = f"DWG Module: {dwg_pmodel} (Excel has no specific model)."
    else:
        status_pmodel = "WARNING"
        note_pmodel = "Module model specifications incomplete."
        
    page1_results.append({
        "category": "Page 1: PV Layout",
        "item": "PV Module Model",
        "excel_val": ex_pmodel or "No Data",
        "dwg_val": dwg_pmodel or "No Data",
        "status": status_pmodel,
        "notes": note_pmodel
    })
    
    # Inverter Model
    ex_inv = f"{excel_data.get('inverter_brand', '')} {excel_data.get('inverter_model', '')}".strip()
    dwg_inv = p1.get("inverter", "")
    
    if ex_inv and dwg_inv:
        n_ex_i = normalize_text(ex_inv)
        n_dwg_i = normalize_text(dwg_inv)
        inv_match = (n_ex_i in n_dwg_i) or (n_dwg_i in n_ex_i) or ("HOYMILES" in n_ex_i and "HOYMILES" in n_dwg_i)
        status_inv = "PASS" if inv_match else "FAIL"
        note_inv = "Inverter model matches." if inv_match else "Inverter model mismatch."
    elif not ex_inv and dwg_inv:
        status_inv = "INFO"
        note_inv = f"DWG Inverter: {dwg_inv}."
    else:
        status_inv = "WARNING"
        note_inv = "Inverter specifications incomplete."
        
    page1_results.append({
        "category": "Page 1: PV Layout",
        "item": "Inverter Model",
        "excel_val": ex_inv or "No Data",
        "dwg_val": dwg_inv or "No Data",
        "status": status_inv,
        "notes": note_inv
    })
    
    # Battery Model
    ex_bat = f"{excel_data.get('battery_brand', '')} {excel_data.get('battery_model', '')}".strip()
    dwg_bat = p1.get("battery", "")
    
    if ex_bat and dwg_bat:
        n_ex_b = normalize_text(ex_bat)
        n_dwg_b = normalize_text(dwg_bat)
        bat_match = (n_ex_b in n_dwg_b) or (n_dwg_b in n_ex_b) or ("LB-16" in n_ex_b and "LB-16" in n_dwg_b)
        status_bat = "PASS" if bat_match else "FAIL"
        note_bat = "Battery model matches." if bat_match else "Battery model mismatch."
    elif not ex_bat and dwg_bat:
        status_bat = "INFO"
        note_bat = f"DWG Battery: {dwg_bat}."
    elif ex_bat and not dwg_bat:
        status_bat = "WARNING"
        note_bat = f"Excel requests battery {ex_bat}, but DWG has no battery!"
    else:
        status_bat = "PASS"
        note_bat = "No battery in both documents (Standard solar system)."
        
    page1_results.append({
        "category": "Page 1: PV Layout",
        "item": "Battery Model",
        "excel_val": ex_bat or "None",
        "dwg_val": dwg_bat or "None",
        "status": status_bat,
        "notes": note_bat
    })
    
    # Battery Quantity
    ex_bat_qty = excel_data.get("battery_qty", 0)
    dwg_bat_qty = p1.get("battery_qty", 0)
    
    if ex_bat_qty > 0 or dwg_bat_qty > 0:
        bat_qty_match = (ex_bat_qty == dwg_bat_qty)
        status_bq = "PASS" if bat_qty_match else "FAIL"
        note_bq = f"Battery quantity matches ({ex_bat_qty} Unit)." if bat_qty_match else f"MISMATCH: Excel requests {ex_bat_qty} units, but DWG drawings {dwg_bat_qty} units!"
    else:
        status_bq = "PASS"
        note_bq = "No battery in both files (0 Units)."
        
    page1_results.append({
        "category": "Page 1: PV Layout",
        "item": "Battery Quantity",
        "excel_val": f"{ex_bat_qty} Unit" if ex_bat_qty else "0 Unit (None)",
        "dwg_val": f"{dwg_bat_qty} Unit" if dwg_bat_qty else "0 Unit (None)",
        "status": status_bq,
        "notes": note_bq
    })

    # ----------------------------------------------------
    # 3. PAGE 2: MOUNTING STRUCTURE DETAILED BOM TABLE
    # ----------------------------------------------------
    p2 = pdf_data.get("page2", {})
    bom = p2.get("bom_items", {})
    
    active_panels = dwg_qty if dwg_qty > 0 else ex_qty
    str_count = p1.get("strings_count", 2)
    jumpers = manual_jumpers if manual_jumpers is not None else p1.get("jumper_lines_count", 0)
    expected_mc4 = str_count + jumpers
    
    page2_bom_rows = []
    
    gl_dwg = p2.get("grounding_lug", 0)
    groups = gl_dwg if gl_dwg > 0 else 4

    # 1. End Clamp: (4 x Groups) + 2 Buffer
    expected_ec = (4 * groups) + 2
    ec_dwg = p2.get("end_clamp", 0)
    ec_match = (ec_dwg == expected_ec) if ec_dwg > 0 else False
    page2_bom_rows.append({
        "Component": "1. End Clamp",
        "DWG Quantity": ec_dwg,
        "SOP / Expected Calculation": f"{expected_ec} pcs ((4 × {groups} Groups) + 2 Buffer)",
        "Status": "✅ TALLY" if ec_match else ("❌ MISMATCH" if ec_dwg > 0 else "⚠️ WARNING"),
        "Notes": f"Matches SOP formula: (4 × {groups} Groups) + 2 Buffer = {expected_ec} pcs." if ec_match else f"DWG shows {ec_dwg} pcs, expected {expected_ec} pcs by formula!"
    })

    # 2. Mid Clamp: (2 x (Panels - Groups)) + 2 Buffer
    expected_mc = (2 * max(active_panels - groups, 0)) + 2
    mc_dwg = p2.get("mid_clamp", 0)
    mc_match = (mc_dwg == expected_mc) if mc_dwg > 0 else False
    page2_bom_rows.append({
        "Component": "2. Mid Clamp",
        "DWG Quantity": mc_dwg,
        "SOP / Expected Calculation": f"{expected_mc} pcs (2 × ({active_panels} Panels - {groups}) + 2 Buffer)",
        "Status": "✅ TALLY" if mc_match else ("❌ MISMATCH" if mc_dwg > 0 else "⚠️ WARNING"),
        "Notes": f"Matches SOP formula: (2 × {active_panels - groups} Gaps) + 2 Buffer = {expected_mc} pcs." if mc_match else f"DWG shows {mc_dwg} pcs, expected {expected_mc} pcs by formula!"
    })

    # 3. Roof Attachment Check: Tile Hook, L-Foot, U-Rail, Universal Cliplock, Shingle Plate, Concrete Ballast
    # User Engineering Formulas:
    # - Tile Hook, L-Foot, U-Rail, Universal Cliplock, Shingle Plate = (End Clamp + Mid Clamp) - 4 Buffer
    #   (subtracting 2 buffer each from End Clamp and Mid Clamp)
    # - If custom panel layout is present (Gambar 1), drafter counts attachment symbols directly, so follow DWG PDF.
    # - For Concrete Ballast: Follow DWG PDF directly.
    # - For L-Foot when used with Concrete Ballast: L-Foot Quantity = Concrete Ballast Quantity * 2.
    
    th_dwg = p2.get("tile_hook", 0)
    clip_dwg = p2.get("cliplock", 0)
    urail_dwg = p2.get("u_rail", 0)
    lfoot_dwg = p2.get("l_foot", 0)
    shingle_dwg = p2.get("shingle", 0)
    ballast_dwg = p2.get("ballast", 0)
    att_name_dwg = p2.get("roof_attachment_name", "")
    att_qty_dwg = p2.get("roof_attachment_qty", 0)
    ex_roof = excel_data.get("roof_type", "").upper()

    # Determine attachment type
    if roof_name_override and roof_name_override != "Keep Original" and roof_name_override.strip():
        detected_type = roof_name_override.strip()
    elif ballast_dwg > 0 or "BALLAST" in att_name_dwg.upper():
        detected_type = "Concrete Ballast"
    elif urail_dwg > 0 or "U-RAIL" in att_name_dwg.upper() or "URAIL" in att_name_dwg.upper():
        detected_type = "U-Rail"
    elif clip_dwg > 0 or "CLIPLOCK" in att_name_dwg.upper():
        detected_type = "Universal Cliplock"
    elif lfoot_dwg > 0 or "L-FOOT" in att_name_dwg.upper() or "L FOOT" in att_name_dwg.upper():
        detected_type = "L-Foot"
    elif shingle_dwg > 0 or "SHINGLE" in att_name_dwg.upper():
        detected_type = "Shingle Plate"
    elif th_dwg > 0 or "TILE" in att_name_dwg.upper():
        detected_type = "Tile Hook"
    elif att_name_dwg:
        detected_type = att_name_dwg
    else:
        if "METAL" in ex_roof:
            detected_type = "Universal Cliplock"
        elif "SHINGLE" in ex_roof:
            detected_type = "Shingle Plate"
        elif "BALLAST" in ex_roof or "CONCRETE" in ex_roof:
            detected_type = "Concrete Ballast"
        else:
            detected_type = "Tile Hook"

    # Determine detected DWG quantity
    if "BALLAST" in detected_type.upper():
        att_dwg = ballast_dwg or att_qty_dwg
    elif "U-RAIL" in detected_type.upper() or "URAIL" in detected_type.upper():
        att_dwg = urail_dwg or att_qty_dwg
    elif "CLIPLOCK" in detected_type.upper():
        att_dwg = clip_dwg or att_qty_dwg
    elif "L-FOOT" in detected_type.upper() or "L FOOT" in detected_type.upper():
        att_dwg = lfoot_dwg or att_qty_dwg
    elif "SHINGLE" in detected_type.upper():
        att_dwg = shingle_dwg or att_qty_dwg
    elif "TILE" in detected_type.upper():
        att_dwg = th_dwg or att_qty_dwg
    else:
        att_dwg = att_qty_dwg or th_dwg or clip_dwg or urail_dwg or lfoot_dwg or shingle_dwg or ballast_dwg

    if table_overrides:
        for k_ov, v_ov in table_overrides.items():
            if str(v_ov).strip().isdigit() and any(x in k_ov.upper() for x in ["TILE", "HOOK", "3", "ROOF", "U-RAIL", "CLIPLOCK", "L-FOOT", "BALLAST", "SHINGLE"]):
                att_dwg = int(v_ov)

    # Evaluate SOP engineering calculation
    if "BALLAST" in detected_type.upper():
        # Concrete Ballast: Follow DWG PDF directly
        page2_bom_rows.append({
            "Component": f"3. {detected_type}",
            "DWG Quantity": att_dwg,
            "SOP / Expected Calculation": f"{att_dwg if att_dwg > 0 else 'Follow DWG'} blocks (Follow DWG Layout)",
            "Status": "✅ TALLY" if att_dwg > 0 else "⚠️ WARNING",
            "Notes": f"Concrete Ballast follows DWG PDF layout directly ({att_dwg} blocks)." if att_dwg > 0 else "No Concrete Ballast quantity detected on DWG drawing!"
        })
        
        # L-Foot with Concrete Ballast: L-Foot * 2 from quantity ballast
        expected_lfoot = att_dwg * 2
        lfoot_val = lfoot_dwg if lfoot_dwg > 0 else (att_dwg * 2 if att_dwg > 0 else 0)
        lfoot_match = (lfoot_dwg == expected_lfoot) if lfoot_dwg > 0 else True
        page2_bom_rows.append({
            "Component": "3b. L-Foot (Ballast Mounting)",
            "DWG Quantity": lfoot_dwg if lfoot_dwg > 0 else f"{expected_lfoot} (Formula)",
            "SOP / Expected Calculation": f"{expected_lfoot} pcs (2 × {att_dwg} Concrete Ballast)",
            "Status": "✅ TALLY" if lfoot_match else "❌ MISMATCH",
            "Notes": f"Matches Concrete Ballast SOP: 2 × {att_dwg} Ballast = {expected_lfoot} pcs L-Foot." if lfoot_match else f"DWG shows {lfoot_dwg} pcs, expected {expected_lfoot} pcs (2 × {att_dwg} Ballast blocks)!"
        })
    else:
        # Standard roof attachments: Tile Hook, L-Foot, U-Rail, Universal Cliplock, Shingle Plate
        # Formula: (End Clamp + Mid Clamp) - 4 Buffer
        sop_att_qty = max(0, (expected_ec + expected_mc) - 4)
        dwg_clamp_calc = max(0, (ec_dwg + mc_dwg) - 4) if (ec_dwg > 0 and mc_dwg > 0) else sop_att_qty
        
        is_formula_match = (att_dwg == sop_att_qty or att_dwg == dwg_clamp_calc) if att_dwg > 0 else False
        th_symbols = p2.get("tile_hook_count_symbols", 0)
        is_symbol_match = (th_symbols > 0 and att_dwg == th_symbols)
        
        if is_formula_match:
            att_status = "✅ TALLY"
            att_notes = f"Matches SOP formula: (End Clamp + Mid Clamp) - 4 Buffer = ({expected_ec} + {expected_mc}) - 4 = {sop_att_qty} pcs."
        elif is_symbol_match:
            att_status = "✅ TALLY"
            att_notes = f"Matches DWG layout symbols directly ({att_dwg} units on drawing layout - Gambar 1)."
        elif att_dwg > 0:
            att_status = "❌ MISMATCH"
            att_notes = f"DWG shows {att_dwg} pcs, expected {sop_att_qty} pcs by formula: (End Clamp + Mid Clamp) - 4 Buffer = ({expected_ec} + {expected_mc}) - 4 = {sop_att_qty} pcs. (Follows DWG layout if custom panel arrangement - Gambar 1)."
        else:
            att_status = "⚠️ WARNING"
            att_notes = "No roof attachment quantity detected on DWG drawing!"
            
        page2_bom_rows.append({
            "Component": f"3. {detected_type}",
            "DWG Quantity": att_dwg,
            "SOP / Expected Calculation": f"{sop_att_qty} pcs ((End Clamp + Mid Clamp) - 4 Buffer)",
            "Status": att_status,
            "Notes": att_notes
        })

    # 4. Grounding Clip: 1:1 Panel Ratio
    gc_dwg = p2.get("grounding_clip", 0)
    gc_status = "PASS" if (gc_dwg == active_panels and active_panels > 0) else ("FAIL" if gc_dwg > 0 else "WARNING")
    page2_bom_rows.append({
        "Component": "4. Grounding Clip",
        "DWG Quantity": gc_dwg,
        "SOP / Expected Calculation": f"{active_panels} pcs (1:1 Panel Ratio)",
        "Status": "✅ TALLY" if gc_status == "PASS" else ("❌ MISMATCH" if gc_status == "FAIL" else "⚠️ WARNING"),
        "Notes": f"Matches 1:1 panel ratio ({active_panels} pcs)." if gc_status == "PASS" else f"DWG value ({gc_dwg}) does not match panel quantity ({active_panels})!"
    })

    # 5. Grounding Lug: 1 unit per array group
    page2_bom_rows.append({
        "Component": "5. Grounding Lug",
        "DWG Quantity": gl_dwg,
        "SOP / Expected Calculation": f"{groups} pcs (1 unit per panel array group)",
        "Status": "✅ TALLY" if gl_dwg > 0 else "⚠️ WARNING",
        "Notes": f"Matches ({gl_dwg} pcs for {groups} separate array groups)." if gl_dwg > 0 else "No Grounding Lug specified on drawing!"
    })

    # 6. MC4 Connector: Strings + Jumpers
    mc4_dwg = p2.get("mc4_connector", 0)
    mc4_status = "PASS" if (mc4_dwg == expected_mc4 and expected_mc4 > 0) else "WARNING"
    page2_bom_rows.append({
        "Component": "6. MC4 Connector",
        "DWG Quantity": mc4_dwg,
        "SOP / Expected Calculation": f"{expected_mc4} pcs ({str_count} Strings + {jumpers} Jumpers)",
        "Status": "✅ TALLY" if mc4_status == "PASS" else "⚠️ WARNING",
        "Notes": f"Matches SOP formula ({str_count} Strings + {jumpers} Jumpers = {expected_mc4} pcs)." if mc4_status == "PASS" else f"DWG shows {mc4_dwg} pcs, verify strings and jumper lines."
    })

    # Calculate SOP R-Railing & Splice (Gambar 2 SOP Formula)
    row_panels = []
    str_configs = p1.get("string_config", [])
    for sc in str_configs:
        m = re.search(r"(\d+)\s*PCS", sc, re.I)
        if m:
            row_panels.append(int(m.group(1)))
            
    if not row_panels:
        num_rows = max(groups, 2)
        base_p = active_panels // num_rows
        rem = active_panels % num_rows
        row_panels = [base_p + (1 if i < rem else 0) for i in range(num_rows) if (base_p + (1 if i < rem else 0)) > 0]
        
    sop_layout_bars = 0
    sop_splices = 0
    for k in row_panels:
        if k <= 0: continue
        L_row = (k * 1.134) + 0.10
        bars_row = math.ceil((L_row * 2.0) / 2.0)
        splices_row = 2 * max(0, math.ceil(L_row / 2.0) - 1)
        sop_layout_bars += bars_row
        sop_splices += splices_row
        
    sop_total_railing = sop_layout_bars + 1 # +1 office buffer standard

    # 7. R-Railing: Follow DWG drawing layout directly
    rr_dwg = p2.get("r_railing", 0)
    if rr_dwg > 0:
        rr_sop = f"{rr_dwg} pcs (Follow DWG Layout)"
        rr_status = "✅ TALLY"
        rr_notes = f"Mengikut susunan fizikal layout CAD DWG ({rr_dwg} pcs) berasaskan SOP AutoLISP AUTOMOUNT."
    else:
        rr_sop = f"{sop_total_railing} pcs (SOP: {sop_layout_bars} Layout + 1 Buffer)"
        rr_status = "⚠️ WARNING"
        rr_notes = "Tiada kuantiti R-Railing dinyatakan dalam lukisan DWG!"

    page2_bom_rows.append({
        "Component": "7. R-Railing",
        "DWG Quantity": rr_dwg,
        "SOP / Expected Calculation": rr_sop,
        "Status": rr_status,
        "Notes": rr_notes
    })

    # 8. Splice: Follow DWG drawing layout directly
    sp_dwg = p2.get("splice", 0)
    if sp_dwg > 0:
        sp_sop = f"{sp_dwg} pcs (Follow DWG Layout)"
        sp_status = "✅ TALLY"
        sp_notes = f"Mengikut susunan fizikal layout CAD DWG ({sp_dwg} pcs) berasaskan SOP AutoLISP AUTOMOUNT."
    else:
        sp_sop = f"{sop_splices} pcs (Rail connector SOP)"
        sp_status = "⚠️ WARNING"
        sp_notes = "Tiada kuantiti Splice dinyatakan dalam lukisan DWG!"

    page2_bom_rows.append({
        "Component": "8. Splice",
        "DWG Quantity": sp_dwg,
        "SOP / Expected Calculation": sp_sop,
        "Status": sp_status,
        "Notes": sp_notes
    })

    # Dynamic display of all additional System Information items (Item 10 L-foot, etc.)
    # matching DWG Page 2 drawing exactly (Gambar 3)
    raw_bom = p2.get("raw_bom_table", [])
    handled_items = ["1", "2", "3", "4", "5", "6", "7", "8"]
    
    for r in raw_bom:
        item_no = str(r.get("item_no", "")).strip()
        item_name = str(r.get("name", "")).strip()
        item_qty = r.get("qty", 0)
        
        if item_no not in handled_items:
            # Check if this item is L-Foot (e.g. Item 10 L-foot)
            if "L-FOOT" in item_name.upper() or "L FOOT" in item_name.upper():
                # L-Foot SOP Engineering calculation:
                # - Concrete Ballast: Ballast * 2
                # - Shingle Plate / Tile Hook / standard: (End Clamp + Mid Clamp) - 4 Buffer
                if "BALLAST" in detected_type.upper():
                    sop_lf = att_dwg * 2
                    lf_match = (item_qty == sop_lf) if item_qty > 0 else True
                    page2_bom_rows.append({
                        "Component": f"{item_no}. {item_name}",
                        "DWG Quantity": item_qty,
                        "SOP / Expected Calculation": f"{sop_lf} pcs (2 × {att_dwg} Concrete Ballast)",
                        "Status": "✅ TALLY" if lf_match else "❌ MISMATCH",
                        "Notes": f"Matches Concrete Ballast SOP: 2 × {att_dwg} Ballast = {sop_lf} pcs L-Foot." if lf_match else f"DWG shows {item_qty} pcs, expected {sop_lf} pcs (2 × {att_dwg} Ballast blocks)!"
                    })
                else:
                    sop_lf = max(0, (expected_ec + expected_mc) - 4)
                    lf_match = (item_qty == sop_lf) if item_qty > 0 else True
                    page2_bom_rows.append({
                        "Component": f"{item_no}. {item_name}",
                        "DWG Quantity": item_qty,
                        "SOP / Expected Calculation": f"{sop_lf} pcs ((End Clamp + Mid Clamp) - 4 Buffer)",
                        "Status": "✅ TALLY" if lf_match else "❌ MISMATCH",
                        "Notes": f"Matches SOP formula: (End Clamp + Mid Clamp) - 4 Buffer = ({expected_ec} + {expected_mc}) - 4 = {sop_lf} pcs." if lf_match else f"DWG shows {item_qty} pcs, expected {sop_lf} pcs by formula!"
                    })
            else:
                # General additional DWG accessories
                page2_bom_rows.append({
                    "Component": f"{item_no}. {item_name}",
                    "DWG Quantity": item_qty,
                    "SOP / Expected Calculation": f"{item_qty} pcs (Follow DWG Layout)",
                    "Status": "✅ TALLY",
                    "Notes": f"Matches DWG drawing layout ({item_qty} pcs)."
                })

    # If L-Foot was found in DWG (lfoot_dwg > 0) but not yet added in page2_bom_rows
    if lfoot_dwg > 0 and not any("L-FOOT" in str(row.get("Component", "")).upper() or "L FOOT" in str(row.get("Component", "")).upper() for row in page2_bom_rows):
        sop_lf = (att_dwg * 2) if "BALLAST" in detected_type.upper() else max(0, (expected_ec + expected_mc) - 4)
        lf_match = (lfoot_dwg == sop_lf)
        page2_bom_rows.append({
            "Component": "10. L-foot",
            "DWG Quantity": lfoot_dwg,
            "SOP / Expected Calculation": f"{sop_lf} pcs ((End Clamp + Mid Clamp) - 4 Buffer)",
            "Status": "✅ TALLY" if lf_match else "❌ MISMATCH",
            "Notes": f"Matches SOP formula: (End Clamp + Mid Clamp) - 4 Buffer = ({expected_ec} + {expected_mc}) - 4 = {sop_lf} pcs." if lf_match else f"DWG shows {lfoot_dwg} pcs, expected {sop_lf} pcs by formula!"
        })


    # ----------------------------------------------------
    # 4. PAGE 3 & 4: SLD, INVERTER CURRENT & BREAKER SPECS
    # ----------------------------------------------------
    p3 = pdf_data.get("page3", {})
    p4 = pdf_data.get("page4", {})
    
    active_inv = dwg_inv or ex_inv
    inv_rule = match_inverter_spec(active_inv)
    
    page3_4_results = []
    
    # Inverter Current Check (A)
    actual_amp = p4.get("inverter_amp", 0.0)
    if inv_rule and inv_rule.get("current_a"):
        exp_amp = inv_rule["current_a"]
        amp_match = (actual_amp == exp_amp) if actual_amp > 0 else True
        page3_4_results.append({
            "Category": "Inverter Specs (Page 4)",
            "Check Item": "Inverter Output Current (A)",
            "Inverter Model": inv_rule["tag"],
            "Expected Value (SOP)": f"{exp_amp} A",
            "DWG Value": f"{actual_amp} A" if actual_amp else "No Data",
            "Status": "✅ TALLY" if (actual_amp == exp_amp) else ("❌ MISMATCH" if actual_amp > 0 else "⚠️ WARNING"),
            "Notes": f"Matches standard model {inv_rule['tag']} ({exp_amp}A)." if (actual_amp == exp_amp) else f"DWG displays {actual_amp}A, expected {exp_amp}A for {inv_rule['tag']}!"
        })
        
    # Inverter Breaker MCB Check
    actual_mcb = p4.get("inverter_mcb", "")
    if inv_rule and inv_rule.get("mcb"):
        exp_mcb = inv_rule["mcb"]
        mcb_match = False
        if actual_mcb:
            actual_u = actual_mcb.upper()
            mcb_match = (inv_rule["poles"] in actual_u and exp_mcb.split()[0] in actual_u)
            
        page3_4_results.append({
            "Category": "Circuit Protection (Page 4)",
            "Check Item": "Inverter Breaker (MCB)",
            "Inverter Model": inv_rule["tag"],
            "Expected Value (SOP)": exp_mcb,
            "DWG Value": actual_mcb or "No Data",
            "Status": "✅ TALLY" if mcb_match else ("❌ MISMATCH" if actual_mcb else "⚠️ WARNING"),
            "Notes": f"Matches standard {inv_rule['tag']} ({exp_mcb})." if mcb_match else f"DWG specifies {actual_mcb}, expected {exp_mcb} for {inv_rule['tag']}!"
        })
        
    # PV Meter & CT Rating Check
    actual_pvm = p4.get("pv_meter", "")
    if inv_rule and inv_rule.get("meter"):
        exp_meter = inv_rule["meter"]
        exp_val = "40" if "40" in exp_meter else ("25" if "25" in exp_meter else "32")
        pvm_match = False
        if actual_pvm:
            actual_u = actual_pvm.upper()
            pvm_match = (exp_val in actual_u and ("CT" in actual_u or "VA" in actual_u))
            
        page3_4_results.append({
            "Category": "PV Metering (Page 4)",
            "Check Item": "PV Meter & CT Rating",
            "Inverter Model": inv_rule["tag"],
            "Expected Value (SOP)": exp_meter,
            "DWG Value": actual_pvm or "No Data",
            "Status": "✅ TALLY" if pvm_match else "⚠️ WARNING",
            "Notes": f"Matches CT rating {exp_meter}." if pvm_match else f"Verify CT meter rating on Page 4 ({exp_meter})."
        })
        
    # Supply Phase Check (Single Phase vs Three Phase)
    pdf_phase = p3.get("phase", "Unknown")
    exp_phase = inv_rule.get("phase", "Three Phase") if inv_rule else "Three Phase"
    phase_match = (pdf_phase == exp_phase) if pdf_phase != "Unknown" else True
    
    page3_4_results.append({
        "Category": "Electrical Phase (Page 3)",
        "Check Item": "System Phase (SLD)",
        "Inverter Model": inv_rule["tag"] if inv_rule else "-",
        "Expected Value (SOP)": exp_phase,
        "DWG Value": pdf_phase,
        "Status": "✅ TALLY" if phase_match else "❌ MISMATCH",
        "Notes": f"SLD diagram matches {exp_phase} system." if phase_match else f"Phase mismatch: DWG '{pdf_phase}' vs standard '{exp_phase}'!"
    })

    return page1_results, page2_bom_rows, page3_4_results
