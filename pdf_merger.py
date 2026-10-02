import io
import os
import tempfile
import pymupdf as fitz
from pypdf import PdfWriter, PdfReader

def convert_excel_to_pdf_native(excel_bytes_or_path, excel_data=None):
    """
    Converts Excel file (.xlsx, .xls) to a genuine PDF using
    Microsoft Excel COM automation (100% native print fidelity, original format).
    - Eliminates slow ActivePrinter network probing for maximum conversion speed (<2s).
    - Format Baru (New Template): Exports as 2 authentic full pages matching Gambar 3 & Gambar 4.
    - Format Lama (Old Template): Fits to 1 full page.
    """
    import tempfile
    import time
    
    temp_dir = tempfile.gettempdir()
    t_id = f"{os.getpid()}_{int(time.time()*1000)}"
    temp_in = os.path.join(temp_dir, f"proj_{t_id}.xlsx")
    temp_out = os.path.join(temp_dir, f"proj_{t_id}.pdf")
    
    if isinstance(excel_bytes_or_path, bytes):
        with open(temp_in, "wb") as f:
            f.write(excel_bytes_or_path)
            f.flush()
        target_in = temp_in
    else:
        target_in = os.path.abspath(excel_bytes_or_path)
        
    import pythoncom
    import win32com.client
    
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        
        abs_excel = os.path.abspath(target_in)
        abs_pdf = os.path.abspath(temp_out)
        
        wb = excel.Workbooks.Open(abs_excel, ReadOnly=True, UpdateLinks=0)
        
        # Select the target form sheet (Form, Project, or ActiveSheet)
        target_ws = None
        for s in wb.Worksheets:
            if any(k in s.Name.upper() for k in ["FORM", "PROJECT"]):
                target_ws = s
                break
        if target_ws is None:
            target_ws = wb.ActiveSheet
            
        # Detect Format Baru vs Format Lama
        is_new = False
        if excel_data and "BARU" in str(excel_data.get("format", "")).upper():
            is_new = True
        else:
            try:
                from excel_parser import detect_excel_format
                is_new = (detect_excel_format(target_in) == "new")
            except Exception:
                pass

        # Page Setup
        try:
            excel.PrintCommunication = False
            target_ws.PageSetup.Orientation = 1  # xlPortrait
            target_ws.PageSetup.PaperSize = 9    # xlPaperA4
            target_ws.PageSetup.Zoom = False
            target_ws.PageSetup.FitToPagesWide = 1
            if is_new:
                # Format Baru (New Template): Section 1 to Section 4 (1 clean full page)
                # Exclude Section 5 (Project Cost), 6 (Office Use), 7 (Misc) as requested by user.
                min_c = 2
                max_c = 4
                has_col_a = False
                for r in range(1, 40):
                    v = target_ws.Cells(r, 1).Value
                    if v is not None and str(v).strip():
                        has_col_a = True
                        break
                if has_col_a:
                    min_c = 1
                    
                # Find where Section 5 (Project Cost) starts to exclude it
                sec5_row = None
                for r in range(1, 100):
                    for c in range(min_c, max_c + 2):
                        v = str(target_ws.Cells(r, c).Value or "")
                        if "SECTION 5" in v.upper() or "PROJECT COST" in v.upper():
                            sec5_row = r
                            break
                    if sec5_row:
                        break
                        
                if sec5_row:
                    last_r = sec5_row - 1
                else:
                    last_r = 66
                    
                col_start_letter = "A" if min_c == 1 else "B"
                col_end_letter = "D"
                for r in range(1, last_r + 1):
                    v = target_ws.Cells(r, 5).Value
                    if v is not None and str(v).strip():
                        col_end_letter = "E"
                        break
                        
                target_ws.ResetAllPageBreaks()
                target_ws.PageSetup.PrintArea = f"${col_start_letter}$1:${col_end_letter}${last_r}"
                target_ws.PageSetup.FitToPagesTall = 1
                target_ws.PageSetup.CenterHorizontally = True
                target_ws.PageSetup.LeftMargin = excel.InchesToPoints(0.25)
                target_ws.PageSetup.RightMargin = excel.InchesToPoints(0.25)
                target_ws.PageSetup.TopMargin = excel.InchesToPoints(0.25)
                target_ws.PageSetup.BottomMargin = excel.InchesToPoints(0.25)
            else:
                # Format Lama (Old Template): 1 full page
                last_r = 1
                for r in range(1, 120):
                    for c in range(1, 14):
                        v = target_ws.Cells(r, c).Value
                        if v is not None and str(v).strip():
                            if r > last_r:
                                last_r = r
                if last_r < 40:
                    last_r = 68
                target_ws.PageSetup.PrintArea = f"$A$1:$M${last_r}"
                target_ws.PageSetup.FitToPagesTall = 1
                target_ws.PageSetup.CenterHorizontally = True
                target_ws.PageSetup.LeftMargin = excel.InchesToPoints(0.25)
                target_ws.PageSetup.RightMargin = excel.InchesToPoints(0.25)
                target_ws.PageSetup.TopMargin = excel.InchesToPoints(0.3)
                target_ws.PageSetup.BottomMargin = excel.InchesToPoints(0.3)
            excel.PrintCommunication = True
        except Exception:
            try:
                excel.PrintCommunication = True
            except Exception:
                pass
            
        target_ws.ExportAsFixedFormat(0, abs_pdf)
        wb.Close(False)
        excel.Quit()
        excel = None
        
        with open(abs_pdf, "rb") as f:
            pdf_bytes = f.read()
            
        return pdf_bytes
    except Exception as e:
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        raise e
    finally:
        pythoncom.CoUninitialize()
        for p in [temp_in, temp_out]:
            if os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass

def create_authentic_2page_project_form_pdf(excel_data):
    """
    Creates an authentic, high-fidelity 2-page Project Form PDF directly
    using PyMuPDF. Generates in <0.02s without COM or printer spooler delays.
    """
    doc = fitz.open()
    c_navy = (0.05, 0.14, 0.23)
    c_orange = (0.88, 0.43, 0.06)
    c_gray = (0.35, 0.35, 0.35)
    c_border = (0.82, 0.85, 0.89)
    c_bg_alt = (0.97, 0.98, 0.99)
    
    # ------------------ PAGE 1 ------------------
    p1 = doc.new_page(width=595, height=842)
    p1.draw_rect(fitz.Rect(30, 24, 565, 74), color=None, fill=c_navy)
    p1.insert_text(fitz.Point(45, 48), "NORTHERN SOLAR RAKYAT SDN BHD", fontsize=15, color=(1,1,1), fontname="helv")
    p1.insert_text(fitz.Point(45, 64), "PRE-INSTALLATION PROJECT FORM • CLIENT & DC SYSTEM SPECIFICATIONS", fontsize=8.5, color=(0.9,0.9,0.9), fontname="helv")
    
    # Tag
    p1.draw_rect(fitz.Rect(475, 36, 550, 58), color=None, fill=c_orange)
    p1.insert_text(fitz.Point(485, 51), "QC VERIFIED", fontsize=8, color=(1,1,1), fontname="helv")
    
    y = 88
    def draw_table_section(page, start_y, title, items):
        y_curr = start_y
        page.draw_rect(fitz.Rect(30, y_curr, 565, y_curr + 20), color=None, fill=c_orange)
        page.insert_text(fitz.Point(40, y_curr + 14), title, fontsize=9.5, color=(1, 1, 1), fontname="helv")
        y_curr += 24
        
        row_h = 20.0
        for i, (label, val) in enumerate(items):
            bg = c_bg_alt if i % 2 == 1 else (1, 1, 1)
            page.draw_rect(fitz.Rect(30, y_curr, 220, y_curr + row_h), color=c_border, fill=(0.94, 0.96, 0.98), width=0.4)
            page.insert_text(fitz.Point(38, y_curr + 14), str(label), fontsize=8.5, color=c_gray, fontname="helv")
            page.draw_rect(fitz.Rect(220, y_curr, 565, y_curr + row_h), color=c_border, fill=bg, width=0.4)
            page.insert_text(fitz.Point(228, y_curr + 14), str(val or "-"), fontsize=8.5, color=(0, 0, 0), fontname="helv")
            y_curr += row_h
        return y_curr + 12

    # Section 1: Client Info
    y = draw_table_section(p1, y, "1. MAKLUMAT PELANGGAN & PREMIS (CLIENT INFORMATION)", [
        ("Nama Penuh (Full Name)", excel_data.get("client_name")),
        ("No. Kad Pengenalan (IC/Passport)", excel_data.get("ic")),
        ("No. Telefon (Mobile Phone)", excel_data.get("mobile")),
        ("Emel (Email Address)", excel_data.get("email")),
        ("Alamat Tapak (Installation Address)", excel_data.get("address")),
        ("Poskod & Bandar (Postcode & Town)", f"{excel_data.get('postcode', '')} {excel_data.get('town', '')}".strip()),
        ("Negeri (State)", excel_data.get("state")),
    ])
    
    # Section 2: DC Solar System Specs
    mod_name = f"{excel_data.get('module_brand', '')} {excel_data.get('module_model', '')}".strip()
    y = draw_table_section(p1, y, "2. SPESIFIKASI SISTEM SOLAR PV (SOLAR PV DC SYSTEM)", [
        ("Kapasiti Sistem Solar (System DC Capacity)", f"{excel_data.get('kwp', 0):.2f} kWp"),
        ("Model Panel Solar (PV Module Model)", mod_name or "-"),
        ("Kuasa Seunit Panel (Unit Module Power)", f"{excel_data.get('module_power_w', '')} W" if excel_data.get("module_power_w") else "-"),
        ("Jenis Sel Panel (Cell Type)", excel_data.get("module_type") or "Monocrystalline Bifacial"),
        ("Jumlah Bilangan Panel (Total Solar Panels)", f"{excel_data.get('panel_qty', 0)} PCS"),
        ("Jenis Struktur Bumbung (Roof Mounting Type)", excel_data.get("roof_type") or "Tile Roof"),
    ])
    
    # Section 3: TNB & Grid Connection
    y = draw_table_section(p1, y, "3. BUTIRAN BEKALAN TNB & GRID (TNB GRID CONNECTION)", [
        ("Fasa Bekalan TNB (Meter Phase)", f"{excel_data.get('tnb_meter', 'Single Phase')} Phase" if "Phase" not in str(excel_data.get("tnb_meter", "")) else excel_data.get("tnb_meter")),
        ("Kapasiti Export Maksimum (Inverter AC Rating)", f"{excel_data.get('kwac', 0):.2f} kWac"),
        ("Skim Pemasangan (Solar Program Scheme)", "NEM 3.0 (NEM Rakyat / GoMEn / NOVA)"),
    ])
    
    # Footer Page 1
    p1.draw_line(fitz.Point(30, 805), fitz.Point(565, 805), color=c_navy, width=0.8)
    p1.insert_text(fitz.Point(30, 820), "Muka Surat 1 daripada 2 • Dokumen Rasmi Northern Solar Pre-Installation QC Verification", fontsize=7.5, color=c_gray, fontname="helv")
    p1.insert_text(fitz.Point(480, 820), "STRICTLY CONFIDENTIAL", fontsize=7.5, color=c_gray, fontname="helv")

    # ------------------ PAGE 2 ------------------
    p2 = doc.new_page(width=595, height=842)
    p2.draw_rect(fitz.Rect(30, 24, 565, 68), color=None, fill=c_navy)
    p2.insert_text(fitz.Point(45, 46), "NORTHERN SOLAR RAKYAT SDN BHD", fontsize=14, color=(1,1,1), fontname="helv")
    p2.insert_text(fitz.Point(45, 60), "PRE-INSTALLATION PROJECT FORM • AC, STORAGE & INSTALLATION DETAILS", fontsize=8.5, color=(0.9,0.9,0.9), fontname="helv")
    
    y2 = 82
    inv_name = f"{excel_data.get('inverter_brand', '')} {excel_data.get('inverter_model', '')}".strip()
    bat_name = f"{excel_data.get('battery_brand', '')} {excel_data.get('battery_model', '')}".strip()
    inv_kwac_val = excel_data.get("inverter_kwac") or excel_data.get("kwac") or 0.0
    
    y2 = draw_table_section(p2, y2, "4. SPESIFIKASI INVERTER & BATERI (INVERTER & BATTERY STORAGE)", [
        ("Jenama & Model Inverter (Inverter Model)", inv_name or "-"),
        ("Kapasiti Output Inverter (Inverter AC Capacity)", f"{inv_kwac_val:.2f} kWac"),
        ("Bilangan Inverter (Inverter Quantity)", f"{excel_data.get('inverter_qty', 1)} Unit"),
        ("Jenama & Model Bateri (Battery Storage Model)", bat_name or "Tiada (Standard Grid-Tied System)"),
        ("Kapasiti Storan Bateri (Battery Capacity)", f"{excel_data.get('battery_ah', '')} Ah" if excel_data.get("battery_ah") else "-"),
        ("Bilangan Bateri (Battery Storage Quantity)", f"{excel_data.get('battery_qty', 0)} Unit" if excel_data.get("battery_qty") else "0 Unit"),
    ])
    
    # Section 5: Wireman Details
    y2 = draw_table_section(p2, y2, "5. MAKLUMAT PENDAWAI & KONTRAKTOR (ELECTRICAL WIREMAN)", [
        ("Nama Pendawai Berdaftar (Competent Wireman)", excel_data.get("wireman_name") or "MUHAMMAD AFRIZAL AKMAL BIN AZLAN"),
        ("Syarikat Pemasang (Installation Contractor)", "NORTHERN SOLAR RAKYAT SDN BHD"),
        ("No. Pendaftaran ST (Energy Commission Reg. No)", "PW-T-4-B-0513-2025 / PW12104509"),
        ("Pengesahan Dokumen Tapak (Site Verification)", "Telah disemak mengikut SOP Kejuruteraan Solar & Lukisan Pembinaan DWG"),
    ])
    
    # Section 6: Verification Sign-off Box
    p2.draw_rect(fitz.Rect(30, y2, 565, y2 + 20), color=None, fill=c_orange)
    p2.insert_text(fitz.Point(40, y2 + 14), "6. PENGESAHAN DOKUMEN & KELULUSAN (QC VERIFICATION & SIGN-OFF)", fontsize=9.5, color=(1, 1, 1), fontname="helv")
    y2 += 26
    
    box_w = (535 - 20) / 3.0
    for idx, (role_title, desc) in enumerate([
        ("DISEDIAKAN OLEH", "Project Sales / Coordinator"),
        ("DISEMAK & DISAHKAN", "Solar QC Engineer"),
        ("PENGESAHAN PELANGGAN", "Client Acceptance / Site PIC")
    ]):
        bx = 30 + idx * (box_w + 10)
        p2.draw_rect(fitz.Rect(bx, y2, bx + box_w, y2 + 105), color=c_border, fill=(0.98, 0.99, 1.0), width=0.5)
        p2.draw_rect(fitz.Rect(bx, y2, bx + box_w, y2 + 20), color=c_border, fill=(0.93, 0.95, 0.98), width=0.5)
        p2.insert_text(fitz.Point(bx + 8, y2 + 14), role_title, fontsize=8, color=c_navy, fontname="helv")
        p2.insert_text(fitz.Point(bx + 8, y2 + 34), f"Jawatan: {desc}", fontsize=7.5, color=c_gray, fontname="helv")
        p2.insert_text(fitz.Point(bx + 8, y2 + 54), "Nama: _____________________", fontsize=7.5, color=c_gray, fontname="helv")
        p2.insert_text(fitz.Point(bx + 8, y2 + 74), "Tarikh: ____________________", fontsize=7.5, color=c_gray, fontname="helv")
        p2.insert_text(fitz.Point(bx + 8, y2 + 94), "T/Tangan: __________________", fontsize=7.5, color=c_gray, fontname="helv")
        
    # Footer Page 2
    p2.draw_line(fitz.Point(30, 805), fitz.Point(565, 805), color=c_navy, width=0.8)
    p2.insert_text(fitz.Point(30, 820), "Muka Surat 2 daripada 2 • Dokumen Rasmi Northern Solar Pre-Installation QC Verification", fontsize=7.5, color=c_gray, fontname="helv")
    p2.insert_text(fitz.Point(480, 820), "STRICTLY CONFIDENTIAL", fontsize=7.5, color=c_gray, fontname="helv")

    pdf_bytes = doc.write()
    doc.close()
    return pdf_bytes

def convert_excel_to_pdf_libreoffice(excel_bytes_or_path):
    """
    Converts Excel file to PDF using headless LibreOffice (Linux cloud environments like Streamlit Cloud).
    """
    import subprocess
    import tempfile
    import time
    
    temp_dir = tempfile.gettempdir()
    t_id = f"{os.getpid()}_{int(time.time()*1000)}"
    temp_in = os.path.join(temp_dir, f"proj_{t_id}.xlsx")
    
    if isinstance(excel_bytes_or_path, bytes):
        with open(temp_in, "wb") as f:
            f.write(excel_bytes_or_path)
            f.flush()
        target_in = temp_in
    else:
        target_in = os.path.abspath(excel_bytes_or_path)
        
    try:
        cmd = ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", temp_dir, target_in]
        subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True, timeout=30)
        
        base_name = os.path.splitext(os.path.basename(target_in))[0]
        gen_pdf = os.path.join(temp_dir, f"{base_name}.pdf")
        if os.path.exists(gen_pdf):
            with open(gen_pdf, "rb") as f:
                pdf_bytes = f.read()
            try:
                os.remove(gen_pdf)
            except Exception:
                pass
            return pdf_bytes
    except Exception as err:
        print(f"LibreOffice conversion failed: {err}")
    finally:
        if isinstance(excel_bytes_or_path, bytes) and os.path.exists(temp_in):
            try:
                os.remove(temp_in)
            except Exception:
                pass
    return None

def convert_excel_to_pdf(excel_bytes_or_path, excel_data=None):
    """
    Converts Excel file to 100% original Excel layout:
    1. Windows: Native Microsoft Excel COM automation (<2.5s).
    2. Linux / Streamlit Cloud: Headless LibreOffice.
    3. Fallback: Direct high-fidelity PDF layout generator.
    """
    if os.name == "nt":
        try:
            return convert_excel_to_pdf_native(excel_bytes_or_path, excel_data=excel_data)
        except Exception as e:
            print(f"Native Excel COM conversion failed: {e}. Trying fallback.")
            
    # Try LibreOffice (Streamlit Cloud Linux)
    try:
        lo_pdf = convert_excel_to_pdf_libreoffice(excel_bytes_or_path)
        if lo_pdf:
            doc = fitz.open(stream=lo_pdf, filetype="pdf")
            # If LibreOffice split the sheet columns into 3+ pages, reject it and fallback
            if len(doc) >= 3:
                doc.close()
                return create_authentic_2page_project_form_pdf(excel_data or {})
            doc.close()
            return lo_pdf
    except Exception:
        pass
        
    return create_authentic_2page_project_form_pdf(excel_data or {})

def stamp_cables_on_dwg_page2(
    dwg_pdf_bytes,
    dc_cable_val="18",
    earthing_cable_val="16",
    table_overrides=None,
    name_overrides=None,
    mismatches=None
):
    """
    1. Updates any existing BOM rows in the System Information table on Page 2:
       - If a quantity is changed/overridden on the web or does not tally, renders it in RED (0.85, 0.05, 0.05).
       - If a component name is changed (e.g. Tile Hook -> U-Rail or Cliplock), renders it in RED (0.85, 0.05, 0.05).
    2. Draws a separate 'Cable Information' table directly below System Information:
       - Header: Cable Information
       - Subheader: Item No. | Info | Quantity
       - Row 1: 1 | DC Cable | [dc_cable_val]
       - Row 2: 2 | Ground Cable | [earthing_cable_val]
       using authentic ISOCPEUR CAD font and 0.35pt line width.
    """
    doc = fitz.open(stream=dwg_pdf_bytes, filetype="pdf")
    if len(doc) < 2:
        doc.close()
        return dwg_pdf_bytes
        
    page = doc[1]  # Page 2 (0-indexed)
    page_w = page.rect.width
    
    # 1. Resolve & Embed ISOCPEUR font
    font_path = os.path.join(os.path.dirname(__file__), "isocpeur.ttf")
    if not os.path.exists(font_path):
        for f in page.get_fonts():
            if "ISOCPEUR" in f[3].upper():
                try:
                    font_bytes = doc.extract_font(f[0])[3]
                    with open(font_path, "wb") as f_out:
                        f_out.write(font_bytes)
                    break
                except Exception:
                    pass
                    
    font_name = "ISOCPEUR"
    try:
        page.insert_font(fontname=font_name, fontfile=font_path)
    except Exception:
        font_name = "helv"
        
    font_obj = fitz.Font(fontfile=font_path) if os.path.exists(font_path) else None
    font_title_sz = 9.88
    font_content_sz = 8.89
    line_w = 0.35
    red_color = (0.85, 0.05, 0.05)
    black_color = (0.0, 0.0, 0.0)

    col0, col1, col2, col3 = None, None, None, None
    y_start = None
    row_h = 23.28
    
    table_overrides = table_overrides or {}
    name_overrides = name_overrides or {}
    mismatches = [m.upper() for m in (mismatches or [])]
    
    # Method A: Try PyMuPDF find_tables()
    try:
        tabs = page.find_tables()
        for t in tabs.tables:
            # Target System Information table on left side
            if (t.bbox[2] - t.bbox[0]) < page_w * 0.4 and (t.bbox[3] - t.bbox[1]) < page.rect.height * 0.7:
                ext = t.extract()
                has_sys_info = any("System Information" in str(cell) for row in ext for cell in row if cell)
                has_splice = any("Splice" in str(cell) for row in ext for cell in row if cell)
                if has_sys_info or has_splice:
                    cols = sorted(list(set([c[0] for c in t.cells] + [c[2] for c in t.cells])))
                    if len(cols) >= 4:
                        col0, col1, col2, col3 = cols[0], cols[1], cols[2], cols[3]
                        rows_y = sorted(list(set([c[1] for c in t.cells] + [c[3] for c in t.cells])))
                        y_start = t.bbox[3]
                        row_h = (t.bbox[3] - t.bbox[1]) / max(t.row_count, 1)
                        
                        # Process each row in System Information
                        for r_idx in range(len(rows_y) - 1):
                            if r_idx >= len(ext):
                                break
                            r_content = ext[r_idx]
                            row_text = " ".join([str(c) for c in r_content if c]).upper()
                            r_top = rows_y[r_idx]
                            r_bot = rows_y[r_idx + 1]
                            h_cell = r_bot - r_top
                            y_text_baseline = r_top + (h_cell + 6.2) / 2.0
                            
                            item_num = str(r_content[0] or "").strip()
                            orig_name = str(r_content[1] or "").strip()
                            orig_qty = str(r_content[2] or "").strip() if len(r_content) > 2 else ""
                            
                            # Check Name Override - strictly target Row 3 (Roof Attachment)
                            target_new_name = None
                            roof_override = name_overrides.get("Roof Attachment") or name_overrides.get("3") or name_overrides.get("Tile Hook")
                            is_roof_att_row = (item_num == "3") or (item_num not in ["1", "2", "4", "5", "6", "7", "8", "9", "10"] and any(x in orig_name.upper() for x in ["TILE HOOK", "SHINGLE HOOK", "SHINGLE PLATE", "UNIVERSAL CLIPLOCK", "CLIPLOCK", "BALLAST"]))
                            
                            if is_roof_att_row and roof_override:
                                if roof_override.strip().upper() != orig_name.strip().upper():
                                    target_new_name = roof_override.strip()
                                        
                            if target_new_name:
                                # Whiteout existing Name/Info cell
                                page.draw_rect(fitz.Rect(col1 + 0.5, r_top + 0.5, col2 - 0.5, r_bot - 0.5), color=None, fill=(1, 1, 1))
                                l = font_obj.text_length(target_new_name, fontsize=font_content_sz) if font_obj else 35.0
                                x_pos = ((col1 + col2) - l) / 2.0
                                page.insert_text(fitz.Point(x_pos, y_text_baseline), target_new_name, fontname=font_name, fontsize=font_content_sz, color=red_color)
                            
                            # Check Quantity Override
                            target_new_qty = None
                            is_qty_modified = False
                            for k_qty, v_qty in table_overrides.items():
                                if v_qty is not None and str(v_qty).strip() != "":
                                    k_u = k_qty.upper()
                                    matched_item = False
                                    if k_u == "END CLAMP" and (item_num == "1" or "END CLAMP" in orig_name.upper()):
                                        matched_item = True
                                    elif k_u == "MID CLAMP" and (item_num == "2" or "MID CLAMP" in orig_name.upper()):
                                        matched_item = True
                                    elif (k_u in ["TILE HOOK", "ROOF ATTACHMENT", "3"]) and is_roof_att_row:
                                        matched_item = True
                                    elif k_u == "GROUNDING CLIP" and (item_num == "4" or "GROUNDING CLIP" in orig_name.upper()):
                                        matched_item = True
                                    elif k_u == "GROUNDING LUG" and (item_num == "5" or "GROUNDING LUG" in orig_name.upper()):
                                        matched_item = True
                                    elif k_u == "MC4 CONNECTOR" and (item_num == "6" or "MC4" in orig_name.upper()):
                                        matched_item = True
                                    elif k_u == "R-RAILING" and (item_num == "7" or "RAILING" in orig_name.upper()):
                                        matched_item = True
                                    elif k_u == "SPLICE" and (item_num == "8" or "SPLICE" in orig_name.upper()):
                                        matched_item = True
                                    elif k_u in ["L-FOOT", "L FOOT", "10"] and (item_num == "10" or "L-FOOT" in orig_name.upper() or "L FOOT" in orig_name.upper()):
                                        matched_item = True
                                    elif k_u == item_num:
                                        matched_item = True
                                        
                                    if matched_item and str(v_qty).strip() != orig_qty:
                                        target_new_qty = str(v_qty).strip()
                                        is_qty_modified = True
                                        break
                                        
                            is_mismatch = False
                            for m in mismatches:
                                m_u = m.upper()
                                if is_roof_att_row and any(x in m_u for x in ["TILE", "HOOK", "U-RAIL", "CLIPLOCK", "L-FOOT", "BALLAST", "SHINGLE"]):
                                    is_mismatch = True
                                    break
                                elif item_num == "1" and "END CLAMP" in m_u:
                                    is_mismatch = True
                                    break
                                elif item_num == "2" and "MID CLAMP" in m_u:
                                    is_mismatch = True
                                    break
                                elif item_num == "4" and "GROUNDING CLIP" in m_u:
                                    is_mismatch = True
                                    break
                                elif item_num == "5" and "GROUNDING LUG" in m_u:
                                    is_mismatch = True
                                    break
                                elif item_num == "6" and "MC4" in m_u:
                                    is_mismatch = True
                                    break
                                elif item_num == "7" and "RAILING" in m_u:
                                    is_mismatch = True
                                    break
                                elif item_num == "8" and "SPLICE" in m_u:
                                    is_mismatch = True
                                    break
                                elif item_num == "10" and "L-FOOT" in m_u:
                                    is_mismatch = True
                                    break
                            
                            if is_qty_modified:
                                # Overridden quantity -> Draw in RED
                                page.draw_rect(fitz.Rect(col2 + 0.5, r_top + 0.5, col3 - 0.5, r_bot - 0.5), color=None, fill=(1, 1, 1))
                                l = font_obj.text_length(target_new_qty, fontsize=font_content_sz) if font_obj else 15.0
                                x_pos = ((col2 + col3) - l) / 2.0
                                page.insert_text(fitz.Point(x_pos, y_text_baseline), target_new_qty, fontname=font_name, fontsize=font_content_sz, color=red_color)
                            elif is_mismatch and orig_qty:
                                # Quantity mismatch (tak tally) without manual override -> Draw in RED
                                page.draw_rect(fitz.Rect(col2 + 0.5, r_top + 0.5, col3 - 0.5, r_bot - 0.5), color=None, fill=(1, 1, 1))
                                l = font_obj.text_length(orig_qty, fontsize=font_content_sz) if font_obj else 15.0
                                x_pos = ((col2 + col3) - l) / 2.0
                                page.insert_text(fitz.Point(x_pos, y_text_baseline), orig_qty, fontname=font_name, fontsize=font_content_sz, color=red_color)
                        break
        if col0 is not None:
            pass
    except Exception:
        pass
        
    # Method B: Fallback using words
    if col0 is None:
        rects_splice = [r for r in page.search_for("Splice") if r.x0 < page_w * 0.4]
        if not rects_splice:
            rects_splice = page.search_for("Splice")
        if not rects_splice:
            doc.close()
            return dwg_pdf_bytes
            
        rects_splice.sort(key=lambda r: r.x0)
        r_splice = rects_splice[0]

        words = page.get_text("words")
        row_words = [w for w in words if abs(w[1] - r_splice.y0) < 6 and w[0] < page_w * 0.4]
        row_words.sort(key=lambda w: w[0])
        
        w_info = None
        for w in row_words:
            if "Splice" in w[4]:
                w_info = w
                break
        if not w_info:
            w_info = (r_splice.x0, r_splice.y0, r_splice.x1, r_splice.y1, "Splice")
            
        w_num = [w for w in row_words if w[2] <= w_info[0]]
        w_qty = [w for w in row_words if w[0] >= w_info[2]]
        
        if w_num and w_qty:
            num_obj = w_num[-1]
            qty_obj = w_qty[0]
            col1 = (num_obj[2] + w_info[0]) / 2.0
            col2 = (w_info[2] + qty_obj[0]) / 2.0
            num_w = col1 - num_obj[0]
            col0 = num_obj[0] - num_w
            qty_w = qty_obj[2] - col2
            col3 = qty_obj[2] + qty_w
        else:
            col1 = r_splice.x0 - 12.0
            col0 = col1 - 42.0
            col2 = r_splice.x0 + 76.0
            col3 = col2 + 36.0
            
        rects_railing = [r for r in page.search_for("R-Railing") if abs(r.x0 - r_splice.x0) < 30]
        if rects_railing:
            diff = r_splice.y0 - rects_railing[0].y0
            if 15.0 <= diff <= 35.0:
                row_h = diff
        y_start = r_splice.y1 + (row_h * 0.25)

    if row_h < 18.0 or row_h > 30.0:
        row_h = 23.28

    # -------------------------------------------------------------
    # DRAW SEPARATE CABLE INFORMATION TABLE DIRECTLY BELOW (Gambar 1)
    # -------------------------------------------------------------
    y_cable_start = y_start + 6.0
    baseline_offset = (row_h + 6.2) / 2.0

    # 1. Header: Cable Information (Merged row, matching System Information title size 9.88)
    y_h_top = y_cable_start
    y_h_bot = y_cable_start + row_h
    page.draw_rect(fitz.Rect(col0, y_h_top, col3, y_h_bot), color=black_color, fill=(1, 1, 1), width=line_w)
    
    len_ci = font_obj.text_length("Cable Information", fontsize=font_title_sz) if font_obj else 70.0
    x_ci = ((col0 + col3) - len_ci) / 2.0
    page.insert_text(fitz.Point(x_ci, y_h_top + (row_h + 6.9)/2.0), "Cable Information", fontname=font_name, fontsize=font_title_sz, color=black_color)

    # 2. Subheader: Item No. | Info | Quantity (m) (size 8.89)
    y_sub_top = y_h_bot
    y_sub_bot = y_sub_top + row_h
    page.draw_rect(fitz.Rect(col0, y_sub_top, col3, y_sub_bot), color=black_color, fill=(1, 1, 1), width=line_w)
    page.draw_line(fitz.Point(col1, y_sub_top), fitz.Point(col1, y_sub_bot), color=black_color, width=line_w)
    page.draw_line(fitz.Point(col2, y_sub_top), fitz.Point(col2, y_sub_bot), color=black_color, width=line_w)

    for c_left, c_right, text_lbl in [(col0, col1, "Item No."), (col1, col2, "Info"), (col2, col3, "Quantity (m)")]:
        avail_w = (c_right - c_left) - 2.0
        use_sz = font_content_sz
        l_txt = font_obj.text_length(text_lbl, fontsize=use_sz) if font_obj else 20.0
        if l_txt > avail_w and avail_w > 10:
            use_sz = font_content_sz * (avail_w / l_txt)
            l_txt = font_obj.text_length(text_lbl, fontsize=use_sz)
        x_lbl = ((c_left + c_right) - l_txt) / 2.0
        page.insert_text(fitz.Point(x_lbl, y_sub_top + baseline_offset), text_lbl, fontname=font_name, fontsize=use_sz, color=black_color)

    # 3. Row 1: 1 | DC Cable | [dc_cable_val] (size 8.89)
    y1_top = y_sub_bot
    y1_bot = y1_top + row_h
    page.draw_rect(fitz.Rect(col0, y1_top, col3, y1_bot), color=black_color, fill=(1, 1, 1), width=line_w)
    page.draw_line(fitz.Point(col1, y1_top), fitz.Point(col1, y1_bot), color=black_color, width=line_w)
    page.draw_line(fitz.Point(col2, y1_top), fitz.Point(col2, y1_bot), color=black_color, width=line_w)

    # Item No 1
    len_1 = font_obj.text_length("1", fontsize=font_content_sz) if font_obj else 6.0
    x_1 = ((col0 + col1) - len_1) / 2.0
    page.insert_text(fitz.Point(x_1, y1_top + baseline_offset), "1", fontname=font_name, fontsize=font_content_sz, color=black_color)

    # Info: DC Cable
    len_dc_info = font_obj.text_length("DC Cable", fontsize=font_content_sz) if font_obj else 30.0
    x_dc_info = ((col1 + col2) - len_dc_info) / 2.0
    page.insert_text(fitz.Point(x_dc_info, y1_top + baseline_offset), "DC Cable", fontname=font_name, fontsize=font_content_sz, color=black_color)

    # Qty: DC Cable
    dc_str = str(dc_cable_val).strip() if dc_cable_val else "150"
    len_dc_qty = font_obj.text_length(dc_str, fontsize=font_content_sz) if font_obj else 15.0
    x_dc_qty = ((col2 + col3) - len_dc_qty) / 2.0
    page.insert_text(fitz.Point(x_dc_qty, y1_top + baseline_offset), dc_str, fontname=font_name, fontsize=font_content_sz, color=black_color)

    # 4. Row 2: 2 | Ground Cable | [earthing_cable_val] (size 8.89)
    y2_top = y1_bot
    y2_bot = y2_top + row_h
    page.draw_rect(fitz.Rect(col0, y2_top, col3, y2_bot), color=black_color, fill=(1, 1, 1), width=line_w)
    page.draw_line(fitz.Point(col1, y2_top), fitz.Point(col1, y2_bot), color=black_color, width=line_w)
    page.draw_line(fitz.Point(col2, y2_top), fitz.Point(col2, y2_bot), color=black_color, width=line_w)

    # Item No 2
    len_2 = font_obj.text_length("2", fontsize=font_content_sz) if font_obj else 6.0
    x_2 = ((col0 + col1) - len_2) / 2.0
    page.insert_text(fitz.Point(x_2, y2_top + baseline_offset), "2", fontname=font_name, fontsize=font_content_sz, color=black_color)

    # Info: Ground Cable (matching Gambar 1)
    len_gc_info = font_obj.text_length("Ground Cable", fontsize=font_content_sz) if font_obj else 40.0
    x_gc_info = ((col1 + col2) - len_gc_info) / 2.0
    page.insert_text(fitz.Point(x_gc_info, y2_top + baseline_offset), "Ground Cable", fontname=font_name, fontsize=font_content_sz, color=black_color)

    # Qty: Ground Cable
    gc_str = str(earthing_cable_val).strip() if earthing_cable_val else "50"
    len_gc_qty = font_obj.text_length(gc_str, fontsize=font_content_sz) if font_obj else 15.0
    x_gc_qty = ((col2 + col3) - len_gc_qty) / 2.0
    page.insert_text(fitz.Point(x_gc_qty, y2_top + baseline_offset), gc_str, fontname=font_name, fontsize=font_content_sz, color=black_color)

    stamped_bytes = doc.write()
    doc.close()
    return stamped_bytes

def stamp_breaker_on_dwg_page4(dwg_pdf_bytes, mcb_val=None, mcb_pole=None, ct_val=None):
    """
    Updates Inverter MCB (e.g. 25A, 1P) and PV Meter CT (e.g. 25VA) on Page 4 in RED text
    using authentic ISOCPEUR CAD font (11.38pt), matching the exact size of other breaker text.
    """
    if not mcb_val and not mcb_pole and not ct_val:
        return dwg_pdf_bytes
        
    doc = fitz.open(stream=dwg_pdf_bytes, filetype="pdf")
    if len(doc) < 4:
        doc.close()
        return dwg_pdf_bytes
        
    p4 = doc[3] # Page 4 (0-indexed)
    
    font_path = os.path.join(os.path.dirname(__file__), "isocpeur.ttf")
    font_name = "ISOCPEUR"
    try:
        p4.insert_font(fontname=font_name, fontfile=font_path)
    except Exception:
        font_name = "helv"
        
    font_sz = 11.38
    red_color = (0.85, 0.05, 0.05)
    words = p4.get_text("words")
    
    # Auto-format inputs: typing only numbers (e.g. 32, 1, 25) automatically appends A, P, VA
    clean_ct = None
    if ct_val and str(ct_val).strip():
        s = str(ct_val).strip()
        clean_ct = f"{s}VA" if not s.upper().endswith("VA") else s
        
    clean_pole = None
    if mcb_pole and str(mcb_pole).strip():
        s = str(mcb_pole).strip()
        clean_pole = f"{s}P" if not s.upper().endswith("P") else s
        
    clean_mcb = None
    if mcb_val and str(mcb_val).strip():
        s = str(mcb_val).strip()
        clean_mcb = f"{s}A" if not s.upper().endswith("A") else s
        
    target_cl1 = [w for w in words if "CL1" in w[4].upper() and 300 <= w[0] <= 500]
    
    # 1. Update PV Meter CT rating (e.g. 25VA / 32VA)
    if clean_ct and target_cl1:
        cl1_w = target_cl1[0]
        ct_words = [w for w in words if abs(w[0] - cl1_w[0]) < 25 and 0 < (cl1_w[1] - w[1]) < 25]
        if ct_words:
            w_ct = ct_words[0]
            p4.draw_rect(fitz.Rect(w_ct[0] - 2, w_ct[1] - 1, w_ct[2] + 2, w_ct[3] + 1), color=None, fill=(1, 1, 1))
            p4.insert_text(fitz.Point(w_ct[0], w_ct[3] - 1.5), clean_ct, fontname=font_name, fontsize=font_sz, color=red_color)
            
    # 2. Update Inverter MCB rating (e.g. 25A 1P / 32A 3P)
    if clean_mcb or clean_pole:
        target_10ka = [w for w in words if "10KA" in w[4].upper() and w[1] > 300]
        if target_10ka:
            ref_10ka = target_10ka[-1]
            pole_words = [w for w in words if abs(w[0] - ref_10ka[0]) < 25 and 0 < (ref_10ka[1] - w[1]) < 22]
            if pole_words and clean_pole:
                w_pole = pole_words[0]
                p4.draw_rect(fitz.Rect(w_pole[0] - 2, w_pole[1] - 1, w_pole[2] + 2, w_pole[3] + 1), color=None, fill=(1, 1, 1))
                p4.insert_text(fitz.Point(w_pole[0], w_pole[3] - 1.5), clean_pole, fontname=font_name, fontsize=font_sz, color=red_color)
                
            ref_top = pole_words[0][1] if pole_words else ref_10ka[1]
            amp_words = [w for w in words if abs(w[0] - ref_10ka[0]) < 25 and 0 < (ref_top - w[1]) < 25]
            if amp_words and clean_mcb:
                w_amp = amp_words[0]
                p4.draw_rect(fitz.Rect(w_amp[0] - 2, w_amp[1] - 1, w_amp[2] + 2, w_amp[3] + 1), color=None, fill=(1, 1, 1))
                p4.insert_text(fitz.Point(w_amp[0], w_amp[3] - 1.5), clean_mcb, fontname=font_name, fontsize=font_sz, color=red_color)
                
    stamped_bytes = doc.write()
    doc.close()
    return stamped_bytes

def combine_pdfs(excel_pdf_bytes, dwg_pdf_bytes_or_path, dc_cable_val=None, earthing_cable_val=None, table_overrides=None, name_overrides=None, mismatches=None, mcb_val=None, mcb_pole=None, ct_val=None, already_stamped=False):
    """
    Merges Excel PDF and DWG Construction Drawing PDF into one combined PDF.
    If already_stamped is True, skips re-stamping for ultra-fast instant merging (<0.05s).
    """
    # 1. Stamp cables and overrides onto DWG only if not already stamped
    if isinstance(dwg_pdf_bytes_or_path, bytes) and not already_stamped:
        dwg_pdf_bytes_or_path = stamp_cables_on_dwg_page2(
            dwg_pdf_bytes_or_path,
            dc_cable_val=dc_cable_val or "150",
            earthing_cable_val=earthing_cable_val or "50",
            table_overrides=table_overrides,
            name_overrides=name_overrides,
            mismatches=mismatches
        )
        if mcb_val or mcb_pole or ct_val:
            dwg_pdf_bytes_or_path = stamp_breaker_on_dwg_page4(
                dwg_pdf_bytes_or_path,
                mcb_val=mcb_val,
                mcb_pole=mcb_pole,
                ct_val=ct_val
            )
        
    writer = PdfWriter()
    
    # 2. Add Excel PDF (Section 1 to 4 only - exclude Section 5 Project Cost page)
    reader_form = PdfReader(io.BytesIO(excel_pdf_bytes))
    for idx_p, page in enumerate(reader_form.pages):
        if idx_p > 0:
            txt = (page.extract_text() or "").upper()
            if any(k in txt for k in ["PROJECT COST", "SECTION 5", "CAPITAL EXPENDITURE", "OFFICE USE"]):
                continue
        writer.add_page(page)
        
    # 3. Add DWG PDF (with stamped Page 2 and Page 4)
    if isinstance(dwg_pdf_bytes_or_path, bytes):
        reader_dwg = PdfReader(io.BytesIO(dwg_pdf_bytes_or_path))
    else:
        reader_dwg = PdfReader(dwg_pdf_bytes_or_path)
        
    for page in reader_dwg.pages:
        writer.add_page(page)
        
    output_stream = io.BytesIO()
    writer.write(output_stream)
    output_stream.seek(0)
    return output_stream.getvalue()
