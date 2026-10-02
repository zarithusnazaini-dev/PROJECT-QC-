import streamlit as st
import pandas as pd
import io
import os
import re

from excel_parser import parse_project_excel
from pdf_checker import (
    extract_pdf_data,
    render_pdf_page_image,
    extract_dwg_system_info_image,
    extract_dwg_pv_layout_image,
    extract_dwg_page2_table_image
)
from qc_engine import run_qc_check
from pdf_merger import convert_excel_to_pdf, combine_pdfs, stamp_cables_on_dwg_page2, stamp_breaker_on_dwg_page4


# High-performance caching functions
@st.cache_data(show_spinner=False)
def get_cached_excel_data(excel_bytes):
    return parse_project_excel(io.BytesIO(excel_bytes))

@st.cache_data(show_spinner=False)
def get_cached_pdf_data(pdf_bytes, hint_client_name=None):
    return extract_pdf_data(pdf_bytes, hint_client_name=hint_client_name)

@st.cache_data(show_spinner=False)
def get_cached_rendered_page(pdf_bytes, page_num, dpi=140):
    return render_pdf_page_image(pdf_bytes, page_num=page_num, dpi=dpi)

@st.cache_data(show_spinner=False)
def get_cached_pv_layout(pdf_bytes, dpi=200):
    return extract_dwg_pv_layout_image(pdf_bytes, dpi=dpi)

@st.cache_data(show_spinner=False)
def get_cached_sys_info_img(pdf_bytes, dpi=200):
    return extract_dwg_system_info_image(pdf_bytes, dpi=dpi)

@st.cache_data(show_spinner=False)
def get_cached_page2_table_zoom(pdf_bytes, dpi=220):
    return extract_dwg_page2_table_image(pdf_bytes, dpi=dpi)

@st.cache_data(show_spinner=False)
def get_cached_excel_pdf(excel_bytes, fmt_name=None):
    return convert_excel_to_pdf(excel_bytes, excel_data={"format": fmt_name} if fmt_name else None)

@st.cache_data(show_spinner=False)
def get_cached_stamped_dwg(pdf_bytes, dc_cable, earth_cable, t_overrides, n_overrides, mismatches_tuple, mcb_amp, mcb_pole, ct_rating):
    stamped = stamp_cables_on_dwg_page2(
        pdf_bytes,
        dc_cable_val=dc_cable,
        earthing_cable_val=earth_cable,
        table_overrides=t_overrides,
        name_overrides=n_overrides,
        mismatches=list(mismatches_tuple) if mismatches_tuple else []
    )
    if mcb_amp or mcb_pole or ct_rating:
        stamped = stamp_breaker_on_dwg_page4(
            stamped,
            mcb_val=mcb_amp or None,
            mcb_pole=mcb_pole or None,
            ct_val=ct_rating or None
        )
    return stamped


# Page configuration
st.set_page_config(
    page_title="Northern Solar - QC & PDF Combiner",
    page_icon="☀️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for modern styling
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #0d233a 0%, #1a365d 100%);
        padding: 22px 28px;
        border-radius: 12px;
        color: white;
        margin-bottom: 25px;
        border-left: 6px solid #e06d10;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
    }
    .metric-card {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px;
        text-align: center;
    }
    .metric-label {
        font-size: 12px;
        font-weight: 600;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .metric-val {
        font-size: 20px;
        font-weight: 700;
        color: #0f172a;
        margin-top: 5px;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        font-weight: 600;
        font-size: 15px;
        padding: 10px 18px;
    }
    .img-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 15px;
    }
</style>
""", unsafe_allow_html=True)

# Main Header
st.markdown("""
<div class="main-header">
    <h2 style="margin: 0; padding: 0;">☀️ NORTHERN SOLAR RAKYAT</h2>
    <h4 style="margin: 6px 0 0 0; font-weight: 400; opacity: 0.95;">Solar Pre-Installation QC & DWG PDF Combiner</h4>
    <p style="margin: 6px 0 0 0; font-size: 13px; opacity: 0.85;">Automated Cross-Verification of Project Form vs Construction Drawings • Panel Quantity Auto-Calculation • Inverter & Breaker SOP Validation • Authentic Excel to PDF Conversion</p>
</div>
""", unsafe_allow_html=True)

# Initialize session state for clean client cycling
if "doc_cycle" not in st.session_state:
    st.session_state["doc_cycle"] = 0

# Sidebar
with st.sidebar:
    st.header("📂 Document Upload")
    
    # Button to upload another document and clear previous client
    if st.button("🔄 Upload Another Document / Clear All", use_container_width=True, type="primary", help="Clears current files, manual overrides, and resets the application fresh for the next client."):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.cache_data.clear()
        st.rerun()
    
    cycle = st.session_state.get("doc_cycle", 0)
    
    excel_file = st.file_uploader(
        "1. Project Form (Excel)",
        type=["xlsx", "xls", "csv"],
        key=f"excel_file_{cycle}",
        help="Supports both New Template (key-value structured) and Old Template (Remark parsed)."
    )
    
    pdf_file = st.file_uploader(
        "2. DWG Drawings (PDF)",
        type=["pdf"],
        key=f"pdf_file_{cycle}",
        help="Exported PDF containing Construction Drawings (Pages 1 to 4+)."
    )
    
    # Auto-detect if new files are uploaded for a different client -> clear previous overrides automatically
    if excel_file and pdf_file:
        current_doc_id = f"{excel_file.name}_{pdf_file.name}"
        if st.session_state.get("last_doc_id") != current_doc_id:
            if st.session_state.get("last_doc_id") is not None:
                st.session_state["doc_cycle"] = cycle + 1
                st.session_state["last_doc_id"] = current_doc_id
                st.rerun()
            else:
                st.session_state["last_doc_id"] = current_doc_id
                
    cycle = st.session_state.get("doc_cycle", 0)
    
    st.markdown("---")
    st.subheader("⚙️ DWG Page 2: System Information (1 to 8)")
    st.caption("Values here follow the exact order of the DWG Page 2 System Information table. Leave blank if DWG is already correct (only modified values or mismatches appear in RED):")
    
    with st.expander("✏️ Modify System Information Items & Quantities", expanded=False):
        ov_end_clamp = st.text_input("1. End Clamp (Qty):", placeholder="e.g. 18", key=f"ov_end_clamp_{cycle}")
        ov_mid_clamp = st.text_input("2. Mid Clamp (Qty):", placeholder="e.g. 30", key=f"ov_mid_clamp_{cycle}")
        
        col_th1, col_th2 = st.columns([1.2, 1.0])
        with col_th1:
            roof_name_choice = st.selectbox(
                "3. Roof Attachment Name:",
                ["Keep Original", "Tile Hook", "U-Rail", "Universal Cliplock", "L-Foot", "Shingle Plate", "Concrete Ballast", "Custom..."],
                key=f"roof_name_{cycle}"
            )
            if roof_name_choice == "Custom...":
                roof_name_choice = st.text_input("Custom Roof Attachment Name:", placeholder="e.g. U-Rail", key=f"roof_custom_{cycle}")
        with col_th2:
            ov_tile_hook_qty = st.text_input("3. Roof Att. (Qty):", placeholder="e.g. 32", key=f"ov_tile_hook_{cycle}")
            
        ov_grounding_clip = st.text_input("4. Grounding Clip (Qty):", placeholder="e.g. 18", key=f"ov_gclip_{cycle}")
        ov_grounding_lug = st.text_input("5. Grounding Lug (Qty):", placeholder="e.g. 4", key=f"ov_glug_{cycle}")
        ov_mc4 = st.text_input("6. MC4 Connector (Qty):", placeholder="e.g. 4", key=f"ov_mc4_{cycle}")
        ov_railing = st.text_input("7. R-Railing (Qty):", placeholder="e.g. 23", key=f"ov_railing_{cycle}")
        ov_splice = st.text_input("8. Splice (Qty):", placeholder="e.g. 16", key=f"ov_splice_{cycle}")
        ov_lfoot = st.text_input("10. L-Foot (Qty - optional):", placeholder="e.g. 54", key=f"ov_lfoot_{cycle}")
        
    st.markdown("---")
    st.subheader("🔌 DWG Page 2: Cable Information (Gambar 1)")
    st.caption("Stamped as a dedicated separate table directly below System Information:")
    dc_cable_input = st.text_input(
        "1. DC Cable Quantity (m):",
        value="18",
        key=f"dc_cable_{cycle}",
        help="Quantity (m) stamped into Row 1 of the Cable Information table on DWG Page 2."
    )
    earthing_cable_input = st.text_input(
        "2. Ground Cable Quantity (m):",
        value="16",
        key=f"earth_cable_{cycle}",
        help="Quantity (m) stamped into Row 2 of the Cable Information table on DWG Page 2."
    )
    
    st.markdown("---")
    st.subheader("⚡ DWG Page 4: Inverter Breaker & PV Meter (Optional)")
    st.caption("Modify Inverter MCB or PV Meter CT values on Page 4. Only modified values that differ from DWG will be rendered in **RED** text:")
    with st.expander("✏️ Modify Breaker & PV Meter Specs", expanded=False):
        col_mcb1, col_mcb2 = st.columns(2)
        with col_mcb1:
            ov_mcb_amp = st.text_input("Inverter MCB Rating (A):", placeholder="e.g. 25 or 32", key=f"ov_mcb_amp_{cycle}")
        with col_mcb2:
            ov_mcb_pole = st.text_input("MCB Pole (P):", placeholder="e.g. 1 or 3", key=f"ov_mcb_pole_{cycle}")
        ov_ct_rating = st.text_input("PV Meter CT Rating (VA):", placeholder="e.g. 25 or 32", key=f"ov_ct_rating_{cycle}")
        
    st.markdown("---")
    st.subheader("📐 Engineering & SOP Verification")
    jumper_input = st.number_input(
        "Jumper Line Count (Green/Cyan Lines):",
        min_value=0,
        max_value=10,
        value=0,
        help="SOP Rule: MC4 Connector = String Count + Jumper Line Count (+1 per jumper line)."
    )
    highlight_mismatches = st.checkbox(
        "Highlight QC mismatches in RED on DWG Page 2",
        value=True,
        help="When enabled, any item whose quantity doesn't tally with SOP calculation will automatically appear in RED text on DWG Page 2."
    )

if not excel_file or not pdf_file:
    st.markdown("""
    ### 👋 Welcome to Solar Pre-Installation QC & PDF Combiner!
    Please upload the **Project Form (Excel)** and **DWG Drawings (PDF)** using the left sidebar to start verification:
    
    1. **Automatic Panel Formula Calculation:**
       * If the Project Form only specifies system capacity (e.g. `8.82 kWp`), the system automatically extracts the module rating (e.g. `630W` $\\rightarrow$ `0.630 kW`) and computes:
         $$\\text{Panel Quantity} = \\frac{8.82\\text{ kWp}}{0.630} = 14\\text{ PCS}$$
    2. **Mounting Structure BOM & Cable Information Table (Page 2):**
       * Validates SOP engineering formulas: Grounding Clip ($1:1$), MC4 Connector ($\\text{Strings} + \\text{Jumpers}$), End Clamps ($(4 \\times G) + 2$), Mid Clamps ($2 \\times (P - G) + 2$), and R-Railing / Splice bar formulas.
       * Stamps dedicated **Cable Information** table directly below System Information: Row 1 (**DC Cable**) & Row 2 (**Ground Cable**).
       * Automatically renders changes and mismatches in **RED** on the DWG Page 2 table before download.
    3. **SLD & Circuit Protection Verification (Pages 3 & 4):**
       * Automatically cross-checks inverter rated current (A), MCB breaker ratings (`25A 1P`, `32A 1P`, `32A 3P`, `40A 3P`), CT meter rating (`25VA`, `32VA` / `40VA`), and single/three-phase electrical alignment.
       * Modified breakers on Page 4 appear in authentic CAD **RED** text matching original font and size.
    4. **High-Performance Delivery PDF Combiner:**
       * Uses optimized native Excel conversion and instant memory merging (<0.05s) to generate a complete client delivery package.
    """)
    st.stop()

# Process data
with st.spinner("Processing Excel Project Form and DWG PDF drawings..."):
    try:
        excel_bytes = excel_file.getvalue()
        excel_data = get_cached_excel_data(excel_bytes)
        
        pdf_bytes = pdf_file.getvalue()
        pdf_data = get_cached_pdf_data(pdf_bytes, hint_client_name=excel_data.get("client_name"))
        
        # Construct table overrides & name overrides
        table_overrides = {}
        if ov_end_clamp.strip(): table_overrides["End Clamp"] = ov_end_clamp.strip()
        if ov_mid_clamp.strip(): table_overrides["Mid Clamp"] = ov_mid_clamp.strip()
        if ov_tile_hook_qty.strip(): table_overrides["Tile Hook"] = ov_tile_hook_qty.strip()
        if ov_grounding_clip.strip(): table_overrides["Grounding Clip"] = ov_grounding_clip.strip()
        if ov_grounding_lug.strip(): table_overrides["Grounding Lug"] = ov_grounding_lug.strip()
        if ov_mc4.strip(): table_overrides["MC4 Connector"] = ov_mc4.strip()
        if ov_railing.strip(): table_overrides["R-Railing"] = ov_railing.strip()
        if ov_splice.strip(): table_overrides["Splice"] = ov_splice.strip()
        if ov_lfoot.strip(): table_overrides["L-Foot"] = ov_lfoot.strip()
        
        name_overrides = {}
        if roof_name_choice and roof_name_choice != "Keep Original" and roof_name_choice.strip():
            name_overrides["Roof Attachment"] = roof_name_choice.strip()
            name_overrides["3"] = roof_name_choice.strip()

        # Run QC Comparison
        page1_results, page2_bom_rows, page3_4_results = run_qc_check(
            excel_data, pdf_data,
            manual_jumpers=jumper_input,
            roof_name_override=roof_name_choice,
            table_overrides=table_overrides
        )
        
        # Identify mismatches for red highlight
        mismatched_components = []
        if highlight_mismatches:
            for r in page2_bom_rows:
                if "MISMATCH" in str(r.get("Status", "")):
                    comp = r.get("Component", "")
                    for key in ["End Clamp", "Mid Clamp", "Tile Hook", "U-Rail", "Cliplock", "L-Foot", "Shingle", "Ballast", "Grounding Clip", "Grounding Lug", "MC4 Connector", "R-Railing", "Splice"]:
                        if key.upper() in comp.upper():
                            mismatched_components.append(key)
        
        # Reflect manual overrides in BOM comparison table
        if name_overrides:
            for r in page2_bom_rows:
                comp_name = str(r.get("Component", "")).strip()
                if comp_name.startswith("3."):
                    new_n = roof_name_choice.strip()
                    r["Component"] = f"3. {new_n} (Modified - Red)"
                    r["Notes"] = f"Component name changed to '{new_n}' (Rendered in RED on DWG Page 2)."
                    
        if table_overrides:
            for r in page2_bom_rows:
                comp_name = str(r.get("Component", "")).strip()
                for k, v in table_overrides.items():
                    if k in ["Tile Hook", "Roof Attachment", "3"]:
                        if comp_name.startswith("3."):
                            r["DWG Quantity"] = f"{v} (Modified - Red)"
                            r["Notes"] = f"Quantity updated to {v} via web interface (Rendered in RED on DWG Page 2)."
                    elif k.upper() == "END CLAMP" and comp_name.startswith("1."):
                        r["DWG Quantity"] = f"{v} (Modified - Red)"
                        r["Notes"] = f"Quantity updated to {v} via web interface (Rendered in RED on DWG Page 2)."
                    elif k.upper() == "MID CLAMP" and comp_name.startswith("2."):
                        r["DWG Quantity"] = f"{v} (Modified - Red)"
                        r["Notes"] = f"Quantity updated to {v} via web interface (Rendered in RED on DWG Page 2)."
                    elif k.upper() == "GROUNDING CLIP" and comp_name.startswith("4."):
                        r["DWG Quantity"] = f"{v} (Modified - Red)"
                        r["Notes"] = f"Quantity updated to {v} via web interface (Rendered in RED on DWG Page 2)."
                    elif k.upper() == "GROUNDING LUG" and comp_name.startswith("5."):
                        r["DWG Quantity"] = f"{v} (Modified - Red)"
                        r["Notes"] = f"Quantity updated to {v} via web interface (Rendered in RED on DWG Page 2)."
                    elif k.upper() == "MC4 CONNECTOR" and comp_name.startswith("6."):
                        r["DWG Quantity"] = f"{v} (Modified - Red)"
                        r["Notes"] = f"Quantity updated to {v} via web interface (Rendered in RED on DWG Page 2)."
                    elif k.upper() == "R-RAILING" and comp_name.startswith("7."):
                        r["DWG Quantity"] = f"{v} (Modified - Red)"
                        r["Notes"] = f"Quantity updated to {v} via web interface (Rendered in RED on DWG Page 2)."
                    elif k.upper() == "SPLICE" and comp_name.startswith("8."):
                        r["DWG Quantity"] = f"{v} (Modified - Red)"
                        r["Notes"] = f"Quantity updated to {v} via web interface (Rendered in RED on DWG Page 2)."
                    elif "L-FOOT" in k.upper() and ("L-FOOT" in comp_name.upper() or "L FOOT" in comp_name.upper()):
                        r["DWG Quantity"] = f"{v} (Modified - Red)"
                        r["Notes"] = f"Quantity updated to {v} via web interface (Rendered in RED on DWG Page 2)."
        
        # Clean & auto-format breaker & CT inputs (e.g. '32' -> '32A', '1' -> '1P', '32' -> '32VA')
        clean_mcb_amp = f"{ov_mcb_amp.strip()}A" if (ov_mcb_amp.strip() and not ov_mcb_amp.strip().upper().endswith("A")) else ov_mcb_amp.strip()
        clean_mcb_pole = f"{ov_mcb_pole.strip()}P" if (ov_mcb_pole.strip() and not ov_mcb_pole.strip().upper().endswith("P")) else ov_mcb_pole.strip()
        clean_ct_val = f"{ov_ct_rating.strip()}VA" if (ov_ct_rating.strip() and not ov_ct_rating.strip().upper().endswith("VA")) else ov_ct_rating.strip()

        # Check if manual breaker / CT overrides are actually DIFFERENT from DWG
        # "kalau dh tally tkyh ubah apa2"
        orig_mcb = pdf_data.get("page4", {}).get("inverter_mcb", "")
        orig_pvm = pdf_data.get("page4", {}).get("pv_meter", "")
        
        mcb_is_different = False
        stamp_mcb = None
        stamp_pole = None
        
        if clean_mcb_amp and (clean_mcb_amp.upper() not in orig_mcb.upper()):
            mcb_is_different = True
            stamp_mcb = clean_mcb_amp
        if clean_mcb_pole and (clean_mcb_pole.upper() not in orig_mcb.upper()):
            mcb_is_different = True
            stamp_pole = clean_mcb_pole
            
        ct_is_different = False
        stamp_ct = None
        if clean_ct_val:
            ct_digits = re.sub(r"[^\d]", "", clean_ct_val)
            if ct_digits and (ct_digits not in orig_pvm):
                ct_is_different = True
                stamp_ct = clean_ct_val

        # Reflect in Page 3 & 4 QC table ONLY if actually modified / different
        if mcb_is_different:
            for r in page3_4_results:
                if "MCB" in r.get("Check Item", "").upper() or "BREAKER" in r.get("Check Item", "").upper():
                    custom_mcb_str = f"{clean_mcb_amp or ''} {clean_mcb_pole or ''} 10kA MCB".strip()
                    r["DWG Value"] = f"{custom_mcb_str} (Modified - Red)"
                    r["Status"] = "✅ TALLY"
                    r["Notes"] = f"Breaker modified to {custom_mcb_str} (Rendered in RED on DWG Page 4)."
                    
        if ct_is_different:
            for r in page3_4_results:
                if "CT" in r.get("Check Item", "").upper() or "METER" in r.get("Check Item", "").upper():
                    r["DWG Value"] = f"{clean_ct_val} (Modified - Red)"
                    r["Status"] = "✅ TALLY"
                    r["Notes"] = f"CT rating modified to {clean_ct_val} (Rendered in RED on DWG Page 4)."

        # Stamp cables, table overrides, and breaker modifications (cached for instant performance)
        stamped_dwg_bytes = get_cached_stamped_dwg(
            pdf_bytes,
            dc_cable_input,
            earthing_cable_input,
            table_overrides,
            name_overrides,
            tuple(mismatched_components),
            stamp_mcb,
            stamp_pole,
            stamp_ct
        )
        
        # Pre-extract visual images (cached for maximum performance)
        p1_pv_layout_img = get_cached_pv_layout(pdf_bytes, dpi=200)
        p1_sys_info_img = get_cached_sys_info_img(pdf_bytes, dpi=200)
        p1_full_img = get_cached_rendered_page(pdf_bytes, page_num=0, dpi=140)
        
        # Page 2 visuals: both focused tables zoom and full layout using stamped_dwg_bytes (cached)
        p2_table_zoom_img = get_cached_page2_table_zoom(stamped_dwg_bytes, dpi=220)
        p2_full_img = get_cached_rendered_page(stamped_dwg_bytes, page_num=1, dpi=140)
        
        p3_full_img = get_cached_rendered_page(pdf_bytes, page_num=2, dpi=140)
        p4_full_img = get_cached_rendered_page(stamped_dwg_bytes, page_num=3, dpi=140)
        
    except Exception as e:
        st.error(f"Error processing files: {str(e)}")
        st.stop()

# Overall Status Metrics
all_checks = page1_results + page3_4_results
fail_count = sum(1 for r in all_checks if "FAIL" in r.get("Status", r.get("status", "")))
pass_count = sum(1 for r in all_checks if "PASS" in r.get("Status", r.get("status", "")))

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.markdown(f"""
    <div class="metric-card" style="border-top: 4px solid {'#ef4444' if fail_count > 0 else '#10b981'};">
        <div class="metric-label">Verification Status</div>
        <div class="metric-val" style="color: {'#dc2626' if fail_count > 0 else '#16a34a'};">
            {'❌ MISMATCH' if fail_count > 0 else '✅ ALL TALLY'}
        </div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Panel Quantity</div>
        <div class="metric-val">{pdf_data.get('page1', {}).get('total_panels', 0)} PCS</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">System Capacity</div>
        <div class="metric-val">{pdf_data.get('page1', {}).get('kwp', 0.0):.2f} kWp</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">System Phase</div>
        <div class="metric-val">{pdf_data.get('page3', {}).get('phase', 'Three Phase')}</div>
    </div>
    """, unsafe_allow_html=True)

with col5:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Excel Format</div>
        <div class="metric-val" style="font-size: 16px;">{excel_data.get('format', 'Detected')}</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Main Navigation Tabs
tab_summary, tab_pages, tab_export = st.tabs([
    "📋 Match Verification Matrix (Page 1)",
    "🔍 Technical Details (Pages 1 - 4)",
    "📑 Combine & Download PDF (Combiner)"
])

# -------------------------------------------------------------
# TAB 1: SUMMARY MATRIX (PAGE 1 & CLIENT INFO + VISUAL SNIPPETS)
# -------------------------------------------------------------
with tab_summary:
    st.subheader("📋 Match Verification Matrix: Client Information & Page 1 (PV Layout)")
    st.caption("Cross-verification between Project Form (Excel) and DWG Page 1. For Mounting Structure BOM and SLD details, please view the 'Technical Details (Pages 1 - 4)' tab.")
    
    rows_p1 = []
    for r in page1_results:
        status_str = r["status"]
        if status_str == "PASS":
            status_badge = "✅ TALLY"
        elif status_str == "FAIL":
            status_badge = "❌ MISMATCH"
        elif status_str == "WARNING":
            status_badge = "⚠️ WARNING"
        else:
            status_badge = "ℹ️ INFO"
            
        rows_p1.append({
            "Category": r["category"],
            "Check Item": r["item"],
            "Project Form (Excel)": str(r["excel_val"]),
            "Construction Drawing (DWG)": str(r["dwg_val"]),
            "Status": status_badge,
            "Verification Notes": r["notes"]
        })
        
    df_p1 = pd.DataFrame(rows_p1)
    st.dataframe(df_p1, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 🖼️ Visual Reference: DWG Page 1 Engineering Drawings")
    st.caption("Direct visual excerpts extracted from Page 1 of the uploaded DWG PDF to inspect layout routing and system specifications:")
    
    col_v1, col_v2 = st.columns([1.3, 1.0])
    
    with col_v1:
        st.markdown("**Roof PV Layout (Front View with Strings):**")
        if p1_pv_layout_img:
            st.image(
                p1_pv_layout_img,
                caption="DWG Page 1: Roof PV Layout Front View (String Routing, Jumpers & Fall Angle)",
                use_container_width=True
            )
        else:
            st.info("Roof PV Layout snippet not available.")
            
    with col_v2:
        st.markdown("**System Information Table:**")
        if p1_sys_info_img:
            st.image(
                p1_sys_info_img,
                caption="DWG Page 1: System Information Table",
                use_container_width=True
            )
        else:
            st.info("System Information table snippet not available.")
            
    with st.expander("🔍 View Full DWG Page 1 Drawing (High Resolution)"):
        if p1_full_img:
            st.image(p1_full_img, caption="DWG Page 1: Complete Sheet & Site Key Plan", use_container_width=True)
        else:
            st.info("Full Page 1 preview not available.")

# -------------------------------------------------------------
# TAB 2: PAGE-BY-PAGE DETAILS (PAGES 1 - 4)
# -------------------------------------------------------------
with tab_pages:
    # 1. Page 1 Summary
    st.markdown("### ☀️ Page 1: PV Layout & System Specifications Overview")
    p1 = pdf_data.get("page1", {})
    c_p1_a, c_p1_b = st.columns(2)
    with c_p1_a:
        st.write(f"**Client Name:** `{pdf_data.get('client_name', '-')}`")
        st.write(f"**Installation Address:** `{pdf_data.get('address', '-')}`")
        st.write(f"**PV Module Model:** `{p1.get('panel_model', '-')}` ({p1.get('panel_type', 'Bifacial')})")
        st.write(f"**Total Solar Panels:** `{p1.get('total_panels', 0)} PCS`")
    with c_p1_b:
        st.write(f"**System Capacity:** `{p1.get('kwp', 0.0):.2f} kWp` / `{p1.get('kwac', 0.0):.2f} kWac`")
        st.write(f"**Inverter Model:** `{p1.get('inverter', '-')}`")
        st.write(f"**Battery Model:** `{p1.get('battery', '-')}`")
        st.write(f"**Battery Quantity:** `{p1.get('battery_qty', 0)} Unit` (Excel: `{excel_data.get('battery_qty', 0)} Unit`)")
        st.write(f"**String Configuration:** `{', '.join(p1.get('string_config', [])) or str(p1.get('strings_count', 0)) + ' Strings'}`")

    st.markdown("#### 📐 DWG Page 1 Drawing Preview (Roof PV Layout & System Specifications):")
    if p1_full_img:
        st.image(p1_full_img, caption="DWG Page 1: Roof PV Layout & System Specifications", use_container_width=True)
    else:
        st.info("Page 1 preview not available.")

    st.markdown("---")
    
    # 2. Page 2 Mounting Structure BOM
    st.markdown("### 🔧 Page 2: System Information (Mounting Structure BOM)")
    st.write("This table compares the quantities listed on the DWG drawing against standard Northern Solar SOP engineering formulas (Items 1 to 8):")
    
    df_bom = pd.DataFrame(page2_bom_rows)
    st.dataframe(df_bom, use_container_width=True, hide_index=True)

    # Dedicated Cable Information Table (Gambar 1)
    st.markdown("### 🔌 Cable Information Table (Gambar 1)")
    st.write("Separate cable specification table stamped directly below System Information on DWG Page 2:")
    
    cable_table_rows = [
        {
            "Item No.": "1",
            "Info": "DC Cable",
            "Quantity (m)": str(dc_cable_input or "18"),
            "SOP / Expected Calculation": f"{dc_cable_input or '18'} m (Configured in Web Interface)",
            "Status": "✅ TALLY",
            "Notes": "Stamped into Row 1 of the Cable Information table on DWG Page 2."
        },
        {
            "Item No.": "2",
            "Info": "Ground Cable",
            "Quantity (m)": str(earthing_cable_input or "16"),
            "SOP / Expected Calculation": f"{earthing_cable_input or '16'} m (Configured in Web Interface)",
            "Status": "✅ TALLY",
            "Notes": "Stamped into Row 2 of the Cable Information table on DWG Page 2."
        }
    ]
    df_cables = pd.DataFrame(cable_table_rows)
    st.dataframe(df_cables, use_container_width=True, hide_index=True)

    # Visual DWG Page 2 Preview
    st.markdown("#### 📐 Stamped DWG Page 2 System Information & Cable Information Tables:")
    st.caption("Visual verification showing both tables with authentic AutoCAD vectors, red highlights for changes/mismatches, and separate Cable Information table (Gambar 1):")
    
    col_p2_table, col_p2_dwg = st.columns([1.0, 1.5])
    with col_p2_table:
        st.markdown("**Page 2 Tables (System Information + Cable Information):**")
        if p2_table_zoom_img:
            st.image(p2_table_zoom_img, caption="DWG Page 2: System Information & Cable Information Tables", use_container_width=True)
        else:
            st.info("Table zoom preview not available.")
    with col_p2_dwg:
        st.markdown("**Mounting Structure Drawing Layout:**")
        if p2_full_img:
            st.image(p2_full_img, caption="DWG Page 2: Full Mounting Structure Layout & Stamped Tables", use_container_width=True)
        else:
            st.info("Page 2 preview not available.")

    st.markdown("---")
    
    # 3. Pages 3 & 4 SLD & PVMSB
    st.markdown("### ⚡ Pages 3 & 4: SLD, Inverter Current & Circuit Protection (PVMSB)")
    st.write("Cross-verification of inverter output current, MCB circuit breaker sizing, and CT metering ratings based on inverter model:")
    
    df_sld = pd.DataFrame(page3_4_results)
    st.dataframe(df_sld, use_container_width=True, hide_index=True)
    
    with st.expander("📖 Northern Solar Inverter & Protection Specification Reference"):
        st.markdown("""
        | Inverter Model | Output Current | MCB Breaker Rating | PV Meter CT Rating | Electrical Phase |
        | :--- | :---: | :---: | :---: | :---: |
        | **HIS-5L-G3** | **25.0 A** | **32A 1P 10kA MCB** | **32VA CL1 3CTS** | Single Phase |
        | **HIT-8L-G3** | **13.3 A** | **32A 3P 10kA MCB** | **32VA CL1 3CTS** | Three Phase |
        | **HIT-10L-G3** | **16.7 A** | **32A 3P 10kA MCB** | **32VA CL1 3CTS** | Three Phase |
        | **HIT-12L-G3** | **20.0 A** | **32A 3P 10kA MCB** | **32VA CL1 3CTS** | Three Phase |
        | **HIT-15L-G3** | **25.0 A** | **32A 3P 10kA MCB** | **32VA CL1 3CTS** | Three Phase |
        | **HIT-20L-G3** | **33.3 A** | **40A 3P 10kA MCB** | **40VA CL1 3CTS** | Three Phase |
        """)

    # Visual Previews for Pages 3 & 4
    st.markdown("#### 📐 DWG Pages 3 & 4 Drawing Previews (SLD & PVMSB Schematic):")
    col_p3, col_p4 = st.columns(2)
    with col_p3:
        if p3_full_img:
            st.image(p3_full_img, caption="DWG Page 3: Single Line Diagram (SLD)", use_container_width=True)
        else:
            st.info("Page 3 preview not available.")
    with col_p4:
        if p4_full_img:
            st.image(p4_full_img, caption="DWG Page 4: PVMSB Schematic & Layout", use_container_width=True)
        else:
            st.info("Page 4 preview not available.")

# -------------------------------------------------------------
# TAB 3: COMBINE & EXPORT (EXCEL TO PDF & COMBINER)
# -------------------------------------------------------------
with tab_export:
    st.subheader("📑 Project Form to PDF Conversion & DWG PDF Combiner")
    st.write("Exports the original Project Form into a clean, full-page authentic PDF (matching exact Excel layout), stamps separate Cable Information table directly below System Information on DWG Page 2, highlights overrides and mismatches in RED, and merges both documents into a single delivery PDF named strictly after the client.")
    
    col_c1, col_c2 = st.columns(2)
    
    with col_c1:
        if st.button("🚀 Generate Project Form PDF & Combine with DWG", type="primary", use_container_width=True):
            with st.spinner("Converting Excel to authentic PDF, stamping Cable Information & red highlights into DWG Page 2, and combining documents..."):
                try:
                    # 1. Retrieve genuine original Excel PDF (instant 0.00s from cache)
                    excel_pdf_bytes = get_cached_excel_pdf(excel_bytes, fmt_name=excel_data.get("format"))
                    
                    # 2. Combine Excel PDF + DWG PDF (stamped_dwg_bytes already includes Page 2 & Page 4 red stamps -> instant 0.03s merge!)
                    merged_bytes = combine_pdfs(
                        excel_pdf_bytes,
                        stamped_dwg_bytes,
                        already_stamped=True
                    )
                    
                    # File name MUST BE client name only
                    clean_client = excel_data.get('client_name', '').strip()
                    if not clean_client:
                        clean_client = pdf_data.get('client_name', '').strip()
                    if not clean_client:
                        clean_client = "Solar_Project"
                        
                    clean_client = re.sub(r'[\/\\:\*\?"<>\|]+', '', clean_client).strip()
                    
                    st.session_state["excel_pdf"] = excel_pdf_bytes
                    st.session_state["merged_pdf"] = merged_bytes
                    st.session_state["excel_filename"] = f"Project_Form_{clean_client}.pdf"
                    st.session_state["merged_filename"] = f"{clean_client}.pdf"
                    
                    st.success(f"✅ Combined PDF successfully generated as '{clean_client}.pdf'!")
                except Exception as ex:
                    st.error(f"Error during PDF generation: {str(ex)}")

        # Download buttons if generated
        if "merged_pdf" in st.session_state:
            st.markdown("#### 📥 Download Generated PDF Files:")
            st.download_button(
                label=f"📄 Download Combined Delivery PDF ({st.session_state.get('merged_filename')})",
                data=st.session_state["merged_pdf"],
                file_name=st.session_state.get("merged_filename"),
                mime="application/pdf",
                key="btn_download_merged",
                use_container_width=True
            )
            st.download_button(
                label=f"📑 Download Project Form PDF Only (Full Page)",
                data=st.session_state["excel_pdf"],
                file_name=st.session_state.get("excel_filename"),
                mime="application/pdf",
                key="btn_download_excel",
                use_container_width=True
            )
            
    with col_c2:
        st.markdown("#### 📊 Export Full QC Verification Report (CSV):")
        # Combine all report rows
        all_rows = rows_p1 + [
            {
                "Category": "Page 2: Mounting Structure",
                "Check Item": r["Component"],
                "Project Form (Excel)": r["SOP / Expected Calculation"],
                "Construction Drawing (DWG)": str(r["DWG Quantity"]),
                "Status": r["Status"],
                "Verification Notes": r["Notes"]
            }
            for r in page2_bom_rows
        ] + [
            {
                "Category": "Page 2: Cable Information",
                "Check Item": f"{c['Item No.']}. {c['Info']}",
                "Project Form (Excel)": c["SOP / Expected Calculation"],
                "Construction Drawing (DWG)": str(c.get("Quantity (m)", c.get("Quantity", ""))),
                "Status": c["Status"],
                "Verification Notes": c["Notes"]
            }
            for c in cable_table_rows
        ] + [
            {
                "Category": r["Category"],
                "Check Item": r["Check Item"],
                "Project Form (Excel)": r["Expected Value (SOP)"],
                "Construction Drawing (DWG)": str(r["DWG Value"]),
                "Status": r["Status"],
                "Verification Notes": r["Notes"]
            }
            for r in page3_4_results
        ]
        
        csv_buffer = io.StringIO()
        pd.DataFrame(all_rows).to_csv(csv_buffer, index=False)
        client_tag = excel_data.get('client_name') or pdf_data.get('client_name') or 'Solar_Project'
        client_tag = re.sub(r'[\/\\:\*\?"<>\|]+', '', client_tag).strip()
        st.download_button(
            label="📊 Download Full QC Verification Report (CSV)",
            data=csv_buffer.getvalue(),
            file_name=f"{client_tag}_QC_Report.csv",
            mime="text/csv",
            use_container_width=True
        )
