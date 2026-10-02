# Northern Solar - Pre-Installation QC & DWG PDF Combiner

Automated quality control and PDF document combiner application designed for the **Northern Solar Project Team** before site installation commences.

---

### 🌟 Key Capabilities:

1. **Dual Excel Format Support:**
   * **New Template:** Key-value structured fields (`MODULE PANEL QUANTITY`, `INVERTER MODEL`, `BATTERY MODEL`, `TNB METER`, `ROOF TYPE`).
   * **Old Template:** Automatically parses system capacity from `Remark` text (e.g. `40 PANELS X 635W`) and checklist tables.
   * **Formula Fallback:** Automatically calculates panel quantity if only kWp is supplied:
     $$\text{Panel Quantity} = \frac{\text{kWp}}{\text{Module kW}}$$

2. **Automated Cross-Verification Against Construction Drawings (DWG in PDF):**
   * **Title Block:** Client Name & Site Address (strictly isolates client from contractor/consultant branding).
   * **Page 1 (PV Layout & System Info):** Total Panels, System Capacity (kWp & kWac), Module Model, Inverter Model, Battery Model & Quantity, and String Configuration.
   * **Page 2 (Mounting Structure BOM):**
     * Grounding Clip ($1:1$ panel ratio).
     * MC4 Connector ($\text{Strings} + \text{Jumpers}$).
     * End Clamp: $(4 \times G) + 2$ buffer.
     * Mid Clamp: $2 \times (P - G) + 2$ buffer.
     * Grounding Lug: 1 per array group ($G$).
     * R-Railing / Structure: GstarCAD `AMSTABLE` verified BOM.
   * **Pages 3 & 4 (SLD & PVMSB):**
     * Electrical Phase (Single Phase vs Three Phase).
     * Inverter Rated Current (A) lookup matrix.
     * Inverter Breaker MCB rating (`32A 1P`, `32A 3P`, `40A 3P`).
     * PV Meter CT Rating (`32VA` / `40VA`).

3. **Dynamic Visual References & Snippets:**
   * **Tab 1:** Embedded high-resolution visual snippets of the **Roof PV Layout (Front View)** and **System Information Table**.
   * **Tab 2:** Visual drawing previews for **Page 2** (Mounting Structure), **Page 3** (SLD), and **Page 4** (PVMSB).

4. **Cable Stamping on DWG Page 2:**
   * Web inputs for **DC Cable** (Row 9) and **Earthing Cable** (Row 10).
   * Stamped directly below Row 8 (`Splice`) in the Page 2 table with exact cell borders and font alignment.

5. **Authentic Excel COM Engine & Delivery PDF Combiner:**
   * Uses native Microsoft Excel COM rendering to produce the exact 2-page print layout matching company standards.
   * Merges the 2-page Project Form PDF with the DWG PDF (with stamped Page 2).
   * Generates download named strictly as **`<Client_Name>.pdf`** (e.g. `Liang Yee Mun.pdf`).

---

### 🚀 How to Run the Application:

1. Open folder: `C:\Users\Husna\Desktop\Software\Solar_QC_Combiner_App`.
2. Double-click **`Run_App.bat`**.
3. The app will launch in your browser (`http://localhost:8501`).
4. Upload your Project Form (Excel) and DWG Drawings (PDF) to view verification results and generate the combined delivery PDF!
