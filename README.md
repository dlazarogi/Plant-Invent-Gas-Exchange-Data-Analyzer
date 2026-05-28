<p align="center">
  <img src="www/images/logo_app.jpg" alt="logo_app">
</p>



# Plant-Invent-Gas-Exchange-Data-Analyzer

Python-based application to analyze Gas Exchange Jyrkki systems from PlantInvent. Compatible with Jyrkki and Jyrkki MAX (partially) output formats, with comprehensive physiological calculations based on established plant physiology formulas.

# User Workflow

This document describes a **recommended step-by-step workflow** for using the Gas Exchange Data Processing Python application.  
Following this workflow ensures **physiological correctness, reproducibility, and clear interpretation** of results.

---

## 1. Prepare your input file

Before opening the app, ensure that your data file meets the following requirements:

- Tab-delimited **TXT** file
- Decimal separator: **comma (, )**
- Time format: **HH:MM:SS**
- Consistent channel labels (e.g., "Channel 1", "Channel 2")
- Correct physiological units (e.g. mmol, µmol, cm/s)

> ⚠️ Any formatting or unit error at this stage will propagate through *all* downstream calculations.

### Batch Processing File Requirements

For batch processing across multiple replicates:
- Files should follow naming pattern: `*_Rep1.txt`, `*_Rep2.txt`, etc.
- All files must have identical column structure
- Optional: Excel mapping file for channel renaming and parameter assignment

#### Excel Mapping File Format (5 columns)

| Replicate | Channel number | New channel name | Leaf area | Absorbed radiation |
|----------|----------|----------|----------|----------|
| Rep1 | Channel 1 | no_cond1-4_nacl | 0.0013478 | 0.000591 |
| Rep1 | Channel 2 | no_cond1-6 | 0.0022008 | 0.000591 |

> 💡 Leaf area and radiation columns are optional but recommended for batch processing, as it is assumed the individual replicates have been checked individually.

---

## 2. Upload and process the file

1. Click **Browse** in the sidebar
2. Select your TXT file
3. Click **Process File** and wait for confirmation

**Processing steps performed automatically:**
- Raw data parsing (28 columns from Jyrkki output)
- Time normalization and continuous time calculation
- All physiological calculations (see formulas below)

At this stage:
- Raw data are loaded into the Raw Data tab
- Calculated variables appear in Interactive Calculated Data tab
- All Interactive Plots tabs are populated
- Processing controls become visible

If you need a clean restart, **reload a file** rather than resetting individual controls.

---

## 3. Inspect raw data

In the **Raw Data** tab:

- Check for missing or unexpected values (NaN)
- Verify channel identity and ordering (1-8 typically)
- Confirm correct time progression within channels
- Verify units match expectations

**Key columns to inspect:**
- `CO2_M`, `CO2_R`: Measured and reference CO₂
- `H2O_M`, `H2O_R`: Measured and reference water vapor
- `absorbed_radiation`: Light intensity (cal/cm²s)
- `leaf_area`: Leaf area in cm²

This step is critical before interpreting calculated outputs.

---

## 4. Rename and select channels

### Channel Renaming

In the **Interactive Calculated Data** tab:

1. Locate the **Channel Parameters** table at the bottom
2. Enter new names in the **Rename** column
3. Click **Rename Channels** button

**Important rules:**
- Rename channels **early** before applying darkness or radiation settings
- Channel names persist across plots, tables, and parameter tables
- Names with parentheses (e.g., `no_cond1-4_nacl (2)`) are automatically cleaned to base names (`no_cond1-4_nacl`) for grouping

### Channel Selection

Use the channel listbox in the sidebar to:
- Select which channels appear in all plots and tables
- Focus on biologically relevant channels only
- Hide noisy or failed channels

> ✔️ Selection applies globally to Interactive Calculated Data table, Interactive Plots, Darkness tab, and Red+Blue Light tab.

---

## 5. Adjust time handling

Use the sidebar controls **before clicking Apply Changes**:

- **Remove Initial Points:** Remove equilibrium artifacts at start
- **Remove Final Points:** Remove end-of-run instability
- **Shift Start Time:** Positive values shift time right (start later), negative values shift left (start earlier)

After entering values, click **Apply Changes** to:
- Recalculate all metrics with new time parameters
- Update all plots and tables
- Adjust darkness and radiation settings to match new time scale

> ⚠️ The slider info label shows maximum allowable removals based on the channel with fewest points.

### Time Range Information

The **Time Range Information** frame displays:
- Overall time range (min)
- Duration and total points
- Time interval between measurements (integer, forced across replicates)
- Points per channel

---

## 6. Apply physical and experimental parameters

Update parameters **one category at a time**, using the corresponding buttons in the **Channel Parameters** table:

| Parameter | Button | Recalculation |
|-----------|--------|---------------|
| Channel names | Rename Channels | Affects display only |
| Leaf area | Update Leaf Area | Affects all area-normalized rates |
| Air flow rate | Update Air Flow Rate | Affects all exchange rates |
| Absorbed radiation | Update Absorbed Radiation | Affects leaf temperature, conductances |
| Boundary layer cond | Update Boundary Layer Cond | Affects conductance calculations |
| Cutic conductance | Update Cutic Conductance | Affects stomatal conductance |

**After each update:**
- Click the corresponding button
- Inspect CO₂ exchange, transpiration, and stomatal conductance plots
- Confirm that changes behave as expected

> 💡 Modified values persist across time adjustments. The Modified Data indicator shows when non-default parameters are active.

---

## 7. Inspect physiological plots

Follow the physiological dependency order in **Interactive Plots** tab:

### Recommended inspection order:

1. **Environment** (air_temp, absorbed_radiation, relative_air_humidity, VPD)
   - Validate environmental conditions
   - Check for unexpected fluctuations

2. **Base Fluxes** (CO₂_exchange_rate, Transpiration_H2O_evol_rate, leaf_temp_c)
   - Primary physiological responses
   - Compare with expected patterns

3. **Humidity/Vapor** (leaf_air_hum_grad, satur_hum_at_leaf_temp_c)
   - Driving forces for gas exchange

4. **Conductances** (overall_conductance, corr_stomatal_conductance, Stomatal_conductance_corrected)
   - Stomatal regulation patterns
   - Temperature-corrected vs uncorrected

5. **Ozone** (overall_O3_uptake_rate_c, corr_stomatal_O3_uptake_rate, corr_cumul_O3_dose)
   - Ozone uptake and cumulative dose

6. **CO₂** (intercell_CO2_conc_in_gas_phase, stom_resist_to_CO2)
   - Internal CO₂ concentrations
   - Resistance components

7. **Mesophyll** (mesophyll_conductance_for_CO2, CO2_comp_point)
   - Biochemical limitations
   - CO₂ compensation point

8. **Efficiency** (WUEi, WUE)
   - Water use efficiency
   - Compare intrinsic vs operational WUE

> ⚠️ Do not interpret downstream variables (e.g., mesophyll conductance) before validating upstream conditions (CO₂ exchange, leaf temperature).

### Plot Features

- **Color palettes:** Select from 13 colorblind-friendly palettes
- **Channel legend:** Automatically positioned to the right
- **Time zero line:** Dashed vertical line at x=0 if within range
- **Download:** Each plot can be saved as PNG (300 DPI) or PDF

---

## 8. Darkness analysis (Light-Dark tab)

### Auto-Detection (Recommended First Step)

1. Go to **Light-Dark** tab
2. Click **Apply Darkness** without entering any values
3. The app automatically detects periods where `absorbed_radiation = 0`
4. Auto-detected periods are shown as black rectangles in plots

### Manual Override (Add Additional Darkness Periods)

If additional darkness periods are needed:

1. In the **Light-Dark** tab, locate the input box for each channel
2. Enter comma-separated integer values of `time_repeated` (e.g., `5, 10, 15, 20`)
3. Click **Apply Darkness**

**Important behavior:**
- Manual entries are **added to** auto-detected periods (not replaced)
- Input boxes show only manual entries (auto-detected remains hidden but active)
- All dependent variables are recalculated after each apply

### Darkness Data Table

- View all modified data in the **Darkness Data Table** tab
- Export with **RE-IMPORT** (28 columns) or **ALL COLUMNS** options
- Data sorted by continuous time (handles multi-day experiments)

> ⚠️ Darkness is a **physiological intervention** (absorbed_radiation = 0), not a visual filter. Leaf temperature and conductances are affected.

---

## 9. Radiation intervals – Red+Blue Light tab

To redefine absorbed radiation for specific time intervals:

1. Go to **Red+Blue Light** tab
2. For each channel, enter intervals in format: `start-end:value`
3. Multiple intervals separated by semicolons: `0-10:0.0006; 10-20:0.0012; 20-30:0.0003`
4. Click **Apply Radiation Intervals**

**Example:** `0-5:0.000591; 5-15:0.000886; 15-30:0.000197`

**Important rules:**
- Intervals should be non-overlapping per channel
- Endpoints are inclusive (tiny epsilon added automatically)
- Values overwrite original `absorbed_radiation` for the specified time range
- All dependent metrics recalculated

### Radiation Data Table

- View modified data in the **Radiation Data Table** tab
- Export with same RE-IMPORT / ALL COLUMNS options
- Sorted by continuous time

---

## 10. Batch processing (multiple replicates)

### Setup

1. Click **Add Files** to select replicate files (select multiple)
2. (Optional) Select Excel mapping file for channel renaming
3. Select output directory
4. Click **Process Batch**

### Batch Output

The following files are generated in the output directory:

| File | Description |
|------|-------------|
| `batch_statistics.xlsx` | Mean ± SEM per channel and time point |
| `batch_report.txt` | Processing summary |
| `averaged_plots/` | Mean ± SEM plots for all metrics |
| `WUEi_Distribution.png` | Violin + boxplot across replicates |
| `combined_data.csv` | All processed data combined |

### Batch Results Tabs

After batch processing completes, the **Batch Results** tab contains:

1. **Statistics Table:** Mean ± SEM per channel and time point
2. **Metric Plots:** Interactive plots with channel selection
   - Same categories as Interactive Plots
   - Error bars show SEM across replicates
   - Channel selector at the top
3. **Darkness Detection:** Auto-detected darkness periods with summary
4. **WUEi Distribution:** Violin + boxplot of WUEi values across all time points

### Channel Selection in Batch Results

- Use the channel selector at the top of Metric Plots tab
- Selected channels filter ALL batch plots (Metric Plots, Darkness Detection, WUEi Distribution)
- Changes apply immediately

---

## 11. Download and document results

### Single File Export

Before closing the app, use the **Download Data** button in Interactive Calculated Data tab:

**Export format options:**
- **RE-IMPORT:** Only original 28 columns, compatible with raw data import
- **ALL COLUMNS:** All calculated metrics plus original columns

### Darkness Data Export

Click **Download Darkness Data** in the Light-Dark tab:

- Same RE-IMPORT / ALL COLUMNS options
- Data includes darkness modifications (absorbed_radiation = 0)
- Handles multi-day experiments with correct sorting

### Radiation Data Export

Click **Download Radiation Data** in the Red+Blue Light tab:

- Same export format options
- Data includes radiation interval modifications
- Intervals applied before export

### Plot Downloads

Each plot has its own **Download** button:
- PNG format at 300 DPI
- PDF format for vector graphics

---

## Recommended workflow order

**Single file:**
1. Upload → Process File
2. Raw Data inspection
3. Channel renaming and selection
4. Time adjustments → Apply Changes
5. Parameter updates (one at a time)
6. Interactive Plots inspection (Environment → Efficiency)
7. Optional: Darkness analysis
8. Optional: Radiation intervals
9. Download data and plots

**Batch processing:**
1. Add multiple files (matching pattern)
2. (Optional) Load Excel mapping file
3. Set output directory
4. Process Batch
5. Review Batch Results tabs
6. Export statistics and plots

---


## 📜 License & Terms of Use

### Dual Licensing Model

This project is available under **two different licenses** depending on your use case:

| License | Who Can Use | Commercial Use | Price |
| :--- | :--- | :--- | :--- |
| **Research License** | Universities, researchers, students, non-profits | ❌ Not allowed | Free |
| **Commercial License** | Companies, startups, commercial entities | ✅ Allowed | Paid |

### 🔬 Research & Academic Use (Free)

If you are:
- A student working on a thesis or dissertation
- A researcher at a non-profit academic institution
- Using the software for non-commercial scientific research
- Teaching or demonstrating gas exchange concepts

→ You may use the **Research License** (CC BY-NC-SA 4.0)

### 💼 Commercial Use (Paid)

You need a Commercial License if you want to:
- Use the software in a paid product or service
- Use the software for commercial consulting or contract research
- Integrate the software into a commercial application
- Distribute the software as part of a paid offering

→ Contact: **david.lazaro.gimeno@gmail.com**

### ❓ Questions?

- **What constitutes commercial research?** Research funded by a company or conducted for a company's benefit requires a Commercial License.
- **Can I share the software with colleagues?** Yes, under the Research License for non-commercial collaboration.
- **Can I modify the code?** Yes under both licenses, but modifications must remain under the same license for Research users.

*Gas Exchange Data Processing App – User Workflow Documentation*
