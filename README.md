<p align="center">
  <img src="www/images/logo_app.jpg" alt="logo_app">
</p>

# Plant-Invent-Gas-Exchange-Data-Analyzer

Python-based application to analyze Gas Exchange Jyrkki systems from PlantInvent. Compatible with Jyrkki and Jyrkki MAX (partially) output formats, with comprehensive physiological calculations based on established plant physiology formulas.

The application supports **single-file analysis, multi-cuvette processing, batch analysis across replicates, interactive physiological visualization, advanced statistical analysis, ANCOVA, linear mixed models, repeated-measures analysis, light-dark experiments, customized radiation intervals, and publication-quality data export**.

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
- All files should have identical column structure
- Optional: Excel mapping file for channel renaming and parameter assignment

#### Excel Mapping File Format (5 columns)

| Replicate | Channel number | New channel name | Leaf area | Absorbed radiation |
|----------|----------|----------|----------|----------|
| Rep1 | Channel 1 | no_cond1-4_nacl | 0.0013478 | 0.000591 |
| Rep1 | Channel 2 | no_cond1-6 | 0.0022008 | 0.000591 |

> 💡 Leaf area and radiation columns are optional but recommended for batch processing, as it is assumed the individual replicates have been checked individually.

The mapping system also supports channel names containing replicate-specific suffixes such as `no_cond1-4_nacl (2)`. These are automatically converted to a common **base channel name** for batch grouping and statistical analysis.

---

## 2. Upload and process the file

1. Click **Browse** in the sidebar
2. Select your TXT file
3. Click **Process File** and wait for confirmation

**Processing steps performed automatically:**

- Raw data parsing (28 columns from Jyrkki output)
- Time normalization and continuous time calculation
- Integer time-interval detection and normalization
- Physiological calculations
- Conductance and CO₂ corrections
- Ozone calculations
- Photosynthesis and respiration-derived metrics
- Water-use-efficiency calculations

At this stage:

- Raw data are loaded into the Raw Data tab
- Calculated variables appear in Interactive Calculated Data tab
- All Interactive Plots tabs are populated
- Processing controls become visible

If you need a clean restart, **reload a file** rather than resetting individual controls.

### Time normalization across replicates

The application determines the measurement interval independently for each channel and rounds the detected interval to the nearest integer minute. Integer time points are then generated consistently across measurements and replicates. This improves alignment of batch statistics and averaged plots. The original continuous clock time is retained in `time_full_sec` and `time_full_minutes`.

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
- `air_flow_rate`: Chamber air flow
- `O3_M`, `O3_R`: Measured and reference ozone

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
- Base channel names are also used when calculating batch statistics

### Channel Selection

Use the channel listbox in the sidebar to:

- Select which channels appear in all plots and tables
- Focus on biologically relevant channels only
- Hide noisy or failed channels

> ✔️ Selection applies globally to Interactive Calculated Data table, Interactive Plots, Darkness tab, Red+Blue Light tab, and batch statistical analyses.

For batch analysis, selected channels are explicitly used to filter the data before ANOVA and ANCOVA calculations.

---

## 5. Adjust time handling

Use the sidebar controls **before clicking Apply Changes**:

- **Remove Initial Points:** Remove equilibrium artifacts at start
- **Remove Final Points:** Remove end-of-run instability
- **Shift Start Time:** Positive values shift time right (start later), negative values shift time left (start earlier)

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

The application retains both normalized analysis time and the original continuous clock time, which is particularly important for multi-day experiments.

---

## 6. Apply physical and experimental parameters

Update parameters **one category at a time**, using the corresponding buttons in the **Channel Parameters** table:

| Parameter | Button | Recalculation |
|-----------|--------|---------------|
| Channel names | Rename Channels | Affects display and grouping |
| Leaf area | Update Leaf Area | Affects all area-normalized rates |
| Air flow rate | Update Air Flow Rate | Affects all exchange rates |
| Absorbed radiation | Update Absorbed Radiation | Affects leaf temperature, conductances and downstream metrics |
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

3. **Photosynthesis & Respiration** (CO₂_exchange_rate, P_gross, R_dark_estimate)
   - Net photosynthesis
   - Estimated gross photosynthesis
   - Estimated dark respiration

4. **Humidity/Vapor** (leaf_air_hum_grad, satur_hum_at_leaf_temp_c)
   - Driving forces for gas exchange

5. **Conductances** (overall_conductance, corr_stomatal_conductance, Stomatal_conductance_corrected)
   - Stomatal regulation patterns
   - Temperature-corrected vs uncorrected conductance
   - VPD sensitivity

6. **Ozone** (overall_O3_uptake_rate_c, corr_stomatal_O3_uptake_rate, corr_cumul_O3_dose)
   - Ozone uptake and cumulative dose
   - Ozone sensitivity

7. **CO₂ Dynamics** (intercell_CO2_conc_in_gas_phase, Ca_CI_gradient, stomatal_limitation)
   - Internal CO₂ concentrations
   - CO₂ gradients
   - Estimated stomatal limitation

8. **Mesophyll & Biochemistry** (mesophyll_conductance_for_CO2, CO2_comp_point, V_cmax_approx, electron_transport_rate_approx)
   - Mesophyll conductance
   - CO₂ compensation point
   - Approximate biochemical capacity

9. **Efficiency** (WUEi, WUE, WUE_gross, WUEi_gross, WUE_vcmax)
   - Water-use efficiency
   - Compare intrinsic, instantaneous, gross and V_cmax-based WUE

> ⚠️ Do not interpret downstream variables (e.g., mesophyll conductance or V_cmax approximation) before validating upstream conditions (CO₂ exchange, leaf temperature, humidity and conductances).

### Plot Features

- **Color palettes:** Select from multiple colorblind-friendly palettes
- **Channel legend:** Automatically positioned to the right
- **Time zero line:** Dashed vertical line at x=0 if within range
- **Channel filtering:** Select/deselect channels
- **Download:** Each plot can be saved as PNG (300 DPI) or PDF

The plotting system organizes metrics into physiological categories including Environment, Base Fluxes, Photosynthesis, Humidity/Vapor, Conductances, Ozone, CO₂, Mesophyll & Biochemistry, and Efficiency.

---

## 8. Calculated physiological metrics

The application calculates a comprehensive suite of physiological variables.

### Photosynthesis & Respiration

- **Net Photosynthesis (A)** — `CO2_exchange_rate`
- **Gross Photosynthesis** — `P_gross`
- **Dark Respiration Estimate** — `R_dark_estimate`
- **Carboxylation Efficiency** — `carboxylation_efficiency`
- **Stomatal Limitation** — `stomatal_limitation`
- **CO₂ Gradient** — `Ca_CI_gradient`

The current implementation estimates:

- `R_dark_estimate` from mesophyll conductance and CO₂ compensation point
- `P_gross` as net photosynthesis plus estimated dark respiration
- stomatal limitation from the relationship between intercellular and corrected ambient CO₂
- carboxylation efficiency as photosynthesis relative to intercellular CO₂

### Conductances

- Stomatal Conductance
- Corrected Stomatal Conductance
- Boundary Layer Conductance
- Leaf Conductance
- Overall Conductance
- Stomatal Sensitivity to VPD
- Mass-flow corrected conductances

### Ozone Metrics

- O₃ Uptake Rate
- Corrected Conductance for O₃
- Corrected Stomatal O₃ Uptake
- Cumulative O₃ Dose
- Ozone Sensitivity Index (OSI)

### Mesophyll & Biochemistry

- Mesophyll Conductance (`g_m`)
- CO₂ Compensation Point (`Γ`)
- Approximate V_cmax
- Approximate Electron Transport Rate (`J`)
- Uncorrected and corrected mesophyll-related metrics

> ⚠️ Derived biochemical variables such as `V_cmax_approx` and `electron_transport_rate_approx` are simplified approximations and should be interpreted accordingly.

### Water Use Efficiency

- **Intrinsic WUE (WUEi)** — A / g_s
- **Instantaneous WUE (WUE)** — A / E
- **Gross WUE (WUE_gross)** — P_gross / E
- **Gross Intrinsic WUE (WUEi_gross)** — P_gross / g_s
- **V_cmax-based WUE (WUE_vcmax)** — V_cmax approximation / E

### Environmental Metrics

- Vapor Pressure Deficit (VPD)
- Relative Air Humidity
- Estimated Leaf Temperature
- Saturation Humidity
- Leaf-Air Humidity Gradient
- Vapor Pressure Components

---

## 9. Darkness analysis (Light-Dark tab)

### Auto-Detection (Recommended First Step)

1. Go to **Light-Dark** tab
2. Click **Apply Darkness** without entering any values
3. The app automatically detects periods where `absorbed_radiation = 0`
4. Auto-detected periods are shown as black rectangles in plots

The batch detection system processes each channel independently and identifies continuous blocks of zero or near-zero radiation.

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

> ⚠️ Darkness is a **physiological intervention** (`absorbed_radiation = 0`), not a visual filter. Leaf temperature and conductances are affected.

---

## 10. Radiation intervals – Red+Blue Light tab

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
- All dependent metrics are recalculated

### Radiation Data Table

- View modified data in the **Radiation Data Table** tab
- Export with same RE-IMPORT / ALL COLUMNS options
- Sorted by continuous time

This functionality is particularly useful for experiments involving controlled Red+Blue light transitions and custom radiation treatments.

---

## 11. Batch processing (multiple replicates)

### Setup

1. Click **Add Files** to select replicate files (select multiple)
2. (Optional) Select Excel mapping file for channel renaming
3. Select output directory
4. Click **Process Batch**

The batch processor identifies replicate names from filenames and applies the corresponding channel mapping, leaf area and radiation information when available.

### Batch Output

The following files are generated in the output directory:

| File | Description |
|------|-------------|
| `batch_statistics.xlsx` | Mean, SD, SEM and n per channel and time point |
| `batch_report.txt` | Processing and statistical summary |
| `averaged_plots/` | Mean ± SEM plots for metrics |
| `WUEi_Distribution.png` | Violin + boxplot across replicates |
| `combined_data.csv` | All processed data combined |

Batch statistics are calculated by channel and normalized `time_repeated`, with mean, standard deviation, sample size and SEM stored for each metric.

### Batch Results Tabs

After batch processing completes, the **Batch Results** tab contains:

1. **Statistics Table:** Mean ± SEM per channel and time point
2. **Metric Plots:** Interactive plots with channel selection
   - Same categories as Interactive Plots
   - Error bars show SEM across replicates
   - Channel selector at the top
3. **Darkness Detection:** Auto-detected darkness periods with summary
4. **WUEi Distribution:** Violin + boxplot of WUEi values across all time points
5. **ANOVA + post hoc:** Parametric and non-parametric group comparisons
6. **LLM Analysis:** Linear mixed-model comparison
7. **ANCOVA Analysis:** Covariate-adjusted group comparisons
8. **ANOVA Diagnostic Plots:** Statistical diagnostic visualizations

### Channel Selection in Batch Results

- Use the channel selector at the top of Metric Plots tab
- Selected channels filter ALL batch plots and statistical analyses
- Changes apply immediately

---

## 12. Advanced statistical analysis

The application includes an integrated statistical analysis framework designed for batch experiments with biological replicates and repeated measurements.

### Standard ANOVA

Available analyses include:

- **One-Way ANOVA**
- **Two-Way ANOVA (Group × Phase)**
- **Tukey HSD** post-hoc comparisons
- **Kruskal-Wallis** non-parametric test
- **Dunn's** post-hoc comparisons
- **Welch's ANOVA** for unequal variances
- **Games-Howell** post-hoc comparisons
- Shapiro-Wilk normality testing
- Levene's test for homogeneity of variance
- Eta-squared and partial eta-squared effect sizes
- Compact Letter Display for pairwise differences

The application automatically evaluates assumptions and provides recommendations about whether standard or robust analyses are more appropriate.

### One-Way ANOVA

Use when comparing a single response metric among independent channel/group categories.

The analysis provides:

- Group descriptive statistics
- Sample sizes
- Normality assessment
- Homogeneity assessment
- ANOVA F-test
- Effect size
- Tukey HSD where appropriate

### Two-Way ANOVA

Two-way analysis evaluates:

- Group effect
- Phase effect
- Group × Phase interaction

This is particularly useful for experiments where biological groups are exposed to different phases, such as light and dark.

### Non-parametric alternatives

When ANOVA assumptions are not adequately satisfied, the application can evaluate:

**Kruskal-Wallis → Dunn's post-hoc**

or

**Welch's ANOVA → Games-Howell post-hoc**

Games-Howell is specifically used for comparisons where unequal variances are a concern. The application also applies multiple-testing correction and validates compact-letter assignments.

---

## 13. Repeated Measures ANOVA

Repeated Measures ANOVA is available for experiments where the same biological units are measured across experimental phases.

### Recommended use

Select:

**Repeated Measures ANOVA (Recommended for Light-Dark)**

The analysis evaluates:

- **Phase effect** — within-subject effect
- **Group effect** — between-subject effect
- **Group × Phase interaction**
- Sphericity
- Greenhouse-Geisser correction

When there are insufficient replicates for repeated-measures analysis, the application can fall back to a Two-Way ANOVA approach.

### Light-Dark interpretation

For light-dark experiments, the primary question is often whether groups respond differently to the transition between experimental phases.

The **Group × Phase interaction** is therefore particularly important when determining whether treatments have different physiological responses to light or darkness.

---

## 14. Linear Mixed Models (LMM)

The **LLM Analysis** tab provides Linear Mixed Model analysis for batch experiments in which repeated measurements or replicate-specific variability should be accounted for.

Available model structures include:

- Random intercept model
- Random intercept + random slope model
- Reduced fixed-effect model
- OLS comparison model

### Model comparison

The application provides:

- Likelihood Ratio Tests
- AIC comparison
- BIC comparison
- Random-effect evaluation
- Model validity/convergence checks
- Automatic OLS fallback when mixed models fail to converge

The application can recommend a model based on the comparison results.

### Interpretation

A **random intercept** model is appropriate when replicates differ primarily in baseline values.

A **random slope** model can be preferred when replicates also differ in their temporal trajectories.

If random effects do not provide sufficient improvement, an OLS model may be recommended as a simpler alternative.

> ⚠️ Model warnings, singular fits and convergence problems should be inspected before interpreting an LMM.

---

## 15. ANCOVA analysis

The **ANCOVA Analysis** tab provides analysis of covariance for adjusting physiological response variables for relevant continuous covariates.

ANCOVA can be used to determine whether group differences remain after accounting for environmental or physiological variation.

### ANCOVA capabilities

The application supports:

- Response-variable selection
- Covariate selection
- Covariate-response mapping
- Homogeneity-of-slopes testing
- Covariate-adjusted group effects
- Adjusted means / Least Squares Means
- Post-hoc comparisons of adjusted means
- Partial η² effect sizes
- Context-aware covariate recommendations
- Scientific interpretation and recommendations

The ANCOVA implementation prevents the same variable from being selected as both response and covariate and validates the selected channel set before analysis.

### Context-aware covariates

The application provides physiologically relevant covariate suggestions depending on the response variable.

Examples include:

| Response | Potential covariate | Interpretation |
|----------|---------------------|----------------|
| VPD | Leaf temperature | VPD adjusted for leaf temperature |
| VPD | Air temperature | VPD adjusted for air temperature |
| Leaf temperature | Absorbed radiation | Leaf temperature adjusted for radiation |
| Leaf temperature | Transpiration | Leaf temperature adjusted for transpiration |
| CO₂ compensation point | Leaf temperature | Γ adjusted for temperature |
| Time | Net photosynthesis | Time adjusted for photosynthetic response |

### ANCOVA interpretation

The recommended workflow is:

1. Select the physiological response variable
2. Select an appropriate continuous covariate
3. Set the significance level α
4. Select the channels/groups to compare
5. Run ANCOVA
6. Check homogeneity of slopes
7. Inspect adjusted group effects
8. Inspect adjusted means
9. Inspect post-hoc comparisons
10. Consider partial η² when interpreting effect magnitude

> ⚠️ ANCOVA should only be used when the selected covariate has a scientifically defensible relationship with the response and the model assumptions are reasonable.

---

## 16. Statistical diagnostic plots

The application provides diagnostic plots to help evaluate the suitability of statistical models.

Available diagnostics include:

- **Residual vs Fitted**
- **Q-Q plot**
- **Scale-Location**
- **Residuals by group/time**
- **Cook's Distance**
- **Effect-size visualizations**

These diagnostics can help identify non-normal residuals, heteroscedasticity, influential observations and other potential model problems.

### Recommended diagnostic workflow

Before interpreting inferential results:

1. Inspect residual-vs-fitted patterns
2. Inspect Q-Q plots
3. Check variance structure
4. Identify influential observations using Cook's Distance
5. Compare model assumptions with biological expectations
6. Use robust or mixed-model alternatives when appropriate

---

## 17. Compact Letter Display

For significant post-hoc comparisons, the application can produce a **Compact Letter Display (CLD)**.

Groups sharing the same letter are interpreted as **not significantly different** under the corresponding post-hoc test.

For example:

| Group | Mean | Letter |
|-------|------|--------|
| Treatment A | 15.4 | a |
| Treatment B | 13.8 | ab |
| Treatment C | 10.2 | b |

This provides a compact way to communicate statistical groupings in figures and tables.

> ⚠️ Compact letters describe the selected post-hoc comparison procedure and should not be interpreted independently of the underlying statistical test.

---

## 18. Statistical recommendations

The application evaluates the statistical results and provides recommendations based on assumptions and significance.

When ANOVA assumptions are met:

- Standard ANOVA is recommended
- Tukey HSD can be used for pairwise comparisons when the overall test is significant

When assumptions are not fully met:

- Welch's ANOVA may be recommended for unequal variances
- Games-Howell may be used for pairwise comparisons
- Kruskal-Wallis and Dunn's test provide a non-parametric alternative

For repeated measurements:

- Repeated Measures ANOVA can account for phase effects
- LMM can account for replicate-level random effects and individual trajectories

For environmental confounding:

- ANCOVA can adjust the response for an appropriate continuous covariate

---

## 19. Download and document results

### Single File Export

Before closing the app, use the **Download Data** button in Interactive Calculated Data tab:

**Export format options:**

- **RE-IMPORT:** Only original 28 columns, compatible with raw data import
- **ALL COLUMNS:** All calculated metrics plus original columns

### Darkness Data Export

Click **Download Darkness Data** in the Light-Dark tab:

- Same RE-IMPORT / ALL COLUMNS options
- Data includes darkness modifications (`absorbed_radiation = 0`)
- Handles multi-day experiments with correct sorting

### Radiation Data Export

Click **Download Radiation Data** in the Red+Blue Light tab:

- Same export format options
- Data includes radiation interval modifications
- Intervals applied before export

### Batch Statistical Export

Batch processing provides:

- `batch_statistics.xlsx`
- Combined processed data
- Averaged plots
- WUEi distribution plot
- Batch report

The complete-data export contains all calculated physiological variables and is suitable for further analysis in external statistical or visualization software.

### Plot Downloads

Each plot has its own **Download** button:

- PNG format at 300 DPI
- PDF format for vector graphics

---

## 20. Recommended workflow order

**Single file:**

1. Upload → Process File
2. Raw Data inspection
3. Channel renaming and selection
4. Time adjustments → Apply Changes
5. Parameter updates (one at a time)
6. Interactive Plots inspection (Environment → Efficiency)
7. Optional: Darkness analysis
8. Optional: Radiation intervals
9. Inspect calculated physiological metrics
10. Download data and plots

**Batch processing:**

1. Add multiple files (matching pattern)
2. (Optional) Load Excel mapping file
3. Set output directory
4. Process Batch
5. Review Batch Results tabs
6. Inspect descriptive statistics and SEM plots
7. Review darkness detection
8. Select appropriate statistical analysis
9. Check assumptions and diagnostic plots
10. Run ANOVA / robust alternative / repeated-measures ANOVA as appropriate
11. Use LMM when replicate-level random effects or repeated trajectories require them
12. Use ANCOVA when a scientifically relevant covariate should be controlled
13. Review post-hoc comparisons and compact letter displays
14. Export statistics, processed data and plots

> ⚠️ Statistical analysis should be performed only after the raw data, channel assignments, time handling, physical parameters and physiological calculations have been checked.

---

## 21. Statistical analysis decision guide

A practical decision sequence is:

```text
                    Batch data
                        │
                        ▼
             Inspect raw/calculated data
                        │
                        ▼
             Select response variable
                        │
                        ▼
             Are observations repeated?
                  /              \
                Yes              No
                 │                │
                 ▼                ▼
       Repeated Measures       Standard
            ANOVA              ANOVA
                 │                │
                 └──────┬─────────┘
                        ▼
                Check assumptions
                        │
              ┌─────────┴─────────┐
              │                   │
          Assumptions          Violations
             met                   │
              │              ┌─────┴─────┐
              ▼              ▼           ▼
        ANOVA + Tukey     Welch/       Kruskal-
                          Games-       Wallis/
                          Howell       Dunn
              │              │           │
              └──────────────┴───────────┘
                        │
                        ▼
             Consider LMM for repeated
             replicate-level structure
                        │
                        ▼
             Consider ANCOVA when a
             relevant covariate exists
                        │
                        ▼
              Inspect diagnostic plots
                        │
                        ▼
              Report effect sizes,
             adjusted means and
             post-hoc comparisons
```

---

## 22. Interpretation and reproducibility notes

For reproducible analysis:

- Keep the original raw TXT files unchanged
- Record any channel renaming
- Record leaf-area modifications
- Record air-flow modifications
- Record radiation modifications
- Record boundary-layer and cuticular conductance changes
- Record removed time points
- Record time shifts
- Record darkness interventions
- Record radiation intervals
- Record selected statistical response variables
- Record α levels
- Record statistical model choices
- Record covariates used for ANCOVA
- Retain diagnostic plots and statistical output

The batch system preserves combined processed data and stores statistics by base channel and normalized time point, supporting reproducible downstream comparisons.

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
