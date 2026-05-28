# multifile Gas Exchange Jyrkki multi-cuvette system
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib
matplotlib.use('TkAgg')
from matplotlib.patches import Rectangle
import io
import os
import sys
import subprocess
import json
from datetime import datetime
import math
import openpyxl
from tkinter import font as tkfont
import base64
import re
from matplotlib.patches import Patch, Rectangle
from matplotlib.lines import Line2D
from PIL import Image, ImageTk
import traceback
import csv
import gc
import shutil
from openpyxl.styles import Font

# ============================================
# DEPENDENCY CHECK AND INSTALLATION
# ============================================

def run_comprehensive_dependency_check():
    """Run complete dependency check and verification"""
    # Import standard library modules that are always available
    import sys
    import os
    
    print("\n" + "=" * 60)
    print("GAS EXCHANGE ANALYZER - DEPENDENCY CHECK")
    print("=" * 60)
    print(f"Python version: {sys.version}")
    print(f"Platform: {sys.platform}")
    print(f"Executable: {sys.executable}")
    
    try:
        # Core libraries
        import pandas as pd
        import numpy as np
        import matplotlib
        matplotlib.use('TkAgg')
        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        from matplotlib.patches import Rectangle, Patch
        from matplotlib.lines import Line2D
        
        # Tkinter libraries
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox, scrolledtext
        from tkinter import font as tkfont
        
        # Image processing
        from PIL import Image, ImageTk
        
        # Excel handling
        import openpyxl
        from openpyxl.styles import Font
        
        # Data handling
        import json
        import csv
        
        # System and utilities
        import io
        import subprocess
        import base64
        import re
        import math
        import gc
        import shutil
        import traceback
        from datetime import datetime
        
        print("\n✓ All required packages are available!")
        print("  - pandas, numpy, matplotlib")
        print("  - tkinter (built-in)")
        print("  - PIL/Pillow")
        print("  - openpyxl")
        print("  - All standard library modules")
        return True
        
    except ImportError as e:
        print(f"\n✗ Missing dependency: {e}")
        try:
            # Import only what's needed for the error dialog
            import tkinter as tk
            from tkinter import messagebox
            
            root = tk.Tk()
            root.withdraw()
            
            # Provide specific installation instructions
            missing_pkg = str(e).split("'")[-2] if "'" in str(e) else str(e)
            
            install_cmds = {
                'pandas': 'pip install pandas',
                'numpy': 'pip install numpy',
                'matplotlib': 'pip install matplotlib',
                'PIL': 'pip install Pillow',
                'openpyxl': 'pip install openpyxl'
            }
            
            cmd = install_cmds.get(missing_pkg, f'pip install {missing_pkg}')
            
            messagebox.showerror(
                "Dependency Error",
                f"Missing required package: {missing_pkg}\n\n"
                f"Please install it using:\n{cmd}\n\n"
                "Or install all dependencies:\n"
                "pip install pandas numpy matplotlib pillow openpyxl"
            )
            root.destroy()
        except:
            pass
        return False

# ============================================
# DATA PROCESSING FUNCTIONS
# ============================================

def calculate_common_metrics(df):
    if df is None or len(df) == 0:
        return df
    
    # Step 1: Mass flow corrections and gradients
    df['saturating_air_humidity'] = 6.04 * np.exp(17.569 * df['air_temp'] / (241.9 + df['air_temp']))
    df['relative_air_humidity'] = (df['H2O_M'] / df['saturating_air_humidity']) * 100
    
    # Step 2: CO2 and H2O related rates
    df['CO2_exchange_rate'] = df['air_flow_rate'] * (df['CO2_R'] - df['CO2_M']) / df['leaf_area']
    df['Transpiration_H2O_evol_rate'] = df['air_flow_rate'] * (df['H2O_M'] - df['H2O_R']) / df['leaf_area']
    df['relative_chamber_air_humidity'] = df['H2O_M'] / df['saturating_air_humidity']
    df['VPD'] = df['saturating_air_humidity'] * (1 - df['relative_chamber_air_humidity']) / 10
    
    # Step 3: Leaf and temperature-related calculations
    df['transp_ind_leaf_temp_depr'] = 0.773 * df['Transpiration_H2O_evol_rate']
    df['rad_ind_leaf_temp_increase'] = df['absorbed_radiation'] / 0.00137
    df['leaf_temp_c'] = df['air_temp'] - df['transp_ind_leaf_temp_depr'] + df['rad_ind_leaf_temp_increase']
    
    return df

def calculate_conductance_metrics(df):
    if df is None or len(df) == 0:
        return df
    
    # Step 4: Leaf and air humidity gradients
    df['satur_hum_at_leaf_temp_c'] = 6.04 * np.exp(17.569 * df['leaf_temp_c'] / (241.9 + df['leaf_temp_c']))
    df['leaf_air_hum_grad'] = df['satur_hum_at_leaf_temp_c'] - df['H2O_M']
    
    # Step 5: Overall conductance and corrections
    df['overall_conductance'] = 1000 * df['Transpiration_H2O_evol_rate'] / df['leaf_air_hum_grad']
    df['saturation_vapor_pressure_inside_leaves'] = df['satur_hum_at_leaf_temp_c'] * 1.013
    df['vapor_pressure_around_leaves'] = df['H2O_M'] * 1.013
    df['mean_vapor_pressure'] = (df['saturation_vapor_pressure_inside_leaves'] + df['vapor_pressure_around_leaves']) / 2
    df['massflow_correction_for_overall_resistance'] = 1 + (df['mean_vapor_pressure'] / (1013 - df['mean_vapor_pressure']))
    
    df['corr_d_overall_conductance'] = (1000 * df['Transpiration_H2O_evol_rate'] * (273 + df['leaf_temp_c'])) / \
        (df['leaf_air_hum_grad'] * 446 * 273 * df['massflow_correction_for_overall_resistance'])
    
    df['b_layer_conductance'] = (df['boundary_layer_cond'] * 446 * 273) / (273 + df['leaf_temp_c'])
    df['corr_leaf_conductance'] = df['corr_d_overall_conductance'] * df['boundary_layer_cond'] / \
        (df['boundary_layer_cond'] - df['corr_d_overall_conductance'])
    df['corr_stomatal_conductance'] = df['corr_leaf_conductance'] - df['cutic_conductance']
    
    df['massflow_corr_for_CO2'] = 1 - (1.62 * (df['saturation_vapor_pressure_inside_leaves'] - df['vapor_pressure_around_leaves']) / 1013)
    df['mflowcorr_Ca_gradient'] = 1 - (1.62 * (df['saturation_vapor_pressure_inside_leaves'] - df['vapor_pressure_around_leaves']) / 2026)
    
    return df

def calculate_stomatal_metrics(df):
    if df is None or len(df) == 0:
        return df
    
    # Step 6: Stomatal conductance corrections
    df['Stomatal_conductance_corrected'] = (df['corr_stomatal_conductance'] * 446 * 273) / \
        (273 + df['leaf_temp_c'])
    
    # Normalise corrected stomatal conductance
    if len(df) > 0 and not np.isnan(df['Stomatal_conductance_corrected'].iloc[0]) and df['Stomatal_conductance_corrected'].iloc[0] != 0:
        df['normalised_corr_stomatal_conductance'] = df['Stomatal_conductance_corrected'] / df['Stomatal_conductance_corrected'].iloc[0]
    else:
        df['normalised_corr_stomatal_conductance'] = np.nan
    
    # Step 7: Change rates
    df['corr_stomatal_change_rate'] = 0.0
    
    if len(df) > 1:
        df.loc[:len(df)-2, 'corr_stomatal_change_rate'] = \
            (df['Stomatal_conductance_corrected'].iloc[1:].values - df['Stomatal_conductance_corrected'].iloc[:len(df)-1].values) / 2
        df.loc[len(df)-1, 'corr_stomatal_change_rate'] = (0 - df['Stomatal_conductance_corrected'].iloc[-1]) / 2
    
    df['relat_stom_cond_change_rate'] = 0.0
    if len(df) > 1:
        df.loc[:len(df)-2, 'relat_stom_cond_change_rate'] = \
            np.diff(df['normalised_corr_stomatal_conductance']) / 2 * 100
        df.loc[len(df)-1, 'relat_stom_cond_change_rate'] = \
            ((0 - df['normalised_corr_stomatal_conductance'].iloc[-1]) / 2) * 100
    
    return df

def calculate_ozone_metrics(df):
    if df is None or len(df) == 0:
        return df
    
    # Ozone-related calculations
    df['overall_O3_uptake_rate_c'] = df['air_flow_rate'] * (df['O3_M'] - df['O3_R']) / df['leaf_area']
    df['corr_conductance_for_O3_c'] = (df['b_layer_conductance'] * df['Stomatal_conductance_corrected'] * 0.62) / \
        (df['b_layer_conductance'] + df['Stomatal_conductance_corrected'])
    df['corr_stomatal_O3_uptake_rate'] = 0.001 * df['corr_conductance_for_O3_c'] * df['O3_M']
    
    df['corr_cumul_O3_dose'] = 0.0
    if len(df) > 0:
        df.loc[0, 'corr_cumul_O3_dose'] = df.loc[0, 'corr_stomatal_O3_uptake_rate'] * 0.001 * 960
        if len(df) > 1:
            for i in range(1, len(df)):
                df.loc[i, 'corr_cumul_O3_dose'] = df.loc[i-1, 'corr_cumul_O3_dose'] + (df.loc[i, 'corr_stomatal_O3_uptake_rate'] * 0.001 * 960)
    
    return df

def calculate_co2_metrics(df):
    if df is None or len(df) == 0:
        return df
    
    # CO2-related corrections
    df['massflow_corr_d_ca'] = df['CO2_M'] * df['massflow_corr_for_CO2']
    
    # Calculate stomatal resistance to CO2
    df['stom_resist_to_CO2'] = 1.62 * 2.24 * (273 + df['leaf_temp_c']) / (273 * df['corr_stomatal_conductance'])
    df['b_layer_resist_to_CO2'] = 1.62 * 2.24 * (273 + df['air_temp']) / (273 * df['boundary_layer_cond'])
    
    # Calculate CO2 gradients
    df['corr_ted_CO2_grad'] = df['CO2_exchange_rate'] * \
        (df['stom_resist_to_CO2'] + df['b_layer_resist_to_CO2']) * df['mflowcorr_Ca_gradient']
    df['intercell_CO2_conc_in_gas_phase'] = df['massflow_corr_d_ca'] - df['corr_ted_CO2_grad']
    
    # mesophyll conductance for CO2
    mesophyll_conductance = []
    for i in range(len(df)):
        if i + 2 < len(df):
            CO2_exchange_rate_i2 = df.loc[i + 2, 'CO2_exchange_rate']
            intercell_CO2_conc_in_gas_phase_i2 = df.loc[i + 2, 'intercell_CO2_conc_in_gas_phase']
        else:
            CO2_exchange_rate_i2 = 0
            intercell_CO2_conc_in_gas_phase_i2 = 0
        
        if (np.isnan(intercell_CO2_conc_in_gas_phase_i2) or np.isnan(df.loc[i, 'intercell_CO2_conc_in_gas_phase']) or
            df.loc[i, 'intercell_CO2_conc_in_gas_phase'] == intercell_CO2_conc_in_gas_phase_i2):
            mesophyll_conductance.append(np.nan)
        else:
            value = ((df.loc[i, 'CO2_exchange_rate'] - CO2_exchange_rate_i2) * 100) / \
                (df.loc[i, 'intercell_CO2_conc_in_gas_phase'] - intercell_CO2_conc_in_gas_phase_i2)
            mesophyll_conductance.append(value)
    
    df['mesophyll_conductance_for_CO2'] = mesophyll_conductance
    
    # Apply threshold for mesophyll conductance
    threshold = 60000000000
    for i in range(len(df)):
        if not np.isnan(df.loc[i, 'mesophyll_conductance_for_CO2']) and \
           abs(df.loc[i, 'mesophyll_conductance_for_CO2']) > threshold:
            df.loc[i, 'mesophyll_conductance_for_CO2'] = 600 if df.loc[i, 'mesophyll_conductance_for_CO2'] > 0 else -600
    
    return df

def calculate_additional_metrics(df):
    if df is None or len(df) == 0:
        return df
    
    # Additional metrics
    df['CO2_comp_point'] = (df['mesophyll_conductance_for_CO2'] * df['intercell_CO2_conc_in_gas_phase'] * 0.001 - 
                             df['CO2_exchange_rate']) / (df['mesophyll_conductance_for_CO2'] * 0.001)
    
    df['corr_uncorr_stom_cond'] = df['Stomatal_conductance_corrected'] / df['stomatal_cond']
    df['corr_uncorr_ca'] = df['massflow_corr_d_ca'] / df['CO2_M']
    df['uncorr_ci'] = df['CO2_M'] - (df['CO2_exchange_rate'] * (df['stom_resist_to_CO2'] + df['b_layer_resist_to_CO2']))
    df['corr_uncorr_ci'] = df['intercell_CO2_conc_in_gas_phase'] / df['uncorr_ci']
    
    # uncorr_gm_prima
    uncorr_gm_prima = []
    for i in range(len(df)):
        if i + 2 < len(df):
            CO2_exchange_rate_i2 = df.loc[i + 2, 'CO2_exchange_rate']
            uncorr_ci_i2 = df.loc[i + 2, 'uncorr_ci']
        else:
            CO2_exchange_rate_i2 = 0
            uncorr_ci_i2 = 0
        
        if (np.isnan(uncorr_ci_i2) or np.isnan(df.loc[i, 'uncorr_ci']) or 
            df.loc[i, 'uncorr_ci'] == uncorr_ci_i2):
            uncorr_gm_prima.append(np.nan)
        else:
            value = (df.loc[i, 'CO2_exchange_rate'] - CO2_exchange_rate_i2) * 100 / (df.loc[i, 'uncorr_ci'] - uncorr_ci_i2)
            uncorr_gm_prima.append(value)
    
    df['uncorr_gm_prima'] = uncorr_gm_prima
    df['corr_uncorr_gm_prima'] = df['mesophyll_conductance_for_CO2'] / df['uncorr_gm_prima']
    df['uncorr_gamma'] = (df['uncorr_gm_prima'] * df['uncorr_ci'] * 0.001 - df['CO2_exchange_rate']) / (df['uncorr_gm_prima'] * 0.001)
    df['corr_uncorr_gamma'] = df['CO2_comp_point'] / df['uncorr_gamma']
    
    # Water use efficiency calculations
    df['WUEi'] = df['CO2_exchange_rate'] * 1000 / df['Stomatal_conductance_corrected']
    df['WUE'] = df['CO2_exchange_rate'] / df['Transpiration_H2O_evol_rate']
    
    return df

def process_time_columns(df, time0_shift=0, remove_start=0, remove_end=0):
    """
    IMPROVED: Force integer time intervals across all replicates
    """
    if df is None or len(df) == 0:
        return df
    
    processed_data = []
    
    # Process each channel independently
    for channel in df['Channel'].unique():
        channel_df = df[df['Channel'] == channel].copy().reset_index(drop=True)
        
        # Convert HH:MM:SS → seconds since midnight
        time_seconds = []
        for time_str in channel_df['time']:
            parts = str(time_str).split(':')
            if len(parts) >= 3:
                secs = float(parts[0]) * 3600 + float(parts[1]) * 60 + float(parts[2])
            else:
                secs = 0.0
            time_seconds.append(secs)
        
        # Detect day rollover
        day_increment = 0
        full_seconds = []
        prev_sec = time_seconds[0] if time_seconds else 0
        
        for i, sec in enumerate(time_seconds):
            if sec < prev_sec and i > 0:
                day_increment += 1
            prev_sec = sec
            full_seconds.append(sec + day_increment * 24 * 3600)
        
        # Continuous time in minutes from actual clock time
        full_minutes = [sec / 60 for sec in full_seconds]
        
        # Calculate the time interval - ROUND TO NEAREST INTEGER
        if len(full_minutes) > 1:
            intervals = np.diff(full_minutes)
            positive_intervals = intervals[intervals > 0]
            if len(positive_intervals) > 0:
                # ROUND to nearest integer! This is the key fix
                time_interval_minutes = int(np.round(np.median(positive_intervals)))
            else:
                time_interval_minutes = 1
        else:
            time_interval_minutes = 1
        
        # Apply remove_start and remove_end
        total_original_points = len(channel_df)
        start_idx = remove_start
        end_idx = total_original_points - remove_end
        
        if start_idx < 0:
            start_idx = 0
        if end_idx > total_original_points:
            end_idx = total_original_points
        
        if start_idx < end_idx:
            channel_df = channel_df.iloc[start_idx:end_idx].reset_index(drop=True)
            full_minutes = full_minutes[start_idx:end_idx]
            full_seconds = full_seconds[start_idx:end_idx]
            time_seconds = time_seconds[start_idx:end_idx]
        
        # CRITICAL FIX: Force integer time points
        if len(channel_df) > 0:
            # Create integer time points starting from 0
            n_points = len(channel_df)
            # Generate integer time points: 0, interval, 2*interval, etc.
            integer_time_points = [i * time_interval_minutes for i in range(n_points)]
            
            # Apply time0_shift
            time_repeated_list = [t - (time0_shift * time_interval_minutes) for t in integer_time_points]
            
            # Add columns to dataframe
            channel_df['time_repeated'] = time_repeated_list
            channel_df['time_sec'] = [t * 60 for t in integer_time_points]  # Convert to seconds
            channel_df['time_full_sec'] = full_seconds
            channel_df['time_full_minutes'] = full_minutes
            channel_df['time_interval_minutes'] = time_interval_minutes
            
        else:
            channel_df['time_repeated'] = []
            channel_df['time_sec'] = []
            channel_df['time_full_sec'] = []
            channel_df['time_full_minutes'] = []
            channel_df['time_interval_minutes'] = time_interval_minutes
        
        channel_df = channel_df.loc[:, ~channel_df.columns.duplicated()]
        processed_data.append(channel_df)
    
    if processed_data:
        result = pd.concat(processed_data, ignore_index=True)
        return result
    return df

# ============================================
# PLOTTING UNITS FUNCTIONS
# ============================================

UNITS_MAP = {
    # Environment tab
    'air_temp': 'ºC',
    'saturating_air_humidity': 'mmol/mol',
    'absorbed_radiation': 'cal/cm²s',
    'relative_air_humidity': '%',
    'VPD': 'kPa',
    'transp_ind_leaf_temp_depr': 'ºC',
    'rad_ind_leaf_temp_increase': 'ºC',
    
    # Base Fluxes tab
    'CO2_exchange_rate': 'mkmol/m²s',
    'Transpiration_H2O_evol_rate': 'mmol/m²s',
    'leaf_temp_c': 'ºC',
    
    # Humidity/Vapor tab
    'leaf_air_hum_grad': 'mmol/mol',
    'satur_hum_at_leaf_temp_c': 'mmol/mol',
    'saturation_vapor_pressure_inside_leaves': 'mbar',
    'vapor_pressure_around_leaves': 'mbar',
    'mean_vapor_pressure': 'mbar',
    
    # Conductances tab
    'overall_conductance': 'mmol/m²s',
    'corr_d_overall_conductance': 'cm/s',
    'b_layer_conductance': 'cm/s',
    'corr_leaf_conductance': 'cm/s',
    'corr_stomatal_conductance': 'cm/s',
    'Stomatal_conductance_corrected': 'mmol/m²s',
    'normalised_corr_stomatal_conductance': 'dim-less',
    'massflow_correction_for_overall_resistance': 'dim-less',
    'corr_uncorr_stom_cond': 'dim-less',
    'corr_stomatal_change_rate': 'mmol/m²s',
    'relat_stom_cond_change_rate': 'dim-less',
    
    # Ozone tab
    'overall_O3_uptake_rate_c': 'nmol/m²s',
    'corr_conductance_for_O3_c': 'mmol/m²s',
    'corr_stomatal_O3_uptake_rate': 'nmolO₃/m²s',
    'corr_cumul_O3_dose': 'mkmol/m²',
    
    # CO2 tab - UPDATED with all required variables
    'massflow_corr_d_ca': 'mkmol CO₂/mol air',
    'massflow_corr_for_CO2': 'dim-less',
    'mflowcorr_Ca_gradient': 'dim-less',
    'corr_uncorr_ca': 'dim-less',
    'stom_resist_to_CO2': 'm²s/mmol',
    'b_layer_resist_to_CO2': 'm²s/mmol',
    'corr_ted_CO2_grad': 'mkmolCO₂/mol air',
    'intercell_CO2_conc_in_gas_phase': 'mkmol CO₂/mol air',
    'uncorr_ci': 'mkmol CO₂/mol air',
    'corr_uncorr_ci': 'dim-less',
    
    # Mesophyll tab
    'mesophyll_conductance_for_CO2': 'mmol/m²s',
    'uncorr_gm_prima': 'mmol/m²s',
    'corr_uncorr_gm_prima': 'dim-less',
    'CO2_comp_point': 'mkmol CO₂/mol air',
    'uncorr_gamma': 'mkmol CO₂/mol air',
    'corr_uncorr_gamma': 'dim-less',
    
    # Efficiency tab
    'WUEi': 'µmol CO₂ / mol H₂O',
    'WUE': 'µmol CO₂ / mmol H₂O',
    
    # Additional variables that might be in your data
    'saturating_air_humidity': 'mmol/mol',
    'leaf_temp': 'ºC',
    'stomatal_cond': 'mmol/m²s',
    'boundary_layer_cond': 'mmol/m²s',
    'cutic_conductance': 'mmol/m²s',
    'leaf_area': 'cm²',
    'air_flow_rate': 'cm³/s',
    'CO2_M': 'µmol/mol',
    'CO2_R': 'µmol/mol',
    'H2O_M': 'mmol/mol',
    'H2O_R': 'mmol/mol',
    'O3_M': 'nmol/mol',
    'O3_R': 'nmol/mol',
    'CO2_exch_rate': 'µmol/m²s',
    'H2O_evol_rate': 'mmol/m²s',
    'overall_O3_uptake_rate': 'nmol/m²s',
    'satur_hum_at_leaf_temp': 'mmol/mol',
    'leaf_air_hum_grad': 'mmol/mol',
    'stomatal_O3_uptake_rate': 'nmol/m²s',
    'O3_cumul_dose': 'µmol/m²',
    'stom_change_rate': 'mmol/m²s',
    'normalised_stom_cond': 'dim-less',
    'relstom_condchange_rate': '%',
    'temp_calibr': 'ºC',
    
    # Time columns
    'time_repeated': 'min',
    'time_sec': 's',
    'time_full_sec': 's',
    'time_full_minutes': 'min',
    'time_interval_minutes': 'min'
}


# ============================================
# COLOR PALETTES
# ============================================
# Define colorblind-friendly palettes
colorblind_friendly_palettes = {
        'tab10': plt.cm.tab10,
        'Set2': plt.cm.Set2,
        'Set3': plt.cm.Set3,
        'Dark2': plt.cm.Dark2,
        'Paired': plt.cm.Paired,
        'viridis': plt.cm.viridis,
        'plasma': plt.cm.plasma,
        'magma': plt.cm.magma,
        'cividis': plt.cm.cividis,
        'inferno': plt.cm.inferno,
        'colorblind1': ['#E69F00', '#56B4E9', '#009E73', '#F0E442', '#0072B2', '#D55E00', '#CC79A7', '#000000'],
        'colorblind2': ['#377EB8', '#FF7F00', '#4DAF4A', '#F781BF', '#A65628', '#984EA3', '#999999', '#E41A1C'],
        'colorblind3': ['#4477AA', '#EE6677', '#228833', '#CCBB44', '#66CCEE', '#AA3377', '#BBBBBB', '#000000'],
    }

# ============================================
# COEFFICIENT TABLES
# ============================================

def get_coefficient_table_data():
    """Get coefficient table data """
    coef_data = {
        'Color': ['RED', 'GREEN', 'BLUE', 'WHITE'],
        '15 cm': [650, 250, 600, 280],
        '10 cm': [1000, 380, 1000, 450],
        '5 cm': [2000, 680, 1900, 800]
    }
    return pd.DataFrame(coef_data)

def get_color_table_data():
    """Get color table data"""
    color_data = pd.DataFrame({
        'Light_Intensity': range(50, 1201, 25),
        'WHITE': [0.000197,0.000295,0.000394,0.000492,0.000591,0.000689,0.000788,0.000886,0.000985,
                  0.001083,0.001182,0.00128,0.001379,0.001477,0.001576,0.001674,0.001773,0.001871,
                  0.001970,0.002068,0.002167,0.002264,0.002364,0.002462,0.002561,0.002659,0.002758,
                  0.002856,0.002955,0.003053,0.003152,0.003250,0.003349,0.003447,0.003546,0.003644,
                  0.003743,0.003841,0.003940,0.004038,0.004137,0.004235,0.004334,0.004432,0.004531,
                  0.004629,0.004728],
        'RED': [0.000186,0.000279,0.000372,0.000465,0.000558,0.000651,0.000744,0.000837,0.000930,
                0.001023,0.001116,0.001208,0.001301,0.001394,0.001487,0.001580,0.001673,0.001766,
                0.001859,0.001952,0.002045,0.002138,0.002231,0.002324,0.002417,0.002510,0.002603,
                0.002696,0.002789,0.002882,0.002975,0.003068,0.003161,0.003254,0.003347,0.003440,
                0.003533,0.003625,0.003718,0.003811,0.003904,0.003997,0.004090,0.004183,0.004276,
                0.004369,0.004462],
        'BLUE': [0.000296,0.000444,0.000592,0.000740,0.000888,0.001036,0.001184,0.001332,0.001480,
                 0.001628,0.001775,0.001923,0.002071,0.002219,0.002367,0.002515,0.002663,0.002811,
                 0.002959,0.003107,0.003255,0.003403,0.003551,0.003699,0.003847,0.003995,0.004143,
                 0.004291,0.004439,0.004587,0.004735,0.004883,0.005031,0.005179,0.005326,0.005474,
                 0.005622,0.005770,0.005918,0.006066,0.006214,0.006362,0.006510,0.006658,0.006806,
                 0.006954,0.007102]
    })
    
    return color_data


# ============================================
# SPLASH SCREEN CLASS
# ============================================

class SplashScreen:
    def __init__(self):
        self.root = tk.Tk()
        self.root.overrideredirect(True)  # Remove window decorations
        
        # Get screen dimensions
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        img_bytes = b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x01\x80\x00\x00\x01\x80\x08\x06\x00\x00\x00\xa4\xc7\xb5\xbf\x00\x00\x01\x81iCCPICC profile\x00\x00(\xcf\x95\x91;H\x03A\x18\x84\xbf$JD"\x16\xa6\x10\xb1\xb8B\xadL\xa3"\x96\x1a\x05\x11"\x84\xa8\x90\xa8\x85w\x17\x13\x03\xb93\xdcEl,\x05\xdb\x80\x85\x8f\xc6Wac\xad\xad\x85\xad \x08>@\xec\x05+E\x1b\x91\xf3\xdf$\x90 Dpa\xd9\x8f\xd9\x9daw\x16\xfc\x87y\xd3r\x9bF\xc1\xb2\x8bNb2\xaa%S\xf3Z\xf0\x85\x10\xad\x04\x89\x10\xd0M\xb70\x16\x8f\xc7h8>\xef\xf0\xa9\xf56\xa2\xb2\xf8\xdfhK/\xbb&\xf84\xe1Q\xb3\xe0\x14\x85\x97\x84\x87\xd7\x8b\x05\xc5\xbb\xc2asEO\x0b\x9f\t\xf7;rA\xe1\x07\xa5\x1b\x15~U\x9c-\xb3_e\x86\x9d\xd9\xc4\xb8pXX\xcb\xd6\xb1Q\xc7\xe6\x8ac\t\x0f\t\xf7\xa4-[\xf2\xfd\xc9\n\xa7\x15o(\xb6\xf2kf\xf5\x9e\xea\x85\xa1e{nF\xe92\xbb\x99d\x8ai\xe2h\x18\xac\x91#OQ\xfa\xcaa\x8b\xe2\x92\x90\xfdh\x03\x7fW\xd9\x1f\x17\x97!\xae\x1c\xa68&X\xc5B/\xfbQ\x7f\xf0\xbb[738PI\nE\xa1\xf9\xd9\xf3\xde{!\xb8\r\xdf%\xcf\xfb:\xf2\xbc\xefc\x08<\xc1\xa5]\xf3\xaf\x1e\xc2\xc8\x87\xe8\xa5\x9a\xd6s\x00\xed\x9bp~U\xd3\x8c\x1d\xb8\xd8\x82\xce\xc7\x82\xee\xe8e) \xd3\x9f\xc9\xc0\xdb\xa9|S\n:n\xa0u\xa1\xd2[u\x9f\x93{\x98\x95\xaeb\xd7\xb0\xb7\x0f}Y\xc9^l\xf0\xee\x96\xfa\xde\xfe<S\xed\xef\x07`\xddr\x9f.\\=\x92\x00\x00\x00\tpHYs\x00\x00\x0b\x11\x00\x00\x0b\x11\x01\x7fd_\x91\x00\x00\x00\xc1zTXtRaw profile type exif\x00\x00x\xdamPA\x12\xc3 \x08\xbc\xf3\x8a>\x01\x01\t>\xc74\xe9L\x7f\xd0\xe7\x17E;\xb1\xed\xce\xb8K\\\xdd p\xbe\x9e\x0f\xb85P\x12\x90\xbc\x99\x16UtH\x91B\xd5\x0b\xc3@\xed\x9cP:w\xd8\xac\xd2\xba\x0f\x1f\x83\\\xd9\x95\xc7\x05\rMs\x7f\\\x98\x9a\xaaW\xf9\x1at\x1f\xc6\xbe\x1aEB\xc9\xbe\x82(\x84[G\xad>FP\x19ALa\xa4\x11P\xe3Y\xa8\xc5\xb6\xeb\x13\xf6\x13WX,h$\xb6\xb6\xfd\xf3\xbd\xf9\xf4\x8e\xec\xffa\xa2\x93\x13\xa33\xb3E\x03\xdcV\x06\xaenP\xe7v\x10\xb9x-\xac\xce\xc4:\xc2| \xff\xe64\x01o\xee\xecY"\xf2\xd8\xd6Z\x00\x00\x0e[iTXtXML:com.adobe.xmp\x00\x00\x00\x00\x00<?xpacket begin="\xef\xbb\xbf" id="W5M0MpCehiHzreSzNTczkc9d"?>\r\n<x:xmpmeta xmlns:x="adobe:ns:meta/" x:xmptk="XMP Core 4.4.0-Exiv2">\r\n\t<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">\r\n\t\t<rdf:Description rdf:about="" xmlns:xmpMM="http://ns.adobe.com/xap/1.0/mm/" xmlns:stEvt="http://ns.adobe.com/xap/1.0/sType/ResourceEvent#" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:GIMP="http://www.gimp.org/xmp/" xmlns:tiff="http://ns.adobe.com/tiff/1.0/" xmlns:xmp="http://ns.adobe.com/xap/1.0/" xmpMM:DocumentID="gimp:docid:gimp:436576ab-dd5f-4849-8ee6-e131c7859c37" xmpMM:InstanceID="xmp.iid:ed6d2298-8186-4987-b2db-9b1780799106" xmpMM:OriginalDocumentID="xmp.did:b09ee49e-ada0-4f94-819a-2f7f1fc618de" dc:Format="image/png" GIMP:API="2.0" GIMP:Platform="Linux" GIMP:TimeStamp="1765352788424891" GIMP:Version="2.10.38" tiff:Orientation="1" xmp:CreatorTool="GIMP 2.10" xmp:MetadataDate="2025:12:10T09:46:26+02:00" xmp:ModifyDate="2025:12:10T09:46:26+02:00">\r\n\t\t\t<xmpMM:History>\r\n\t\t\t\t<rdf:Seq>\r\n\t\t\t\t\t<rdf:li stEvt:action="saved" stEvt:changed="/" stEvt:instanceID="xmp.iid:03de2e2c-1cd6-4949-b1d6-343f1dfe7abc" stEvt:softwareAgent="Gimp 2.10 (Linux)" stEvt:when="2025-09-21T16:12:00+03:00"/>\r\n\t\t\t\t\t<rdf:li stEvt:action="saved" stEvt:changed="/" stEvt:instanceID="xmp.iid:2ec3c20e-5c04-4f85-9731-8ca57b386dd5" stEvt:softwareAgent="Gimp 2.10 (Linux)" stEvt:when="2025-12-10T09:46:28+02:00"/>\r\n\t\t\t\t</rdf:Seq>\r\n\t\t\t</xmpMM:History>\r\n\t\t</rdf:Description>\r\n\t</rdf:RDF>\r\n</x:xmpmeta>\r\n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                                                                                    \n                                  <?xpacket end=\'w\'?>\xc4\xc3\xd1\xfe\x00\x00\x00\x06bKGD\x00\xff\x00\xff\x00\xff\xa0\xbd\xa7\x93\x00\x00\x00\x07tIME\x07\xe9\x0c\n\x07.\x1cM\x9b\xaei\x00\x00\xee\xdcIDATx^\xec\x9dw`\x14E\x17\xc0\x7f\xbbw\x97K\xef!$4\xe9\x1d\xa4w\x10\xa5WA@\xc5^\xb1\x016DAD, \xc5\x02H\x11\xb1\xa0\x88\xf0\xd1T@@\x8a\xd2\x04\x95\x0eJ\xef$\x01B \xbd\\\xd9\xf9\xfe\xd8\xbd\xcb\xdd\xe5R\x804\xe0~z$\xd9\x99\xdb2;\xf3\xde\xcc\x9b7o$\xab\xd5*\xf0P8\x08@r=X\x8a(\xed\xf7wK\xe3)|\x0f\xa5\x0f\xd9\xf5\x80\x87\x1b\xa0\xb4\xb7\xef\xd2~\x7f\xb74\x9e\xc2\xf7P\xfa\xf0(\x00wx\xc6D\x1e<x\xb8\r\xf0(\x00wx:k\x1e<\xdc\xdax:y\xe0Q\x00\x1e<x\xb8-\xf1t\xf2\xc0\xa3\x00<\xdc,x:l\x1e<\x14>\x1e\x05p\x93!\xc4\xed)\no\xbc\xc3\x96\xb3\xdclGr\xa6\xdc\xbe\xdc\xae\xf5\xebv\xc5\xa3\x00n2$\xe9\xc6E\xe1\xedI\xcer\x93\x10\xaa\xf8\x17\xc2\xa3\x044<\xf5\xeb\xf6B*\xacu\x00B\x08O\xe5\xf1P\x8c\xdc\x80_\xbd\x10$\xa7\xa4p\xe0\xc0\x01\x12\x12\x12h\xdb\xb6\r!\xc1\xa1H\xf2u\x9e\xcf\x83\x87\x9b\x94BS\x00\x1e<\x94f\xd4J.\xc8\xca\xcc\xe2\xcc\xd93|\xfd\xd5\xd7,\\\xb8\x90\xac\xac,\xfa\xf7\xef\xcfs\xcf?O\xb5\xaaU\xf1\xf1\xf1)\xf2\x8e\x8c\xa7\xb3\xe4\xa1\xb4\xe0Q\x00\x1enm\xb4\x81\x82\xa2X9{\xee<\x7fn\xdb\xc6\x94)S\xd8\xb7o\x9fS\xb6r\xe5\xca3r\xe4\x1bt\xe9\xd2\x85\xca\x95+\xa3\xd7\xeb\x9d\xd2=x\xb8\x15\xf1(\x00\x0f\xb7<iii\xfc\xf1\xc7\x1f\xfc\xf0\xc3\x02\xd6\xad_\xc7\x95\x84\x04\xd7,\x00\x04\x06\x04\xd0\xf1\x9e{\x18x\xdf\x00z\xf5\xee\x85\xbf\xbf\xbf\xa7\xa7\xee\xe1\x96\xa6\xf4*\x80\x1b0\xf1z\xb8\xc9)\x84w/\x84Bf\x96\x89\x93\'O\xf2\xe5\x9c9\xfc\xef\x7f\xff\xe3\xe2\xc5K\x05\xf2\xf9\t\x0b\x0b\xa3W\xaf^<\xff\xfc\xf3\xd4\xaa]\x0b\x7f?\x7f\x90$\x84\x10\xc8\x1e\x85\xe0\xe1\x16\xa2\xf4*\x00\x0f\x1e\xae\x13!\x14N\x9e:\xc5\x1f\x7f\xfc\xc1\xdc/\xe7\xf2\xf7\xdf\x7f\xbbf)\x10\x95\xee\xa8\xc4\xcb\xc3_\xa6c\xc7\x8e\xd4\xacY\x0b\x83\xc1\x80G\xfe{\xb8\x95(z\x05P\x08\xbd9\x0f\x1e\n\x82\x10\x82\x94\x94\x14~]\xfd+\xcb\x96.c\xe5\xca\x95dee\xb9f\xbb&$I\xa2s\xe7\xce\x0c\x1c8\x88\xde\xbd{\x11\x1e\x1e\xee1\x0by\xb8e\x90\xac\x16\xab\xf0\x08h\x0f73V\xab\x85\x8c\x8cL\x8e\x1f?\xcew\xdf}\xc7\xdc\xb9sIKO\xcb\xd7\xda\x13\x16\x16Ftt4\x07\x0e\x1cpM\xca\x81\xbf\x9f?\xad\xdb\xb4f\xc4\xeb#h\xdc\xb41\x01\x01\x01\xe8$\x19uH\xe0\xe9\xe5x\xb89)\xfa\x11\x80\x07\x0fE\x88\xa2(\x1c=z\x945k\xd7\xf0\xe5\x97_r\xf8\xd0a\xd7,ny\xf0\x81\x07\x180p !!!,]\xba\x94\x193f\xb8fqKTT\x14\xcf?\xff<]\xbav\xa5A\xfd\xfa\x18\x8dF\xd7,\x1e<\xdc4x\x14\x80\x87\xdc)\xc5\x1d[\xa1\x08\xae&^e\xcd\x9a5,^\xbc\x98_~\xf9\xc55K\x0edY\xa6v\xed\xda<\xf9\xd4\x93\xdc\xd7\xff>\xca\x96\x8dD\xa7\xd3s\xe5\xca\x156n\xdc\xc8\xa7\x9f~\xca\xde\xbd{1\x99L\xae_\xcd\xc1]w\xddE\xcf^\xbd\x188`\x00\xd1\xe5\xa2\xd1\xc9:\xd7,\x1e<\x94zJ\xb9\x02(\xc5\x12\xc8C\x89 \x84 --\x8dC\x87\x0e1w\xee\\\x96-[\xc6\x95+W\\\xb3\xe5\xa0j\xd5\xaa\xf4\xea\xdd\x9b\xc7\x1e}\x94:u\xea`0\x18\x9c\xd2\xadV+\'O\x9ed\xd9\xb2e,_\xbe\x8c\x7f\xfe\xd9\xe9\x94\xee\x0e\xa3\xd1H\x9f>}x\xea\xa9\xa7h\xd2\xa4\t\xc1A\xc1\xc8\xb2\\DU\xd6\xd3\x16<\x14>\xa5\\\x01x\xf0\x90\x8d\xd9b\xe6\xc4\xf1\x13\xacY\xb3\x9a\x8f>\x9aH||\xbck\x16\xb7<\xf8\xe0\x83\x0c\x1e<\x98V\xadZ\x11\x1c\x1c\x92\xab\'\x8f\x10\x82\x8c\xf4t\xf6\xec\xdb\xcb/?\xff\xc2\x94)S\\\xb3\xe4\xca\xcb/\xbf\xcc\xa0A\x83\xa8\xdf\xa0>~\xbe\xbexb\xaay\xb8\x19\xf0(\x00\x0f\xa5\x1e!\x04\t\t\t\xacX\xb1\x82\xe5\xcb\x97\xb3f\xcdj\xacV\xc55\x9b\x13:\x9d\x8e\x86w6\xe4\xb9\xe7\x9e\xa7[\xb7\xaeD\x96\x89D\xa7\xd7\x83\xd0\xe6m\xdd \x14\x01\x92z\xbd\xabW\xaf\xb2c\xc7\x0ef\xce\x9c\xc9\xd6m[IMIu\xcd\x9e\x83\x96-[\xd2\xabW/\x1ez\xe8!\xca\x97/\xaf\x8e\x06<x(\xc5x\x14\x80\x87R\x89\xd0\xfc\xf9\xd3\xd33\xf8\xef\xe0\xbfL\x9f1\x9dU+W\x91\x94\x94\xe4\x9a5\x07\xf5\xea\xd7\xa7g\x8f\x1e\xdc\x7f\xff\xfd\xd4\xabW\x0f\x9dN\xce\xd5|\x92\x9baE \x10\x8a .\xee\x02\xab~]\xc9\xe2\xc5\x8b\xd9\xb8a\xa3k\xb6\x1c\x18\x8dF:t\xe8\xc0\x90!\xcf\xd2\xb6];BCB\x91e\xd9\x13\xff\xe7f!\xb7\nq\x8b\xe2Q\x00\x1eJ%f\xb3\x89CG\x8e\xb0a\xddz>\xf8\xe0C\x12\x13\xaf\xbafq\xcb\x90!C\x180`\x00\xcd\x9a5\xc3\xdf_]\xc1{c\xedYR\xef\xe5\xf0!V\xaf^\xc3\x84\x8f&\x90\x92\x94\xec\x9a\xc9-\xc3\x86\r\xa3_\xbf~4n\xd2\x04??_n\xf4N<x(l<\n\xc0C\xa9B\x08\xc1\xe5\xcb\x97Y\xf1\xcb\n\xfe\xb7\xf8\x7f\xac[\xbf\x9e\xfc\x0c\xea\x06\x83\x81\x9a5k\xf2\xfa\x88\xd7\xe9\xde\xad;AAA\xe8t\xbaB5\xc1\xd8&\x9fw\xef\xd9\xcd\xec\xd9\xb3Y\xbbf-\x89\x89\x89\xae\xd9rP\xbfA\x03\x06\r\x1c\xc8}\xf7\xf5\xa7Z\xb5\xea\xe8t\x1eo!\x0f\xa5\x07\x8f\x02\xf0Ph\\\xdf\xe8Y \x84M\xc0\xa6\xf3\xcf\xae\x7f\xf8x\xca\x14\xb6l\xdeBZZ\x9ak\xe6\x1c4n\xd2\x98\xde\xbdzs\xff\xfd\x83\xa8R\xa5j\x0e\xef\x9e\xeb&\x97\x87\x11\x8aB\xc2\x95+\xac\\\xb9\x92\x9f~\xfa\x89\x15+V\xb8f\xc9\x81\xc1`\xa0A\x83\x06\x0c\x7f\xf9e\xbat\xe9BhH\x08\xb2N\x97\xbdPMr{)\x0f\x1e\x8a\x1c\x8f\x02\xf0P\xe2\x98L&\x0e\x1e<\xc8\xea5\xab\x995k\x16q\xb1q\xaeY\xdc2|\xf80\xfa\xf5\xefO\xa3;\x1b\xe3\xef\xef\xe7\x9a\\$\xa8\xb6|\x19\xa1X9r\xf4(\xeb\xd7\xadg\xf6\x17\xb39t\xe8\x90kV\xb7<\xfd\xf4\xd3\x0c\x1c8\x80f\xcdZ\x10\x18\x18\x00xv\xe1\xf2Prx\x14\x80\x87\x12\xc3\xaaX\x89\x8b\x8dc\xf1\x92%\xac\xf8\xe5\x176m\xda\xe4\x9a%\x07\xde\xde\xde4m\xd6\x94W_y\x95V\xad[\x11\x16\x1af7\xab\x14\xcfD\xab@h\xd6|EQ\xc8\xcc\xcc\xe4\xd8\xb1c|\xf7\xddw|\xfd\xcd7\xa4$\'\xe7\xbb\xafn\x9d:\xb5\xe9\xd3\xa7/\x0f<\xf0\x00\xb5j\xa9A\xe6\x8a\x8a\\\x062\x1e<\x80G\x01dS\xd4\xc2\xa3\xa8\xcf\x7f\xf3 \x10\n$&\'\xf2\xf7_\x7f3}\xc6\x0c~\xdf\xb0\x81\xcc\xccL\xd7\x8c9h\xd3\xa6\r\xfd\xfb\xf7\xa7_\xff\xfeT(j7\xcbk\x90\x9cB\x08RSSY\xb3f\r\xcb\x97/g\xc5\x8a\x15\xa4\xa7\xa7\xbbfsB\xaf\xd7S\xa3Fu\x86\r\x1bN\xcf\x9e=)S\xa6\x0c:\xbd\xbe\xa0\x97\xf4\xe0\xa1P\xf0(\x00\x0f\xc5\x86\x04dfe\xb1\x7f\xdf~\x96-_\xc6\xa4I\x93\\\xb3\xe4\xca\xc87\xdf\xe4\xbe\xfe\xfd\xa9S\xa7.\xde\xde\xde\xb9\xfa\xf2\x97$\x8a\xa2p\xee\xdc96n\xdc\xc8\xfc\xf9\xf3\xf9\xe3\x8f?\\\xb3\xb8\xa5\xff}\xfdy\xe6\xe9gh\xd2\xa4\taaa\xae\xc9\x1e<\x14\x19\x1e\x05\xe0\xa1XP\x14\x85\xf3\xe7\xcf\xb3t\xe9R\x96-[\xc6\x9f\x7f\xfe\xe9\x9a%\x07\xfe\xfe\xfet\xe8p\x17/\r}\x89\xa6M\x9b\x12\x12\x1c\x8c$I\x88<\x16s\x95\x14\xc2\xf6\x8f\x04\x99\x19\x19\x9c9}\x86o\xe7\xcd\xe3\x87\x1f\xe6\x13\x17\x17\x97\xafY\xa8V\xadZt\xea\xd4\x99\x87\x1e\x1a\xcc\x9d\r\x1b\xe2\xe5\xe5U\xfa\x1e\xd2\xc3-\x87G\x01\xdc\x82\x94\xb4\xb9I\x15\x86\xea=(\x8a\xc2\xd5\xabW\xd9\xbcy3\xb3g\xcff\xdb\xb6mddd\xb8~%\x07\xed\xdb\xb7\xe7\x81\x07\x1e\xa0W\xaf\x9eDGG#Iy\x98{\xae\xc1\\S\x9cdff\xb2e\xf3\x16\x16.Z\xc8\xda\xb5k\x89\x8b\xcb{r[\x96e\xaaT\xa9\xc23\xcf<\xcb\xc0\x81\x03\x88\x8a\x8e\xc2\xa07\xa8\x8fV\x82\xef\xb3$(\xa5\xaf\xb4\x04)\x9a\x12\xf1(\x00\x0fEFjj\x1a\xbbw\xef\xe2\xa7\xe5?1u\xdaT\xd7d\xb7x\x1b\xbdx\xf3\xadQ\xf4\xec\xd9\x93z\xf5\xeaa4\x1a\xf3\xed=\x97f\x14E\xe1\xc2\x85\x0bl\xdb\xb6\x95\x05\x0b~,P\xd4R\x80\xbe}\xfa\xf2\xc0\xe0\x07i\xd3\xa6-\xd1QeKT\xa1{\xb8u\xf1(\x80"\xc4Yg\x17\x8d\x06/\x8dX\xadVN\x9c8\xc1\xe2\xc5\x8bY\xbcxq\x816\\\t\n\n\xa2O\xef\xde<\xfd\xcc3\xd4o\xd0\x80\x00\xff\x00d\xf9V)/A\x96\xc9\xc4\xc5\x0b\x17Y\xbcx1\xb3g\xcf\xe6\xf4\xe9\xd3(J\xde\xf1\x8c\xa2\xa2\xa2\xe8\xda\xad\x1b\x0f>\xf0\x00\xed\xda\xb7\xc7\xe8\xe5\xe5\x9a\xc5\xce\xedS\xbb<\x14&\x1e\x05\xe0\xe1\x86\xb1\x99\x9c\x14E!!!\x81u\xeb\xd73m\xeaT\xf6\xee\xdd\x8b\xd9lv\xcd\x9e\x83\x9e={\xd2\xa7O\x1f\xfa\xde\xdb\x97\xb0\xd0\xb0\xa2\xf5\xee)I\x04d\x99\xb2\xd8\xf1\xd7\x0e\x16.\\\xc8\xfau\xeb9y\xf2\xa4k.\'$I\xa2|\xf9\xf2<\xf7\xdcs\xdcw\xdf}T\xacX\t//\xc3\xcd;"\xf0h\xaaR\x85G\x01x(\x14RRS\xd9\xb7w/\x0b\x17.d\xd6\xacY\xae\xc9n\xa9\\\xb92C\x86\x0c\xa1{\xf7\xee\xd4\xae]\xfb\xb6\x08\x93 \x84@\x08A\xfc\xe5xv\xec\xd8\xc1\xdc\xb9s\xf9u\xd5\xaf\xae\xd9\xdcr\xcf=\xf7\xf0\xd8\xa3\x8f\xd2\xbeC\x07\xca\x95+w\xeb*\xca\xdb\x96\xe2\xd7\x8e\x1e\x05\xe0\xe1\x9aq\xac\xa6&\x93\x89#G\x8e\xb2d\xc9b~\xfd\xf5Wv\xef\xde\xed\x92;\'\xbe\xbe\xbe<\xf6\xf8\xe3<\xfe\xd8c\xd4\xacQ\x13??_54B>\x14\x7f\xf3(|l\x1eLB\x80\xc5l\xe6b\xfcE\x96/[\xce\xb4i\xd38s\xe6\x0cV\xab\xd5\xf5+N\x04\x05\x06ro\xbf~<\xf8\xe0\x83\xb4m\xdb\x16\x1f\x1f\x1f\xd7,\x1e<\x14\x18\x8f\x02\xf0\xe0@\xc1D\xac\x00\xacf\x0b\x17/]d\xf5\xea\xd5|\xf6\xd9g\x1c?v\x1c\xb3%os\x8fN\xa7\xa3W\xaf^\x0c\x180\x80^={\x11\x10\x18p\xf3\x9a2\n\x91\xcc\xacL\xf6\xee\xdd\xcb\xf2e\xcbY\xf5\xeb\xaf\x1c\xfa\xef?\xd7,NH\x92Ddd$O?\xfd4\x0f>\xf8 \x95*\xdd\x81\xb7\xb7\xd1^\x96Bhk\x95=E\xeb!\x07\xcem\xdc\xa3\x00<\\3\x97\xe2\xe3\xf9\xfb\xaf\xbf\xf8\xea\xab\xb9\xfc\xf2K\xfe\xc1\xd0\x00\x9a5k\xcec\x8f=\xca=w\xdfC\xd5jUo\x0bs\xcf\xb5 \x10\\\xbd\x9a\xc8\xce];\xf9\xfe\xbb\xefX\xf0\xc3\x02\xd7,ni\xdd\xba5O>\xf9$\x1d;v\xa4R\xa5JZ\xe3.Y7`\x0f7\x0f7\x91\x02(X\xef\xb44R\xd2~\xf9\x85\x82\x00\x93\xc5\xcc\xbe\xbd{\xf9\xfe\xfb\xefY\xbdzu\xbe\x13\x98\x00e\xca\x94\xa1o\xdf\xbe<\xff\xfc\xf3T\xadZ\x15??\x7f\xe0\xb6sk/\x10B\x80\xa2X\xb9t\xe9"+W\xad\xe2\xf3\xcf?\xe7\xf0\xa1\xc3X,\x16\xd7\xacN\x18\xf4\x06\xfa\xdfw\x1f\x83\x06\r\xa4c\xc7\xbb\x08\x08\x08\xf4\xcc\x0f\xdc\x14\x94\xbcL\xbb\x89\x14\x80\x87"\xc7M}\x14\x02\xcc\xa6,.\xc5\xc7\xf3\xf3/?3s\xc6L\x8e\x1e=\x9a\xaf\x0bchh(\x9d;w\xe6\xfe\xfb\xef\xe7\x9eN\x9d\xf0\xf3\xf3\x83R\xb8\x82\xb7\xb4b6\x9b9t\xf8\x10K\x16/a\xc9\x92%\x1c9r\xc45\x8b\x13\x92$\x11Q&\x82A\x03\x07\xf1\xd4\xd3OS\xb9re\xfc|\xfd\x90$\xe9\x16.s7\x15\xd6\xc35\xe1Q\x00\x1erE\x08A|\xfce\xfe\xfcs\x1b\xdf}\xf7=?\xff\xfc\x93k\x16\xb7\xd4\xaf_\x9f\x11#F\xd0\xa6M[*T(\x7f\x93\x99{T\xa1b\xb5ZU\xe1)\xcb%&b\x84\x10$%%\xb1{\xf7n\x16.Z\xc8Ws\xbfr\xcd\xe2\x96\xf2\x15\xca1\xea\xad\xd1t\xe9\xd2\x95\n\x15\xca\xa3\xd7\xeb]\xb3x\xf0\x00\x1e\x05\xe0!7\xccf3\xbbw\xeff\xde\xbcy\xacY\xbb\x963\xa7O\xbbfqB\x92$\xa2\xa2\xa2x\xfc\xf1\xc7y\xf4\x91G\xa8X\xa9\x12^Fc\x89\t\xcf\xebA\x08Arr\x12\x7fl\xda\xc4\xde={\xa9^\xa3:]\xbbt!,,\xdc5k\xf1!@\x11\n\x97.]b\xe3\xc6\x8d\xcc\xf9r\x0e\xdb\xff\xdc\x9e\xafY($8\x84N\x9d:q\xdf}\xf7\xd1\xb5[W\x02\x02<\x13\xee\x1er\xe2Q\x00\x1e@3\xf5\x00X\xac\x16bbbX\xbel\x19s\xe6\xcc\xe1\xd8\xb1c\xf9\x86b([\xb6,=z\xf6\xe0\xde\xbe\xf7\xd2\xae}{\x02\xfc\xfdoJa\x93\x92\x92\xc2\xacY\xb3\xf8\xe4\xd3O\xb8\x92p\x05Y\x96y\xf9\xe5\x97\x195j\x14\x81\x81\x81%jrP\x84\x82\xb0\nN\x9c<\xc1\xf7\xf3\xe7\xb3~\xdd:\xfe\xfe\xfbo\xd7l9\x08\x0b\x0b\xa3G\xcf\x1e\xbc\xf8\xc2\x0b\xd4\xacU\x0b\x7f?\x7f$Y\xf2\xecO\xec\x01<\n\xc0\x03\x9aXC\x08\xe2\xe3\xe3\xd9\xbam\x1b\xf3\xe7\xcf\xe7\xe7\x9f\nf\xeei\xd3\xa6\rC\x9e\x7f\x8e{:\xdeM\xd9\xb2e\xed\xca\xe2fT\x00111T\xaaT\xc9I\xe1\x95)S\x86\xa5K\x97\xd1\xbauk\x84\xa2 \x95dx\n\xa1\xce\x9e\xa7\xa7\xab{\x13/Z\xb8\x88\x993g\xba\xe6r\xcb\x9d\r\x1b\xf2\xd8\x13O\xd0\xa9S\'\xaaW\xabV\xa4\x9b\xd0x\xb8y\xf0(\x00m\xe8\x7f3\n\xac\xc2"++\x8b]\xbbv\xb1p\xe1B\x16-\\\xc4\xe5\x84\xcb\xaeY\x9c\x90e\x99r\xe5\xa2y\xe6\x99g\xb9\xff\x81\xfb\x89\x8e.\x87\xb7\xb77\xb2\xdd\x0f\x1d$\xa9\xe4z\xcb\xd7\xcb\x993g\xa8Z\xb5\xaa\x93\x02\x88\x88\x88`\xfe\x0f\xf3\xe9|O\'UQ\x96\x10\x8e\xa5)\x00\xab\xc5Bbb"\xdbwlg\xca\xe4)\xfc\xf3\xcf?dee\xb9|\xcb\x19\x9dNG\x8f\xee\xdd\x190h \xbdz\xf6"((\xe8\xb6\xae\xf7\x1e<\n\xa0\xf4P\x0c\xf2\xd2\xf6\xa2%\xedb&\xb3\x89\xd8\x98X\xfe\xf7?5@\xd9\xb9sg\xf3\xf5\xee\t\t\ta\xd0\xa0A\x0c\x1e\xfc M\x9b6\xc3\xdb\xdb\xdb5\xcbM\x89\x10\x82\x98\x98\x18\xaaU\xab\x8a\xd9\x9cm_\x0f\x0f\x0f\xe7\xfb\xf9\xdf\xd3\xa5s\x97"zE\x05=\xab\xfb|VE\xe1\xec\xd93,Y\xbc\x98\x9f\x7f\xfe\x85\xed\xdb\xb7\xbbf\xc9\x81\x9f\x9f\x1f\xf7\xdd\xd7\x9f\x17_|\x89Z\xb5j\xe1\xe3\xe3\x8bNgs\x1bu\x7f\x9d\xdc\x8f{\xb8\x99\xf1(\x80\xdb\x08E\x08$\t$\x01\xf1\xf1\xf1l\xd9\xba\x95\x993g\xf2\xfb\xef\xbf\xbbfuK\xf7\xee\xddx\xf2\xc9\xa7h\xd9\xb2%\xd1\xe5\xcaeO\x1c\xdc"\xc4\xc4\xc4P\xa5J\x15\xa7\t\xd6\xf0\xf0p\xbe\xfb\xee;\xbat\xee\xa2\xf6\x96K\xa9\x0c\xcc\xc8\xcc`\xff\xbe\xfd\xac\\\xb1\x82\x993g\x90\x98\x94\xec\x9a%\x07\xd5\xabW\xe7\xd9!\xcf\xd2\xad[7j\xd6\xa8\x85\xac)\x81R\xfa\x88\x1e\x8a\x00\x8f\x02\xb8m\x10\x08\xa1nR\xb2\xfd\xcf?Y\xbcd\tk\xd7\xae\xe5\xcc\x993\xae\x19\x9d\x90$\x89jU\xab\xf1\xcc\xb3\xcf\xd0\xaf_?\xcaW(\x8f\x97!\xf7\xb0\xc473\xb1\xb1\xb1T\xa9R\xc5)\x82\xa9M\x01t\xed\xda\xd5)oiDQ\x14\x92\x93\x93\xd8\xb5s\x17\xe3\'L`\xc7_\x7f\x91Y\x80\xcdwz\xf5\xecE\xbf\xfe\xf7\xd2\xb3W\xaf[;\x1a\xab\x87\x1c\x14\xbd\x02\xf0\x8c\x1c\x8b\x87\\\xcaY(\x02$\xd5\xce\x1f\x1b\x1b\xcb\xc2E\x0b\x99\xfb\xe5\\\xce\x9d;\x97\xaf\xb9\xa7b\xc5\x8at\xeb\xd6\x8d\xc7\x1e{\x8c&M\x9a\xdc\xf2\x13\x87\xb1\xb1\xb1T\xadZ\x15\x93\xc9d?\x16\x16\x16\xc6\xbcy\xf3\xe8\xde\xbd\x9b\xfb\x02.\x85(\x8aBll,\xcb\x96/\xe7\xe7\xe5\xcb\xf9c\xd3&\xd7,9\x08\n\x0e\xe6\xde\xbe}y\xf6\xd9g\xa9W\xaf>~\xfe~vS\xa1\x8d\xdb}\xae\xac\xf8\xc8\xa51\x17\x01E\xaf\x00<\x94(\x8a"\x88\x8d\x8da\xcb\x96\xcd\xcc\x9c9\xab@{\xf1\x02\xf4\xed\xdb\x97\x87\x1e~\x98\xb6m\xdbR&"\xe2\xb6h\xf8\xea\x1c@5\'\x05\x10\x1e\x1e\xcew\xf3\xe6\xd1\xb5[W\xadY\xde<\xe5\x90\x9e\x9e\xc1\xe1\xc3\x87X\xfd\xeb\xaf\xbc3v\xack\xb2[\xc2\xc2\xc2\x185z4=z\xf4\xa0J\xe5\xca\x9eEd\xb78\x1e\x05p\xcb"HO\xcb`\xeb\xb6\xad,X\xb0\x80\xdf~\xfb\x8d\x8b\x17/\xbafrB\xa7\xd3Q\xb3FM\x9e\x7f\xf1\x05z\xf7\xeeMdd\x19uO\xda\xdb@\xf8\x93\xcb\x08 \\\x1b\x01t\xed\xd6\xad\xd4\x97\x83k\xbfQ\xdd{@!%%\x95]\xbbw3\xe3\xf3\xcf\xd9\xb0a\x03)))\x0e\xb9r\xe2\xed\xedM\xe7\xce\x9d\x19p\xdf\x00\xbau\xefFXXX\xa9\x7fv\x0f\xd7\x87G\x01\xdcT\xb86q\xf7dffr\xee\xdc9\xbe\xfe\xfak\xe6\xcf\x9fO\\\\\\\xbe\x8b\xb9j\xd7\xaeM\xd7\xae]y\xe8\xa1\xc14h\xd0\x00Y\xa7\xb7\xbbu\xe6E\xbew\x94o\x86\xd2CLL\x0c5\xabW\'\xc3\xc1\x9d2<<\x9co\xbf\xfd\x96n7\x81\x02\xc8\x0b!\x04\x17.\\\xe0\x97_~a\xc1\x8f\x0b\xd8\xbae\xabk\x96\x1c\x84\x04\x87\xd0\xb5[W\x9e\x7f\xfey\xeelx\'~\x01~7\xd5\x08\xc8C\xfex\x14\xc0-\x84UQ8{\xe6\x0c\x9b7o\xe6\xfb\xef\xbf/\xb0w\xcf\xa0\xfb\xef\xe7\xf1\xc7\x1f\xa7y\xb3f\x84\x84\x84\xb8&\xdf6\xb8\x1d\x01\x84\x87\xab#\x80\xae]oj\x05`#3+\x8b\xff\xfe\xfd\x97_~Y\xc1\x87\x1f~\x90\xef<\x10\xda(h\xf4\xdbos\xd7\xdd\x1d\xa9U\xa3&^y\xecM\xec\xe1\xe6\xc2\xa3\x00n\x01\x84\x10\xa4\xa5\xa5\xb1\xf1\xf7\x8d,X\xf0#\xabV\xac =\x1f\xef\x0f\xbd^O\xf5\xea\xd5y\xe5\x95\x97\xe9\xd5\xb3\x17!aax\x19\xf4yv\xd7o\xf5I\xc0\xbc\x14@\xb7n\xdd\x9c\xf2\xde\x8c\xd8G\x81\x02RRS\xd8\xb3g\x0f\xb3g\xcff\xe5\xca\x95\xa4\xa7\xa7\xbbfwB\x92$:u\xea\xc4\x80\x01\x03\xe8\xdd\xbb\x0fe\xca\xdc\x1e\xf3B\xb7:\x1e\x05p\x13!\xb4\xfd\x04%\xd4F\xac\x08\x85\xcc\x8c\x0c\x8e\x1e;\xc6\xf7\xf3\xbfg\xfe\xf7\xf3\xd5U\xbc\xf9\xbc\xd1F\x8d\x1a\xd1\xb3G\x0f\x06\x0e\x1cD\xad\xda\xb5\x1c&\xfa\x9c\xed5VE!+3\xd3\xee\x17\x9f\x9e\x9e\x8e^\xaf\xc3\xcb\xcb\x88,\xc9\x18\xbd\x8d\xe8\xf5z$IB\xa0\xedBu\x13\x13\x13\x1bK5w\n\xe0\xbbyt\xed\xdc\x05I\x92\xf3\xd2\x8f7\x1dB\x08b\xe3bY\xbel9k\xd7\xae\xe1\xd7_W\xbbf\xc9\x81\x9f\x9f\x1f\xed\xda\xb5c\xe4\xc8\x914l\xd8\x90\x00\x7f\x7fd\x9d\xee\x96\xef\x1c\xdc\xaax\x14\xc0M\x8a\xa2\x08N\x9d:\xc9\x86\r\x1b\x98={6\xfb\xf6\xeds\xcd\xe2\x96\xc7\x1e{\x8cG\x1f}\x94\xa6M\x9a\x12\x10\x18\xe0vn@\xb1*\xc4]\xb8\xc0\xe1\xc3\x878y\xf2$g\xcf\x9e\xc1\xaa(\x1c>t\x08??\x7f*T\xa8@`@ 5k\xd5$<"\x9c\xea\xd5\xaa\x13\x1e\x1e\x81^\x7f3\x85}\xce\x89\xbb\x11@Xx8\xf3\xe6}K\xb7\xae7\xf7\x1c@^\x98\xcdf\x8e\x1e;\xc2\xaf\xabV3g\xce\x9c\x02m\xf4S\xae|y\x86<\xfb,}\xfa\xf4\xa1v\xed\xda\xe8\xf5\xba[\xa0\x0bp\xfb\xe1Q\x007\x19\x8aPHJLd\xfd\xfa\r,Y\xbc\x98\xa5\xcb\x96\xb9\x15\xe2\x8e\x18\x0c\x06\x1a4h\xc0+\xaf\xbc\xc2=\xf7\xdcChh\x98[a-\xb4\x80p\xbf\xfc\xf2\x0b[\xb6l\xe1\x8f?\xfe\xe0\xfc\xf9\xf3\xae\xd9\xec\x18\x0c\x06\xaaV\xadF\x8b\x16\xcd\xe9\xdc\xb93]\xbat!,,\xcc5\xdbM\x83;\x05\xe08\tl\x1f}\x15\x1bE<\x83.P\x97\x85#!\x84 ==\x8d\x7f\xff;\xc4\x8c\x19\x9f\xf3\xd3\xf2\x9fHMMu\xfdF\x0e\xbav\xedJ\xdf\xbe}\xb9\xb7\xdf\xbdD\x84Gx\x16\x91\xdddx\x14\xc0\xf5P\xc4\xed\xd2\x15\xdb\xf0:==\x9d#G\x8e\xf0\xe5W_\xb1t\xf1\x12\x12\x12.\xe7+\xfc\x1b4h\xc0\xc0\x81\x03\xe9\xdf\xbf?\xd5\xabWw\xbf9\x8b\x10\x98-\x16\x0e\xec\xdb\xcf\xb8\xf7\xdfc\xc3\x86\rd\xe43\x87\xe0JHH(\xfd\xfb\xf7\xe3\xe5\x97_V\xa3M\xde\x84\x13\x85\xe7\xcf\x9f\xa7F\x8d\x1aNA\xd5\xc2\xc2\xc2\xf8\xe6\xdbo\xe8\xd1\xbd\xc7-;\x02pDh\x9b\xd0\xacZ\xb5\x8a%K\x96\xb0r\xd5*\x14\xab\xd55\x9b\x13\xfe\xfe\xfe\xf4\xe8\xd9\x83\xe7\x9f{\x9e\x86\r\x1b\x12\x18\x18\xa8\xedBv\xeb\x97\xd7\xcd\x8eG\x01\xdc\x04\x98\xcdfN\x9f>\xcdo\xeb\xd6\xf1\xd1\xc4\x8f\x88=\x1f\xe3\x9a\xc5-/\xbc\xf8"\x03\x07\x0c\xa0I\x93&\xf8\xf8\xfa\xe6\xea\xd6i6\x9b\xd9\xbau+\x1f\xbc\xff\x01\x7fl\xfa\xc35\xf9\x9a\xe8pW\x07\xde\x19\xf3\x0e\xed\xdbw@.\xc9\xd0\xc9\xd7\x81\xbb\x11@XX\x18_\x7f\xf35={\xf4P\xe7\x00n\x13,\x16\x0b\'N\x9c`\xdd\xfau\xcc\x9c9\x8b#\x87\x0f\xbbf\xc9\x81\x8f\x8f\x0fo\xbe\xf9&\xbdz\xf5\xa2V\xadZj\xa0\xc0b\xee,\x95>Jw\x01x\x14\x80Fi|MB\x08\xae\\\xb9\xc2\xaaU+Y\xbe\xfc\'6l\xd8@ZZ\x9ak6\'\x0c\x06\x03\xad\xdb\xb4f\xc8\x90!t\xbc\xab#\xe1\xe1a\xc8\xb2\x9b^\xbf\x86\xa2(\xec\xf8\xeb/F\xbf\xf5\x16\x9b\xb7lqM\xb6#I\x12\xb2,#\x84\xc8\xd7u\xb0C\x87\x0e\x8c{\xef=\xda\xb4i\x8d\x84t\xd3\xf4\x9csS\x00\xb7\xd3\x08\x00\xd4\xc6 4O\x82\xf4\xf4t\xfe\xfb\xef?f\xce\x9c\xc9\xaaU\xabHHHp\xcd\x9d\x83\x0e\x1d\xeeb\xd0\xa0\x81\xf4\xee\xdd\x9b\xa8\xe8hd\n?\x88\x9eMh\x15\xf2io;<\n\xa0\x94!\x14\x81\x82 #=\x9d\x03\x07\x0e0k\xd6,V\xadZEbb\xa2k\xd6\x1c\xb4l\xd5\x92\xce\x9d;\xf3\xe8#\x8fR\xa9REt\xba\xbc\x97\xf1\x0b!\xf8\xef\xbf\xff\x182dH\xaea\x84%\xa0s\xd7\xaeT\xafV\x8d\x88\x88\x08\xacV+\'N\x9c\xe0\xf0\xa1C\xec\xde\xb3\xc75;h\xfb\x05<\xf6\xf8c\x8c\x1d;\x96\xe8\xa8h\xf7f\xa7R\xc8\xb9s\xe7\xa8V\xad\x9aS4\xd0\xb0\xb00\xbe\xfe\xfa+z\xf4\xe8\x91\xa7"\xbd\xd9\xc9\xcb\x8bG\x08AJJ\n\xabW\xaff\xc9\x92%\xfc\xb6\xee7RS\xf2\x9e\x1f\x08\x08\x08\xe0\xee\xbb\xef\xe6\xa5\x97^\xa2I\xd3\xa6\xaa\xb7\x90,\x97\xca\x8e\xd6\xed\x8cG\x01\x942\xb2\xb2\xb28v\xfc8\xeb\xd7\xadc\xe2\xc4\x89\\\xbat\xc95\x8b[^}\xf5U\xfa\xf5\xebG\xc3;\x1b\xe2\xe7\xe7\xa7\xf5\xe2\xf2nlIIIL\x9e4\x89I\x93&cUr\xday\xabU\xab\xce\xb0aC\xb9\xfb\xee\xbb)W\xae\x1cAA\xc1(\x8a\x95\xb8\xb8\x0b\x9c<q\x82\xe5?\xfd\xc4g\x9f}\xea\xfa5\x00\x02\x83\x82\xf8|\xfat\xfa\xdf\xd7\x1f\x1fo\x1f\xd7\xe4R\x89;7PU\x01|M\x8f\x9e=4\x13Z^%zkc\xb1Z9}\xe64kV\xaff\xd2\xc4\x89\xc4\xc4\xc4\xbafq\xcb;\xef\xbcC\xcf^=\xa9W\xb7.\xde%U\x17$\x89[-|ya\xa0\x1b;v\xec\xbb\xae\x07og\xf2\x13\x9aE\x85P\x04\t\t\t,_\xbe\x9ci\x9fMe\xd6\xec\xd9\xf9za\x18\x8dFZ\xb7n\xcd\xb4i\xd3\x18<\xf8!\xaaV\xad\x8a\xd1\xcb\xcb\xfe\x04y=\x87\x10\x82\x9d;w\xf2\xf1\'\x9f\xb8U2\x11\x11\x11\x8c\x1b7\x8eG\x1ey\x98\xe8\xe8h|||\x90\xb4\xde}`@ \xe5+\x94\xa7y\xf3\xe6\xf8\xf8\xf8\xb0\xc9M\xb4\xc9\xac\xac,\x12\xae$\xd0\xa3{\x0f\x02\x02\x02\\\x93K%\xc9\xc9\xc9L\x9b:\xd5\xc9\xc4\xe5\xeb\xebK\xbf~\xf7R\xa3F\xcd|J\xf4\xd6G\x92eB\x83C\xa8W\xa7\x1e=z\xf6\xc4\xa0\xd7s\xe6\xcc\x99|\x17\x91m\xda\xb4\x89\xa3G\x8e\x92\x96\x9eF\xc5J\x95\xf0\xf3\xf5\xcbu\xb4Q\xd8(\x8aBFF\x06\xe7\xcf\x9f\xe7\xe2\xa5\x8b\x18\xbd\xbd1hkW<\x94\xea\x11@I\x89\xe2\xe2\xc16\xe4V\x84BZj\x1a\x07\x0f\xfe\xcb\x84\x8f&\xb0\xe9\x8f?\xf2\x15\xfc\x00m\xdb\xb6\xa5O\x9f>\x0c\x180\x90\n\xe5\xcb\xdb7\xf3((\xa9\xa9)\xcc\x9a5\x9b7\xdf|\xd35\t\x80\xd1\xa3G\xf3\xf6\xe8\xb7\xf1\xf6\xc9{\xc7\xaf\x84\x84\x04>\xfah\x02S\xa6|\xec\x9a\x04\xc0\xfc\xf9\xf3\xb9\xff\xfe\xfb\xb3\xdd\x03K\xf1k\xcd-\x1a\xe8\xb7\xf3\xe6\xd1M\xdb\x0f\xc0\xbd\xe0(\xc5\x0fUD(\x8aBjj*\x1b7nd\xc1\x82\x05,_\xfe\x13\x8a\x9bQ\xa4#\xde>\xde\xb4h\xd1\x92\xd7^{\x8d6m\xda\x10\xe0\xef\xef`\x1e,\xcc2\x14\x08E\x90\x98\x94\xc4\x8e\x1d;\xf8\xf6\xdbo8|\xf8\x08f\x8b\x85*\x95+\xf3\xf0#\x8f\xd0\xa9S\'BCB\x90\xa4\xec9*\xf5\x0e\n\xf3>J?\xa5x\x04p\x8b\xbf\x04I\xc2b1s\xe8\xd0a~\xf8\xe1\x07\xdez\xebMv\xee\xdc\xe9$|\xdc\x11\x12\x1a\xca\xab\xaf\xbc\xc2\x0b/\xbe@\x9f>}\t\x0f\x0f\xcfE(\xe5\xcd\xe5\xcb\t|\xf8\xe1\x87\x9c;w\xce5\x89\xbb\xef\xbe\x9b1c\xc6P&\xb2\x8ckR\x0e\xbc\xbd\xbd\t\x0e\t\xe1\xf0\xe1\xc3\x9c={\xd65\x19??_z\xf6\xec\x95\xdd\xd0\xaf\xfdV\x8b\tAJr\n\xd3?\x9f\x8e\xd5\xc1\xed\xd1\xd7\xd7\x97{\xef\xbd\x97\xaa\xd5\xaa\xe5\xeaEU\x8a\x1f\xaa\xc8\x90$\t\xa3\xd1H\xcd\x9a5h\xdc\xb8\x115k\xd6\xe2\xcc\x993\xc4\xc7\xc7\xbbf\xb5c\xb1X8s\xfa4\xeb\xd6\xad\xc3b\xb1\xe0\xe7\xefOxX\x18z}\xe1\xee3!\x84j\xce\x9b5k\x16\xcf=\xf7\x1c\xff\xfdw\x88K\x97.\x91p\xf92\xc7\x8f\x1fg\xd9\xd2\xa5\xa4\xa7\xa7Q\xb5JU\xa7H\xa7\x8e\xff\xde.\\[\xb7\xd1C\xa1\xa0(\n11\xe7\x995k\x16#\xdfx\x83\xd1\xa3G\xbb\x15\xc4\x8e\xf8\xf8\xf8p\xd7]w\xf1\xddw\xf3x\xfd\xf5\xd7i\xde\xb49F//\x84\xa2h}\x97\x82#\x84\xe0\xd8\xb1c\xb9\xee\r\xd0\xbbw\x1f\xa2\xa3\xa3]\x0f\xbbE\x96e\x1a4h@\x9b6m\xdc.\x02\xda\xbf\xff\x00\x17/^\xbc.%U\xbc\xa8\x9e*V\xc5\xb9,mf\xe3\xd2~\xf7%\x85$\xc9T\xadZ\x8dG\x1f{\x84o\xbe\xfd\x86\x97\x86\x0e%$8\xef\x80\x82\x97/_\xe6\x83\xf7\xdf\xe7\x8d7F0}\xfatN\x9f9\x9d\xafg\xd9\xb5\x90\x96\x96\xc6\x92%K\x99:u\xaak\x92\x9d\x993f2o\xde<\x12\x93\x92\\\x93r\xe1\xda\xda\xd8\xcdB\xce\x16\xeb\xa1\xc8\x10\x8aBrr2[\xb6l\xe1\xf5\xd7^g\xd4[\xa3X\xbbv\xadk\xb6\x1c\xdcsO\'&N\x9c\xc8\x82\x05?\xd0\xadk7\x82\x82\x82\x90u\xb2\xc3\xf0\xf5\xda\xc5SL\xac\xfb\t\xbc\x8a\x15+\xd2\xb0a\x03||\n>Y\xe7\xe3\xedM\xb3fM\t\n\nrM\xe2\xe2\xc5\x8b\x9c={\x16!\xb4XF\xa5\x15\xed\xd6r.]\xc8\xeb\x9e\xf3J\xbb}\x90$\t?_?\x1a7j\xcc\x07\xef\xbf\xcf\xec/\xbe`\xd0\xa0A\xf9N\xf8n\xdd\xb2\x95w\xdey\x87W\x86\r\xe7\xb7\xdf~\xe3\xca\x95+(\x8ar\xc3\xf5\xe4\xf2\xe5\xcb|\xf6\xd9\xa7$\'\xe7\xbd/\xf2\xa4I\x93\xd8\xb7wo\x01\xaf\x97\xa3b\xdc\x12x\x14@1a6[\xd8\x7f\xe0\x00\x9f}\xf6\x19\xdd\xbave\xf1\xe2\xc5\x05Zm;\xee\xbd\xf7\x98<y2O=\xf5\x14e\xcaD\xba\xede_+V\xab\x95}\xfb\xf6\xbb\x1e\x06\xa0a\xc3\x86\x94/_^\xf5\x9a(\x00\xb6\xa6S\xa1BEu\xe1\x8f\x0b\xc9\xc9\xc9\xec\xdf\xbf\x1f!\x94\xd2=\n\x90r\xb3\xef\xe7\xa5\xb8\xdc\xe7\xbf]\x91$\x89\x80\x80\x00\xfa\xf7\xef\xc7\x07\x1f|\xc0\xd4\xcf>\xa3y\x8b\xe6\xae\xd9\x9c0\x99L\xfc\xb2r\x05={\xf6\xe4\xfd\xf7\xdf\xe3\xef\xbf\xff\xc6d\xca\xba\xee\xb2U\x84\xe0\xc0\xc1\x83\xf9\x8e\xa8m\xfc\xf4\xd3ONn\xbf\xb7\x1b7.M<\xe4\x89\x10\x82\x0bq\x17\xf8\xe6\x9bo\x181b\x04\xe3\xc6\x8d\xc3\xe4\xb0\xe9\xb8;\x02\x02\x02\xe8\xd3\xa7\x0f\xeb\xd7\xafg\xd8\xd0\xa14hP\x1foo\xefB\x0b\xb5e\xb5Z\xd9\xf9\xcf\xdf\xae\x87A\xb3y\xab+8s\x13z\xceH\x9a\x12\x88(S\x86\x88\x88\x08\xd7d\xb2\xb2\xb24\xbb\xb0\xb0/.*\x8d\xe4ugBh\xe6\x89\xbc2y\x00\xad\xbeK\x92D\x95\xcaU\x18\xfc\xd0`\xe6|1\x87!C\x86\xb8\xad\x1b\xaeL\x9b6\x9dW_{\x8d\xcf>\x9bj\xdf\xb3\xfa\x9a\x8b\\\x08\x92\n\xb0f\xc6\xc6\x81\x03\x07\n\xd5\xfct\xb3\xe1Q\x00E\x84\x10\x82\xabW\xaf\xb2v\xedZ\x9e\x1d\xf2,o\xbc1\x82\r\x1b6\xb8f\xcbA\xa7N\x9d\x984i\x12\xd3\xa7O\xa7C\x87\x0eZ\\\x15M\xf0\x17\x8e\xfc\xc7j\xb5:Mt\xda\x90$\x89\x88\x88\x08\x02\x02\x03\xaf\xe9R\x12\x10\x1d\x15\x95\xebf2j\x0fZ*4\x05V\x14\xe4ugB\xa8\xc1\xd2\xd4\xc0i\x1e\xf2\x92\xca6\xb3\xa4$K\xf8\xfa\xfaR\xbf~}&O\x9e\xcc\xacY\xb3\x180p \xa1\xa1\xa1\xae_q\xe2\xaf\x1d;x\xef\xbd\xf7xn\xc8\x10~\xfd\xf5W\x12\xaf^\xc5\xaa)\x82<.kG\x92\xa4k\xda\xb0&*:*\xd7\x91\x9f3\x05\xb9\xfa\xcd\x87G\x01\x14\x01\x19\x19\x99\xfc\xf3\xcf?L\x9c8\x91\x9e={\xb2j\xd5\xaa|\xf7a\r\n\n`\xfc\xf8\xf1L\x9a<\x99\xc7\x1f\x7f\x82\n\x15*\x14\x8a\xb9\xc7\x1dV\xab\xd5m\xafG\x96e\xc2\xc3\xc3\t\n\x0c\xcaG$\xe6${>\xc2\x19!\x04\x99YYdee\xe5aJ)\xddHZ$\x03q\x8der\xcbr\x8d\xc5\xe0\xe7\xe7G\x9f>}\x18\xff\xe1\x87L\x99<\x85\xea\xd5\xab\xbbfq"33\x935k\xd7\xd0\xb7o_&\x7f<\x85\xed\x7f\xfeIfzz\x81/\x1bU6\xca\xf5P\xae\xdc}\xf7=\x05\xdc\xf8\xbe\xa0W\xbf\xb9(\x1a\ts\x9b\xa2(\n\xc7\x8f\x1fg\xfa\xf4i\x0c\x1f\xfe2\x93\'O\x86|\xaaNpp0O<\xf1\x04+V\xac\xe2\x85\x17^\xa0^\xdd\xbaxy\x19\xdc\n\xe8\xc2B\x96\xd5\t\xe4\\\x91\n\xcfX#\x84 -5\x95\x8c\xcc\xcc\xbc\xafY\xea)\xcd\xe3\x97\xd2\x8fN\xa7\xa3r\xe5\xca\x0c\xbc\x7f K\x96.e\xf8\xf0\xe1\x94)S&\xdf:1q\xc2G\xbc\xfc\xf2\xcbL\x9d6\x8d\xc3G\x0e\xe7\xdb.$I"*\xba\xac\xeba\xb7\x94/_\x9e\xf6\xed\xda\x15YG\xebf\xe0\xf6}\xf2BDQ\x04\x89\x89\x89\xfc\xfc\xf3\xcf<\xf6\xf8\xe3\x8c\x1d;\x96\xbf\xff\xfe\xcb\x9e\x9e\x9b0\xed\xd5\xab\x17\x9f~\xf2)\x13\'N\xa4U\xabV\x04\x04\x04\xd8\xfd\xe5\xf3k\x187\x82^\xaf\xcf3>\x8f\x1a\xc0\xcd\xf5\xe8\xf5#\xcb:t\x92\x8cpp\xb1\xb4\xfd\xa6(\ni\xe9i\xa4\xa5\xa5\x91\x92\x92BzFF\xa9\x1b)\x08Q\xba\xe7/\nL\t?\x82,\xcb\xf8z\xfbR\xb7N\x1d\xc6\x8c\x19\xc3\xcc\x993\x190p`\xbe\x02x\xcf\x9e=\xbc\xfb\xee\xbb<\xf1\xf8\x13,_\xbe\x9c\x8b\x17/\xba5a\xa2\x05\xaf[\xb4\xe8\x7f\xae\x87s\x10\x10\x10\xc0\xa8Qo\xa9\x0e\x0f\xb71y\x97\xbc\x87|IMM\xe5\xaf\xbfv\xf0\xd1G\x1f1`\xc0\x00vl\xdf\x9e\xefb\xae\x1a5j0i\xd2$\xc6\x8f\x1f\xcf\xc3\x8f<BXXX\xbe\x8d\xa00\xc9_\xc0\xe6\x97^p$I\xc2\xdb\xdb\x88\xde\xa0Gh6t!\x04V\x8b\x85\xf31\xe7\xf9\xf3\xcfm|\xfb\xcd\xb7L\x9f>\x9d\x193f\xb0r\xc5/\xec\xf8\xeb/RRR\np\x9f\xc5\x83\xa4\x99\x81nzJ\xc33h\xdeV!!!\xdc{\xef\xbd\x8c\xff\xf0C\xbe\xf8\xe2\x0bZ\xb7i\xed\x9a\xd3\t\xb3\xd9\xcc\xdf\x7f\xff\xcd\xa0A\x83x\xe5\xd5W8|\xf8p\x8e\xfaa6\x9bX\xbbv-\xd3\xa7Ow:\xeeJHH(o\xbf\xfd6\xbd{\xf7q\xeb\xb9v;Q|R\xe7\x16\xc3j\xb5p\xf4\xe8Q>\xfb\xec3^}\xf5U\xbb\xb9\'7$\xd4J\xff\xc2\x0b/\xf0\xfdw\xdf3\xe4\xd9!\xd4\xae]\x1bYV\x17\x9f\x17\'\x92\xb6\x03\x94;\x84\x10jZ.\xe9\xd7\x83^\xafWW{j\xa7\xb4Z\xad\xfc\xb6\xee7F\xbd5\x8a!C\x860l\xd80F\x8f\x1e\xcd\xe8\xd1\xa3y\xf8\xe1Gx\xf1\x85\x17\x99={6W\x13\xaf\xba\x9e\xaaH\x11B\xe0\xb2\x0e,\x9b\xdc\x8e{\xb8n$I\xa2r\xe5*\x0c~\xe8!\xbe\xfd\xe6[\xdey\xe7\x1d\xca\x95+\x97\xe7\xe8\x14`\xd1\xc2E\x9c8q\xc2\xa9\x0e\x0b!\xf8\xf7\xdf\xff\x98={v\x8e\xd5\xc8m\xda\xb4\xe6\xab\xaf\xbf\xe6\xc7\x1f\x7fd\xe9\xd2\xa5l\xdd\xba\x85!C\x86\x10\x15U\xd0\t\xe0\xa2\xa2\xe4+\x95G\x01\xe4\x83\xf3+\x12\x98\xcdf.\'\\\xe6\xe7\x9f\x7f\xe6\xc1\x07\x1f\xe4\xc3\x0f?\xe4\xef\xbf\xdd\xbbT\xda\xf0\xf2\xf2\xa2\xff}\xfd\x99\xf3\xe5\x1c\xc6\x8f\x1fO\xe3&\x8d\xf1\x0fP\xc3\xe3R\x02\x1d3I\x96\xdc\xf6|\x14E!>>\x9e\xc4\xc4\xa4k\xee\xf2fdfbv\xe3\xdej\xf3\xca0x\x19\x90e\x19\xab\xd5\xca\x8e\x1d\xdb\x19\xff\xe1x\x16,X\xc0\xe1\xc3G\x00\xf0\xf1\xf6\xc1`0hk\x14\xf6\xf2\xe6\x9bo2\xe3\xf3\x19\xeaZ\x89bj\'\x02\x90\xdcx\xfa\x08\x8as\x08\x90\xf3\xfa\xb72\x92\x04\xdeF#U\xabV\xe5\x8d7\xde`\xc1\x82\x1fx\xe9\xc5\x17\x89\x8c\x8ct\xcd\n@hh(c\xdey\x87v\xed\xda!9\x8c\x9a/]\xba\xc4\x97_~\xc9\xfa\xf5\xeb\x9d\xf2\xd7\xaf_\x9f\xd1\xa3F\xf3\xd0\xe0\xc1\x0c\x1c8\x90\xbe}\xfbR\xabV-\x02\x02\x02JX\xf8S\x02-?\'\x1e\x05\x90\x0f\x8e\xaf(5-\x8d\xcd\x9b\xb70f\xcc;\x0c\x1c8\x88\xbd{\xf7\xe6k\xeei\xde\xbc9\x9f\xcf\xf8\x9c\x0f\xc7\x8f\xe7\xde\xbe\xf7\x12\x10\x10P\xac\xe6\x1ew\xc8\xb2L\xfb\xf6\xed]\x0f#\xb4=\x81\xf3\xf3Xr\xc4f\x1b?s\xe6\x0cq\x17.\xb8&#I\x92\xda\xa3\xd3zk\xe9\xe9\x19\xfc\xf1\xc7&\xb6o\xdfn\xef\xc1}\xf8\xe1x\x96.]\xca\x8f?\xfe\xc8\x90g\x9f\xc5`Pc\xc3\xcc\x993\x87\x9d;w\x16[$_\t\x90D\xceFY\xbc\x82\xa28\xafU\xba\xf0\xf1\xf1\xa1m\x9b\xb6\x8c\x1a5\x8a9s\xbe`\xf0\xe0\x07]\xb3\xf0\xdcs\xcf\xf1\xec3\xcf\x10\x12\x12\xa2zf\tAFF\x06K\x97.e\xf6\xec\xd9NyeI\xe6\xf1\xc7\x1e\xa3]\xfb\xf6\x18\x0c\x86\\=\xd5ngJV\x12\xdd\x0c\x085\x0e\xfa\x81\x03\x07\x980~\x02#G\xbe\xc1\x9c/\xbep\xcd\xe5\x84$I\x94-[\x96\x17^|\x81/\xe7\xce\xe1\xa1\xc1\x83\xa9V\xb5j\x89\x0b~\x1b:\x9d.WW\xbc\x13\'Op!.\xce\xf5p\xae\xd8|cRSSHw\xb3[\x99\xbf\xbf?5k\xd5\xb47<\xbd^G\xabV\xadx\xe5\xe5\x97\xb9\xe7\x9e{\xf8\xe4\x93Ox\xf6\xd9g\xe8\xdc\xa53\xdd\xbbw\xe7\xe5W_\xa5^\xbdz\x00\\\xba\x14\xcf\xb6m\xdb\x10\x88b\xe9\x80\xe76\xd9+I\xc58\x00p 73\xdd\xad\x8a\x10\x02!Axx\x04]\xbbvc\xe2G\x13\x99={6-[\xb6\xa0z\xf5\xea\xbc\xfb\xee\xbb\x0c\x1d:Tu\xf3\x14B\xed~\x08\xc1\xc6\x8d\x1b\x992e\x8a\xeb\xe9x\xfa\x99\xa7\x18\xfc\xd0C\xf8\xfa\xfa\xba&y\xd0(\x1d\x12\xa9D\xc8\xa7q\t\xd5V}\xe1\xe2\x05\x16-\\\xc8\x8b/\xbe\xc8\x94)S\xd8\x93\xcb.X6"##\x19<x0_\x7f\xfd5\x93\'M\xa6n\xdd\xfax{\xfb\x94\xba\x9eGpp\xb0\xeb!\x00\x0e\x1e\xf8\x97\xd8\xb88,\x16\xf7^\x16\xee\x10B\xe1\xcfm\xdb\xb8r\xe5\x8ak\x12~~~T\xaaXI\xed\xd8\n\xb5\x97ww\xc7\x8eL\x988\x91\x95+W\xf2\xe2\x8b/\x12\x1a\x12\x82,\xcb\x18\x8dF*U\xacH\xa3\xc6\x8d\x01P\x14+\x19\xe9\xe9X\xady\xbb\xfe\x15\x06\x02\x81\x94\x8b{\xac\x10`[\x0c\\\x9c\xb8\xbb\x97\x82`\xeb\x15\xa7\xa7\xa5c6\x9b\x8aF\x91\x14\xc1)%I[,(\xa9[\x9bF\x97+\xc7\x93O>\xc9\xb2e\xcb\xf8\xe5\x97_\x181b\x84\xea:*K\x80\x84\xa2\x08\x0e\x1f9\xcc\xdc\xb9s9\xe3\x12\x89\xb6M\x9b6<\xfb\xec\xb3.+\x90\x8b\xe0\xa6ornc\x05\x90w\xe3JLJd\xfd\xfa\xf5\xbc\xf5\xd6[<\xfa\xe8\xa3l\xdb\xb6-\xdf\x98!\xed\xdb\xb7\xe7\x93O>a\xfc\xf8\xf1t\xe9\xdcY\x0b\xdfP\xfa\x90$\x89Z\xb5j\xd1\xb4i3\xd7$L\xa6,V\xae\\IrrA\xa3$Bbb"\xa7N\x9dr;\x07\x10\x1a\x16FTT\x14\x88\xec}a%Y\xc6\xa0\xd7\xe3\xe5\xe5\xa5.\xc2\x91\xb2\'\xc2\xd3\xd2\xd39y\xe2\x04h\xa6*\x1f\xdf\xe2Q\x9e\x92\xc8\xcb\xcf\xdf\xfd\xc8\xa0\xb4!\x04\xc4\xc7\xc7\xb3y\xcb\x16\xe6\xce\x9d\xcb\x8c\x993X\xb6t\x19\xff\xfc\xf3O\xbe\x81\xd1\xae\x99\xdc\x0b\xabP\xd1\xe9tDF\x96\xa5F\x8d\xea\xce\xf3V\x12$&^\xe5\xbby\xdf\xf1\xcb/\xbf8\xd9\x08k\xd6\xac\xc9+\xaf\xbeB\xdd\xba\xf5\\\xeaN1\xdd\xf4M\xc4m\xac\x00r"\x84\xc0b\xb1\xb0\x7f\xff~>\xfcp<\xc3\x87\x0f\xe3\xbb\xef\xbes\xcd\xe6\x84$IT\xa8P\x81w\xc6\x8c\xe1\xab\xaf\xbe\xa2\x7f\xff~\x94/_\xdei\x82\xaa4\x12\x14\x14D\x8f\x9e\xdd]\x0f\x03\xf0\xd5W_\xf1\xe7\xf6\xed\xf9.\xbaA\x9b8\xde\xb7w?\xbf\xff\xfe\xbb\xdb\x9ef\xdf>}\xb4}\x05r\xa6\xa1y\x04e\x99Ldd\xa4s\xf6\xdc9\xe6|\xf1\x05\xfb\xf7\xab\x81\xea*U\xaaD\xfb\xf6\x1d\xaey\xb3\x9b\xebB\x1b\xa1\xe0\xe6\x91o\x06\xdb\xb1\xa2(\xec\xdb\xb7\x971c\xde\xe6\x99\xa7\x9f\xe6\xe5\x97_\xe6\xcd7\xdf\xe4\x91G\x1fe\xe8\xd0\xa1L\x9b6\x8d\xb8k0\xed\x95>\x9c\xcb\xdfl6\xb3|\xf9O9\xbc\xef\xbc\xbc\xbc\xb8\xef\xbe\xfb\xe8\xd4\xa9\x13^^\xd9\xfb\x0cX\xadV\xccf\xb3\xdb:Z\x9c\x94\xec\xd5\xdd`\xb5Z\x85\xe7c\x15&\x93I\x9c={V\xcc\x9c5K\xd4\xabWO\xc8\xb2\xac\x1a\x19\xf3\xf8DGG\x8b\'\x9e|B\xacY\xb3Z\xa4\xa4\xa4\x08\x8b\xc5\x92\xe3\xbc%\xf2\xb1\xb89\xe6\xf2\xb1X,\xe2\x8f?\xfe\x10\xb5j\xd5\xcc\xf1\\\x80h\xd5\xaa\x95\xd8\xb7\x7f\x9f0\x9b\xcdBQ\x84[\x14E\x11\'N\x9c\x10\x8f<\xf2H\x8e\xef\x03\xc2\xc7\xd7Gl\xde\xbc%\xcfr\xf9\xfd\xf7\xdf\xc5\xddw\xdf-\xda\xb7o/\xee\xbc\xf3N\xe1\xe3\xe3#dY\x16]\xbat\x11\xf3\xe7\xcf\x17\x99\x99\x999\xbeST\x9fs\xe7\xce\t///\xa7g\x08\x0b\x0b\x13?\xff\xfc\xb30\x9b\xcd9\xf2\x97\x9e\x8fE\x1c<xP\xb4n\xdd:\xc7;p\xfc\x8c\x1c9R\\\xbcx1\xcf\xf7Qz>\xb9\xdf\xa3\xc9d\x12\xbf\xae\xfeUT\xacT1\xc73>\xf8\xe0\x83\xe2\xf4\x993\xf6\xbcf\x8bE\xa4g\xa4\x8b\xaf\xbe\xfaJ<\xf3\xec\xb3\xe2\xcf?\xffT\xdb\xaa\x9b\xf3\xe6\xfe\xc9\xfd^n\xf6\x8fG\x01X\xac"!!A\xac\\\xb9B\x0c\x1e<8G\x85\xca\xed\xd3\xb1cG1\x7f\xfe\x0f\xe2\xdc\xf9s\xc2r\x93V\x90\xcb\x97/\x8b\xe7\x9f\x7f^H\xaa\xefc\x8eO\xf7\xee\xdd\xc5\xda\xb5k\xc5\x95\xabW\\e\xbf\xc8\xca2\x89\x03\x07\x0f\x8a\xc7\x1e{,\xc7\xf7l\x9f\xe1\xc3\x87\x8b\x84\x84\x84\x1c\xd7\xb5},\x16\x8b\xf8\xed\xb7\xdfDPP\xa0\xd3\xf7\xea\xd4\xa9+\x96.]*2\x8aQ\xf8[\xf3P\x00\xbf\xfcR\xba\x15@rJ\x8a\xf8\xe0\x83\x0f\x84\xc1`\xc8\xf1\x0e\x1c?\x92$\x899s\xe6\x88,\x93IXo\n%\x90\xf3c2\x9b\xc5\xde}\xfbD\x8f\x1e=r<_\xf3\xe6\xcd\xc5\xdf\x7f\xff\xad*8\x8b\xda\xb6M&\x93\xd8\xbe\xfdO\xbbrlP\xaf\xbe\x986m\x9a8p\xf0`\xb1v.J\xeb\xa7\x14o\tY\xf4\x98L&v\xed\xde\xc9\x9c9_2e\xca\xc7l\xd9\xb2\xc55\x8b\x13:\x9d\x8e\xf2\x15\xca\xf3\xe6\xc8\x91\x8c\x1c\xf9&\xcd\x9a5%$X\xddW\xb4\xd4"r\x8c\x9e\xedx{{\xe3\xe7\xeb\xcb\x96\xad[\xb9z5\xe7\xa2\xab\xe3\xc7\x8f\xb3w\xef^N\x9f>MRR\x12\xe9\xe9\xe9\xc4\xc4\xc4p\xe2\xe4)\x96\xff\xb4\x8c9s\xbe`\xc9\x92%\xae_\x03\xa0Y\xb3f\x0c\x1b>\x8c\xaaU\xaa \xcb\xee\x17\xf6H\x12$%%\x11\x1b\x1b\x8bN\xa7\xc3j\xb5jn\xb5\x82\x8c\xccL\xfc\xfc\xfc\xa8R\xb9r\x9e\xdeSy<\xde5\x93\x92\x92\xc2\xb4i\xd3\xdcn\tY\xa3F\x8dB~\xcfy\xdcy^IB\x80\x94=_!\x04\x1c?q\x9c\xaf\xbf\xfa\x8a#G\xd45\x15y\x11\x1a\x1aJ\xabV\xad\xf0\x0f\x08pM\xba)H\xbcz\x95/\xbe\x98\xcd\xb7\xdf~\xebt<<<\x9c1c\xde\xe6\xae\xbb:j\xf3J P#\xf2~\xfe\xf9\xe7,]\xba\x14\x80\x8b\x97.\xb1f\xcd\x1aN\x9fVw!\xabT\xb1\x12Foc!\xbf\xdb\x9b\x08W\x8dp\xab\x7f,\x16\x8bHOO\x17g\xcf\x9e\x13S\xa7N\x15U\xabV\x15:\x9d.Go\xc2\xf5\x13\x16\x16&^x\xe1\x05\xb1y\xf3f\x91\x96\x96v\xd3\xf6\xfa]?\x99\x99\x99b\xee\xdc\xb9""""\xc73\xdb>\xb2,\x8b\xa0\xa0 \x11\x15\x15%\xa2\xa2\xa2Ddd\xa4\xf0\xf6\xf6\xce\x91\xcf\xf6\t\t\t\x11\x9f~\xfa\xa9HNI\xcew\xf8l6\x9bEzz\xba8z\xec\xa8X\xb5j\x95x\xf9\xe5\xe1\xc2\xdf\xdf_\x00\xa2f\xcd\x9ab\xd3\xa6M9\xbeST\x9f\xdcF\x00+V\xac\x10\x16\x8b\xe5\xfa\xdey\x01\xccq7\xf21\x99Lb\xf5\xaf\xabEPPP\x8e\xf7\xe0\xees\xdf\x80\x01\xe2\xe4\xc9S7\x89\x19(\xfbc\xb1XDZZ\x9a\x98>}z\x8eg\xd2\xe9t\xe2\xe5W^\x16\x07\xff=(.]\xba$\xb2\xb2\xb2\x84\xc5b\x11\x19\x19\x19\xe2\x9bo\xbe\xc9\x91\xdf\xf6\xf1\xf3\xf3\x13\x0f>\xf0\x80\xd8\xf1\xd7_")9\xa9P\xca\xc4\xa2}\\\x8f\x97\xd6O\xee]\xab[\r-\xc4\xc1\xe5\xf8\xcb\xac]\xbb\x96\x17^x\x8e\xe1\xc3\x87s\xe2\xc4\x89\\\x03K\xd9\xe8\xd3\xa7\x0f\xb3\xbf\x98\xcd\xe8\xd1\xa3i\xd3\xa6M\xa1n\xceR\xd2\x18\x0c\x06\xfa\xf5\xef\xc7\x1bo\xbcA@@\xa0k2h\x13\x8cIII\xc4\xc5\xc5\x11\x17\x17\xc7\xc5\x8b\x17\xc9\xcc\xcct\xcdfg\xd4\xa8Q<\xf6\xd8\xa3\xf8\xf9\xf9\xe5\x1bB\xd9\xe6\xfeY\xb9re\xbat\xe9\xc2\x1b#G\xd2\xa7O\x1f\x00\x8e\x1d=\xca/\xbf\xfcR\xa0\x9d\xd3\x8a\x12\xa1\xd5\x1dw\x8b\xc4\xf2\xe5:\xber-\xc8\xb2\x8cN\xaf#++\xcb5\xc9-\xb2$kk\x1a\x84kR\xa9FQ\x146m\xda\xc4\xc4\x89\x13]\x93\xe8\xdb\xb7/\xfd\xfb\xf7\xc7\xcf\xcf\x8f\xd4\xb4T\xe2\xe3\xe3IMK\xe5\xe8\x91\xa3|\xf0\xc1\x07\xae\xd9\xed\xa4\xa5\xa5\xf1\xe3\xc2\x85<\xf8\xc0\xfd\xcc\x9c1\x83\x83\x07\xff\xc5b\xb1\xdcP\xc9HE\xff\xca\x0b\x95\xdbB\x01\x08!\xc82\x99\xd8\xbau+\xef\x8e{\x97W_{\x95\x95+W\xb9fsB\x96ej\xd5\xaa\xc5\xe4\xc9\x93\xf9\xf4\x93O\xe9\xdd\xab7\x91emafo\xa4\x8a\x94>\x82\x02\x83x\xf2\xc9\'\xf9x\xcad*T\xa8\xe0\x9a\\`\xca\x96-\xcb\xd4\xa9Sy\xec\xf1\xc7\x08\n\n\x06\x91{cp\xf5\xc6P\x05\x93DxX8w\xddu\x17h\xdb\xfb\xc5\xc6\xc6j\xeb\x0bJ\xb2\xcc\xd5\rNJ\xeb\x860\xe1\xe1\xe1\xb4l\xd9\xd2\xf5\xb0[.\\\x88\xe3\xe2\x85\x0b(\xb9\x06=*}\x08\xa1p\xe4\xc8\x11\xe6\xcc\x99\xc3\xf9\xf3\xe7\x9d\xd2\xa2\xa3\xa3y\xe2\xc9\'(W\xae\x9c\xbdN\x99\xcc&\x12\x13\x13I\xcbH\xa3]\xbbvx\x19\xbd\xf24\xf1\x9c:u\x9a\xb7\xde\x1a\xc5\xc8\x91o0{\xf6,\xe2\xe2\xe2\xd4\xf2\xb9y\x8a\xe8\xba\xb9\xa5\xe7\x00\x84\x80,S&\xe7\xcf\x9fg\xee\xdc\xb9\x8c\x1d;\x96\x8d\x1b7r\xf5JN{\xb7#\x95*U\xe2\xe1\x87\x1ff\xdc\xb8q\xf4\xee\xdd\x9b\xb0\xb00t:\x9d\x830\xcb\xbd2\xdd\x8cH\x92\x1a\x1b\xa8~\xfd\xfa\xd4\xadW\x17\xc5\xaa\x90p\xf92\xa9\xa9\xa9\xaeY\xddb4\x1a\xb9\xb7\xef\xbd\xbc\xf9\xe6\x9b\xdc\xdb\xef^\x82\x83\x82\xdd68\xa1\xed\x92v\xfe\xfcy\xb6o\xdf\x8e\xd5j%44L\xeb6\xa9}\'\xb3\xd9\xcc\xea\xd5\xab\xd9\xb4i\x13\x005k\xd5\xa4\xd3=\x9ds]\xb8\xa6"\n\xe5\x9d\xe4>\x07\xd0\x97\xea\xd5\xabk=\xe7\x1b\xbfN\xc1\xc9\xff\xb9\x14!HHH\xe0\x8f?6q\xfc\xf8q\xd7\xe4\x1c\x9c?\x7f\x9e\x03\x07\x0e\x12]>\x9a\xb0\xf00\xbc\xbcT\xe1\xe8\xee}\x95$\xaa0W\x83\x16\xc6_\x8eg\xeag\x9f\xe5p\xc9.[\xb6,c\xde\x19C\xb3f\xcdr\x04\x90\x13B\xe0\xef\xefO\x93\xa6M\xa8Y\xb3&AAA\x9c9{\x06SV\xee\xa1[N\x9c8\xc1\xa6M\x9b8u\xfa\x14A\xc1A\x84\x85\x85a\xf02 Ir>o\xe1\xe6\xe5\x96U\x00\x8a\xa2p\xe1B\x1c\xbf\xfd\xb6\x96\xf1\x13>b\xce\x9c9$&&\xe6\xe8y\xbar\xdf}\xf7\xf1\xfa\x88\x11<\xf2\xf0\xc3T\xadR5G\xc5\xba\x95\xd1\xe9t\xdcQ\xb92\xcd\x9b5\xa3q\xe3\xc6T\xa8P\x81\xd8\xd88\x92\x92\xdc\xef\xb1\x1a\x16\x16\xc6\xd0\x97\x862t\xe8P\x1e~\xe4aZ\xb4h\xa1.\xd6q\x98\xa4t\xc4l\xb10\xfa\xed\xd1<\xfa\xe8\xa3,\\\xb8\x90\x94\x94\x14\x1a6l@pp0\x92\xb6\x93\xda\xc1\x7f\xd5\x90\x1b\xf1\xf1\xf1H\x92D\xc7\xbb;\xd2\xa3{w\xb7\xc1\xeb\xb2qw\xb5k\'99\x99iS\xa7buX\xff`\x9b\x04\xaeV\xad\x1a\x92T\xdc\x03\xe6\xbc\x9fKh#\xa4\xd9_\xccf\xfe\xf7\xdf\xbb&\xbbE\x08\xc1\xf9\x98\xf3\xfc\xb8\xe0G\xacV+\xc1\xc1\xc1DD\x94A\xaf\xd7\x15@\xdd\x14\x1f6e\x9b\x95\x95\xc9\xa2E\x8b\x18;v\xacS\xba^\xaf\xe7\xa9\xa7\x9f\xa2g\xcf\x9e\xf8\xf9\xf99\xa59\xe2\xe5\xe5E\x8d\x1a5h\xd4\xa8\x11uj\xd7\x01\xe0\xd8\xb1c\xae\xd9\xecX,\x16\xfe\xfb\xf7?~\x98\xff\x03z\xbd\x1e\xff\x00\x7f\xc2\xc2\xc2\xf1\xd2g\xaf)\xb8\x95\x90\xacVk\xde\x12\xf1&C\x08\xc8\xcc\xcc`\xcb\xe6-,Z\xb4\x90\xb5\xbf\xfd\x96\xef\x02\x18\x9dNG\xed\xda\xb5y\xe9\xa5\x97\xe8\xd9\xb3\x07\x11\x11e\xec\x01\xc9\n\x03\x9b\rY\xdd\x8a\xd1\x8a\x10\xaa\x89\t!\xd0\x17B\x90*!\x04\x8a\xa2\xd8{\xae6%\'\xc9\x12\x06\xbdz\xfe\x82\xa2\x08\xd5jo\xb5XHNI&%%\x953g\xcep\xf9\xf2e\x14EA\x11\n\xb2$S\xae\\9\xca\x97/O`` ~\xbe~\xe8\xf42\x92$\xab\xb6\xf2\\\xaegU\x14\x16\xfe\xf8#\x8f>\xfa(h\x1e)\xf7t\xeaD\xc3\x86\r)\x13Q\x86c\xc7\x8e\xb2k\xf7n\xfe\xf8\xfdw\x14E\xa1B\xf9\xf2|9w.\x9d;wv=U\x91\x10\x13\x13C\x95*U\x9cV|\x87\x85\x85\xf1\xed\xb7\xdf\xd0\xad[\x0f-&\x90\xfbg+\t\xce\x9f?\xcfG\x1f}\xc4\xacY\xb3\\\x93\nL\x9b6mx\xe8\xa1\x87\xe8\xd3\xa7\x0f\x91\x91e\x91\xe5\xd2\xf3|\x8a\xa2\xb0f\xcd\x1a\x86\x0e\x1d\xca\xe9\xd3\xa7\x9d\xd2Z\xb6j\xc9\xe4I\x93)W\xbe\x9c\xd3\xf1\xbc\xb0X,$$$\xf0\xeb\xaa_Y\xb0`\x01g\xce\x9c\xc9wu\x7f\xcbV-\xe9\xdb\xa7/\x0f<\xf0\x80\xb6\xc0\xb3t\xefm}\xad\xdc\xf4\n\xc0\xb1\xd7\x92\x9e\x91\xc1\x99\xd3\xa7\x99;w.\x8b\x16-\xe2\xc2\x85\x0b\xf9\xf6\xf8k\xd7\xaeM\xd7\xae]y\xe8\xa1\x87h\xd0\xa0A\x01\xf7\x07\xcd\x7fp.\x84\xc0\xaa(\xa4\xa7\xa5q\xf2\xc4I\xfe\xfd\xef_\xd6\xaeY\xcb\xf9\x98s\xa8\x9e|\x12AAA\xf4\xbd\xb7/\xf5\xeb7\xa0r\xe5\xca\x04\x05\x05i\x8a!\xef\xe0cB\xa8\xc6\xf5\xcc\x8cL\x92\x92\x928~\xec8\xfb\x0f\xecg\xf5\xea\xd5\xa4i\x01\xd9\x84\x10\xb4m\xd7\x96\xe6\xcd\x9aS\xa7n]\xca\x96\x8d\xc4\xc7\xdb\x07$\xd5\xde^Pl\xe5g\x1b\x92K\xdav\x91\xea\xaeay\xdcd.$$$0s\xe6Lf\xcc\x98a\x8f\xdbn0\x18\xf02\x1a\xc9HO\xb7\xaf>\x0e\x8f\x08\xe7\xdd\xb1\xef\xf2\xf4SOcpX\xd1Y\x94\xc4\xc4\xc4P\xadZuL\xa6\xec\t\xd5\xb0\xb0p\xbe\x9d\xf7\r\xdd\xbbv\xb3\xc7\xa0)il=\xff\x8f>\xfa\x88/\xbf\xfc2G\x08\x0e\xa3\xd1\x88^o\xc0j\xb5\xe0\xe3\xebC\xf3\xe6-HNNb\xfb\x9f\xdb\x9d\xf2\xd9\x08\x08\x08\xa0W\xaf^\x0c\x1b6\x8c:u\xea\xe0\xe7\xe7\xa7\xbe\xdb\xfc*\xb9[\xae\xebK9P\x14+\x07\x0f\x1e\xe4\xad7\xdfb\xcd\xda\xb5Ni\xf5\xea\xd6c\xe2\xe4\x89\xd4\xaaU\xeb\xba\xea\xa0\xc5b\xe1\xc8\x91#l\xdc\xb0\x91U\xabVq\xf4\xe8Q\xd7,Nx{{\xd3\xe1\xae\xbbx\xe1\xf9\xe7i\xd9\xb2%\xa1\xa1\xa1\x9akr\xe1<kN\x8a\xea\xbc9\xb9\xe9\x15\x00\x80P\x04\xe7\xce\x9fc\xed\xda\xdf\xf8\xe6\x9b\xaf\xf9\xeb\xaf\xec\xed\x18\xf3b\xf0C\x83y\xf8\xa1\x87i\xd6\xbc\xb9\xe6\xcf\xef\x9a\xe3\xfa1\x99\xb2\xd8\xbdg/k\xd7\xac\xe1\x9bo\xbf\xe1\xdc\xd9s\xaeY\x9cx\xfd\xf5\xd7\xe9\xdd\xbb7M\x9a6\xc5\xdb\x98\xb7_\xb2\xa2(\xc4\xc6\xc6\xb2q\xe3\x06\x96,Y\xca\xaaUyOh7n\xd4\x88G\x1f\x7f\x9c\xbb\xee\xea@\x8d\x1a50zy\x15[\x05sE\x08\xc1\xa5K\x97\xd8\xb0a\x03\x93&M\xe2\xc0\x81\x03\xaeY\xe8xwG\x9e}\xe6Y\xbat\xe9\x92\x8f\xed\xbfp\x89\x8d\x8d\xa5j\xd5\xaaN!\xbe\xc3\xc2\xc2\xf8\xe6\x9bo\xe8\xde\xbd{\x9e\xeb\x11\x8a\x93\xb3g\xce2c\xe6L\xa6L\xc9\xb9\tQ\xabV\xadx\xfa\x99\xa7\t\t\n\x06Y\xc2\xcf\xd7\x8fr\xe5\xcb\x91\x9a\x92\xca\xca\x95+\xf9\xf0\xc3\x0fs\xed\x14\xd5\xabW\x8f\xa7\x9e~\x8a\x1e=zR\xa5J\x15ud\xea\x9a\xa9\x98HHH`\xc2\x84\t|\xfa\xe9\xa7N\xc7+W\xae\xcc\xab\xaf\xbdJ\xe7\xce\x9d\xf1\xf2\xf2rJ\xbbV\xd2\xd3\xd3\xd9\xb9s\'3g\xcc\xe0\xef\xbf\xffqMv\xcb\x88\x11#\xe8\xd5\xab\x17\x8d\x9b4\xc1\xd7\xc7\xc75\xf9\xa6\xe3\xa6V\x00B\x08\xd2\xd2\xd2X\xbf~=K\x96,\xe1\x97_~\xb1\xf7\x80sCo\xd0S\xa3z\rF\xbc\xf1\x06\xdd\xbav%44\x14\x9d^\x8f\xa4\xe9]\xf20a\x14\x94\x94\x94\x14\x16.\\\xc8w\xdf\xcd\xe3O\xad\xd7%\x15 \xa6}\x93&Mxc\xe4\x1b\xf4\xec\xd1\x13\x9f\\*\x97\xa2(\xfc\xf7\xdf\x7f\xcc\x9a5\x8b\x1f~\xf8\xa1\xc0\xb1\xfb%I\xa2s\xe7\xce<\xff\xfc\xf3t\xed\xda\x15\xa3\xd1\xe8\x9a\xa5\xd8\x10B`6\x99\xb8\x14\x1f\xcf\xf1\xe3\xc7\xedQD\x85\x10D\x96\x8d\xa4z\xb5\xea\x84\x04\x07\xa37\xa8\x9b\xc8\x14\x17\xb9)\x80y\xf3\xe6\xd1\xbd\xbb\xfb\xb8I\xc5\xcd\xf9\xf3\xe7\xf9\xf0\xc3\x0f\x993g\x8ek\x12\xcd\x9b7\xe7\xcd7G\xd2\xf0\xce;\xd1\xe9t\x04\x06\x04\x12\x18\x18\x88$IX\x15\x85\x94\xe4\x14\xf6\xec\xde\xc5\'\x9f}\xc6\xba\xdf~\xcb1r\x000x\x19\xe8\xd7\xaf\x1fO=\xf9\x14m\xda\xb6\xc5\'\xcf\xb9\x97BF\xeb\xf8\x9aL&\xe6\xcd\xfb\x8e\xe7\x9e\x1b\xe2\x94\xac\xd7\xeby\xf2\xc9\'y\xe1\xc5\x17\x08\x0ct\xef\xb2|\xadx\x1b\xbdIJJb\xfe\xfc\xf9|7o\x1e\xe7cbrU\x906\x1a7n\xcc}\x03\x06\xd0\xef\xde{\xa9^\xbdz\xb1\xd6\xd1\xc2\xe6\xa6Q\x00\x8e\x83"\x01d\xa4\xa7s\xec\xe81\xbe\xfb\xfe;\xe6}7\x8f\xc4\xab\xf9O\xf0\xd6\xad[\x8f\xbe}\xfb0\xe8\xfeA\xd4\xaaY+\x0f;\xff\xf5\x0f\xc1L&\x13\xbf\xfe\xba\x8a\xd7_\x1f\xc1\xa9S\xa7\\\x93A\x1bR\x9aL&\xb7\xc1\xd6"""\xf8r\xee\x97t\xe9\xdc\x05///\x84\x10\xf6\n\xa6(\n\xa7N\x9db\xf4\xe8\xd1,Y\xb2$\xdf\xe7uGdd$_\x7f\xfd5w\xddu\x17\xdeFc\xde\xb6\xa6b\xc0\xdd3\xdc\xa8\x02\xbe^T\x13P\xb5\x1c\n\xe0\xdbo\xbf\xa5{\xf7\xee%v_\x00\x8aU!\xeeB\x1c\x1fM\x98\xc0\x97s\xbf\xc4lv\xb6]\xb7i\xd3\x86\x97\x86\xbeD\xab\x96-Q\x84@\x085\xe0_\x88\xcb\x08JQ\x14\xce\x9c=\xc3\x82\x1f\x16\xf0\xd3O?\xb1{\xf7n\xa7t\xb4\xf2\x8f\x8e\x8e\xe6\xb9\xe7\x9e\xe3\xe1\x87\x1f\xa6l\xd9\xb2\xf6\rU\x8a\x12\xa1\xcde\xfd\xfa\xeb\xaf\xbc\xf4\xd2K9\\>;w\xee\xcc{\xef\xbf\x97\xebna\xd7C@@\x00!\xc1!\xa4\xa7\xa7\xb3u\xebV\x96,Y\xc2\x9a5\xab\x89\x8d\xcd{\xde\xd0`0\xd0\xa8Q#F\x8c\x18A\xdbv\xed\x08\r\rA\xaf+\x98\xf9\xb8\xc4q\x10o7\x8d\x17\x90\xad\xea)\x8a\xc2\xf1c\xc7X\xb6l\x19\xa3G\x8ff\xc5\x8a\x15df\xe4\xbe(\xc9\xc6\xd3O?\xcd\xeb\xaf\x8f`\xe0\xc0\x01T\xacP1\x1f\xef\x9e\xeb\xac\xe8\x02\x0e\x1c<\xc0\xa8Q\xa39x\xf0\xa0k*\x0f>\xf8 /\xbe\xf8"\xcf<\xf3\x0c=z\xf4\xa0B\x85\n9\xccU\xe9\xe9\xe9l\xdb\xb6\x8d\x86w\xdeI\xa5Jw8\xc5\xa8\xbfz\xf5\n\xef\xbd\xf7\x1e\xf3\xe7\xcfw\xfa\x8e\x8d{\xee\xe9\xc4#\x8f>B\x97.]\xa8X\xb1"{\xf7\xeeu\xcdBZZ\x1a[\xb7n\xa5I\xd3&\xdcQ\xf9\x8e\x12\x9f\xd0\x92\xb4\tp\xc7OI\x91\x9a\x9a\x9a\x8b\x1b\xa8\xea\x05DI*\xa7\xd8\x18\xa6O\x9f\xce\xd4i\xd3rt\x1c\xda\xb4m\xc3[\xa3F\xd1\xb8q#\xbbB\x95\xb4\x8eFN\xef)\x89\xe0\xe0\x10\xee\xbc\xf3N\x1a4l\x08B\xb0o\xdf>\x97<\xea(\xf6\xf7\xdf\x7f\xe7\xe4\xa9\x93\x84GD\x10\x19\x19\xa9\xfa\xd3\x17a}\x11B\xf0\xdf\xa1\xff\x984q\x12\xbbv\xefrJk\xd1\xa2\x05\xc3\x86\x0f\xd3\xbc\xb1\n\xef\x1e\x8cF#\xbe>\xbexyyQ\xb5jU\x9a7oN\x8d\x1a5\xb8p\xe1\x02\xe7\xce\xe5n\xb6U\x14\x85\x98\x98\x18\x16/^LzZ\x1a\xc1A\xc1\x84\x84\x86\x14y\x19\x15\n\x0e\xb7w\xd3(\x00\xa1\xf9\x90\xff\xfc\xf3\xcfL\x9d:\x95\xcf>\xfb,\xc7\xe6\xcf\xae\x18\x0c\x06\x9a4n\xcc\xa4\x89\x93x\xf2\xa9\'\xa9]\xa7\x16>>>\xaa\xa9\xe7\xfa\xc5|\xae\xa4\xa5\xa7\xf1\xc1\x07\x1f\xb0b\xc5\n\xa7\xe3e\xcb\x96\xe5\xe3\x8f?a\xe8\xb0a4m\xd2\x84Z\xb5jR\xbdFMZ\xb5jE\xbbv\xed\xd8\xbau\xabS\xbc\xf6\x94\x94d$$\x9a6kFPP\xa0}\x13\xf7\xb5k\xd70e\xca\xc7\xa4\xa7\xa7;\x9d\xbfr\xe5\xca|\xf6\xd9g\xbc\xf1\xc6\x08:\xb4\xef@\xcb\x96-\xe9\xd4\xa9\x13\x83\x06\r"33\x83\x83\x07\xffu\xeai\'\'\'#I2\xcd\x9a6)\xb4\xa1\xf4\xad@rr2S\xa7Nu\x12\xb0\xbe\xbe\xbe\xf4\xed\xdb\x97\x1a5\xaa\x97\x80\x1b\xa8JlL,\x13>\x9a\xc0\xb4i\xd3r\xd4\xd96m\xdb0\xea\xadQ\xd4\xab[\xd7\xe9\x1d\x0b\x04\xdeFo\xbc}\x9c\x15\x80\xed\xfbF//\xcaW(O\xf3\xe6\xcd\xa9X\xb1"\'O\x9et\xbb\xa1\xcf\xd1#G\xd9\xb3{7\xf1\xf1\xf1T\xa9R\x95\xc0\xc0\xc0"3y\\NH`\xf6\xac\xd9\xcc\xfbn\x9e\xd3\xf1\xa0\xa0 ^|\xf1E\xda\xb4mS`\'\x8d\x82b4\x1a\xed\xe6VI\x92\xf0\xf7\xf7\xa7Z\xb5jt\xe8\xd0\x81\x90\xb0P\x8e\x1f;NZZ\x9a\xdb\x91\xaa\x8d\x9d;w\xb2\x7f\xff~\xe2\xe3/Q&2\x92\x90\xe0\x90|:\x98\xa5\x87R\xaf\x00\x84PHMMe\xcf\x9e=|\xf4\xd1GL\x9d:\xd5m\xcf\xd6\x95\xc6\x8d\x1b\xf3\xc2\x0b/0j\xf4h\xda\xb5kg\xdf\x8bWh6~u\x0b\x12\xd7\xe6t\x03\x08\xc1\xa6M\x9b\x98<yr\x8e\x05T\xef\xbf\xff>O>\xf9\x14AA\x81\xda\xc2\x1b\x19\xbd^\x8f\x8f\x8f\x0fwT\xba\x83;*Wf\xd7\xae]N\x01\xd9._\xbeL\xdb6m\xb8\xe3\x8eJ\xc8\xb2LRR\x12\xdf|\xf3\r\x9b7ov:w\xb5j\xd5\x984i\x12\x03\x06\x0c ((\x08\xa3\xd1\x88\x97\xc1\x0b\x1f_\x1f\xa2\xa2\xa2h\xd7\xae=\xdeF#{\xf6\xecq\n\x17p\xfa\xf4)\x9a5mF\x95*Un\x9a\xcaZ\xd4\xa4\xa4\xa40}\xfa\xf4\x1c#\x00\x9b\xad\xb70{\x9e\x05\xc1\xaa(\xc4\xc6\xc40\xe1\xa3\t\xcc\x9d;7G\xcf\xbfu\xeb\xd6\x0c\x1b:\x94F\x0e=\x7fG\xbc\x8d\xda\x08\xc0\xf1\xb6\xb5\xed-%IB\x96e\x02\x02\x02h\xd2\xa4\t\xcd\x9b7GhkgRS\x9d\xe7\xd1\xe2/_f\xc7_\x7fq\xe0\xc0\x01\x82\x82\x83)S\xa6\x0c\xdeFU\xb1\xd8v\xf0\xbaQ233Y\xb0\xe0\x07\xde~\xfbm\xa7\xe3\xb2,3\xe4\xb9!\x0c\x1c80\xd7y\xb1\x1b\xc1Q\x01\xd8\xd0\xeb\xf5\x84\x85\x87\xd3\xb2E\x0b\x9a6m\x8a^\xaf\xe7\xd4\xa9Sy\x86?\x89\x8b\x8bc\xfb\xf6\x1d\xfc\xf1\xfb\xefx\xfb\xf8P\xbe|y\x8cF\xa3:\x82w\xcd\\\x88\xd8\xe4Y~\xe4\x96\xafT+\x00\x8b\xc5\xc2\xe1\xc3\x87Y\xb4h\x11/\xbd\xf4\x12\xdb\xb7o/P\\\x98\xd7_\x7f\x9da\xc3\x86\xd1\xf7\xde{)[6\xd2\xa9\xc7\x92]\x089\x0b\xe3\xfa\x11d\x9a\xb2\x98\xff\xfd|~\xfb\xed7\xa7\xc68p\xe0\x00\x86\x0f\x1bNhH\xa8\xdbK\xca:\x1de"\xcbp\xee\xec\x19v\xee\xcc\x1e\xf6\xa6\xa4\xa4\xd0\xa4I\x13\xea\xd7\xab\x8f\xc1K\xcf\xa1C\xff\xf1\xf5\xd7\xdf\xe4\x18\x96\x8e\x1c9\x92\xc1\x83\x07\xe7\xd8-\t\xedY}||\xa9T\xa9\x12\'O\x9d\xe2\xd0\x7f\xff\xd9\xb3defQ\xae\\9Z\xb6l\x99\xa3\x01\xdc\xae\xa4\xa4\xa40u\xea4\x14\xc5Y\x01\xf4\xbd\xf7^\xaaU+>\x05`\xeb\x9a\xc4\xc6\xc42}\xfat\xa6\xb91\xfb\xb4h\xd1\x82\xb7F\x8d\xa2i\x93\xa6n\x85?H\x18\xbd\x8dj\xdc\xaa<\xee[\x92$5\xcam\xf9r\xb4h\xde\x82\xea\xd5\xaa\xb3{\xf7n\x92\x92\x9cw\x84\xb3\xcd?\xfdo\xd1"\x0c\x06\x03a\xe1\xe1\x84\x04\x87\xa07\xdcx\x8f\\Q\x14\xfb\xee{\xae;\x97\xdd\xdb\xef^\x9ex\xe2\t\x97\xad\x1d\x0b\x0fw\n\x00\xad\xfc\r\x06\x03\x95*U\xa2q\xe3\xc64l\xd8\x90\xf4\xf4\xf4<W[+\x8aB||<+W\xac\xe0\xe2\xc5\x0b\x04\x04\x06\x10Y\xa6\x0c\xde\xde\xde\xda;-\xe4N\xa7\x93<\xcb\x9b\xdc\xf2\x95J\x05 \x84\xe0\xf2\xe5\xcb,Y\xb2\x94i\xd3\xa61{\xf6\xec\x1cf\x0fW\xbc\xbd\xbd\xb9\xeb\xae\xbb\x988q"\x83\x07\x0f\xa6Fu\xd5\xdd1\xb7\x07/<\xd4\x97z\xe1\xe2E\x96.]\xc2\xfe}\xeanV\x00:\x9d\xcc\xc8\x91o\xd1\xb4y3\xf4\xfa\\z\xd9\x92\xbaO\xae\xd1\xcb\x88N\xa7\xe3\xb5\xd7_\xe7\xd9g\x9f\xe5\xcd\xb7\xde\xa2y\xf3\xe6\xf8\x07\xf8c\xb5*l\xdb\xf6\'\x9f\x7f\xfe\xb9Sco\xd3\xb6-\xcf=\xff\x02\x15+V\xcc\xf59%\t\x82\x02\x031\x9bMl\xdc\xb8\xd1i\x14\x10\x7f)\x9e\xfb\xee\x1b@hh\x88\xd3wnWr\x9b\x03\xe8\xa7\x85\x83\x16y4\xa4\xc2D\xd2&\xa4?\xca\xc5\xec\xd3\xb6];F\x8f\x1aE\xbdzu\xf3\xd8\xaaR\xe0c\xf4v+\xdc\xdc"I\x04\x04\x04P\xa3fM\xbat\xe9\x82\xc5b\xc9\xb5\xd7\xbbe\xcb\x16N\x9c8\x81\xd9b\xa6r\x95\xca\xf9*\x99\xbc\x10\xc0\xd1\xa3G\x99<q\x12\x7f\xff\xf3\xb7SZtt\x14o\x8f\x1eM\xb5\xea\xd5\x9d\x8e\x17&\xb9)\x00;B\x10\x18\x14D\x8d\xea5h\xd3\xae\r\x15+V\xe4\xf4\xa9Sn\xc3\xa7;\xb2\x7f\xff\x01\xf6\xec\xde\xcd\x85\x0b\x17\x89\x8c\x8c$,,,\xd7\xb0\xe8%I)R\x00jENIIc\xe7\xce\x9d|\xf0\xc1\x07\xcc\x98\xf1\xb9\xdb\xc9TWZ\xb5j\xc5s\xcf\r\xe1\xad\xb7F\xd1\xbcysu!\x8b\\8C\xd3\xfc\x91\x10\x08\xce\x9e=\xc3\xf2e\xcb8y2\xdb\xf3\xa7j\xd5\xaa<\xfc\xf0`*\xdfQ9\xd7\x06"iB\xa5R\xa5J\xf4\xea\xd5\x8b\x06\r\x1aP\xb5jU\xcaDD\xe0\xe7\xe7\x87N\xa7#9%\x85U\xbf\xfe\xca\x1f\xbf\xff\x9e\xfd=I\xe2\xee\xbb\xef\xe6\xbe\xfb\xfa\x13\x10\xe0\xeftNG\x84}o]_V\xff\xba\xdai\xde\xe4\xea\xd5\xabt\xec\xd8\xb1\x08b\xdd\xdf\x8c\x08\x92\x92\x92\xf9\xfc\xf3\xcfs(\x80{\xfb\xddK\xb5b\x8a\x05d\xb5*\x9c?\x7f\x9e\x89\x1fM`\xeeW_\xe5\xe8\xf9\xb7k\xd7\x8eW^~\x99;\x1b\xdd\x99K\xcf\xdf\x86\x84\xd1\xdb[s\xe3\xcc\xbf\xe7iK\xd5\xeb\xf5DDD\xd0\xb6][*V\xaaHJrr\x8eU\xb8\x00\'O\x9ed\xfd\xba\xf5\x9c\x8f9O\x85\xf2\x15\x08\n\x0eB\xaf7\xd8\xcbH \n4\x19z%!\x81/\xe7\xcca\xee\xdc\xb9N\xc7\xa3\xa2\xa2\x18\xf3\xce\x18\x9a7o^\xa4&\xca\xfc\x14\x80\xa49&\xe8\xf42!!\xa14m\xda\x94\x96-[b4\x1aIHH\xc8S\x11\xc4\xc7_\xe6\xaf\xbf\xfeb\xfb\xf6\xed\x18\x0c\x06"##\xf1\xf1\xf1A\x96u%\xed|g\xa7hfs\xae\x19Aff&\x07\x0e\x1c`\xce\x9c9\xb4o\xdf\x9e\x1f\x7f\xfc1\xc70\xd4\x1do\xbd5\x8aO>\xf9\x84\xa1C\x87Q\xb1\x92j//n\x84"\x88\xbf\x14\xcf\x9e\xdd{\x9c\x8e\xd7\xab_\x8f\xe8r\xe5\x9d\x8e\xe5\x86^\xaf\xc7\xa0\xf9\xbd;\nc!\x04\x99\x99\x19\xfc\xb5c\x87S~Y\x96\xb9\xe3\x8e\xca\xf9\x0e\x8dmg\xaaT\xb1\x12\xe5+\xe4\xbc\x97\xfc&\xd2o\x1f$u\x1fg\x17\xa1*\xa9\x9d\xc0b\xdb\x16>&\xe6<\x9f\xcf\xf8\x9cY\xb3g\xe7\xf0\xd3o\xdb\xb6-o\x8c\x1cI\x93\xa6Mr(\x06\xb7\xd8o\xf8\xda\xa4\x8d,\xcb\x84\x04\x87\xf0\xf0\xc3\x8f0q\xe2DF\x8f\x1aEXH\xceQbfV&\xdf\x7f\xf7=O<\xfe8\xdf\xcd\xfb\x8esg\xce\xda\x8b/o\xe1/@\xa8m~\xe9\xd2\xa5\x8c\x1f?\xde)Uo\xd0\xf3\xc0\x83\x0f\xd0\xb2e\xcb<\\\xb5\x8b\x1b\xf5y\x0c\x06\x03-Z\xb4`\xdc\xb8q|\xfc\xf1\xc7\xf4\xbd\xf7^\xd7\x8cNX\xadV\xf6\xee\xdd\xcb\x90!Cx\xfd\xf5\xd7\xf9\xf5\xd7_\xb9\x1c\x1f_,u\xa9 \x14\xbf\xb4tA\x08\xc1\x85\x8b\x97\xf8\xea\xab\xafy\xfd\xf5\xd7\x199r\xa4k\x96\x1c\xf8\xf8\xf80h\xd0 \xd6\xaf_\xcf\xeb\xaf\xbfF\x93&\x8dU\x1b\xb8PJ\xa4`%-\x86N\x86\xcbp\xd9\xcf\xd7O[u{\xfdH\x92\x84\xa2(\xfc\xf1\xfbF\xa7\xe3\xc1A\xc14i\xdc\xc8aIz\xde\xe8t:\x06\xdc7\xc0\xf5\xb0=*\xa7\x07M\xd8\xbb\x1c\xb3\xfd\xadF\x80\xc9K\xa8\xdd8g\xce\x9ca\xc2\x84\x8f\xf8x\xca\x14\xd7$Z\xb7i\xc3\x88\x11#\xa8[\xa7N\xc1\x84?\xc2\xcd\xd3\x14\x1cI\x920\xe8\xf54j\xdc\x98\xd7G\xbc\xce7\xdf}K\xdb\xb6m\xdd. \xfc\xef\xd0!^z\xe9%\xde\x1e3\x865k\xd7\xe69O\xa7\xde\x91\x84"\x04\x1b7nd\xd2\xc4I\xaeYh\xda\xb4)\xfd\xfb\xf7/\xd6U\xe0\xd7Jhh(\xdd{tg\xda\xd4\xcf\xf8\xf4\xd3O\xa9Y\xb3f\xbe#\x95\xff\xfd\xef\x7f\xbc\xf2\xf2+\x8c\x1d;\x96\x1d\xdb\xb7\xe7P\xf0%A\x89*\x00\x01\xec\xdb\xb7\x8f\xc7\x1e}\x94\x91#G\xb2~\xfd\x06\x84\xc8\xbbrw\xe8p\x17\xe3\xc7\x8fg\xca\x94)\xdcu\xd7]\x04\x07\x07\xa3\xb3/\xc0(\xea&\xea\x06[\xd8p\xa18\r\xc9eY\xa6l\xd9\xb2\x9a\xed\xcf\xa5\x98\xaf\xa1]\n!\x88\x8b\x8b#\xcb\xe4\\Y\xbc\x8c^D\\\xc3\x82\x18Y\x96\tv\xd3\x8b;w\xee\\>\xa6\x84\xdb\x85\xdc\xcb@\x15\xa5\xb9\xa7\xdf(V\xab\xc2\xa9\xd3\xa7\x18?~<_\x7f\xfd\x95k2\xed\xda\xb5\xe3\xb5\xd7^\xa5\xe1\x9dw^kg^\xe5\x06n]\x96e\x02\x03\x83\xe8\xd6\xad\x07\xf3\xe6\xcd\xe3\x83\xf7?\xa0Q\xa3F\xae\xd9\x00\xf8\xf1\xc7\x05<\xff\xdcs|\xf6\xd9g\x9c={\xd6)\xa6\x12B\xed\rgff`2\x999r\xe4(s\xe7\xce\xe5\xf4\x19g\xf3R\xbdz\xf5\x183f\x0c\xd1\xd1\xd1N\xc7K\x16\xf7\x05h\xd0\x1b(_\xbe\x02\xcf\xbf\xf0<\xf3\xe7\xcfg\xe8\xd0\xa1\xd4\xa9\xa3F\x1c\xcd\x8d3g\xcf0\xe7\xcb9<\xf1\xf8\xe3\xcc\xff\xe1\x07L%\xac\x04JV\x01(\n+W\xaed\xc3\x86\ry\xf6\x1a\x00\xa2\xcb\x95\xe3\xbd\xf7\xc61y\xf2$^x\xe1y\xca\x95+W:\xec\xd6\x92*\xa4\xcdfK\x8eM6|}}\xed\x1e\x00N\\\xe3m_\xbe\xec\xdeLc?M\x01w\xaa\x92\xdd\x94W\xc1z\x93\xb7\x03j\xd9\xe4,!\x90T\x1bP\x91\x11\x1b\x1b\xc3\xecY\xb3\x99;wn\x8e\xe8\x94\x1d;v\xe4\xcd7\xdf\xa4i\xd3\xa6N\xdeI\xd7\x84\xbb\x87\xbaFt\xb2\xcc\x1dw\xdc\xc1\xf3/<\xcf\xc7\x1fO\xe1\xb9\xe7\x9es\xcd\x82\x10\x82s\xe7\xce\xf1\xf6\xdbo3\xe2\x8d7X\xfb\xdbo$%%\xa1X\x15\x8e\x1c;\xca\xcaU+Y\xb0\xe0G~Z\xbe\x8co\xbe\xf9\x9a\x9f\x7f\xfe\xd9\xa9\xf3Q\xb9Je\x9e{\xee9\xaaV\xadZ:\xda\xb6\x9d\xbc\xef\xc5\xa07\xd0\xb8qc\xde\x1e3\x86)S\xa6\xf0\xe4\x93O\xbafqB\x08\xc1\xb1\xe3\xc7\xd9\xb0~=\xf1\x97.\x15i\xe7"?JT\x01H\x9a/r^C\'___:u\xea\xc4Ws\xe72l\xd8p\x1a7j\x84\xbe\xd4\xc5\xe6\xce\xe9\xe0e{6\xc9\xb5\xf7\x7f\x8dX\x15k\x0e\xc5b\xc3\xdeF\xf2\xae\x9f\xd7\x8d\xa2((\x8a:\xb2\xb9]F\tn\x9fR\x92\xb40\xc9nSo\x88\xb8\xd8X&N\x9c\xc8\x147f\x9fV\xadZ\xf1\xdak\xafQ\xa7n\x9dB7\xd3]\xef\xfb\xf4\xf1\xf1\xa6m\xdbv\x8cy\xe7\x1d\xe6}\xf7\x1d-Z\xb4p\xdb~\x97,^\xcc\xeb\xaf\xbd\xc6\xb4i\xd3X\xb8h\x11#\xdfx\x83\xfb\xef\xbf\x9fg\x9fy\x86\x87\x1f~\x84\x8f?\xfe\xd8)\xbf^\xaf\xa7G\xf7\x1et\xb8\xab\xc3\r\x07y+)\x82\x83\x83\xb9\xfb\xee\xbb\x19\xff\xe1\x87\xcc\x9f?\x9f\xfa\xf5\xeb\xe79\x87a\xb1XP\xac\xd6\x92\xb0[\xd8\xb91\xe9t\x83\xa8B\xd2\xfd\xc3\xebt::u\xea\xc4\xc4\x89\x13\xf9~\xfe|:u\xeaD@@\xc0\r\x0bT\x1b\xd7W\xfd\x1dq<\x83\xa2\xce\x14\xba<\x8a\x9a#\xbf]q\xf3A\x90cb\xd2F\xf6a\xf7\xe9\xae(n\xce\xe3\xae\xf2\x99Lf\xe2\xe2\xe2X\xbfa\x03K\x97.e\xd9\xf2e\xec\xd8\xbe\x9d\xcb\x97/c\xb1\xaa=T\xdb\xb3\x15\x88\x02f\xcb\x9bB9\xc9u!\xc4\xf5\x0bL\'\x1cNaQ\xac\x9c={\x96\xf7?\xf8\xc0\xc9\x03\xc6\xa6\xd4\xdb\xb5m\xcb\xd0aC\xa9[\xcfy\x85\xef\xb5\xa0\xf6,\xdd\x7f\xf7\xfa{\xd8\xea\xba\x81\xc8\x882<\xf8\xc0\x03\xcc\x9a5\x8b\x17_|\x91\x9a5k\xb8f\xe4\xf8\xf1\x13L\x980\x81\x97^z\x91\x15+V`\xd6\xcc\x98V7#\x99\xee\xdd\xbb\xf3\xc8\xa3\x8f\xe4\xb9\xb9KiG\xd2&\x89#\xca\x94a\xd0\xa0A,\\\xb8\x90\xb7\xdez\x8b\xc6M\x9a\xb8f\xcd\xe6\xba\xdfC\xe1P8\xd2\xf4\x06\x10\xb9\xc4\xbe\xef\xd9\xab\'\x1f}\xf4\x11O?\xf5\x14e""r\xda\xd1o\x107\x97\xbcF\xb2\xcf \x90\x10\x9a)\xc8~L\x08LYY\x98\xb2L\xd7\xdd\x80\xd1l\xb0\xa1\xa1a\xae\x87\xaf\x0f7\xf3+:\xbd\xce\xe9\xfeRSSY\xb5j\x15\x8f>\xf2(\xdd\xbbu\xe3\x81\x07\x1e`\xd0\xc0A\xb4m\xd7\x8e\x91#G\xb2m\xdb\x9f\xf6`i\x05Vm\x05\xcc\x967\x85r\x92kG\x02Y\x96\xb4g\xbd\xc1{p\xf8z\\l\x1c3f\xcc\xe0\x8b/\xbep\x9a\x0c\x14\x02:t\xe8\xc0\x9bo\x8d\xa2U\xabV7Tw\xd4\x0b\xde\xe0=\xe7\x86\xaczM5h\xd0\x80\xb7\xc7\xbc\xcdG\x1fM\xa4s\xa7\x9c\x1b\xf7dee\xe5\xeb\xcd\xd7\xa6m\x1b\x9e\x1d\xf2,e\xca\x94qM\xbai\xd1\xe9t\xd4\xacU\x8bW_}\x95\xc7\xb4\r\x90\xdc"\xec\xff\x94\x08\x85+U\xaf\x9b\x9c\x95\xf4\x8e;\xee\xa0n\xdd\xbax\xb9\xf1:(mH\x92j\x07t\\\xe8!\x84\xe0\xc2\x85\x0b$\\Ip\xca{\xad\xc8\xb2L\x99H\xf7\r\xc3^m\nX\x7f,n\xec\xfd\xa1\xe1\xa1\xf6\xdf322\x98?\x7f>\xaf\xbd\xf6*\x1b5\xaf#\x9dNgW\xbe\xdf~\xfb-o\x8e\x1c\xc9\xbau\xeb\x10J\x81\xc5\xffM\x8d$\xd4\x8dQ\xa4\x1b\x12\xc4\xce\xc4\xc4\xc40\xc9\x8d\xd9G\x92$\xda\xb5k\xc7\x9bo\xbeI\x9d\xba\xb5oP\xf8\xab\xad*\xcf3\xe4\x99\x987\x92mNI\x82\xd0\xe0P\xbau\xed\xc6\x9c/\xe70n\xdc8\xa2\xa2\xa2\xaei\x84Q.\xba\x1c\xe5\xca\x15|g\xaf\x9b\x05\t\xf0\xf7\xf7\'**\xca5)\x1b\xc9\xfeO\x89PJ\x14@N\xd4\nVjo\xcf\t\xdb\xa6\xd1\x8eu^\x08AfF\x06\x8aU\xb9\xe1\x17\xec\xeb\xe3K\x8b\x16-\x9c\x8eefe\x12\x1b\x1b\x03\x14l\x0e\xd8b\xb6\xb0u\xcb\x16\xd7\xc3t\xbe\xa7\x13\x06\xbd\x1e\x81\xe0\xcf?\xffd\xc2\x84\t\x9c9s\x06ooo\x06\x0e\x1c\xc8\xa8\xb7F\xf1\xca\xab\xaf\xd0\xaauk\x00\xfe\xf9\xe7\x1f~\\\xb0\x80\xf3\xe7\xcf\xe5j\x9a*<\x8a\xfa\xfc\xce\x08E\xc9a&S\xfd\xffE\x81\xca\xd8Fn\x82[Q\x14\xce\x9e;\xc7\x07\x1f~\xc0\x97_~\xe9\x9aL\xcb\x96-yU\xb3\xf9\xe7v\x8ekA\xd8k^.\xe7\xba\x86g\xca\r\t\tI\x96\xf02zQ\xb1bE^}\xf5U\xbe\x9d7\x8f\xfb\x06\xdc\xe7vn\xc0\x1d\xbf\xfd\xf6[\x9e\x9b\xb5\xdf\xcc\x14\xf4=\x16,W\xe1S*$\xac\xdb\x87\x17 $\xb7)\xa5\x0e\t\x81\x9f\x9f\x1fU*Wq:\x9e\x98\x94Djj\xea5\xcd\xf2\x0bE\x8d\xe5\xee\x88\xb7\xb77]\xbatq:\x96\x94\x98\xc4_\x7f\xfdE\x96)\xcb\xad\x1d\xdf\x95\xf8\xcb\xf1\xfc\xb9\xedO\xd7\xc3\x94+W\x1e$\x89\xa4\xa4dV\xaeZi\x8f\xc1\xfe\xfc\xf3\xcf\xf3\xe1\x87\x1f2f\xcc\x18\xde\x1d;\x96q\xe3\xd4\x05\xe3B\x08v\xee\xda\xc5\x993g\xc8en\xba\x10\xc9\xff\xb9\n\x15\xc9\xfd\xae=\x92\xa4m\x8fX@r\xeb\xfd\xc6\xc6\xc52\xe3\xf3\xcf\x99\xf3\xc5\x9c\x1c>\xe0\xadZ\xb7f\xc4\x88\x114i\xd2\xb8\xd0\'|\x8b\xb3\x1c}}|\xb8\xbbcG\x9e{\xee\xf9\xbc{\xbe\x0e$&&\x92\x9e\x91w\xa8\x97\x9b\x95\xdc\xea\x82+\x05\xcbU\xf8\x94\x02\x05 \xdc\x0e\xaf\x15\x8a\xd6\xf5\xaep\x91\x88\x88\x88\xa0J\x15g\x05\xb0w\xef^\xce\x9c9\x83pczqD\x00\xa7N\x9ff\xdb\xb6m\xec\xdb\xbf\x8f\xa4\xa4D,\x16\x8b\xbd\xf7\xe0\xed\xedM\x8d\x1a5\x9c&\xcc\x15E\xe1\xc0\xfe\x03\x9c<q\xb2@\xc5\xf4\xd7_\x7f\xe7\x08$g\xf42R\xad\xba\x1a\xe7\xde\xcf\xd7\x97W_y\x95=\xbb\xf7\xb0x\xf1b^}\xf5U5R\xa8^\r#Q\xa3F\r\xf4\x9aGCLL\x0c\x99YY\x85Xk\x9d\x9f@h^G\x05y\xaeBE\x92r\xb8\xca\xaa#;\xd5\xa3\xab`\xb8<\x8b\xf6<\xe7\xce\x9ee\xc2\xf8\tn\xcd>m\xda\xb4\xe1\x8d\x11#h\xd0\xb0a\x0e7\xd0\x1b\xa1\xd0^\xcf\xb5\xa0\x95U\x93\xc6\x8d\xa9\\\xb9\xb2k\xaa{\xb4\x80t\xb7*y)\x01[]/)\nZ\xab\x8b\x0c\x81\x16\x9f\xd6\x05\xa9\xa4*\xf0uR&"\x82\n\x15+8\xbd\xec\x0b\x17.\xb0u\xebV\x92\x92\x93\xf3|\xc9f\xb3\x99\x1f\x17,\xe0\x81\xfb\xefg\xc0\x80\x01\xf4\xed\xdb\x97)S\xa6\xb0}\xc7v\xb2\xb2\xb20\x1a\x8dT\xa9R\x85\xfa\xf5\x1b8}o\xcd\x9a5\xac]\xbb\x96\xb4\xd4\xd4\x1c\x82\xc7\x86\xd0BK\xff\xf6\xdbZ\x12\x12\x9c\xe7#\x1e\x1c<\xd8\x1e\xdf\xdd`0P\xa1B\x05\x1a4l@\xbf~\xfd\x88\x8e\x8e\xd6\x9eE\xed\x15\xc7\x9c\x8f\xc1\xa2\xf5Z+\x94/\xaf\xee&\x96\xcb5\x0b\x8a\xd5\xaa \x14\x81\xd5\xaa\x90\x9c\x94Lll,\x87\x0e\xfd\xc7\x8e\x1d\x7f\xb1u\xebV\x0e\xfd\xfb\x1fqq\xb1$\xa7\xa4`\xb6Xn\xf8z\xf9!\x84@\xb8\x08z\xa1-\xe7\xcf\x7f\xbd\x84zo9^\xb3\x10\x9c?\x1f\xc3\x84\x8f>\xca\x11\xef\x06\xa0u\xabV\xbc\xf6\xfak4j\xdc(W\x8f\xb8\x9b\x11???z\xf7\xee\xe3z\xd8-\r\xea\xd7w\xb3q\xcd\xed\x81T\xc2\x9b \x95\xa8\x02\x10\xb8k1*\x02\xb5gp3 I\x12\x01\xc1A4i\xda4\x87\x1b\xdb\xf8\xf1\xe3\xf9k\xc7\x0eu\xc8/\xa9\xeb\x05\x9c\x11\x1c;z\x84\xad[\xb7\x12\x1b\x17\xc7\xa9S\xa7\xd8\xbau+3g\xcd$5%U\xed\x19IP\xbbvm\xda\xb5k\xe7TY\x14Ea\xc2\x87\x1f\xf2\xc7\xa6?\xc8\xc8eW4\xab\xd9\xc2\x8e\xed\xdb\xf9\xfd\xf7\xdf\x9d\x94\x90$A\x87\xbb:\x10\xe0\x1f\xe0\x94_MS\xafa\xcb\x9f\x9c\x9c\xcc\xb2e\xcb\xeci\xed;t\xa0v\xed\xda7,\x8f\x15a\xe5\xcc\xd93\xfc\xb1\xe9\x0f>\xfb\xf43\x1e}\xec1z>\xdd\x84\xe1_\xf4`\xe4w}\xe9\xfd\\\x13\xbaw\xed\xc6\xecY3\xd9\xb7w\xaf;\'\xa6BE\x92$\xb7\xa3Q\x1by$9uWl\x96$!\x04\xe7\xcf\x9fg\xfa\xf4i|\xf1\xc5\x179z\xf7\xed\xdb\xb7\xe7\x8d\x91#i\xd4\xa8Q\x01\x14\xcc\xcd\x85,\xcbt\xeat\x0f\xcd\x9b7wM\xca\xc1\xbd\xf7\xdeK@@\xcezx\xab\x90W\xb5\x11\x0e\xed\xac$(Q\x05\xa0\xf6/\xdd\tE\xb5dJ\xc0\x08p\xdd\xc8H\xb4o\xd7\x8e\xfa\r\x9c{\xe9\x00\x13&L\xe0\xbb\xef\xbf\xe3\xe2\xc5\x8b(V+B\xdb\xfb\xd4b\xb1p\xe0\xc0\xbfL\x9b6\x9d\x8d\x1b6\xd8\xf3K\x92D\x97\xce]\xa8_\xbf\x1e:\x9d\x0e\t\x89\xc0\xc0 z\xf7\xeeM\xf9\xf2\xce\x01\xdd.]\xbe\xcc\xd8w\xc62\xe7\x8b/\xb8p\xe1\x02V\xabj:R\x84 ==\x9d\x9f\x7f\xfe\x89\xcf\xa6N\xe5\xc8\x91#N\xdf\xeb\xd1\xa3\'\xcd\x9b5\xcds\xd1\x8d$I\xa4\xa4\xa4\xf0\xcd7\xdf\xf2\xcd7\xdf\x80\xe6\x9d5h\xd0 \xc2\xc2\xae\xdf5U\x08A\xdc\x85\x0b,\xfcq!\xa3F\x8db\xf8\x98!lN\xfe\x9a\x8a\x0f\xc7s\xf7\xf35\xa8{wE\xea\xb4\xabH\xc7g\xabS\xf7E\xc1O\x07\xa6\xf1\xe6\xa8\x91\xac^\xb3\x1a\xb3\xc5\\d\xb5"\xb7\xee\x86\xad\x97&\xe51\'\xa5(\n\x99Y\x99\x1c?~\x9c\x9d\xff\xec\xe4|\xccy\xce\x9d;\xc7\xf8\xf1\x13r,z\x02h\xd1\xb2%\xc3_\x1eN\xbd\xfa\xf5\x8aN\x00\xe4\xf6@\xf9P\x18\xf7#I\x12u\xea\xd4\xe6\xddw\xc7\xd2\xa2E\x0b\xb7\xbd\\___\x1e~\xe4az\xf4\xe8\xee6\xc6\xd0\xadB\xce\'\xcfFr\xe8p\x95\x04%\xbe)\xfc\x84\t\x13xw\xdc8\xbby\xc1\xc6\xb0a\xc3\x98<yr\xa1o\x01W\x94(\x8a\xc2\xcaU\xabx\xfe\xb9\xe7\xb8p\xe1\x82SZpp0\r\xea\xd7\xe7\xfe\x07\x1f\xa0LD\x19\x14E\xe1\xd4\xe9\xd3\xfc0\x7f>\xff\xfe\xfb\xaf\xd6\x03\x94\x00ATT\x143f\xcc\xa0O\xdf>Z\xe5Q\xff5\x99L|\xfe\xf9\xe7\x8c\x181\xc2\xe9\xdch\x9b[\xd7\xacY\x93\xc7\x9fx\x94\xb2\x91\xd1ddd\xb0z\xf5j\xd6\xad[\x97#\xe2ghh(\x13\'N\xe4\xb1\xc7\x1e\xcb\xd3\xf6\x9a\x92\x92\xcc\xa2E\xffc\xcc\x981\\\xbat\x89\x8a\x95*2v\xecX\x1e\x1a\xfcP\x9e+\x1c\xdd"@A!3#\x8b\xdd{v1}\xda4\xfeK\xd8\xc6\x1d\xcd}\xb9\xa3q\x08V\xab\x15S\xa6\xd5\xa1\x9b\xadu\xa3%\t\x83AG\xec\x914.m\xf6\xe5\x9d7>\xa0S\xa7\xce\x18\xbc\xae\xf1\xfa\x05 6&\x96\xaa\xd5\xaa\xbal\n\x1f\xcaW_}M\x8f\x9e=\x90\x91\xd50\xe3.dfe\xb2k\xe7.~\xfc\xf1G6m\xdaDRR\x12e\xcaD I2\xfb\xf7\xef\xcf\xd1\xf3\xef\xd0\xa1\x03\xcf\xbf\xf0<\xcd\x9a5+\xb2\xc6/\x84 (0\x88\x107\xf1\x9f\x8a\x93,\x93\x89-\x9b7\xb3l\xe9R\x0e\xfe\xfb/\x07\x0e\x1e (0\x88z\xf5\xea\xd2\xaauk\xbav\xed\x9aoD\xdb\xa2$  @\xdd\xac\xa9\x88\x10B\xb0|\xf9r\x06\x0e\x1c\xe8\x9a\xc4\xc0\x81\x03\x994y2\x15+TpM*8"\x1f\r\x93\x0f\xa5V\x01\x0c\x1d\xfa\x12\x93\xa7|\x8c\xe1&R\x00h\xe6\x92O>\xf9\x84\xf7\xdf\x7f\xdf5\xa9\xc0L\x9d:\x95\xc7\x1e{\x8c\x80\x80\x80\x1c\xef7&&\x86\xc9\x93\'3}\xfat\x87\xa3\xd7\xc6\xc4\x89\x13y\xf2\xa9\xa74\xe1\xe0.n\xbb\x1a\x1b\x7f\xe1\xa2\x85\xbc\xf0\xfc\x0b\x00\xdcqG%\xde~{\x0c\x03\x07\x0c\xc0\xff:\x87\xeb1\xb11\xac\\\xf5+Sg~B\xd5N&\xc2\xab\xfb\xe3\xed\xaf\xc7\x9c\xa5\xd8\x07\xca\xb6\x7fUU\x08\x08\x19\x10\xe8\x8d2\x17\x0e\'\x91\xb2\xbd\x02\x9fL\xf9\x8c;\xeflxc5\xdf\r\xb1\xb1\xb1T\xad\xea\xaa\x00\xc2\xf8\xfa\xeb\xaf\xe9\xd1\xa3\'\x12\xc2i%\xba\x10\x02\xb3\xd9\xcc\xda\xb5k\x192d\x08\x17/^\xb4\xa7\xe5\xc6=w\xdf\xcd\xab\xaf\xbdJ\xdd\xbaus\xb8\x9c\x16&\xa5E\x01\xa0\xddKBB\x02\xe7\xce\x9e#6.\x96\x8c\xcc\x0c\xa2\xa3\xa3\x89\x8c\x8c,\xf1\x0e^q(\x80e\xcb\x973(W\x050\x89\x8a\x15*\xba&\x15\x1b%j\x02\x02\xcd\xcc\xe3\xb6!H\xe41\xe2.\xb5\xf8\xfb\xfb\xf3\xc4\x13O\xf0\xc2\x0b/\x10\x18\x18\xe4\x9a\x9c\'\x11\x11\x11|\xfa\xe9\xa7<\xf4\xd0C\xea\\\x82\xc8\xb9\xd8***\x8aW_}\x95\x17^x\xe1\x9a\xc3\xe5\x86\x86\x852q\xe2$\x1e\x7f\xfcqBBB\xd4s\xbb)\xe3+W\xaf2\xf7\xab\xb9\x8cy{\x0c\x00e\xa3\xa2\x183\xe6\x1d\x06\r\x1c\x84\x9f\xff\xb5/\xd5\x17Bp\xe8\xd0!>\x9a\xf0\x113\x17\x8f\xa3\xc9\xd3\x06\xca7\tB\xef%k\xc2?\x1bY}\xf3\xf6\xbfmkp-YV\xc2\xab\xfa#\xd5:\xc1\xb4i\xd3\xb8z5\xd1\xe9{E\x89$\x81$\xe3\xb6\xf7\x7f\xe2\xc4\t\xbe\xf8\xe2\x8b|\x85\xbf$I\xb4o\xdf\x9ea\xc3\x87S\xabv\xed"\x15\xfeE\xc1\x8d\xdc\xad$I\x84\x87\x87\xd1\xa8\xd1\x9d\xb4\xef\xd0\x9ef\xcd\x9aQ\xae\\\xb9\x12\x17\xfe\xc5E\xceZ\xe3B\t\xd6\x85\x12W\x00\x12\xb9\xcc\x82K\x05)\xb9\xd2\x87,\xcbT\xaaT\x89\xb1c\xc72y\xf2d\xfa\xf7\xefOh>\xf6rI\x92\xe8\xd5\xab\x17S\xa6L\xe1\xd1\xc7\x1e%$$Du;tS.\xb2,S\xb1\xa2j\x8a\xf9l\xeaT\xfa\xf5\xebGXx\xde\xe7\x07\xb8\xfb\xee\xbb\xf9\xe0\xfd\x0fx\xe6\xd9\xa7\t\x0f\x0f\xb7\x17\xadk\xd9\'\'%1\xff\xfb\xf9\x8c\xffp<\t\t\t\xd4\xaf_\x9f\xf7\xc6\x8d\xe3\xc1\x07\x1e\xc0\xd7\xdf\x0f\xc9aq^Al\xc5\x16\xab\x85\xc3G\x0e3i\xcaG\xec\xbc\xbc\x94\xd6OD\xe2\x13\xa4\xc7\x92!P\x84\xa29\xfcj\x01\xe7\xb4\x05W\x0eF U\xf4Hj\x9c%\x19\x99\xf2uB\xf8\xeb\xc4ol\xd8\xb0A\xfbN!\xe3\xa6E\x08!\x10n\x16=\x98L&v\xfc\xb5\x83\xb5k\xd6\xba&\xe5\xa0\\\xf9r\x0c\x1f>\x8c\x06\r\x1b\xe4(\xf3"\xa3\x10/s\xe3\xa7R\xbd\xfdr\x8e6K\x0f9\xdfp\xd1#\xa1\xf50J\x087\xd5\xbd\x98Q\xc7\xf9\xaeG\x0b$\\J3\xe1\xe1\xe1<\xfe\xf8cL\x980\x81\xe9\xd3\xa61v\xecX\xee\xeax\x17\x06\x83\x01\x83\xc1\x80^\xa7\xee\xe85z\xf4h~\x98\xff\x03\x93&O\xe2\xa1\x87\x1e"8(\xd8my\xb8\x12\x1e\x1e\xce\xe0\x07\x1fd\xd2\xa4I\xcc\x9a9\x8bQ\xa3GS\xbdz\ruW1\xbd\x1eo\xa37\x11\x11\x11\xbc\xf0\xe2\x8b|\xf3\xcd7|\xfc\xf1\xc7<\xf1\xe4\x93\x04\xe51*\xb1Z\xad\xacX\xb5\x8aW^y\x85\xc4\xc4D\xee\xb9\xe7\x1e&N\x9c\xc8C\x0f=\x84\xd1\xdb;W\xa5\xe1\x0e\x9by\xe4\x9f\x7f\xfe\xe1\xdd\xf7\xc6\x12\x17\xb8\x85\xc6\xbd\xa2\xb1d)(&\xc5>\xc1o\x7fR\xed\x94Bu\x9f\xb1\xa7\xa91\x96\xd4?\x14\xab\x82o\xb0\x81\xa8\x86^\xfc\xfe\xc7\xef\\M\xbcZ\xf8\xe2$\x873\x8ez/\xea#;_\xcdb1s\xf1\xc2EM\x91\xe5\x8d\xaf\x8f/QQ\xd1\xc5[\xaf\x8b\xf1R\xb7\x02\x85^\x97\nD.N0\xc5D\xc9+\x00I\xb8-z{\xef\xef&F\xaf\xd7S\xb5jU\xfa\xf5\xeb\xc7+\xaf\xbc\xcc\x82\x1f\x16p\xf2\xe4IN\x9c8\xc1\xb1\xe3\'\xd8\xbcy\x13#F\xbcN\xff\xfb\xfaS\xb3FM$IR\xdd\x07\xdd\x94\x87;t:\x99*U\xaa\xd0\xa7O\x1f^\x7f\xed5\xd6\xad\xfb\x8d\x13\'Np\xe2\xd81\xfe;\xf4\x1f\xbbv\xed\xe2\xbd\xf7\xc6\xf1\xc0\x03\x0f\xd0\xa0A\x03\xbc\xb4\x89\xdb\xdc\x84\xd0\xd93g\x99=k6hB>44\x14\x8b\xd5\xc2_\x7f\xff\xc5\xc6\x8d\x1bY\xb7~\x1d\xeb7\xac\xe7\xcf?\xb7\xdbW\x0c\xe7\x86\xc9db\xc3\x86\xf5\xbc3v\x0c\xe9\x95\xf7Q\xb9I\x08V\x8b\xea\xfd\x94\xe3\xf1\x1c\x84\xab\x1b9\xabI_\t\x10\x98\xb2\xac\x94\xaf\x1d\xc0\xb6\xbf6\xf3\xef\xc1\x7fs\r\x95]x\xd8F\xa8R\xae\xe5V\x10\x92\x93\x93III-\xe6\x1epq^\xcb\x83;\xf2\xaf3\xa2DU@1+\x007\x0f\x9aK\xa8]\xd5\xd4\xe0\xbe\x02\xbb\xcb_Z\x91$\t\xa3\xd1H@@ \x91\x91\x91DGGS\xae\\9*T\xacH\xb9r\xe5\x08\x08\x08t\xf2\xa8\x91\xaei/{5\xa7\xc1` ((\x88\xf2\xe5\xcbS\xae\\9*V\xacH\xa5J\x95(W\xae\x1c!\xc1!9\\=s\xf6\xe0U\xf3\xcb\x85K\x178y\xf2\x84zD\x08~\xfe\xf9g\x9ez\xf2)\x1e\xb8\xff\x01\x06\x0f\x1e\xcc#\x0f?\xc2\xc3\x0f=\xcc\x83\x0f<\xc0\x91#G\xb4\rJ\xb2\xdf\x85\xdayW0[\xccl\xdb\xbe\x8d\xe1o\xbc\x80O\xa3\x18*5\x08\xc0l\xb2`v\x0cq\xe08\xbf!T\xe1n;\x87\xda\xf5W\x93l\xf3@\x92$\xd4\xfb\x16\x82\x80H/\xb2\x02\xcer\xf8\xd0a\xe7]\xa7\x8a\x00I\xb2\xd5E\xed\xfa\xce\xa9.\x7f\xe7\x8e\x9f\x9f\x1f\x81\x81\x01\xc5\xec\xda\\\x9c\xd7\xf2p}\xd8:7%C1+\x00\xe7\x07uh\xe7n\xc9--gC\xbc\xf9(\x8a\'\xb8\x91r\x11B\x8dC\xe4\xe8\xbej2\x99\x88\x8f\x8f\xe7\xd2\xa5K\xc4\xc7\xc7\xdb?\xe7\xce\xab\xdbH\xca\xb2\xde\xe9I\xd4\x8e\xb2\xc4\xbf\x07\x0f2\xe5\xe3\x8f\xa8\xdaU\xa2\\\x03?2\xd3,\xa8=i[v\xdb\xde\t\xb6\xef:\xbcima\xb8M\'\x08-\xd9\x96CB"3\xd5L\xb5\x96\xa1\xec\xda\xb9\x9b\xf4\xf4\xa2\x8d!\xa3\xf65\xdc\xd7D\x9d^GTT\xd9\x02\x85\x89h\xd9\xb2%Q\xd1\x05\x8b\x8d\xe3\xe1\xd6Ar\xbb\xf83\x9b\xe2\xed\x10\xe4$\xff\x9a[\x0c\xb8\x13[\xd7\xd6\x13\xf6pc\xa8\x86\x89\xc8\xc8H\xfe\xb7\xe8\x7f,\xfe\xdfb\x16/\xd6>\xda\xef\xff\xb3\xfd\xad}*\xddQ)\xc7\xdc\x95\x10p\xee\xec9f\xcd\x9aM\x82\xdfA*5\x08\xc1\x94\xe9\xe8\x03/\xeco[U\xfe\xb6\xbf\x1d\x0c#v\xa1\xafU\x00I\x02Is\x0b\xd5\x14\x88\xb0\n"\xab\x06\xf1\xcb\xca\x9f\xc9\xcct\xbf\x02\xbaP\xc9\xa5\x8dzyy\xd1\xa2EKz\xf7\xee\x9d\xa7\xf25\x18t\x0c\x180\xe0\xa6\x1a\xb9z(<r\xaf\x19%o\xcd(Q\x05\xa0\xb6o\t\xe1\xa6\xf1\x94p\xb9\xb8\xa1\xd4\xddP\xa1"I\x12U\xaaT\xa1\xff}\xfd\xe9\xd7\xbf\x1f\xfd\xfai\x1f\xc7\xdf\x1d>\xd5\xaa\xaaA\xe4\x1cIMO\xe3\xd7_\x7fe\xd3\x7f?Q\xbf[$f\xabb\x17\xf7\xd9\x91]5\xb1/\xd0\xa4\xb9z\xdci\xc3\x15U\x038\xfc\xae\xa5;\x1c7\xfa\xc8\\\x95.r9>\xbe\xc8\xdf\x8dU\xa81\x8b\\\xaf"K2\xd5\xaaU\xe3\xc5\x17_\xe4\xd1G\x1eA\xe7f$P\xb5j\x15>x\xffCj\xd7\xb9\xf1\xf8\xfe\x1en>\xf2}\xe7\xd7\x12g\xbc\x08\xc8YcK\x02\xb7\x85$\x8a\xbca_\x1b%\xfb\xa2\x8a\x03\xc9\x1e\xf2 \xe7Gvs\xcc\x86\xd0VA\xef\xd9\xb5\x8b\x17G\xbfH\xe3~\x11H\xb2\x1a__\x15\xf8\xaapW\x97\x9c\xa9\xbd}\xb5\xde\x0bm\x14a{\xd7\x9ai\xc8\xf6\x11\x9a\xa6\xd0\x04\xbf:\x05\xa0^\xd7b\x11Tn\x04\'O\x9dr_}\n\x13\xfb-8_Hh\x13\xfdw\xdfs\x0f\xe3\xde{\x8f\xcfg\xcc\xa0{\xf7\xee\xf8\x07\xf8\x13\x1c\x1c\xccSO?\xc5\x07\x1f|@\xef\xbe}n\xe9P\x07\x1e\xae\x1f\xe9\xda"\x8d\x17:\xa5B\x01\xb8\x1b>\xab\xe2 \xe7q\x0f\xd7I\x11\xd62\t8\x7f\xfe<\xd3\xa6M\xe5\x9eG\xcb\xe3W\xc6\x80b\xd1\\#\x9dLyj/\xda\xa6\x10\x9c\xed\xa3j.{^\xcd\x04\x94\x9b\xe7\xb8P\xacDU\xf2\xd26\xdc)Z$\x9brts\xdc\xf63*:\x9a\xee\xdd\xbb\xf3\xd1G\x13\xd8\xb4i\x13\x9b7o\xe6\xf5\xd7G\xd0\xbcE\x0b\xbc\xbd\xbd\xf3\xef\t\x16\x019\xc7,\x1e\x8a\x1bw\xb2\xcd\x95\xfcs\x14\x1d%\xa8\x00\xb2+\xe7\xcdXM\x15-\x8e\xb7\xd5b%%-\x8d+W\xae\x10\x1b\x1bKLL\x8c\xfdgLl,qqq$&%\x93\x91\x91\x8e\xd0\x82\xb4\xd9\x9e\xb7 \xfe\xe3\xd7Gv\x89*\x8a\x82\xa2(X\x95\xec\xfd\x05\xf2"\xff\x1c91\x99\xcd\xacY\xbb\x86c)\x7fR\xb1N\x08\x98m\xb5Z[\xc7kW\x02\xd9\xffJ\xf6n5H\xc8x\xfb\xea\xc1\xa2\xf6\xf0m\x9e\xc1\xb6\x86!4\xe1\xeb(\xd0dYFo\xf4\xbf\xbe\x1bvK\xee\'\xb2\x87\x7f\xc8\xa31K\x80\xde`\xc0? \x80\x90\x90\x10\xfc\x03\xfc1\x1a\xbd\n$\x00\x8a\x8e\x92\xbc\xb6\x87\x82\xe0\xdc\t*~JP\x01\xd8\x9b\xb7\xdb\x10\xbc\x12y\x87\xe6-i\xac\x16\x0bG\x8f\x1ea\xdd\x86u\xcc\xfc|\x06#F\xbcA\xe3\x06M\xa9^\xb7*5\xebW\xa3V\x83\xaaT\xaf^\x8d\xfb\xfa\x0f\xe2\xa3\xf1\x13\xf8\xea\xab\xafY\xb7n\x1dG\x8f\x1e%99\xefM\xb2o\x14\xdb\xc8I(\x82c\xc7\x8e\xf3\xfb\x1f\xbf\xf3\xc7\x1f\x9b\xd8\xbd{7\xf1\xf1\xf1(\x8a\x92\xab`r\x7f4w\x84\x80c\xc7\x8e\xf2\xd3O\xcb\xa9\xda:\x10/?=\xc2\xaa\xbd?I\x13\xf4\xae_p\xd9\xeb\xc7\xcbW&\xe9b&^~j`:[\x9a\x9aU\xfbK\x00";"\xa7\x90@\xc9\xddS\xf8:\xc8\xfdDR\x017\x84q5\x8dy\xf0P JP\xce\xe5_\xab\x8b\x1a\xe1~\x1eD`s\x15,~\xf2z\x1dB\x08\xd2\xd2\xd3X\xb4h\x11#F\x8e`\xe8\x98\xa7\xf8\xf9\xdf\x8fI\xa9\xf5;w\x8d\r\xa6\xff\xc7\x95\xb9\xef\x93*\x0c\xfc\xb4*\x03&U\xa1\xc2}W8\x15\xb1\x92Y+\xdf\xe3\x89W\xeeg\xc4\x1b\xaf\xf3\xfe{\xef\xf3\xcb/\xbf\x90\x9c\x9cR\xf8q\xe0\x85Z\xa1\x84\x10l\xdd\xb6\x8dQ\xa3\xdf\xe2\xd97\x1e\xe2\xb9\xb7\x1f\xe4\xb5Q/3\xfa\xed\xd1|\xf3\xcd7\xc4\xc4\xc4\x14hD\x90\x1f\xe9\x19\xe9\xfc\xfe\xfb\x1f\x9cJ\xfb\x9bru\x030eX\xb5p\x0e\x9a\xe8\xd6\x84\xbd\xedJ\x92$i\xd6\x1d\tE\xa8=\xf9\xcc\x143:\xbd\x84\xc5\xa4\xad\x02V\x8d\xfdZ7 \xfb\x1c6\xd9*\x00+\x02\xab\xd9Zx\xa3\xa8|\x8a"{W\xb7\xdc3\x96Pu\xcd\x95\xd2v?\x1er"\x84m\x1e\xacd(q\x05`\x13\x129\x10\xb9h\x86B\xc4\xdde\xc9\xa7\xe1dfe\xf1\xdd\xbcyL\xfev4\x86\xa6G\xe8\xf6V\x14w\xf6,O`\x84\x1f~\xfe\x06\x8c\x92\x17:\xab\x1e\xc9\xaa\xc7\xe0\xa3# LOP\xa4\x9e\xbb\x9e\xacD\xaf7+\xe3\xdd\xfa8\xdb.\xfc\xc8\xa8O_\xe0\xfe\x07\x06\xb2b\xc5\n.^\xbc\x84\xd5\xbe\x05\xa4\xa3\xb8\xbc\x0e\xb4I\xa5\x98\x98X\xe6\xcd\xfb\x96\xabe\xfe\xa1\xf7;\x15\xe9;\xa6\n\xb5\x1eN\xe7R\xe4f\xa6/\x7f\x87\xe1/\xbf\xcc\x8e\x1d\x7f\x91\x9e\x9e\xa1}\xf1\xfa\xaey\xf1\xc2Ef\xce\x98\xc5\x9d\xbd\xa2T\x7f\x7fY\x01I\xc9~\xb1\x8e\xb6s\xed\xd1\xd4y]\x81N\x02\x9d^\xe2\xe2\xd1T\x82\xa3\xbd5{\xbe*\xfc%\xb2\x9f\xc5~o\x0e#\x07a\x11X\xd3MZ\xcf\xdc\xdd\xbd\xbb;\x96;\xb6\xd9\tWl\xbdz\xa1\x99\xfd\xf2\xaa\x1d\x85\xa1P\x0b\x17\xcf,@I\x93_\x9dPG\x8c\xb9\xd7\xa9\xa2\xa6\x84\x14@\xde\x85\x02Z\x96"\x0e\x07z\xad\xc5nU\x14v\xfe\xb3\x93o\x16|M\xe3\x81\x81DT\xf6\'\xf3\xaa\x85\xac4\x0bBQm\xd7\xaa\xc0p\x10`\x02\xacf\x85\xccT\x0bV\xb3Bh\x947\x8dz\x95\xa5\xe9\x03a\xf8\xb68\xcd\x88\x8f\xfb\xf3\xe6\x9b#\xf9m\xddo$%%j_\xba\xd6;s\xc6l2\xb3q\xe3\x06\x0e\xa5\xfcF\x8d6a\xa4&\x98H\x8c\xcb\xc2bV([\xcb\x8f\xa6\xf7E\x90Zy\'/\xbd\xf2<\xff[\xfc?\x92\x93\x93\xb5\xd2p/\x04sCQ\x146o\xd9\x82\xbe\xc6I\x82\xca\x18\xb5\x1d\xbb43\x88\xd6\xadQ\xcf\xe6P\x1e\x12v\x97PI\'\x91t!\x8b\xe8\xdaA\xa4%\x9a\xd5[\xb0\xd9\xfe%\xf5~le\x8a&\x84\xd5\xb3JX\x15\x85s\'34\x05\xe0\xae\xbc\xdc\x1d\xcb\x9d\xbc\x1a\xa2$;^;\x1f\n^|\xc5@\xce\x89k\x0f\xc5K~6~q{\xee\t\x9c\x7f\xb5,\xb9"\xc9\x9d\xcc\xccL\xfe\xf9\xe7\x1f\xbc\xab\xc6\xe3\x1d`\xc4\x92\xa5\xa8\xe6\x0c\x9b\xc0\x93\xc8\xbes\xd5\x0e\xa2\xca4IB\xd6\x84\x9e\xd5*0eX1x\xcb\x84W\xf6\xe3\x9e\xa7\x1bq5j\x1b\xaf\x8cy\x81\x89\x13\'\xf1\xef\xc1\x7f\xb1X\xaf\x7fcp!\x04W\xaf^\xe1\xfb\xef\xbf\xa7\xd6\xdd\xe1X,\x8e\xb6t\x819\xc3\x8a$A\xc5\x06AT\xed\x9f\xc1\xf8\xe9\xa3\xf8\xfa\xebo\x1c\x94@\xfe\xef\xc6\xc6\xe5\xcb\x97Y\xb4h\x015\x9bVR\xafa\xd7\x1f\xda\x1f\xda\xef\xb6h\'\xd9\x13\xbf\xaa`\xd2I\niI\x99\xa0W\xa3\x81\xda\xae\xad\x15\x9dz\x1a\xed\x94\x92P\x0bSm+\x02/\xa3\x17q\x07\xb1\xc7P*Z\xb4Ih\xed\x99\x1c\xfb\xd5\xb6\xc6+l\xd1LKY\xcd\xb5\t\x98\xec\xfb\xb2\xbf\xa4\x9c\xcfb\xff\xcd=\xb6w\\\xd0gT\xafk\xff\xcb)\xedv\xa2\xa8k\xe7\x8dP\n6\x84\x19\xcf\xb8q\xe30\x9b\x9d\x85\xdeK/\xbd\xc8\xc7\x1f\x7f\x8c^\x7f=;?\x89B/v!\x04\xf1\xf1\xf1<\xf5\xe4\x93xw8AP\x84\x97j\xaeF\xbb\x94\xcd|\x8d\xcdkE\x15|\xb6{\xc9\xfe\xdbvH\xfbKHx\xf9\xe8\xc8H\xb5\xf2\xdf\xc6K\x88\xb3\xe5\x18;z\x1cm\xdb\xb6%(0\xd0\xf9&r%\xfby\x85\x10\xac\xdf\xb8\x81Qs\x07\xd2\xb0{e\xd0\xdc$\x05\xaaIM\xc2fc\x07/?\x1d\x89\xe73\xd9\xb5\xe8*\xaf=3\x96\x87\x06?\x82\x9f\x9f\xaf\xcb\xb9sb\xab0\xbf\xad]\xcb\xc8)\xcf\xd2x@0\x92A\xd6&h\xb5F/\xb0\x17\x8c\xbdwm+\x0e\x04:/\x99\xabg\xd3\xd1y\xcb\xf8\x87\x1a\xb3\xdd95a\xae=\x8dz\xcc\xe1u\nTe*$\x03\x1b\xdf\xbb\xc2\x96\xcd[)W.ZM\xbcAbcc\xa8Z\xb5\x9a\xd3\x860!!!L\x9f>\x9d\xb6m\xdb!I \xcb\xea\xe8\x03$t\x92l\xf3V%\xcb\x94EVV\x16V\xab\x15Y\x92\xb5i\x0c\x81,\xcbj/P\x08,\x165v\x92\x10 \xebdl\xaf\x02IBQ\x04B(\xf6\xc5d\xea\x93\x0b\x84\xa2\xf5"\x85\x82$\xeb\x00\x81,\xeb\x90P=\xca\x14%\xdb\xdc&\xc92\xb2\xad\x9c\x91\xf02\x1a\xd1\xebt\xf6\xb8\xfb6\x8f0\xf5\xda\x02\x8bEA\x92%t\xb2\x8c,\xcb\x0e\xca\x0c\xf5]\xaaU\xc6~\xdcj\xb1j\xcf#c\xb5\xaa\xae^:\x9d\xce\xfe\x9c\x0e\xd5]\xeb\x0c\x08dY\xf3\x04So\x80\xf4\x8c\x0c\x92\x92\x93J\xb4\xd7\xeb\x88\xbf\xbf?\xfe~\xfeH\xb2\x84,\xc9d\x99Ljy\xe8t\x18\xf4z\x82\x82\x82\xf0\xf6\xf6\xbe\xeeN\x86\xc8cG\xb0\x01\x03\x06\xa8;\x82U\xacPbc\xb5\x92W\x00\xe3\xc7\xf3\xee{\xe3\xb0\xb8*\x80\x17_\xe4\xe3\x8f?Ao(\x1d\x9bF(\x8a\xe0\xf4\x99\xd3T\xafQ\x8d\xc7\xbe\xac\x83,4\xd3\x83\xad\xa1d\xcb(;j\xc1j\xdeLN\x15\xc8\xd63\xb2\x1b\xb9\x91u\x12\x06/\x99\xb8\xa3\xa9\xec\\\x1e\xcbK\x0f\xbe\xcdc\x8f=Jdd\xd9k\xaa|\x99YY\xbc\xf6\xea+\x9c\x8bXKd\xb5 \x84\xd9y\x92\xd4\xf6\xb2m\xeb\xab\x8c~z\x92b\xd290\x1f>\x9d4\x9d\xbb:v@\xa7\xcb\xbb\xcc\x85\x10\x98L&\xde\xff\xe0}\xfe87\x8f\xfa\x9d\xcb`\xca\xb0"i\xbbKj\xf2HE\x93G\x92\x1eP$\x84U \xeb$\xacf+\xff\xfe\x11\xcf\x9d=\xa20g(\xaaB\xb4\x17\x87\x83\xdb\xa7v\x1e\x07\xd9\x02:A|\\2\x15c\xfa2v\xec;\x84\x17`?\x84\xfc\x10B\x10\x17\x17\x97cG0\x00\x9dN\x87l\x17\x92 K\xea\xdd(\x9a\x89J\x96e\xca\x95/GddY\xf4\x06=Y\x99Y\xf6\xfb\x97$\xd9^\x14BQ\xec]\x01{g\x01\xd0\xc9\xb2\xf6\xac\xea\x13\xaa\xdfS\x7f\xa2\x99\xc0\x14\xab\x15\xec\x8b\xf0\xb0{DY\xadj\xef^\xd8\xba\xe7\xb6r\x12\x02\x93\xd9\x8c\xc9lF\xa7\xd3a\xb1\x98I\xbc\x9aLJr\x12V\xa1*0\xbb\x90\xd7\x9e\xdf\x86$\xabQi%Mi\xe8d\xf5\xc5*B\xf5\x1e\xb3)\x12\x00I\x92\xd5\xf7\xab)=E\xa8\nI\xd6e+>\x84\xba\xeb\x83\xed\xfe\xec\x05RJP\x14U\x11\xca\xb2l\x7f6\x9d,sg\xa3F\xbc\xf4\xd2K\xf4\xebw/~~\xfe\xae_+\x10"\xbf\x1d\xc1&M\xa2\xc2\xed\xa8\x00l\x8dy\xc2\x84\t\xda\x08\xc0yK\xc8\x17_z\x91O>\xfe\xa4\xd4\xec\x1a$\x14\xc1\x9e\x03{h\xd6\xb2\x19O~S\x072Ua\x80$\xa9\x9b\x85H\xd9RJ}\x95Z\x03\xb6\t-\x07\x87H[\x8a\xe4\xa0\x00@R+\xa1^"+\xc9\xca\xdf?$pO\x93\xbe<\xf3\xcc\xb3\xd4\xabW\xaf@n\x88\x028|\xf80\x83\x9f\x1c@\x9d\x012\xbe\xe1z\x84\xcbB)M\xc68T7\x81\xe4%qn\x7f\x1a\x1c\xa8\xca\xd4O\xa7R\xb3fM{\x9acNG\xce\x9e;\xcb[\xa3\xde\x82&{\xf1\r\xf6R\xaf#\xab\x82I\xa8R\x0b\xb4g\x95\xd0\xa1\xf3R\xb0X\x04\x92"a\xf0\xd6\xf3\xdf\xc6\x8b\x94\xaf\x15\x88O\xb8\x17\x8a5\xdb\x13\xc2Vfv4A\xa7\n5\xf5\x8e\xbc|\xf5\xfc\xf3\xd3i\x1ej1\x86!C\x9e\xc5\xd77\xffQKAp\xb7%\xa4\x87\xdb\x9b\xa8\xa8(\x16-ZD\xeb\xd6\xad\xaf\xa9#f#\xaf\x11\x80\xaa\x00&S\xb1b\xf9\\\xdbY\x91\xe0\xd0\xac\xf3\x97*ED~\x8f+gw\x07K\x07\x12$^M\x02\x13\xa0\xd8\xec\xaa\xa8\x02\xd4V\x8aN\xc2\xdf\xb6\x91U\xf6\x84\xa8\xd6\xd7\xb2w\x84\xd4/h=8IA\x08\x05\xab\xd9\x8a!\x00Z?\x1d\xc1\x9e\xa4_xi\xe8\x0b\xec\xdd\xbb\'G\xe8eW\x84\xe69\x13\x7f\xf1"rH*\xbeA\x12X\xd4P\x0cZ\x0e\xad\x97g\xbbk\x87\x8f\x19*\xd6\r\xe0J\xc0~\xe6\xcc\xf9\x82+W\xae8\xd7\x12\x97\xe1\xbaUQ8\xf4\xdfa\xf6\xc7l&\xac\xa2\x8fj\xbe\x91p~\xabB\xed\x1d\x83\xd6\x1bT$t\xc8\xc8\x92\x8e\x84s\xe9\x983\xad\x84V\xf2E\xb1\xaaQ5\xb3\xcdc\xb6\x12\xcc.3\xe1\xf0\x97$\xd4\x9eZ\xcaI?\xea\xd6\xad\x83\xb7\xb7\xb7\x96r\x83\x08\xb0Zs_\x1f\xe1\xe1\xf6$..\xce)Bna\xa3V\xb7b\xaes\x0e\x97+1\x05\x90\xdf\x9e\xa8\x02M\x82\x96"|}|A\x02=z\x84Vt\x0e2\xd5~\xbb\xaaj\x10\x08\x9b\x07\x80\xcd\x92\xa1\xba\xc0\xd8\x0c\xbfj.\x079kCX\xd4=h\xeb\xde\x13\x81T\xf3\x0cc\xc7\x8ee\xf7\xee=X\xf3\n{\xa0\x8dD\xf6\xed\xdf\x8fYd`\xdbs\xcb\xfeQ\x0bT\xbd\x96\xe3M\xa3MV\xcb\x82\x1a\xedBY\xf7\xc7o\xda\xb5\x1c\xca\xdeE(feer\xf4\xc8\x11\xa2\x9aJd\xa5[\x9c\xd2\x85\xa4\xe5W\x1fQ\x15\xe5:+\xe9\xc9\x99\x18|tH\xb2\xe0\xf4\xae+\xd4j_\x86\xb4d3\x06\x83\x0c\x8a\xad4m\xa5\xe2\xa0\x04$\xdb1\x15\x9dQ\xe2\xe2\xe9\x14\xaaG\xd6\xa7B\x85\n9\xee\xed\xba\x91\xc0\xcb\xe8E\xbf\xfe\xfd]S<\xdc\xe6X\x1d\xf7\xb1(D\x84}\x82\xbe\xe4(1\x13\x90\x8d\t\xe3\'0\xee\xbd\x9c&\xa0\x97\x86\xbe\xc4\xc7S>.5& \x04\x1c;~\x8c\xda\rk\xf1\xf8\xec:\x08\xab:!\'9\x88R\x1cD\x97\xdaoUsH\x926j\xd0\xec\xba\x92\xd0\xe6\x0el\xd8,&B\x9d@\xb6i\rI\x96\xd0\x1b\xf4\x1c\xdf\x98\x821\xae\x1a\xef\x8e{\x97&M\x9a\xb85\x07\tT\x1b\xf3\xfb\xef\xbf\xcf\xca\x7fg\xd3\xa4o4\x96,\x8b*\x90\xb5\x99J\xf5\xda\xea\xb5\xec=n[\xffZ\x80\xc1[\xe6\xf4\x8e\xab\xd4\x90\xfa\xf2\xce\xd8w\x08\x0b\ru\xb9\x8a\xca\x95+Wx\xfd\xb5\xd7I\xad\xbb\x83\xc00/m\xf4!i\xee\x9d\xb6\xdez\xb6\x01\xcc\xcb \x93r%\x13\xffH#W\xcfeq\xf5l*w4\n\x03\x9d\x04V\xd0\xebe,\x8a\x15\xc5\xaa\xee\t\xac\x16\x85\xed\\\xeaY\x04\xea\\\x8a\xce\xa8c\xef\xaa\x0bt\xac\xf4(c\xc6\x8c\xc1\xc7\xe8\x8b,;\xd9\xb4\xae\x1b\xb3\xc5\xcc\x9f\xdb\xfed\xfc\xf8\xf1\xec\xdb\xb7\xcf5\xb9\xd8q\xb4\xcb\xdb\x84\x85\xebqw\xb8\xa6;\x18\x1c\xf3F\xd2J<\x8f\xf3\xabg\xd2\xde\xadZ\xb9\\\xb3\xe4\x8a\xfa\x1d\x97\xfc\xd7v\n[\xe5U\x7f\xb8\x9e+\x1f\xb2kd>\xcf\xa8\xa5)B\xa1{\xf7\x1eL\x9e4\x89J\x95*\xb9f+\x10\xf9\x99\x80&O\x9aL\x85\x8a\x15\\\x93\x8a\x8d\x12W\x00\xe3\xc7O\xe0\xbdq\xefb\xb68O\x02\x0f\x1b>\x8c\xc9\x93&\x97\x1e\x05\x00\xc4\xc4\xc4P\xb7V]zN\x8a\xc6\xdb\xdb\xe0`\x9f\xb6M\xede\xcb!\x81:1&\x14\x9b\x01C\x15n\xe08c\xac\xe5\xb6WF\x9b S3HZ\x92\xd1\xdb\xc0\xbf\x1b/b=Z\x81iS\xa7\xd1\xa0ACt\x92\x0eIvn\xd2)\xa9)\xbc\xfb\xee\xbb\xec\x88\xff\x91\xc6}\xca\x92\x9elq\x18\xe4i\x93p\x926\n\xb1\x1d\xc3\xd6\xa8T\x85#Y$\xfe\xf7\xde~~[\xbc\x8d\xe6\xcd\x9b\x83$;\xf5\xca\x01N\x9c<I\x93\x16\xf5\x194\xa5:VESt\xea\x03\xa3\x89jm^D}\x06\xbd^G\xe2\xc5,"\xca\xfb\xf0\xc7\xfcS4\xea\x1e\x85o\xb0\x17\xe9\x89&$\t|\x02\xbd0k\x1e2\xceb\xca\xde\xd2A;Of\xba\x85]\x0b\x93\x99<r6]\xbbv-t\x93\x8d\xd5j\xe5\xe4\xc9\x93\x9c9s\xc6A\x89g\x17\x95\xeda\xd5\xe7%\xfb\xfaB\x150\x85y?\xaag\x90\xeaIc\xb1X\xb1\x98\xcd\xea$\xa5P=\x86\xec\xb5\xcev/Z\t:N\xd2\xda\xbe\xabN\x1e\xabQ]m\x02\xce^\xb4\xb6\xc9eY\xaduV\x8b\xd5\xf1u\xdb\xcb\xc0\xae\x8e\x85\x92=\xa2t\x98\x0c\x07u\xb2W\xfd\xaaz\xf2\xecb\xb3\xdd\xabzD\x08\xb5\xacl\x13\xc5\xeayU\x1b\xaa\xa4=\x83\xed4\xf6:!\xd4\xea\xacVc\xdb\xf9\x9d[\x9f\xfaz\x1c\xdb\x98\x86\xe6I\xa5\x9eRB(\x8a\x9aE{_\x8e\nV\x11\x02\xc5jE\xa7\xd7\xd3\xbaU+\xaaV\xad\x86,__\x98\x0f\xa1(,\xff\xe9\xa7\xdc\x15\xc0\xe4\xc9T\xa8P\xccs\x00\x0e\x94\x02\x05\xa0\xba\x81Z\\\x14\xc0\xd0aC\x992y\xf2u\xba\x81\x16\r\xf1\xf1\xf1\xbc\xf1\xfa\x1b\xa4\xd7\xfb\x8b\x80\x08\xb5\xe7\x9b\xdd\xec\xb2e\x84\xede\xea$\xb0Z\x04\xb2QBX\xed\xf5\xdb\xde\x13\xb1co%\xd9\xbd\x7f\xbbiC\x80^\xa7C \xd8\xf5\xf39j\xf9t\xe1\xc3\x0f?\xa4R\xa5;\xecY\xd4l\x82\x8bq\x17x{\xcc\xdb\x1cQ\xd6\xd2\xb0G$\x19)\x16{\xb5\xb2\x9d\xd6Q\xa0\xda+\x9d\xc3hD\xaf\xd7sxs\x0c]\xeex\x91\x11o\x8c\xc4\xa078U|!\x04?\xfd\xfc\x13\xcfM\x18@\xaf\xe7\xeb\xaa\xa3\x19\xa1\x86w\x00\xe1\xa0\\l\x8dWm\xa0Iqf\xd2/\x9aH5\xa5Q\xafc\x14\xb1\x87S\x08\x08\xf2&\xa0\xac\x17\x99\xe9V\x872q\x1c\x9b\xd8\xcaK\xbdm\xbdNG\xdc\x91d2\xff\xa9\xc2\x82\x1f\x7f$4$Ls\xa7-\x9c\xc6\xa3\xde\x81@FFA\x15\xb0\xea\x11\xf5N\x84P\x90%5\x84\x85Ma\xca\x92\xe6>)\xa9\x8e\x026\xaf\x18l#.\xfb9\xb4b\xce~\xad\xee\xef\xda\xe1\xe1\xb3\x05\xb5*$\x15m?\xe5\xec\xe7\xb5\xc5FRK\xc9\xae\x845\xe1\xaa\x9dF\xf5>r\x88S\xa4\xde\xaf\xa6\x0c\x1c/\xa9\x8d\xe4\x9cnL{-\xae\xca\xcd&\xc0m?\x9d\xb2;\xa49W\xd2\xdc\x15\xa4P\xd4\xf2t:f;\x8fc=\xc8\xbd\xe4\xc0v\x9f.=\xfc\xdc\xae\x99}but\xeexnUY\xb8\x17\xfa\x8e\xe7\xce\x0f!\x14\x96/\xcfK\x01L\xa2|\x85\x92\xf3\x02\xcaiK(nryn\xf5p.\x89%\x84\xd1h\xa4v\x9d\xda\x9c\xd9u\x15\x83Q\xa7\xde\x9d\xadG\xa3\xe5\xd1\xfa HZ\x0f\xce\xcbO\x879\xcd\x8c\x97Q\x97\xed\x13-\xb2sj_R\x11B=\x93c\xa5\x93\x04VE\xad\xd6wv\x8fb\xf7\xe9\xdf\x997o\x1eI\x89W\xb3\xf3hg\xb3X\xadddd\xa2\xd3\xe9\xd5\xb8:\x80\x10j\xb5\xce\xd6UZ/K\x92\x90da\xf7\xddW3\xab\r\'\xbcR0\xcbW,\'\xe6|L\x8eFi\xb1X\xd8\xb4\xe9\x0f\xaa\xd7\x8dD(\xea\xf3do\xf6\xe2\x80Z\x0cZ#\x92I\x8a\xcf`\xdf\xe6\x8b\xd4\xefT\x81\x13\x7f\'\x10\x1e\xea\x8bo\x98\x9e\xcct\xd5\xbej3\xf5\xa0\xddbv99\x14\x0f\x12\xdb~=\xc5#\x8f<JpP0\x92\xa4N\n\x17\x16\xb6:\xe7(\xb8m)\x12\xea\x060\xa0\xed\x8d \xcb\x9a\xbf\xbf:)-K2:Y}\xc7\x92\xd6[\x94Q\xd3%\xfb~\nZ\xb9\xe7\xb2\xbf\x82$I\xd8\xff\xb3\x9dW[G \xcb2z\xbd\x1e\xbdN\x8fN\xd6\xa1\xd3\xe9\xd0\xe9dd\xd9\xe6\xa2\xaa\x1d\x93d\'\xb7U\x9d\xf6=\xa7c:\x1d:\xdby\x1d\xae-\xcb\xb2\xfd\xde\xed\x1f\xedyl\xf7\xe1\x94\xd7\xe1\xa7\xfd\xb8\xd31Y\xfd\xdb\xf5\xb8\xf6|N\xdf\xd3\xdcF\xdd^\xc3\xe1\xdc\xaa\xbbi\xce\xef\xdb\xbfc{\'\x92\xac\xbe#\x97\xfb\x93$\x87\xeb\xdb\x9eUV\xef\xd3\xf1\xdc\xb6ru\xfd\xee\xb5\x93\xffw\xf2\xcfQt\x94\xbc\x02\x10\x92\x93\xbc\xb3\x1fv\xf87\x07\xb9\x1c.j\xfc\xfc\xfd\xa8\xd7\xa0\x1e\xb1\x87\xb2\x10\x80b\xc5Et\xa9=\x10\xdbO\x01\x982\x14|B\xbcH8\x9b\x89)\xd5\x8aNv\x18v:|\xcf^\x0bl=8\x87^\xba\xc0\x8aP\xac\xc8\x06\x1d-\x1e)\xc3\x17K\'\xb0u\xdb\xd6\x1c\xf3&\xb6^\xa2\xa4S\xcd\x04H\xd9\xe7u*c\xa1\xf6fU\x81\xa5\x03\x91]\r\x14\x0b\x04\x96\xf5\xc6\x12r\x99\xbd\xfb\xf6f\x0f\x16\xb4\xf3gdd\xb0g\xf7^\x8c\xfe^HB]\xcc\xa4%:\tl\xfb\xa5\x15\xd0\xeb\x0c\xecYu\x81\xea\x8d\xc28\xbb#\x89\xd0h_\xe4`\x19\x93Y\xf5|\xb2\x95\x95\xe3w\xc1\xe6\x1a*#\x14\t\x9dN\xe6\xcc\xbe+\xd4\xf6mC\xc7\x8ew\xd9}\xd3o\xb8\xf5\xb8\xd4%w\xa7+\xe8\xb1\x1c\x07s\xfc\xedz\xe0:\xc9\xeb4y\xa5\x1539\x1e\xd7\xf1%\x175yM\x0b\xe5\x96P"r%\xb7\x9b)\x1eJ^\x01\xd8\x86\xae\xae\x88<^H\t\x95\x99,\xe9(\x17]\x8ej\x11\r\xb9r6\x1d\x9dA]\xc0\x93\xeb\xedHj-4\xa5\x0b\x82\xa3\x8d\xa4^\xcd$\xe9B\x06z/\xad\x7f)\xd0\x1eF\xfbH\xda1\xc9q\x9bL\x9bBQ\x05\xb0\x97\x9fL\x8b\x01w\xf0\xe6\xd87\xd8\x7f\xe0\x80\x83\xe8\xd4rI\xaa\x0b\xad$9\x98\x93r\xf4\x92US\x85\xa2\t_[\x88e\xf5\xfa\nF\x1f\x1d\x86@3;\xff\xd9IVVV\xb6)B\x96\x89\x8f\xbfLb\xd2UB\xa2\x8c\xe8\x0c\xd8M\x10\xd9Wp\x10\xe7\x12\xe8\xbdu$\x9c\xca$\xe1D\x16\x87\xfe\xbc\x84O\x84Dp\x94\x0f\xe6,\xcdE\xd5V\x10\xdaO\x81m\xce\xda\xa6X\x05\xb2\x1e\xccY\x82\x9d?]\xe0\x89\'\x1f\',,\xa2\xf0\xea@a\x9d\xe7&&\xb7f\x96\x8dc-\xcb\xaba\xde\x9c\xd8\x9f\xa6\x04\xeaBI\x97d\x89+\x80\xdc\xca\\\xa0\t\xd0\xd2\x84\x04U\xaaV\xa1Q\xc3\xc6\x9c\xdb\x97\xac\nZ\x1cL\xe8\xb6\xfbUo^\xf5\xb8\xd1d\x9b%K!\xa2J\x00\x16\xab\xe0\xf4\xeeD\xf4F\xd5\x9c\xe0(8\x11\xb6\x93\x89\xecq\x85\x93\xf0\x16(f+e+\xfb\xe2_=\x89\x85\x0b\x17\x92p9\x014\x8f\x05\xbd\xde\x80\xd1\xcb\x0b\xb3IA\xa7w\xf8\x9e\xe4z\x1eu\xf5iV\xaa\x05Io\x13\xe0\xda\xb3\xe8@\x12\ne*\xfar>\xe6\x1cW\x12\x12\x90l\xb6^\xe0\xca\x95\x04\x923\x12\xed\xabEm\xb7,$\x87\x91\x9c\xed\xa7\x00\x14\x85\x7fV\x9d\xc1\'\\\xa6\xe3\xb3\x15\t\x8c\xf2!+C\xb3e;\xd6~\xc7Vh[\x05\xab\xd9ru\x06\x89\xe3\xdb/\xd3\xae~W\xda\xb4i\x83\x97\xd1\xcb\xe1\x8b\xc5\x8fj"P\xcd\x03\xf9r\xc3-\xfc\x86O\xe0\x16\xf5\xacy\x9d[S\xc6\xda\x84\xb7\xa4\xcdI\xd8j\xa7\rE\x0bIaS\xfb\xb9\xd9\xc7\xd5\xf9\x00\xc7s8\xa49\xfc\x9e/Zf\xe1\xf0\xfb\x8dR\x80\xb7h\xe7\x9a\xde\xfdM\xc0u+\x80B*{\xd0\n\xd5\x15g\xc3J1\x92\xc7e% \xc0\xcf\x9f{\xee\xb9\x87\xd4\xf3\x06L\xc9Vt\x92\xce\xc1\xa9F\x0bg\xac\xfd-$u3t\x81\x00\x9d\xc0\x9ci%\xa2\xb2\x1f\xe1\xe5\xfd8\xb5\xf3*\xc2\xaay^\x08\x9b\x08\xd6\xae\xa2M\xac\xaa\x02R\xb5\x8f\xaa\xadFF\x08\t\xb3"\xa8\xdf\xa5\x0c\xab\xb6,`\xc7\x8e\x1d\x98\xadVd$\x02\x03\x03\x88\x88(\x83bU\xd0\xc9\xb6s\xa8#\x01{/_C\xd2KX\xb2\x14\xf4^z\xed\xdcjhfI\x911[\x04\xa1U\xfd\xb8\x9cp\x99\xa4\xa4Dm\x94\xa0\xde]VV\x16\xfe\xe5%\xbc\xfd\r(\x96l\x1b\xbc$\xdb&\x82m:L`\xf0\x969\xbb/\x89\xf4\xa4Lz\xbeZ\x8d\xc0p\x1f\x14\x8b\x82@U\x00\xd9\xa3\x1c\r\xfb@@\xbdW!\x04z/=I\x17M\x9c\xdfe\xe1\xde{\xfbQ\xb5ZUdg{\x96\xc3\xef\xee\xc9M(]/\xee\xeak\xae8e\xcdy\x1f9\x8f\xb8\xa2\x9e \xff|\xd7\x86\xbd\xae\xb9\xde\xa2\x1d\xadC`\x1bM\xda\xe6/\x9c:- \xdb\xe6D\x1c\xf2\xbb\xc3v\\\xfd\xe9<jv\xff\x8d\\\xd02K\x0e\xbf\xe7\x8es\xa9\xddP\x19\xba|9\xb7\xe7\xbcV\n\xe7,\xd7\xcfu+\x80\xe2\xb9\xf1\x1bze\xd7G\x1e\x0ff\xeb\xf1\xb6h\xd1\x82\xba\xe5[\x11\x1f\x93\x06\x92MP\xe3\xdc\xcbv\x11:\xb6fc\xceP\x08,\xebM\xb9Z\x81\x1c\xdb~\x99\xf4$\x13z/5M\x13{\xd9\xcaO\x02\xc5jE\xd2\\\xe2TA&!,\x02\x9dQ\xa6vw\x1f~Z\xfe\x13\xc9\x89\x89\x08\xc0\xdb\xdb\x1b???,Y\x02\xab\xc5v6\x15uR\xd56\x91%\xa1X\x04\xb2^Bdh\x93\xce\x12\xe8\xbdd2S\xb30\x18e\x8c\xbezN\x9c?\xc2\xa5K\xf1\xa0y\xefY\xacV\xf6\xef\xdf\x87\xeck\xc2\xe8\xaf\xcb>\xb9,aN7\xe3\xed\xa3*\x13\xa1J\x0bL\xa9VN\xefI\xa4a\x8f(B\xca\xfbb6k\xcb\xd3\xecB\xc0\xc5|f\xeb\xf9\x0b\xb5,%$\x14\x0b\x1c\xde\x16K\xdb\x06\x9d\xe9\xd5\xab\x17\xfa\x1cq\x8a\xf2xa\x1a\xd7\xddX\x0b\xbd\xfa\xe5\xbc\x8f\x9cG\xdc\xe3._\xa1\xdf^1\xe1\xacB\xd4zmw\xfb,T\x9c\xafc\xfb\xeb\xba\xca\xcd\xdd\x0b(\x04l#\xac\x92\xe2\xba\x15@a\xa0\xb6\xf3lo\x0f\xd7\xb4"+\xf5\xebD\xd2zA!!!<\xfe\xc4c\xec[w\n\x1cL\xed\xaah\xcd\x16\xde\xb6#\x80*A\xb5?\xcd\x19\n^\xbe:\xea\xb4\x0f\xe7\xc0\xda\x8b\\\x8dKG\xa7)\x01T\x19\xaf\x9eW@p\x84/\xe6\x0c+:\x9d\xac\xcaMM\xdb\x98\xb3\x14B*\xfb\xb2;}\t\xeb7\xacS\x03X\xe9\xf4T\xadZ\x85\x8cd\x89\x8c4\xcd\xadV3)\xd9n\xd0>\x00W\x04\xde~:\x8c(H\x92\x82\xd1W\xcf\xa5\xe3i\x04F\xfa`J\x17\xf8\x04x\x91\xa5K$++K\x0b\x02\xa6\xce-\xc4\xc7_F\x08+\x8aE]\xd7 \x04x\xfb\xca\xc4\x9fN\xb3?\x83\x9aW\xe6\xea\xb9\x0c\xf4F\x89\xaa\xcd\xc31gi\xbe\xe5\xb6\x11\x83\xed\xae\xec\xef\xdev\xc4Vb\x02/\x1f\x1d1\xff&\x13\xb7\xcd\xc8\x90!\xcf\x11\x12\x1aZ\xe8\xbd\xf9<\xb9\x96\xeaW\x8c\xb7e\xc3\xe9\xf6l\xabJ\xb5\x8fm-@\xb1\x96\xd7ub\x1fe\x14\x13\xc5w\xa5\xfc\x91\xd0\x1a\xcc5S8\xef\xb5D\x15\x80\x84\x83`t\xc1\xee\xe5Q\n\x91e\x99f\x8d\x9b\xd1\xbb\xc5s\x9c\xfa\xe7*z/\x9d\xb6\xba\xd7&\xdc\xb2\x85\x9c\xfd\xd1\xec/Y P\xb0\x98\x04B\x86\xd6\x03*qp\xfd%b\x0f\xa7\xa07\xa8\xeehBs\xc7\x94\xb4\r\xd7/\x9dJ\xc7;P\xd6\xec\xec\x02I\xd6&{\x85\x9e\xaa\r"\x99\xf3\xd5l\x12\x12.\xa3(V\xea\xd6\xad\x8bH\xf1\xc5\x94\xaa\x86\xfd\xb5\xf7\xb64\x93\x92\xda\xaf\xd6"T\xcap\xec@\n^\xdez\xd2\x12\xcc\x18\xfc$\xccf[\xfc \x05d+\x19\x19\x19h\x01\x1e\xb1X-X\xadV\x02\xc3\xf4\xf8\x05\x1a\xd4\x9d\xbf$\xd0{\xcbXL\n\xb2\xc1\xf6\xce$\x94,\xc1\xbf\x9b/\xd1\xe4\xde\xf2\x9a\x02\xd3!\xeb$\xad\xd7\xaf\xbaY\xaa\x95_\xfb\x8aC\x85\x16H\xe8d\x1d\xe9\x89\x16\xf6\xafL\xe4\xed\xb7\xc7\xd0\xf0\xce;\xb3\'\xb7\x8b\x9a\xebi[\xc5p[\xb9\x93s\x95\xb0\xa2(7\x95\xad\xfaf\xb9\xcf\xc2F\xc0u\x9a\xbb\x0b\xa7\xbcJT\x01\xe4\xcd\xf5\x14J\xf1\x11\x14\x14D\xdf>}\x88\xddn +\xd9\xa2\x9aSl\xd6\x0bM\xc0\xa1\xfd\xae\xbe,\x9b\xb9#\xfb\xb8\xd5"@\x0f-\xef+\xcf\xb1-\x978\xbb?\x11I\xaf\xc6\x01\xb2\xe53\xa5[\x08\xaf\xe0\x87%\xcb\x16\xf7=\xdbp"\xac\n\xa1\x15\xfcH\xf2?\xc1\xbau\xebQ\x84\xa0B\x85J\xf8[\xc3\xc9\xb8bR\xf3jf#\xdb\xc8A=\x85f\xd3Qd.\x9fK\xc5\xdb\xcf\x8b\xf4\xabY\x84\x94\xf1\xc5jV\x85\xba\xc5dA6\x9a9q\xea$\x99\x99\xe9 \x04\x99\x19\x99\xa4\xa7\xa7\x83,\xa95G\xeb\xcd\xa7\'[IO\xb2`[2\xe0e\x949\xb4\xf9\x125\xdaF\xe0\x13\xe8\x85\xc1(P,\x16\xbc\xf4\x12:\xbdz\x13jYe\xdb\x8eU$@\x9d\x17\x018\xf8\xdb\x05\xda7\xeeL\xcf\x9e=\xf1\xf3\xf3\xb3\xe7*r\n\xa7m\x15:N-\xc2\xfe\x87\xfa~%-\x1e\xbfM\xe8_\xeb\nz\xd7\x91\x82\xe3\xdf\xaei\x85AQ\x9c\xb3\xb4\x92\xf7\xa3\xba\xb7\x80\x14\x17\xa5X\x01\x94n\xf4z\x1d\x8d\x9b4\xe6\x9e6]9\xb8\xf1\x12\x8a\x1a\xae]\x13\x8a6/\x16\xd4\xc6\xa9\xfdn?"K\xf8\x06\xea\xd1{\xc9(\x16\x81\x97\xaf\x81\xf6\x8fU\xe1\xf0\xe6x\x8emO\xc0`T\xcd=\x02u4\xe0\xe5/a\xca\xd4\xe6\x014\xf1\xaf\xaa\x14\xf5\x97*\xcd\x82X\xf3\xdb\x1a\xe2\xe2\xe2\x08\x0e\x0eb\xc0\x80\x01\\\x8e\xbb\xaa\xf6\xe2\xedw\xa1\xe6W\xccZ\xc8\x07IB\xd6\xcb\xe8\xf4:2\xaf*\xc8z\x19aV3IB]\xc1\xec\x17l$=-\xcd\xbe)Ijj*\xf1\x97\xe3U\xf3P\x86b\x8f\'\xaf\x98\x15uw4\x01^\xbe:.\x9eH!#\xdd\xc4\x1dw\x06\x92z9\x93\xb3{\x93\xf0\xf2\xd1a\xb1\x80bu\x88\x8fo\xef\xb9\xaaJ\xc0\xa6\xdadI\xe6\xfc\xa1$\x94sey\xfd\xb5\xd7\xa9X\xa1B\x896\x92\xd2\x80\xbd\xcc4\x8f\x9bl%\xe5\xa8>\xaf_s\xb9\xf6\xc0\x1d\xffvM\x03\xf5\x86n\xe4\x95\xd8\xce\xe9:r)\x1c\xb2\xcfU\xa8\xa7\x05\xa7s\x17\x04A\xde\x16\x1eI\xca\'C\x11s\x03\n\xe0\xda\n"7$\xcd5\xcc\x95\xfc\xa2\x85\x96<\x12\x11\xe1\x11<\xf1\xf8\xe3\x18\xe3\xab\x12w8\t\xbd\x97jrQ\x85\x99\xed\xa5:\x8al\xadF\nH\xba\x94\x859UAQT\x85 \x1b\xf5tz\xb6:g\xf7\'qls\n\x06\x83N}9\x8a\x84\xac\xd3\x91z\xc9\xac\xae\x1f\x10\xaa\x00\xb5\tOK\x96BD5#\x07/n\xe2\xc0\xfe\x03H\x92L\xbbv\x1d8\xb0>\xc1\xe9\xbaB\xdb\xc8\xc4\xe8\xa3W;\xef\x02\xd5\xbf\xde\xac\xf0\xcf\xca\xb3\xf8\x04\xeb0i1`\x84\x04:\x9d\x8c%\xd3\x82\xbf\x9f\x1f:\x9d\x8c\x02\x98-\x16Lf\x13\xe5\xea\x06b\xcd\xb2\xd8\x17\x90I\x92@\xa7\x97\xc1,c\xc9\x14\x9c\xf8\xfb\n\r\xbb\x95#\xf9\xa2\x99s\xfb\x92\xa8\xd8(\x94,\x93\xc0\xaah\x1e@\xa8\xab\x87mu\xdfVB\x08\t\x9d^&#Q\xe1\xe8/\n\xaf\x0c{\x8d\x06\r\x1a\xa0\xd3\xe9J\xb2\x8d\x94\x12\xb4x;\x9a\xc7M\x89\x93\xd7B\xab<qn\xd76E\xe0V\xc9\\7\x0e\xe7*\xf4\x0e\xf6\xb5\xdd\xa7;\xd9V\x9a\xb8\x01\x05pm\x05\x91;\xb9\x9cG3\x13\x94v\x1a4l\xc8\xf0\x97\x86\xb3f\xde\x192\x12-\xe8\x0c\xceu\\m\xaeZ\xaf\r\xb5\x01\x0b\xc07\xc0\x80O\xb0\x1es\xa6\x95\xcb\xa7\xd3\xb8|&\x05\xb3E\xe1\x9e\xa7\xaa\x11{\xf4*\xfb\x7f\x8bC\x08\x81N\'!\xe9\xc1\xac\x08\xb5Cos9r\xf082\xa5Y\xa9\xd5\xc1\x8f\xf5\xeb\xd6\x93\x92\x96J\xd5*U\x19p\xf7S\xc4\x1cL\xc2\xe0m\xf3\xd6QC?X\xccVt^\xaa\xb9@g\x90\xd1\x19@\xef-\xf0\x0b3\xa8\xbdK\xed\xf4zo\x99\xcc4\x99r\xe5\xcac\xf0\xf2rTa\xe8\r\xb2Vu\xb4\xb9\x08$\xf4:5\x80W\xcc\xc1$"\xca\x05\xa1\x97u\xc4\x1dI\xa1\xce\xdde0eY@\xb1u\x19\xb5\xf5\x02\x9a\x97\x8fz9\tE\x91\xd0\x1btd\xa5X\xd9\xf0\xcd\x11\x9ey\xfc9\xfa\xf5\xef\x8fNWz\xe7\x82\x8a\x03\xa1j{\xb5\xec\xaf[\xe8\x96&\xd4\'\xd0\x1e\xab\xd0\xc8M\xd0:\xd6\xdb\x92\xe0\xfa\xec\xfb\xc5\xc7\r(\x80B\xc2\xe6\x15\xe8J\xf6\x9cj\xa9F\xd6\xe9\xe8\xde\xb5;\xd3\xdf\x9d\xc6\xce%\xf1\x983l\x8b\xb0T\xa1o\xab\xe8\x8e#\x02\x00\x8bY!+\xc3\x82\xd1OGh%\x1f\xc2"}H<\x95\xc6\xa9]W\xa9\xdb&\x8a\xd3\x7f\'\xb3}\xe1\x19$-\xda\xa6\xc1\xa0V%\xbb\xdcw\x10\x06\x8a"\xf0\xab\xe0\xc5\xd6\x93+8s\xe6\x14\x81\x81\x01\xf4\xeb\x7f/\xc77d\x91\x95nA\xd6\xab\x15QX\xc1\xe0#s%6\x03\xa3\x9f*\xfcu\x06u\x04\x82\xe40\x89-\x81\xac\x97@\xd1\x11\x1c\x1c\x8cN\xafG\xd2\xe2\x00\x99\xcdf\xb2\xd2\xb3\xb4+\xa3\xaesP\x04\x08\ts\xba`\xff\xbax\xac&\x89\xa3;.Q\xadE(\xa6,m_[T\x8f/\xf5\xe9\xb3\xff\xb55P\x9dA&+\xc3\xcc\x9e\x95q\xdc\xd7\xeei\x1ey\xe4\x11\xfc\xfd\xfc\xf2\xed\x19\x96l\xf3*\x9a\xab;\n3\xd5D@\t\x8b\xb1\xc2\xa7\xb0\x9f&\xbfzr}\\\xef\xfbu\xfc^\xde\xf7\xa5\x9a\xf4J\x8e\x92W\x00\xb8\xef\xe9\xcbnb\xde\x97F$  0\x80\xbe}\xee\xa5c\xdd\xfb\xd8\xbd2\x0eEA\xdb\xb8[\x15\xfbZ|H\xbb\xf7\x8b\x10\xaa\xa1]h+x\xad&\x81"At\x83 j\xb4\t%%-\x95\x90\xf2>\x1c\xde\x98\xcc\xca\xcf\x8e`JSPL\x8a\x1a\xf7\xdei\xb1\x94\xad\xe0$tBOh\x8dT\xb6n\xde\x02\x124n\xd2\x98nm\xfaq\xf8\xf7\x04\x0c\x06=\xc8:\x14\t\xcc\x99\n!\xd1\xde\\<\x9e\x86\xc1\xcb\x88$\xebU\xb3\x82\xb6w\xad\xda;\x17\x98S\xad\x18,\xfe\xe8\xb4\xc9D\x01$$$p1\xe1<F?/@]\x1c&\x0b\tI/\xd4P\rKb\x89\xfb7\x9d\xf4\xcc4\xeav\x8a\xc0l\xb5\xda#Q\n\tm\xa1\x9cz\xcf\x12 \x0b@H\xe8\xf52V3\xfc\xb3\xfc\x1c\xcd\xcb\xf7e\xe4\xc8\x91DEE\xb91u8\xfbL\x0b\xa1\x0e\x13\x15\xedc\xd5\xdc\x1em\x9b\xa5\xdb\xb2\xe6\xd6;\xcc\x8b\x82}\xc3\xf5\xfe\n\x87\xa2\x11f\xa5\x8fky\xca\x82\xbd\x8f\x82S\xb0\xf3]\xcb\x1d:rm\xdf\xbb\xb6\xdc\x85K\x89JY\xfb\x84\x96\x1b\xae\xa7\xd1\x96$\xd1\xe5\xca1\xf4\xa5\x97\xb83\xac;\xfb\xd7\x9dB\xd6\xdc\x1e\xb1W6\xdbh\xc0\x16\xe4LU\x08\xb6\x7f\x05`\xca\xb4`\xcaT\xa8\xda<\x94\xe6\x03\xa3\xe8\xf4b\x05\xb2R\x15v,<K\xc6\x15\x13WN\xa7c\xf0\xb2\xb9pf\xf7\xa6m\x93\xa7e\xef\x88`\xc9\xaf?p\xf9r\x02e"\xca0\xe8\xfeA\xa4\x1c\xf1\xe5\xca\xd9\x0cdm\xc51B`\xc9\x12\x84\x95\xf7%).\x1ds\xba\x05\x9d,\xd0\xa9\xd2X\xfd_\x82\xac435\xca\xd7\'\xda&\x88\x85 ))\x89\xb8\x94\xd3\xf8\x04\x1a\xec\xae\xaa\x08\x19\x84LjR\x1a\xbbV\xc6\xd1\xe6\xd1\xb24\xed[\x1e\xb3YQ7x\xd1\xfe\x93\xd4\x07V\x1fT\xa8\xc3\x17\x01\xe8\x8d2\x19\xa9&\xb6-:A\xb7\xbaO\xf1\xe6\xc87)_\xbe\xbc*\x04\xb5rRKI}\xca,S\x16\xb1qq\x9c>}\x86\x7f\xfe\xf9\x87\x8d\x1b7\xb0a\xc3\x06~[\xbb\x96\xf5\xeb\xd7\xb1~\xfd\x06\xd6\xaf[\xcf\x86\r\x1b\xd8\xb8q#\xbbv\xef\xe2\xcc\xe93\x9c?\x7f\x9e\x8c\x8c\x0c\x14\x87m\x1f\xf3\xaaa%\xd9(\xf3C5\xd5\xe5u\xf7\xc5\x83\xed>l\xf5\xb7\xc8\xb0\xcfy\x17\xfc*\xf9\x95Oq\xbd\xdf\xfc\x95y~\xe9EK\t\xee\x07\xa06\xeb\t\x13&\xf0\xee\xbb\xef\xba\xdd\x0f`\xf2\xe4)\x18\xae\xd1\x9d\xad$P{\xf4\xaa\x90<r\xf4\x08\xef\x8e}\x97\x13\xe6-4\xbe7\x02\xa1\xad\xba\xcd\xce\xe2\x10\x13][He\xab\x02\xea\xef\xd9\xe9\xb2^B(\xf0\xf7\xe2\xb3\x9c\xfb/\x8d\xc8\xca~tz\xa1\x8a\xb6\xd1\x8b&\xacQO,#c\xf0\x92\xd9\xb7"\x917\x1f\xfa\x94\x1e={\x90\x99\x91\xc1\xcf+W0n\xf63\xdc\xf5xm\xacV\xb3\xd6\x86\x04\x92N\x87\xa4\x08\xfe^~\x8e\x80\x08=\x8dz\x97#\xed\xaa\x15I\xdbv1\xeeP\n\xc1\xe7Z1\xfe\xc3\tDGG\x83\x04\xbfo\xf8\x9dgF\r\xa2\xd7[UHK0\xa9\x93\xd7:\x89\xd4\x04\x13\x9b\xbe>E\xd9\x9a\xfe4\xef_\x01\xc5"\xb0Z4\xbb\x8f\xf6\xc9\x1e\xb0h\xe2\\\x02\x83Q&\xe1\\:\xff\xaeNc`\xc7\xa7y\xe6\xe9\xa7\x89\x8a\x8aB\x96\xb5\xd1\x88\xb6AKJJ*\xa7O\x9fb\xe7\xae]\x9c8~\x9c\xd8\xd8X\xd2\xb3\xd29{\xee\x0cg/\x1eG\xe7\x03\x06\xa3\x1a\xca\x02I\xc2\x94\xa1`JS\x90\xac\x06\xaaW\xacET\xd9(\x0c\x06/\xa2\xa3\xa2\xa9Z\xad\x1aM\x9a4\xa1\xf2\x1dw\x10\x1c\x1c|]\xbe\xf2\x05\x1d\x9d\xda\x9e\xc1\x95\\\x8fkB\xebZ\xef\xa70P/\xad^\xdfv\x1fB\x80\xd5jq\xaa\xb3j]\x07\xec#Y\xd4-=e\xf5E\xab\xa7Q\xf3\xebd\x87M\x8b$\xd4\x11\xb1=\xbe\xbf\xa3\xd8)\xfe\xe7-\x0c\n\xb2zYQ\x14~\xcasC\x98IT(_\xb1\xc4\x8a\xa0\x04\x15\x80\xca\x84\xf1\x13x\xf7\xbdqX\\B\x1b\x0f\x1f>\x8cI\xa5lG\xb0\x82 \x84\xe0\xe8\xd1\xa3L\x9a4\x91\x1dg\x7f\xa2i\xffrx\x87\xe8\xb0fH\xaa\xe0\xb5\x87;\xd0\x1a\x19\xdaq\xd7\x1a i\rI\x96\xf1\xf6\xd1qrW"\xff\xfct\x8e\xb6\x0fU\xa2L\x95\x00\xacf\xab\xbd\xf7\xa56-uK\xc7s\xff&S_\x19\xc8\x98w\xde\xc1\xd7\xdb\xc8\xd5\xc4$>\xfb\xf4S\xfe\xbc\xb8\x98*\xed|\xb0h+r\x85\x10\x18\x8c:.\x1eN\xe1\xca\x85t\x1a\xf5\x8a"5\xc1\x8a$K\x18\xfd\xf5\xec^\x11G\xcf\x1a\xcf\xf1\xd6[o!\xebd\xacV+\xab\x7f]\xcd\xab\x13\x1f\xa7\xdb\xab\x95I\xbfj\x01$\x0cz\x1d\x97N\xa5pp\xe3EZ\r.\x8f\xd1\xcfKUx\x9a\xb7\x0f\xa8\xf1\x82\xecw\xaa\xc5~\xd7!\x11w8\x99\xab\x7f\x95\xe5\xe9\xc7\x9e\xa1\xff\x80\xfe\xf8\xf9\xfa\x02\x12f\xc5JFZ:W\xae$\xb0e\xcbV\x16.\xfa\x91s\x97\x8f!\x05\xa7\xe0\x13\x0c~aF\xca\xd6\xf4\xc7\xdbO\x8fbU0\xf8\xea\xf1\xf1\x911x\xeb\x11\x12d$[\xc8H6c\xb5Z\xd1I2\xa6\x0c\x85\x0b\xc7\xd3HK4\x93qEA\xa4\xf8\xe2\xaf\x94\xe1\x91\x87\x1f\xa1e\xcb\x96\xdcQ\xb92\x01\xfe\xfe\xdajkU\x90\xe5EA\x15\x80Jn\xe2\xbe\xa4\x11(\x8a 33\x13\xab\xc5JZz\x1aVE!5%\x99\xc3G\x8eb6\x9b\xc9\xc8\xc8\xe0\xe0\xc1\x83\x1c\xd8\xbf\x9f\xcc\xccL\xb7\xbdj\xd7\xa7\x93$\to\x1fo\xa2\xa3\xa2i\xda\xa4\t\xe1\x11\x11\xe8t:*V\xac@\xd9\xa8(t\xb2\x0e???\xf4z=\xdeF\xa3\xa6 n\xb4|\\\xef\xa2hp\xbdJA\x14\x80\xc8cK\xc8\x01\x03\x060y\xf2d*V\xac\xe8\x9a\xa4\xe1zE\x07\xf2H\xba\x16JT\x01\x08`\xc2\x84\xf1\x8c{7\xe7\x8e`\xc3\x87\x0fc\xe2\xa4I\x18J\xd1\x8e`\xd7\xc2\xd1\xa3GY\xb4h\x113\x7fz\x97\xb6\x03k\x10V\xce\x17\x8b\xd5\xea\xd4\xd3R\xc5\xb6*\x1a5Qi\xef\x1d\xd9\x8e\t@FF6H\\:\x9e\xcc?\xcbci\xde\xbf<\x915\x020gX5\xbb\x8e\x96S\x80\xdeG\xc7\xc5\x9f*0\xf1\xfd\xc9T\xafQ\x03\xa1(\x1c;v\x94\x91o\x8dDn|\x88\x88J~X\xcc\x8a\xbd\x97\x96\x95f\xe1\xd4\x9e\xab4\xea\x19\xad.\xe6\x92\xd5\xe5\x03;\x97\\a\xdcs3\xe8\xd5\xab\x17\xb2$\x91\x9e\x99\xc1\xb7\xdf~\xcb\xb7\x9b\xc7\xd1\xa8g\x14\x96,\xf5\x19\xf4:\x1d\xdb\x17\x9f\xa5q\xf7h|\xc3t\x982\xb4\xf5\n\x92\xed9\x1db%I\x12\xde>2iW,\x1c\xdaz\x9e\xba\xde\xf7\xf2\xcc\xb3\xcf\xd0\xb8qc\xbc\xbc\xd4\x08\x9f\x17.\\`\xef\xde\xbd\xfc\xfd\xf7\xdf\xfc\xf8\xe3\x8f\x88\x8aG\xa9\xd6\xe4\x0e\x82"\xbc\xf1\t\xd4c\xf4\xd3\xa3(\xaa\xfb\xabm\xcf\x03[\xc0;$\xb4En\xd9\x9b\xd4\xa8\x1b\x91H\x18\x8c:\x042\x99\xa9f\xcc\xe9VR/\xa5s\xfa\xe89\x12wGq\xff\x80\xc1\xdcuW\x07\x9a4iBxx\xb8v\xb3\xb9sm\n\xa0t\x91\x96\x96\xc6\xe5\xcb\x97\x89\x899Obb\x12\'N\x9c !!\x81m\xdb\xb6\xb1u\xf3f2M&\xd7\xaf\x14\x1az\xbd\x9e&M\x9b\xd2\xbd[7"#\xcbR\xf9\x8e;(S&\x82\xb2QQ\x84\x87\x87\xabsU\x85!\xd5\x8a\x89\x1bU\x00\xd9[B\xde\xce{\x02O\x98\xa0)\x80[c\x04\x80M\xbc\x0b\xc1\xe5\xcb\t,\xfe\xdf">\xff\xfaS\xca41Q\xbdM\x18\x92,\xa1X\xac\x0e\xe2\xdd\xa1\xc2\xdb~ux#\xaa\x02P\xd3d\x9dD\xf2\x15\x13\xdb\xbe?C\x93\x1eQD\xd7\x0b$3\xcd\xaaE\xe4\xcc^\x88ujS\x16\xaf\x0f\x98J\xc7\xbb\xefF\'\xcbX\xac\x16\xd6\xaeY\xcb\xdb\x1f\xbfL\xfd\xfe\xde\x18\x03du\xabJ\xcd<\xf5\xdf\x96K4\xea\x19\x8d)CAo\x90I\xba\x94\xc9\xe9\x9f\xfd\xf8\xdf\x0fK\xa9X\xa1"B@Zj\x1a\x9f\xcf\xf8\x9c]\x19\xf3\t\xac\xa4\x03!0\xfa\x1b9\xb9\xe32\xa6t35\xdbG`6\xa9\xf7\xa0Y\xc3\xb4\xe7\x11HB\xf58B\x86s\xbb\x93\xb9\xb8\xcf\xc0S\xf7\xbd@\xff\xfe\xf7\x13\x11\x1e\x81$I\xa4\xa5\xa5\xb2n\xdd:V\xaeZ\xc9\xe6\x1d\x1b)\xdb\xccB\x95\x96a\x18\xbc%5\xb2\xa9\xad\xc1)\xaa\xd7\x93\xe3\x1c\x81\xfa\xbbvA\xfb\x85\xb3\xff\xb6\xc5\x9b\xb2)\\\xbd^u\x83\xb5\x98\x04f\x13\x9c\xdcq\x95\xb8\xdd\x82>\x9d\xefc\xd0\xfd\xf7\xd3\xe8\xce;1z\x19s\x95E\xa5U\x018>9\xa8\xdeY\x16\xab\x95\x94\xe4\x14\x0e\x1f\xfa\x8f\x7fv\xee\xe4\xdc\xb9s\x9c={\x96\xc3\x87\x0fs\xfc\xf8q233]O\xa3\xa2\xc5s\xb2\x99~\xf4z=\xb2\xb6\xc3\x98m\x07-\x90\xb4\xf5:\x02\xa1(X\xacV,\x16\xd5d\x84&\xfcl\x1fw\xf8\xf9\xf9\xd1\xa8Q#\xaaT\xa9B\x85\x8a\x15i\xdc\xb8\x11w6\xbc\x93\xe8r\xe5\xf02\xa8[\x91J\x14\x8d\xef\xab\x93)\xf6:)\x98\x02\xc8oK\xc8\xdb^\x01\xe42\x02\x186\x8cI\x93o\x16\x05\xe0"\xc8\x1d0\x99Ll\xde\xbc\x99\xb9s\xbfdO\xecF\x9a\x0c\x88$ \xc4K\x9d\x1a\xb5\xf9\xc6k?\xb2\x1b\xafZA\x01d\x9d\xbaZXg\xd0\xfc\xe7%\x99\xd4\x04\x13{V\x9e\xa7R\xc3 *\xdd\x19\x86\xd5d\x9bjU+dVr\x16u\xd2\x1fa\xd8\xb0\xe1\xf8\xfb\xfb\xab\xc7\xb2\xb2\x989s&\xdfnx\x9b\xa6\xf7U\xb5\xfb\x93\n\x01\xfbW_\xa4q\xefh,V\x05\xbd^O\xcc\x7fITN\xed\xca\'\x9f|b\xef\x99\xa7\xa4\xa61\xed\xb3\xcf\xd8g\x98OP\x847\x92\x0e\xb2R\x15\xb6|\x7f\x92\xd6\x0fV\xc47\xc4\xa8\xda\xfdmF-\xa1y@\xe9\xd5QMz\x82\x89\x7f\xd7$\xd1\xbazo\x1e~\xe8\x11\x1a7Q{\xfd\x99\x19\x99\x1c:r\x84\x1f\xbe\xff\x9e\xa5\xeb\xbe\xa7l]=\xb5\xee*C@\x88\x9e\xac\x0c\xab\xa6L\xb3\x0bG\x9d?\xd6\xe6T\xec\xe5\x95-\xf0\x9dm\xcb\xb64\xc77\xa4\xfd&\xa1\x9eL\x08\x8c~\x06R\xaf\x989\xb2\xe9"IG\xfd\x19?f\x12\x9d;wV\xcb\xce\xa6\xd1\x1c(]\n\xc0\xe1\xc9\x84*p2\xd23HLJ\xe2\xf8\xf1#\xfc\xfa\xeb\x1a\xb6l\xd9\xc2\xe9\xd3\xa7\xb9z\xf5*&\x93\xc9\xad\xe02\x1a\x8d\x04\x05\x05\xd1\xf0\xce\x86\x84\x04\x87\xe0\xe5\xe5Et\xb9h|||\x90e\x19_\x1f_|||0x\x190\xe8\xf4\xe8\xb5\xb5!V\x8b\x05\xabb\xc5b\xb6\x90\x91\x99Ijj\xaa\xda\x96\x85\xaa|.]\xbcHRr2&\x93\x89c\xc7\x8fq\xe2\xf8\t222\\/\x8fN\xa7#0(\x88\xe8\xa8(Z\xb4lA\xff~\xfdi\xd0\xa0\x01\xa1\xa1\xa1\xf6{(m\xb8+GW\x14\xa1\xf0\x93G\x01\xb8G\xe4a\x02\x1a:t(S\xa6\x94\xaeM\xe1\xaf\x17!\x04\xc7\x8f\x1fg\xcd\x9a5\xcc\x9c9\x83\xf0\x96\x89Tj\x18NPYo,fE\x13\x9c\xf6\xdc\xd9\x02JQ{\xcdV\x0bx\x19%\xd2\xae\x9a\xf1\xf2\xd6\xe3\x13\xaa\xc7j\xb1p`\xe5e|CuTn\x12\x8a\x10j\xcf\x18@gP\xb8\xfcGyf\x8f\xff\x862\x91\x91\xb6\x9b &6\x96)\x93?fW\xca"\xeav,\x83\xc9\xa4\x86s\xf8k\xe9y\x9a\xdd[\x1eI\x96\xd0\xeb\x0c\xac\xfbz\x0f\x0b\'o\xa7Y\xb3f\xf6^RbR\x12\x1f\xbc\xf7\x01\xe7\xca\xffB`\x88\x0f\x92"\xb1\xf3\x97\x18\xc2\xef\xf0\xa5J\xb3P\xccYj\xc4Pu~C\xed\xb5\xcb:\x99\xc4\x8b\x19\x9c\xdcu\x81\xe0+Mx\xfa\xa9\xa7i\xd3\xa6\xad\xea\xe2)I\x9c;w\x8e_V\xfc\xc2\x8ci3\xf0\xa9{\x81\x9a\xad\xa3\t\x8a\xf4\xd6"Yj\xc5a_\xf8\xa6\x9e\xdbu\xaeD\x80=$\x85\xa3J\xc8\x16\x8b\xd9\xf3\x0f*j\xd9J\xb6L\xdaye\x9d\x8c^\'\x11w<\x95\xbd?%\xf2\xce\xb0\x8fx\xf0\xc1\x071\x1a\xbd\xb5\t\xcelJ\xa30\x12B!6.\x8e\x13\xc7O\xb0{\xcf\x1eV\xaeX\xc1\x86\r\x1b\\\xb39\xd1\xbcy\x0b\xee\xbe\xe7n\xc2\xc3\xc2\xa8]\xbb6>>>\x94\x89\x8c 0 \x08\xbd^GHH(F\xa31\xfb\x0b\xb6b\x10\xb6\x7f\x9c\xcb\x05\xd4\xd1\x98\xad\xf3b\xb5ZILL$##\x03\x8b\xc5L\xdc\x85\x8b$\'\'s\xea\xe4\tb\xe3\xe2X\xbfn=;v\xecp=\x85\x9d\xf6\x1d:\xd0\xabgO\x9a7oN\xed:\xb5\xed\xa3\xc5\xd2BA\x14\x80\x10\x82e\xcb\x973\xc8\xa3\x00\xdc3~\xfc\x04\xc6\x8ds\xe3\x054t(\x93\xa7\xdc\x1c^@\xb9a\x1bf\xaa\x93\xb5j0\xb5\xc3G\x0f\xf3\xdd\xfc\xafY\xf9\xfb2"jB\x8dv\x11\x04\x85{\x91\x91aFX\xb1\x0b0\xfbJ[\xd0\xa2\xc3\xa9\xe1\x9b\xadf\x89\xe4\xb8\x0c\x12.\xa4Q&\xda\x8f\x7f~\x89#8\xda\x8b;{Eb\xc9R]\xed\x85\xacp\xfe@,ot_\xc4\xdd\xf7\xdcmo4V\x8b\x85\xbd\xfb\xf6\xf1\xf6;\xa3\xa1\xd6\x11*5\t\x01\x05\xfeZr\x96\xa6}\xa2A\xa7#=\xce\x8a\xe1pc&O\x9eLxD\x84\xbd\x89\xc7\xc5\xc5\xd1\xa9S\'\x9a\r7b\xf4\x16\xc4\x9fLc\xdf\x9a\x0btz\xb1\x1a\x16\x8b\xa2M\xf4\x82$)\x18\xbc\rd&Y8\xf6\xe7U,\xe7"xt\xe0\x93t\xeb\xd6\x8d\xe8\xa8h\x8c\xdeF\xcc\x163[\xb7l\xe5\xfb\xef\xbfg\xd5\xb6\x05\xb4|\xb02\x91\xd5\xfc\x10\x9a\x89\x07\xcdUV\x95#6A\xe38\x91l\x1b!\xd9D\xbc*\xc4\xb5\xc2r\x12\xf9\x8e\x02\xc9\x86\xa3\x12\x11\x80\xa4M\xa6\xcb\xc8\xe8\xbcd.\x9dO\xe1\xec\xf2\x00>\x9e\xfc)m\xda\xb4A\xef\xb2\x1a\xb9\xb4(\x00\xa1\xadw\x88\x89\x89a\xd5\xaaUl\xdf\xbe\x9d\x9d;wr\xe8\xd0!{\x1eI\x92\xf0\xf6\xf6&,,\x8cJ\x95*\xd1\xb2e\x0b\xda\xb4iK\xddzu\t\x0c\x08\xc4\xc7\xc7\x17\xbd^\x87\xb7\x8f\x0f\xbabx.!\x04\x99Y\x99X\xccf\xd2\xd2\xd2\xb8\x92p\x85={\xf7\xb0f\xcd\x1a\x0e\x1c8\xc8\x99\xb3gIIN\xc6j\xb5\xda\xbfS\xbdz\r\xda\xb7oG\xa7N\x9d\xb8\xab\xe3]\x94\x89(\xe3\xe4)VR\x14T\x01\xe45\x070i\xf2d*\xd8\xdc\x9eK\x00\xdd\xd8\xb1c\xdfu=X\x1c\xd8^\xde\xd6-[\xf9c\xd3\x1f9\n\xb3u\xeb\xd6t\xee\xdc\xb9\xd44\xb6\x02\xa1\t\x12\xdb\xcb\xb4\xff\xd4\xdc\xdf\xf4z=Qe\xa3h\xdb\xb6\x035+\xd5%\xf1\x94`\xef\x9a\x18\xce\x9e=\x8fO\x98 8\xc4\x1b\x05\xd5\x9dQX\xb5\x11\x00\x92&\xd5\xc1b\xb6\xa2X\x15\xfc\xc2\x0c\x84\x96\xf3A\xb1\x82\xd5\xaa\x90p&\x95\x8b\'\xd3)[-\x00Y\xb3\xcfJ:\x03\xa6s!\xb4j\xd1\xca\x1e\x16Z\x96e\xca\x96\x8d"\xb2L$?-\xf8\x1d\xdf\x083\xa1e}9\xf1\xd7\x15B\xcb\xfb\x11R\xc6\x9b\xbf\x16\xc6\xf3\xd4\xe0\x97h\xd8\xa0\xa1\x93\xf9->!\x81\x8f?\xff\x80\x16\xf7Ea\xcd\x94\xd8\xb2\xe0\x0cm\xee\xaf\x88\xc1G\x87\xb0\x82N\'!\x0b\x05\xabUpb{"Y{\xee\xa0W\xab\x87yu\xd8\x08\xbat\xeeLXX\x18z\xbd\x8e+W\xae\xb2p\xe1B>\x980\x8e\xe4\xb0\x83\xb4~\xb0\x12\x81e\x0c(V\xf5\x19\xd59\\\xc7\xa6\x9d\xfdS\x96$5\x90\x9d\xfd\xaePCJ84\x1e\xe7T5|\xb6Z\x86h\xa6\x1cuRZ[\x8a\xa0e\xb3\x99\xd7\xd4QTP\x19#g\xce\x9dE\x97\x1cJ\xa3\xc6w\xe2\xe3\xed\xebd\x05*\xb6\xc6\x9a\x9b\x84\x13\x90e2\x11\x17w\x81%K\x16\xf3\xea\xab\xaf\xf1\xe3\x8f\x0b\xd9\xbd{\x17\xf1\xf1\xf1\x00xyy\xd1\xbau\x1bz\xf7\xee\xcd\x03\x0f<\xc0+\xaf\xbc\xc2\xd0\xe1\xc3\xe8\xde\xad;u\xea\xd6&,,\x1c\x7f\xbf\x00\xbc\x8dF\xbc\xbc\xbc\\vY+|l\x9d!I\x920\xe8\r\x18\x8dF\xfc\xfd\xfd\t\x8f\x88\xa0^\xddz\xf4\xee\xd5\x9b^\xbd{\xd1\xacY3*V\xacHXX\x18\x17.\\ ++\x8b+W\x12\xd8\xbbw/\x1b6l`\xf7\xee\xdd\xf8\xfa\xfa\x11\x1a\x1a\x8a\xb7\xd1[\xdd-\xaf\x84p\xedX\xe4\xc6\xe1\xc3\x87X\xbcx\xb1\xeba\xea\xd6\xadK\x97.\x9d\t\n\nvM*6JL\x01\xd8\xaa\xdb\x96\xad[\xd9\xecF\x01\xb4h\xd1\xe2\xa6S\x00\xb6\xb5\xa7\xd9\xbdOY\xf3TQ\xc5\x92M\x11\x18\xbd\xbc\xa8Y\xb3&-Z\xb4\xa0i\xe3\xa6\x04)\x15\xd8\xf2\xfdq\xb6o\xda\x8f\xb7\xd1\x80l\x95\xd0{\xeb\xf0\x0e0 \xc9\xda\x02Zl\xb3\xab`5\x0b,f\x81\xdeK&\xb2\x8a?Q5\x03\xc8\xbcj\xe1\xf4\xfe\xabDV\xf6Gg\xd0\xe1\xe5kd\xf7\xd6#\xdc\xdd\xa2;!\xc1!\xf6{\x94$\x89\n\x15*\xe2\xa3\xf7\xe7\x87\xf9\xf3\xa8\xd48\x14K\xa6\xc0\'\xd0\x80\xac\xd3!\x9d\xaa\xca\xc3\x0f=B\x99\xc8H\xbb\xa0\x13B\xb0k\xd7?l\x8f\xfd\x8e\xaa\x8d\xca\xb1\x7f\xed\x05\xc2\xef\xf0%\xaaF\x00\x16\xb3\xc0\xe0-a1Y9\x7f0\x91\x7fW\x9a\x18\xd8f(\xcf\x0fy\x9e\xae]\xbb\x12U\xb6\xac\x16\xccM\xe6\xf8\xf1\xe3,X\xb0\x80\xf7g\xbcB\xfd\xde\x01Tm\x11\x8a\xacC\xddZR\xab\x10\x92\x94m\xe6\xb1\xc9?\x81\x1a\x9cN\xd6K\x08kv9\xd8\xbe\xa4\xfe\xc8\x96\x96\xce{3hG%\x87E\x17\xea\x01\xb0-PS\xdf\x8c}\xfeX\xb1\x08\x02"\xbdY1\xfdo\x1e\x7f\xf2q\x82\x02\x03\xd5\x9b\xd3(2\x05\xe0*\xf0\x9d.\xa3&\n!8}\xfa4\xabW\xaf\xe1\x9dw\xdea\xfa\xf4\xe9\\\xb8p\x01\xb3\x83\x1b\xf5\x88\x11#\x18\xfa\xd2K<\xf2\xf0#\xdc\xdb\xff^\xda\xb5oO\x85\xf2\xe5\xf16\x1a\xd1\xeb\xf4HRv\xd0B\xb7\n\xa6\x08p_fj{\x90d\x19\x9dNOpp\x10\xb5k\xd7\xa1u\xeb\xd6\xb4j\xdd\x9a\xb6m\xdarG\xe5J\xec\xdf\xbf\x9f\xac\xac,2339v\xec\x18\xff\xfb\xdf".]\xba\x84\x97\x97\x17\xc1\xc1\xc1\xf8\xf9\xfb\xb9(\xfe\xe2\xa1\xa0\n\xe0\xd0\xa1\xc3\xb9*\x80\xce]\xba\x10\x14\x14\xe4\x9aTl\x94\xa8\tH\xd8\xe6\x00\xc6\xbd\x97s\x1d@)\x99\x04vgw\xb6!4\x0f\x07\x93\xd9\x04\x02.]\xba\xc4\x85\x8b\x17\xb0Z\xaddee\xe1\xe3\xed\xab\xda\x97\x1d{\xa8\x92\x84\xd1h$""\x82\xd0\xd0P$!a2gq\xe5\xea\x15\xf6\xec\xdf\xc3\x8a\x9fW\xb2\xfd\xafm\xe8\xa3R\xf0\x0eU\x88\xae\xedOxE_|\x83\xbcHO1\xab\x02P{c\x12\xa8\x93\xad:\x90e8\xbd\xf3\n\xb1\xc7\xd3i\xd6\'\x9a\x80\xf2F\x0e\xaeO`P\x9d\x91<\xf3\xcc3\xaa\xdc\xb3\x0bN5\xac\xc3\x94)S\xd8\x1a;\x8f\xda\xad\xca\x92x5\x93\x94\xcbYT3u\xe7\xbd\xf7\xde#88\xc4\xdeh\xadV+\x9fM\xfd\x94?\x93g\xa2X\xf5\x1c\xf9\xf32-\x06V\xc4\xe0\xad\xc3\xe0-\x11\xf3o\ng\xff1\xd1\xec\x8e\xce<\xfd\xd43\xd4\xaaY\x1b\xff@?\xbbt1\x99L\xec\xd9\xbd\x9bY\xb3g\xb1?\xe17\x1au/\x8b\xdeWB\xb1\xed4\xa3\tg\xa1\xd9\xf2%\xdbb!I\x80"\xe1\xe5\xab\xc3\x94fF\x00^\xde:\xac&\xb0\xa2\xa8B\xdb\xf1\xfdh\xe15\x9c=c\xd5\\8,R\xb2\x15\x9e-\xc5V\x96h\xe9B\x02I/\xb3p\xf8a\x8e\x1e<\xa1.\x84s\xa0\xb8;%\x02\x10B!))\x85M\x7f\xfc\xce\x82\x05?\xb2t\xe9\x12\x14-\xccF\x80\xbf?\xb5j\xd7\xe6\x91G\x1e\xa5c\xc7\xbb(_\xae<~\xfe\xaa\xaf\xfduc+\xc3k GY\x16\x84<\xaec\xb1ZHKM#\xee\xc2\x05~\xfe\xe9\'\x16,X\xc0\xb1\xa3\xc7\xc8\xcc\xcc@\x00wT\xaa\xc4=\x9d:1\xf8\xa1\xc1\xb4i\xdd\x06///\xad\xcef\x9f\xd46\xf2(\n\\;\xad\xee\xc8\xcf\x04t\xdb\xcf\x01L\x18?\x9ew\xc7\xe5\x9c\x04\x1e6|(\x93\'M\xb9\xb1J\\\x088\xd6O\x81\xba22##\x93\xd4\x94\x14b\xe3b9z\xf4\x18\x1b\xd6\xff\xc6\x89\x13\'IIK%)3\x1e\xdfH\t\x9d\x11t\xb2\x8c\xba\xb1\x99\x84\xa2\x80\xd5$0\xa5\tH\xf7\xc2\xcf\x10\x84\xbf\xaf?\xf5\xea6\xa0C\x87\xf6T\xa9V\x95\x88\x88p\x8c^Fbcc8v\xec8\xbbv\xed\xe6\xdc\xf9\xb3\x1c;s\x98\xcc\xa0STl\xe2CX\xd9p\xbc|e{dP\xc5\xa2\xd8\x83\x84\x1a\xbcd\x8e\xfd\x95\xc0\xa9\xbdWi\xf5`%\xfc\x83\x8c\xb0\xa3)\xef\x8e\x19Gd\x992N\rA\x08\xc1\xfe\xfd\xfbx\xe3\xad\x11X"\xcf\xaa\xabrcR\xe9\xdd\xec\x19F\xbf\xfd6\x92\xa4\x9a\\\x00.\'$0`P?\x82\xea%p\xe6\xbf\x04\xaa\xb4\x0f\xa4\\-\x7f\x12/e\xb1\x7fY:M\xef\xe8\xc8\xfd\xf7\xdfO\x87\x0ew94B\xf5\x1a\x99\x99Yl\xde\xb2\x99W\xdez\x892\xcd\xac\xd4l\x13\x84\xa2\x08m\x92\xd7V\xf5T%i\xd7\x07Z\x8a\xc1\xa0CX\xc0j\xb2b\xf0\x91P$u\x9f\x02\x81\x1a_I\xa8\x12>\xfb\x05I\x80\xa4N@+V5|\xb2\xdeW\x879S\xd1&\xc8\x155\x93\xd0\nL\xbb\x92j\x12rP\xf32\xe8d\x03\xf3\x9f\xfd\x97\xff\x0e\x1f\xa6b\xa5\x8aN2\xaa\xb8\x15@FF\x06G\x8e\x1e\xe1\x8b\xd9_\xb0l\xd92.]\xba\x04@hh(\xbdz\xf5\xa2\xd3=\xf7\xd0\xf1\x9e{(\x1b\x19\xa9\x8e\xb6T\xedV"\xe4!\xcf\xaf\x13\xd5\xe3\xcbj\xb5r\xfc\xd81V\xaf^\xc3\xba\xf5\xeb\xd8\xb8q#YYY\xda\x88\xb6\x02\xcf?\xf7\x1c\xf7?\xf0\x00\x15*T@\xaf\xd7\xe7\xd9q+,\n\xa2\x00\x14Eu\x03\x1d4(\xa7\x02\x180`\x00\x93\xa7L\xa1B\x85\nE|\xa7\xb9S\xe2\n`\xfc\xf8\xf1\x8cs\xa3\x00\x86\x0f\x1f\xce\xa4I\x93J\\\x01\xd8\xc8\xcc\xcc\xe0\xd4\x993\x1c?z\x8c\xc3\x87\x8f\xf0\xc3\xfc\x1f8pq?wv\x90\x08*S\x86\xc0\xf0\x00\xbc|U\x81\xe2\x1b\xe8\xa5\xc6\xeeGB\xa7\xd3\xbc\r\x15-\x02h\x9a\x15k\x96@\xd2\t\xb2\xd2\xad$]N\xe5\xe2\xf9K\x9c\xdc\tu\xc2\x1a\xd1\xad{7\x9a4mBTT\x14\x15\xcaW@\x11\nqq\xb1\x1c;v\x9cc\xc7\x8e\xb1h\xd1B\xa4\x8a\'\x88\xac\x1eL\x99J\xe1\xf8\x05ya\x0c\xd0!\x84\x84\x92\xa5\xa0\xf7\x969\xb1\xeb\n\x877_\xa0\xfbsu\x88\xdd\xaeg\xf8\xfd\xe3\xb9\xab\xe3]9B+\xa7\xa7\xa7\xf1\xf1\xc7\x9f\xf0\xe1\x07\x1f"y)x{\xf90\xf1\xa3)<\xf9\xc4\x13\x9a\xd9F59\xac^\xb3\x9a\x17^~\x16\x8c\x19Tm\xe6O\xfdNe8\xf8g\x0c\x863uyl\xf0Sth\xdf\x812\x91e\x9c\xce/\x84\xe0\xf2\xe5\xcb\xac]\xbb\x96w\'\xbfE\x9d>>\x94\xab\x15\x8c\xd9i\xa1\x91\xad\xea\xd9z\xe7\xea_\x06/\t\x83\x97\x8e\xc4\x0bYH\x80o\xa8\x01\xabU\xf3\'w\x94\xf9\xda/:\xbd\x84\xd1WGf\x9a\x05K\x16d\xa6da\xd0\xeb\xf0\x0e4`\xb6ZUs\x87C\xec\xa4\x1c\xd8\xa3\x1bH\xea\x9c\x82\x01\x12\xe3\xd2\xf1\xda\xdb\x94\xcf\xa6~FDD\x84S\xf6\xe2R\x00\x8a\xa2p\xee\xdc9\xd6\xfe\xb6\x8e\xf7\xdf\x1b\xc7\xf9\xf3\xe7\xedi\x83\x06\x0e\xe4\xc1\xc1\x83i\xde\xbc9QQe\x91l\xdb\xa8\xddj\xd8+\x86*"\xcd\x163\xc7O\x1cg\xd3\x1f\x9b\xf8x\xca\x14\x8e\x9f8a\xcf\xda\xa7O\x1f\x9e}\xe6\x19\xda\xb6kW,f\x95\x82(\x00!\x04\xcb\x96-e\xd0\xa0\xfb]\x93\xec\x93\xc0\x15Kp\x04Pbs\x00h\x85\xb3e\xeb\x166m\xda\x94\xa30[\xb6lY\xe2s\x00B\x082\xd23\xd8\xb5{\x17\xdf\x7f?\x9fo\xe7}\xcb7K\xa6s\xd1g\x17U:\xca\xb4\xec[\x81\xb25\xc2\x08\x8e\xf2\xc1?\xd4\x0b\xdf =^\xbe\x86\xec pV\xb5\xe7\x8f\xa2\xf6:m+R}\xfc\r\xf8\x04y\xe1\x1b\xecEh\x94?\x15jEp\xe7=\x91\x845\xb1p"i\'\xff\xfb\xdfb\xd6\xaf\xda\xc2\xd1#\xeaB\x9dJ\x95*\xd1\xbcE\x0b\x9a5m\xca\xc0\x81\x83\xb8\xbb\xe9@\xc2Em\xf6\xac=\xcf\xb9#\x89\\8}\x15\xc5\xa4\x10\x18n\xc4\xcbWOpYo\x8c>^\xec]\x11\x87\xc5j!=^\xa6y\xf3\x16\xf8\xf8\xf88\x8d\x02\xf4zu\xb1\xcd\xaa_W\x91\x9c\x98BP`0C\x86\x0c\xa1r\xe5*\xf6^drr2\xdf\x7f\xf7\x1d\x1b~\xfb\x1d\x9d\xaf\x95\xba\x1d\xa2\xf9w}\x12w\x06\xf5e\xf4[\xef\xd0\xbauk\x82C\x82T[\xae\xc3\xb9\xe3/\xc73k\xe6,f-\xf8\x84f\x0f\x05\x11V\xc9\x07\xc5\x9c\xed\xd7\xaffU\x05\xae\xbd\xc3\xaa\x03\xff\x10/\xd2\xe2\xb38\xf1\xcfU\x82"}\xf0\x0b\xd1c\xb1j{!h_\x91%\xd0{\xc9\xf8\x06\x19\xb0\x9a\x14\xd2\xaffqb\xe7\x15b\xfeK\xc1\xa0\xd3\x11\x1a\xed\x8b\xceO\x87P\x14t\x06\x19\x14\x9bG\x95v\x7f\xb6\xdbt\x14\xfc\x9269\x0c\xf8\xf8\x1b\xd8\xb54\x86\xfb\xbb?E\xd3fMs\xacP\xcd\xcb\xa4P\x18=`!\x04\xa6\xac,\xd6\xaf_\xcf\xc7\x1f\x7f\xcc\x17\xb3gs\xe9\xd2%dY\xa6^\xbd\xba\x8c{w\x1cC\x87\r\xa5q\xe3\xc6\x04\x05\xa9e\x7f\xcb"\xd9\xde\x8f\xa6\xecu:\xc2\xc3\xc3\xa9W\xaf\x1e\xdd\xbbwG\xaf3p\xfa\xf4i\xd2\xd3\xd39r\xe4\x08[\xb7m%%%\x95\xa8\xa8(\xd5\xc4Z\x84\xfbH\x17d\x0e@\x08\xc1\x91#G\xdc\xce\x01\xd4\xa9S\x87N\xf7t"8\xf86\x9c\x04F\x8b\x8b\xb3u\xcbV6\xbbQ\x00\xcd\x9b7\xa7K\x97.\xc5\xaa\x00ls\x84\x8aPHKK\xe3\xc0\xc1\x83L\x9f>\x9d\x8f\xe7|\xc8\xa9\xcc\xbf\x08i\x92J\x93>\x91\x84W\xf6\xd3B\x0b(H\xc8vA\x96\x1d\xedX\xadp\x92\xac\t:[\x05\xb6]\x07\xd5\x04"\xb4^\xad@\xdd/E\x96e\x82\xcb\xf8P\xa3m(!5-\\\xc88\xca\xa6\xbf\xd6\xb3t\xf1/\x1c\xdaw\x98\xa0\xa0`*T\xac@\xe5*Uh\xda\xb4\x19\xf7\xf6\xe9G\xeb;;\x12LEL\xe7\x03\xf8\xef\xf7\xab\x1c\xdd\x7f\x82\x80\xb2\xdeT\xaa\x1d\x84E\x08\x8e\xec\xb8\xc4\xc6\xd5\x7f\xe2\xebk\xa4A\x83\xfax\x19\xbd\xed\xa6\x1dI\x02oo\x1f~\xfd\xf5Wbcc\t\x08\x08\xe0\x91\x87\x1f\xa6B\x85r \x04\x16\xa1\xb0g\xf7nf\x7f\xf1\x05\xb11\xb1\xe8t\x12\xa9\xa7}x\xf2\xbe\xe1\xbc\xf2\xca\xab\x94+\x17\xad\x85\xeaP\x1b\x99\xa2\xad\x06\x8d\x8b\x8d\xe5\x93O?\xe1\xe7]3h\xf7d\x05\xbc\xfc\xb2M7v\x01)\xa1*G\t\xf4\x06\x1d22\xe6L\x0bG\xb7\xc4cJ\xb1R\xady\x18z\xa3\x84\xd5\xa2\xa8\x9b\xe2H \x90\xd0\xc9`N\xb7\x90\x14\x9b\xc9\xbf\x1b\xe2\xb9p<\x85\xacL\x85\xf2\xf5\x82\xa8\xd6,\x04\xdf\x10\x03f\x8b\xa2\xd6\'\xa1\x96\xab$9L\xfe:\n\x03M8\xd8\x92\x90%tz\x89\xc4\xd8tL\xffV\xe0\xf1\xc7\x9f\xa0b\xa5J9L\t9\x05J\xb6\xd8wM\xb9V,f3g\xce\x9ca\xc6\x8c\x99\xbc\xfb\xeeX\xb6\xef\xd8AfF\x06\xd5\xabW\xe7\xb1\xc7\x1e\xe3\xbd\xf7\xfe\xcf\xdey\xc7WU\xa4\xff\xff}\xce\xad\xe9\t$!\t\t\t\xbd\xf7&\x88\x80RT\xec\x05\x14\x0b\x96]\xcb\xda\xfbZ\xd6\xb5!(\xb8k\xdb]w\xed\x8abC@i\x82\xa2\x02J\x93\xdeK \x90\x10H\xef\xed\x96s\xe6\xf7\xc7\x9c{s[\n5\xeeo\xbf\x9f\xd7\xeb&\xf7N;s\xa6<\xcf\xcc3\xcf<\xcf\xf3\x9c\x7f\xfe\xf9\xc4\xc6\xc6\xa0\xfe\x7f\xec,\xa71F\xaa\xa0`1\x9b\x89o\x9d\xc0\xf0\xb3\x87\xd3\xb3G\x0ft]g\xd7\xae]\x94\x97W\xb0v\xedZV\xacXAZ\xbb4\x12\x12\xe3\xb1\xd9\xec!\xfa\xec\xe4\xd1\x1c\x06\x00\xb0gO\xc3\x87\xc0\xe7\x9f\x7f\xfe\x19\xd9\xad4\x84\x16d\x00\xb2CV\xfd\xf2\x0b?\xff\x1cJ\x0bh\x08\xe3\xc7\x9f\x7fF\x19\x00H\x9b\xf2[\xb6n\xe3\xb3\xd9\x9f\xf1\xf4_\x9f\xa2\xbc\xcdz:\x9f\x17M\xfb\xfe\xad\x88\x88\x93\x1e\xbct\x97\xa87\xa4-<D\xc5#\x9b0\xbe\x07\x8d7\xc5\x90Jz\xc8\x8dg\x19\xec\x93P\x80\xd0\x05.\x87\x8e\xd9\xa4\x12\x97\x1aFR\x97\x08Zu\xd0)\x12{\x99\xf5\xf9\x7f\xd8\xb6\xea0\x9a\xa6c\xb7\xdbHHL %%\x85\xc1C\x860l\xd8p\xce\x1a4\x8c\x8e\x89\xfd)\xda\xa22\x7f\xfe\x02R;\xc5\x91\x90\x1a\x87\xaa\xc2\xbc\xd9K\t\x0f\x8f\xa0{\xb7nDx\x1d\xad\xc8\x03\xe9U\xabV\xb1c\xc7\x0e"##\xb9\xf6\xdakHm\x9b\x06\x8a\xa0\xa2\xbc\x82\xff\xfc\xe7?\xcc\x9b7\x0f\x80\xd6\xb1I<\xf7\xd7\x17\xb8\xee\xba\xeb\x88\x8e\x8e6\x16\xef\xf5\xf5\xd7t\x8d\x1d\xdb\xb7\xf3\xc2\xd4\xa9l\xab\xfa\x8a\x01\x97\xb4\x03U\xa0i\xbe\xda6\xf5*\x98&\xb3\x8a\xd5\xaeR\x91\xef\xa0p\x7f5\x9a\x1b\xd2z\xc6\x92\xda;\n\x018\x1c\x1a\xc2\xadSY\xec\xa4<\xaf\x96\xa2\x035\x1c\xd9Q\x81\xbbN\xc7\x16a\xa1m\x8f\x08\xda\xf6\x8c!\xa1C\x04\n\xe0\xa8s\xa3\xb9}\x0f\x85=\xcf\x94\xef\xaa\xe0\xd1\x10\xf2V\xd9\xa7V\x86kNT\xb6..\xe4\xf2so\xe4\x92K/\xc1j\x917\xa1\xfd\xd2\x07\x11\x93\xc0\xdf\xcd\x85?\x99+**f\xf9\xf2\xe5\xbc0\xf5\x05\xde}\xf7]*++\x01\x98|\xedd\x1ey\xec\x11n\xbc\xe1F\xd2\xd3\xd3Q\x0c\xd5\xc7\xe0z\x9c\t4F\x9a\x1b\x8b;>4]\x8a\xec5\xbb\xddF\xb7n\xdd\x194p\x10\x1d;v$33\x93\xa2\xa2"\xf2\xf3\xf3\xf9\xec\xb3\xcf@\xc8\xb3\x92\x84xi\x94.\x10\'S\xe3\xe62\x80\xdd\x8d0\x80\xf1-\xac\x05\xd4\x82\x0c@b\xd5\xaaU!w\x00g\x9a\x01\x08!(--\xe1\xd3\xd9\x9f\xf2\xda\xeb\xaf\xf2[\xf1|\xfa\\\x13MR\xc7hlv\xb3<l\xd5\x8c%\xab\xaa\x18\n\xebFf\x0f\xbd\xf1\x9b\x90\xbeF\t<\xfa\xe8\xbei\x8c2\xbc\xc3\xaf\x9e0)\xc8\xf3J\xdd%@\x08\xccaV\xa2\x12\xect\x18\x1c\x8f\xa3U6s\xe6\xcec\xc3O\xbb)\xc8/ 11\x91\xd8\x98\x18"""IJN\xa6[\xf7n\x9cs\xceH\xc6\x0c\xba\x82\xdd+\x8b\xd8\xbdg\x1f\x15\xc5UT\x15\xbb\xd8\xb9s\'&\xb3\x89\xde\xbdz\xfb\x89\x83\xca\xcaJY\xb8p!\x91\x91\x91L\x9a4\x89\xb4\xb44t]\xe7\x87\x1f~\xe0\x81\x07\x1e\x00 )9\x89\x17\xa7N\xe5\xea\xab\xaf&22\x12a\x1c\xc4z\xe0v\xbb\xf9m\xc3o<\xfb\xec\xb3\x94Dl\xa4\xcf\x05)\xe8B\xfa\x04\xf0\x88\x93<P\x00k\x98\x8aV\xab\xb1\xfd\xfb\x02Z\xc7\x85\x11\xdf9\n[\xa4\x89\xca\xd2:\xf6\xae.!wW9\xd9\xdb*\xc8\xdf[\x85\xa3B\xa7U\xdbp\xdat\x88 \xa1s\x04\xd1\t\xe1\x84\xc7\x99QL*\x9a&\x0cs\xdb\xf5\xea\xb6\xf5O\xf1|\xf31]\xeciv#\x85\xdcY\xc8\xef\xaa\xaarlo\x05U\xdbZ\xf3\xe7\xc7\x1e\'%%%$\x91\r\x15vb\x90\xf5r\xbb]l\xde\xb2\x99\x7f\xbd\xf5/\xa6M\x9b\xc6\xd6-[Q\x14\x85\xd4\xd4T\x1e\xff\xf3\xe3<\xfc\xc8\xc3\xf4\xef\xd7\x8f\xf0\xf0\x08#W\xf3\xeb\xd0 \x81\xf3\x8b\x90\xdaV\x82z\xdb=\xba\xaeK\x0fp\x9a\xf4\x02\xe7r\xb9p8\x1cTVVQT\\HIi\t\xc5EE\x94\x96\x96RQQ\x8e\xcb\xe9BU\xcd\xe8\xba\x86\xa6i\xe8\x86s\x1e\xcf\xa3\xbcW\xf3\x8cg\x1aK\x01\xa3\x1a!k\xd84\x8cE\x8c\xa2(\xc4\xb5\x8a\xa3w\xef\xde\x0c\x1b6\x8c#Gr\xc8\xcd=\x8a\xdb\xed\xe6\xd7_\x7fe\xef\x9e\xbd\xe8B#9%\x99\x88\xc8H\x9f>\xaf\x7f\xf2\x89\x1c\x1a7\x8f\x01\x88Fw\x00-\xcd\x00Z\xfc\x10x\xfa\xb4\xe9<\x1b\xe2&\xf0\x199\x046.\x1d9\xea\x1c\x1c<x\x90\x7f\xbc\xf9&?n\x9bO\xb7\xf3\xa2i\xdb#\x12\xa7\xc3#\xb3\x96\xc3\xa3~\xd8\xd47\x99\xf0\xe1\x05\x9e\xc5\xbfQ\xb4\x17rXy\x06\xbc\xfc/\x89\xa2\x9cx\xf5\xe3NA1)(z\xbdh\xa8\x1e2\x9f\xaa*X\xc2L\x14fU\x91\xb5\xae\x1c\n\x92\xb8\xfd\xa6\xbb\xb8\xe8\xa2\x8bh\xdd\xba5\x16\xabEj\xd4\x08\xa8\xaa\xaef\xc1\x82oY\xb8`!?\xff\xf4\x13\xf9\x86\xf6\xc8\x93O<\xce\xedw\xdeIrR\n\x8a\x02s\xe7\xcd\xe5\xdak\xae%>>\x9e\x8f>\xfa\x88\xb1\xe3\xc6r\xe8\xd0an\xb9\xf9\x16V\xaf\xfe\x95V\xadZ\xf1\xc4\x13O\xf0\x87?\xfe\x81\xe8\xa8\xe8\xfa\x1b\xb9\x80\xdb\xadQQY\xc9\xb7\x0b\xbe\xe5\xcd\xb7\xffN\xc2\xe0\n:\x0em\x85\xb3\xd6\xb8\xc9\xe9\xb3\xcb1[\x14\xdc.PMplw%k\xe7d\x13\x97\x12\x8e\n\x94\x17\xd4Q[\xa5\xd1\xb6k$\xe5\xf9u$dD3\xe8\x92\xb6\x84\xc7\xaa8\xdd\xbaW\xff\xdf\xed2n\x0b+\xba\xa1\xf6i\x10\x01\x0c\xf5\xcd\xa0\xd1\\\xdf)\xf5m\xef\t\x97\xb27\x81\x8a\xc9\xacR]\xe1b\xc3\xdb\xd5<\xfd\xf8s\\{\xcd\xb5\r\x8e\xbd\x93Y\x94\xc8\x11d|\x17\x82\xd2\xd22\x96-\xfb\x8eg\x9e}\x96\x03\x99\x07\xd04\rUQ\x98x\xcd$\xa6L\xb9\x89\x91#\xce!"2\x12\x827.\xc1\xf0-\xbc\x19\xd0u\x1d\x87\xc3!\r\xc6UV\xa2kn\x0e\x1f\xce&//\x8f\xd2\xb2R\xd6\xad]K\xd6\xa1\xc3\xb8].?\xa6PWW\x87\xa6K"\xaf M~\xd8ma\x84\x85IQ\x8b\xa2($\xb5Ib\xe4\xa8\x91$$$`\xb7\xdb\xe9\xd4\xa9\x13QQQ\xd8\xec6\xecV;V\x9bE\x9ay\t\xd8E\x9e,4]#??\x9f9_\xcd\xe1\x8b/\xbe`\xf5\xea\xd5\x00DGG3f\xcc\x18\xee\xbf\xff~\x06\x0c\x18@Td\x94O[\t4MG=N\xdf\x10\x81\x8b\xd6Ph\xca\x16\xd0\x8c\x193Hk\x97vJ\xdb\xe0x\xf0\xbbe\x00\xf7\xdew/\xaf\xcc<\xfdj\xa0uu\x0e~\xfa\xe9G^ye&\xe5q\x1b\xe95*\x83\xc8D\x0b\x8ejM\x12:\x05\xcf\xf5TIF\x0cz\xe2!+\xb2\xdb<\x84\xc8\x83\x00"c\x84I\xd4\x87\xc9\xfc\x06S1\xfc\xf0\xba\x1c\xba\xa1>*0\xdbT\xeaj4\xf9<C7\xdes\x9ci\xb1*(:\x94\x1cs\xb0gu&q\x15#\xb8\xe5\xe6\x9b\x18v\xd60\xdaw\xec\xe85_\xe0\xd64r\xb2\xb3Y\xbdf\r3f\xced\xc7\xb6m\x00\\\x7f\xfd\xf5\xdc}\xcf\xdd\xf4\xeb\xd7\x8fE\x8b\x162\xf1\xeaIDFF\xf2\xf7\xbf\xfd\x8dk\xae\xbd\x96\x7f\xfd\xeb_<\xf3\xcc38\x9dN^x\xe1y\xfe\xf4\xa7\xbb\x88\x8d\x8bE1\xce+*+*8z\xf4(\xfb\xf6\xef\xe3\xdb\xf9\x0bX\xb2\xe9#\xce\x99\xd4\x99\xb8\xb40I\xa4\x15\x81\xd5\xaeb\xb6\xab8*u\x9c\xb5n\x9cU\x1a\xc5\xb95d\xae-%k]%\x11mT:\x9f\x15G\xab4;\xb6\x08\x0bQ\xf16"c\xac8\xaa\\\x92\xb1\x14\xb9\x08\x0f\xb7\xa2\t\x1dM\x08\xc2\xa3\xcd\x84E\x9a\xb0G[1\x9b\x15\xeaj5\xdcn\x81\xe2Y\xdd\x1b\x06\xee@2vi%\xb5\x9e\x91zb\xeb\xfbJ\x1e.\x9bmf\x8a\x0f\xd7\xb2{\x9e\x9b\xc7\xee{\x8a\x89\x13\'\x12\x1e\x16n\xa4\t\xc6I1\x00C/]\xd7u\xf6\xee\xdd\xcb\xac\x8fg1\xfd\xa5\xe9\xde\xf8\xde\xbdzq\xf7\xdd\xf70\xfe\x82\xf1d\xa4g\x1c\x17Aj\x08\xbe|\xc1\xe9rQX\x90O^^\x1e\x85\x85E\x1c>|\x88\x82\x82B\x16.X\xc0\xc6\x8d\x1b\xd1\x9aA\xd4N\x14\xc9\xc9\xc9\\y\xe5\x15\xa4\xa7\xa7\xd3\xa1}\x07\x12\xdb\xb4!==\x9d\xf8\xd6\xad\t\x8f\x90\xbb\x1b/\x841\xef\x8e\x930zr8\x1c\x0e\xb6m\xdb\xca\xfc\xf9\xdf0m\xda4\xbf4\x7f\xfb\xdb\xdf\xb8\xf8\xe2\x8b\xe9\xd2\xa5\x8b\xccs\x82w\x05\x9a\xc5\x00\x9at\x08\xf3\xbf~\x0f\xa0\x11\x8f`\xa7\x95\x01\x08A\x9d\xc3\xc1\xf2\xe5\xcby\xf6\xf9g\x08\xefUH\xb7sZ\xa1\xa8&\xdc\x9a\xdb;\xect\xe3\xb0\x1a<*\x88\xf2\xbb\xef_\x0c\r\x12\x14\xcf\xb6\xd0\x9f\x9f\xcb\x06\x96\x7f=\xf9\x0c-v\x9f\xb4F>\xb3\x82pkd\xae+\xa5u\xaa\x9d\xa4N\x91(f\x13uUn\x04\n\xaaQ\x0ftC\x12eVA\x08J\x8e:\xd8\xfeC>=\xe2Fr\xce\x88Q\\t\xd1E\xa4\xa7\xa7c\xb1\x98\x11\x02\\.\x17\x07\x0f\x1e\xe4\x9dw\xde\xe1\xd3O?\xa5\xa8\xa8\x88sF\x9e\xc3MSn"/?\x9f\xbf<\xf5\x14&\xb3\x89G\x1fy\x94\x11#F\xf0\xd8c\x8f\xb1k\xd7.\xae\xbe\xfajf\xcc\x98A\x9b\xa46\xb8\\.\x8e\xe4\x1ca\xeb\xb6m\xac]\xb3\x86\x83Y\x07\xd9\x99\xbd\x8e\x1e\xe7\x87\x93\xd2/\x96\xc8\x08\x1b\x9aKP]\xe9\xc4\xe5\xd49\xb6\xb7\x82\xeaR\x07\xb5\xa5\xd2AK\xe1\xa1\x1aZ\xb5\xb1s\xf6\xa4v\xb4\xca\x88\xc0\x85\x0bG\xb5\x8e\xdb)P\x10\x98\xcc*\xba&\x9b\xc2lQp\xeb\x02\xc5\xad`\t\x03\x93\xc5D]\x99\x86IQ8z\xb0\x92\xdaZ\x07\xad\xdb\x86\x13\x11g56Q\xf5\xb2\x1c!|vX\x863\x9a\xfa\x16\x97\xed\xed\x19\xf4f\xb3J\xde\xbe*v-\xaa\xe5\xa1\xdb\x9e\xe4\xc6\x1bn0\xc4\x04\r\x13\x84`\x06\xd0|B%\x04h\xba\x9bM\x1b7\xf1\xe2\xb4i,\\\xb0\x00]\xd7\t\x0b\x0b\xe3\xba\xeb\'s\xef=\xf7\xd1\xb5kW\xecv{`\xd6\xe3\x86\x10\xd2g\x82\xc3\xe9$7\xfb\x08\x1b6n`\xc7\x8e\x1d\x1c:\x94\xc5\xc1\x83\x07\xd9\xb6};5\xd55\xde\xf4\x8a"\xcd\x85\xd8\xed6"#\xa5Y\x91\x88\x88\x08\x12\x13\x13\xb1x\xcc3\x1b\xed\xa2\x9a\xe4\xb8\x13B\xb6\xb7G\x1c\xe2v\xb9))-\xa1\xbc\xbc\x1cM\xd3\xa8\xa9\xa9\xa1\xa6\xba\x06M\x97\x0e\x8c<0\xa9*\x1d:ub@\xbf\xfe\xb4Mm\xcb\xa0A\x83\xe8\xdb\xb7/\x1d;v\x94\xcf\xf2\xf8\xbfn\xa4\x1f\x9a\x82.t\xca\xcb\xcaY\xb1r\x05\xff\xfc\xc7?X\xb9r\x15N\xa7\x13UU\xb9\xe6\x9ak\xf8\xd3]w1t\xc8P\xac\xd6\x1338\xd9,\x06 \x04\xf3\xff\xef"X\xc3h\x90\x01\xdc{/\xaf\xbcr\x1a\x18\x80\x00\x1d\x1dG\x9d\x83\x9f~\xfe\x91\x07\x1e\xbd\x97\xf6cU:\x0c\x8d\xc5\xedt\xcb\x01\x8dg\xf5XO(<d\xdb\x88\xf5N\xf7\xfa_\xf2\xbf[\xd7\xa5V\x8b\xaa\xa0\xb9t\x831(~\x1e\xb2\xe4\xea\xb4\x9ef\xc8\xfc\x86\x16\x11\xd2\x10\x9a5\xccBuQ\x1d\xb9{\xcb\t\x8f\xb1\x91\xd0\xde\x8e-\xda\x8a\xb3Z\xa0\xeb\x1a\x08\x15\xb3YAQeA&\x93\t\xd5\x02y{+\xc9\xdfYGmn\x0c\xa3\x87\x8dg\xe4\x88\x91t\xed\xda\x95\xc8H\xe9\xe0\xdd\xedv\xf1\xd9g\x9f\xf1\xd1\xc7\x1f\xb2k\xc7n\xa2\xa2\xa2\x10\x08\xaa*\xab\x00HLL$::\x9a\xcc\xccL\xcc&3\x0f>\xf4 \x03\x06\x0c`\xdd\xfau\xac_\xb7\x9e\xe2\xe2b\xf2\xf2\xf2\xa9\xa9\xab\xe4\xbc?\xa4\x11\x19\x1f\x86-\xd6\xc4\xc1-\xc5\x14\x1c\xa8Ds\xea\xd8#-X\xc2\x14\xac\xe1\x16l\x11R\xed5\xa5{\x14f\x8b\x82\xcd\xaaRW\xab\x91\x97Y\t\x8aJ|\xbb\x08"b-\xb8t\x81*@\xb8\xa51;O\xcbH*P\xdf\x11\x8a\n&\x9b\t[\xb4\tPp\xd5i8+\xddR\x94\xe6%\xc4\x81\xab~\xcf>\xcb\xf3WH\xc3o&3\xbb\x7f.\xa4p}\x18\x8f>\xfc\x08\x93\'_\xdf,\xc2\x1b\xcc\x00\x9a\x0f\x87\xd3\xc1\xcf?\xfd\xc4\x13O<\xc9\xe6\xcd\x9b\x01\x189j$7O\xb9\x99+\xae\xbcB\xaau"\xc7\xcc\x89\xc2\xad\xb9\xa5\xe3\x97\xc2"\xb6\xef\xd8\xc1\xe7\x9f}\xc6\xb6\xad[9z\xec\x98\x14\xf5\xf8\x10\xae\xc8\xc8H\x06\x0f\x1eLRR\x12)))DGE\x91\x94\x94D\xb7n\xdd\xb0\x87\x85\x11\x1e\x1eNjj*aaar\xa4\x1a\xd5\x12Hn\xe6\x1d\xff\xb2\xd28j\xeb\xc8\xcb\xcf\xa3\xbc\xbc\x0c\xb7K#\xf3@&Y\x87\x0eQUYInn.\x85\x85\x85\xac[\xb7\x8e\x8a\x8a\no\x1d\x14E%&&\x9a\xc4\xc4D\xce\x1bs\x1ec\xc7\x8ca\xe8Y\xc3\x88\x8d\x8d!<,\x1cE\xf5=m:\x1e\xc8\xda\tM\xe3\xd0\xe1lf\x7f\xfa\ts\xe7\xcec\xf3\x96\xcd(\x8aJzz;\xde~\xfb?\xde\x0b\x8c\xcdg\xe3\x12\xcda\x00B\x17\xcc\x9b\xdf0\x03h\xe9{\x00\xbfc\x06p\x0f\xaf\xbc\xf2\xb7S\xcf\x00\x8cI\xb8\xf4\xbb\xa5<\xf8\xe4=\xf4\xb8\xc4Fj\xafX)\xe7T\xe4\xa0\xf6\x88\x12\xfc\x89?>$\xc4\x80\'\xbd\'\xd6,\xd5\x14\xcb\x8f:\x10B\x10\x11m\xc3\x12)e\xf6(\x86\xbcY\x91\x97\xc3\xccf\x05\xcd\x13&@\xd1$\xe1B(h\x9a\xc0\xed\x14hn\x81\xc5\x10\x03\xa95:f7\xa8f;\xd9\x87Kpi\x1aqIVP\x05\xc2\xad T\x15\x97\xa6\x11\x11a\xc5\\kb\xdb\xca"\xb6|w\xcc\xfb\x06\xd1\xb11\x8c\x1b3\x96~\xfd\xfa\xa1i\x1a\xdf/_\xca\xaf\xab\xd6x\xe3\x8f\x171\x19\x16\xc6\\\xdb\x1e\xdd\xa4SQYGx\xac\x9d\xf00\x15T\xb0\x86[\xb0\x86I\xcb\x9a\xd60\x13\xae:\r\xb7C\x12\x0b\x0ccw\xb6p\x13.\xa7\x8e\xbbV\xc7Q\xe9\xc2\x16k\xc6d\x92\x8ej\xea\x0f\x0f\x03%\xa3>\xad\xef\xf9j4\x9b\xa7\xb7\xea\t\xbfg\xe5_\x9f\x1cC\xf1\xcal1Q\x9e_\xc7\x9e\x15E\xc4\xd7\xf4\xe5\x81\x07\x1f`\xd4\xa8Q^\xdf\tM\xa1)\x06\xe0\xa9\x7f\xe0\xea\xd5\xad\xb9\xf9\xf5\x97_y\xfc\xf1\xc7\xbdf\x90\xa7L\x99\xc2\xbd\xf7\xdeK\x9f>}\x0c\xdf\x0b\xc7G\x86|S\xbb\\.rrr\xd8\xbdg7k\xd6\xae\xe5\x93Y\x9fr\xf8PV@\x0e0\x9b,\xdcu\xcf]t\xed\xda\x95\xce\x9d:\x91\x94\x94D\\\\\x1c\xb1\xb1qDD\x84\xfb\x10t\x03AU\xf2\r0Z\xdc\xd0\xae\x12\x1e\xe56O\xac\xd1\x16u\xb5u\x94\x94\x96PV^NNv6\x07\x0e\x1c\xe0\xc7\x1f\x973w\xae\xd40\x0b\xc4\xc0\x01\x03\xb8v\xf2d\x06\x0c\x18H\xcf\x9e\xddi\xd3&)0I3Q_\xd7\xba\xdaZ\xd6\xae[\xc7\x1bo\xbc\xe1\xd5lKMK\xe5\x9d\xb7\xdfa\xec\xd8\xb1\xde\xcb\x8f\xa1s\x07\xa3Y\x0c\xa0\ts\xd0-m\r\xb4\xe5\x19@\x03\xa6 \xee\xb9\xf7\x1e\xfev\x1a\x18\x80\xd3\xe5\xe4\xc7\xe5?\xf2\xd7\xe7\x9e&\xf1\x9c2\xd2\xfa\xc5\xe2vi 4\x04\x8a\xd4\x1b\xf7\x8a\x0c\xa8\x1f\xe0\xde\xc3^\x83(\t\x8f@^\x12mE\x91\t\x14\xab\xb4\xc6YS\xe8\xe0\xf0\xe6r\\N\r\xc5\x82\x8cC\x01EAU\xa4\x83\x17]\xc8m\xba\x10\x80\xa6\xa0\xaa\xc8\x1b\xbd\x9a@s\n\xeaj\xdc\xd2\xe8\x9bE%\xccf\xc6\xaa*\xb8\xdd\nG\x0fW\xe0t\xb9\x8d\x89&\xd0\xf5\xfa\xc1\xa3\x1a\xfa\xa9u\x15:%\xd9\x8ez\x8a\xd8L(\x8aBrJ2\x95\x95\x95\x84\x87\x87\xd3\xa9S\'\xdc.7\xa5\xa5\xa5TWW\xa3i\x1a.\x97\x8b:S\x05\x03\xc6&\x92\xd2+\x92\xe8v\xf2^\x84V\xe76\x0e\nu\xe3pU\xc8{\x12\x01\xb4\x02\xa4:\xa6Pe[\xa8&\x10H\x03w\x92\xf6(!\x88\x8cq\x17@\xe0w\xda^\xaf\x8ce\xec\x140DA>\x0f\x13F:\xd5\xacb2)dm(e\xe7\x8f%\xf4K\x1f\xc9\xd4\xa9S\xe9\xde\xad\xbbW\xc4\xd1\x1c4\xc9\x00tQ\xef\x0c\xdd\x13&\x04;v\xee\xe0\xcf\x8f=\xce\x92%\x8bQ\x14\x85\x9bo\xbe\x99\xbf\xfc\xe5/ddd4YfC\x10B\xe0vk\xec\xde\xbb\x9bo\xe7\x7f\xc3\xe6\xcd\x9bY\xb3v\r\xc7\x8eJ\xe6o2\xa9\xc4\xc6\xc6\xd1\xb6m[.\x9cp!\xe7\x8c8\x87n\xdd\xbbIk\x9av;\xf6f\xea\xc87F\x08\x8f\x17B\xe8\xd49\x1cTVTRRR\xcc\x9a5kY\xbcx\t\xeb\xd7\xaf\xa3\xb0\xb0\xd0\xcfaL\xc7\x8e\x1d\x199r$\xe3\xc7\x8d\xe3\xdcs\xcf\xa5u|\x02&\x93jL\xc8\x86\xea\xd4pm5M\xe3\xc0\x81\x03\xbc\xfa\xda\xab\xfc\xfb\xad\x7f\x03p\xc9\xc5\x173\xfd\xa5\x97\xe8\xde\xbd{P?4v>\xd0\\\x06\xd0\x98-\xa0\x193f\x92\x96\xf6\x7f\x0c \x98\x01\xdcs\x0f\x7f\xfb\xdb\xa9g\x00\x87\x0e\x1f\xe6\xf1\xc7\x9f\xa4\xa0\xd5*\xba\x9d\x93\x80\xb3\xce\xd0V14}\xea\xc9\xbe\xe7\xbbO\xc7\xa8\xd2\x06\x8f\xa2\x187LU\x81\xbbNC\xe8&\xac6\x13B\xe8TU8\x11\x0e\x05\x93\xa2\x10\x19c\xa3`\xaf\x93_\xbe:@E\x99\x03\xb7C\xa0;\x05B\x87\xf8NV\x12\xdb\x87\xa3\x98\x14\xc2[[0\x9b\xe4h6[\xcd$\xb6\x8f \xa2\xb5\x19\xcc*v\x9b\x05\xbd\xda\xcc\xdeU\xc7h\xd3%\x8c\xe8\x14;\x96p\x13\xaa!\xfbG\x084\x1d\x10\xd2\xf7\xadI1Q\x91\xeff\xff\xea"\xf2\xd6)\xf4\xec\xd9\xdbp\xc2m\xc1f\xb5\x12\x9f\x10\x8f\xddn\xe7\xd7_\x7f\xe5\xc7\x1f\x7f\xac\x7f7\x03f\x8b\x99;\xee\xba\x8d\xa2\xc2b\x1c5N\x1ey\xe4\x11222(-\x93\x0c@\xd7\x05E\x85\x85l\xdc\xb8\x91\xfc\xfc|r\x8f\xe5\xb2\xf3\xe0F"\xd3\xea\xe8~n\n\xd1\xf1f\x14U\x9a\xa9\x16\x8a!\xd2\xf1\xae\x88%\xa1\xf7\x88\xbc@\xa6\xc1X5\xd6\xd3m/\xb7\xad\x0f\x97,\xc0G\xb8\xe3\xdb3\xb2\x10\xe1\xb9\xcd+\xa87\x03\x8d\xbc\x81\x8d\xaeP]Z\xc7\xe6\xc5G\x88\xa8\xec\xc2\xb5\x93\xae\xe5\xfa\xebo\xa0\x8d\x8f\xd5Sh\x94vx\x11H$\x02Q\xff\x8e\xf5(++\xe3\x96[na\xfe\xfc\xf9\x98L&\xa6L\xb9\x91\xc7\x1f\x7f\x82\xce\x9d;\xcb7k\xe2\x99A\x10\x82\xea\x9ajrr\x8e\xf0\xe5W_\xf1\xe9\'\x9fp(\xeb\x904L\x08\xc4\xc6\xc52d\xc8\x10\xfa\xf7\xef\xcfYC\xcfb\xd0\xa0A\xd2\\\x87j\x92\xf7\x1d~\x07\xa6#<\xe3B\xd7u\x9c.\x17\xdb\xb6nc\xc3\xc6\xdf\xf8m\xfdol\xdc\xb8\x91\x1d;v\x00rQ\x12\x1d\x1dM\xef\xde\xbd\xb9\xf3\xce;9\xf7\xbc\xf3H4\x9c\xcd\x9f\x08\xe1\x14Bp47\x97i\xd3\xa7\xf3\xaf\x7f\xfd\x0b\x80;\xee\xfc\x13\xcf=\xf7W\x12\x13\x02\xc6C#8\x15\x0c@\x9e\x01\xa46=\xe8N\x13Z\x9c\x014\xe4\x12\xf2T3\x00!\x04N\xa7\x93W_}\x95\xcf7\xfc\x85\xa1\x97\xf7\x91\x9e\xb8\x8c\t\xef{@%\xbb\xc2\x97\xfc\x0b\x14U\xc5\x16a\xc2\xed\x12\xb8\x1dn\x1c\x95n\x84\xa6Sv\xac\x0e4\x13\xaaE\xfaE\x15(h\xc20\xaf\xac\x80jQ\t\x8f2\xa3W\xa8l\xfd\xa9\x80#\x99\xa5\x0c\xba4\x99\xb8\x94p\xa2Z\xdb\x10:X\xa3M\x98M\x92\xc0\xb9\\\xbad*B\xfa\x00.\xca\xaa\xa5`\x7f%\x9d\xcf\x8e\xc7\x16m\xc6Y[\xdfN\xbe\xec\xcad1\xe1\xaa\xd6\xc8\xd9U\xca\xb1\x15\x11L\xb9\xfef\x06\x0f\x1eLjj*6\xbb\x1d\xb3\xc9\x84\xd9b!::\x8a\xea\xaaj\xfe\xf2\xf4\xd3\xbc\xf3\xf6\xdb\xc4\xc4\xc40r\xe4H\x16,X\xe0\xf3\xf6\x82\x8b/\xbd\x84\x1f\xbe\xff\x81\xbe\xbd\xfb\xf0\xc2\xd4\x17\x18>|86\xbb\r\x93jB\x08ax{*\xa5\xa4\xb4\x94CY\x87\xd9\xbd{7\x9b7o\xe1\xe3/>d\xe0E\nm:e\x10\x9dd\xc7\x1ei!,R\x8a{4\x97\xc7`v\xfd"^\x18\xdf\xeb\x89\xbcD}o\xd4\x8bq<\xe1\xfeS\xc5(\xc8\xe7\xbbd\x00\nf\xab\xf4]\\r\xa4\x86#;\xcaY\xf7Y\x01w\xdey\x07\xd7^;\x99\x81\x83\x066\xaa\xe9\xd3\x18\x9ab\x00^\xb1\xa01\xae\xf2\xf2\xf2\x99?\x7f\x1eO=\xf5\x14\xa5\xa5\xa5\x0c\x1a4\x88\xf7\xde{\x9f>}z\x07\xe6\x0c\xf9\x86\x81p8\x1c\xec\xda\xb9\x93\x1f\x96/\xe7o\x7f\xfb\x1b\xf9\xf9\xf9~\xf1\xf7\xdcs7\xe7\x9e{\x1e\xdd\xbbw\xa3}\xfb\xf6\xd8\xeda\xde\xb8\xa6Ko.\x02{\xec8\xd1@Et]\xa7\xa0\xa0\x80\xcc\xfd\xfb\xd9\xb0q#\x9f\x7f\xfe9\xeb\xd6\xad\xf3K3i\xd25L\x9e|-C\x87\x0e%99\xd9\x1b\xde@\x91!\xa1\xeb:;v\xec\xe0\xa9\xbf\xfc\x85\x85\xc6\xd8\x7f\xe7\xddw\xb8\xfa\xaa\xab\x9bm\x9a\xe1T0\x80\xff\x13\x015\xb8\x03\xb8\x9b\xbf\xfd\xed\xef\xa7\x8e\x01 \xd8\xbau\x1b7\xdd>\x99nWAt\x8a\r\xe1\x92\xaf\xae\xa0\x18\xb6\xfc\x8d\xdf\xdec\x00\xa9\x93\x8f\xe1\xff\xd6U\xab\x11\xd1\xdaLX\xb8\x99\x8ab\x07\xf6H\x0bV\xbb\x94[\xcb\xdb\xae\xd2\xc4\xb1\xd0\xc1l6\xa1\x9aM\x98,\x90\xbb\xb3\x82\xcac\x0e\xe2\xda\x85\x11\x93*]\xecI\xb9\xbf\\\xb5z\xb7\x99\x1e&\xa4H\'\xdc\xfb~- \xccf&m@+t\x0c\x06\xe3C4\xd1\xe5\xedUK\x98\x89\x82\x83\xb5l_R\xc8\xe8\x9e\x17q\xdb\x1f\xef\xa6{\xf7\xee\xf2\x00M\xf1\x88g\xe5\x00+++\xe3\xdd\xf7\xdee\xfa\xb4\xe9\x94\x95\x95q\xff\x03\x0f\xd0\xaf__n\xb9\xf9\x16\x00.\xbe\xf8"n\xbf\xfdv\xbe\xfdv\x01\xb3g\xcf\xa6\xa6\xa6\x86a\xc3\x86q\xcd\xb5\xd7r\xc3\xf5\xd7\x13\x1b\x13\x83\xa2z\x94QA\x17\n\x08\xa9O^]]\xc3\xd1\xa39\xec\xdc\xb1\x8b\x8fg\xcd\xe2hY\x16\xc4U\x10\x95\xa8\x92\xd2%\x92\x98\x14;a\xd1f\xeaj4t\x0f\xf3\xf5\xd1\xce1^\xdd;\x95\x05\xf2@\xdc\xb3M\xf0\xec\xca|wj\xbeS^\xa6W\x08\x8f\xb6R[\xe1\xa4\xf4H-Y\x9b\xcb(\xddk\xa6o\xe7!\xdc|\xf3M\x0c\x1b6\x9c\xd6\xad\xe3\r{M\x81\xeb\xf4\xe6\xa1)\x06P\xbf\xe3Qp8\x1c|\xf9\xe5\x97<\xf9\xe4\x93^cn\xcb\x96.\xe5\xdc\xf3\xce;\xee\xb1\xadk\x1a{\xf7\xef\xe7\xcb/\xbe`\xf9\xf2\xe5\xacZ\xb5\n\x8c\xfa\xb4m\xdb\x96\t\x13.\xe2\x96[n\xa1s\xe7NDGEa\xb64\xa2\xdd\x12\x82Z\xcaZ\x1b"M\xe4\xbd\x8a\xc0\xef>270Z\xd0\x9b\x0fC\x0cj\x8co\xfc.\xe6\xf9d\x0b\xd8\xa1\x85\x82\x10\x82\xda\xdaZJJ\x8aY\xbe\xfcG\xdey\xe7]v\xec\xd8NEy\x05\x02A\xabV\xad\xb8\xfc\xf2\xcb\xb9\xfa\xea\xab\x199r$aa\xe1(\x1e\xc7>\xcd\x84\xa6i,^\xbc\x84\xdbn\xfb#\xf9\xf9\xf9\xf4\xec\xd5\x93\x97\xa6\xbf\xc4\xf9\xe7\x9f\x8f\xa5\xb1\xb63p*\x18\xc0\xcc\x993ImAk\xa0-~\x13\xf8\x97U\xab\xf8y\xe5\xca\xa0\xc6\x1c:t\xe8\xa9\xbd\t,`\xcd\x9a\xd5\xac\xdc\xfd\r\xed\x07F\x19D\xc5 \xa8\xc6\xa0U\x08\x18\x93\x9e1*\xc0\x1a\xa1\x12\xde\xda\x8c\x10\n\x9a[`\xb2\xa8\x08\x01n\x97@\xd7dG\x0b\xa4\xf6\x90\xb4\xd1\xaf\xa2;t2\xd7\x97`\xb7\xab\xb4\x1f\x12\x835\xca\x8c\xee6\xe6\x98\xc7\xee\xbc!#W\xe4\xac\x90~ym&\xb26\x96\x92\xd4)\x9a\xc4\xae\x91\xb8\xdd\x9a4gl\xd4M\xfeS0\x99\xcd(B\xb0\xff\x97Rj6\xa5p\xef-\x8fq\xef\xbd\xf7\xd3\xbe}\xfbz\xb3\xcc\x8a\xf1fBj\xfa|\xf1\xc5\x17<\xf7\xdcs\x94\x95\x951f\xcc\x18\x9ex\xe2I\x8a\x8b\x8b\x99\xfb\xf5\xd7\x00\xf4\xed\xd3\x8f\x1bn\xb8\x91\xb1c\xc7\x10\x1d\x1dMnn.[\xb6la\xdd\xbau\x94\x97\x97\xd3\xa9S\'\xa9Qd2\x1bg\x1f\xf2\x11\x16\x8b\x85\xf0\xf0p\xda\xb4iC\xcf\x9e\xbd\xb8\xfc\xf2\xcb9w\xc48\xdaFu"\xac6\x89\xfc\xed\n\x9b\x96\xe4phw.6\xbb\xd4\xe7\xb7\x86[\x10\xba@\x8a\xcb\x15\xccV\x15\xb3I\x95\x04T\x00\xa8\x92\x01\x0b\x05\x84\x8a\xaa\xaa\xa8\xaa\x82\xd9\xac\xa2\x9aU\x84.5\xa7\x14\x05\xcc&P\x85\x8d\xe2\x82\nv\xfdp\x94\x8d_\x97\x12W\xd5\x8b\xc1\x9d\xc7\xf2\xc0\xdd\x0f\xf3\xa7?\xfd\x89~\xfd\xfa\x11\x19\x19\x81\xaa\xfa8D\tE\t\x9b@S\xab5O\xb4\x87\x01,[\xb6\x8c\xf9\xf3\xe7\x03\xf0\xf0\xc3\x8f0y\xf2d\xc2\xc3\xa5VMs\xe0r9)((`\xc1\x82\x05\xfc\xf5\xaf\x7f\xe5\xf3\xcf?\'+K\x1e\xee\x0e\x1d:\x94\xc9\xd7N\xe6\x89\'\x9f\xe0\x8f\x7f\xfc\x03\xed\xda\xa5\x11\x1e\x1e\xde\xb4\x9d \xe3\xd1^1\x8c\xa6SWW\x87\xd3\xe1\xc4\xe9tR]]CyY)\x15\x15\x95TVVRYYAUU\x15\x95\x15\x95TVUQUUEEE\x05e\xe5\xe5\x94W\x94SYYI\x957\xbc\x92\xea\xea*\x9cN\'\x9a\xae\xe3v9q:]hnM\xf6\xa1\xa7\xed\x1biGEQ\xb0X,DEE\xd3\xabwo.\xbax\x02\x9d:v\xc4f\xb7\x91\x9d}\x98\xf2\xf2r\xb6l\xd9\xc2\xb2e\xcb\xa8\xae\xae&11\x91\xa8\xc8H\xcc\xc7y\x96\xd3\xae]\x1a\x07\x0e\x1e`\xf3\xa6\xcd\x14\x17\x15c2\x99\x184hP\xb3\x8c\xec\xf9J\r\x1aCc7\x81\xc7\x8e\x1bGlLt\xb3\xc7\xc2\xa9\xc6\xefd\x07\xf0<n\xb7\xbfC\x98{\xef\xbd\x87\x99\xaf\xfc\xed\x14\xf9\x04\x96b\x81\x17^x\x81/V\xbf\xce\x90\x89mQ\x14\x15\xd5\xb86Z/d0\x9a\xc2\x87&\xd4\xafa\x0c\xd1\x82\x0cl\x00r\xb9\xad\xa8\x92\x80U\x97\xb8\x88heA\xb5(hN\xe1\xbd\xd8"\xd7\xb0\xfe\xa4G \x0fF\xadv\x95\xe2\xc3\xd5D\xb6\t\xc7lUq;uPt/\xd1G\x80P\xc1b6Q~\xd4\xc9\xa6\xef\xf722\xfd&n\xbb\xfdv\xfa\xf6\xed\x8b\xcdj\rY\xbf\xaa\xaa*\xbe\xf9\xe6\x1b\x1e}\xf41\xf2\xf3\xf3\xe8\xda\xb5\x0b3g\xbe\xc2\xd8\xb1cY\xb0p!\xd7L\x9a\x04\xc0\x84\t\x13x\xe5\x95W\xe8\xdc\xb93U\xd5U,\\\xb0\x90g\x9e}\x86\x83\x07\x0e\x02p\xe9e\x97q\xcd5\x93\x186l8i\xa9\xa9\x8d2h)\x12\x83\x9a\xeaj\x8e\x1c=\xca\xd1\xa3G9\xb0?\x93\r\x1b6\xb0\xe0\xdb\x05\xe4\x97\xe5\xd3e$\xb4\xed\x9a\x84%\xccFL\x82\r\xd5\xa4\xe0v\x08l\xd1R\x85\xd4l\x91m\xa3\xbbt\\u\n\x8e\x1a\x97\x0c\x10\x82\x9a\n\'u\x15\x1a\xd5U\x95l][J\xcd\x1e\x182p(\x17L8\x9f\x01\x03\x06\xd0\xa1c\x07\xd2\xd3\xd3\x89\x8ch\x9ev\x8f\x17M\xf0\x84\xc6\xde\x19Y5/m;z\xf4(\x0f<\xf0\x00_}\xf5\x15a\xe1a|\xfe\xd9\x17\\t\xd1\x84\x90vi\x02\xa1\xeb:\xc7\x8e\x1dc\xc3o\xbf\xf1\xd5\x9c9|\xfa\xe9\xa7\xde8\x93\xcd\xca\xb3\x7fy\x9a\xf1\xe3\xc7\xd3\xa3G\x0f\x1f\r\xa6\xc6+\xef\xd64\xca\xcb\xca\xa8\xa9\xae\xa6\xba\xa6\x9acG\x8f\xa1\xe9:U\xd5Ud\x1f\xce6\xceyt\x8e\x1e\xcdc\xcd\x9a_\xd9\xbf\x7f?n\x87\x1b\x1d\x1d\xa1Jf\xac*\x8a4\xcej\x98{\x10B\xa0\xaa\xaa!\x93\x87\xb8V\xad\xe9\xd8\xb1\x03=\xbau\'9%\x05\x8b\xd9\x8c\xc9d"*:\x9a\xb4\xb44\xc2\xc2\xc2\x08\x0b\x0b#99\x19\xab\xcdF|\xebV\xd8\xec\xf6\x06\xf7c\x9e\xf6\x14Bp\xf8p6k\xd7\xae\xe6\xddw\xdfe\xf9\xf2\xfa3\xac\xae]\xbb\xf2\xc7?\xfe\x91s\xcf;\x97n]\xba\x11\x11\x19p\xb1,\x14\x84\xd4\xd3_\xb7n-\xc3\x87\x0f\x07\xa0m\xdb\xb6\xbc\xfe\xfa\xeb\\z\xe9\xa5r\x17\xd0Hs\x06.ZC\xa19;\x80\xff\xed{\x00\x8d0\x80S\xa9\x06*\x04,\\\xb8\x80\'\xdf\xbc\x8b\x81\x13\xe3\xa4&\x81A\x90\xa5\x98\x81 \xc2\x1c\xf8\xdfC\x80\xe5\x0f\xdff3V\xefBA\x17\n\xaa\tlaPW-\x07\x88\xd0\xa4\xa8\xa6~ y\xf2\xca\x00\x8f\x08\xc4\xac\xaa\xb8]n\xccv3B\x97\xf7\x08P\xf0n\xab\xf1,\xe8Q\xd8\xf7k\x11\xf9\x1b\xcd\xdc\xf7\x87G\xb9\xe2\x8a+h\x93$\xd5\xe4\x14D\x90m\xf8\xca\xcaJ>\xf9\xe4\x13^{\xf552\x0fd\x12\x1d\x1d\xcd\xd4\xa9S\x992e\na\xe1a|=\xe7k\xae\xbb\xee:\x00:v\xec\xc0\xbb\xef\xbe\xc79#\xceAG\xc7\xe9p2o\xde<\xa6O\x9f\xce\xae]\xbb\x00\x88\x8fo\xcd\x981c\xb9\xe2\xca+\x18=z4\xf1\xad\xe2!\xd0W\xaf\x01]x\xc48\xf2w]\xad\x83\xea\xea*\xaak\xaa9t(\x8b\x9d\xbbv\xb2b\xc5*\x8e\x1d=JEe9e\x15%\x98#\x14\xc2\x13L\xd8"UT\xb3l\x1d\xdd%p\xd7(\xd4\x96j8*t\xac\xaa\x8dV\xad\xe2\x89\x8a\x8c\xe6\x82\xf3\xcf\xa7c\xa7\x8et\xed\xd6\x95\x98\x98X"\xc2#\x88\x88\x08?ec\'\x10M1\x00\x0ft]g\xe1\xc2\x85\\v\xd9e\x00\\x\xe1\x04\xfe\xfe\xf7\xbf\xd3\xad[W\xbf!\xe0\x1dg\x061ukn\xb2\x0fg\xb3f\xedZ\x96-]\xca\x8a\x15+\xc8\xc9\xc9AUU\xda\xb4i\xc3\x85\x17^\xc8\x9dw\xdeI\xb7n\xdd\x88\xf4\xb9\xb4\x16X\x8e\xae\xeb\xb8\xdcnrr\xb2\xc9;\x9a\xc7\xe6-\x9b)*.\xa6\xb0\xa0\x80\xd2\xd2R*++\xd8\xb2e\x0b\xe5\xe5\x15\xb8\xdc.\\N\xff9x:`\xb5Z\xb1Z\xad\xb4k\xd7\x8e.]\xbb\x10\x11\x1eA\xdb\xd4\xb6DFD2`\xc0\x00222H\xcf\xc8\xc0n\xb3\xd5\x1f\xf2\x06\x10a\xa7\xd3IAA>\xdf|\xbb\x80\x7f\xfd\xe3\x9f\x1c\xcc\xca\xa2\xaeNj\x0e\x8d\x1c5\x92Q#G1v\xecX\xfa\xf7\xef/wZ\xaa\x8a\xc9k\xbd\xd1\xbf,\x81\xa0\xa8\xa8\x98G\x1e}\x84\x8f?\xfc\x08\x80\x193fp\xf7\xddw\x13\x1e\x1e\xee7\xfb\x03\x11\xc8\x00B\xa5\x14B:\x84i\x88\x01\xcc\x989\x83\xb4\xd4\xb4&w\x1b\xa7\x0bg\x96\x01\xf8\xb4\x90g\xfb4}\xfat\x9e\x7f\xeey\\\x81\x0c\xe0\xbe{xe\xe6\xa9c\x00\xa0p\xe8P\x16\x7f\xb8\xed\x0f\xd8\x06\x1c\xa6m\xafX4\xa7T\xa5\x944\xb5^\x12\xed\xb7\xb3\xf3\xca*\x03\xbb\xb7^"\xed\tWd\xb0q~\x10\x02\x1e\x19\xaa\xa7<\x81A\xd1\xe5wE\x91\xccC\x17\x06\xd30\xb2\xa8\x8a\'\x8bBm\xb5\x8bM_\x1f\xa3g\xebs\xb9\xe1\x86\x1b\x193f,\x11\xe1\xe1>j\x87\xf5\x87\xa0B\x08J\x8a\x8b\xf9\xfc\x8b/x\xe1\xf9\xe7),*\x02\xe0\x8e;\xee\xe0\xb9\xe7\x9e#!!!`\x85\xa2`\xb5Y\xf9\xea\x8b/\xb9\xe0\xc2\xf3\xa5\xad\x16\xc3\xd0\xdb\xbau\xeb\x981c\x06\x0b\x17.\x04\xe4\x01u\xd7\xc1)to;\x94k\']K\xff\x01\x03HH\x88\xc7b\xb1`\xb5Y\xbdZ&\x02\x0cS\r\x81\x8d"\xeb\xa9\xeb:\x9a\xa6Q[WKQa\x11%%\xc5\xde\xfa\xeb\xba\x8e\xcb\xe5\xf2\xb6\xb1\xcd*-s*(\xd8\xecV\x92\x92\x92\x88\x89\x89\xc5l6K\xb1N\xd03N\x0f\x9a\xcb\x004M\xe3\xf5\xd7_\xe7\xe1\x87\x1f\x06c\xf7\xf4\xe2\x8bS\xe9\xda\xb9\x8b\xd7\x0e\x8e\xae\x0b\x1c\x0e\x07.\x97\x93\xca\xcaJ\xb6o\xdf\xce\xcf+~f\xf1\xa2\xc5\xe4\xe4\xe4P^^\x0e@jj*\x17_|1WO\xbc\x9a\xb3\x86\x9ee\\\x90\x92\xef+\x84@sk8\x9c\xd2onuu5\xbb\xf7\xecf\xcf\x9e\xbd,[\xba\x94C\x87\xb2(*.\xa1\xa2\xbc\x1c\xa7\xcb\x890lM\xfd^\xe0\xe9\xbb\xd8\xd88Z\xb5\x8a\xa3mj[.\xb9\xe4\x12\xce\x1ev6\x1d:\xb4\'2:*H]U\x08\x1d\x97\xcbMV\xd6!\x16/^\xc4k\xaf\xbdFvv6 \xcf\xcf\x92\x92\x92\xe8\xd1\xb3\'7M\x99B\xdf>}ILL$<<\x02\x8b\xd5b\xf8w\x90c\xd3QW\xc7\x91\xdc\\^x\xfeyf\xcd\x9a\x05\xc6%\xd4g\x9ey\x86\xd6\xad[{\x9f\x17\n\x81\x0c \x14\x9a\xb7\x03\xf8_\xd6\x02\x9a6\x8d\xe7\x9f{\x0eW\x88\x8b`\xa7N\x04${[\xd3\xdc\xcc\x9b?\x9f?<}\r\x17\xdd\xde\x87\xe8D\x15w\x9d\x86\xa2\xd4;t\x0f\xd9\x0f\x9e\x16\xf2\xc6\xf9\x12\xed\xfa\xe8 \x1e\xe1\x07\xb9B\xaeO"\t\xa0oz\x05\xc31\xb9\x87/\x18\xe9,6\x0b\xce*\x17GvV\x90\xf5\xbd\x89\xdb\xffp\'\x97\\|\t]:w1\xf2\x07\x1c~\x19\xb2\xe7\xfd\xc6\x81\xe1\x8b/\xbe\xe8\x8d\x1a;v,o\xbe\xf9\xa6\xd7\x0e\x8a\xae\x0b\xbe\xf8\xe23\x1ez\xf1F\xdc\x15vJ*\xea\x98\xfe\xe4t\xee\xbe\xe7n"\x0c\xeb\x93\x00\x9a\xae\xb3f\xf5j\x1e{\xec1\xd6\xad_\x07B\xe1\xc2;\xbb\x10\x9bn\xe2\xc0\xee\xc3\x14m\x8a\xe3\xfa+n\xa1s\xe7.\xb4MK\xa5]j*\xb1qq\xc4\xc6D\xa3\xaa\xaa\xe1\x92Q\xaa\xaa\x86\x82\xaf(\x01$\x03\xf0\x84{_N\xc8?\n\xc6\xb1\x80!\xff?\xd3h.\x03p8\x1c<\xfb\xec\xb3\xbc\xf4\xd2K\x804\x9f0\xe5\x86)L\x980A\xdaUR\x15\x9c\x0e\'\x99\x07\x0ep$\'\x87%\x8b\x17\xb3c\xe7\xce\xc0b\x00\xe8\xde\xbd;\xf7\xddw?\xe7\x9dw.\x91\x91\x11DEGS]UMmm-y\xf9y\xe4\x1d\xcbc\xdf\xbe}\xec\xd9\xb3\x87o\xbe\xf9\x86\xd2\xd2\xd2\xc0"\xfe+1y\xf2dF\x8d\x1e\xcd\xe0\xc1\x83Hm\x9b\x8a\xd5j\xa5\xb2\xa2R\x9e\x95\x99\xcd\xd4\xd5\xd5\x91\x95\x95\xc5C\x0f=\xc4\x96-[\x02\xb3{q\xdd\xe4\xeb\xe8\xd7\xbf\x1f\xe9\xe9\xe9\xb4j\xd5\n\x8c\xb1\x95\x99\x99\xc9\xd2e\xcb\xf8\xc68\xa3\x01\xe8\xdc\xb93K\x96,\xa1C\x87\xf6A\xbbi_4\x97\x014u\x11,\xf8&p\x10\xc1\tB\xa3\xa4\xe68\xd0b\x0c\xc0\xf3\x02\xd3\xa7O\xe3\xb9g\xcf\xc4\x0e@\xae\xec\xcb\xca\xcb\xf8\xfc\xf3\xcfx\xfd\xa3i\x0c\x9c\x14Ct\x8a\rW\xad\x86\xf0\x1e\xb2\x1a\xda8J\xfd\xfa\xbe\xf9\xa8\xd7\xff\x16\x18&\xa0E\xfd\xcb\xfav\xab\xfc\xeey\x96\x11n\xd8\xadA\x05UH\xbb@f\x15r\xf7V\xb3\x7fU)\xa9\xb6\xfe<\xfb\xd7g\xe9\xd3\xa7O\xbdi`Ex\x0f\x92\x85\x10\xb8\xddn\x0e\x1d:\xc4\xa2E\x8b\xf8\xe1\xa7\xef\xc9\xaf\xdd\xce\xde\xcdyT\x97ht\xe8\xd0\x81W_{\x95\xf3\xc7_\xe0]\x05\xd5\xd6\xd5r\xd7\xdd\x7fB\xa4\xed!sk\x01)\xddb(\xdf\x92\xc0G\x1f|L\x92!V\xf2\xc0\xe1p\xf0\xc1\x07\x1f0\xf5\xefO\x93\xda;\x82\xe2c\x95t\x1d\x11Kj\x9fH\\\xb5\x82\xaa2\x17\xa5Y\x0e\x8a\xf6\x9a\xe9\x98\xd0\x9b6\t\xc9$\xc4\'\xd0\xabWO:t\xe8H\x87\x0e\x1d\t\xb3\xdb1[\r\xbb2F\xb9\xbe\xabQ\xb9\xe5\xf7\x1c\x90\x9f\x8a!~\xea\xd1\\\x06P[[\xcb\xc4\x89\x13Y\xb4h\x91_\xb8\xc9d\xc2f\xb3\xa1\xa0\xe0\xd6\xdd8\xea\x1c~\xf1\r!""\x82\xd1\xa3G\x13\xd7*\x8e\xf8\xd6\xad)--\xa3\xb2\xb2\x92\xcc\xccLv\xed\xda\x15\xa4I\xf7\xff\x13\xce>\xfbl\xbau\xebFX\x98\x9d\xc2\xc2"\xccf\x13\x16\xab\x95\x8a\x8a\n\xb6m\xdb\xc6\x81\xcc\x03\xcd\xda\xd5x\xda^NBAMM\xfd\xa53\x0f\xcc\x163\x87\xb2\xb2h\xdb\xb6\xad\xcf\x8c\x0e&\xba\xcde\x00\x8d\xed\x00f\xcc0D@\r,\x8cN7Z\x8c\x01x0}\xfat\xa9\x06\xea\xf2g\x00\xa7\xcb\x1c\xb4\xae\x0b*++\xf8\xe0\xfd\xf7\xf9\xe7\x173\x18<1\x8e\xc8d\x1b\xbaS\xf8\xc9~\x02;;\xb0\x91\x8c\xf1\xe3\x8d\xf3\xaeD}\xf9\x88\xe2C\xd4\x8d\\\xf5\xbf\xe4w\xcf\x16B\x9e?\xa8 \xe4\xedM!T\xaa\n\xeb\xd8\xb6\xf8(\x19\x11g1\xfe\xfc\x0b\xb8\xf2\xca+HNN1\xb4X\x8c:\t\x9d\xba:\'\x95\x95\x15d\x1e\xc8\xe4\xd3\xd9\x9f\xb2\xed\xd0Z\x94\xf8"\xba\x0eO gm5k\x17\x1f\xc2\xedP\xb8\xffO\x0f\xf2\xd8\x9f\xffLtT\xb4\xdc\xf5\x08A\xd6\xc1\x83\\s\xe3U\xf4\x9a(\xd8\xbf\xbe\x98\xfec:\xf0\xdd\xfb\xebxw\xdaRF\x8e\x1aex\x0f\x93\xb5S\x14\x85\x9c\xecl\xee\xb8\xfd\x0e\xac}\x0f\xa2\xda\x156/:B\xff\x0bSH\xec\x14\x8e\xdb\r&\x93\x825\xd2LUY\x1du\xe5\x82\xbc}\x15\xd4\x94*\x88j;\x11\xae6t\xce\xe8\xca\xd81\xe3\xe8\xde\xb3\x07\t\x86\x15H\x9b\xd5\x8a\xcdf3\xf8\xae/\x9b\xfc}\xa29\x0c@\x00U\x95\x95\\t\xd1\x04V\xad\xfa%0\xfa\xa4\xe1\x11\xf1\xfd/\xc1oWx\x1aa6\x9b\xc9\xca\xca"5550\xca\x0f\xcde\x00\x8d\xed\x00\xfe\xe7E@\r\xd9\x02\xba\xff\xfe\xfb\x981c\xe6)g\x00\x18\x93\xb3\xa2\xbc\x9cOg\x7f\xca\xe33\xefe\xc45\xe9\xa4\xf6\x8cEA\xda\xe0\xc1P\xcb\xf4%E\xfed\xc9s\\\x1c\x1a\x1e\x06 3\xf9h\xae\xfbr\r\x83\x05@=\xb71[M\x08\rJsk\xc8\xdap\x0c\xf3\xd1\x1e\xdcx\xe3\r\x9c;\xfa\\:t\xea\x88\xcdb\xad?_\x10:5\xb5u\xec\xd9\xb3\x87\xdf\xd6\xff\xc6\xea\xb5\xbf\xb2\xa3b.\xa9\xbd\xe2i\x9d\x1cCD\xbc\x8d\xbc\xdd\x15l^X@F\xcfD\x8en\x16|\xfc\xee\xa7\x0c\x1a4\x08a\xdc;\x10\xba\xce\x9c\xaf\xbff\xfa{\x0f1\xfc\xfa6lZr\x94n\xc3\xe3\xa9.\xaeC\xdd=\x80\x193g\xd0.\xad\x9d\xdf\xbbk\x9a\xc6\x9c\xaf\xbf\xe6/\xff\x9c\xcc\x84\xbb\x87Pp\xa8\x82\xb5s\xb3\xe9yN"i\xfdcp\xd6\xc9{\x10\xa8\xd2\xd2\xa6\xc5\xa6\xa0\xb9\xc1Y\xa3Q[\xe9\xa4\xb6\xc2IIN\x01\xd9\x9b\xdc\xb4\x8f\x1d\xc2%\x97^J\xf7n\xddh\x97\x9eAZj*QQ\xd2\n\xe5\xe9\x9e\xe4\'\x83\xe60\x00\x80\xac\xac,.\xbf\xfcr\xb6\x19&\xb8=\xe8\xd5\xab7W\\y\x05a\x1e\xc3sB\xee\xc4\xd6\xaf_\xcf\x8f?\xfe\x88+`1\xe4\x0b{X\x18u\xb5\xb5XmV\x9c\x0ey\xeb\xf7\x7f\x116\x9b\xcd{F\xd4\x14:v\xec\xc8\x95W^I\\\\\x9c\x11"\'\xa7\xf4\xbb\xbc\x9c_~\rf\xd0{\xf7\xed\x95"\xd6F\xd0|\x060\x97I\x13\xa5\x96\x9d/\xfe\xe7\xb5\x80\x84\x10L\x9f>\x9d\xe7\x9e\x7f\x0e\xb7\xcb\x9f\x01\xdcw\xdf\xbd\xcc<\r\xe6\xa0\xbd/+\x04UUU,\xfb~\x19o\xbc\xf9:uI\x99t\x1d\xd1Z\xfa\x94uj\x9e\xf5\xb9\x97`\x072\x81\xa0\xef\xde\x03^\x0f\x93\xf0\\\xf0\xf2y\xa6\x11\xa9\x08y\x91J\xd2}\x05\x93UjF\xe7\xed\xab$g[%jA\x12\x93\xae\xb8\x8eK/\xb9\x84\xb4\xd4v\x84\x85\x87\xc9\xf4\x8a\xf4#\\UY\xc9\xd6\xed[\xf9\xee\xbb\xefX\xbb\xe9W,\x1d\xb3H\xeb\x13\x87\xd9bFW5PT\xea*\xdc\xfc2+\x8bA\x17\xb6\xa3\xba\xbc\x86\x81Q7\xf2\xc4\x93O\x1aDG\xfa$.-)\xe6\xb9\xe7\x9fcK\xf57\x0c\xbe,\x99uss\x88\x8c\xb3\xd3a`kV\xbe\x9f\xcd\r\xe3\x1f\xe0\xde{\xef#\xc2\xc7V\xbb\xae\xeb\x1c\xce>\xccC\x0f<\x84m\xf8~Z\xa7\x85QU\xe8`\xd3\x82#\xb4\xe9\x1aE\xa7\xb3Z\xe3r\n\xf9v\xc6\x06\xc7\xd3\x1e\xaaYA\x15&\xa9\xfbo\x85\xda\x1a\xc1\xb1\xdd\x15d\xae*\xa3K\xe2@\xbau\xed\xce\xf9\xe3\xcf\xa7o\x9f>$\xb6i\xd3,U\xc9\x96@s\x19\xc0o\xbf\xfd\xc6\x15W\\Ann.\x00\xe9\xe9\xedx\xea\xa9\xa7\x18\x7f\xfe\x05\xc4\xc5\xc6\xfa\xbd\x9f\xae\xeb\xd2trM\r{\xf6\xeca\xee\xd7_\xf3\xddwK\xc9=\x9a\xeb\xc7\x0c?\xf8\xf0\x03\xda\xb7oO\xc7\x8e\x9d().b\xd5/\xbf0{\xf6lr\x8f\xe4RQ^NyEE\xb3\x08\xd3\x7f\x13TU%..\x8e~\xfd\xfar\xd3\xcd73h\xa0\xd4\xd3\xd7t\x8d\xac\x83\x07y\xfdui\xdc\xcd\xd3N6\x9b\x8d\xf4\x8ct\xae\xb9\xe6ZF\x8f\x1aE\xd7\xae\xdd\x88\x8c\x8c\xf0\xd2\x129\xef\x04\xba.\xef\xc6\xac\xfae\x15/\xbf\xfc2\x9b6m\x02 -\xad\x1d\xcb\x96-\xa5[\xb7n~\xf5\x08Ds\xda\xb9)\x11\xd0\xcc\x993IMMk\x91\xb3,Z\xfa"\x98\xa2(\xf2"\xd8\xcf!\\B\x9eu\x16\xe3\xc7\x8dCUO-\x11P<\x1fE\xc1f\xb5\xd1\xb5kW\x06\r\x1cL\xc5aX\xfd\xcd\x01jk\x8b\x88L\n\xc3j\x93\x96-\x11R\x8c\x83\xa2\xa2x\xce#\xbd\x8e\xc4\x15\xafh\xc4\xb7\xff\x14$\xd5W\x14#^(\xd2A\xbcPPP%!4\x19\x97\xbe\x14;\xd9{JY\xfbY\x0e\xd1E=\xb8\xf2\xbc\x9bx\xe2\xcfOp\xe1\x85\x17\x92\x90\x98\x80\xd5jAQ\x04\x9a\xe6\xa6\xa4\xa4\x84\xf5\xbf\xad\xe5\xb57^\xe5\x93\xa5oR\x11\xb7\x83\x0ecM\xc4\xa5\x85#\x14i\xe6Z\x08i\x7fg\xdb\xa2|bS\xect\x1e\x96\xc8\x82\xb7v\xf3\xdc\xe3/\x91\x91\x9enl\xa3%s\xdb\xbam+\x1f\xccz\x8f\xae\x17\xa8X"L\xe4\xed\xaf\xc2Q\xe3"\xbdo4\x960X\xbf8\x8b\x9e\xdd{\x93\x96\xda\xd68\xbc\x95\xefj\xb3\xd9\xa9\xaa\xaa\xe4\xab/\xbf\xa4\xcb\xd9\xadQ\xcd*\xed\xfa\xc7rhS\x19e\xc7ji\x95\x12\x8ejQ\x0c\xbbI\xb2e\x04\xd2H\x9a@G 5\x7fTU\xd0:\xd5N\xe7sb\x11q\xc5l\xdd\xfb\x1b\xdf.X\xc8/\xdf\xaf#<<\x9c\xc8\xa8(\xc2\xeca\xa8&y\xe9\xce_\x0b\xa4!#]rz\x9fN\x84~n0\x8a\x8a\x8a\x98;w\xae\xf7@\xb6G\xcf\x1e\xdcz\xeb\x1f\xe8\xde\xbd\x1b6\x9b\xcd\xab\x12i5D`\x11\x91\x11\xc4\xc6\xc6\xd2\xb1cG.\xb8\xe0|\x86\x0c\x19\x8a\xc5b\xa1\xb0\xb0\xd0k>\xf9\xf0\xa1C\xf4\xee\xdd\x9b\xbe}\xfb\xd2\xb6m*\x83\x06\x0c\xe4\xea\xab\xaf\xe6\xfc\x0b.`\xc4\xd9\xc3\xe9\xd8\xb1\x13\x19\x19\x19$&&\x92\x9d\x9d\x134\xaf\xfe\xdbp\xf6\xd9gs\xd5UWq\xdf}\xf7\xf3\xd8c\x8f1x\xe0 \xe2\x13\xe2\x89\x8a\x8a\xc2l2s\xe8\xf0a\x16-Z\xe4\xd5\x00:k\xd8Y\xdc~\xdbm<\xf7\xcc\xb3\\~\xd9\xe5t\xec\xd8\x81\xa8\xe8(\xecv\xbb_{{\xda<**\x92\xb6)m\xc9=r\xc4\xeb=l\xe4\xc8s\x988q"QQQ\x01\xb5\xf1Gsw\xa9\x8d\xf9\x04\x1e7n<11\xd1\xcd\x1eS\xa7\x1a-\xca\x00\x00~i\xcc)\xfc\xb8\xf3\xa5\xe3\x89\xd3\x05E\xae.\xda\xb4i\xc3\x90!C\xe8\xdd\xa3/eYf\xe6\xbd\xb7\x1c\xb7^\x85\xd5f\xc3\x16f\xc1\x1eeF5+\x08\xb7\xdc\x11x\xd6\xee\x1e\x16\xe0!\xa8\xf2\xb7bp\r\xdfS_\x81\xc5b\xc2\x1an\xc6\xed\xd2\xa9\xabtS\x9a[\xcb\xa1\xcd\x05|\xf3|&\xfd\xda\x9c\xc7\xed\xd7\xdf\xc7\xcd7\xdd\xc2\xf8\xf3\xcf\'!>\x01\xb3\xd9\xece$\xa5ee\xfc\xfa\xeb\xaf\xbc\xf3\xf6\xbb\xfck\xde\xe3\x84\xf7*\xa1\xc3\x90(bSm\xb8\x9d\x86\xafb\xe3YV\xbb\x89\x1d?\xe6SW\xa51\xe0\xe2\xb6T\x96\xd4\xd1\x851\\\x7f\xfd\rDD\xd4\xdb\xbe\xa9s8X\xb6\xf4{\xb6\x96\xcf%\xado\x0c\x8a\xaapl_\x05\x8a\x02I\x9d"\xb1EZ\xa8S\x8a\xc9\\[\xce\x90\xc1\x83\x88\xf6\xf1[j6\x9bqk\x1a+\xbe\xfb\x8d\xe8\x8e\x0e\xd96B!\xb5[4\xf9\xfb\xaa8\xb8\xa9\x8c\x94n\xd1\xd8#\xcc\xf2.\x83\xd1*\n\xbe\xb4Y\xb6\x9d\xae\x81\xe6\xd6\t\x8b\xb6\xd0\xb6{4\t\x1d-\xd4\x9a\x8b\xf8\xfb?\xfeC\xe6\xc6,lv\x1bQQQDD\x84c\xf2Y\x0c4<a\x1a\n?uh\xf8\xd9\xfe\x08\x0f\x0fg\xdf\xbe}\xde\x95\xe5\x91\x9c#\x8c\x1b;\x96.]\xbb\xf8\xbd\x8b\x07\x8ag1\xa1H\r\x97\xf4\xf4\x0c\xce:k(\xdd\xbaw\xa3\xba\xaa\x9a\xfd\xfb\xf7\x93\x97\x97\xcf\xa6M\x1b1\x9bM\xa4g\xb4\'&:\x1a\xbb\xddNrr2=z\xf6\xe4\x9c\x11\xe7p\xf6\xd9\xc3\x195z4c\xcf;\x8f.]\xba\xb0\x7f\xff~\xaa\xaa\xa4\xbf\x87\xff\x16\\\x7f\xfd\xf5<\xf4\xd0C\xdct\xd3M\\y\xe5\x95\xf4\xed\xdb\x17\xbb\xdd\x8e\xa2\xca\xf6\xc9\xcc\xcc\xe4\xcb\xaf\xbe\xe4\xfa\xeb\xae#\'\'\x07\x80\xdbn\xbb\x8d??\xf6\x18\x97_~\x05)))\x98\xcc\xf2\x1eAc\xfd\xa5\x0b\xc1\xfe}\xfb\xb9q\xca\x140\xfa\xf6O\x7f\xfa\x13g\x9f}v\x93\xe6 \x9a\xc3\x00\x84\x10\xec\xde\xbd\x9b9\r0\x80\xf1\xe3\xc7\x13\x1d\xfd\x7f\x0c \x88\x01\x0c\x192\x94\xf1\xe7\x9f\xfa\x1d@(\x08!\xb0\xdbm\xa4\xa5\xa61|\xd80.\x9b0\x11\xb5,\x91\xdf\x16\x1c"\xe7@\x11\x05\x87\xcbpT\xbb\t\x8b\xb1`2K\x97\x84\x16\x9b\t\xc5\xf0\xe7+5V\x0c}}\xd5\x90\x7f[U\xacv3V\x9b\tW\xad\x9b\xca|\'\xfb\xd6\x94\x90\xb5\xa1\x84Ck\xeb(\xdb\x12\xc9\x84\xa17\xf2\xd4\x93Oq\xeb-\x7f`\xd0\xc0A\xde[\x93\x9e+\xf3nMc\xeb\xd6-\xfc\xe3\x1f\xff\xe0\xbd\xcf\xfe\x8d\xe8\xb5\x8b\x9e\xa3R\x08\x8f\x96\xf6\x84\\n\xe9\x93\xd5\xc3\x86\xacv\x13\x877\x97rhC\t\x83/k\x8b5\xc2B\xe1\x81*\xfa\xb5\x1d\xc7\xd9g\x0f\x97\xda\x0f\x06\n\x0b\n\x98:\xf5E\xd2\xcf\xd5\xa5\x93^\x05\xf22\xabPM\n)]\xa3\xa9\xab\xd4\x89I\xb3\xb2i\xc7\x06\xcc\x95\xad\x198p\xa0Wd\xa1(R\'\x7f\xf7\xce\xddd\x15\xec \xa9S\x94\xb4\xf7/ \xb9k\x0c\xa5\xb9\xb5\xec_[\x84\xd5\xa6\x10\x16g"\xb2\x95\r\x84<\x80\x17\x18b!\xc5\xf8\x02 \x14\x84.\xd0\xdd`\tS\x89M\t\xa3\xe7\xb06d\xe6\xeca\xce\x87\xdf\x91\x95y\x98\xc8\xa8H\xd2\xd2\xd20\x9b\x9a/\x12<]{\x81\xe6NV\x8b\xc5Bll,\xdb\xb6m\xe3\xd81i\x9e\xb9MR\x12\xc3\x87\x0f\x97\xc4\xac\x89r\x14Ej\xfet\xee\xdc\x89\xb3\x86\r\xc3\xe5r\x92\x99y\x80\xbc\xbc|\x96-[FEe%\xad[\xb5\xa2m\xdb\xb6^\xb1\x94\xc9l"*J:W\xe9\xd4\xb93\x83\x07\x0f\xe6\x86\x1bn\xe0\xb2\xcb.\xa3[\xd7\xaeh\x9af\x88B\x04N\xa7G\xec\xda4!;\x1dPU\x15\x8b\xc5B\xabV\xad\xe8\xd4\xb9\x13#G\x8d\xe2\xe1\x87\x1eb\xfa\xb4i\\y\xf5U\x0c\x1a8\x88\xd4\xd4T\xecv\xbbW\xa4\xe8p\xd4\xf1\xeb\xaf\xaby\xe5o\xaf\xf0\xea\xab\xaf\xa2\xeb:\t\x89\x89<\xfc\xd0C<\xf0\xe0\x03\xf4\xe8\xd1\xd3\xf0\xab\xd0<x\x0cD\xaeY#}c\x0c\x1c8\x90[o\xbd\x95\xf6\xed\xdb7\xd9?\xcda\x00\x00{\x1b\xdd\x01\x8c#\xba\x05w\x00-z\x06@#\x87\xc0\xf7\xddw\x1f3g\x9e\x9eC\xe0\xa6\xa0\xeb\xbat.\x9d\x97\xcf\xce\x9d;Y\xb5j\x15\xd99\xd9\x94\x96\x95\xb0m\xffo\x88\xb0\x1ab\x93\xcd\xb4\xeb\x19OL\xab0\xdcfp\xbb4\xcc&+u\x8e\x1a\xea*k)\xcb\xa9\xa3\xe8`-\xee:3j]\x04\xfd{\x0e&1!\x91\xe4\xe4dz\xf6\xecE\xbf~\xfdHMK\xc3\xe6\xb1\xd9\x83\xdc]`hw\x14\x15\x15\xf3\xfd\x0f\xcbx\xfb\xc3\xb7P2\x0e\xd3iXkT\x8bb\xf8.@\x926\x0f5\x15\x02\x8b\xc9Lyq-\xbf\xce>L\xdfqI\xa4\xf6\x8a\xc1d2\xf1\xfd?2\x99\xf6\xe0\xbb\\r\xc9\xc5\xde\x89\xa1\xeb:\x0b\x17-\xe2\xa1\xd7/g\xd4\xe4\xde\x08t\xcc6\x95}\xbf\x94 4\x8d\xfe\x17\xa5PY\xe4\xc2dQqTjl\xf8\xac\x82\xa9\x0f\xff\x8dK.\xbd\x14\xb3Yz\xe2\xd2\x85\xe0\xdfo\xbd\xc5\xbf\x17>\xcf\xe0\xab\x92\xe4;\x08i\x07\xdf\x12\xae\x90\xbb\xb5\x925s\xb2\xa9\xab\x16d\xf4\x8d!\xbdO\x1cI]\xc21\xdb\xa48G\xd7\xa4\xa7&\xcf\xfeI\x18\x1a\xb1 \xdfIQ\x04\x16\x9b\x89\xd2\xc3u\xecZU@\xd5\x81H\xee\xbd\xe3>\xae\xbe\xfajRRRQ\x8d\x95`K \xf8\x0c\xa0aVSUU\xc5\xc7\x1f\x7f\xcc\xcb/\xbf\xec\x15S|\xb7\xf4;\xc6\x8e\x91\x0eH\x9a\x05CK\xac\xac\xac\x8c\x8fg}\xcc\xecOg\xb3~\xfdzTU\xa5O\x9f\xde\xdc\xff\xc0\x83\\p\xfe\xf9$$$\xc8\x9d\x85OU\x84\'\xbf\xd0qk\x1aN\x87\x83CYY\xec\xcf\xccd\xdb\xb6m\x14\x17\x15QPTHEy9;w\xee\xa2\xa8\xa8\x88\xba\xba\xba\xa0\x05\xd9\xc9\xc2n\xb7c2\x99\x88\x8fO\xa0k\xd7\xce\xc4\xc5\xb5"5-\x8d6\x89\x89\xf4\xef\xdf\x9f\x8e\x9d:\x92\x9c,\x17@\xaa\xaa\x1ak\x03\xcf\xbc\x00\xb7\xaeQTT\xc4g\xb3g\xf3\xfa\xeb\xaf\x93\x9d\x9d\x8d\x10\x82\x11#Fp\xfb\xed\xb7s\xd5\xd5W\x13\xd6\x0c\xa6\xea\x0b!\x04+V\xac`\xca\x94)\xe4\xe4\xe4\xd0\xb3gO\xfe\xfc\xe7?s\xd9e\x97\x11\x1d\x1d\x1d\x98<\x08\xcdi#\xbdY>\x81\xff\xa7\xb5\x80\xa6\xf1\xec\xb3\xa1LA\xdc\xcb\xccW^9u\x17\xc1\x9a\t\xdf\xa9\xac\x18\xc4\xd8\xe9pRTRDEy\x05999\x94\x96\x95\x91u\xf0\xa0\xf7FeII)\x95\x95\xd5X,f\x12\x13\x12\xc0\x04\xe8\x82\x846\x89\xa4\xa6\xa6\x12\x1e\x1eNz\xbbv\xc4\xc6\xc6\x10\x1b\xd7\n\x93G\xcb%\xe0\xec\x00\xc3\xab\xd3\x8e\x1d;\xf8\xe4\x93OX\xb2\xfdMz\x9d\xd7\x99V\xedlh\x0eI\x04\xa4\x01;C\xd0d\x10]\xd5\xa4 \\\n+g\x1d \xadG\x0c\x9d\xcf\x8e\xc7\xe5\xd0\t\x8b\xb2\xf2\xc3+\x05\xfc\xeb\xe5\x0f\x18=z4\xaa"5\x92j\xaakx\xe8\xc1\x878\xd2\xfa{\xdat\x8a\x014\xcca*\x99\xab\xca\xd0\xdcN\xfa_\x9cBE\x91\x0bU\x91\x96F\xb3\xb7\x95\xa1\xec\xe8\xce\xcb/\xbdL\x8f\x1e=\xbc\x93,33\x93\xabo\xb8\x9cnW\t\xa2\x12m\xe8\x9af0&\x81=\xdcJ\xce\x8eRV|\x94\xcd\xd0\xcb\xda\x92\xd2+\x9a\xca\x12\'\xaa\x0ea\xd1\xd2M\xa4\x1a&\xdb\xd7\xed\xf4\\\xc3\xf3\xb4\x86q\xd9KH\xef]B\x87\xc2C\xd5,}?\x93\x1b\xce\xff#\xb7\xdd\xf6G\x06\x0c\x1c o\x1b\x87h\xc3\xd3\x8d`\x06\xe0\x0b\x7ff \x84 ??\x9f\x85\x0b\x16\xf0\xe8c\x8fQVV\xc6\x80\x81\x03\x985k\x16=\xba\xf7\xf0\xcb\xd90\xea\xcb\xac\xae\xaef\xc3\x86\r\xbc\xf1\xc6\x1b\xcc\x9d;\xd7\x9b\xe2\xd6?\xfc\x81\xc9\x93\'3d\xf0\xe0\xc6\x89\x97O\xf5t]\xa7\xb2\xa2\x82\x8a*i\xbc-\xefX>.\x97\x8b\x92\x92b\n\n\xa4c\x16\xdd\xb0\xf7\xa3\x8bz5i!tt\xe3\x12\x9e\xe2a2\xc6E7\xc5x\x86\xe21lh\xb6\x10\x9f\x90Hbb\x02V\xab\x15\x8b\xd9LBB<QQ\xd1\xb4j\xdd\x9a\xf0\xf00\xef\x0e\x16\xa3j\x81\xec\xb4\xa6\xa6\x86\xf5\xeb\xd7\xf3\xf1\xac\x8f\xf9\xe0\xfd\x0f\xbc\xe17\xdcp\x03w\xdf}7\x03\x06\x0c\xc0j\xb5\xf8\xe9\xec7\x07999<\xf1\xe4\x93\xcc\xfe\xf4S\x84\x10\xbc\xfd\xf6\xdb\\{\xed\xb5\x01\xb2\xff\xc0\xda\xd4\xa39\x0c\xa0I\x97\x903f\xd0\xae]Z\x83\xcf8\xdd8\xad\x0c\xa0\xe1\xa6\xabG\xc3\x0c\xe0\xd4\xda\x02:\x19\x18\xb4\x1a\x8cNw\xbb\xdd8\r\xf53\x05p\xbb54\xcd\x8dbX0\x14BnW\xcdV3\x16\x8b\x15\x93\xcf\x8aFn\x1b\xeb\xed\x02\xf9\xb6\x8f\xc3\xe1`\xe9\xd2\xa5\xbc\xf9\x8f\xd7q\xa6\xed\xa3\xe7\xa8\x04tE\x95\xab~Ob\x8f\xd5Q\xa5^\xf6d\xb1\xa8\xac\x9b\x93\x83\xd9\xa6\xd0g\xbc\xb4\x8f.\x90\xb2\xf5\xb2e\xed\x99\xf1\xd2L:u\xec\xe4%\\\x1b7n\xe4\x0f\x0fM\xa6\xe7\x15V\xec\x91\x16\x84\xa2c\tS\xd8\xb1\xac\x10\xd5\xac3\xe0\xe2\x14*\x8b5\xc3\x04\x85\xc0n\xb7\xf0\xdb\xec\xa3\x8c\xea~\x03\x8f?\xf98q\xb1q\x80\x82\xd3\xe1\xe0\xc5\xa9/\xf2k\xf9\xdbt\x18\x90l\x1c\xf2z\xa6\xa1\x82\xc5\xa6rd{\x19[\x97\x1d\xe3\xack2\x88o\x17\x81\xcb\xe1\xc6b1\xa3\xeb\xd29NU\x85\x83\xf08\xb3A\\<\x9aW\xc6\xc6\xc6\x90\xa8\xa1H[J.\'l\xf9\xe6(\xf1u\xbd\xb9\xf7\xbe\xfb\x18;v\x1c\xe1a\xf5\xb6\xee\xcf\x14\x1ac\x00\xa1\xc6\xbc\x10Rs\xeb\xc5\x17\xa7\xf2\xf2\x8c\x99\x98L&n\xbe\xe9f\xfe\xfc\xf8\x9f\xe9\xd4\xa9\x931\x1c\x02s\xf9 \xa0PM\xd3\xc8\xc9\xc9\xe1\x9f\xff\xfa\'\x1f}\xf8\x11\x85\x85\x85`\xdc\x16\xbe\xf0\xc2\x0b\xb9\xfa\xea\xab\x180`\x90\xa1@\xe0\xc9\xdeH\xf9\x06<"\r\xb7\xdb-Mp\xf8\x888\x840\x98\xb2O\xfa`\xc8X\xef\xb3\x14\xc08\xcf\xb0\x98-\x8d\xb6\x9b/\x84\xe1l^\xd75v\xef\xde\xc5\xe7\x9f\x7f\xc9\x92%\x8b\xbd>\x94SRR\xb8\xfb\x9e\xbb\xb9\xf5\x96[\xe5\xae\xa7\xb9;)\x03\xba\x10\x94\x95\x96\xf0\xc6\x1bo\xf0\xcf\x7f\xfe\x8b\xa2\xa2"n\xbc\xf1\x06^zy\x06\xc9I\xc6n\xb6\x19h\x0e\x03\xd0\r-\xa0\xc6\xef\x01\xb4\x9c\x1a\xe8i=\x03hN36t\x06p\xd6YC\x197~<\xa6f\x0e\x9aS\x0b\xff\x19\xe7;\x1e\x14E:`\xf7\xd5&\xb0\xdbm^\x0b\x87\x1e\xed\x02\x9b\xcd\xea\x1d\xf4\x92X\xcb\x82\x14\xc3f\xb9$\x91\xa0\x1b\xf6\x0c\xea\xeajY\xf6\xc32\x1e}\xf1\x1eRF\xd6\xd2ix\x1cN\x87\x8e\xd0\x04(\xd2(\x9c\x9c^\xf5\x0cDA\xc5j3\xb3we!\xc5\xb9\xd5\x0c\xba\xbc\xada\xa6Z\xee\nj\x8a\x9dDVvd\xc2\x84\x8b\x88\x8e\x8aF\x00B\xd7\x98;o.\xdb\xf3\x7f$\xbd\x7f4\xba[:\xa8\x17\xa8do-\xc5\x1a\xae\x92\xd49\x1aG\xada\x88\x0e\xe9\xa3\xb8M\xd7h\xe6~>\x9f>\x1d\x87\xd1\xb1C\'L&U\xae\xf8\x14\x85\xa9\xd3\xfe\xcd\x80q\th:x\xae\xa8\t\xc3|D|F\x04\xc5G\x1c\xac\x9e}\x84\xc88\x95\xd6m#pk\x9a\\A\x9a\x04\xf6\x18\xb3\xb1\x932\xc8\x8b\x0f\x85Q|\x82\x84\x10\x98\xcc\nm{DQP\x91\xcd\x82Y\xab\x88\x8fkM\xfb\xf6\x1d\xb1\xd9\x9a/\xf3=\x15h\x8c@\x84\x8aQ\x14\x05\xab\xd5Jjj\x1a\x87\xb3\xb3\xd9\xbbg\x0f[\xb7n\xa5\xaa\xaa\x8a^\xbdz\x11\x1b\x13\x03\x8d9>\x0f\x88P\x14\x95\xd8\xd8\x18F\x8c\x18A\x87\x0e\x1d\xd0u\x9d\xacC\x87\xc8\xcb\xcbc\xdd\xba\xf5,_\xbe\x9c\xfc\xfc<\x12\x13\x13\t\x0b\x0f\xc7j\xb1J\xa2\xdc\xe0\x03$\xe4\xd8\x94\xe3\xdbb\xb1\xf8\x8dq\x8b\xc5W\x83\xc6\x82\xd5j\xc3j\xf3\xd7\xaa\xb1Zm\xc6\xc7\'\xccb\xc5l27\xdaf\xbe\x10\x86\xa7\xb3\xec\xecl\xde{\xff=\x9e~\xfai\x16|\xfb-G\x8f\x1e\xc5b\xb1r\xc5\x15\x97\xf3\xecs\xcf1\xf9\xdak\x89\x8b\x8bk6S\xf1@\xd34\xf2\x8e\x1d\xe3\xcd7\xdf\xe0\x85\x17\xa6RSS\xc3\x90!Cx\xec\xcf\x7f\xa6{\xf7\xee\xc7Eo|\x19d\x83\x10\x8d\x9b\x83\x1e?~\xbcq1\xb3y\xeds\xaaqZ\x19@s\xb0\xea\x97_X\xd1\x00\x03\x18?n\xfcqw\xf0\xa9\xc1\xa9\xe8\x8c\xa0Y\xeb\xff\xdb\x80bl\xeb\x17.X\xc8\xb3\xffx\x88~WE\x11\x9f\x11\x81\xb3\xc6\xb0\x06\x87O^\x81q\xe4+a\x0f3\x91\x7f\xa0\x8amK\xf3\x18~M:\xe1\xb1\x16\\.\xc9 l\x91&\x0e\xac/\xa1w\xdbQ\\x\xe1\x04\x14\xc3\x06~II)_}\xf9%U1\x07I\xcc\x88@s\xe9\xa8f\x95\xbar7y\xfb*\xb0\x84\x99H\xea\x1c\x85\xdb\xa1\xfbY05[\xa0uR4_\xfe{9g\r\x1dNJJ\n\x8a\xe1\xaa/gw.\xe5\x96\xc3\x84\xc7\x98\xbdu\x94,\x004\xb7 1=\x8a\xa8X+\xa5G\xeb(/\xa8\xa5Uj\x18&\xb3\x8a[\x13\xe8n_\x9bH>\x9cQ\xf1\xe3uR\x15\x17@Qi\xd3)\x02s\x8c\x93\x19\x8f|@\xfb\x8c\x0c:v\xeah\x1c\x14\x06\xb5\xfai\xc1\x89LV\x81\xa0U\\+\xbav\xed\xca\xe1\xc3\xd9dff\xb2y\xf3fJ\x8aKHNI\xa2Mb\x9bf\xefv\x15E\xd6\xc1b\xb1\xd0\xbd{w\x06\x0f\x1eL\xfb\xf6\xed\xd9\xbce3\x95\x95\x95\x94\x95\x95\xb1z\xb54\x99l6\\@\x9a\xcdf""\xa4/\x84\x13\x81\xe2\xdb/\x9e\x15\xcc\xc9 \xa0\xb3\x84\x10\x14\x16\x16\xb0e\xeb\x16\xbe\x99?\x9f\xbb\xee\xba\x9b\xb9s\xe7RTX\x84\xae\xeb\xf4\xe9\xd3\x87g\x9e\xfd+\xb7\xddv\x1b\x83\x07\x0f\xc6ji\x80\xe972\x08\\.\x17\x9b6mb\xea\xd4\xa9^W\x90\xa9\xa9\xa9\xcc\x9c9\x93Q\xa3FamB\xeb\'\x10\xcda\x00B\x88\xc6\xd5@\xc7\xb7\xec!p\x8b3\x80\x86v\x00\xa7\xeb\x1e\xc0\xef\r\x95UU\xcc\x9e=\x9b\x7f~\xfcw\x06\\o\',\xc2\x82\xe6\x96\xber=sM\xe0K\x0c\xe5A\xab-\xccL\xf6\x8eR\xd6\xcd\xc9%,\xd2L\xbf\t\xc9\xd4V\xb9\xbc\xa3\xdf\xac*\xe4l\xaf\xa4W\xea\xd9\x8c\x1a9\xcaK4rrr\x98;\xffkZ\r\xae\x96^\xb1\x84\x8a\xd9\xa2\x90\xb3\xa3\x9c\xf4\x9eQ\x1c\xdd[E\x9b\x0eQ\xd2\x08\x96gL\n\x81.tb\xda\x84Q\\THu\xae\x99\xc1\x83\x07a\xb7\x87a2\x99\xa8\xaa\xaa\xe4\x9bo\xe6\xd1iH+4\xb7\xa7\x9eFf!\x1d\xbd\x84E\x99I\x1f\x12G\xc1\xc1*2\xd7\x14\xd3\xaeW\x14f\x8b\x19aH\xb70\xc4Z\x92\xd8\xfb\x10\x1b\x00\xe4\xa1\xb0\x82<_\xd05AL\x9a\x8d\xa4na|\xfe\xd6R\xc2,\xe1\xf4\xe8\xd1C\xda\xd79\x1d\x13)\x80\xa8\x04>\xc37Z\x04\xdcO\xf0\xfd\xad(\nIm\x12\xe9\xd2\xb5+;w\xed\xe2\xc8\x91#l\xdb\xbe\x8d\xdd\xbbw\xa3(\n\x1d;v\xc0f\xb3{\xc5\x85\xbeh\x88\xae\xa9\xaaJ\xabV\xad\xe8\xd5\xbb\x17\x17\\p\x01V\xab\x95\xa3G\x8fRUU\x8d\xaek\xac\\\xb9\x92\x95+W\xb1w\xef\x1e\xf6gf\x12\x1e\x1eNDD\xa4\xd7\x16\x14\xc8A\x16\xaa\xec\xd3\x05\x81<;\x13B\xa7\xb6\xd6\xc1\xc1\x83\x07\xf9\xfa\xeb\xb9|\xf2\xc9\'\xbc\xfb\xce\xbb\xcc\x9a5\x8b\x8a\x8a\nTU%%%\x85)7M\xe1\x85\x17\xa62\xe6\xbc\xf3\x82}8\xfb\x94j\x0c\xbc \x08!\x0f\xd0\xbf\xfe\xfakf\xce\x9c\xc9\x82\x05\x0bPPh\xd7\xae\x1d\xef\xbd\xf7\x1e\xe7\x9dw.6\xab-\xa8\xef<y\x03\xc3<h.\x03h\\\x0bh|\xb3]P\x9e\x0e\xb48\x03X\xb5\xea\x17V\xae\x08f\x00g\x9d5\x84q\xe3N\xa1G\xb0\xdf\x15\xe4\xc0\xa9\xabs0\xf7\x9b\xf9L{\xf7!\x86\xdf\xd8\x06\xac\x1a\xba\xe7H\xc6g,{V\xfd&\x93\x8a\xd9lB\xd7L\xfc6\xef\x08\xdb\x97\xe6c\xb1\xa8\xf4\x9e\x90\x82=Z\xb6\x93\xa2\xc8\xe2uM#w\x9b\x93I\x17\xdcB\xb7\xae]Q\x14\xd0\x84`\xcb\xd6\xad\xbc\xf7\xe5\x9b\xf4\x1c\x17\x83\xe6\x90\xaa\xa3\xe5yN\x12\xd2\xc31\xd9T6.8F\xdb\xce\xd1\xd8c\xea\xcf2<\xcf\xd7\x05\xb4j\x17\xc6w\x9f\xaf\xa3[zO:v\xee\x88\xd9b\xc1\xe5v\xf3\xd3\xc2uDt\xac\xc5\x12f\x02\xe1o\x12[5IW\x99\x96H3\x89\x19\x91\xe8.\xd8\xbc\xe8\x18\xa9\xddb\xb0\xd8M\xe8\x9a\xe13\x00\xa5~\xb7\xe3s\xd9\xce\x08\x00\xa3\xe5\x04\x02\xe1\x12\xc4%\x87a\x8et\xb1\xf4\xb3ut\xec\xd0\x89N\x9d:\xd5\xcb\x83\x1b\xa2\x98\'\x82\x80r\x02\t\x82\xef/\x0f\xc3\x96\x97\x07=a\x86\xb5VEAUM\xa4$\'3x\xd0 \xf2\xf3\xf3\xd8\xb7w\x1f999,]\xb6\x94\xca\xca*\xd2\xd3\xd3\x89\x8b\xf3\xbf%Lp\x15\x82`\xb5Xi\xd3\xa6\r\xe7\x8c<\x87!C\x06\x13\x1b\x17\x8b\xe6\xd68z\xf4(eeel\xdf\xbe\x83_~Y\xc5\xe2\xc5\x8b\xf9\xe9\xa7\x9fpk\x9aT\t5\x99\x11BGQ\x94zq\xe5)\x86\\\xc0\x08\\.\x17\xb5u\xb5\x94\x97\x95q\xe8\xd0a~\xf9\xe5W\xa6N}\x9e7\xdex\x93/\xbf\xfa\x8a\xf5\xeb\xd7\x93\x97\x97\x07\xc0\x90\xa1C\x998q"O=\xf9\x04\xb7\xdcz+))m\x9bX\xa1\x87\xa8\xb7\x00\xa7\xcb\xc9\xbe}\xfb\x986m\x1a\x7f\xff\xfb\xdf\xd9\xbd{7\x00W^y%S\xa7Ne\xe4\xc8\x91X\xadV9\xd2B\xbc\xbb\xa2xn\xfa\x04\xa39\x0c@\xd7u\xf6\xee\xdd\xdb \x03\x18?~<1>wl\xce4Z\x96\x01\x08)\x02\xfa9\x04\x03\x182d(\xe3\xc7\xb7\x94\x08\xe8tC\xc1\xe1t\xb1l\xe9R\x9e}\xe5Q\x06NJ\xc0\x1e\xad\xca\x0b]\n\xd2\x01\x8c1\x18\x05\xd22\xa85\xcc\x84^\xa7S\x94]\xc7\xb6%yl_R\xc4YW\'\xd3\xa1\x7f+b\x93m\x98\xac>CT(\x08T\xf2\xf7\xb8\x99|\xc9-\xa4\xa6I\xa3VN\x97\x8buk\xd7\xf0\xcb\xce\x85t9;\x1eg\x9d\x8e\xd9\xa2RYXGtR\x18.\x97`\xcb\xe2\x02,\x91*I\x9d"\r\xe2\xef\xf1\x07+5\x1a\xecQ\x16,\xd1N\x0e\xae\xabb\xe49#\x89\x88\x08\'2"\x82\xbd{\xf6\xb1\xe7\xc8\x16\x92;E\xa1\xbb\x85\xf1\x1e\x92\xa0\xab\x02\xec\x91V\x84[C\xd7tZ\xa5\x86\x13\x11gf\xf5\xec\xc3\xc4&\xd9\x89M\n\xc3\xad\ty\x04l\xe4\xf1\xb4\x93\xe2\x9d|\xbe\xd3\xd0\x102\xb9\x05\t\xed#\xb1\xc6h\xfc\xf4\xc5V\xbav\xebJ\xbb\xb44\xff\xdd\xcbi@(BQ\x0f\x83]z\x9d\r\xc9\xb4\xd2\x1c\xb6|5UUILL\xa0\x7f\xff\x01DFE\xf2\xcb/\xbf\xa0k:\xeb\xd7\xafg\xe7\xce\x9d\xb4n\xdd\x9a\xd4\xb44\xacVk\x83\xc4\'\x14\x14\xe3v{\xfb\xf6\x1d\x186\xec,\x86\x0c\x19J\xcf\x9e=\xc8\xce\xc9\xa6\xa8\xb0\x08M\xd3(//\xe7\xc0\x81\x03|\xfb\xcd7\xbc\xf3\xce;l\xd9\xb2\x99\xdc\xa3\xb9\x14\x16\x16RTT\x84\xdb\xed\xa6\xb6\xb6\x06!\xa4\xf3\x16\xc5P\xcfm\xfc\x9d}`Xr\x15\xba\xa0\xb8\xb8\x98\xa2\xa2"\xb2\x0f\x1ff\xc7\xce\x1dl\xd8\xb8\x91\xef\xbf\xff\x9e\x7f\xff\xfb-\xee\xbf\xff~\xbe\xfc\xf2Kv\xed\xdaMQQ\x91\xd7\xa6\xcfyc\xc6\xf0\xf0\xc3\x0fs\xdbm\xb7q\xf5\xc4\x89t\xe9\xd2\x15\x93\xd9d\x18%<>\x94\x96\x95\xb2x\xf1b^x\xe1\x05\xe6\xcc\x99Cm\xad\xb4\xfc\xf9\xdcs\xcfr\xdf}\xf71h\xd0 \xcc&\x93\xcfx\x0b\x8d\x86b\x9b\xc3\x00\x84\x10\xec\xd9\xd30\x03\x187~\\\x8b2\x80\xd3\xaa\x05\xd4\x1c4\xec\x0f@\xda\x02\xf2\xdb\xaa\xfe\x7f\x02M\xd3Y\xb3f5\x8f>\xfe\x08\xc9c\xcbH\xec\x1c\x89\xdb\xa1#\x90r\x7f\xcf\xaaWQ\xc0dQPU3G\xf7\x96\xa1\xea*\xb1m\xc2\xd9\xfeC\x01%\xc7\xaa\x18\xf3\xa7\x8e\x1c\xdbUNB\x87hT\xab\xdcV+(\xa8&\x95\xea\x12\x17[\xbf\xa8c\xd6[_2\xa0\x7f\x7ft\x04\xd5\xd55\xfc\xe3\x8d7\xf9v\xf7\xeb\xf4\x1c\x93\x04(\xd4\x96;\x88O\x89\xa4\xb2\xc2AX\xa4\x99ySw\xd1:%\x92\x1e\xa3[\x13\x97\x16\x86\xd3\xa1yW\xe3\x9e\x81\xa2\xda\x15~yo\x1f\x7f\x7f\xe8+\xce\xbf\xe0\x02\x14Ea\xee\xd7_\xf3\xd7\x7f\xdc\xcf\xa0I\xf1\xd8\xc2\x154\xc3\xfe\xbfb8\xb4\xb1F\x98\xd0\xea44M\x12\x14\x93Y\xa10\xbb\x8e\xd5\xb3\x0f\xd1gl\x1b:\x0fo\x8d\xc30\xcb-\xcdkxZ\xcbC\x02\x8dp\xe3\xa7G\xdbHQ\x15\xccV\x13\xbbV\x1c\xc3z\xb0\x07\xaf\xcc\x9cI\xdf>}\x9b\x9c\xd4\'\x83\x93Y\x94Hb\xea\xf9.E\x13\xcb\x96-\xe5\xc9\'\xff\xc2\xa1CY\xe8\xbaN\xd7\xae]\xb9\xe6\xdak\xb8\xe3\xf6;HJNBQ\x1a9 n\x04B\x08\xeajk)(,`\xc3\x86\x8d\xcc\x993\x875k\xd6P\\\\LMM\x8d\xdf\xa2KQU\x12\x13\xe2\x192d\x08\x11\x11\x91$$$\x10\x17\x17\x87\xc9d"&&\x86\xf4\xf4t\x92S\xa4\x86Y(\x81\x91\xcb\xe5"\xfb\xf0a\x0e\x1f\xce\xc6\xe1t\xa0\xeb:yyy\x94\x95\x95QTT\xc4\x96-\x9b)++\xf7{\xa6\xaa\xaaDFF\x92\x92\x92B\xef\xde}\xb8\xe9\xa6\x1b\xe9\xdb\xb7\x1f\xadZ\xb52\x9c\xbc\xfb=\xa2Y\x90\xdaCn\xf6\xed\xdb\xcf\xc7\x1f\x7f\xcc\xbb\xef\xbdKQa\x11\xaa\xaa\xd2\xbf_?\x1e|\xf0A&\\4\x81\xd8\xd8\xb8\xe63\xb5\x06\x10\xb8h\r\x05]\x08\xe67b\x0bH\xaa\x81\xb6\x0b\x889\x1e\xb6\x7frhq\x06\xd0\xd0E\xb0\xfb\xef\xbf\x8f\x19/\xcf\xc4\xfc\xff\t\x03\xf0\xdcdD\x91\xfa\xc7\x8f<\xf4\x08E\xad~\xa3\xc7\xd8\xd6\xd4Vi(\xc2sl\n(\x02U5a\rS\xa9,r\xe2(w\xd1\xbaC\x04f\x8b\x89}\xab\x8a\xd8\xbf\xa6\x90\xf3n\xeb\x80P!wG\x05\x1d\x87\xc4QWc\\\x10S@5\x99(>RE\xe9\xf2\x14\x16/Z\x82\xdd\xb0:YPT\xc4\xd3O<EI\xe7\xd5\xc4$\x98\xb1\x84\xabT\x97\xd4\x11\x16i\xc1\xed\x14D\xc7\xdbX\xf2\xfa>\xfa^\x98L]\x85\x8b\xa4N\x91\xe8\x8a\x8a\xae\x81b\xb8\xccQ\x90\xfe\x88\x8f\xee/C\xdb\xd2\x83Of\xcd"&:\x9acG\xf3\xb8\xf2\xea+H\x19[FB\xa7(4\xcdUO\xc8\x85\xdc\x11\x08E*4y\x06\xb7\xc9l\xa2<\xb7\x9a\x15\x9f\x1c\xa6m\xb7h\xfaOH\xc6dVp\xeb\x86f\x92W\x85\xd1C8e\x0b\x19ki\xb9c\x10\x92\t\xb8\x9d:\x9b\xbe=\xc6\x88\x8cI\xfc\xf5\xe9\xa7ILLh\xd4\x99\xc7\xc9\xe0\xe4\x18\x80\xafLY\xbeI]\x9d\x83\xcd\x9b6\xf3\xfe\x07\xef\xf3\xd9g\x9fQ]]\x8d\xd5j\xe5\x92K.\xe1\xa1\x87\x1e\xa2\x7f\xff\xfe\xcd\xba9\xdc\x10\x84\x10\x08]\xe0t9\xd9\xb1c\'\x9b\xb7lf\xeb\xd6-\x1c\xc99\xc2\xc6M\x9b8b\x98R\xf0\x85\xa2(Rq@Q\xb1X\xcc\x84\x87\x85\xf9\xdd$\x0f\x84\x94\xe7\xd7\xf91\x16\xdd\xb8C\xe0\x0b\x9b\xcdF\xc7\x8e\x1d\xe9\xd9\xb3\'\xed\xdb\xb7\xa7{\xf7\xee\x0c\x1b6\x9cv\xed\xe4\xa5HT5h\xb5\x7f<\xe4\xb0\xa2\xa2\x82u\xeb\xd6\xf3\xec\xb3\xcf\xb0~\xfdz\xdcn7\xf60;7\xdep#7\xdf|\x0bC\x87\x0e\t\x12\xaf\x9d(\x9a\xc3\x00\x84\x10\xcc\x9b;\x8f\x89\x93\x1ac\x00-w\x0f\xa0\xc5D@B\x97\x13\xa1!-\xa0\xa1C\x872~\xfc\x7f\xdf!pC\x83U1\xc4\x01\xd5U\xd5\xbc\xf7\xc1\xfb|\xbf\xf7\x1d\xfa]\x9c\x86\xdb\xe1=\x05\xf5\x91|H\xca\xa7\xbb\x04a\x91\x16"Z[\x10n\xc8\xdb_\xc1\xa6E\xb9\x0c\xbf\xb6\x1d\xb6\x08\x0b\xba\xaeSq\xccEL\x92\x1dM\x93\x8e\xe3\x05\xd2\x81z\xe9\xd1jb\x1d]\xb8\xfa\xea\xab\x8c\x01/(++\xe7\xd1\x07\x1e#c\xb4\r\x8b]\xc5Y\xe3\xc6\x1emFs\xc9I\xaa\xb9\x05\xe1Q&*K52\xfa\xc7RY\xe8\xc6\x1aa2\x08\xb0\xec\x1f\xcf\xea\xcflS9\xb0\xa5\x98\xd1\xc3\xc6\xd1&)\x89\xf0\xf0\x08JJJX\xb9\xee\x1b\xd2\xfa\xb4A\xd3\xa4\xa75)?\x92y<\xf5\x93?\xa5\xf7\xb3\xf0V6\xdad\x84\xb1v\xceQ*\xf2k\t\x8f\xb5\x12\x11c\xc1\x1anBs\xd7\xb7\xa5l\x1bE^\x113\xe4\xeb\x86 \x08!\xc0b7\xd1\xba]\x18k\x16\xee#\xa3m\x07:u\xea\xe4uiy\xaaq<\x848\xd4!\xa2\xffo\x05\xb3\xc9LZZ*\xfd\xfa\xf5\xa3C\xfb\x0e\xfc\xf6\xdboTTT\xb0{\xf7n~\xfbm=aaa$%\'\x11\x1d\xd5\xc8\x05\xaf\xc6\xa0HOlf\xb3\x99\x94\x94\x14\x06\x0e\x18\xc0\xf0\xe1\xc38\xe7\x9c\x91\x9c7f\x0cc\xc7\x8ee\xe0\xc0\x01$$$\x90{\xf4(uuu\xe0e\x1c\xf2\xdeKmm-UUU\x8d|\xaaq8\xe4\xca_\x18\xde\xdd0\x1c\xb0\xf4\xea\xd5\x8b[n\xbd\x85\xcb/\xbf\x82?\xddu\x17\x93\xae\x99\xc4\xd5W]\xcd\xc5\x17]\xc4\xc0A\x83\x89\x8fo-\xed\xee\x84 \xfe40\x9f\x02\xa1\xeb:\xd9\xd99|\xf8\xd1\x87\xdc~\xe7\x1d\x1c\xca\x92\xbb\xa9\xce]:\xf3\xdcs\xcf\xf3\xa7;\xef\xa4[\x8f\xee\'\xc5\xbc\x03\x11\xc8\xdc\x1aBcj\xa0\xe3\xc6\x8d#::&h\x8c\x9c)\xb4\xf8\x0e`\xda\xb4\xe9<\xf7\\\xf0\x0e\xe0\xbe{\xefe\xe6+3O\xdb$>\xd3\xf0\xdc\xa4\\\xbat)\xf7?\xfe\'\x06^\x1fML\xb2\xd5\xff\xbd}=a\t)G\x16\x02\xccf\x13\xa5\xb9\xd5\xac\xfa\xe40\xbd\xc6\xb6\xa1\xd3\xe0V\xb8\x1c:\x8a\x19r\xb7U\x93\xd2;\xd28\xb0\x95+m]\x87}\xeb\xb3\xf8\xc3Y\xafq\xf3\x94\x9b\xbd\x06\xf5\x8e\x1e=J\xe7\xce\x9d\xb8\xe2\xef\x1d\xb0Y,\x80.o\x07\x1b\xa6\x06\x14U\xa1\xae\xc2A\xc1A\x07\x19\x03\xa4)\x89\xda\n\x17\x11\xf1f\xea\xaa\xdd`\xd8>R\x84\x82\xc9\x0e\x9b\xe6\x15p\xd3\xa8\'\xb8\xeb\xee\xbbP\x14\x85\x03\x99\x99\x8c\x1eq\x1e\xa3\x1e\x8f#\xac\x95\x8a0\xce\x02\xe4=\x03\x83\t(\x00R\xef\xdf\xa0\xe9\x98T\xa8,r\xf1\xed\x8c}\xe8n\xe8<4\x8a\x8eC\xe2h\xdb-\x167\x02w\x9d\xb1\xa24D\\\xde\xe6\x02\xc3\xdf\xb0q\xf7\xc1\xacPx\xa4\x9a\xfc\x05\xady\xfb?\xff\xa1W\xef\xde\xa7eb5\x87\x88H\xda`\xecU|\xea\xe0%\x1a\x92;\x06\x89SjjjX\xbfn\x1d3_y\x85\xef\x97-\xc3\xe5v\x13\x15\x1d\xc5u\xd7]\xc7MSnb\xc0\x80\x81\xc6\x05\xafS\xf3^\x02p9\x9d8\x9d\x0ejk\xeb\xe4Amm-\xfb\xf6\xed\xa3\xa6\xa6\x86\xca\xca\n6o\xd9B\xe6\xfeLjkk\xbd\xf5\x0f\xb5\xd0\x89\x89\x89f\xf8\xf0\xb3\xc9\xc8\xc8\xc0l\xb1\x90\x92\x9cL\xdb\xd4TL\xaaJxD\x04\x16\xb3\xd9k\x12\xe2T\xc0\xd3\xc6.\x97\x8b_W\xaf\xe6\x9d\xb7\xdff\xe9\xd2\xa5\x94\x94\x94`\xb5Z9\xfb\xec\xe1L\x9d\xfa"\xfd\xfb\xf7\'\xec4\\\x18\x0c\\\xb4\x86\x82\xf8\x9d\x9b\x83\xfe\x1d0\x80i<\xf7\xdcs\xc1\x0c\xe04\xf9\x03h)\xe8@\xfe\xb1<\x9ey\xe6iv\xbb\x96\xd0kL\x12n\x87\x86P\x8c\x03S\x9fI%\xf0\x88J\xc0dV\xa9.w\xb2\xe8\xef\xfb\x89O\xb31\xfe\xdeN\xd4U\xe9\x80\xc0\x1c\xa6\x92\xb3\xa5\x9c\x94\x1e1\x08M\x97\xb9\r\x9b\xd5\xeb\xbf\xde\xce\xd4?\xcc\xe5\xd2\xcb.\xf5\x12\x8b\xdc\xa3G\xe9\xd1\xab3W\xfc\xad\x03&a\xdcN\xae\x97L\x81\xa2PY\xe4\xe0\xe8\x9eJ\xba\x9c\x1d\x8f"\xc0\x1an\xa2\xaa\xa4\x8e\x88\xd6V\x9c5\x06A\x03\xc2\xa2\xcdl]\x92\xc7\xf0\x84\x1bx\xe6\x99g\xb0\x87\xd9\xa9\xad\xa9\xe5\xed\xb7\xdf\xe6\xd5\xcf\x9e\xe4\xe2\x87{\xe0\xa8\xae\xefS\xe1\xbdI\xeaY\xbdK\x91\x97bLd\xb3\xd9DIv\r?\xbc\x9dEE\xae\x9b\xa8\x143\xdd\x86\xc6\x13\x9eh\xa3}\xff\x18\xcca\xd2`\x9c\xa2*h.\xcf\x8dc\x0f\x89\xf50\x02\xb9\xd2\xdd\xf5s.CZ\xdd\xc0\xb3\xcf>\xeb\xbd\x00w*\xe7Wc\x0c\xc0\xefYF\x9dB\xc1\xc3\x1f\x02\xa3%\x1f\x97>\x17\xfe\xf5\xcf\x7f\xb1`\xe1B\xf6\xee\xd9\x83\xaa\xaa\xb4o\xdf\x9eG\x1fy\x94\xcb\xaf\xb8\x9c\xd6\xad[\x9f\xb6\xb9\xe1\xbb\x82\x17B\xa0i\x9a\xf7\xb7b,\x18\x8c\x9e\xf4\xb6\xbf\x8c\x04UQ%\x817\xcc\xa8\xa8\xc6\xe5\xb2\xfa\x1b\xe2\x8d\xa1\xf9=%\x00\xcd\xadq\xec\xd8Q\xe6\xcf\x9b\xc7\xab\xaf\xbd\xc6\xa1C\x87\x10B\xd0\xa1C\x07n\xb9\xe5\x16n\xb8\xe1\x06\xda\xb5k\xe75e~\xaaq\xaa\x18@K\xde\x04nx$\xb70\x9a\xbb\xbd\xfao\x81\xb3\xae\x8e\xdf6\xacg\xf5\xe6\x15t\x1b\x99\x80\xa6\xe9~\xc4_N\xa8z\x02\x8b!fqT\xbbX\xf1\xc1!j\x8b\x05}\xc6\xa4\xe1\xae\x13\xdei\xa7**\x91qVLA\x03\\a\xfff\xe8\xd0\xb1\x83\xdfJQ\x01\x14\x13\xd8m\x86\'*oI\x8a\xb4+\x04D\xc5\xdb\xc9\xdeZ\x81\xd9"\x8d\x8a9j5lQV\x9c5:\x8a\x97p\x0b\\\x0e\x8d\x88Vv\x8e\x1e\xcd\xa5\xa8\xa8\x08\x04\x84\x85\x87q\xeey\xe7\xd2)\xea,\x0eo,\xc1l\x93D\xc0K\xac\xbd\xd5\xf4Lt\xf9LE\x01\xdd-\x88k\x17N\xffK\x92H\xe9\x19\x85\xd9j"\xbdG+\xd2{GQx\xa8\x82\xb2\x9cZJr\xaa\xa8\xabr\xa1 \xc5]F\xcd\xe5\xf9\x89b\x10.\x04\x9d\xcfJb\xd1\xbaY,Y\xf2\x1dn\xddc<\xef\xcc\xc0\x8f\xcc\x18D0\xd4X\xf6\xdeu\x08\x0cW@U\x14\xdag\xb4\xe7\xa9\xa7\x9e\xe2\xa5\x97^b\xd4\xa8\xd1\xe8\xba\xce\x81\x03\x07\xb8\xf3Ow\xf2\xd8c\x8f\xf1\xd3\xcf?Q^^\x1e\x98\xfd\x94@Q\x14TE\xca\xff=\xb7\xde\xe5\xedvy\xcb\xd7f\x93\xbf\xad\xc6\x7f\xef\xc7j\xc3l\xdc~W\x15\x05\x93Z\x7f\x98\xdf4\xf1\'t\x834\x80\xca\x8a\nV\xac\xf8\x99\'\x9f|\x92\xfb\xee\xbf\x9f\xac\xac,\x84\x10L\xb8\xf0B\xfe\xf6\xb7\xbfq\xf7=\xf7\x90\x91\x91!\x99up\xf3\xffn \xc7l\xcb\xe1w\xcb\x00\xfe\x7fCII\t\x1f\xbf\xff!\x1d\xcf5c\xb2\xa8\xe8\xc6\x19H=<++\t\xd5\xa2\xe2\xaas\xb3qA.\xdd\x86\'0\xe1\x81\x8eD%Yq\xbb%1U\x14\xb0\xdbU\xc2\xa2lre\xec-E\xa1\xae\xd8E\xc7\xa8\xae\xc4\xb7j]_\xa0\x01]Sq9\xdc`\xc8\xce\xbd\xdb\x00\x03\xaa\n\xc9\x1d#q;\xdc\xa0\xcai\xab\xbb%\xb5\x12\xc2C\xc5\xc1\xed\x12\xc4\xb6\xb5\xe3t\xbb\xd0t\xe3\x1cC@\xf7\xee=\xb8f\xd2\xb5\x1c\\[\x8d\xa3B3.\x9b\x19\x86\xeb$\x1b\x90\xff\xfdF\xbd\xe7\x82\x97F\xf7\xe1\xf1t\x1d\xde\x9aq\xb7uf\xcb\xca#\x14\x1d\xae\xa4\xe3\xa08b\xd3\xc2Hl\x1fETk\x1b\xb6Hs\xfd\xca\xd9\xe7\xce\x81\x00\x84\x06a\xd1\x16\xba\x9e\xd7\x9a9s\xbe\xe4\xd0\xc1\x83\xc7CW\x8e\x0f\xcd\x98\xb9JC\xf6\xe8\xe5F(4\x8c\xf4\xb1\xb1\xb1L\x980\x81\xb7\xdez\x8b\'\x9ex\xc2k\xf7\xe8\xe3\x8f?\xe6\x81\x07\x1e\xe4\x85\xe7_`\xe3\xa6\x8d8\x9dN\x84\x08>p=)\xf8j\xe36\x08\xdf\x04\xf2\xd9\xfey<}~j \x84\xc0\xe9t\xb2y\xcb\x16^~\xf9e\xee\x7f\xe0\x01>\xf9\xe4\x13\x00\x12\x12\x12x\xf0\xc1\x07x\xe3\xcd7\x99p\xd1E\xc4\xf9^\xae\n\xf1"\xa7\xb4\xadN\x02\x9eEUK\xe1\xcc2\x80\x10\xef\x19\xa2o\xc08<\xfa\xaf\x87a\xd4J\xd346l\xdc\xc0\xcf\x07\xbf!\xa9G\x84a\xdcM\xaeV}\xa1(r\xb5\xa2\x9a$\x91\xde\xf8\xcd1\xc2"\xadt\x1a\x1eG]M\x1da1\xb2\xb1\x04 \x84\x82\xcb)0\xd9@\xd3dIB\x80bR\xa8,\xac\xe5\xea\x89\x13\x89\x88\x8a\xf4+_\x08\x81Z\'\x9d\xd2\x18J\x94\xc6\x83=\xf1\xf2\xd3\xa6k8\xaej\xe30\xaf>7 \xd0\x15\xe4\xb0\x11\n\xe11\xd2\xc6\x8b\xe7\x80[A\xc1j\xb1p\xc9%\x97\xd2\xb7\xedy\xe4\xef\xad6\x0e\xf1\xa5\xfa\xa6\x02>\xaa=\xf5"\x04!@(\xf2Y\x0e\x97F\x87\xc11\x08Ug\xd4\xad\xed)8\\\xc3o\xf3\x8f\xe1v\xbaq\xbb\x05\x8eJ7\x8e*iK\xc8K\xfc\x8dJzZ\xc7Y\xab\x91\xd63\x92]\xa5+X\xb5\xea\x17\x9c\x0e\x87\xf7-N)\x1a\x18\xbb\xa7\x12V\x8b\x85n\xdd\xba\xf0\xf8\xe3\x8f\xf3\xde\xfb\xefs\xf1\xc5\x17\xa3(\n\xbbv\xee\xe4\xb5\xd7_\xe3\xdak\xaee\xda\xb4i\xec\xd9\xb3\x87\xaa\xaa\xaazf|\xc6\xd1\xbc\xc6\x08A\x02\xa0\x91p\x8cq[[[\xcb\xa1C\x87\xf8\xd7\xbf\xfe\xc9\xb5\xd7\\\xc3\x8c\x193\xd8\xb9c\x07\x00\x13&L\xe0\xd5W_\xe5\xaf\xcf<K\x87\x0e\x1d\xa4A\xc6\xc0B\xa8\x7f\x88\xa0\x11\xc2s\x8a\xd1(\xa3\x11\x86\xfc\xb5\x05qf\x19\x80_\x9b{\xde\xfc\xcctD\x8b@\xb2w\xca\xcaJ\xf9\xf2\xcb/\x19:\xa1=\xa8>r\xc3\x80W\x17\x02LV\x05\xcd\xa9\xb1u\xd11\xea\xaa5\xfa]\x94\x82\xd0\x04\xb60iASf\x93\xb2\xfe\xdaj\x97T\xb2\xf7\x19D\xf603\xbb\xd7\x14\xd3.\xbd]\x90G#\x8b\xd5\xca%W\\Lm\xa5\x03E\xad_wHF \t<\x08\xa2\x12l\x14\x1c\xaa\xc6b\x97\xb7z\xbdi<7Z\x8dt\xba1\xb8\xbd+\\\xe3_RR\x1b\xee\xbd\xf7>\x0e\xfdd\xa1\xbc\xb8\x16\xb3\xc5\x9fJ+\x06\xb1\xf7\xb2\x17_\xdb\x0fB\xa0X\x14\x84\xaeaQT\x06]\x9aJT\x82\x95\xcd\xf3\x8fQYP\x875\xc2\xb8\xf1\x8cG\x19\xc8\xc7\xb8\x9e"\xeb(\x10\xb8]\x1a}&$\xf1\xd5\x97s()\x91.\x19\x7fo\xf0\x88\xac\x9a\x82\x82Jtt4W_}53f\xced\xe6+3i\xd5*\x0eM\xd3\xc8\xcc\xcc\xe4\xb9\xe7\x9ec\xd4\xe8sy\xe3\x8d7Y\xb5j\x15\xf9\x05\x05\x8d\x13\x9e\xe6\xe0$\xb37\x84\x86f{C\xe1UUUl\xd9\xb2\x85\x0f?\xfc\x80\xc9\x93\'\xf3\xe0\x83\x0f\xb1o\xdf>\xdcn7\xed\xdbw\xe0\xa5\x97^b\xc6\x8c\x19L\xbev2\xb11R\x9b\xc6\x18I\xc1\xf0\x0cS\xe3<\xca\x17\'\xdd^\r \xe4\xee\xcf\x03c\xac\x9e\x9e\'7\x0fg\x96\x01\xf8@\xe0\xe1\x80\xberp\x9fx\x83\xd8\xfd\xf7C\xf0\xeb\xea\xd5\xec8\xfa\x0bI\xed\xa3\xbd&\x93\xa1\x9e\xfb\x0b\xe4\xe0\xb4\xd9M\xd4\x14;\xa9.\xd1):\\\xc3\xc0K\x92$\xa1\xd6\x15"\x13l\x86{E\xcf\xaaW\x01\xdd\xb0\xbdi\xd0WEQ\xa8)wb\xadL\xa4C\xc7\xf6Az\xdb\xe1\xe1\xe1\xf4\xe8\xd5\x83\xd2\x1c\'\x16\x8bj\xf4\x81\xec\x02\x832\xa3k\x10\x9d`\xc3Q\xe3"2\xda&\x83}\x9fa\x10-O^\x19\xe8\xf7\x18P\x14\xfa\xf5\xed\xcb}\x7fz\x90\xcd\xf3\n\xd0\x91;\x13\x8cs\x0eI\xee\x8d\x02\x8dg\x0b\xc5s\x06\xa2\xa0\xeb:\xad\xd3\xc3)8T\x8dIU\xc8\x18\x14G\xc7\xe1\xad\xd80?\x97\xccu\xc5R5Uj\xb6\xca\xba\x08\xa95%|\xca\xd3\\\x82\xf08\x13\xc7,+\xd8\xb4y\xbdW\x8d\xf5Dp\xdc\xc4!Dzat\xb9/\x14\xc5\x10\x8f5\xb5\x124\xda\xd7l6\xd3\xbd[7\xee\xb8\xe3\x0e\x16/Y\xc2\x94)SHLLDQ\x14\n\x0b\n\xf8\xcb_\x9e\xe2\xde{\xee\xe5\xa9\xa7\x9e\xe2\xbdw\xdfc\xff\xfe\xfd\xd4\xd5\xd5\xc9\xf3&\x9f\n\x84\xa8^0\x02\xfb\xb4\x11\xf8\x16\xd7`\xd1\rFx \x07\xa2\xaeK\xc75\xe5\xe5\xe5,\xf9\xee;^\x9c\xf6"\x0f>\xf8 w\xddu7\xeb\xd6\xadCQ\x14bbb\xb8\xf8\xe2\x8b\xf9\xe4\x93O\xb8\xfb\xee\xbb\xe9\xd1\xa3G\x93\xaec=\xef/\x87G\xf0\xcb\x1dw\x1f7\x17M\x95k\xec\x9c[\n\x8d\xb7\xdai\x84\x82\x87\x03608Z\xaeM\x8e\x1f\xa1\xea\x8f4vUUU\xc5\xca\x95+\x89L\xd3\xb0E\xab\x92\x10\x19\xe9\x05R\rDQ\x14\xeca\x16jK\x9dD\xc4Y\xa9*u\xd0eDkb\xda\xd8\xd1\xdc\x1aB\xd1\x89jm1\x06\xae\x97b\x1b+gc5\xa3\x80\xd5\xa6\x92\xbb\xab\x92~=\x06\xd1\xb1c\xe7 m\x15\x9b\xd5J\x97.]8\xb0\xa6\x04k\x98\xc9\xd0\xaa\xaf\'\x0c m\xf2\xe8\x02\x12:DRYT\x07Bj0y\x9e)\x8cA\xa3*\xe0\xaa\xf5\xdc\r\xf0\x87\x02\xd8\xeda\x8c\x1f7\x9e\xd1\xdd\xaeb\xcf\x0f\x05X,&\xc0T\xaf\xfe\xe9I\xe9\x95\x81z\xc6\x82G\xf7\\\',\xc1\x8cKw\xe3\xac\xd1i\x9d\x16\xcey\x7f\xec@qn\x1d\xbf|t\x18E(\x98-\xaaTM5\xea%/\x8f\t\xa9>\x0b\x98\x15\x13\xed\xfb%\xf3\x8f7\xffIEE\xb9\x0f\x838>\x84\xa0\x17\x8d#D\x06E\t\xd1PH& w.\x811\r#2"\x92!\x83\x06\xf3\xc6\x9bo\xf0\xce\xbbos\xd3\xcd7\xd1\xbe}{\x00v\xec\xd8\xc1\xfb\xef\xbd\xc7\x83\x0f>\xc8\xa5\x97^\xca\xe3\x8f?\xce\xb2eK9r$\x87\xca\xcaJ\x9cN\xe7)\x979\xfbV\xbd\xc1\xd7\x08\x19!\xc7\xb2\x10\x82\x9a\x9aZ\x8a\x8bK\xd8\xb9s\'\xb3f\xcd\xe2\xca+\xaf\xe4\xe6\x9bnb\xe6\x8c\x99\xacX\xb1\x02\x80\xb8\xb88\xae\xbf\xe1z\xfe\xfd\xef\x7f3k\xd6,\xce:\xeb,""#\x83V\xd9\xbeo\xe7\xfd\xae\xc8\xd1\x1e\xb2\x1aMhv\x9d\x0c\x9aj\xe9\x86\xeas\xa6pz\xde\xba\x19h\xaaa\xfe\xab\x10\xb2\x17\xe5\x8a:\';\x87]\xbbv\x92\xda+\x1a\xb7K\x93\x04\xd0K\r\xa4\xc8C\x05\x1c\xb5.\xc2Z\x99p\xb9\x04\xe5Gk\xe8<\xac5\x0e\x87\xa1\xff\xae\x830I\x9b:\xf5y\xa9\'\xfe\xc6\xd0\xd65AaV\r\x9d;w!\xa9MbP\x1b\x9b\xcdf\xda\xa5\xb7#9\xbc+\xce*\x8f\xd9N\x8f\x08\xc5C\x84\x05\xaeZ7\x89\x19\x91T\x14\xd4a\xb6\x19f\x08\xbc\xe2\x1f\xf9]U\x15\xaaK\\\xc6\xca;\xb8CUU\xa1C\x87\xf6\xdcy\xe7\x9d\xd4\xecm\xc5\xd1\xbde(\x16\xcf*\xdc\x97\x04y\xbe\x1b\x7f\x8d\xd7\xd3\\:V\xbb\t\xb7C2%\xb7S \x14\x85A\x17\xb7%\xado,\xabg\x1f\xa2\xb6\xc4\x89\xc9\xec\x99\xd42\xdeS&\xc6\xbd\x81\xc8\xc40\x0eW\xecf\xd3\xa6-\xc7Mh\xebqB\x99N\t\xfc\x9a\xd5\xe7\x87\xa2\xaaDG\xc5p\xf1E\x970}\xda4\xfe\xf3\x9f\xff\xf0\xc4\x13O\xc8dBPU]\xc5\x9e={x\xfd\xf5\xd7\x990a\x02\x17_t13g\xbe\xc2\x9c\xaf\xbff\xed\xbau\x1c:t\x88\xb2\xb2\xb2FV\xbe"\xf0\xe9>\xe1\xc7\x81\x06\x8a\xd1\x85NQq1\x99\x07\x0e\xf0\xf3O?\xf1\xc1\x87\x1f\xf2\x97\xa7\xffB\x9f>}\xb8\xf5\x96[\xf8\xf1\xc7\x1f)((@\xd3\xe4\x99\xc6\xf5\xd7_\xcf{\xef\xbd\xc7K\xd3_\xe2\x9ak\xae!66V\xba\x05\r,8\xa0\xb7|GD\xa8\xb4\xa7\x1b\x81\xcc)\x10\xa1v\x86g\x12-v\x13\xd8\xd3u\xabV\xadb\xc5\x8a\x15A:\xb5g\r\x1b\xca\xb8\x16\xf3\x07pj t\xc1\xa6M\x9b\x98\xfd\xdd;\xf4\xbd\xb8\x15\xce\x1a\xe3v\xac\x0fT\x83\xe8\xaa&\x15MS\xc8\xdd]AF\xdf\x180\xa9\x08]\x92G\x05\x05U\x95\x13[7T^\x14I\x87\r\xbdz\xc3\xfeO\x99\x93\xc3\xab5\xee\xbe\xf5\x01:u\xea\xe4%\xea\x1eH\x91\x83\xc2\xe6\r[)\x16\x07\x89I\xb2\xa3\x1b\x07\xc8\xfe\xd5RPL*\x95\x85.b\x93\xad\xb8\xea<\x07\x8br\xbb*\x10Xl*\xd9[\xcaH\x8f\xec\xc7\xf9\xe7\x9fOXXX\xf0`\x17\x10\x1f\xdf\x9a\xce\x1d\xbb\xf3\xf1[\xdf\xd0\xa6\x8b\xd9k-T\xc6{\x08r\xfdD\xf6\xbc\x9db\xd8\x82\x10\x9e\x9b\xb4\x1eQ\x89\x021\x89v\x12;\xc7\xe0p8QU\x93\x14/y\n\xf0\xbe\x8c|W[\xa4JQn\x05qj:\xc3\x87\r\x0f\xae\xe3\t\xa0\xb9e\x04\xb7\xeb\x89\xc1[F@a\x8a"\xeb\x12\x19\x19EF\xfb\x0c\x86\x0c\x19\xca\xc4I\x93\xc8HO\xa7\xbc\xbc\x9c\xda\x9a\x1aCCH\x90_P\xc0\xca\x95+\x99?o\x1e\xbf\xfd\xb6\x9e\r\x1b6\xb2\xe1\xb7\rl\xd9\xb2\x85\xea\xea\x1aJ\xcb\xcap\xbb]\xd8lv\xc3\xb5\xa7|\x80\xaekr\xd7\xea\xd9\xad\x1b\xe1\n\x92zI\xb7\xa4\xd4\x9f\xe3\x18+zy\xbeax\xd0\xd3\xdc\xb857\x85\x05\x85\xec\xdd\xbb\x97\xcd[6\xf3\xc5\xe7_0\xff\x9b\xf9\xcc\x9b;\x8f\x7f\xbf\xf5\x16\x9f}\xf6\x19\x1b6l\xf0\xbe\x9b\xd5j%\xa5m\nc\xc7\x8c\xe5\xd5W_\xe5\xd6[o\xa5o\xdf\xbe\xc4\xc6\xc66\xbb\xfdO7\x82\x98g\x03\x1d\xde\xd8M\xe0\xf1\xe3\xc6\x11\x13\xfb\xbfl\x0cn\xfa4\x9e{6\xc4E\xb0\x07\xeeg\xe6\xcb3N\xdbe\x973\x81\xca\xaa*^{\xf55V\x14\xbeK\xa7!q\xf2\xf6\xae\xb1\n5\xa4\x91(B\x91tPQ0+fr\xf7\x94\x91\xd4-\xd2`\x88\n P\x15U\x9e\x1d\x18DQ1V\xbb\x8a1\x08\x15EA\xd1U\xb2\xb6\x97\x10\x95\xdd\x9fY\x1f}L\xa4\x9f_\xd3z\xe8\xba\xce\xeb\xaf\xbf\xce\xe7\xeb\xfe\xc6\x80\xcb\xda\xe0\xac\x95\x1aIrNI\xe2\xae\xa0\xc8\xbd\xa1\xaeb\xb6\x83\xb3V3\x0e\x81\x91\xa3\\\xa8\xd8"Ll\xfa\xf6(\x17t\xbe\x9d\xbf\xfe\xf5iC\xdb\'\xf4\x0c\xa8s\xd4\xb1x\xd1b^z\xf3yz]\'\xa4\th\x97@P\xef\xf7\x00\xf9j\xfee\x08|\xca\xf31\x19m\xfc\xd6u\xb9\xd3\x00\x8fJ\xab7\xca\xfb#,\xda\xcc\x8e\x1f\n\xe8\xa6L\xe0\xc5\xa9/\xd2\xbau\xb0j\xec\xf1\xe2d\x17%^\xa6\xd6\x08<\xfd\x10\xbaE\x1b\x86.t4\xb7\x9b\xbc\xbc\x02\xd6\xad[\xcb\xbauk\xd9\xb3g/+V\xac\xa0\xb2\xb2\xd2/\xadbx\xff\x8a\x8d\x8d\xc5f\xb3\x11\x1f\x1fOll,\x89m\xda0j\xe4H\xaf\xed\xfd\xb0\xf00R\xdb\xa6\x11\x17\x17\x83jh\xe7\xc9\x95\xab\xa1\xc9e\xf0\x00\x15\x85\xaa\xaa*\xf2\xf2\xf3)//\xc3\xe5tSSS\xc5\xba\xb5\xeb\xd8\xb5g7\x15\x15\x95\xe4\xe7\xe5QWWGYi)Z\x08[A\xbd\xfb\xf4\xa1G\xf7\x1e\x0c\x18\xd0\x9f\x11#F\xd0\xa7o\x1f\xc2\xc3\xc2~\x97&a\x02\x17\xad\xa1\xa07\xe2\x14\xfe\xaa\xab\xafb\xe6\x8c\x99\xa4\xa7\xb7;\xce^>uhq\x06 }\x02\x073\x80\xfb\xef\xbf\x9f\x97g\xcc8\xe3N\xe1O%\xf2\xf3\x0b\xb8\xee\xba\xebH\xbf\xb6\x04\xd5\xac\xcb\xd5\x11\xf5}\xed%|(D\xc4X\xc8\xde\\N|\x87\x08\x14E\x91\x96A\x8d\x84\x8a!\x06\xf2\xe8RJ\x92\'g\x9db\x98\x8bPQY\xf4\x9fm\xfc\xe3\xf1\xaf\xb8\xfc\xb2\xcbC\x12)IL\x14\xb6m\xdb\xca\x8dwO\xa2\xcf$\x0ba\xd1f4\xe3\x1e\x81\x94\x91\xfbPd#\xbd\xb1\xf0\xf6\x1e\x04c\x10\x80\x8ds\xf3\xb9}\xc2\xd3\xdcu\xd7\x9f\x02\x88s0\xaa\xab\xaay\xeb\xdfo\xf1\x9fy\xcf2\xe2\x86N\xd8\xa3U\x9c\x0e\xddX9\xcar\x15\xb9\x80\x94\xe5\xf8\xd4\xc1\x13gp\x08\xcf?\xe3i\xf2\xfd\xeb\xf3\xfaW\xc3dV\xa8(r\xe1\xf8\xa53\xaf\xccx\x85\x0e\x1d:6I|\x9bB\xa8\xb6=\x1ex\x88zs\x18\xc1q\xc3\x8f\x13BUu\x15G\x8f\x1de\xff\xbe\xfd\xec\xdd\xbb\x97-[\xb70\xeb\xe3Y~Y\x9a\x0b\x93\xc9\xe4}w\xc5\xe49\xcf2,\xb4\x1aK\x1a\x8fY\xe7\xe3AtT\x14\xb7\xdcz+=z\xf4\xa0w\x9f>\xb4m\x9bB\xdb\xb6\xd2\x84\x84\x07\xfeo\xf5\xfb@s\x18\x80\x10\x82y\xf3\xe62q\xe2\xa4\xc0(\xae\xba\xea*f\xcc\x98Azz\xfa\xa9\x1f\x07\xcd\xc4\xc9\x8d\xe4\xd3\x08\x11BU\xeb\xf8\xd0\xa2|\r!\x04\xb9\xb9GXw\xe8\'l1\x18T\xcc\xd8\xb6#\xa9\x98\x00\xb9\x92W\xc1Q\xe9@U\xc0bU\x0c/Y\x8a\x97\xb8IB\xe1\xf1\xb7+\xe7\xb8W\x95M\x80\xc9\xa2R~\xac\x8e\xa1\xc9\x97r\xd6\x90\xb3\x1a$P\x92\xe8\xe8t\xe8\xd8\x81K\xce\xbd\x8ac\xfbJ<\x81\xb2 \xe1c\xbd\xd3\xa3Y\xe3!\xfe\x18\xaf\xa0H\x07\xed\xba\xa6\xa3\xb8m\xf4\xea\xd5\xd3x^\xe3\xbd\x15\x1e\x11\xce\xcd\xb7\xdc\xc4\r\x17<\xc2\x9a\xcf\x8eRW\xee\xc0l\xa5>\x9f\x8f8K\xd6\xc5\xc8\xe8\xdd1\xc9\x9d\x12\xc6N\xa0\xfei>?\x14\xdf\xe2\x14\x14\xa4%S{\x94\xca\x9e\xfd\xbb8\x92\x9b\xeb\xcd\x15\n\xa7{\xc4\x04\x96\x7f<\x93>p\xa5\xdc <\x8dc|"##\xe9\xd2\xb9\x0b\x17^x!w\xdey\'\xaf\xfe\xfd\xefdf\xeeg\xe1\xc2\x85\xfc\xfd\xef\x7fg\xdc\xb8q\xf4\xef\xdf\x9f\x94\x94\x14\x12\x12\x12\xb0\x19\xb6z\xd4\x10\xcea4M\xc3\xe5r\xe1r\xb9p\xd69p9\\\xd2\x8e\x90\xcb\x89\xd3\x08\xf7\x85\xa2(\x98Ty\x9b8\xccn\'11\x91\xb4\xb44\x86\r\x1b\xc6\xa5\x97^\xca;\xef\xbc\xc3\xf7\xdf\x7f\xcf\xf6\x9d;x\xe6\xd9g\x982\xe5F\x86\x9d5\x94vi\xed\x82|\xf36\xbf\xa5\x8e\x0fr\xe47\xb3mO\x18\xa1k/\xa5\x01r\x9e\xb5\x14~\x07;\x80\xe9\xc6\x0e\xc0\x7f\xf0\xdc\x7f\xff}\xcc\x981\xf3\xb4\x8a\x80\x9a^U4\x9d\xa2!H\xce?\x8f\xbf/\xbb\x8d\x1eg\xb73n\xd3\xca\x8d\xbd\x97\xc0\x19\xd7\xfem\x91*GwW\x90\xdc9\nG\x8d4\xbb,\xe9Z\xfd\xb2\xd6\x13\xe6;V%CP0\xdbT~|3\x8fg\xee\x9f\xc9\xd5W^\x85\xb5\x11\xb3\xbd\x18u\xdb\xb6m;\x13\xaf\xbb\x82\x11\x0fE\xa3\xa8\x869\x05\xea\x9f\xe1y\x94\x87\x06\xc8\xeaJuM\xb3\xd9D\xf1\x91Z\x0e\xce\xb32\xe7\xb3yt\xec\xd8\xb1\xd9\x04\xaa\xbc\xa2\x9c9s\xe6\xf0\xd6{o\x92:\xdeAR\xd7\x08y\xeb\xd8\xb8\xd8\xa5(\xf2\xab|l\xfd\xd4\xf4\xd6\xc7S\x90\x11\xe1UO5\x12\xc9`E\xda\x06\x02\x04\n&\x0b,~\xe9\x10\xef\xbf\xfa%c\xc7\x8em\x90A6\x17\'\x9b\xffx\xe1\xdb\xb6\x81D\xf9x!\x8c\xf2\x84.\xb5\xd44M\xa3\xb0\xb0\x80c\xc7\xf2(++g\xe7\xce\x1d\x94\x97\x97\xe3v\xbb),*\xa2\xbc\xac\x0c]\xd7q:\x9dTVVPYYEyy\xb9\xa1^\xaaa\xb1X\xb0\xdb\xed\x84G\x84\x13\x1d\x1dM\xab\xb88L\x86\x13\xf8\xe8\xe8h\x92\x92\x920\x99L\xc4\xc5\xc5\xd2\xa7O_\xa2\xa2\xa3\xe8\xd0\xbe=\xe1\x11\x91\x98M*(\x92\xd1\x84\xb2\x02z\xaa!\x8c\xadb\xd0\x93\xfc\x06Vc\xf0O\xd8\x9c\x1d@c"\xa0\x89\x13\'2c\xe6L\xda\xb5\xa0-\xa0\xdf\t\x03\x08\xb6\x06z\xff}\xf71c\xe6\xe9e\x00\xa7\x13\x9a\xa6\xf1\xde\xbb\xef\xf2\xf9\xee\xa7\xe980\t\xa1\x19\x07\x99\x9e\x04\x06uUP0\x99\xa4\xecV14y\x84\xe7\x10\xd4\x90\xc9{\xc6\x9d\x97!x\xff)X\xec&\x0em.%:{\x10\xd3\xa7\xbfD\x87\x0e\x1d\xfc+\x12\x02\x02\xa8\xae\xaa\xe2\x8d7\xde`\xc9\x9e\xb7\xe8=>\x11\x87\xd3-\x17#\x01i\x8dc\x07\x83\xa8\x1a\x01\x8a\xca\x91\x1d\xa5t\xa8\x99\xc0\xeb\xaf\xbf\x86\xc5r<\x9e\xab\x04\xd555,]\xba\x8c\xc7\x9e~\x90\x8e\xe3\xa1\xfd\x80\xd6\xe8B\xa09u\x14\xd5\xbf$\xb9\x134\xceI|j\xe7\x11\xfb\xc8\x12%<L\xc2\x13\xa8\x18/\x14\x16kf\xc1\xb4\x83\xbc=\xfds\xc6\x8e\x19g\x9c\x1b\x9c8\xce4\x038\xb5\xf0v\xa8\xd1\xcc!\xda[\x91\x86\xd6\xca\x8c\x83d\x1d\x81\xdb\xe5\xa6\xba\xba\x9a\xda\xdaZ\x1c\x0e\x07n\xb7\x1b!\xa4\xdf\n\xb3\xd9\x84\xc5b!,<\x8c\xa8\xa8h\xe9e\x0b\x08\xb3\xdb\x89k\xdd\xca\xb0+\xe4}`P\x7f\x9d\\oH\x9c\x16qZ\x138%\x0c`\xc6L\xda\xb5K=E\xadp\xfc8\xc3#\xd9\x87\xbc\x04R\x9a\x004\x11\xfd\xbb\x87\xa6i|\xff\xfd\xf7\x84\xd9#\r9\xba\xec\xe0\xc0nV\x14p\xeb\x02\xcc\x92\xf8\xcb\xb0\xfa\xb7W\x8c4\x18DOx\x88\x1fRS\xa7\xa6\xdc\xc9\xa1\xb5N.\xb9\xf42RSS\x9b\xb5\x12W\x80\xf0\xf0\x08.\xbd\xe4Rb\xcb\xfapl_\x15V\x9b\xc5[;I\x17\x14\xf0\x8a]$\xe4\xb3\x15\x14\xa1ppg\x0e\xa3G\x8f\xc6d\x92\x0c:\xf0\xbd\x1a\x86BxX8\x13&L\xe0\xbd\xb7>"\xee\xe8P~\x9bs\x14G\x95\x1bk\x84/\xb3\xf7\xb4\x05\x86\xd9\x08I*\x8cP\x9f6\x91L4\x10\x8a\xe2\xa90\xd2\xa3\x96a\xe0,D\xd2S\x84\xa6\xdb\xfdTC\x18\x03\xc2\xf7\xc9M\xd7\xc2GD\xe4\xf9\xed\x1bk4\xac\xc9l\xf6\xba\xa6l\x97\xd6\x8e\x0e\x1d:\xd0\xbbwo\x86\x0c\x19\xc2\x88\x11#\x18=z4\xe7\x9e{.\xa3F\x8d\xe4\xec\xb3\xcff\xc8\xd0!\xf4\xe9\xdd\x87\xf6\x19\x19\xa4\xa5\xa5\x91\x96\x96F|B\x02&\xd5d\x94Y\xff\x1c\xdf_\xfeOo\xba\xf6\r\xe1L\x13\xffS\t\xcf\r\xff\x96\xc0\x19f\x00>\x9d\xd4D\x7f\xfdw\xaf\xb2$J\xcaJ\xb1G\x18\x97\x95|E9>\xfa\xfbB\xe8(B\xa0;\x8d\x95\xbdA\xe0=\xb2\r\xcf\xd0\x90dN\xe6S\x8c\xcbT\xba\x80\x9d?\x1c\xe3\xc2\xc1\x13\xb9\xe2\xf2\xcb\r\x1f\xaeM4\xac\x01UU\xe8\xd1\xb3\x07\xb7\xde\xf2G2W9\xa8*u\xa2\x9aU)V\xf2=\xed5\x9e\xec9\xe8ST\xa8\xacp\x11\x91\xd7\x87!C\x86\x9cP?)\x8a\x82\xddfc\xe4\xd9\xe7\xf0\xf2\xcb/s\xe5\xe0\xbb\xd9\xf7\xb9\x9d\xcc5%\x08EA5)\x92h\x0b\x15Y#\x1f\xe2\xe1}?\xf9\xbf\x9eA\xd4\x9f\x05\x181\xf2\x8b\n\x8eJ\'&G8\x8a\xd4\xa5=a:\xd38sm^\xbb\x9fJH\x19\xb2|r}o\x9d\x1cd9\x9eRB\xbf\xafb\\^\xf3\xfb\x9c\xe0\x93\xfd\x9f\xd0\xbc2t]\xc7\xe5r5\xd1\x1fM\xe3\xf8s\x87\xca\x11*\xac\x1eM\xcdGE\x91w:Z\n-\xf7d\xa3\xdd\x1a\xec\xc3\x06#\xfe\xbb\xa0ZA1\xceH\xeb\xdf\xc8\xe7\xe0R\x18\xbf\x8d\xef\xde\x95\xadq\x9b\xd5\x9b\xc4\x10\xfd\xa0\xa8\xa0\x82\xc9\xa6rlk9}b.\xe1\xde{\xee#""\xc2H\xdd<\x08!m\xeb\x8f\x193\x86[.\xbd\x8fm?\x1e\x02\xcdp\xe0\xeb\x117\x195\xf3eC\x16\xbb\x89\xcc\xf5\x85L\x9e<\x99\xb8\xd8\xb8\x93\xea&EU\xc9\xc8h\xcf}\xf7\xdf\xcf\xcb/\xbcBF\xf5X~|o;\x05Y\xd5h\x9a\xc0\x12!w r\x81d\x909\xd9\x18\xc6\xc7\xd3\x8e\xdeF\xf3!D2\xad5\\%\xff`5\xdd\xd2\xfb\xd06%\xf9\xa4\x1c\xc675\x99[\x12\xa7\xaaf\xfe\xe5\x9c\xaaR\x1bF\x83Ohd\\)\x8a\x82\xc5r\xf2Nq\x8e?w\xa8\x1c\xa1\xc2\x8e\x0f\'_\xc2\x89\xa3\xe5\x18@\x08y\xb3/\x8c\xa3\xd0\x13\xc6\xc9\xae\x0eN\x16\xf2\xf9\x02\xc5d\xf8V5H\x93\x82\xb1\x1b0h\x98\x97\xb8K\xfa%\x7f\xfb\xd00\xef\xe0\xf0n\xa3\xa5\x1e}\xee\xaeJ\xeav\xb4\xe3\xf6\xdbn\'\xbd]\xba\'U\xb3!\x0cK\xa5Q\xd1\xd1L\xb9q\n\x97\xf5~\x80\xcd\x0b\xf2\xb0\xd8UT\xb3\xb4\x13\xe4y\xb6b\x88\x82LV\x85\xf2\xfc:\xecG:0r\xe4H\xa2\xa2\xa2\xbc\x0c\xebD!\x84 *2\x8a\xb3\xcf>\x9b\xe9\xd3_\xe4\x9fO|K\xed\xba\x0c~\xfb\xa2\x80=+\x8a\xd0\\\x1a\xb6\x08\x8f\x0cY\xc2\x7f\xe2\x07\xf6\xb3\xe4\xa2\x02P\xcd\n\xb5\x15n\xf6\xad,g\xc4\x88\x91ddd\xb4\xecl\xfb_@`w4\x81\x06\x937\xd2O\'K\xf8\x7fohIR\xd5r\x0c\xc0#\xce\xf0l\xdfO1Zz\x90\xc8\xe7+(\xba\xc9\xcf^\xbf@\x1e\xa2\x1a"v#q=\xf57\x04.\x86HC\x8ac\x14\xc5\xc3@T,a\n%\xb9U8\xb7\xb7\xe3\x89??\xc9\x80\x81\x83\xa4\xcd\xfd\xe3\x84j8\xedP\x80\xd6\xad[3\xe5\xc6)\x8c\xeep\x1d\xbf~\x98Om\xb9&5\x83\x8c\xad\xbd\xa2H&\xe6\xae\x13\xec\xf8\xae\x88\x8b.\xb8\x84^\xbdz5i\x80\xebx`2\x99\x88OH\xe0\xc2\x0b\'\xf0\xe9\xc7\xb3y\xe2\xe6W\xe8\xe8\x1e\xcf\xae\x8f\xed\xfc\xf2\xc9A\n\x8f\xd5\xa0\x98MXl\xd2\x17\x80\xa2(X\xac&lV\x0b\x16\x9b\t\xab\xdd\x84\xc5f\xc2d5I\x13\x01\x8a\x82\xd3)\xd8\xb6 \x9fQ]\xaf\xe4\xa6)S\xb0\xdb\xc3\x8ec\xa45<+C-.B\x04\xf9\xa3\xa9\xf83\x88\xa6\xd4\x1e\x1b\x8fm\x02F\x037Z\x86Od\xf3\xfb\xa3\xe5\xe1\xfbN\xa1\xc6\xc0\x89\xe2L\x90\xaa\x86j{\xeaf\xf0qC\xbeuC"\xd9\x930\xde\xf8\xbb\x81\x10\x02W\x9d@7,\x80\xfa\xb2\x01o\x9f\x1b_\x14I\xfde:cp)\x9e?\xc6n\xc2d\x83\x92\xec\x1a\nW&\xf2\x97\x87\x9ee\xcc\xb9c\xb0Z\xcd\'=\x82\x14E\xa1S\xa7N<\xf2\xc8#\xdc}\xe5_\xd9<\xbb\x96\xec\x9d\x058k4LV\x13f\xbb\tg\x95\x9b\xdd+r\x18\x94|\x11\xd7\xdfp=Q\x01\xbe\x06N\x15\x14E!\xbeu<\xd7]w\x1dS_\x98\xcak3\xdf\xe0\xc1+_%>s\x14_\xdd\xbf\x9f\x85onf\xe7\xcf\xfb\xc9\xd9^JIn5\xe5y\xd5T\x15\xd6PY\\Ky~5%\x87k9\xb6\xbf\x8c\x03k\xb2Y\xfbz\r\xd7\x8cx\x84\'\x9f|\x8a\xe4\xe4\xe4\xc0G5\x81\x86\xdbT\x00B\x97\x17\xfb<h\xb2\x0b\x9a\x8a\xf7A\xa8\xf9\x002B\x08)\x03o\xceG\xd344\xb7f\x98s\x90\x966u]\x07\xbd\xfe\xbb\xf08q\xa7~Wz\x1cUm\x10\xc1e\xc8\xc2\x85\xb1\xbb\x0f\xackS\x1fM\xd3\x82\xc2<\x9fSI\x8c\x9b\x820\xea\xdeH/\xfd.\x11\xdc\x1f\x12-\xc8\x00$\x1a^\xfd7\xb5N\t\xc6\x99\x1c\x08MAU\x15\xbau\xedBEA5\xba\xbb~\xb8xv\x06\xde\xdf\x06}\xf7\x9e\xbb\x1a\xea\x8b\xbe\xf1\x8a\x02f\x9b\x89\xc3\x1b\xcb\xc8_\x9e\xc0_\x1ex\x8e\xb3\x86\x0f3V\xe0\r\xb5\xdf\xf1#11\x91\x1b\xae\xbf\x91\x0f^\xfb\x94\x94\xa2\xf3\xd9\xfcU%\xbf\xcd?\xca\x86\xf9\xb9l\xfa\xb2\x8cn\xd6Ky\xec\xd1\xc7\x0cm\xa3\xc0\xdc\xc7\x89\xc6\xf2\x1b\xaf\xd4\xa6M\x1b\x86\x0c\x19\xca\xf5\xd7]\xc7K/Og\xcb\xe6m\xbc\xfb\xdcw\xfc\xe9\xbc\xff\x10\x9b;\x8c\xfc%m8<\xbf5YsZ\x93\xf5e+r\xe6\xc7S\xbc<\x89\xf0}\x83\xb9q\xe8+\xcc\x99\xfd5w\xddu\xd7\t\x10\x7f\x1a\xaf\xa0qg\xa2\x91\x14\xc1hn\xe2\xc6\xd2\x19\x9bB\xcf\xce\xac\xa9\xbeW\x94\xfa3\x12\xdf\xb9\xe1\xfb\x08a\xcc\x1b\xa1\x0bt\xa1\xcbO\x00\x81\xc5\xd7\xbeO\xc0\xc7C\x80=\x0b\x17\xf9\xd1\x11F9\xc2C\xc0u].\x84\x02\x18gs\xd1\xe0\x8e\xdeh\x0f?\x1cW\xf1\xcdO\xacP\x9f\xfc\x04^!$\x8e\x9f\xca\x9dZ\xb4\xf8=\x00i\x0b\xe8\xf9\x10\x17\xc1\xeeg\xc6\x8c\xff^[@B\xe8|\xfb\xed\xb7<?\xeb*\xfa_\xd2\x13\xa1\x1bV@}4ld\xe7\x1b,P1\xd4z\x0c\xad\n)\xfe\x01\xd5\x04.\xa1\xb1we1\xc9\x15C\xb9\xe3\xb6;9\xf7\xdcs1\x99\rO[\x81\x0f>\x05\x10BPYY\xc9\xa6\xcd\x9bY\xbbf\r.\x97\x8b^\xbd{1b\xc4\x08\xe2[\xc5\xa3\x1c\xaf\x1e\xbd|M\xff\xa0\x13\xd4\xdb\xf6\x10\x19\xcf\xadT\xa7\xd3\xe9\x9dD\n\nv\xbb\x1dEQ0\x9b\xcd\xa7\xdd\xab\x9c\xe2\xe9\xab\xe3D\x88\xe68a\xe8>\xe6\xc5\xff\xd7\xd0t\xfb\x1b+\xab\xc6\x924\x07>\x1d\xe6\x19\x7f\xcd\x85hQ\xa7\xf0M\x8f\xb4\x96\xdf\x014\xbf-\xffk \x00\x14\x95\xee=zb;6\x14G\x85\x90\xea\x92B\xf1\xda\xaa\xc7 X\xd2\xb0\x9b1\xa8\xbc\x83Y`2\x81\xc5\xaaR\x9a\xe7`\xedG9\x8cJ\x98\xc2\xd4\xe7_d\xec\x981\x98L\xa6f\xac\xff\x8e\x07\xfer8\xc5\xb8\xc59z\xd4(\xfe\xfc\xe7\xc7x\xea/\x7f\xe1\xf2\xcb/\'>>^j\t\x1d/Bdi|\xe26\x0cEQPU\x15\x8b\xc5Bxx8\xb1\xb1\xb1\xc4\xc5\xc6\x11\x17\x1bGll,v\xbb\x1d\x9b\xcdv\xda\x89\xff\xc9\xe0\xc4\xde<4TEEU\xa5yn\x0fA\x0cj\xdb\xff\x8f\xe6\x98\xa7\xffC\x99\xaa\x08\x89f$i\x12\'[F\x13\xed\x7f<\x0c\xe5\xf8\xd0t\xc5[\x9c\x01H\x047@\xd3U\xff\xfd\xc2C\x9cS\x92\x93\xb9\xe1\xfa\x1b\xd8\xba\xf8\x18&\xb3\xc7+V\xfd\xbb\t\x90\xb7~\r\x1f\x01\x8a\x02\xaaI\xc1\x12a\xa2\xaa\xdc\xc9\xc6o\x8e\x91\xbd0\x9c\xc7\xa7\xbc\xca#\x8f<L\xe7\xce\x9d\xebw\x08\xa7\xb4\x85\x1a^%)\x8ab\xd0|\xf9\xcc\xd37X\xff\x0f\xa1\xe6A\xf3 \xfb\xc6\x97\x01x\x88\xa4\xaa\xaa\xf5;\xb6F\x8a\xaf_\xd96\x92\xe8\x0c!\xf0=\x145\x04\xd1\x17\xcd!\x9c\r\x0c\xea3\x8dF\xaa!\x8c\xf7m)\xb4 \x03\xf0\xbd\xc5\x18\xdc\x00\xba!g\xfdo@\xfd@\xf4\xafqDD\x04c\xc6\x8ca@\xd2%\xac\xfd,\x07\xad\x0e0IU~\xb3U\xc5j3\xa1\x98%\x13P\x15\x05\x14\x0b\xe5\x05N~\xf90\x9b\xcc\xcf"\xb8\xb0\xfbm|\xf0\x9fY\\\x7f\xc3\r\xb4j\xddZ\xae\xfcO\xcb`i\xac\xa5\r\xe2b|;\xa56[\x1a{\xec\x19A\x8bW\x00\xfcj\xd1@\xdb\x9ed5\x15C\x8bK\xf1\xec\x12\x02\x13\xf8\x10\xdd\x06\xebp:\xe19\xd70\x1c\r\xf92\x00\x19\xedk\x04\xc4?Oc\x08dg\r}oI(\x1e\xfa\xd1B\x15jA\x06P\xbf\xea\x0c\xd5\x8d\xbe:3\xbfw\xd4\x0fD\xff\x1a+\nt\xea\xdc\x99\x87\x1fz\x88\x91\x1d\xaf\xe3\xf3i[\xd9\xb9|?\x85\x87\xaa(\xcb\xaf\xa6<\xbf\x86\xa2\x03\xd5\x1c\xd9V\xc6\xce\x9f\xb2y\xff\x8eM\x94,K\xe5O\x17?\xcf\x1b\x7f\x7f\x93G\x1ey\x84\xee\xdd\xba\x9f\x01qF\x0b\xb5t\x0b=\xb6\x1e-^\x01hN-\x9aL\xd04\x14c\x9c\xd63\x03\xb9\x9a\x96\x1f\xe3\xbbLU\xbf\xcb\xf4\x10`o:\x19\xe4G\xa0\x8d,>\x91F}e\x84oZ\xc50\xc7\xe1\xfd\xae\xca\x85\x8fJ}XC/\x1b:\xb4q\x04\x96\xd6\xd0\xf7\xd3\x8d&\x8c\xa46\x92\x00\x00\x84\x05IDATi{`E\xcf Z\x90\x01\x18Ph\xa4\x89\x1a\n\xffo\x814\x87\xdb\xb3gO\x9e~\xfai\x96\x7f\xf6\x0b\xd7\r~\x89\xba5\x9d(\\\xdc\x96\xfcE)T\xaeJ\xc3\xb4\xb7\x07\xd7\x0f}\x81\xc5\xdf,\xe1\xc3\x8f>\xe4\xa6\x9bof\xc8\xe0!DFF\x9e:u\x83\xff\x1f\xd0`S4\x18qza\x88!\xfe\x1b?\xfe\xe3\xca\xf8\xeeK\xbf\xbdt=\x90l\xfaS*\xe3\xa6\x887F~\xea\x99E |\xf7\xf5\x9e\x85\xaf\xbf\x06\xd1\xef\xfd\xe3\xf7:\xcdB\x88f\x08@\xd3)N\x0cMW\xb6\x05\x19\x80\xac\x9c\x12`p\xcc\x03\x05\xfe\x8b\xf6\x00\x8dCUU\x12Z\'p\xd6Ygq\xd7\x9dw\xf1\xed\xb7\xdf\xb2d\xf1w,Y\xbc\x84%K\xbec\xfe\xfcy\xdcq\xfb\x1d\x8c\x1f7\x8e\xb6\xc9\xc9\x84\x87\x85\xf9\xcf\xc2\xff\x83D\x83m\xd1`\xc4\xc9\xa3\x919\xf4\xdfE\xb8\xfe\xefs*>\x8d\x0e\x88\x13\x80~"\x1c\xa5\xd9hz^\xb4 \x03\xf0\xad\\\x88\x8a\nBI\xfe\xfe{al\x9dM&\x13V\xab5\xe8c2\x99\x8e_\xbd\xf2\xffp\xfa\xf1\x7f]\xf2\x7f8I4N\xc5N=S9\x1e\xb4 \x03\xf0\xc0c\xea7 T\xf9\xffg\x07\xf0\x7f8\xf5\x90\xab\xb1\xff\xc3\xff\xe1\xf7\x8f\x86\xa8\x98\x10\x9e\x1b\xa0\r\xa58\xfdhq\x06\xd0\xd0\xae\xca\xa4J]\xf7\xdf\x0fBT\xf2\x7f\x1c-\xd9"\xfe\xb2\xe9\x13CK\xd6\xff\xbf\x1b\xa7\xb7\xe5\x8e\xbf\xf4\xe3\xcfq\xa6\xd0X\xcd\xbc\x07\xeb\x8d\xa6:\xbdhq\x06@\x03\x8d\xf4{Q\x02\xd5u\x1d\xb7[;\xa1\xc3\x9f\x16\x83q8\xa9\x9ff\x83J\'O\x82[\x18\xba\xbcQ\xfc_\xd5\xb7\xa7\x1c\'\xf2\xf2\xa7\xa7\xe7].\x17\xba\xae\x9f@\xe9\xc7\x9f\xe3LAi\xa2\x85\x15\xa4\xdf\xed\x96\xc2\xef\x82\x01\x84\xec\xc0\x96S\x8dE\xd7t\x1cN\'\xa5\xa5\xa5l\xdf\xbe\x9d\x1d;\xb6SZZ\x82\xd3\xe9\x94\x04\xe3$k&|\x0csI\x06\xe3\xc6\xedv\x1bF\xa6N\x1c\x9a\xa6\xe1t:)+/\xe3\xc0\x81\x03\x1c<p\x90\xba\xba:\xa3\xde\'W\xb6\xef\x1bk\xba&\r\x8d\x9d\xc0\xc7\xf7\x1d\x85a\xd2\xc1\xf3\xfe\x9ex\x8f}\x99\xc6\xa0\x1b\x06\xc2\xdcF\xb9n\xb7\x1b\xb7A@\x1a\x83\xae\x0b\x1c\x0e\x07\xa5\xa5\xa5l\xdc\xb4\x91\x9d;wR^^\x8a\xc3\xe9h\xf2\x99\x1ex\xfaO\xbeS}\xdd\x9b\xca/\x84\xf0\xa6u\xbb\xdd^\x07\xeb\x9a\xae\x05&m\x14\xbe\xe3\xc7S\x07M\x93\x86\xdf\x8e\x07\xb2\x0c\xff\xb2<\xf5k\x0e\x02\xf3\x05\xbaum\n\x9e\xfew:\x9d\x14\x14\x16\xb2m\xfb6\xb2\xb3\xb3\xa9\xa9\xa9\xf5:|i\xbcE\xeb\xe1)\xebD?M\xf5\xdd\xc9 \x04u\xf3"\x94\x02\xcc\x99D\xcb\xd9\x02\x12\xd2\xf4\xc1\xb4i\xd3x\xee\xb9\xe7\x82\x06\xcf\xbd\xf7\xde\xcb\xccW^\xc1r\x86l\x01\t!\xa8\xad\xab\xe5\xd8\xd1c\xac_\xbf\x9e\xe5\xcb\x97\xb3k\xf7n\n\xf2\xf3Q\x14\x95\x84\xc4x\x12\x12\x12\x19?n\x1c\xe7\x9e{.iii\x84\x85\x87I\x07\xd8\x81\x85\x19\x10!\xec\xdd\x94\x94\x94\xb0r\xe5\n***p\x18\x0c\xa5\xa8\xb0\x08\x97\xdb\xcd\xf0a\xc38\xf7\xdcs\xb1X-\x86b]3 \x04U\xd5\xd5\x1c\xcb\xcb\xe3\xc7\xe5\xcbY\xbat)\x05\x05\x05TTT\x80a\xea955\x95\xf3\xcf?\x9fa\xc3\x87\x91\x92\x9c\x82\xddn\x97\xba\xdeH\xab\x96\xc7\xe3\x91\xc8\xe9t\xf2\xf1\'\x1fS[S\x8b\xa3\xce\x81\xa2\x82\xdb%\'\xbf\xaa\x9a\xa4\xb3\x1aU\x05EAh\xd2H\x98\xeas\xf8\xdd.#\x83\x8b.\x9c\x80\xc9d\xa2\xaa\xba\x8a\x1f\x7f\xfc\x91\xec\xecl*++\x89\x8e\x8a\xc6f\xb7\x11\x1d\x1d\xcd\xd0!Ci\x97\x9e\x1e\xf2\xe2\x99\xd3\xe9d\xce\xd7_S\\TDeU\x15f\xb3\x89\xaa\xaa*\x9c.7\x7f\xba\xf3\x0e\xd2\xd3\xd3\xfd8\x96\x00\\N\x07yy\xf9l\xd9\xb2\x85E\x8b\x16\xb3}\xfb6\xf2\xf3\xf31\x99\xcc$$\xc6\x93\x96\x96\xc6\xf5\xd7\xdf@\xdf>}IJN\xc2b\x96\x8e\xcd\x03!t\x9dm;\xb6\xb3n\xdd:j\xaak\xa8\xab\xad\xc5\xe5r\x93\xd8\xa6\rW\\q\x05\x89\x89\t\x81Y\xbc\xc8\xcb\xcf\xe3\xf3\xcf>G\xd34\xaa\xab\xab\xa8\xae\xae\xa6\xbc\xbc\x9c\xdbo\xbf\x9d\xfe\x03\x064\xab\xcfKKKY\xb1b%55\xd2?omm\r\x0e\xa7\x13U5\xd1\xabg/\xce;\xef<\xaf\x8e}c\xd0u\x9d5k\xd7\xb0\x7f\xdf~\x00\\n\x17\x15\x15\x15\x14\x16\x14\xd0\xbbwo\xae\xbez"\x16\xab5d\xfb\x03\xd4\xd6\xd4\xb2r\xd5Jr\x8e\xe4P]]MAA\x01\n\nw\xdcqG\x93\xb6m4\xcdMee\x15;v\xec`\xe3\xa6M|3\x7f>5\xb5\xb5\x14\x16\x14\x10\x13\x13CLL\x0c\xed\xdb\xb7\xe7\x8a+\xaf\xa4w\xef\xde$\'\'Kow\x08\x1fr\xea\xfb\x1dJJ\x8a\xf9\xe2\x8b\xafp\xbb\xeb\x17\x01\x9a\xa6\x81\xa2\xc8\xdf\x06\xbdA\xde\xbdGUM\x86\xaa\xaa\t\x8b\xc5\xc4\xe8sG\xd3\xad[7\xe964\xa0\xec\x93\x81.\x04\xf3\x1b\xb1\x054c\xe6L\xda\xa5\xa6\x9d\xaa\xc7\x1d?4M\x13-\xf9y\xe1\x85\x17\x84\xd9l\x16F\xab{?\xf7\xde{\xafp\xd4\xd5\x05\xa5?]\x9f\xac\xac,\xf1\x8f\x7f\xfeS\xf4\xed\xdb7\xa8.\x81\x9f\xc46m\xc4\xab\xaf\xbd*\xb2\xb3\xb3\x85\xdb\xed\x0e*\xab\xc1\x8f\xae\x89%K\x96\x04\x95\xe7\xf9\xdcr\xeb\xad"//_\xb8\x03\xf35\xf8q\x8b\xdc#G\xc4[o\xbd%\xda\xb6M\t*/\xf03t\xe8\x10\xf1\xce;\xef\x88#G\x8e\x08M3\xea}<\xf5\xd74QSS#\x86\x0f\x1f\x1eTvs?\xaf\xbe\xfa\xaap:\x1d\xc2\xedv\x0b\x87\xc3!\x16/^,\xd2\xd2\xd2\x82\xd2\xcd\x9a5K\xd4\xd4\xd4\x04=\xdf\xedv\x8b\xdd\xbbw\x8b\x8e\x9d:\x05\xe5\x99\xf1\xf2\xcb\xa2\xce\xe1\xf0\xcf\xa3\xeb\xa2\xa0\xb0P|6\xfbS1~\xfc\xf8\xa0<\x81\x9f\xf1\xe7\x8f\x17_~\xf5\xa5(++\x93\xcf\x0b|\xbe\xcb%\xbe\xfb\xee\xbb\xa0|W]u\xa58x0K\xb8\xdd\xc1m\xe6\xa9wvNvP\xbe\xd6\xf1\xf1b\xcd\xda5B\xd7\xf5\xa0<\x81\xf95M\x13+W\xae\x08*\xc3\xf3\xb9\xed\xb6\xdbDAaaP\x9dC}\xdcn\xb7\xf8\xe2\x8b/\x83\xca\x00\xc4\xe8\xd1\xe7\x8a-[\xb6\x08\xb7\xdb\x15\x94\xcf\xf3\xa9(\xaf\x10/\xbc\xf0\x82_\xbe\xc4\x84D\xb1~\xfdz\xa1iz\x83\xf3\xc2\xe5v\x89\xd5\xabW\x8b\x07\x1f|0\xe8\xb9\xa1>\x97]v\x99\xf8\xf6\x9boDUUUPY\xbe\x9f\xec\xec\xe0\xb6m\xee\xa7c\xc7\x8e\xe2\xe7\x9f\x7f\x16.\x973\xa8\xdc\x93\xfd\xb8\xddn\xf1\xd5W_\x05=\x13\x10\x13\'N\x14\x07\x0f\x1e\x0c\xcas&?\xcd_\xfa\xb5\x04\xce\x10W\xcc\xc9\xc9a\xda\xb4i\xdcs\xf7\xddl\xdd\xba50:\x08\x05\xf9\xf9<\xfb\xcc\xb3\xbc\xf9\xe6\x9b\x14\x15\x17\x07F\x87\x84\x10\xd2\xe4\xee\xe6\xcd\x9b\xb1X,\x81\xd1\x00d\xee\xdfO\xee\xd1\xdc\xc0\xe0\x06QRR\xca?\xff\xf5/\x9e|\xeaIrs\x8f\x06F\x07a\xdd\xba\xf5\xdcv\xdbm\xbc8u*\x07\x0ff\xc9\xc0\x06Vx\rA1\xdc\xf1\x9d\x08\xc2\xc2\xc2\x8c\x15\xa5\xbc\xf5i2\x99\xe9\xdb\xaf\x1f\xe3\xc7\x8f\x0fZm/Y\xbc\x84\xca\xcaJ\xbf0a\xe8b/_\xbe\x9ccG\xfd\xdf\xb7m\xdb\xb6\\~\xc5\x15>;F\x01@EE\x05\x1f|\xf0\x01\xf7\xdc{?\xcb\x96-\xf3\xcb\x13\n\xcb\x96.\xe3\xee\xbb\xef\xe1\xbd\xf7\xde\xa7\xa2\xa2"x\x08\x1a\xf6i\xc2\xc3\xc3\xfd\x82UCi\xa1\xa1\xe6T\x8c\x9b\xb6V\xab\xd5/\xdcb67\xcb)\xb8\xa7}6n\xdc\x88\xddf3\xc2\xfc\xd3\xec\xd9\xb3\x87\x9c\xecl\xb9\xb3k\x86H\xc3\xb3K\x08\xac\xf2\xda\xb5k\xd8\xbcy3N\xa7\xbf\x85^_\x08\xe3~\x8b/\xac6\xf9nre\xed_\xaaGl\xf5\xe3\xf2\xe5<\xfe\xf8\xe3\xbc\xfa\xea\xab~\xf1\r\xe1\x9bo\xbe\xe1\xe9\xbf\xfe\x95\xa5K\x97RWW\x17\x18\xed\x85\xa2\x04\xb7ms\xa1kR\xe4x<;\xe1S\x05EiY\xfbZg\xfe\x8d\x9b\t\xd3)\xb6u\x1f\x08!\x04\x9a\xaeST\\\xcck\xaf\xbd\xc6\x87\x1f~\x18\x98\xc4\x8bP\xa6\x18\xca\xcb\xcb\x999s&\xff\xf9\xcf\xbf\xa9\xa9\xa91B\x1b\xeeHEQ(*.\xe6\xc0\x81\x03A\xe2.\x0f6l\xd8@\xee\x91#\xd2fz#\x10\xba\xa0\xb6\xb6\x96\x7f\xfd\xeb_L\x9b6\x8d\xd2\x92\xd2\xc0$\x98\xcd\xe6\xa0\t\xea\xc1;\xef\xbe\xcb+\xaf\xbc\xc2\xb1c\xc7\xe4\xd6\xf8\x0cAQ\x14\xccV\xc3\x97\xab\x10(\n$&&p\xd9e\x97\x92\x98\x98\xe8\x97v\xf6g\xb3\xd9\xb9k\xa7w;/\x8c\xd1\x90\x9b\x9b\xcb\xaf\xbf\xfe\xea\xd3\xe6\x12\xf7\xdd\x7f\x1fm\xda\xb4\xf1\tQ\xa8\xad\xabe\xce\x9c9\xbc\xfc\xd2K\x14\x17\x17\xf9\xc4\xd5#\x14\xd1(,(\xe0\xc5\x17\xa7\xf2\xf5\x9c9\xd4\xd5\x06\x13\x9d\x86&\xac\xb4\xf4\xdd\xd0\xe1U\xc8@\t!\x1a\xcef@\x08AEe%Y\x87\x0e\xe3p:\x8d0\xff4\xeb\xd6\xad#\xebP\x16n\xb7;\x88\x00\x87\x82\'{\xe0s\xeb\xea\xea\x987o\x1e\x05\x05\x05\x011\xf5P\x14\x11"g\xc3\xd0\x85`\xf3\x96-\xcc\x981\x93U\xabV\x05F\x83w\xce\x07c\xeb\xd6\xad\xdc\xf5\xa7\xbb\xd8\xb4i\x93O\xdb\x07<\xbb\xf9U\t\x82\xbc\xd0\xd7\xb8Y\xed\x86\xfa\xbc9h4\xa71\x0fZ\n\xa1[\xfcw\x80\x93h\xeffAQ\x14tMc\xdd\xba\xb5,Z\xb4\x08\x97+x\xb53\xe5\xa6\x9bx\xeb\xad\x7f3{\xf6l\x9e\x7f\xfey\xbat\xe9\x12\x98\x84\x7f\xbf\xf5o\xd6\xae[k\xf0\xaa\xc6{\xb2\xa8\xa8\x88#G\x8e48\x98jkk\xd9\xb3g\x0f\xb5\x8d\xact\x00\x14Ua\xc3\x86\r\xfc\xe7?\xff\t\x8c"9%\x99\x17\x9e\x7f\x9e\xd9\xb3g\xf3\xe1\x87\x1f\xf2\xd0\xc3\x0f\x05&\xc1\xedv\xf3\xce\xbb\xef\xb2t\xd9R\xea\x9c\x8e\xc0\xe8\x13\xc2\xe0!C\x185\xea\x1c\x86\x0f\x1b\xc6\xa0\x81\x83\xe9\xdc\xb93\xadZ\xb5"""\x82\x88\x88H""#\t\x0f\x0b\'""R2&\xc3\xff\xb1\xaa\xa8\x8c\x18q\x0e\x17]tQ`\x91|6\xfb3\x1cuu2\xa1\x108\x9dN\xb6m\xdf\xce\x9a5k\xfc\xd2\r\x1b>\x9cQ\xa3Fa\x0f\xb3\xfb\x84\n\x0e\x1e8\xc8\xdbo\xff\x87\x92\x92\x12\x9fp\x89\x1bo\xbc\x81w\xdf{\x979s\xe6\xf0\xdcs\xcf\xd1\xa9s\'\xbf\xf8\x92\x92\x12\xde~\xf7\x1d\xb2\x0e\x1b;%/\xa4vU\xa8>T\x08\xb0\xa5\xe0\x07c\xa5\x17\x90O\xca\xebUIPC\x94\xe9\x8b\xdc#G\xd8\xb7wo\xc8gc\x9c\x8d\xec\xd8\xb6\x83\x9a\x9a\xda\xc0\xa8\x90\x10\x8dh\x89}\xfb\xed\xb7\xac\\\xb9R\xca\xd1C d\x15|\x02\x03\xa3\xab\xab\xab\xf9n\xc9\x12\x96/_\x1eT\xffN\x9d:\xf1\xe4SO\xf1\xc9\xa7\x9f\xf2\xe1G\x1fr\xeb\x1fn\xf5\x8b\x07\xc8/\xc8\xe7\xfd\xf7\xdf\xa7\xb4\xcc\xb3\xd8\t\xd8a\x04=Q.\x82.\xba\xe8"\xae\xbc\xe2\n.\xb8\xe0\x02z\xf7\xe9C\xeb\xf8\xd6DGG\x13\x1e\x1e\x8ej,\xec\xa2\xa3\xa31[,A\xf5\xf2Es\x18j\x83h\xac\\\xf5\xf4.t\x9bD\xa0L\xe8L\x7f\xa6N}A\x98\xcd\x16a\x8c\x19\xef\xe7\xbe\xfb\xee\x13\x8e@y\xee)\xfe\xe4\xe5\xe5\x89\xdbn\xbb-\xe8\xd9\x89\t\x89\xe2\x9f\xff\xfc\xa7\xc8\xce\xce\x16UUU\xa2\xae\xce!\xca\xca\xca\xc4\xda\xf5\xeb\xc4\xd5\x13\xaf\xf6K\xab\xaa\xaa\xb8\xff\xfe\xfbEaQ\xa1p{\xe4\xea!>.\xb7K|\xff\xfd\xf7")))\xe8y\xbe\x9f1c\xc6\x8ac\xc7\x8e5(Cu\xbb5\xe1t\xb9\xc4\xd4\xa9S\x83\xceN&O\x9e,6m\xda$JKKEuM\xb5\xa8\xaa\xaa\x12\xc5E\xc5\xe2\xd7_\x7f\x15\xd7^{\xad0\x99L~\xe9\xcf>\xfbl\x91\x95\x95\x15\xf4\x8c\xc6>\xb5\xb5\xb5b\xd4\xa8QA\xf5\x9e=\xfbS\x91\x9d}X\x1c:|Hd\x1d\xca\x12YYY\xe2`\xd6A\xf9\xff\xe0\x01q\xe0@\xa68p S\x14\x17\x17K\x19\xb5\x8f\xac\xdc\xedv\x8be\xcb\x96y\xcb2x\x83H\xcf\xc8\x10k\xd6\xae\xf5\xca\x94\x0b\n\n\xc4\x9dw\xde\xe1\xf7\\\x93\xc9$\x9e|\xf2IQYY)\\.\x97\xb7\\G]\x9d\x986mZP=\xdb\xb4i#\xde~\xfbmq\xec\xd81QUU)\x1c\x0e\xa7(..\x16\xabV\xad\x12\xc3\x86\r\x0bJ?}\xfatQ\xe7s\x16\xe5v\xbb\xc5\xd2\xa5KEXX\x98_\xba\x89\x13\'\x8a\x83M\xb4evN\x8e\xb0Z\xad~\xf9\x92\x92\x92\xc4\xaf\xbf\xfe*4]oTv\xefr\xbb\xc5\x0f?\xfc \x92\x93\x93\x83\xea\xe8\xfb\x19:\xf4,\xe3\x8c\'\xb8\x0c\xdfOc\xb2i\xcfg\xf4\xe8\xd1"\xf3\xc0\x01\xa1\x87\xa8[yy\xb9x\xf1\xc5\x17\xfd\xd2\xb7m\xdbV\xac_\xbf>\xe8<\xc3\xedv\x8b\xd5\xabW\x8b\xd4v\xa9A\xcf\x185z\xb4X\xfe\xc3\x0f\xa2\xa4\xb4D\xd4\xd5\xd5\x89\x9a\x9aZ\x91\x97\x97\'\xe6|\xfd\xb5h\x1d\xdf\xda/mBB\x82\xf8\xf6\xdbo\x83\xdeE\xd34\x91\x13\xa2m\x13\x13\x13\xc5\x92\xef\xbe\x13yy\xc7Dnn\xae\xc8\xce\xce\x16\x99\x99\x99\xe2\xc0\x81\x03b\xff\xfe\xfd"33S\xec\xdb\xb7W\x1c8p@\x94\x94\x94\x08\xa7\xf3\xd4\x9f\x01\xb8\\.\xf1e\x03\xed<q\xe2D\x91\x95u(\xa8m\xcf\xe4\xa7\xc5w\x00\x929\x8a\xc0`\xe3\xe4>0\xf4\xd4A\x18[\xd2w\xdey\xc7/\xdcf\xb3q\xcf\xbd\xf7p\xdd\xf5\xd7\x93\x9c\x9cBXX\x18\x16\x8b\x99\xa8\xa8(\x06\xf4\xeb\xcf\x9dw\xdcI\x8f\x1e=\xbc\xe9u]g\xf3\xe6\xcdd\x1f\xcent\x05Q[S\xcb\xae]\xbb\xc8\xcb\xcb\xf3\x86)\x8aBDD\x84_\xba\xe5\xcb\x7f //\xaf\xc1\x15\x87\xa2\xc0\xd1\xdc\\v\xec\xd8\xe1\'J2\x99T.\xb9\xf4R:w\xe9Btt\x14v\x9b\x9d\xb0\xb00b\xe3b\x192d\x08\xcf\xbf\xf0<\xe3\xc6\x8d\xf3+k\xc7\x8e\x9dl\xdd\xb6\xadI\xd5\xc9\xe6 **\x9a\xe4\xe4\xb6\xa4\xb6M%-5\x8dvii\xb4KkG\xbbvi\xb4k\x97Nzz\x06\x19\x19\x19\xc4\xc6\xc6\x06m\x96\x14E\xa1o\xbf\xbe\xdcq\xc7\x1d\xe03\x1a\x8e\xe4\xe4\xf0\xed7\xdfP[[\x83\xae\xeb\x1c8\x98\xc9\xbf\xff\xed\xbf\xebi\xdf\xbe=\x97]v\x19\xe1\xe1ar5\xa5\xc8\xbe\xcd<p\x80Og\x7f\xea\x97\x16\xe0\xe1\x87\x1f\xe6\xfa\xeb\xaf\'11\x91\xb0\xb0p\xccf\x13\xb1\xb1\xb1\x9cu\xd6Y\xcc\x9c9\x936II~\xe9g\xcf\x9e-\xc5*\x9a\xe6w\xaf"d\xff\x84\xea~\x9f0\x99\xc3?\x9f\xa2(\xa8&\xb5\xc9\xcb@N\xa7\x83cy\xc7\xfc\xc6O(\xac[\xb7\x96\x83YY\x8d\x8e\xc5\xe6b\xc3\x86\r,\xff\xe1\x07\\nW\xf0\x1a5(\xc0\xd8\x01y#\xea\x9f\xefr\xbb\x98\xf5\xc9\'\x1c\xc9>\xe2\r\x03HNN\xe6\x91\x87\x1ff\xe4\xe8Q\xc4D\xc7`\xb1X\xb0\xd9\xac$$$0n\xdc8\x1e\xff\xf3\xe3~\xe9\x8b\x8b\x8bY\xb9r%\xd5\x01\xe2?\x8c\xf9\x1c\x08\xd5\xa4\x12\x1b\x1bCbb\x1b\x92\xda$\xd1\xb6m[222HOO\xa7}\xfb\xf6\xb4o\xdf\x9e\x8e\x1d;\x91\x9e\x91AtL\x8cwGp*\xd1\xe0\x86\xd0\x80\xa2\xd0\xa2\x0e\xd0[\x9c\x014\x06q\x1a\xdc\x85\xc9\x12\x05.\x97\x8b\xef\x96,\t\x8c\xa6}\xfb\xf6\\u\xd5UDGE\x07\xa9\xd2\xa9\xaaB\xff~\xfd9w\xf4\xb9~\xe1+W\xaed\xdf\xbe}\xb8]\xa1e\xfb\x18\x97\\\xb2\xb2\xfc\xc5\t]\xbbu\xe3\xa9\xbf\xfc\xc5/\x0c`\xcd\x9a5\r\x12ea\x0c\xf6\xc0\xad\xb9\xc5b%&&\x06\x9b\xd5\x1a4;UU%==\x9d;\xee\xb8\x03\x93I\xe5\xbe\xfb\xee\xe5\x9f\xff\xfa\'s\xe7~M\xff~\xfdB\x13\xb3\x13\x80\xe7\xf0O\x91_\x8cr}L\x0772\x15\xa2\xa3b\x980a\x02\x1d:t\xf0\x86i\x9a\xc6\xaf\xbf\xfe\xca\x9e={p:\x9d,\\\xb0\xc8/\x0f\x86(\xa7O\x9f>R\xcb\xcf\x08S\x14\x85\xdf~\xfb\x8d\xecC\x87\xfd\xd2\x8e\x1a}.\x97_q\x05v\xbb=\x88\xdc\xaa\xaaJ\xb7\xee\xdd\xb9f\xd2$\xbf\xf0\xed\xdb\xb7\xb3{\xf7nc,4L\xa6=\xe1A\xf1>\xaf,B\xca{=#\xb2\xb1\xd6\x91\x0b\x88\r\xeb7\xf8\x11\xba\x94\x94\x14\xae\xbd\xf6Z\xbft\x00\xdb\xb6n\rI\x10\x03\xd1T\x8a\xaa\xaa*\x96}\xbf,\xe4\xe2&\xd4\x9d\x83\xfa\x14\xfeos\xf4\xe81V\xad\\\xe9\xfd\xed\xc1\xf9\xe7\x9f\xcf9\xe7\x9c\x83\xaa\xa8A\xe5\x87\x87\x85s\xdey\xe7q\xce9\xe7x\xc3t]\'33\x93\xfc\xbc\xbc\xa0\xf4\x18\xfd\xee\xf7\xdbk\xa7\xb4\xbe:\x9e\xb1\xe8\x9b\xd632\x1bk\xff\x13E\x88j\x06AQZ\x8e\x0c\xb7\xdc\x93\x9b\x81\xe6\xe8E\x1f7\x84\xd4\x07\xce\xcd\xcde\xf9\xf2\xe5\x81\xb1\x0c\x1f>\\\xea\x91\x87\x98\x1e\x8a\xa2\x12\x13\x17\xcb\xc8Q#9\xef\xbc\xf3\x185j\x14W]y%\xaf\xbe\xfa*\xc9\xc9\xc9!\x07%\xc6\xc4\xaf\xaa\xaab\xe1\xa2\x85~\xe1\xc9II\\~\xf9\xe5A\x07\x91K\x97.\ry&\x81g\x90\x86h\x96\xba\xba:v\xed\xdaEm\x90\xfcW\xd6\xc9\xa4\x988\xef\xbc\xf3\xd8\xbbo?O?\xfdWn\xbe\xe9fF\x8d\x1aEjjj\xd0\xc49!4\xf0\xee\xcd\x85\xd9bb\xc0\x80\x01\x8c\x1e=\xda\xaf>+W\xae\xe4\xb7\xdf~c\xe3\xc6\x8d|\xfe\xf9\xe7~y\xda\xb6m\xcb\xa4I\xd7\x1aZI\xf5\xc4\xb9\xba\xba\x9a\xad[\xb7\x06\xad\x14\xcf\x1f7\x8e\xa46\xf2\xb09\xd4\x1bGGF2f\xcc\x18\x9ex\xe2\t^z\xf9e\xde}\xf7]>\xfb\xfcs\xe2[\xc7\xa3k\xf2\x86\xaa4u\x1c\x9c[\xa19\xab\xbdP\x1a\x1f\xdeR\x1bm\xc3\xf2\xf2r\xbe]\xf0\xad_X\xa7\xce\x9d\xb8\xfd\xf6\xdb\xfc\xc2\x00\xbe\xfb\xee\xbbF5f<h\xac\xae\x1e\xfc\xfc\xd3\xcfl\xd8\xb0\x11W\x80\xe2B\xe8\xfb\x01"\x88\xf8k\x9a\xc6\xa6\x8d\x1b\xc9\xcd\r\xd6n;\xff\xfc\xf3\x89\x88\x94;\xe0\xc06U\x14\xe8\xd2\xa5\x0b\xe7\x9e{.c\xc7\x8e\xe5\xb1\xc7\x1e\xe3\xad\xb7\xde\xe2\xc6\x1bo$:&&\xe8\xfc"t\xdb\xfa0\xe6\xc0\xa8\xc0\xdf\xa7\t!\x9b\xc9\x07\xe28.\xbb\x9d\x0e\xb48\x03hpa\xa8\x9c\x9eNR\x0cu\xbc\x82\x82\x02\x8a\x8a\x825C\xce;\xef<l\x86\x9a]((\xc0\xe5\x97_\xce\xa2E\x8bX\xb2d\t\x9f|:\x9b;\xee\xb8\x93\xe1\xc3\x877\x9a\xef\xe0\xc1\x83d\xee\xcf\xf4\x0b\x1b;v,1\xd1\xd1\x9c{\xae\xff\x8eb\xc7\xf6\x1d\xec\xdd\xbb\x0f \xe4\xe1\x96\xc5l!,,,0\x98Wf\xced\xd5/\xab\xa8\xab\xab\x93\x93A\xe0m\\EU\x88\x8c\x8c\xa4}F\x06\xadZ\xb5\xf2:O\x0fB\xf0\xe3|\xd0Hd\xa8\xb2\x8e\x03\xaa\xa2\x92\x9c\x9c\xcc\xe8sG\x13\x1d\x13\xed\x17\xf7\xddw\xdf\xf1\xe1\x87\xefs\xe4\x88\xbf\x08\xe1\xc1\x07\x1f$99\xd9X\xd1\x19CF\x08JK\xcb\xc8\xcb\xcb\xf3\xdbEEFF2`\xe0\x00\xc2\xc2\xfc\xd57}a2\x9b\xb9\xe0\x82\x0bx\xf6\xd9gy\xe0\x81\x07\xb8\xe9\xa6\x9b\x98x\xf5U\x0c\x1f>\x1c\xb3\xf7b\x98\xd4\x1e\x0bEl\x9aZ\xb0\xb85\r=\xe0F\xb6\x10\x86\x06J#.\x90\x84\x10\xec\xdc\xb93h\x07y\xd5\x95W\xd1\xb9K\x17\x86\x0e=\xcb/|\xfb\x8e\xed\xec\xd9\xb3\'\xf4\xbcj\x02QQQ\xd2\xf5\xa8\x81\xe2\xe2b\x16|\xfb-%>\xea\xce\x02\x08\xbd?\xad\x87\x97\xf0\x02\xdb\xb6o\xf3^L\xac\x87\xc2\xe0\xc1\x83\x8c\x8bW\xc1P\x14\x85\xf0\xf0p\x1e}\xec1\xbe\xfd\xf6\x1b^x\xe1\x05\xfe\xf8\xc7?r\xe9\xa5\x97\x12\xdf\xba\xb5\xbcth@ ULC\x8egO]\x02\xa5\t\xa1\x93\x9e\x164f\x92E\xd2\xa3\xe0\xb1t\xa6\x10\xba\xf5\xcf(\x8c#\xbf@\x84\n;E\xd0u\x9d#G\x8e\xe0p\x04k\xc0\xc4\xc7\xc778\x90<0\x9b\xcdX\xadVl6\x1bV\xab\x94[\x86R\x15\xf5@\xd7u6m\xda\x14\x18L\xbf~\xfd\xa4\xf3\xf5\xd1\xa3\xfd\xc2\x8b\x8a\x8b\xd8\xb5k\'\x84 *\x02ABB<\x1d;v\xc4\x1cpK:??\x9f?\xfe\xf1\x8f\xcc\x9f?\x9f\xa2\xc2\xa2\x90\xcc\xa3I4\xfe\xea\'Vf\xb3\xa10z\xd4h:\xb4\xaf\x17\x03a\xe8\x82\x7f\xf0\xc1G~\xfd5l\xd80F\x8c\x18Axx0#t:\x1d\xd4\x05\xf4m\x8f\x1e=\xe4\xf9C3\xfa\xd6l6c1\xd4h=\xde\xb2<\xd0\xf5\x86\xb5uB\x87\x1a\x10\x92\t+\x01\xaa\x8e\x8a!6\xf3\xe1\xd5A\xd04\x8d\x8d\x1b7\x06\x06\xd3\xbf\x7f\x7f\xa2\xa3b\xb8\xe0\x82\x0b\xfc\xc2KKJ\xd8\xb5kg\xb3\xee\x17\x04\xa2G\x8f\x1eL\xb9\xe9&\xbf\xb0\xd9\x9f\xcdf\xe5\xca\x95^\x86\xaa4PU\xff\xb6\x95w6\xaa*+)//\x0f\x12iv\xef\xd1\x1d\xbb=\xb8\xef\x02\x11\x1e\x16\x86\xcdf\xf7\xaa5\x87RmVD\xc3;\x00\x84\xec\xaf\xc0yt\xe6\xd08y\x97un\xa9\xba\xfd\x1e\x18\x80\xe7\x86v\x00\x14\xef\x9fS\x0f!\x04{\xf6\xee\t\xbah\x84q\x06\xd0\x14\x91h.<\x03\xd2\xe5r\xf1\xfd\xf7\xdf\xfb\xc5\xd9\xc3\xect\xeb\xd6M\xfe\xef\xde\xddO\x0cTYY\xc9\xb6\xad[C2(\x05\x05\xb3\xd9\xc2\xa8\x91#III\t\x8c&//\x8f\'\x9ex\x82g\x9ey\x86U+WI}\xf9\xc6F\xe0q\xc1G\xa6\x1a\x009\xcf\x82\x1dh\xe8\xc6j9\xe4\xe4\x0c\x80\xa2@RR\x12w\xdc)\x0f\x83\x1b\x82\xc9db\xd4\xa8Qt\xef\xd1#\x98 (p,/\x8f\xbd{\xf6\xf8\x05\xa7\xa6\xa6\x92\x10\x1f\x1fz\xb0\x1d\x07\xa4\xcf\xda\xc0P\xcf\xfb\xd7\xdb1\xd2\x03\xdf\x1f\x81\xe6v\xa3\x07\x9c\xddx\xda\x8dF\x86\xbb\xdb\xedf\xc5\x8a\x15~a\xad\x13\x12\xc8\xc8\xc8 <<\x8c>\xbd{\xfb-\x06*+\xab\xf8m\xc3F*+\x03W\xddM#22\x92K.\xbe\x98\xe4\xe4d\xbf\xf0\x0f>x_\x1e@+r\xc1\x16\xaa?}\xc3\x14\xe49Pee%\x85\x05\x05A\xe9/\xb8\xe0\x02\xec!v\xb1\'\x84\x00\x99\xbe\x07\xb2\xdd\xe5\xa2\xc5;&\xf5\xe01z\xba\xd1\xd0\x9c!\x88i\x9ey\xb48\x03\x90\x9d\x13\x18\n\x8a\xb9\xe1\x15\xf5\xc9BQ\x14\xdcnwP\xe7[,\x16i\'\'\x00\x9et~\x03\xc7K\xd8\x8cx!\x82\xb6\xf7\x9e\xce=|\xf8\x10\x87\x02\xb6\xef\xd7M\xbe\x9e\x98\x98hL\xaa\x99\xc4\x84\x04:w\xa9\xdfv\xeb\xbaN\xe6\x81Ly\xbb7T\xe3\x00C\x86\x0e\xe5\x8e;\xef\x0c\x0c\x06 ;;\x9b\xb7\xdfy\x9b\x9bn\xba\x89\x17\x9e\x7f\x81\xdd{vSSS\x13\xf4\xbe\xa7\x12k\xd7\xade\xde\xbcy\xcc\x9b7\x8f\xaf\xbf\xfe\x9a\xb9s\xe71\xf7\xeb\xb9F\xd8\\V\xafY\x83\xc3!/05\x06\x93\xc9\xc4E\x13.\xe2\xec\x11#\x02\xa3\xbc\xc8\xc8\xc8\xe0\xaa\xab\xae"2""x\x02\t\xa8\xae\xaa";;\xdb/\xd8d2\xa1\x9aLM\xac\xc7\x9a\x86\x80\xa0]\x01\xc0\xe1C\x87X\xb4h\x11\xf3\xe6\xcfc\xee\xbcy\xcc\x9b;\x97\xb9\xf3\xe61\x7f\xfe|\xe6~=\x97\xb9s\xe7\xb2`\xc1\xc2\xa0\xc3{E\x91~q\xeb\x05\x01\xfe\xf5\x13B\xb0}\xfbv\x0e\x1c8\xe0\x17~\xf3\x94)DGGc2\x99HHL0\xce\xad\xea\xf3\xe4\x1e9BIq\xc9q\xef\xd8\x84\x10\xb4m\x9b\xca\x83\x0f>\xe8\x17\xben\xddz~\xfa\xe9\'\xdc.\x17B\t}92\x10\x8a1\x96\xdd!\x0e\x8cc\xa2\xa3\x83\xec|y\xe6\x13\x1e\xba`\xb4\x86w\xfe!/A\x86B\xa8\xd0\xaa\xaa*~\xf8\xfe{\x9fq8\x8f\xaf\xe7\xcee\xee<\xf9{\xee\xd7s\xd9\xb9s\xe7i\x9d\x17BH[X\xbfW\xb48\x03h\x889*\xa7\xb1S \xf4\x88Q\x90N\xab}\xe1t:\xd9\xb7o?\xbb\xf7\xeca\xff\xfe\xfddff\xb2\x7f\xff~\xf6e\xeeg\xcf\x9e=\xec\xd8\xb1\x9dM\x9b6\xb1w\xdf\xfe\xa0\x832\x0f\xb2\xb2\x0eQ\x18p\xde0`\xe0\x00\xc2\xc3#P\x14Ho\xd7\x8e\xbe}\xfa\xfa\xc5\xef\xda\xb5\x9b\xbc\xbcc\r\xb6\x8f\xcdfc\xca\x94)\xfc\xf9\xcf\x8f\x05F\x811\xf0\xb2s\xb2\x991s\x06\xb7\xfe\xe1V>\xfa\xf8#v\xec\xd8I\x9d\xc3\x11\xea\xd5O\x1a\xd3^|\x91I\x93&1q\xe2D\xae\xb9\xe6\x1a&M\x9a\xc8\xa4k&\x19a\x93\xf8\xe9\xc7\x1fq\xb9\x9bf\x00\x00\xadZ\xb7\xe2Ow\xdeAJ\x8a\xff*\xd4\x83\xc9\x93\'\xd3\xabW\xaf\x06\'\xae\xaa\xaaA"\x87z4\xd0\xa0\xc7\x01\xcf"\xc0\x17\xbfm\xf8\x8d{\xef\xbd\x97I\x13\'1i\xe2D&M\x92\xff\'N\xaco\x87\x07\xee\xbf\xdf/\x0f\x1e\xa2f\xd4\xc9\xf7o}\xbc`\xf7\xee]\x94\x95\x95\xf9\x85w\xef\xde\xdd{\x8e\x93\x92\x92B\xef\xde\xbd\xfd\xe2w\xee\xdc\x19\xf2\xe0\xd5\x1f\xc1\xed\'\x84 ,\xcc\xce\xe8\xd1\xa3\x196l\x987\xbc\xa2\xa2\x82\xa5K\x97\x92\x9f\x9fo\xec\xce\x9b\xd7\x8eB\x88\x90\xa2(S\x80\xb1=\x81 \'\'\x87\xad[\xb7\xb2s\xd7.\xf6\xed\xdb\xcf\xc1\x03\x078|\xf80\x07\x0e\x1ed\xcf\x9e\xbdl\xdb\xba\x95\xcd[6\x93\x9b{4\xa8\xfdC\xd5\xa6\xaa\xaa\x8a\xbf\xfe\xf5\xaf\xdeq)\xc7\xe6$&M\x94\xbf\']3\x89\xbd{\xf7\x06f;\xb5h@\xc2\xed\x81|\x8f\xc6R\x9c^\xb4<\x03h\xf4\xdd\x1b\x8d<a\xe8\xba\x8e\xa6\x870\x01\x1bb\x14\xb9\\.\x96,Y\xcc\xc3\x0f=\xc4\xc3\x0f?\xcc#\x8f>\xca#\x8f>\xc2#\x8f<\xc2\xa3\x8f>\xca\x83\x0f<\xc8\xfd\xf7\xdf\xcf\xe2E\x8bp\xb9\x835w\x84\xd0\xd9\xb2u\x8b\xdf!XXX\x18\x1d:\xb4\xf7\x1e\x1a\'$&\x92\x9e\x9e\xee\'\xce\xd8\xbd{7\xfb\xf6\xedkt\xd5\x9c\xd4&\x89{\xef\xbb\x8f\xbf\xfd\xedoR\x9b\'0\x81\x81\xf5\xeb\xd6s\xef=\xf7\xf2\xe8\xa3\x8f\xf0\xef\x7f\xff\x9b\xc2F\xae\xf87\x85\xa06k&L&s\xb3\xbb\xd3f\xb5\xd1\xaf_?\xfa\xf6\xf5g\x8a\x1e\xf4\xe8\xd1\xa3A\xb3\x01J\x03\xe2\x00\x89\x86\xc2\x9b\x0f\xa5\xc9g\x1c\x1f\x14\x80\x10\x0c\xc5\x03\x87\xc3\xc9\xee\xdd{\xa8\xae\xae\xf6\x86\xc5\xc5\xc5\xd1\xae];,Vi\x93)>>>H\x9bk\xdf\xbe}\xec\xde\xbd\xbb\xd1\xf1\x13\xaa=\x84\x10\xa8\xaaJ\xf7\xee\xdd\x195j\x94w\xa5/\x84\xe0\xa7\x9f\x7fb\xc3\x86\xdfp\xb9\x9c!r\x06C \xcb\nT\xa7\xc6\x98\x83~\xbb\x13\x01\x9b7o\xe1\x9e{\xee\xe1\xe1\x87\x1e\xe2\x91\x87\x1f\xf6\xce\xb1??\xf6\x18\x8f<\xf20\xf7?p?\x0f\xdc\x7f?{\xf6\xfa\x8b\xf7N\n\xa7\xa8\x1f\x1b\x82\xd0\x1b:\xda\xafG\x03]\x7fF\x10z\x16\x9dA(rF\x05\x06\x1b\xce\x92\x83\xc3O\x05TU\x95f\x9c\x03\x9e+\x84\xc0\xa5\xb9\x83\x18O]]\x1dK\x97.e\xf1\xe2\xc5,\\\xb0\x80E\x0b\x17\xb1x\xd1b\x96,Y\xc2O?\xff\xc4\x9a5k\xa4=\xf9\x10\xab\xce\xd2\xd2R\xb2\x0fg\xfb]\xda\xea\xdd\xbb7\xf1\xf1\xf1\xde\xad\xa1\xd5j\xa5}\xfb\xf6A\x9a=\xdb\xb6o\xa3\xb66P\xad\xb3\x1e\xaa\xaa\x90\x9c\x94\xcc]w\xdd\xc5;\xef\xbc\xc3\xc4I\x93h\xdd\xbau`20\xde\xed\xfb\xef\xbf\xe7\xe1\x87\x1e\xe2\xf57^\xe3\xd8\xb1c\r\x12\x9d\xc6\x10\xd8f\xcdG\xf3\x9f\xa5i\x1a\xfb\xf6\xedg\xed\xdau\x81Q\x00\xcc\x9a5\x8b\x8a\x8a\xaa\x90\xe3F\xae8\x1b\xd0\xd2\t\xd4\x049\x014&\xcf=\x11xw\x13\r\x14\x9b\x9f\x97Gff\xa6\xdf\x8ef\xf8\xf0\xe1\xb4KO\xf7\x8a\x8db\xa2\xa3IOO\x0f2\xd2\xb7g\xef\x1ej|\x18\x87\x1f\x04h!\xc6\xab0L\x98\x87\x87\x87s\xf9\x15\x97\xfb\x89\x96\x8e\xe6\x1e\xe5\xeb\xb9\xf3)))\r\xd9\xbe\xfe\x90\x07\xaf\r\x8d\x17\xcd\xed\x0e\xda\x19\x08\xa1\xb3f\xcdj\xbe\xff\xfe{\x16/^\xcc\xb7\xdf~\xcb\xdc\xaf\xbff\xfe\xfc\xf9,Y\xb2\x84U+W\xb1a\xd3&\\Ng\xb0(\xa8\xa9\xea4\x80\xd0\xb5;\xb5h\xac\xad\x14\xc3zHK\xa1\xc5\x19@\x93{\xa4\xd3\x009\xe9\x82\x07\xbf\xa2\xa8\x98M\xa6\x80a\xd1\xfc\xca\x05\x8a\x1dt!(((\x0c\xb2\xff\xd3\xad[wR\xdb\xa6\xf9\xf5|\x9f\xbe}\x88\x8e\xf6W\x7f\xfc\xe2\xf3/\xa8\xa8\xa8ht\x00a0\x90\xb1c\xc70}\xfat\xde\xfc\xc7?\xb8\xf9\x96[\x02\x93\xf8\xe1\xa5\xe9/3\xfd\xa5\x97\xc8\xcd\xcd=n\x19qc0[\xad\xd8\xed\xf2\x06\xb2\xd5j\xc5f\xb7a\x0b3,W\xa26\x9bx\x16\x14\x140g\xce\x1cJK\x83\x8d\xdc\x01\xac^\xbd\x9a\x1f\x96\x7f\xdf\xe0\xd2\xa9\xa1\x15\xba\xa6\xc9C\xc0\x93\x81\xef\x81\xa2/\x14E\xc1n\xb7\x13\x19\x19Ntt\xb4\xdf\'&&\x86\xe8\xa8\xe8\x90\xe7K\xfe\x08$\x88\x82#\xb9G\xfc\xc4\x14\x8a\xa2\xd0\xbe}{\xda$&\xd6\xb3#Ea\xe0\xc0\x81\xc4\xc4\xc4\xf8f\xe7?\xff~\x9b\xb2\xf2r\xbf0\x0f\x84\x02z\x13\x8ehz\xf7\xea\xcdu\xd7]\xe7\x17\xf6\xe9\'\xb3X\xdd\xc8EE\x0f<\xcd\xa3\xaaj\x90\xb6\x1a\r\x18XlL]\xd2\x03O\xaf\x06\x8e\xdb\xc0\xdf\x1e\x98-&\xcc\x16\xa9\xb5g\xb7\xdb\t\xb3\xdb\xbd\xda{4A\x9cO\x15\x1a\x7f\xc6\xff\xbcG\xb0\xc6\xb6H\xa7\xafeB\x11#]\x97\x1e\xa9|;LQT\xe2\xe3\x13\x98|\xddd\xae\xb9\xf6\x1a.\xbf\xe2\n\xce\x1e1"\xe4\xb5\xf1\x104\x87\xbc\xbc<\xb6m\xdb\xe6\xfd\xad\xaa*\xa9\xa9\xa9\xc4\xc7\xfb\xaf\xd4\xbbt\xe9BFF\x86_X^^\x9e\xd4\xe7n\x06T\xd5Dzz:W^q\x05/M\x9f\xce\xf2\x1f~`\xf2\xe4\xc9\xc4\xc6\xc6\x06&\x05\xe0\x93Y\xb3\x98?\x7f~\x88\x8bc\'\x86w\xdf}\x8f\x9d\xdb\xb7\xb3k\xd7.v\xee\xda\xc9\x9e={\xd8\xb9s\'\xbbw\xed\xe6\xf0\xe1\xc3\xdc|\xf3M\xcd \x80\x92\x89\xae^\xbd\x86\xd9\xb3g\x07FyQQQ\xc1\xe2E\x8b\xa55\xd3\x00(\x8a\x82\xc5b%>>\xde/\\z\xbar\x9f\x92%_\xe0\x18\x01\xb8p\xc2\x85\xfc\xf8\xf3O\xec\xdc\xb9\x9b\xed\xdb\xb7\xb3m\xdb6v\xec\xd8\xc1\xf6\x1d\xdb\xd9\xb2e\x0b\xdbwng\xd5\xaaUX-\xc1\xd6G\xf1\x8eG\xff\xca\xb9\xdcnr\xb2s\xd8\xbd{\xb77\xccd2\x91\x9c\x9cLLt\xb4\xdf\xfch\xdf\xbeC\x90VXmm\r\xbbw\xed\n\xaa+\xc6\x93L>\xba\xf4\xf5\x90\x1aL\x18\xe7L\x93\xae\xb9&\xa8-?\xfc\xf0C\xca\xcb\xcbC2Y\x0f<q\xf60;\xd1\xd1\xd1Ai\x85\xf7O=\xbau\xed\xc6\x07\x1f~\xc0\xd4\xa9S\xb9v\xf2\xb5\x9c3rd\x906\x12^\x06\x1f\xa8N\x1b\\\x97\xf8\xf8x\xbe\xf8\xe2+\xf6\xed\xdb\xc7\xbe\xfd\xf2\xccn\xa7!Z\xcd\xcc<@VV\x16c\xc6\x8c\t\x997\xa8r\'\x08\xa5\x81\xbaa0\x06\xd1\x80\x16\xe4\x99B\xcb1\x00\xaf\xda`h\xee-\xf4P\xa1\xa7\x06\x8a"/E\x05\xaeB\xdcn7u\xb5\xb5~\x131<<\x9c\x9bo\xbe\x89\x0f\xde\xff\x80\x8f>\xf8\x90\xcff\xcff\xfa\xb4\xe9$&\x04{~\n\x1c\x94\x9a\xa6QPTH\xae\x8f\xedz\xbb\xddNrr\x12\x0e\x87\x83\xaa\xaa*\xef\x07\x01mS\xdb\xfa\xe5\x07X\xbf~=nM3\x06d\xe3-"\x89\x9f\x85\x84\x84\x04F\x8d\x1e\xcdk\xaf\xbf\xc6\xccWfr\xc1\x05\x17\x04\r\xc2\xf2\xf2r^\x9e9\x83\xfc\xfc|\xbf\xf0\x13EBB<\x1d:v =\xbd\x1d\xe9i\xedHOO\'#=\x83\xf4\xb4v\xa4\xa6\xa6\xd2&\xa9MP{\xfb\xc2\xf3f\xa5\xa5\xa5\xbc\xf5\xd6\xbf\x02b\xfd!\x84`\xe1\xa2El\xde\xb29\xe8\xc6\xb4\xae\x0bZ\xb5nE\xb7\xee\xdd\xfc\xc2\xcb\xcb\xcb\xa9\xac\xac<\xe9\xd5\x96\x10Bj\x13\x05\xb4gDx\x04\xc9m\x92HMM%55\x95\xb4\xb44\xda\xb65l#\xa5\xa5\x91\xda6\x95\xc46m\x08\xba\xf7\xa4(\r\xda\xa1\xaf\xac\xacd\xe3\xa6\x8d~\xab\xed\xb8Vq\x0c\x1a4\x88\xea\xdaZ\xaa\xab\xab\xbd\xe3\'&&\x9a\xb4v\xc1\x9e\xb8\xb6l\xaeo\xa3P\x8c \x18\xf5\xef\xa5(\n\xe9ii<\xf4\x90\xbfE\xd9_\x7f\xf9\x85y\xdf\xcc\xc7\xd5\x88\xe9\x13\tA\x98=\x8c\xc8\xc8\xc8\xa0\xf6:x\xf0\x00\x0eg\xfdmeEQ\xe8\xde\xbd;7\\\x7f\x03\x8f=\xfa\x18\x1f\xbc\xff\x01\x0b\x17.d\xca\x94)A\xaa\xbe\xa2\x015\xd4@X,\x16RRRHO\xcf \xcd\xe8\x93v\xed\xd2HMM\xa5m\xdb\xb6\xb4k\xd7\xae\xc1\x05R 3>Q(j\xb0\xc6\x98\x07\x8a\xa2\x18\x17\xd4B\xc7\x9f\t\x84\x1eyg\x02\x8a\xf7O\x83PN\x93\xc3LEQ\xe8\xd9\xb3g\x90\xc8\x05 333\x88\xd0\x9a\xcdf,\x16\x0b\x16\x9f\xcb_\xa1:5\x90e\xd5\xd6\xd6\xb2s\xfb\x0e?\xddo]\xd7\xd9\xb5k\x17\xff\xf8\xc7?x\xf3\xcd7\xbd\x9f\xb7\xdez+\xe4\xb6z\xf1\xe2\xc5TVT\x18\x84+\xf8\x99\rAQ\x14\xe2[\xc7s\xc3\xf57\xf0\xc2\x0b/p\xd9e\x97\x05&\xe1\xe8\x91\\v5\xb0B<^(\x86*#\xd4\x0b5\xe5\x00\xf7\xfd\x1e\x90\xc9\x07\x8aA\xa0~\xfe\xf9g\xb6l\xd9\xe2\x177f\xecX&N\x9a\xe8\'\xe3.).f\xc17\xdfR\x1e \xe2PU\x85\x98\xe8hb\xa2\xfd\xc5!;w\xec\xa4\xb4\xa4\xe4\xd4\xac\xb6\x1ah\xafPc\x02\x9fp\x05\x08r\xcd,\xa9Y\xc8>p\xd4\xd5\xb1f\xb5\xbf\xe9\xeb\xca\x8aJ\xe6\xcf\x9f\xef7v<\xe3\'\xd4\xbd\x91\x05\x0b\x17\xd6\xdfw\t\xa8\x9fg\xa5\x1f\x08\xdf\xf7\x08\x0b\x0fg\xd4\xa8Q\x0c\x1c8\xc0\x1bVZZ\xca\xac\x8f>f\xedZ\xff\xba\x05C\xc1f\xb3\x11\x1e\x1e\x1e\xf8h\xe6\xcf\x9b\x87+\xc0\xe1\x8cb\x98\xc66\x99MX,V\xaf\xb8\xc6\xb7>r\xd1\x18\xd0^\x8d0\x04cD\xfa\xfc\x0e\xddG\xa7\x13\xa1kV\x8f\x06\xaa~F\xd0r\x0c\x00\xd93\r\xbd\xbb\xa0\xe1\x83\xb1\x93\x86\xa2\x10\x1f\x1f\x1f\xd2t\xc3\xd1f\x1c\x8e\x86\x12\x01\x10bpUUV\xb1\xf4\xbb\xa5~auuu\xbc\xf5\xd6[<\xf1\xc4\x13<\xf9\xe4\x93\xde\xcfSO=\xc5\xdc\xaf\xe7\xfa\xa5\x058v\xec\x18\x07\xb3\xb2P<\xb7EC\xa01\xa7\xd6V\xab\x95\xbe\xfd\xfa\xf1\xf8\xe3\x8f\xd3\xb5k\xd7\xc0h*+\xab\x1a\xcc{\xdc8\xc9bJJJ\xf8\xf1\xc7\x1f\x83\x88\xfa\xf8q\xe3\xb8\xef\xde\xfb\xfcL\x14`8\xb6\xd9\x11\xa8\xc7- *:\x9a\x84\x84\x04?\xc2\x91_\x90\xcf\x91\xdc#\x8dz\xb9\x02c\xe7\x19\x8a\xc8\x18P\x94\xc6m\xf64\x06\x01A\x96\x1f\xbd\xcc!`\xac\x0b!\xc8\xcd\xcde\xed\xda\xb5~\xe1\x9e\xf1\xf3d\x88\xf1\xf3\xfd2\xff\xcb\x86\x18\xe3\'+\xeb\xa0\xc1\xa0\xfd\x11B9G\xc2\xe7\xfdTU\xa5[\xb7\xae\x8c8{\x84\xdfJ|\xd3\xa6M,\taL1p\x0c(\x8aB\xb7n]\x89\x8c\x8c\xf2\x0b\xaf\xa8\xa8\x0c\xc9\xb0|\xe1t\xb9p8\x9d\xc1\xfd\xa0\xc8[\xd5~\xbf\x03\x1b\xd0@P\xde\x16@\xe8\x9aI\xc8\x83\xf2\xc0\xd03\x87\x16e\x00\xf2\xbdCwPC\x1dzJ \x04]:w\xf6\xd3p\xf0`\xe1\xc2\x85\xd2\x96N`\x84QS!\x04G\x8e\x1c\t25\x10\xb8z\x17B\x90\x93\x93\xcdo\x1b~\xf3\x0b?^\x14\x14\x14\xb0z\xf5\xaf\xc6\x05"Y+!\xa4G\xb0\xa2\xe2"\xb6m\xdb\xc6\xab\x7f\x7f5h\xd5\xec\x0b\x93\xaa\xd2\xb7o_\x86\r\x1b\x1e\x18E\x9d\xe3\xf8\xce\x00B\x1d\x9e\xe3\xe9\xaf\x93\xe82\xcf\xce\xe8\xd7_W\xfbM\xdan\xdd\xbaq\xd6\xd0\xa1\x0c\x1a4\x88\x9bn\x9a\xe2\x97\x07`\xc1\xb7\xdfR[[+\t6rP\xc5\xc6\xc4\x90\x91\x91\x11\xac\x15\xb3go\xa3F\xd2\\n\x17\x9b\xb7nf\xceWsX\xb8p![\xb6n\xe1pv6EE\x85\x08C\xbb\x08\x02\x88O\x00\x02w\x81\xbe\xf0\xacpCA\xber}^!\x04\xbbv\xef\xf2MrB(**b\xcd\x9a\xb5A\x17\xd0h@?\x9f\x10s/&&\x86\xf1\xe7\x9f\x1f\xe4\xb5-$\x8c\xac\x9ev0\x99Lt\xe9\xd25\xe8\x80\x1a`\xc3F_\xf1V \x91\x17\x94\x97\x97\x05Y\xfe4\x99LX\xcd\xe6\x80\xa1\xd6H\xab\x9fN:\xd2\x0c\xc8\xba7X\xbbF\xa3\xce\x04B\x8f\xc63\x88\x06\xbb\xe74s\xee\xb0\xf0p\x86\x9dU\x7f\xd1\xc5\x83m\xdb\xb7\xb1u\xeb\xd6\x90\xe3F1\xb6\xbf?\xfe\xf4#\xe5\x01\x17s\xf0vv\xfd\xf7Sq\xc9\xc4\xe1pp \xf3\x00U\xd5\xd5\x80\x82\xd3\xe9d\xcb\x96-|\xf4\xd1G<\xf1\xc4\x13\xf4\xef\xdf\x9f??\xfeg>\xfb\xec3\xa95\x13\xd0l\x9e\x95\xac\xc9d\xa2_\xff~\xfe\x91\x86\xec\xfax \xc7sp\xdf\x04\x87\x1c\x1f\xaa\xab\xabY\xb5j\x15\xdb\xb7\xd7\x1f\x98+\x8a\xc2\xe8\xd1\xa3\xe9\xd1\xb3\'V\x8b\x85\x8b/\xbe\x84n\xdd\xbb\xfb\xe5\x9b3g\x0e\x1b7n\xf4#Z\x8a\xa20h\xd0`Z\xb5j\xe5\x97\xf6\xa5\x97^b\xf7\x9e\x86E^5\xd55\xcc\x9f\xff\r\xd7N\xbe\x96\xcb/\xbf\x9cA\x03\x07\xd1\xabgOf\xcf\xfe\x0c\xb7\xa6\xf9<#\xc4\xe00\x10\xbc\xce\xf6 \xb4\xb2\x83\xa7\x7f\x02\xe1v\xbbY\xf1\xd3\xcf\x81\xc1\xc7\r\x87\xc3\xc1\xde\xbd{)\x0f4\xc6\xd6\xc8\xd9\x83/\x14\xe3\xc0u\xd8\xb0a\\x\xe1\x85\x81\xd1! \xdf\xc5#\xbe\xd5u\x9d\x8c\x8c\x8c\xa0\x03j\x80%K\x16{/\xb8\x05\xb6\x8e"\xe0H\xce\x11\xf6\xec\xd9\x13\xd4>"85\xba\xa6\x85d\xae\xa1\xda\xfcL#\xb0\xb6\xbe\x90\xde\x8f\x82\xfb\xffL!\xb8\xc5\xfe?D\xe0\x00\xc2XI\\p\xc1\xf9\x81\xc1\xe4d\xe70w\xee\\\x8a\n\x8b\x82\xc8\x9a\xd3\xe9\xe4\xe7\x15+X\xb6lYp\x99\x1e\xb1\x81\xf1\xd3\xedv\x87tB\xbex\xc9bv\xed\xde\x15\xf2\xb3q\xd3&\x1e\x0c8p\x13\xc6\x8e\xa3\xb8\xb8\xd8X\xb5\xe9\xec\xd8\xb1\x93{\xef\xbd\x8f\xf7\xdf{\xdf\x9b\xee\x9bo\xbea\xc5\x8a\x15\xd4\xf9\x1c\xac\xe1\x9d\xc0\nGrsY\xb8p\x81_\x1c\x86\xed\x97\x90\xdc\xae\x01(J\x03\xb3*\xc0\xee\x8f0~{,gz~\x0b\xe1\xb1R\xea\x9bUp\xf8\xf0af}<\xcb/<22\x92\xc1C\x86\x10\x1b\x1b\x8b0\xec\xdf\x07\xda\xeb?v\xec\x18\x0b\x17.\xa4\xac\xbc\xcc\xafZ}\xfb\xf6\xa1{\xf7z\xc7=\x1e\xfc\xe3\x1f\xff \xf7h.\xc2 $\x1e\xe8\xba\xce\x81\x03\x07\x985\xebc\xbf\xf4\xd1\xd1\xd1\xf4\xea\xd5\x0bK\xc0\xcd\xd5@\x04\x8d\x87 \x84\xce\xeb\xbb+\xf0\x96\xa0@VV\x16\xeb\xd6\x05\xdf\x83\x987o\x1e\xbbv\xf9\x8c\x9b]\xc6g\xf7.\xb6o\xdf\xce\x8c\x193\xfc\xd2\x0bC\x94T\x1c\xc2\xf2\xed\xf1 **\x8ak\'\x07\xfb\x1e\x08\x84\xef;`\xbc_LL\x0c7N\xb9\xd1\'\x95\xc4\xe2E\x8bY\xb9J\x1a\x99\xf3m\x1d!\x04\xc5%%\xcc\x9b;7H\x04&\x84@\xd3\x02.\x91\x19\xcf\t\xec\x83\xfaq\xe7/\xd6\xd3ui\xb3I\xd7u\x84\xde\xf0\xbd\x91S\x81\x86\xa6\x0b\x9e\xfa\x858\xf7;\x93hA\x06 \x1b\xbc\xa1Iu*\xfb#\xd43TU\xa5w\x9f\xdeL\x9c8\xd1/\\\xd34\xde\x7f\xff}>\xfd\xf4S\x0e\x1f\xce\xa6\xb2\xb2\x92\xaa\xaa*Ih\x16-\xe2\xe1\x87\x1e\xe2\xe0\x81\x83~y\x00\xaf\x06\x90b\xfc)(,\x08R\xe1\x1c8p }\xfb\xf4\xa5k\x97\xae\x01\x9f.t\xed\xd2\x85\xde\xbdz1r\xc49A^\xc2V\xaf^\xcd\x81\xccL\x84\x10\xd8\xedv\x86\r\x1f\xc6\xd5W_\xed\xf7^\x99\x99\x99<\xf8\xe0\x83\xbc\xfd\xf6;\xec\xdb\xbf\x8f\x92\x92\x12\xca\xcb\xcb)\xaf(g\xdf\xbe}\xbc\xfc\xd2K\xfc\x1c\xb0\xa2\xec\xd7\xaf\x1f\xbd{\xf7>.[%\xf2=\x83\xd3\xef\xd8\xb1\x83\xa5K\x97\xb2d\xc9\x12\x16-Z\xc4\xb7\x0b\xbee\xce\x9c9\xcc\x9f7\x8f\xb9\xf3\xa4-\x9c/\xbe\xf8\x82U\xabVI\x9f\xba>y5\xb7\xc6\xf2\xe5\xcb\xd9\xb7_\x9a\xc0\xf6\xa0}\xfb\xf6\x8c\x1a9\xd2\xb0\xe1#5\xb7F\x8e\x1c\xe9w\x96\xa1i\x1a+V\xacd\xf7\xee=~"\x8d\xe8\xe8\xe8\x90\xb6\xf2\xe7|5\x87\x17\xa7\xbe\xc8\x9e\xdd\xbb)-)\xa5\xb2\xb2\x92\xe2\xe2b\xd6\xad[\xc7\x03\x0f<@\xf6a\x7f\x1bB\xe3\xc6\x8dc\xc8\xd0!\xde\xdf\x9e\xcbM\x81cJQ\x94 \xa2\x14\x08M\xd7\xd1\x82\x06\xb6\xef\x0e\xa0\xbe\xccc\xc7\x8eR\x12p\x0f\xe2\xea\xab\xaeb\xe8\xd0\xa1\xc6x\x91c\xa7K\x97\xaet\xed*\xbfw\xef\xde\x8d\x81\x03\x07\x92\x96\x96\xea\x97o\xd5\xaaU\x1c8p0HL\xe9n\xc0\xe7D(b\xa8(\n\xfd\xfa\xf6\xe3\xf6;n\x0f\x8c\xf2C\xa8\x1d\x90\xc9\xa4r\xce\x88s\x82vo\x85\x85\x85\xccxy\x06s\xe6\xcc\xe1\xd8\xb1cTTT\xf0\xff\xda;\xef\xf8*\x8a\xae\x8f\xffvo\xbf7\x15\x08i\x10z\x07i\xd2CWP\x14\x01\x11\x0b\x8ab\x03\xc1\xf2X_}T\x04\x15l\x88>J\xb5+\x88\x14Ai\x8a \x1d\xe9\xd2{\x12R!\x90@zn\xdf\xdd\xf3\xfe\xb1{7wo\t\x01\x02$\xb0\xdf\xcfg\x08wwvgvvf\xce\xec\xcc\x99s\x8a\n\x8bp\xec\xe81|\xf8\xe1\x87\xf8\xe8\xe3\x8f\x15\xf1!\xe5C\xab\xd1\xfa\xd5A\x96\xf5w*c\xb3\xd9\xb0}\xfb6\xfc\xf9\xe7\x1fX\xb5j\x15V\xfd\xf1\x07\x96/_\x81_\x97,\xc1Os\x7f\xc2\xe2\xc5\x8b\xb1d\xe9R\xfc\xbad\t\xf6\xfc\xfb/\xac6\xab\xb7\x08\xab\x1c\xcaY\xbb\xf3\xd4\xa3@ev\xcd\xf0\xf5\x11y\xad\xc3\xe4\xc9\xef\x93.\x80O\xe0\xf1\xcf\x8e\'\xa7\xb3\xcc\x17\xeb\xd5\x08\x0e\x87\x83\x96-_N\r\x1b6\xf4K\x1f\x00uh\xdf\x9e&L\x98@\xef\xbf\xf7\x1e\r\x1c8\xd0\xef\xbcw\x988q"\xe5\x17\xe4\xcb\xf7\xfe{\xdd:\x8a\x8a\x8aR\xc4y\xf3\xad\xb7\xa8\xb0\xb0\xd0/\x1f\x9e\xc0\xf1\x1c\xed\xdb\xb7\x8f\xfa\xf4\xee\xe3w\xffY\xb3fQII\tq\x1cGn\xb7\xe8\x1f6!!\xc1/\x1e\x00j\xd0\xb0\x01=\xff\xc2\xf3\xf4\xde{\xef\xd1\xc4\x89\x13\xa9_\xbf~~q\x00\xd0\x97_N\xa7\xd2\xd2\x12\xbf|\x94\x17\x82\xf9\x04\xaehx{\xc2\xdb\xc4q\\Y\xe09JNJ\xa66\xad[\xfb\xc5}\xe7\x9dw\x14~\xa19\x8e\xa7\xdc\xdc\xf34v\xecXb\xc4m\xbdr\x982e\n\x15\x15\x15)\xf2\x9ay:\x8b\x1e}\xf4Q\xbf\xfb\x02\xa0&M\x9b\xd2\xb3\xcf=G\xefL|\x87\x1e~x\xa4\xdfy\x00\x04\x96\xa1-[\xb6(\xfc3s\xee\xe0>\x81/\xe6_9=#\x83\xb4Z\xa5_\xe6\x98\x98\x18\xda\xb1}\x87\xc2\x8f\xae\xdb\xed\xa6\xc9\xef\xbf\xef\xe7\xf3\xf9\x93O\xa6R\xa1\xcf3z\x07\x8e\xe3\xe8\xd8\xb1ct\xd7]\x83\xfc\x9e\xe5\xcb/\xbf\xa4\x92\x92\xb2w\xcdqn\x9a\xff\xcb/~\xf1z\xf7\xeeMv\xbb\xdd\xef\xde<\xcf\x93\xd3\xe9\xa4\x95\xabVR\xc3F\x8d\xfc\xae\x83\xe4\x13x\xd7\xae]\xc4\xf3J\x9f\xc0<\xcfSaQ\x11}2u\xaa\xdf5\x9e0d\xc8\x10z\xf3\xad\xb7\xe8\xa5\x97^\x0e\xda\x1e\x01\x90\xd1h\xa4\xe5+V\x90\xcb\xe5V\xdc?33\xd3\xcf\'\xf0\xa5\x84\x99\xb3fQqq\xb1_\xbe\xaf4\x94\xe7{y\xc4\x88\x11\x94\x96^~\x9d\xb9\xda\xe1:~\x01H\x90\x9f0\x97\x8e\x07;Qyh\xb5\x1a$\xf6H\xc4\xb0a\xf7\xfa-\x18\x02\xc0\xbe\xfd\xfb\xf1\xfe\xfb\xefc\xc2;\xef`\xcd\x9a2m\x1e\x8dF\xdct\x15\xebe\xac\xacL\x92\x8b#\xba\xfd\xfb\xf6\xf9i\xb34j\xd8\x10z\xbd\xbf\xe6\x91\x0c\x01\xb5k\xd7\x0eh\x92z\xdf\xbe}\xb0\xd9\xed\xd2\x94\x01\x83\xc4\xc4D\xcc\x993\x07\xf1\xf1\xfe{\x07\xd2R\xd30\xfd\xcb\xe9x\xe7\x9dw\xf0\xee\xbb\xefb\xc3\x86\r\x8a\xf3\x1a\x8d\x06\x8f\x8c\x1a\x85Aw\xdd\tC\x056gU\x16\x1a\x8d\x06\x1aV\x03\xc6\xe3:\x12\x00\xe7\xe6\xb0|\xc5r$\xa7(\x9d\xe5\xb0,\x8b\x11#F\xf8\xd8\xfc!DF\x86\xe3\xb6\xdbn\xf3\xd3\xdf\xfe\xf1\xc7\x1f\x91\x9a\x9a\xaa\x18\x05\xc6\xd4\x8e\xc6\x981c\xd0\xabW/\xbf\xf2LNJ\xc2\xcc\x193\xf0\xde\xbb\xefa\xfe|\xffMg\x11\x11\x11\x985c&\xda\xb6\xf5q\x99\xe9\xed\xf6\xf2\x12\x11x\x1e|\xa0Et\xc5\xad\x08yyyHKOW,\xdc\x1a\x0c\x064m\xd2\x04\xe6rL(3\x92a\xb8\xc6\x8d\x9b\xf8\xe5o\xd7\xae]\n{B\x00\x13\xd4\x9eR04\x1a\r\xda\xb5m\x87\x9e\x89\x89~\xf7\xf7\x86\t\xa0\xddg6\x9b1\xf8\xee\xbb1|\xf8\xf0\x80\xd7._\xbe\x1c\x1fL\x99\x82\xcf?\xff\x0c\xa9\xa9e_\xd8\xd1\xd1\xd1x\xf9\xe5\x97\xe5\xf5\x1c\x87\xc3\x81\x9c\x9c\x1c\xbf\xfd\x1f\x81\xeey)\\\xd9\xd5W\x8aoi];.\xad\x06\\\x05Dq\xe8_\x00W\xfaB+\x02\xc3\xb0\x88\x8c\x8c\xc0\xb3\xcf=\x8b1c\xc6TH\xcb\xc1\x12b\xc1\xd3O?\x8d\xb9s\xe7\xe2\xe9\xa7\xca\xa6\x18\xd2\xd2\xd2`\xb3\xdb@\xd2BqZZ\x9a\xa2\x92\xc6\xc7\xc7\xa3^\xfd\xfa\xd0\x1b\x02\xef\x04\x85\xd4p\xa2\xa2\xa2P\xbf~=\xbfMS\x0b\x17.\x84\xd5Z*\xff\xd6\xe9t\xe8\xd3\xa7\x0ff\xcf\x9e\x83\x01\x03\x07^\xd2g\xe4c\x8f=\x86\x97_~\t\xf5\x13\xea\xf9Y?\xbd\xda\x88\xf3\xb0e\xff\xcfHO\xc7\x96-[\xfc\xb4s\xc6\x8c\x1d\x8b\xd8\x98\x18\xc5\xe6:Q\xf8i\xd0\xa9s\'\xb4\xef\xd0^\x11?99\x19+W\xae\x84\xcb]f\xfcL\xa3\xd1\xa0K\x97.\x980\xe1m\xdcv\xfbm\x8a\xf8\xe5\x11\x1f_\x07o\xbe\xf9&F\x8c\xb8\x1f!\x92\xcbBo<s\xc9\xbe\x04\xaa\xc7\xde\xe8t:\xe84\xfe\x03\r"\x92\xae\x15\x07=yyy\xc8\xce>\xa3H\xa3A\x83\x06\xa8Q\xb3F\xc0\x85NoBBB\x10\x1b\x1b\xebg~a\xfe\xfc\xf9~\x1e\xf0\x02\xef\x04\x0e\x0e\xc30\x88\x89\x89\xc1\xddw\xdf\xed\xb7;\xd8\x83\xa7\xdd\xfa\xd6F\r\xcb\xa2q\xe3\xc6x\xe6\x99gp\xd7\xddw\xf9\x9c\rL\xe3\xc6M0\xe9\xddw\xf1\xe2\x8b/\xe2\xee\xbb\xef\x96\x8f\xdbm6\xf0A,\xef^.\x0c\xcbVhQ\xbc\xd2!\x04(\xadJ\xa6\x9cjy\x1d\x9e8\x00A2\x18\xe4p\xa5BDH\xa8\x9b\x80\xb7\xdf~\x1b\x9fN\x9b\xa6pB\xed\xcb\xad\xb7v\xc4\xe7\x9f}\x8e\xb7\'L@\xa7N\x9d\x14\x0e\xcc\x7f\xfa\xe9\'8\xec\x0e\x90\x00\xa4\xa7\xa7c\xef\xde}\x8a\x06\xdc\xa2Es\xc4\xc6D\xfby\xa6S\xc2@\xa3\xd1\xa0}\x87\x0e\n\xe3q\x90L\xdb\xa6\xa5\xa6+\xeei0\x18p\xdbm\xfd\xf1\xe9\'\x9f\xe0\xf3\xff}\x8en\xdd\xfd\xb5\x9a\xbc\xe9\x91\xd8\x03_}\xfd\x15&N\x9c(\x9a\x9f\xbe\xcc\n\xefr\x95ga28</\x9a\xda\x00\xc4\x0f<\xa7\xcb\x89]{\xf6`\xc5\n\xa5\xaf[H\xee2\x03m\xd4\x03\x80\xb8\xd88\xdc\xef\xb3\x18\x0c\x00\xd3gL\xc7\x89\x13\'\x15k\xda\x0c\xcb\xa2O\xdf~\xf8\xe8\xa3\x8f0y\xf2\xfbh\xde\xdc\x7f/\x847\xa3G\x8f\xc6\xf4\x19\xd3\xf1\xd4\xd3O\xa3FM\xa5\x16\x91\x07"\xc1\xcfH\x9f\xdb\xcd\x81H|\xc6@\x90\xf4e\xe8[v\xa2\xebJ\x1e\x8cd\xd8\x8d\x04\x01\x19\x99\x19X\xbd\xfa/E\xbc\xd6\xad[\xa3nB\xdd\n\xb5\x89\xb6m\xdb\xfa\x8d\x90\x01\xd1\xda\xa6\xf7:\x80\xef\x9a\x00*\xf0nY\x96E\xef\xde\xbd\x03\xee)9{6\xb0\xb3v\x0f,\xcb\xa2O\x9f\xde\xf8`\xca\x14|\xf2\xc9\'\xb8\xe5\x16\xa5\tko\xee\x7f\xe0\x01|9\xfdK<<r$BBB\x14fR\n\x0b\x0b\x15\xc2\x96$\xc3v\x17\xcb{\xb9PpOoW\x93k1\xd0-O\xbe0<\xcf_\xfb\xa7\xf6\xe2\xc3\x0f?\xc0\xa4I\xef\xfaux\xcf>\xfb,\xa6M\x9b\x16pj\xa6r\x11G]\x82 6\xce\x82\x82|\x1c:|\x08\xfb\xf6\xef\xc7\xb6\xad\xff\xc0\xe5r\xa1g\xcf\x9eh\xdd\xba5:v\xec\x88\xa8\xa8(\x18\x8dF\x08$ \xfbL6\x0e\x1c8\x00\xa7\xcb\r\x06\x84.]:#66\x0e\xa7N\x9d\xc2\xe9\xac,huZp<\x0f\x9b\xd5\x06\x93\xd1\x88N\x9d:#,,\x0ct\x11\x95\xf9\xac\xac,\x1c9r\x18&\x93\x19F\x93\x11,\xc3\xc2\xe5r\xc9\x1a)\xf2(P2\xa3\xc10\x80\xcdf\xc7\x85\xbc<\x9c8~\x1c\x17\xf2\xf2\xb0s\xc7\x0e\x9c<y\x12\xb7\xb4m\x8b\x16-Z\xa0^B\x02\x9a6i\x8a\xa8\xdaQ0\x1a\x8c\x17\xcdC08\x8e\xc3\xea\xd5\xab\xe1v\xbb!\x08\x82bD*\x96\xa4\x88\xe0\xf1\xd1\xea\xb3\x01\xa9VT-$\xf6\xe8\x01\x96\xd5\xc0n\xb7c\xcf\x9e=8s\xe6\x8c8=$}\xf5h4Zt\xed\xd6\x15\xd1\xb5k\x07l \x04\xe0lv6\xb6o\xdf\x0eF\xd2\xfe`\x19\x16\x04BLl\x0c\xbat\xe9"-\x14\x96!\x08\x02J\xadV\xe4\xe7\xe5\xe1\xf0\xe1\xc3\xc8\xcd\xcd\xc1\x96\xad\xdbp\xeel6\xda\xb5o\x87\x0e\x1d:\xa2I\xd3&H\xa8S\x17aaa\xd0\xe9\xf5\x01\xcb\x87\x88\x90\x92\x92\x82\xfd\xfb\xf7K\x8b\xd3\x00\x03\x06&\x93\x11\x9d:wA\x8d\x1a5\x02\x9a?\x06\x80\x82\xfc\x02l\xdc\xb4\x11\x1c\xc7\x89\xb6\xa4Ht\xae\xde\xf1\xd6\x8e\xa8_\xaf>H\xf2m\x9b\x94\x9c\x84\xcc\x8cL\xe8t:\xb8\xddnX\xadV\x84\x87\x87\xa3s\x97.0\x9b\xccA\xef\x0f)\x7f\xa7O\x9f\xc6\xbf\xff\xfe+w\xc6</\x80\x91\xbc\xadu\xed\xd6\x15Z\x8d\x16D\x84\x83\x87\x0e\xe2T\xca)0\x8c\'\x0e\x0b\x83\xc1\x80A\x83\xee\xf4\xfb\x02\xf5\x86H\xc0\xee\xdd{\x90\x99\x99\t\x8dF#\xf5\x9b\xa2\xe9\xe7\xf6\xed\xdb\xfb\xd9\xb4\xf2 \xd6\x0fQ+\xc7n\xb7\xe3\xc2\x85\x0b\xd8\xb7o\x1f\x92\x92\x92\xb0u\xebVX,!\xe8\x91\xd8\x03\xed\xdb\xb5C\xb3f\xcd\x10\x11\x11\x01\xbd^\x0f\xa7\xd3\x89\xbd{\xf7\xe2\xec\xb9s`\xa4i\xa1[o\xbdU\xb1\x913//\x0f\x7f\xff\xfd74Z-4R\xbd\xf3 \x08\x02x\x9e\x87N\xa7S\nPA\x10\xe31@\xbd\x84zh\xd5\xaa\x15LF\xd3\xe55\x8c \x10\x11~\xff\xfdw?e\x13\x00\x181b\x04\xa6N\x9d\x8a\xbau\xfdMx\\+\xae\xbb\x00\xf8\xe0\x83\x0f\xf0\xee\xbb\xef\x81\xf3\xb1\xa5\xff\xec\xb3\xe3\xf1\xe9\xa7\xd3\x14\xae\x12\xaf\x15\x02\x11\x04i\xb4J\x92\x0e\xbd\x86\xd5\x04\xdc\x00\xe4;\xe2\xf1tH\xbe\x90dg\xbd\xa2\x04\xba\x07*0b\xf0\\\xe7\xf1x\xc6J~T\x19\xcf\x1aE\xf9\x97W\x88`y\xab(\xde\xcf@\xd2t\x8ax\xa8\xec\xf8\xc5\x9e\x13\x8a|\x94\x89\x1e\x82(\x0c\xca\x83$\x97\x8d\x1c\xc7\x01\x9e\xf7+ux\x15O\xd7\xd3\x9d\x95\xe5\xba\xa2\xd7\x92\x8f\x010\xdf\xeb\x82\x95\xafo\xbc\xf2\xf0\xbeG\xb0<*\xd2\xf1\x8aT\x91t\x82\xe5\x11\x15\xb8^|\xdf\xd2\xfb"R|\x19j4\x1a\xb9\xbe\xfa^\xe3\x8d\xefyx\x0f:\x02\xe1[\xe8\x01\x08z\xed\x15@\x17\x11\x00\x9f|\xf2\t\x12\x12\xea*\xea\xfe\xb5\xa4\xfc\x96r-` \xd5>%\xde\x95\xe4Z\xc32\x8c\xc2\xf1\xbbV\xab\r\xd8\xf9C\xaa4\xde!\xd01\xa6\x9c\x1d\xa0\xc1\xf0\xbd\xde\xfb\xfe\xe5\xe1\x89\xa7\xd3\xe9\xa0\xd7\xebeg\xda\x0cS9\x9d?\xca\xc9[E\x83\xef\xbd\xc4\xfcIB*@\x9c`\x94\xc5/\xbb\xf6b\x9d?\xa4\xeb4,\x0b\x83\xf7\xfb\xbd\xe4t\xc54\xd9\xcb\xc8\xb3\xe8W8\xf8u\xde\xe7\xca\x8bW\x1e\xde\xd7\x05\xcb\xa3\xe2\xfe^y\xaa\x08\xbey\xf3\x0e\x17\xc3;\x0e\xe3\xd5\xd6\xf4z=4\x01\x0c\xedy\xe2],\ry\xa0\x13(\x94w\xae\x9c{^Ud\x1bF\xd7!m\x89\x80\xad\xc5\xbf;\xbe\x8a\x04I\x8ce\xc5Oz\x95\x1b\x8dk\xffN\xaf}\x8a*U\r\xdf/\x88*A96\x8c.\x9fK{\xce\x80\x02\xa0\xb2\xb3\xe4OY&\x83\rL\xcf_\xc8\xc3\xe9\xac,\xb8\x9c\x01\x8cA\xa9Tc\x02\xbd\xed\xab\xcb\xb5O\xb1r\xa92\xf5\xbf\x8ad\xe3r\xa8\xfc\x8e\xf6\xe2\x10\x11\xacV\xab\xdf\x86>o*?_\x97v\xbf\x80\x02\xe0\xea\xa3\xccd\xa0z\xb5h\xe1B<\xf1\xc4\x93X\xb1r9\xce\x9f?\x0f\x04\x89\xa7\xa2r\xa3S\xf9\x9d\xc4eRE\xb2Q\x1d\xe0y\x1eiii\x98?\x7f>\xa6L\x9e\xec{Z\xe6z\xf7i\xd7I\x00x\xc3\x04]\x9b\xd9\xbau+^\x7f\xe3\rL\x9a4\t[\xb6l\x81\xcb\xe9\xac:\xa3!\x95*\x89Z?\xaa\x1b7\xde\xfb*--\xc5\xb2e\xcb\xf0\xdf\xff\xbe\x817\xdf|\x13\x99\x99J\xf3"J\xae\xef\xf3\x07\x17\x00\xd7"_$n\x8e\tf\x96\x16\x00\xd2\xd3\xd2\xf1\xcd7\xdf\xe0\x89\'\x9e\xc0\x97_~\x89\xcc\x8c\x0c/}\xdf\xc0\xd7]\xefN\xa0"\xa9W$Ny\\\xecz*\xc7IFy\x90\x97\xd9c\xe5\t\xdf\x03\x1e\x82\x9e\xa84\xc4\x14|\xd2\xf1\xfa\xe9}&\xd0h\xd9\xb7\x1c\x02\x99\xb4\xf6\x8dS\x11\xe8\n\xca9\x18\xdew\xf2\xec\x1d\xb8d.\xe3\x12oH\xde\xe8&\x95|\x90\x8do\x95\x83\xff\xfb\xaa\xae\xd8\xecv\x1c8x\x10\xaf\xbc\xf2\n\xc6\x8f\x1f\x8f%K\x96\xca\xd6N\x03\xc1\xc0S\xc4\x15(\xdb\nD\xb9\x1c\xae\xbb\x1a\xe8\x9e={\xf0\xc9\xd4O\x02:C\t\xc4=C\x86`\xe4C\x0f\xa1w\xef\xde\x88\xaa]\xbb\xcaU\x1f\xf2\xf8\x9f\xe58Q\xaf\x9b\x17u\x8d\xb5\xac\x06\xc4\x88*\x8a:\x9d\xe4Q\x8c\xb9\xf4\x17\xebr\xb9`\xb3\xd9\xc00^[\xf9\x19\x06!\x96\x10E<A\x10P\\\\,\xedY\x107\xb9\xf8v\xec\xe4\xe90I\x00$#\x9d$i_Y,\x16\x85\xe6\x92\xdb\xed\x82\xc3\xe1\x94:XQk\x82\xd50\xd0\xebD\xcd\x8d\xca@\x10\x04\xd8\xacV\x08 0\x10\xb5\x83X\x96\x05\xab\xd1@\x1fD/\x1fRG\xc9s\xbc\\\xee\x00\x03\xadN\x1b\xd0\xe1\x8f\x07\x97K|\x1eH\x9d\xac@\x02\xc2\xc3\xc2\xcb\xd5\xd6\x12\x04\x01.\xb7[,9b\xe0t:\xc00,L&\x13t:\x7f\xc7\xe7\x1e\x04A\xdc8&v\xa2\xe2K\xd7h40\xf9\x98u "\x94Z\xadR\x87+\xaa \xb3\x8c\x06\x16\x8b9h\x19\xf3<\x0f\x9b\xcd\x06\x92\xd4)\x19\x06\xd0\xe9\xf50\x9b\xcc\xa2\x8bS\xa7Si\xec\x8f\x00b\x08\x0c\x89~\xa4=\xdam\x0c+\x1aY\xf3\x0c\x004\x1a\rX\r\x0b\x06\x0c\x9cN\'\xdc\x9c[|\xefR>!\xfd\xf2V\xd4`\x18\x06\x82 @\xa7\xd3)<y\xf1<\x0f\x87\xd3!Z\xde$\xd182IZ\x99Z\x9d\x0e\x06\x83^\xd4\xde\xf2\xa8P_G\r\xc0\xcb\x81\xe78$\xa7\xa4`\xfd\xfa\xf5\xf8\xfe\xfb\xefp\xe0\xc0A\xdf(\x01\x99<y2\xc6\x8f\x1f\x8f\xb0\xf0\xb0K\xda\xc9_\x99\\S\x01\xe0\xe9\\\xbc\xe1y\x1ei\xa9i\x98\xf7\xf3<\xfc4\xf7\'\x9c\xce:}\xd1\xd1F\x9d:u0h\xd0]\x18\xf5\xc8#\xe8\xd4\xb9\xd3U\xdd,\xe6i\xb2\x15\x85\xe3y\xec\xfd\xf7_,_\xbe\x1c\xbc\xc0\x83\xe7D;\xf2\x9e\xad\xf9,\xcb\xc2b\xb1\xa0u\xeb\xd6h\xdc\xb8\x11\x1a7n\x12p\xafC\xa0t\x89\x04l\xdc\xb8\tK\x7f\xfb\rz\x9dN\xde\x10d4\x9a\xf0\x9f\xff\xfcGa\xcaB\x10\x04\xfc\xf8\xe3\x8fHJJ\x92\xf73(\x04\x80\xd4\xe13\x8c\xf7\x08\x96\x01\x91\x80\x0e\x1d;\xe2\xde{\xef\x85\xc5l\x92\x8e\x11\xf6\xed\xdb\x87y\xf3\xe6\xc9jz\x0c\xcbB\xab\xd5\xe0\xe1\x87\x1fF\xebV\xad\xcb\xee\x1b\x90@O\xa3\x84\x88p&;\x1bs\x7f\xfa\t\xb9\xe7\xcf\x8b\x02FRg\xbd\xeb\xae\xbb\xd0\xa3G\x8f\xa0\x9d3\xcf\xf3\xf8w\xef^,_\xb6L\xfe:\xd4\xeb\xf5\xe8\xd3\xa77n\xbf}\x80_\x9d\x13\x04\x01\x87\x0f\x1f\xc6\xc2\x85\x8b`\xb3\x89\x1d\xae\xc5b\xc6\xa4I\xef\x96+4\x0e\x1d:\x84\xc5\x8b\x17\x8b\xbe\x01\x00\xd8mv\xd4\xa9[\x07O>\xf9$"##}\xa3\x03\xd2s\x9d;w\x0e?\xfd\xf4\x13rss%\xd3\x16\x84n\xdd\xba\xe3\xde{\x87)\x9e\x89\xe7y\xfc\xf0\xc3\x0f8|\xe4\xb0\xdc\xd9\x86\x87\x87\xe3\xf9\xe7\x9f\x97M/\xf8\x96\xe4\xf1\x13\'\xf0\xd5\x9c9\xb2\x10$\x120r\xe4\xc3\xe8\xda\xad\x1b\xce\x9d=\x8b\xb9s\xe7!\xef\xc2y\xf1EK\x1d/\xa4\xf7\xaea\xc5N^\xfc-\x965\x11\xa1F\x8d\x1ax\xf0\xa1\x07\x11\x13\x1d\x03\x81\x04\xec\xd9\xbd\x07K\x97.\x05\xef\xd9\xb1\x1c\xb0}\x8aS\xb9$\x10\xda\xb6k+\xd6\x9f\x10\x0bX\x86\xc5\xa1C\x071\x7f\xfe/\xf2\xe6A\xf9\n\x86\x81^\xaf\x87\xd1hDXX(\xba\xf7\xe8\x81\x96-Z"44\xd4\xef\x9d\x05"P\x7fr-!"\x14\x17\x15c\xe5\xaa\x95X\xb2d\t\xfe\xf8\xe3\x0f\xbfA\x96/z\xbd\x1e\xbd{\xf7\xc6\x8b/\xbe\x88\xae]\xbb\xfa\xd9\xb4\xba\xe6\xf8Z\x87\xbb\x1e\x81\xe38*))\xa1\xb5k\xd7\xd2\x13O>I\xf5\xea\x05\xb6r\xe9\x1d\x18\x86\xa1\x86\r\x1b\xd2\xd4\xa9S)%%\x85\x1c\x8e\xabk9\xb4\xa2\x81\xe3yZ\xbbv-\x85\x84\x84\x10\xcb\xb2\x01\x83V\xab\xa5\x1a5jP\x93&Mh\xdc\xb8q\xb4\xff\xc0\x01\xb2Yme\xf7\xe1\xfc\xef\xcb\xf3<\xb9\\.\x9a>c:\x01P\xdc/66\x966m\xdeL\x9cw>8\x8e>\xf8\xe0\x032\x18\x0ceq\x19F\xfa+\x06F:\xce0\x8c\x14X\x02@\xef\xbe\xf7.\x15\x14\x16\xc8\xf7r\xbb\xdd\xf4\xed\xb7\xdf\xfa\xa5\x1b\x12\x12B\xb3g\xcf"\xb7\xdb\xe5\x97\xd7\xcb\t;v\xec\xa4\x8e\x1d;\xca\xf9aY1?\x1f\x7f\xfc1\x95\x14\x07\xb7Z\xcaq\x1c\xad\xf1)s\x8dFC\x8f?\xf1\x04\x9d9s\x86\xdc^\xd6<y\x9e\'7\xe7\xa6\x7f\xfe\xf9\x87\xda\xb4i#[\x15\xed\xd3\xa7OP+\x98\x9e4\xbe\xf9\xe6\x1b\n\x0b\x0b\x93\xd3`\x18\x86\x8c\x06#\x9dL:\xe9\x17\xbf,p\x94\x96\x96F\xf7\xddw\x9f\xfcLaaa\xf4\xfd\xf7\xdf\xfb\xc5\xe58\x8e\xf6\xee\xdd+[y\xf5\xe4\xed\xcf\xd5\x7fz=C\xd9\xb38\x9d\x0e\x9a6m\x9a"\xee\x1dw\xdcAg\xce\x9c!\x8e\xe7(5-\x95F\x8c\x18\xa1xg\x15\t\xc3\x87\x0f\xa7\x94S)r\x9e\xd6\xae]K\x16\x8b\xc5/^\xa0\x00\xc92\xae\xc7\xea\xad \x08\xb4m\xfb6\x8a\x8d\x8d%VQ\xd7\xca\xde1\xcb\xb2\xa4\xd7\xeb\xa9n\xdd\xba\xf4\xfc\x0b/\xd0\x89\x93\'\xc8\xedVZ\xfb\xac\x12Az\x07\x1c\xc7QQa\x11\xed\xd8\xb9\x83F=\xfa(Y\xccf\xbf\xfe)P\xe8\xd1\xa3\x07}\xf8\xd1G\x94\x9a\x9aJ\xaeJj3W\x1a\x02\x0f\xa9\xae1\x0c\xc3\xc0l6\xa3_\xbf~x\xff\xfd\xf7\xf0\xe9\xa7\x9f\xa2o\xbf~\xbe\xd1\x14\x10\x11RSS\xf1\xdak\xaf\xe1\xc5\x17\xff\x83U\xabV\xe1\xdc\xb9sbQ_G\x18iJ\x81\xe7\xcb\x9cN\xf8\x06\x8e\xe3\x90\x9f\x9f\x8f\xe4\xe4d\xcc\x9e=\x1b\xe3\x9ey\x06[\xb6n)3\x87\x11\xc8`\x10\x89\xde\x9d\xfe\x92l\xc4x\xdf\xaf\xa0\xa0\x00\xfb\xf6\xfe\x0b\xf2\xf1\rLD\xf2\xa8K\x10\x04q\x87\xb3 @ 1\x90\xc7)\x86\xec,C\x1c\xbd\xb0\x8c\xb8\xf3\xd9\x83\xdb\xed\xc6\x96-[\xfc\xd2\xb5Z\xad8q\xe2\x04\xac6\x9b\x1c\xf7r\xe18\x0e\x99\x99\x19\xd8\xbf\x7f\xbf\x9c\x1fA\x1aM\x1d<t\x10v\xc7E\\W\x92\xb2\x8cE?\x01\x9bp\xf0\xe0A?\xfb<\x0c\x89\xa3X\xcf\xb3C*\xab\xf2())\xc1\xc1\x83\x07QRR"\xa7ADp8\x1d8v\xecx\xd0\xeb=G\x19iz\xc3;ME\x04)N\xa3F\x8d\xf0\xd8\xe8\xd1\xe2))\xde\xbcy\xf3PTT$M\xb7xF\xbc\x84\xec\xecl\xfc\xf4\xd3O\x8a\xb8#G>\x84\xc8\xc8\x08\x88\xfb\xbe\xcb\xd2\xbc\x94\xc00\xa2M*1OR\xefU\xc1\xfb@\xaa#\xde\xa5\xc12\xa2\xb9\x08\xc1\xc7a\x90\xf7=].\x17\xb2\xb2\xb20\xfd\xcb/1q\xe2$\x9c8q\xe2\xbaM\x8b\x04\x85a\xc0\x0b\x02N\x9c8\x81\xd9sfc\xe4\x83#1o\xee\xdc\x8b\xd6\x7f\x93\xd1\x88W_{\r\x9fL\x9d\x8a\xe7\x9f\x7f\x1e\t\xf5\x12.\xd9\x10\xdf\xd5\xa2J\x08\x00\x0f\x0c\x03\xc4D\xc7`\xf0\xe0{\xf0\xfdw\xdf\xe1\xfd\xc9\x93Q\xaf^\xbd\xa0\x9f\xfe\x1eV\xad\xfa\x03\xcf\x8c\x1d\x8b\xc9S&c\xdb\xf6m\xb0\x97\xe3\xf7\xf5\xea\xe3\xd9YXv\xc4d2\xa1W\xaf^\xe8\xdb\xa7\x0f\xbaw\xef\x8e\x84\x84\x04\x18\xbd\xcc0\xef\xdc\xb9\x13\x93&M\xc2\xe6\xcd\x9b\xa5O\xf4\x00\x15\x9f\x01222\x90\x9c\x9c\xec{\x06N\xa7\x13G\x8f\x1c\xc5\xf9</k\x8f\x0c\x03\xa3QL\xb7w\xef\xderh\xdd\xba\xb5\xe2\xb3\xb9i\x93\xa6\x8a\xf3\xbdz\xf5\x82\xd1d\x94\xa7\n\x88\x08\xa7\xcf\x9c\xc1\xbe}\xfb\xcb\xee-AD\xc8\xcc\xccBn\xae\xa8\xa6{%\xb8\\.\xa4\xa5\xa5\xc9\x9d\x887\xbf\xcc\xff\x05999\xe5\xca\xf6@n\xf7RO\xa5b\xe5\xca\x95(\xf61\xcb\ri\xda+X\xa7\x1d\x88\xac \xee\t\x01`\xd5\xca\x95\x01\x8d\xaf!\x88,\xf7\x10\xe8U\x9b-\x16\xf4\xe9\xd3\x1b-[\x96y3\xdb\xb8q#v\xef\xda%/8\x8b\x82\x1c\xd8\xba\xf5\x1fddd\xc8\xf1\xee\x188\x10\x1d\xdaw\x80^\xaf\x17E\x05\xc3\xa0f\xcd\x9a\x8a\xf7\xeb\x1b\xbau\xeb\x86\x9a5k\xca\xf7`\x18\x06\xa1\xa1\xa1^f\xcb\xfd\xbba\xa3\xd1\x84\x9e={\xfa\xddK\xae?~&\xc6I\xf1\x9c\xa1aa\xe8\xdc\xb93\xfa\xf4\xe9\x83\xde\xbd{\xa3\xcd-m\x14\xd7,^\xb4\x08?\xfd\xf4\x93\xe8\x05\xaf\x8a H^\xca\xe6\xcd\x9d\x8b\xd7_\x7f\x1do\xbc\xf1\x06\xd2\xd2\xd3|\xa3)0\x99L\xe8\xdb\xa7/\x16.Z\x84\xff\xbe\xf1::w\xea\x04\x93\xd1\xe8\xff\xd2\xaf\'\xbe\x9f\x04U)\xd8l6\xda\xbau+\x8d\x193\x86"##Ij3A\x03\xcb\xb2T\xb7n]z\xf7\xbdw)%%\x85\xacV\xab\xe8\xccC\x9eRQN\x05Tv\xf088Y\xb3f\r\x99\xbd>\x0b{\xf5\xeaE%%%\xe4t:\xa9\xb0\xb0\x90v\xee\xdcE\x9f~\xfa)\xb5o\xdf^\x8e\xc30\x0c=\xf1\xc4\x13\x94\x95\x99Inw\xe0|\xfe\xf9\xe7\x9fT\xabV-\xbf\xe7\x06@\xb7\xdf~;\x1d:t\x888\xce\xf3\xe9\xcc\x91\xc3\xe9 \x9b\xcdFv\xbb\x9d\xecv;Y\xadV\xfa\xe3\xcf?\xc9\xec\xe5\xccd\xde\xbcy\x8a8v\xbb\x9d\x1c\x0e\x87\xec\x04\x85\xe38Z\xb5jU\xd0t\x1b6lH\xeb\xd7\xad\xf3\x9bf\xb9\xd4p\xee\xdc9\x1a2t\x88x_\xc90\xa6w\xf8\xf9\xe7\x9f\x83N\x0bp\x9c\x7f\x99{\x875k\xd7\x92\xdb\xa5\xfc\xe4\xde\xb2e\x0b5k\xd6L\x8eS\x9e#\x14\xb7\xdbM\xab\xfe\xf8\x83\xc2BC\xfd\xee\r\x80:t\xe8@\x99\x99\x99\xd2\x14\x9c\x7f9\xa4\xa5\xa5\xd1\x88\x11#\xe4\xf8\xa1\xa1\xa1\xf4\xfd\x0f\xfeS@<\xc7\x93\xcb\xed\xa63\xd9g\xe8\xf1\xc7\x1f\x97\xa7uX\x96\xa5\x17_z\x89\n\n\n\xc4:\xc6\xf1\x94}\xe6\x0c=\xf9\xe4\x93r\x1c\x9dNG\x13\'M\xa4\xd2\xd2R\xf9~\x1c\xc7\x91\xc3\xe1P\xbc[\xef`\xb3Y\xe9\xc0\xc1\x03t\xcf=\xf7\xc8y\xabY\xb3\x06\xcd\x9b7O~\xff</N\xaf)\xeas\xcf\x9eTXX(\xdd\xc7&\xd7\x1f\xcf_\xef\xfa#\x08\x02\xed\xd8\xb1\x93bbb\xe5\xeb\xef\x192\x84\x8e\x9f8AN\xa7\x93\xec\x0e\x07\xa5\xa6\xa5\xd1\x94\x0f\xa6(\xdax\x87\x0e\x1dh\xdb?\xdb\xc8\xe9*s\x08t\xcd\x03\'\xbe\xfb\xc2\xc2B\xda\xbcy3=\xf8\xd0C\x15\xea\x87\x00P\xbf\xfe\xfd\xe8\x7f_\xfc\x8f\xb2\xb3\xb3\x83\xd6\xdb\xaa\x10\xca\x1fZ_g\x0c\x06\x03\xbav\xeb\x86\xb7\xdf~\x1b_}\xfd5\x06\r\x1a\xe4\x1bE\x81 \x08\xc8\xca\xca\xc2\xc4w&\xe2\xc9\'\x9f\xc0\xca\x95+q.\'\xc7K\xe0^]\xc9\xcbH\xdf\xcb\xbe\xa3K\x96ee[\'\xe1\xe1\xe1\xe8\xd2\xa53\x9e{\xee9\xbc\xf7\xde{\xb2C\x17"\xc2\xa6M\x9bp\xe0\xe0!\xbf)\x0bH\x0b\x84\xbbw\xef\x96\xd5\xca\xe2\xe2\xe3\xe5\xa9\x02\x00\xd8\xbcy3N\x9dJ\x85\x9b\xf3\\\xcb@\xa7\x15\xb51<i\xeb\xf5z\xe8|\xb4I|\xcf\xeb\xf5\xfa2-%\xc9]\xe3\xce\x9d;E\x87\xf3\xd2"\xd6\xb0{\xef\x15\xad&\x02HMME\xd6\xe9\xd3Wf\x8aW2\x8b\xbc|\xd9r\x00\x00\xcb\xb0\x184h\x10\xda\xb5m+\x9f_\xfb\xf7\xda\xa0_vt\x91\xc5\xc0E\x0b\x17\xa2\xa4T\xf4\xa5\xe0i\xa1\x9e\x7f+\x82\xd3\xe9\xc4\xa9\x94\x14\x14\x97\x94\x00R\x99\r\x1e|\x8f\xbcx\x9f\x91\x91\x81\xc3\x87\x0f\x03D\x01\xbfD\x02A\xb2\x0f\x00/\x18\xb1\xaeD\xd7\x8e\xc6\xc0\x81\x03es\xd8\x82 `\xd7\xce\x9d8q\xe2\x844\x8d"\xe0\xf8\x89\x13\xd8\xb1}\xbb\\\xcfbbb0`\xc0\x00\xd1\xc1\x8ft\xcc\xb3\xb0\xeb\xfb~\xf5z=tz=\\.\x0e\x1b\xd6oP\x98\xe3\xbe\xfd\xf6\x01\x188p\xa0\\\x9eD\xa2\x89j\xef\xfa\xcch4^\xf5\xc6 \xff\xdf\xf3\xd7S\x7f\xca\x9eQ\xf9\x05\xc0\x00\xb2\xfd\x1f\xa3\xc1\x80\x06\xf5\xeb\xe3\x99g\x9eA\x87\xf6e>\x1e\xf6\xed\xdb\x87\x93I\'\xe1r^Y\xbd\xba\\\x18\x00n\xce\x8d#G\x8e`\xfa\xf4\xe9\xb8\xfb\xee\xbb\xb1p\xc1\x02\xb9\x1d\x94\xc7\xbb\xef\xbe\x87\xa9S\xa7\xe2\xe9\xa7\xc7 \xaav\xed\x8b\xce`\\O\xaan\xce$X\x86A||<\xee\x19<\x18_~\xf9\x05>\xfe\xf8c\xd4\xabW\xaf\xdc\x06\x0f\x00\x9b7o\xc1k\xaf\xbd\x86\x89\x13\xde\xc1\xbau\xeb\xbd\xd4\xf0\xae>\x15I\xc7`0\xa0\x7f\xff\xfex\xe4\xd1G\xe4c\xa9\xa9\xa98p`?\x9cN\xff\x8e.//\x0f\xe9^^\xa2\x1a6h\x80\xe7\x9f{N\xd6\x80r\xb9\\8q\xf28\x9c\xe5t\xc4,+\xaa\xda)w\xde\x95_\x8e\x17\xf2. ##CN7::\x1a\xf7\r\x1f\x8e\x84z\tr\x9cS\xa9\xa9W$\x00\x04A\xc0\xb1c\xc7\x00)7f\xb3\x05O?\xfd\x14\x06\x0c\x1c(7\x9e\xed\xdb\xb6##-Mz\xef\xbe\xe5+\xaa\x1f\x06c\xe5\xcaU\xd8\xb8i# u\xcf\xe2\x13\x97\xff\xdc\xde\xd8\xedv\xec\xdf_6\x05\xd6\xb8qc\x0c\x1f~\xaf<uR\\\\,\xaf]\x04\xc2\xb7\xae\x92\xf7X\xd1\x07\xcf\xd3\xf5\xe9\xdb\x17\x1d:t\x90\x8f\xef\xd8\xb1\x03;w\xee\x84\xd3\xe9\x84\xcdn\xc3\xae]\xbbp\xec\xf8q\xf9\xfc\xf0\xfb\x86\xe3\x966\xb7\x88*\x9f\x17i\x1b\x80\xa8\xads\xe2\xc4q\xcc\x993G>\x16\x12\x12\x821c\xc7\xf8h4\x89\xc2U\xf1\x0c\x15\x9c>c\x18\xc9\x02m\xa0\xd2&\xf9\x1f\x00@xX8\x86\x8f\x18\xee\x1d\x03g\xcf\x9d-w\x8f\xd0\xd5\x82\x88\x90}\xee,~\xf8\xf1\x07\xbc\xf6\xdak\x980a\x02J$\xe1\x1f\x8c\x90\xd0\x10\xdcs\xcf=\xf8{\xdd:\xbc\xf0\xc2\xf3h\xdb\xa6-\x8cz\x83\xffsW1\xaa\xbc\x00\xf0\xa0\xd3\xe9\xd0\xa0AC\xbc\xf0\xc2\x0b\xf8\xe6\xdbo\xf0\xe4\x93O >\xce\xdf\x1d\xa27YYY\xf8\xfe\x87\xef\xf1\xf4\xd3Oa\xdag\x9f!--\rN\xd7\xd5\xb5-\xc40\xa2I[\xdfF\x1f\x08\xa3\xd1\x88\xdbo\xbb]q\xec\xcf?W\xfb\xb8\xee\x13\xc9/\xc8\xc7\xb9se\x0e7z\xf4\xe8\x81\xe8\x98\x18\xf4\xec\xd5K\x8e\xb3i\xe3&\x94\x96SQ\x89\x08$(\x1b\xaf \xf8\x7fmx "\xe4\xe6\xe6\xe2\xf4\xe9\xd3\xf2\xb16m\xda\xa0o\xdf\xbex\xf0\xc1\x07\xe5c\xbf-]\x8a\x92\xe2\xe0\xe9\x96\x07I\x8b\xcck\xd6\x8a.7\t@\xbbvmQ\xaf~}\xc4\xc7\xc7\xcb\xa3\xec\xbc\xbc<\x9cL:)\xe5\xdd\xb7l\xfdM\xfdz\xabs^\xb8p\x1e\x7f\xfe\xf1\'rr\xcb\xd6\x11X\x96\x85N[1\xf5\xe1\x82\x82\x02\xfc\xfc\xf3\xcf\xf2\xef\xfb\xef\xbf\x1f\xfd\xfb\xf7G\xe3\xc6M\x00)\xff\xc7\x8e\x1d\xc3\xd9\xecl\x10\xcf+:v\xcf\xbc\xbd\xb2\xce\xf9\xfeV\xc22\x0cj\xd6\xa8\x81\xe7\x9e{Nq|\xf5\xea?\x91\x9f\x9f\x87sgE\xb5R\x0f\xa1\xa1\xa1x\xe0\xfe\x07\xfc\xf6\x15\x04DJ6\xbf \x0f3f\xceDRR\x92|\xea\x8d7\xde@\x87v\x1d\x14#V"q\xfe\xdb7\xbf\x0c#\xee\x01\x10G\xf9R\xc4 \xf86\x05\x86a\xa5\xfd\x07e\'X\x86E\xcd\x1aQ\x8ax\x9c\x9b\x0b\xb8i\xefj@\x92\xdf\x84\xa2\xa2"l\xdc\xb0\x11\xe3\xc7\x8d\xc7\xff\xbd\xf6\x7fX\xb7n\x9doT?\xfa\xf5\xeb\x8b\xa9\x9fL\xc5\xf4\xe9\xd3\xd1\xb7Oo\x84\x85\x85\x8a\xcfwU\x8c\xbdU\x02^\xaf\xaa\xda\x08\x00\x0fz\xbd\x1e}z\xf7\xc1{\xef\xbd\x87i\x9fM\xc3=C\xee\xf1\x8d\xa2\x80\x88\x90\x99\x99\x89\x89\xef\xbc\x83\x91\x0f?\x8c\xa5K\x96"\xebtV\xb9#F@YHW\x0b\x86\x01\x9a6m\x06\x9d\xaelZf\xc7\x8e\xedp8\x9d\x8ax\x9c\xb4Wb\xab\xa4\x89\x03\x00\x9d;wFxX\x18n\xeb\xdf_>\xb6f\xcd\x1aded\xf85V\x0f$v=\xcacDA\xdb\xae \x08HKK\xc3?\xff\xfc#\x1f\xbbs\xd0 \x84\x87\x87#!\xa1\xec\x0b\xe0\xe8\xd1\xa3\xc8\xc8Pz+\xab0\x92\x909\xb0\xff\x80|\xa8A\xc3\x86\x88\x8d\x89E\xc7\x8e\x1d\xe5\xd1hQQ\x11v\xee\xdc)z\xe2\n\x90\x8eo3{\xf6\xd9g1t\xe8P@z\xc6\x8d\x9b6a\xdf\xbe\xfd\xe0yq\x83\x9eV\xab\x81\xc1\x18\\\xe7\xdf\x03\x11\xe1\xe4\xc9$\x85\xc3\xa2\x06\r\x1a\xa0F\x8d\x1a\xb8\xe3\x8e;\xe4c)))8\x93}Ft+\xe8\x95\x19&\x80\x7f\\H\xf7\xf5\xcb\xb4\x17,\xcb\xa2[\xb7n\x189r\xa4|l\xed\xda\xbf\xb1k\xd7nl\xd9\xb2\x05\'O\x9e\x94\x8f\x8f~l4\x1a7n\\\xb1\xce\x86\x11\x9d\xb1\xacZ\xb9\nK\x97,\x91\x0f\'&\xf6\xc4\x1dw\xdc\x01K@\x17\x98\xfem\xc5\xf3\xae\xe5\xa9"\xb9.\xf9\xc7\xf5{P\x06\x001\x8a\xd7(\x90\x80\xf4\xf4t\xefX\xd0\xeb\r\xd7L[\xc6n\xb3\xe1\xc0\x81\xfd\xf8\xf4\xd3i\xb8}\xc0\xedX\xb1b\xc5EG\xfdF\xa3\x11\x93&M\xc2\xd4O?\xc5\xe8\xc7G\xa3N\x9d:\xd2>\x8f\n\xbc\x87\xeb\x89W\xf6\xaa\x9d\x00\x80\xe44"::\x06C\x86\x0e\xc1\xf4/\xa7c\xda\xb4ih\xd8\xa8\xe1E\xe7\xda\xf6\xec\xde\x8d\x17_\xfc\x0f\xdez\xf3m\xfc\xbd\xeeoXmVP\xb0O\xcc\xcb~\x87\x81/\x0c\x9c\n\x03q/\x972\xdfi>\xce\xcd]N\'N\x9eLR\xa8\x9b5o\xde\x02&\xa3\t-Z\xb4\x90G\xc9\x00\xb0w\xdf\xbe\xa0\xc2\x8d\xf1j\xb8\x1e\xc4\xaf\x15\xc5!\x19\x87\xd3\x81\x94\x94\x14\x85\x86K\xf3f\xcd`0\x1aQ\xaf^=DGG\xcb\xc77n\xdc\x18p\xed\xa2"\x9c<y\x12g\xcf\x9e\x05\x00\xb0\x1a\r\xea%$\xa0V\xadZh\xd9\xb2%\x1a4h\x00H\xc2(==\x03\xb9\xb9\xb9 \x9f\x0c\xfbMQ\x00\xb8\xb5S\'\x8c\x1f?^\xfe\x9dz\xea\x14V\xacX\x81\xa2\xc2"0\xd2\x1c\xb4\xbf\xb6\x8a?<\xcfc\xe3\x86\x8d\xf2\xef\xda\xb5k#>>\x1eF\xa3\x11-[\xb6\x90\x8f\xff\xfb\xef\xbfHN9\x05\xb7\xdb\xdfW\xad\xc7\x1e\x7f\x19\xe2\xc8\xd9_\xbfFI\xcd\x1a5q\xef\xb0{\x11\x11\x11.\x1f\x9b3g\x0e\xbe\xfd\xee[\xf9w\xadZ\xb50`\xe0\x00\x84\x87\x97\xc5)\x0f\x81\xc4)\xbb%K\x96\xc8n--\x16\x0b\x1ex\xe0\x01\xb4l\xd9\xca\xaf\r1\x92\xef\x04o233\xf1\xd5\x9c\xaf0c\xfaL\xcc\x981\x033g\xcc\xc0\x8c\x19\xe2\xff\x97.\xfd\r\x02\t>B\xda\xbf\xf6Kk\xd7\xf2\xef\xbc\x0byX\xfd\xe7\x9f\xf2\xef\x90\x90\x10\xb4h\xd1\xbcBB\xfaJ\x10\x04\x01\x99\x99Y\x98\xf3\xd5Wx\xf5\xb5\xd7\xf0\xc1\x07S|\xa3\xf8a6\x9b1r\xe4H\xacY\xb3\x06/\xbc\xf0<n\xb9\xa5-\xf4:\xff\xcd\x9c\xd5\x81\xf2{\xcc*\x8e^\xa7G\x9d:u0~\xfcx\xfc\xf0\xfd\x0f\x183v\x8c\xc2Oo \xf2\xf2\xf2\xb0`\xc1/x\xf2\x89\'\xf1\xee\xa4wq\xfc\xc4qy}\xc0\xbf\x9a^:\x9e\xcfb_\x825u\x92\xd4F\xbd),,T\xdc\xc3&\x8dN<\xf4\xee\xd5\x1b5kFB\xab\xd5"66\x16\xad\xbcv\xe2.]\xba\x04v{p\xbdd\xd6\xc7\xdb\x92o\xda\xde\xd8mv\xec\xdc\xb9S\xfe]\xaf^=\xb4h\xd1\x02\x1a\x96E\xa3\xc6\x8d\x91\x98X\xe6?\xf9\xef\xbf%\x81*\x1f\xa9\x18\x04`\xf7\xae\xdd\xc8\xcf\xcf\x07\x00\x84\x86\x84\xa0U\xabV\xd0j\xb5\x08\r\tA\xd7\xaee~\x8ew\xed\xda\x85\xd4\xb4TP\x00\x01\xe7[\xe6:\xad\x16\xed\xdb\xb7\xc7\xf3\xcf?/\x1f\xfb\xfa\xab\xaf\xb1\xcf# \x19\x16\x1a\x1f\xc7\xe9\xbe\x10\x11JJ\x8a\xb1c\xe7v\xf9X\xcf\x9e=\xd1\xb0\xa18\xd8h\xd1\xaaLUS\x10\x04\x9c<q\x1c6\x9b\xff\xf4\x9d\xaf@&o\x9dx\xc5\x19%Z\x9d\x16\xb7v\xba\x15w\xdcq\xa7|l\xc3\x86\r\xd8\xb5s\x17 \xbd\xbb\x81w\x0cD\xdb\xb6m+<R.-)\xc1\xb2\xdf\x7f\xc7\xea\xd5\xab\xe5c\xbdz\xf5\xc6\xe0\xc1w\x07\xdd\x05\xcd@Yg\xd2\xd2\xd2\xf0\xdf7\xdf\xc4\x1bo\xbc\x8e\xd7_\x17\xc3\x1bo\xbc\x8e\xb7\xde|\x13\xc7\x8e\x1d\x93LD(\xef\xe0\x8d\xc0\xf3\xe0\xa5\xbd\x1a6\x9b\x15\x17.\xe4a\xf9\xf2\xe5\xd8\x7f\xa0\xec+\xb0\x7f\xff\xfeh\xde\xbc9X\x96-\xb7\x8c.\x17A\x10\x90\x9f\x97\x87U\xabV\xe1\xd1GG\xe1\xed\xb7\xde\xc2\x96\xcd\x9b}\xa3\xf91p\xe0@|\xf4\xd1G\x98:u*\xbaw\xef\x8e\xf0\xf0\x08\xa5\xa9\x8djF\xb5\x16\x00\x1e\xf4z=\x12\x13{`\xd2\xa4I\xf8\xdf\xff>\xaf\xd0\xb4\xd0\xd9\xb3g1m\xda4<\xf5\xd4SX\xb8p\xa1\xa8O\x1d\xa0\xe3\xbed.\xf1\x16\xfe\x8dE\xac\x9c\xde\r.\'\'\x07\xf3\xe6\xcd\x93\x7f\xdfv\xfbm\xb0XB\x00\x06HH\xa8\x876\xb7\xb4\x96\xdbXff\x16R\xd3\xd2\xfd\x1a\x1d ~\xca\xbb\xddn\x08\x15\x18\xa9\x13\x08yyy\xf8\xed\xb72\x1bM\x83\x07\x0f\x86\xd9b\x06\x00DFD ..V\xcegZZ:\x92N&\x95+P\x02Q\\T\x88\xec\xb3\xd9\xf2\xd7C\xadZ\xb5\xd0\xba\xb5(\xd04\x1a\rz\xf7)[\xe3\xc8\xca\xcaBfFf@m \xdf2$\x10\xc2\xc2\xc20d\xf0=h\xd5\xaa\x95|\xfc\xc7\x1f\x7f\xc4\x85\x0b\x17D\xcd\x96\x00\x82\xc4\x1b\x86ep\xea\x94\xa8\xe5\x04\xa9\xc3\x8d\x8b\x8bC\x8d\x1a5@ \xd4\x88\x8c\xc4\x03\x0f< \xc7\xff\xe3\x8f?Q*i\x1by\xe3[$$\xd9\x1eB\x10_\x18\x1e\x88\x80\xd8\x98X$&&\x06\x9c\xdf\x0f\r\r\xc5\x9dw\xdc)\xae\x85\x95w#\t\x81\x08\xbbw\xef\xc1\xc4\x89\x13\x15\xc7_x\xe1\x05\xc4\xc7\xc7\xfbM\x0f\x02b\xc1\xca\x9a<^8\x1cv8\x1c\x0ee\xf0\x9e\xba\x94L\x8d\x04\xb8\x14\xe7\xce\x9d\xc5\xf6m\xdb\xb0r\xe5J\xcc\x9e=\x07\xff\xf7\x7f\xafa\xcc\xd81(*\x125\xdcjE\xd5\xc2\xa3\x8f=\x86\x06\r\x1aH\xdb\xd9*\x07\xcf \xc1j\xb5a\xe7\xae]\xf8\xe8\xe3\x8f1l\xd80l\xdd\xba\x15N\x9fiW_\xea\xd6\xad\x8b\xc9\x93\'\xe3\xe3\x8f?\xc6\x981c\x10\x13\x13\xe3\xf7^\xab#7\x84\x00\x80\xd4Xj\xd5\xac\x85\x81\x03\x06b\xc6\xf4\xe9\x98={6\x1a5j\x14\xd4\x88\x96\x87]\xbbv\xe1\x85\x17^\xc0\x84\t\x13\xb0v\xed\x1a\xd8\xed\xa2a\xad\xcb%\xd0tDy0\x0c\xc0\xfaM\xcbx/\xc2\tHMM\x95\x7f\xeb\xf5z\xb4n\xdd\x06F\xa9C\x88\x8a\xaa\x89z\te;\x0b\xcf\x9e=\x8b\x9d;v\x80\xe0\xdf\xb9\x91\xa4-\xe3\x9d\x9a\xef\xe8\xd4\x03\t\x84\xc3G\x8e(\x8eu\xea\xd4\t\x16\xb38G\x1c\x16\x16\x86\x84\x84\x04Y\x0b\xe9\xc2\x85\xf3\xd8\xb5k\xd7%\tQ\x02\x90\x9d}\x16)))r\x99\xc7\xc7\xc7\xcbN\xc5\x19\x86E\x8b\xe6-\x14]\xc0\xbe}\xfb\xe0\x0c \x00|\xcb\x9c!\xd1\x98\\\xdb\x0e\xed\xd1\xa7o_yj\xe3\xaf\xbf\xfe\xc2\x1f\x7f\xfe\t\x86\x05\xb4\xda\xf2\xeb\x06\t\xc0\xde\xbd{\x91/mH2\x18\x0ch\xdc\xb8\tB\xc3D\xe3]&\x93\t\xb7x\xa9\xaa\x1e8p\x00G\x8e\x1eS\xcc\x83\x93\xf4\x1c\xbexv\xea\x96\x07\xc3\x00\x1a\x9d\x06\x03\x07\x0c@\x8b\x16e\xd3M\x1eZ\xb6h\x81>\xbd{C\x9aG,\x17\x06\x0c\xb2\xcf\x9c\xc1\x8c\x19\xd3\x15\xc7\'L\x98\x80.]:+\x17~\xbd\xce\x93\xbci\xae\xec\x18\xcb\xb2\x88\x8d\x8dE\\\\,bcc\x11\x13\x1b\x83\x98\x98\x18DGGK_\x98\xa2\xf0d \x8dC|\xb2\xb7g\xcf\xbf\x18;v,\x1ex\xe0\x01\xbc\xfa\xea\xab\xf8\xe1\x87\x1f\x00I\xe0\xd7\xabW\x0fsf\xcf\xc1\xc0\x01\x03D\xfbY\x17\x7f\xb4\n"\x1a\xcbKII\xc1\x97_~\x81W^~\x05\xd3\xa6M\xf3\x8d\xe4\x878=v?~\xfe\xf9g<\xf7\xdcsh\xd5\xba5\xb4:-\x08\x14\xf0\xbdV7*\xf6\x04\x15o\xd3\xd7\rO\xc7\xab\xd5\xe9\x10\x1f_\x07\x8f\x8d\x1e\x8d\x9f\x7f\xfe\x19/\xbf\xfc\x12\x9a6m\xea\x1b]\x81\xcdf\xc3\x82\x05\x0b\xf0\xf8\xe3O\xe0\xc3\x0f?\xc2\x91#G`\xb5Z\xc1\x0b\xe2|\xae\xd89U\xac\x10\xc4\xc6\x12x\x1a(\x10D\xfe]uLl\x9c\xfc\x7f^\x10\x90\x92R\xb6\xfb\xb7A\x83\x06\xa8U\xab&XI\xcfZ\xab\xd5\xa1m\xdbv\x08\x0f\x17u\xc6\x1d\x0e\x07N\x9c8\x11P+\x87H\x80\x9bS\x1a\xe4\xf2\xf5\xa1\xeb\x19\x05r\x1c\x87\xad[\xb7\xca\xc7\xe3\xe2\xe2P\xb7n]\xe8t\xa2\xc5H\x86a\xd0\xa7w\x1f\xc4D\xc7\x00\x92\xae|J\xca)\x14\x17\x17\x97SV\xca\xe3$\xf0\xc8\xc8\xcc\x14\x05\x87\xc4\x80\x81\x03a4\x18\xa5\xe9\x11BDx\x04\x86\x8f\xb8O>\xff\xcb\xcf\xf3Q\\\xec\xb3\xb3\xd7\xf3\x8e\xbc;\x0bI\x1b22"\x02\xf7\x0e\x1b\x8a\xbau\xeb\x02\xd2b\xf2\x96\xcd[p>\xf7\x82\xf8\x15U\x0e\x0e\x87\x1d\xa9\xa9\xa9pH\x02\'::\x1a=zt\x03\xa4\xdd\xb8\x06\x83\x11M\x9b6UX\x02MIN\x86 (\xc7\xd2\xbe\xf5\x81\x88$\xeb\x9a\xc1\xca\xa9\x0c\x06\x0c\x12\xea\xd7\xc7\x93O>\xe9{\n\x8f=>\x1a\xb5\xa3\xa3+$t\x9dN\x07\x96\xafX\x81\r\x1b\xcb\xd63\xfa\xf7\xef\x8f\xa1\xc3\x86"4L,\x07O\xf1)\x8b\xd1\xff\xdem\xda\xb4\xc1\xa2\xc5\x8b\xb0t\xe9R,\xfeu1\x16/\xf6\x84E\x18>\xfc^\x9f\xec\x88\x1d\xaf\x7f{`\x14kK\x1a\x8d\x06\x0f=\xf4\x10~\xfe\xf9g\xdcy\xc7\x1d0\x99.\xbe>S\x11\xc4\xb2\xe6p!/\x1f\x7f\xfc\xf9\x07\x1e{\xec1\xbc\xf7\xde{\xd8\xbd\xbb\xac\xce\x05\x82a\x18\x0c\x180\x003f\xcc\xc0\x8c\x99\xb3\xd0\xbd{w\x84\x86\x86\x8a\xeb9\x15\x10\xde\xd5\x85\x8a\t\x80j\xf8\xac\x06\xbd\x1e\x9d;w\xc2\xeb\xaf\xbf\x81/\xbe\xfc\x12\x0f?\xfc\xb0o\x14\x05D\x84\x9c\x9c\x1cL\x992\x05O<\xfe\x04\x16,X\x80\xcc\xcc,\xb9\xb3\xf3\xad\xbe\x17\xc3wD\x1a\x8c\x82\xfc|EC`\x19\x06u\xe2\xe3\xe54\x8b\x8b\x8b\xf1\xeb\xafe\xda\x1a\x91\x91\x91HMM\xc5\xfa\xf5\xeb\xb0~\xfdz\xac_\xbf\x1e9\xb99\xf2K\x12\x04\x01)))\x92\xe9\x04\x9f\\KrLYy}\x9fL\xd4\xce\xc8\xcf\xcf\xc7\xb6m\xdb\xe4\xa3\xe1\x11\xe18q\xf2$\xfe^\'\xa6\xb9n\xdd:$\'\'C#\x8d\xa2\x89\x08\xc9\xc9I\xc8\xce>\xebw\xc72\x94e\xe2p:\x91\x9a\xea\x11\x1a".\xa7\x13\x1b7m\xc4\xbau\xeb\xb0n\xdd:\xec\xda\xbd\x0b\xa1^\x9a)\xf9\x85\x058v\xdc\xc7\xf6\x0eI&\x17\xbc\x0f\x11\xc9\xeb+]:w\xc1#\xa3F\xc9\xc7\xb7o\xdf\x8e\xd5\xabW\xe3\xc2\x05/\xd3\x19\x01\xc8\xc9\xc9AZZ\xba,05\x1a\rRRNa\xdd\xbauX\xbfn\x1d6m\xda\x84\xd4S\xa7\x14#\xc1\x9f\x7f\x9e\'\xafg\xc0\xb3-\x8c\x81b\x1eH\x10\x04\xb8\xdd\xaer\xcaI\x89V\xa3A\x7f/m/\x0f\xbdz\xf6\xba\xe8\x17.\xa4\xf4\xf6\xed\xdf\x8f_\xe6\xcf\x87U\x9a\xa2\xb2X,\xb8\xef\xbe\x11h\xd1\xbc\x85\xdf \xc0\x1b\x82G\x98\x96\xe5?22\x12\x9d;uF\xe7.]\xd1\xa3{\x0f$\xf6HDbb"\xbau\xef\x8e\x96-[\x96\xc5\x95\x1e\x90\xd5H\xfbO$:v\xec\x88\x193\xa7\xe3\xa5\x97_\x96\x15\x18\xb4Z-Z\xb7n\x8dv\xed\xdaJ\x0b\xbf\x15k?\xc1\xf0\xd4\x8f\x92R+\xb6n\xd9\x82w&L\xc0\xbd\xc3\xee\xc5\xce\x9d;/\xbag%..\x0e\x9f\x7f\xf69\xa6N\x9d\x8a\x87\x1ez\x085"#\xfd\x16\xc7o\x18|\xb7\x06\xdfHA\xdc6\xcf\x91\xcb\xe5\xa2\xac\xac,\xfaj\xceW\xd4\xa6M\x1b\xd2h4\x9e\xee"h\xd0iu\xf4\xd0C\x0f\xd1\xd2\xdf\x96R~~>\xb997qA\xact\xfa\x86\xb5>[\xe7\xfb\xf4\xe9CN\xa7\x93|\x11\x04\x9e~]\xbaD\x91\xee\xed\xb7\xdfN\xa7O\x9f&\x8e\xe3H\x10\x04\xfaw\xcf\xbf\xb2uHH&#\x8cF#\x99\xcd&\xb2\x98-d\xb6\x98\xc9`0(\xeeQ\xa7N\x1dZ\xb7n\x9d\xdf\x16t\xbb\xc3N\xbf\xfe\xfa+i\xb5Z9\xee\xe2_\x17{m\xfb\xe7\x89\xe7\xc52\xdb\xbcy3\xc5\xc6\xc6\xc9\xf1X\x96%\x93\xc9Df\xb3\x99\xccf\x0b\x99\xcd\x16\n\t\t\x91M\x11\x00\xa0\x98\xd8\x18\xfa\xeb\xaf\xbf\xfc\xd2\r\x16r\xcf\xe7\xd2\xe3O<.]/\x994\xd0j\xc9d6K\xe9\x98\xc9d6\x91F#Z\x99\xf4\x84q\xe3\xc6)\xac\xbfr\x1cG\x7f\xfd\xf5\x97\xa2\xcc\x17-ZTf\xce\xc2\xed\xa6\xc3\x87\x0fStt\xb4\xe2yX\xc9\xfa)\x02\x98\x82\xe08\x8e6o\xda\xac0\x17\xc10\x0c\x85\x84\x84\xc8y3\x9b\xcd\xa4\xd3\xe9\x14y\xabS\xa7\x0e\xed\xdc\xb9\x93\x04A\x90\xef\x95\x9e\x96N\xf7\xddw\x9f\x1cGo\xd0\xd3\xd4O\xa7J\xa6C\xfc\xcb%P\xc8\xca\xca\xa2\xa8\xa8\xa8\xb2{\xe8u\x94\x95\x95\xe5\x17/P\xb8p\xe1\x02\xbd\xf2\xca+\x8a|\x0e\x180\x80\xce\x9d;\xe7\x17\xd77\x042\xb3\xe1)+\xb9|\xe5\xf8\xfef0x\xc9\xcak\xac\x97)\x88A\x83\x06\xd1\xe1\xc3G\xe8\xf4\xe93\xd4\xaf_?\xf9x\xbbv\xedh\xc3\xc6\r\xe4\xaa`\xfd\t\x16\xdc\x9c\x9b\x1cN\x07\x1d9r\x84^\x7f\xe3\r\x85\xc9\x95\xf2BDd\x04=\xf2\xc8#\xf4\xcf?\xffPqqqP\xab\xbc7R\xb8A\xc5\x9a\x88gZH\xa3\xd1 ..\x0e\xa3\x1f\x1f\x8d\xb9\xf3\xe6\xe1\xd5\xd7^\xbb\xe8\xb4\x90\x9bsc\xe1\xc2\x85\x187n\x1c\xdez\xeb-\x1c>|\x18\xa5\xa5\xa5~\xdb\xe2})\xef\x9c7\x04\xa0\xa8\xb8\x04\xdf|\xfd\x8d\xe2x\xff\xdbn\x93\xed\xa1\xf3\x02\x8f\xbd\xfb\xca\xe6\xa0!\xdd\xdf\xe1p\xc0f\xb3\xc3j\xb3\xc2f\xb5\xf9-`\x9d>}\x1a\xc7\x8e\x1d\x83\xd5\xaa\xd4\x06\x12x\x1ev\xbb]\xa1\xae\xe9\xfb)K$\xfa\x04\xd8\xb5k\x17\xf2\xf2\xcb\xd2\x15$\xa7&6\x9b\r6\x9b\x156\x9bU,\x0f\xaf\xe7=w\xf6\x1cN\x9c<\xe9\x97\x9f`\x14\x17\x15c\xc1\xc2E\xd2/\xf1>n\x8e\x83\xddf\x93\xd2\xb1\xc1n\xb3\x83\xe7\x95\x93d\xfb\xf6\xedC\x91\xaf\x817\x9fbgXV\xce\x1b\xc3\xb2\xa8[\xb7.^~\xf9e\xf9\xbc \x88\x16Q\x83\xe1r\xb9\x90\x99\x95\xa90\xbeGD(--\x95\xf3f\xb3\xd9\xfc\x8c\xc0\x9d\xcf\xcd\xc5\x96-e\x96]I\xdaH\xa5\xf8"$q\x8d\xe5R \x12|T=\x19\xff\x87\xf6\x81@\xe0x\x0e;v\xecP\xccw7l\xd8\x10o\xbe\xf9\xa6\x97\x11\xb8\xf2\xeeC\xe2\x94V\x80z\xedy\xa6\xb2\'\xf3\xadK\x9e]\xc4>\xa7\x18@\xa3\xd5\xa0vtm\xbc\xf2\xea+\x88\x8b\x13\xa7<\x0f\x1c8\x80o\xbf\xf9\xaeB>A\x02AD\xe0y\x0e\xb99\xb9\xf8u\xf1\xaf\x187n\x1c\xa6}\xfa\xa9b\x17w BC\xc30t\xe8P|\xfe\xd9g\xf8\xfc\x7f\x9f\xa3k\xd7\xae\xb0X,\xbe\x8fS)\\\xces]Mnh\x01\xe0\x8b\xe73\xf3\xb5W_\xc5\x17_\xfc\x0f\xa3G?\xe6\x1bE\x01\x11!7\'\x17\xb3g\xcf\xc6\xddw\xdd\x8d\x9f\xe6\xfe\x84\x94\xd4T\xbf\x0e\xc9\x17"\x8f\nDp\x1cv\x07V._\x81]^j\x96\x8d\x1a5B\xbb\xb6ma2\x8b\x9a66\xabM\xec\xc8%\xfd\x7f\x9dN\x87\xd1\xa3G\xe3\x83\x0f>\xc0\xe4\xc9\x931e\xca\x149\xbc\xf8\xe2\x8b\xf2} \xe9\xa5;|L(\x934W\x7f\xb1Jh-\xb5"++\x0bn/\xe7*c\xc6\x8c\xf1Ks\xca\x94)x\xe9\xa5\x97\x14\xd7\x1e<p\x006\xebEL7Ke\x94\x94\x94\x04\x87\xd7\xde\x86\xdbn\xbf\r\x13\'M\xf4K\xe3\xf5\xd7_G\x9b6m\xe4x\xa9\xa9\xa9\n\xd3\x0b\x9eNV\x89\xb2\xd3\r\xb1X\xd0\xbf\x7f\x7f\xf4\xf2\xda9]\x1ev\xa7\x03G\x8f\x1eU\xac\x97\xbc\xf8\xe2\x8b\x98<y2&\xbf\xff\xbe\\\x16\x93\'O\xc6\x13O<!Oe\xb8\xdcnddd\xc8\x9b\x88\xc4,\xf8\xe6M\x12\xbc\x15\x9c&\x04\x00\x06,\xb4~\xd3=\xc1\xaf\x17\xa7\xff\x18ddd\xe0\xfd\xf7\xdfW\x9c{\xfa\xe9\xa7\xd0\xbe};i\xfa\x88\x11w\xe6J\x1d\xb5\x7f\x96\xfcU\x95<\x03+F\xba\xc0\xf3W\x8c]\xa6\x04!\xfe\x15\xd7r\x14H\xf6\x85X\x86E\x8f\x1e\x89xf\xdc8\xf9\xd4/\xbf\xcc\xc7\xfc\xf9?\xc3f\xb7\x01\x97\xb0\x96\x06\x00EE\xc5\xd8\xb8q\x13\xdez\xeb-\x8c\x1a5\n[\xb7nUl\xe0\x0bD\x9d:u\xf0\xd1G\x1f\xe2\xe3O>\xc1C#G\xa2Fd\r\xa5\xb0\xaed\xae\xe6\xbd/\x0b\xdfO\x82\x1b=p\xd2\x14\x87\xc3\xe9\xa43g\xce\xd0w\xdf}O\x9d:uRL\x8b\x04\x0ba\xe1\xe14t\xd80\x9a\xf7\xf3\xcft\xee\xdc9\x9f\xa9\x93\xb2\xe0;\x05$Z\x03-\x15\xad\x1f\xda\x1d\xb4g\xcf\x1e\xfat\xda4j\xdb\xb6\xad\x1c\x87a\x18z\xe2\xc9\'\xe8\xcc\x19q\xfa\x87\xe7yJII\xa1a\xc3\x86\xc9qZ\xb5jE\x87\x0f\x1f\xa6\x92\x92\x12*,*\xa2"\xafp\xf4\xe8Q\n\t\t\x91\xe3\x86\x84\x84PRr\xb2"_\xa5\xa5\xa54k\xd6L\xc53\xfd\xfa\xeb\xaf~\xcf\x91\x92\x92LC\x87\x0e\x95\xe34i\xd2\x84\xf6\xec\xf9\x97\n\x0b\x0b\xa9\xa8\xa8\x88\x8a\x8b\x8b\xa9\xa8\xa8\x88\n\xf2\xf3\xe9\xd0\xa1C\xa4\xd1\x96M\xa9\x99\x8cFJNN\xf6\xbb\xa7op:\x9d4a\xc2\x04\xf9:\x9dNG\x0b\x16,\xa0\xc2\xc2B*,,\xa4\x82\x82|***\xa6\xe2\xe2":{\xee,\xfd\xe7?\xff\x91\x1d\x8e\xe8t:\x9a<e\n9]N\xe287\xb99\x8eV\xaf\xfe\x8b\xcc\xe62+\xa7\x8b\xfc\x9e\x8b#\x9b\xcdNS\xa6L!\xbd^\xaf(\x03\x04\x98\x02J\xcf\xcc\xa0[\xda\xde"\x9f\xef\xd0\xbe=\x9d8q\x82\n\xf2\xf3\xe9\xfc\xf9\xf3TPP@\xc5\xc5ETXXH\x1b\xd6o\xa0\xf8\xf8x9\xee\xe0\xc1\x83\xe9\xf8\xf1\xe3\xc4\xf1\x1c\xf1^\x8eY\xbc\x9fu\xe2\xa4\x89\x94\x9b\x9bK\x05\x85\x05~\x16[KKK\xfd\x1c\x1ceeeQ\x8b\xe6\xcd\xe5{\xe8\xf5\xfa\x8bN\x01Y\xad6\xfa\xe8\xa3\x8f\xbd\xa6\x08\x19j\xd7\xae--X\xf0\x0b\x1d:|\x88\x0e\x1f>,\x85Ct\xe8\xe0A:x\xf0 \x1d8\xb0\x9f\xf6\xed\xdfO\xf9\x05\x05\xc4\xf3\x1c\xb9\xddnZ\xbdz5\x99\xbc,\xc8\xdez\xeb\xadt\xe0\xc0~:r\xe4\xb0\x14\x8e\xd0\x91\xa3G\xe9\xd8\xb1c\x94\x9e\x96F\xd9\xd9\xd9\xe4t:\x89\xe7\xc5i\xcc\x9d;vRl\xacr\n\xe8\xc8\x91#\xe4vs\xe4\xe68:|\xe4\x08\r\x192D\x9eN\x8c\x8f\x8f\xa7e\xcb\x96U\xd8\xc2\xac\xcb\xe5\xa2\xa3\xc7\x8e\xd2[o\xbdE\xcd\xbd\xca(X`\x18\x86j\xd7\xaeM\xcf>\xfb,\x1d<xPt6t\x13L\xf7\x04\n7\x9d\x00\xf0\x0e\x1c\xc7\x91\xd3\xe9\xa4\x03\x07\x0e\xd2\x7f\xff\xfb_\xba\xf5\xd6\x8e~\x95%P\x88\x8c\x8c\xa4\x87\x1f~\x98v\xee\xd8I\xc5E\xc5~\xf7\xf4\x9d35\x99L\xd4\xabW/\xea\xd3\xa7\x0f\xf5\xee\xdd\x9b\xea\xd7\xafOF\xa3Qq\xcf\xae]\xbb\xd2\x9a\xb5k\xca\xee\xe3\xe6h\xdd\xdf\xeb\x14\xf3\xffC\x86\x0c!\xbb\xdd\xcbs\x98W\xc8\xcd\xcd\xa5\xd1\x8f\x8fV\xdcs\xc1\xc2\x85\xc4y\x99\x96.--\xa5\x193g(\xe2\xf8\n\x007\xe7\xa6u\xeb\xd7+:\xb4a\xc3\x86\x91\xcd\xe6\x9f.\xc7\xf1\x94\x9d}\x96\x86y\t\x0b\x00\xb4\xd0\'\xdd@!??\x9fz$\xf6\x90\xafi\xdc\xb81m\xda\xb4)\xa0\xe0p\xba\\4y\xf2dy\xbe\x9da\x18z\xea\xe9\xa7(\'G\x14\xc2nN\xec\xa4\xbc\x05\xc0\xe2\x00\x02\x80\x17\x04:r\xf4\x08u\xee\xd2Y\x91_\xf8\x08\x00\x8e\xe3h\xeb\xd6\xad\x8a\xf3S\xa6L\xa1\xa2\xa2"\xbf\xbc\xf1<OEEE\xd4\xbf\x7f\x7f9ndd$\xadZ\xb5J\xf6\xfa\xe4\xf1\x08\xe6}\xbf\xd8\xb8X\xea\xd2\xa5\x0b\xf5\xe8\xd1\x83z\xf7\xee\xad\x08]\xbbv\xa5e\xbf/S\xac\xa5ddfR\xb3f\x97&\x00\x0e\x1d<H\xb7\xddv\x9b"]\x9dNG\xf5\xea\xd5\xa3\xa6M\x9bR\x93&M\xa4\xd0\x98\x1a5jD\x8d\x1a5\xa2\x86\r\x1bR\xabV\xadi\xc3\x86\r\xc4\xb9\xc55\n\xdf\x01\x8dV\xab\xa5&M\x9aP\xd3\xa6\xe2\xf5M\x9b6\xa5\xa6M\x9bR\xf3\xe6\xcd\xa9m\xdb\xb6\xd4\xa9S\'\xc5:\xd6\xce\x9d;(66F\xbe~\xd0]\x83\xe8\xd8\xf1cr>\xddn7-]\xbaT\x16\x12\x0c\xc3\xd0\xd0\xa1C\xe9\xe8\xd1c\xe5\xac\'I\x83\xb8\xecl\xfa\xf6\xdbo\xa9}\xfbv\x15\x1a\xc4i4\x1a\x1a9\xf2!Z\xbat)\xe5\xe5\xe5]\xb4\x9e^\xbdp\xbd\xd2U\x86\x9bj\n\xc8\x17F\xf2\xd5\xdb\xa6Mk\xfc\xf7\xcd7\xf1\xf1\xc7\x9f`\xec3c}\xa3\xf9QPP\x80\xf9\xf3\xe7\xe3\xf1\'\x9e\xc0\xf7?|\x87c\xc7\x8e\x89s\xc1\xd2\'\xab\xe0\xb3\xc9\xc8n\xb7c\xcb\x96-\xd8\xb4i\x136o\xde\x8c\xf4\xf4tY\xb5\x10\x00:u\xee\x8c\x89\x93&\xa2o\x9f\xbe\xe2\x01\x02\\n\x17N\x9f9\x8d\xcc\xccL9\xde\x1dw\x0c\x846\x88\x013\xd14A\xd9\xceT\x00\xd8\xbd{78\xbe\xec\x13\x98\xf18\xdd.\x07\xce\xcd\xe1\xec\xd9\xb3\xc8\xce>#\x1f\xeb\xd7\xaf\x9f\xec\xd3\xd8\x1b\x86\x01Lf#:t,\xb3Z\t\x00\xff\xfc\xf3\x8f"]oH*\xa3\xe4\xe4d\x9c\xce\xca\x92\x8f\xb7n\xdd\x1a\xb1q\xb1\x8a\xb8\x1e\xb4\xac\x06\x1d:t\x90\xcd#\x13\x11N\x9e<\x89\xd3\xa7\xcf\x88\x9f\xd4\xe4y6\xaf\x8b\xfc\x9e\x93\x01\t\x84\xc6\x8d\x1bc\xec\xd8\xf2\xdf\xb1@\xa2\xcf`o\x1a4h\xe0\xe5$E\x89\xd1hD?/\x0fv\x05\x05\x05HOO\x87\xcb\xe5Y\x1f\xf0\xcd\x0bp6\xfb,v\xed\xda\x85m\xdb\xb6a\xf3\xe6\xcd\x8a\xb0s\xe7NQM\xd4k\xba@\x9cOW\xdc\xe2\xa2\x9c\xbfp\x01G|\xf6r\xb8\xa5)\xaa\xa4\xa4$$\'\'K!\x05\xa7N\x9d\xc2\xa9S\xa7D\xf3\xdeY\x19\xe2\xd4I\x90\xf48\x8eCrr2\x92\x92\xc4\xeb\x93\x92\x92\x90\x94\x94\x84\x13\'N\xe0\xe0\xc1\x83\xf2n\xeb\xb2\xf5%F9]%-\xf9{\xea"\xcb\xb2\xe8\xdf\xbf?\xc6\x8cyZ<M\x84e\xcb\x96\xe1\xf7e\xbf\x054\x8c\xc80\x0cJK\xad\xd8\xbem;\xde\x990\x01O=\xf5\x14\xf6\xef?p\xd1\xe9\x9e[o\xbd\x15\xb3f\xcd\xc4\xe4\xc9Sp\xcf\x90!\x08\x0f\x0f\x97\x8c\xd2]\x0f\xaeW\xbaJnj\x01P\xd6.\tf\xb3\t\xdd\xbau\xc3{\x93\xde\xc5\xd2\xa5K\xd0\xa9S\'\x85\x8d\x9d@\x1c?~\x0c/\xbf\xfc\n\xfe\xef\xf5\xd7\xb1p\xe1B\x14\x15\x17\x03\x0c\xc0\xb2\x0cX\x8d\x06\x9arBXX\x18\x1a5j\x84\xf1\xe3\xc7\xe3\xdbo\xbeF\xdf>}\xc0Js\xbc\x04\x12m\xd0\xa7\x9e\x02\xa4\x06\x02\x00\xf5\xea\x89\x1b\xa4\x02a0\x18\xd0\xb0aC\x18\x8dF\xc9\xbe\x0f\x83-\x9b7\xe3\xc2\x85\x0b^\xf3\x8eek\x13\xc1\xd4\x07\xed\x0e\x07\x8e\x1f?\x0e\xa2\xb2t\x1b4h ;\xa0\xf7\xc5h0\xa2I\x93\xa6\xd0\xeb\xf5r\xba\xde~\x0b|a\x18\xb1#>p\xe0\x00\n\n\n\xa1\xd1hd\x93\x16\x11\xe1A\x1cd3@\xf3\xe6\xcdQ\xa7n]0\xd2\xa2\xfe\xe1C\x87q6\xfb\xacbA[\xa3\xd1\xc8y\xf6\x15t$\t,\x9dV\x87\xdbo\xbb\x1dC\x86\x0c\x91\xaf\x81\xcf\xdc,\xcf\xf1X\xbf~\xbd|\xbe}\xfb\xf6\xa8W\xaf\x9eB\xdf\xdf\x1b\r\xabA\xf3\xe6\xcd\x01\xc9\x86\x8d\xc9d\xc6\xfe\x03\xfbQR\\,\x169\x89\x03P\xdf:\x10,\xc8x=\x03\xcb0 i\x87\xb8\xc7\x9c\x87\xef3z\xe3\x19\x888\x1c\x0e\xbf\xfb_,h\xb5Z\xb9<<\xa5\xc2\xb2\xac_\xbc@\x81eYhuZ\xa9\xc0\xcb\xf2\xe7yo\x1a\xa9\x8e\x88g\xca\xca<44\x14\xa3F\x8d\x92\r\xf8\xb1,\x8bY\xb3fa\xf3\xe6\xcd\x8a\x9d\xeb\x1c\xc7\xe1\xf8\xb1\xe3\xf8\xf4\xd3O\xf1\xca\xab/\xe3\xbb\xef\xbe\x93\xcf\x05\x82a\x18\xc4\xc6\xc6\xe2\x95\x97_\xc1\xf7\xdf\xff\x80\x87\x1f\x1e%\x19m\xbbt\xd5\xee\x1b\x11\x86\xe7y\xb5\x1c|\xe0y\x1e\x19\x19\x19X\xb0`\x01\xfe\\\xfd\'v\xee([\xa8\r\x86\xc9d\xc2#\xa3F\xe1\xc9\'\x9f@iI)6l\xd8\x00\x9dN\'\x8f\xdc<m\x95\x91\x16\xcd\x9a4i\x8c\x96-[\xa1y\xf3\xe6\x01\x05\xcd\xb9\x9cs\xf8\xf6\x9bo\xc1q\x9cTY\tO=\xf5\x14\xea\xc4\xd7\xf1\x8d\n\x00\x10\x04\xc2\xbe\xfd\xfb\xf0\xd7\xea?\xa5\x85K\xb1\xa3\x18>|8Z\xb4l\t"\x01.\xa7\x13k\xd7\xac\xc5\xbe\xfd\xfb\x00iW\xf0m\xb7\xdd\x86\xc4\xc4D\xb9\xc1_\xb8p\x01s\xe6\xcc\x11\xbfbH\xdc\xc4\xf3\xf4\xd3O\xa3n\x9d\xba\x01\x07-\x02\t\xd8\xb7w\x1fV\xaf^-\x8f\xee5\x1a\r\xee\xbd\xf7^\xb4h\xd9\x12\xf0\xd1\x80\xf1t\x86\x0b\x17.BRR\x99\xe9\x88&M\x9a`\xc4\xfd\xf7C\xe3c\xab\xc8\x83\xd3\xe5\xc2\xf4/\xbf\x84MZ4&"\xb4n\xdd\x06\x83\x06\r\x82\xd1h\xc4\xbe\xfd{\xb1j\xe5*\x10\x08\x02/\xa0O\x9f>\xe8\xd3\xb7O@\x1dw\x97\xcb\x85\r\x1b\xd6c\xf7\xee=\xf2\xbd,\x16\x0b\x9e\x7f\xe1\x05\xe8u:\xe4\xe4\xe4\xe0\xab\xaf\xbe\x02\x81\xc02,\xc2\xc3\xc31\xfc\xbe\xe1\x88\x8f\xaf\x13\xa8\x08@\x92}\x9c\x15\xcb\x97#,,\x0c:\x9d\x0e\x1a\x8d\x06w\xdey\'"#"\x90\x99u\x1aK\x96,FIi\xa94*&\xbf\x11\xa0\xf7\x11\x9e\xe7\xd1\xbf\x7f\x7f\xf4L\xec)\x8fP\xcf\x9d;\x87o\xbf\xf9\x06<\x95\x8d\xac\xc7\x8d\x1b\x87\xa8(\xa5\te\x0f\x02\t\xd8\xbfo?\xfe\xf8\xe3\x0f\xe9\x88\xa7\x99\x8b\x1d\x9f\'-\xb9\xf1{\xf6\xbbH\x7f\xef\xba\xfbnth\xd7\x0e\x04`\xef\xde}\xf8\xe3\x8fU~\xba\xf0\x9e\xb8De\x9b\xf0H\x10Ms?;~<j\xd5\x8a\x02\xcb\x02G\x8f\x1e\xc3\xd2\xa5Ke\x81\x15\x1a\x1a\x8a\xfb\xee\xbbO\xde\x9c\xe7\xc1\xe5ra\xf3\x96\xcd\xd8\xbem\xbb\xfc\xf5c\xb1\x981\xfa\xb1\xd1\x88\x08\x0fG~A>\xfe\\\xbd\x1a\x9fO\xfb\x1c\'\x93N\xfai`\xf9b\xb1X0t\xd80<<r$z\xf6\xec\t\x93\xc9\x14\xb0n\xdd\xf0\xf8W7\x19U\x00\x04\x85\x81\xd5Z\x8aC\x87\x0fa\xe9\x92%\xf8\xfc\xf3\xff\xf9F\x08H\xc3\x86\r1v\xecX\xdc~\xfb\xedh\xdd\xba\xb58ZS4\xbf2\xad\n\xb9\xf1\x04\x80\x08\xb2I\x01O<\xdf\x06\xa8\xc4k\xbb\xbe\x94`\xa0\xca\xee\xe9\xa4E\x18\x80\x11;9\xdf8\x9e\xbf\xbc @\xc3\xb2A\xd3&H\x99\x95\x12\xf5\xdcZ\x8c\xee\x9f\xbe\x07Y\xbb\x86\x91\xbaD\xe9B\x8f\x80\x0c\x84\x7fY\x89iz\xca\xa7\xec\xd9\x18\xb0l\xf0\xfb\x88\x88y\xf5\xbe\xd6\xf3\x8c\x0e\xa7\x03 Q\xebJ\x1emKW\x95wGx\xe5Q\x10\x04)\x0f,\x88\x00\x9ewC\xa3\xd1\xcae\xceH\x0e\xdb\xe5{{\tJ\x92F\xef\x1a\xaf\xaf\x01\xefw\xc2\xb0,<\xd3\x8d\xc1\xde\x0b$\xdb?\x0c\x03\xd1\x9e\x03<SH\x9e\x0e[\xaa$\xd2\x83y\xf2\x02\xd1\xc8\x01@$\xed6\x07\x04\x10@$\x9a(!O\xe5\xf5\xaao^\x04z\x8f\xe2\xbb&\xd9T2\x91\xf8%\x13\xe8\xfd\xc8\xefX:\xc7H\xbb\xb7w\xed\xda\x85\x85\x0b\x17*\xfc \x94GbbO<\xfe\xf8h\xf4\xef\xdf\x1fu\xea\xd4\x95\xac\x8f\xfa\xa7w\xb3\xa3\n\x80\xa0\x88\x15F\x10\x04\x14\x14\x14`\xdb\xb6m\x98:u*\xf6\xee\xdd[!=\xf7;\xef\xbc\x03\xc3\x86\xdd\x8b{\xee\xb9\x07\xb5j\xd5*k\xdc\xe2\xfeT\xf1\xff\xbe\xfa\xe1\x95\x88\x98\xfb\xb2N.\x10\x95\x95\xbe\xa2iU\xa0\x9d\x89y\n\x10Q\xee\x98\x82Ab\x07\xe4w\x99\xb2\xf3\x84\xdc\x7f\x94w\xaf\x8b\xe3{\xdfK*+\xbfg\t\xf0\xbc\x97\x85\xe7>\x9ef\x1b\xfc\x9e\x97\x9cg/\xbc;r\xe9H\xb9i\x05\xa3,\x97\xde\xd7\x13<\xdb |-i\x12I\xc2\x11\x02\\.7\x0e\x1e<\x88_\x7f\xfd\x15k\xd6\xac\xc1\xd1\xa3G\x15q}a\x18\x06111x\xf8\xe1\x871\xfa\xf1\xd1hP\xaf\xbel3K%0\xaa\x00\xa8 <\xcf#==\x1d\xbf\xff\xfe;V\xadZ\xa5\xb0\x93\x13\x8c\xb0\xb00\x0c\xbas\x10\xc6>\xf3\x0c\xda\xb5k+o\xf0\xba\xdc\xc6t\xd3sS\x14\xdbM\xf1\x90A!\x1087\x87\xdc\xdc\\\xacX\xb1\x023%\xafe\x17\xf35\x11\x15\x15\x85\xdbo\xbf\x1d\x8f<\xf2\x08z\xf6\xec\t\xb3\xb4\x97F\xa5|T\x01p\t\x08Dp:\x1c8v\xfc\x18~\xfb\xed7|\xf4\xe1G\xbeQ\x02b0\xe8\xf1\xdf\xff\xbe\x89{\x86\x0cA\xcb\x96-\xa1\xd7i\x03~>\xab\xa8\xdc\xcc\x10\x80\x0b\xe7\xcfc\xfb\xf6\x1d\xf8y\xfe<\xfc\xb6\xb4\xcc\x14yyt\xeb\xd6\r\xe3\x9f}\x16\x89=z >>^1u\xa6R>\xaa\x00\xb8\x04<\x05E\xbc\x80\xa2b\xd1E\xe1\xcc\x993\xb1y\xf3fyq\xb2<\xfa\xf6\xed\x8b\xfb\x1f\xb8\x1f\x83\x06\xdd\x85\xb8\xb88\xbf\xcf_\x15\x95\x9b\x12\x12M\xaf\x1c8x\x10?\xfe\xf8#\xfeZ\xbd\xda\xcf=\xa4/\xacd\xde\xe3\xb1\xd1\xa3\xf1\xc8\xc3\x8f\xa0N\x9dx\x18<j\xbaj\xb3\xaa0\xaa\x00\xb8\x02\x04A@\xf6\xd9l\xfc\xb6t)~\xfdu\t\xb6o/\xf3\x1c\x15\x8c\xd0\xd0P\xf4\xeb\xdf\x1f\xcf?\xf7\x1c\xda\xb5o\x8f\xf0\xf00\xbfEX\x15\x95\x1b\x19\x92\xfea\x18\x02\xc7q8s&\x1b\xbf\xfd\xf6\x1b\xbe\xf9\xe6\x1b$%%\x95-\x04\x07!!!\x01\x03\x06\x0c\xc0\x83\x0f=\x88\xae]\xba\xc2d4\xa9\x9d\xfee\xa2\n\x80+\x85a\xe0\xb4\xdbq\xe0\xe0A\xfc\xfe\xfb\xef\x98:u\xaao\x8c\xa0L\x9a4\t\x83\x06\xdd\x89\x96\xadZ\x89\x95XE\xe5&"??\x1f[\xb7n\xc5\xbc\xb9s\xf1\xfb\xb2e\xbe\xa7\x03\xd2\xb7O\x1f\x8c\x1d7\x0e={&\x8a\xbe(n\xee%\x93+F\x15\x00\x95\x84 \x08(..\xc6\xee\xdd{\xf0\xd5\xd7_c\xcd_\xabe\xa7\xdb\xe5\x91\xd83\x11\xc3\x86\x0e\xc3\xf0\xfbF >.\xb6\\\xb5\xbeKCm\x19*U\x13\xa7\xd3\x89\x03\x07\x0e`\xc9\x92%X\xb4h!\xce\x9c\xc9\xf6\x8d\xa2\x80e5H\xa8\x97\x80q\xe3\xc6\xe3\xbe\xe1\xc3\x11\x17\x1b\x03\xbd\xc7\x7f\xb1\x9f\xb6U\xf5\xe1J\xb4\xb4*\x0bU\x00T2\x82  \xe7\\\x0e\x96,]\x82\xe5\xcbW`\xe3\xc6\r\xbeQ\xfc0\x1a\x8d\xe8\xdc\xb93^~\xf9e$&&",<\\\xb1\xa3U]+P\xa9\xaex:9\x02\xc1\xedv\xe3\xdc\xd9sX\xb2t\tf\xcd\x9c\x85\xcc\xcc\xcc\x8bj\xf7\xd4\x8a\x8a\xc2C\x0f>\x88\x07\x1e\xb8\x1f\xed\xda\xb7W\xbf\x94+\x19U\x00T2D\x00\xa4\xb9\xcd\xa3G\x8e\xe2\x8f?Va\xce\x9c9\xc8\xce>\xeb\x1b\xd5\x8f\x10\x8b\t\xcf?\xff\x1f\xdc=x0n\xb9\xe5\x16\x18\x8d\xa62\xdf\xaa**\xd5\x10A2_q.7\x07\xdb\xb7m\xc3\xecY\xb3\xb1\xd1\xcb-ey\x0c\x1e<\x18\x8f>\xfa(\x12\x13\x13\x11\x15\x15u\xddG\xcb7"\x95&\x00.e\xc2\xe1R\xe2V7\xc4g\x137,\x11\x04\x14\x15\x15c\xf7\xee\xdd\xf8\xe4\x93O\xb0k\xe7N\xd8*0-\xd4\xb5kW\xdc3\xf8\x1e\x0c\xbf\xef^\xd4\xaf\xdf \xa0!6x\xd9\x08RQ\xa9\xaa\x90dNd\xfc\xf8\xf1X\xbbv-222|\xa3(`\x18\x06\x8d\x1a5\xc2\xd8g\xc6\xe2\xdea\xc3\x11\x13\x13\r\x83\xd1 o\x10S\x85@\xe5Ri\x02\xa0\xear\xfd\xc5\x8d \x08\xc8\xcd\xcd\xc5\xd2\xa5K\xb1|\xf9r\xd9\xd0Xy\x18\x0c\x06t\xe8\xd0\x01o\xbe\xf5&:w\xee\x8c\x88\x88\x08hY\xad\xe2QT\x01\xa0R]\xd0h4e\xe6?\x82\xd0\xa0A\x03\x0c\x180\x00\x8f>\xfa(:t\xe8\x08\x9d^w\x9d[\xee\x8d\xcfM \x00\xaa\x0e\x0e\x87\x03IIIX\xb5r%&\xbc\xf3\x8e\xef\xe9\xa0\x8c\x1b?\x0e\xf7\xdf\x7f?:\xdd\xda\tf\xb3YV\x93S\x05\x80Ju\xe1b#\xf7\xa1C\x87\xe2\xb1\xc7\x1eC\xb7\xee\xddQ\xa3F$4\xac\xba\x99\xebZ\xa0\n\x80k\x0c\x11\xa1\xb0\xb0\x10\x07\x0e\x1c\xc0\x8c\x193\xb0~\xfdz\xd9}`yt\xec\xd8\x11w\r\x1a\x84Q\x8f>\x8a\x84\x84\x04h\xb5ZU\x00\xa8T\x1b\x02\t\x00\x96e\xd1\xb8qc\x8c\x1f\xff,\x86\x0e\x1d\x8a\xe8\xe8\xda\xd0\xeb\xf5UB;\xe6fA\x15\x00\xd7\x01\x81\x08\x0c\x80\x9c\x9c\x1c\xacX\xb1\x02\x8b\x16/\xc6\xa6\n,\x8c\xe9t:\xf4\xe8\xd1\x03c\x9fy\x06}\xfb\xf4Ftt\x8co\x14\x15\x95*\x89o\x87\xde\xa4IS\xdcq\xc7@<2\xea\x11\xdc\xd2\xa6\rtz\xbd\x97\x03\x99\xaa\xc2\xf5\x9f>\xbe\xda\xa8\x02\xe0:\xe3v\xbbp\xf4\xd81\xacZ\xb9\n\x1f|\xf0!\x9c\xce2Oa\xe51n\xdc8|\xf1\xc5\x17\xd0\xe9\x02{\x08SQ\xa9Jx\x0b\x80\x07\x1ex\x00\x0f?\xf2\x08\xbat\xee\x82Z\xb5j*\xe2\xa9\\[nJ\x01P\xd5\xe4\xba \x08(.)\xc1\xa1\x83\x070{\xf6\x1c\xfc\xb1j\x15\xac\x15\xb0-TXX\x88\xf0\xf0p\xdf\xc3**U\x0e\x9dN\x87\x86\r\x1b\xe2\xe5W^\xc1=\x83\xefAdd\xa4\xe8aM\xf2\x07qyT\xb5\x96\\\xfd\xb8D\x01\xa0\x16\xf8U\x83D/N999\xf8\xed\xb7\xdf\xb0r\xd5J\xfc\xbd\xf6o\xdfX\n\xf2\xf2\xf2P\xa3F\r\xdf\xc3**U\x8e\xd7_\x7f\x1d#F\x8c@\xab\xd6\xad\xca\x8c\xb6U\x16j\xb7t\xd9\\\xa2\x00P\xb9\x9ax\xb4{\xdcn7N\x9c8\x81\x95+W\xe2\xeb\xaf\xbf\xc2\xe9\xd3e\x0e\xda\xbdQ\x05\x80Ju!??\x1f\x11\x11\x11\xd5\xdat\xc3\x8d\x88*\x00\xaa\x18\xb2%D"\x94Z\xad8z\xf4(>\xfa\xe8#l\xd8\xb0\x01V\xabU\x11W\x15\x00*\xd5\x05\xcf\x1e\x00o\x8fx*\xd7\x1fU\x00T\x06W\xf1\x13\x94\x88p\xe1\xc2\x05\xacX\xb1\x02\xbf\xfd\xf6\x1b\xd6\xae]+7\xa6\xea&\x00\x9c<!\xd7!\xc0)\x005\xf4\x0cj\x18\x94j\xac6\x9ep\xde!\x80\x13\x80\x1a\x06\x06\x91zU\xcd\xf5F\xe1b\x9b\xc0T\xae\x0f\xaa\x00\xa8\x06\x10\x11\\.\x17RRR\xf0\xf7\xbaux\xe7\xad\xb7`\xb5\xdb/K\x00\x14\xbb\t\xc7\x8b8\\p\x08(\xe1\x00\x97 \xaa\xa4\x02\x80E\xcb \\\xcf \xc1\xa2A\xe3PM\xa5\xcb\xb4\x8d\xe7\\xc\x7f)rK\x05<\xda\xd4\x84w\xdbZ\x14\xe7\xff\xcav\xe2\xcd\xfdV\x14:\x04<\xdb\xc2\x8cWZ\xaan\xfdn\x14T\x01P5\xb9\xe9\x04@u\xdcd"\xce\n\x89\xaf\xc9j\xb5\xe2\xc0\xc1\x83\x983{6f\xce\x9c\x89\xc8\xc8H\xdf\xe8~\x08\x04\x1c)\xe4\xb0\xe6\xac\x0b;\xcf\xbbq\xb0\xc0\x8d\xf3\x0eB\xb1[\x00<\xaf\x9fa`\xd0\x8a\xa3\xee\x04\x0b\x8b\xee\xd1:\xdc\x15g@\xafh=*k \xfes\xaa\x03O\xee(\x82\xab\x94\xc7\x83-C\xb0\xa0g\x98\xdf\xf9\xd1\xdb\x8b\xc0;\x04\xbc\xd06\x14_t\nQ\x9cW\xb9\xb6X9\xc2\xa2t\x07\x0e\x14\xf0h`a\xf1ls\xf3e\xd7\x05U\x00TMn:\x01p#\xe0\x99\x16\x8a\x88\x88\x80\xc1c\x17=\x08\xa7m<~<e\xc7\xc2t\'\x8e\x16p\x80\x9b\x00\r`\xd0\xb3\x88\xd030\xb0\xa20t\t\x84\x02\x17\xc1\xe9\x16\x00\x0e\x00\x0b\xc4\x87i\xf1h#\x13^jnB\x94\xf12[\xbe\x17\x0b\xd2\x1cxfW1\x8a\xad<F5\xb7`n\x0f\xa5\x00X\x90\xe6\xc0\x98\x9d\xc5(u\nx\xa5M\x08>\xed\xa8\n\x80\xebI\xba\x95\xc7\x9d\xeb\nq\xe2\xbc\x1b5\xc34H\x19Z\x0b\x11\xfa\xcb\x1b<\xa9\x02\xe0"\\\xc5i\xe4\xf2\xb8\xf2V\xadr\xcda\x18\x06QQQ\x17\xed\xfc\xf7\xe4qx~w\t&\x1d\xb0\xe2\xe8y7B\xf5,\xfa\'\x18\xf0n\x87\x10,\xe8\x19\x8e_z\x86c~b\x98\x18z\x86cA\xcfp\xbc\xdb>\x14\xbd\xeb\x1a`4\xb08S\xc4\xe1\xe3\xc3\xa5\x98p\xc0\x8a\x0bN\xb5\x01\xdfl\x10\x01V\x9e\x00\x8e`\xe7\xc5\x1d\xec*W\x89\xeb\xd0\xf9\xa3*\n\x80\x8b\xf9\x03U\xa9\x18\xff\xe6\xb9\xf1\x9f=%X\x96\xe6\x00/\x00}\xeb\x1a\xf0m\xd70\xfc\xd23\x1c\xef\xdc\x12\x82au\r\xe8\x17\xadGbm1\xf4\x8b\xd6cX]\x03\xde\xb9\xc5\x82_z\x86cj\xc7\x104\x8c\xd4B\xe0\x08\xdf&\xdb\xf0\xe9Q\x1b\xec\xea\xc7\xe2MEM\x03\x8b\x17\x9a\x990\xb2\xa5\x19/\xb50\xc1\xa2\xadr\xdd\x85\xca\x15R\xe5\xdehu\x9b\x9f\xaf\x8a\xe4\xd8\x05|x\xc4\x8a\x1dg\x9d`5\x0c\x1eml\xc2\xd7]\xc3p\x7f}\x03jW`*\'\xce\xc4\xe2\xb9ff|\xd4!\x04u\xc24\xe0\xdd\x84y\xa9\x0e\xec\xba\xc0\xf9FU\xb9\x81\t\xd31x\xb5\x95\x05\xf3{\x85cr\xbb\x10\x18T\x03\x9d7\x1c\x17\xef\rT\xaa\x15n\x810\xe3\xa4\r\xcb3\x9c\x00\x0b\x8cld\xc2\x94\xf6!h\x1c*\xb6\xdeR\x8e\xb0)\xc7\x85\x9fS\x1d\x98\x9b\xea\xc0\x9al\x17\n]\xe2\xc8\xfex\x11\x8fUg\x9cH-\x11\xdd\xf4\rO0\xe2\xb1F&0:\x06\xd9\xa5\x1c\xfe<\xe3\x84KP\xbf\x02.\x17\xb5\xe4T\xaa\x1a\xea"p5&\x909\xe8\x1d\xe7\xdd\x18\xf9O\x11\xd2\xf39t\x8f\xd7\xe3\xfb\xeeah\x16&z\x14;Y\xc4\xe3\xf3\x13V\xac\xccr!\xdb\xca\x01\x04D\x185x\xa0\xbe\x11\xa3\x1a\x1a1+\xc9\x86\x95YNLn\x1f\x8a\x17\x9a\x8b\xbeW\xb7\xe4\xba\xf1\xd0\xd6Bd\x17\xf1\xe8_\xd7\x80\x05=\xc3\x83.\x08\x97\xba\tk\xce:\xb1\xeb\x02\x87\x1c\xbb\x00\x01@\x94\x81Ebm\x1d\xee\x8c\xd7c\xf5\x19\'\x1e\xdf~e\x8b\xc0\x02\x80\xb4\x12\x0e\x9bs\xdc8^\xc4#\xd7)\x80\x17\x00\x93\x96A\xb4\x91A\xd7Zz\xf4\x8d\xd1\xc1\xa2\xf5\xff\x92\xfc\xe3\xb4\x13\x8b3\x9c\xe0\x03lF\xf5\xfc$"$\xd6\xd6\xe3\xb1FF\x185e\x91\xac\x1c\xe1`\x01\x87\xed\xe7\xddH+\xe1Q\xcc\x89\xcd&D\x0b\xd4\x0f\xd1\xa2o\x8c\x0e\x9dj\x06v`R\xe4&\xcc:a\xc7\xe1\x027\x06\xd51\xe0\x91\x86F\x1c,\xe0\xb0$\xc3\x89t+\x0f\x1d\x0b\xdcZC\x8b\xfb\xeb\x1bQ\xcbgoDr1\x8f\r\xe7\\8Z\xc4\xa1\xd0E \x00\x06\x16\x881\xb1\xe8TK\x87\x9eQz\xd40\x04J\xb5\x8csv\x01\x7f\x9du\xe2@\x1e\x8f|\x97\xf8^|\xdf \x01 \x8100\xde\x80Q\r\x8d\x00\x00;G\xf8\xe1\x94\x03\xff\xe4\xb8\xd00L\x83\tm.\xff+@]\x04\xae\x9aT\x0b\x01p\x9d\x16\xc8\xab\x0eA\n\xc0W\x00\xb8\x05\xe0\x8d\xfd\xa5\xf8\xecp)\xccz\x16_v\x0e\xc5\x93\x8d\xc5\x8e<\xdd\xca\xe3\xc5=%X\x9e\xee\x04@\xa8\x1b\xa6E-\x03\x83\xe4b\x1e\xa5n\xc2\xad\xb5t\xc8\xb2\xf2\xc8\xb1\x0b\xf8)1\x1c\x8fJ\x9d\xc0Y\xbb\x80Q\xdb\x8a\xb0>\xdd\x81\xd6\xd1z\xfc\xd1/\x02\t\x16\xff^ \xb5\x84\xc3\'\xc7lX\x9c\xee@\x81]\x00\xa4\x0e\x12,\x83P3\x8b\xd1\x8dL\x883\xb1\x98z\xd4\x86\xfcR\xee\xb2\x04\xc0i\x9b\x80\x1fO9\xb0$\xc3\x8e\xa3\x85<8\xb7 \xea\xb8\x92\xa8\xc6\n\r\x83(3\x8b\x11\xf5\x8dx\xa3\xb5\x05u\xcd\xca\xf2y}_)>\xd9W\xe2_\x9e\x8c\x98O\x90\xa83;\xb0\x81\x11\xbf\xf4\x0cG\rI\xe7\xf1\xdf<7f\x9e\xb4\xe1\xef\xb3.\x9c\xb1J\xcf\xe6\xf9\x12b\x00F\xcf\xa2I\x98\x16\xaf\xb7\xb2\xe0\xb1F\x06h|\xa4K\xb6]\xc0]\x1b\nq \xcb\x81\x07ZY0\xb2\xbe\x11\x93\x0f[\xb1\'\xd7%\xde\x8b\x80\x84\x9a:\xfc\xd23\x0c=\xa2\xf4\x80\xb4_\xe3\xe7T\x07\xbeK\xb1\xe3p\x01\x07\xb7K\x10\xa5\x1fH\xcc\xaf\x86A\r\xb3\x06}\xa3\xf5x\xb1\x85\x19\x89\xb5\x03[\x85\xdd\x9c\xe3\xc6\xc7G\xac\xd8\x9c\xe3\x82\xcd!xt\x8a\xcb\xca\xccSD$V\xa0\x07[Z\xb0\xa0\xa7h`0\xdf%\xe0\x91\x7f\x8a\xb1:\xd9\x8e\x86\xd1:\x1c\xbc\xbb&B\x02\x08\xd6\x8a\xa0\n\x80\xaa\x89\xef@\xa0JryU\xee\x06\xa2\x82\x05\x90i\xe5\xf1\xf7Y\'\xc0\x03=j\xebqg\xbc\xa8%\xc4\t\xc0\xec\x936,\xcfp\x00\x0c0\xa2\xa1\t?\'\x86aa\xcfp|\xd1)\x0c\xf5C5\xf87\xd7\x85\x1c\x9b\x808\x8bF\xd1q\xc6\x9aX\xd41i\x00I#\xc4\xd3\xaf{s\xde!\xe0\xed\x03V|u\xdc\x86\x02\x9b\x00\x8b\x81E\xcf\xbaF\x0cil\xc2\xad1z\xd8y`\xfaq\x1bf%\xd9\xc5\x85d\xdf\xe1w\x05\x99\x9fj\xc7\x84}%8\x98\xe3\x86Y\x0bt\x8e\xd1\xe3\x9eF&1\x9dX=Lz\x06\xe7\xad<f\x9f\xb4\xe1\xeb$\x1b\xdc>}N\xb7\xda:\xdc\xdb\xc4\x8c\xc1\x8d\xc5k\x8646aX\x13\x13\x8656\xa1q\xa4\x0e\xd02\x80\x96A\xb4Q#\xab\xc7f\xdbyL:d\xc5\x8f\xc7\xec8S\xc2#\xde\xa2A\xbf\xba\x06\x0cml\xc2\xa0\x86&4\x8b\xd2\x83\x00$\xe5\xb9\xf1\xf6\xc1R\xfcy\xc6\xa5LT\xc2\xa8\x11\x87\xee\x87\xf2\xddxs\x7f)\xf6\x9cs\xa1\x86I\x83.\xf1\x06$\xd4\xd4\xc2\xa2c\xc0J/\xba\x94#L=j\xc5\x1b\xfbJ\xb0/\xc7\x05\x8e\x08\xad\xa2t\x18\xd4\xd0\x80!\x8dM\xe8\x1ag@\x88\x81E~)\x8f\xa5\xa9v\xbc\xf4o\t\xb6\x9fw\xfb&\x89\xe3\xc5<\xde:P\x8a\xd5\x99\x0e\xd8\xdc\x84\x8e1z<\xdd\xca\x82\xf1m,\xe8\x95`\x84\xd1\xa4\x01X\x06:=\x8b.\xb1z\x0cijF\xbfXQ\x00y\xd0\xb1\x00t\x0c\x8c\x1a\xd5\x80\xc3\x8dH\xb5\x10\x00W\x83\x00\xfdX\xb5gs\x8e\x0bIE<\x18\x1d\x83\xbe1:\xc4\x99\xc4\xd7\xbb7\xdf\x8d\xc5\x19N\x80\x03\xba\xd6\xd6\xe3\xbd\xb6\x16\xf4\xaa\xadG\xd30-\x9ehl\xc4\xa4[B\x10ib\x01\x9e\x10e`P+\xc8\x14O0\x16\xa7;\xb0$\xc3\t\x10\x10\x1b\xaa\xc5G\x1dB\xb0\xa8W8~\xe8\x1e\x86\xa5\xbd\xc3\xf1\xc5\xad\xa1h\x10\xaaAV\x11\x0f;\'\x8d`/\x83<\x17\xa1\xa6\x91\xc1\xa8\xe6&\xccO\x0c\xc7\xef\xbd\xc54\xbe\x97\xd2y\xaf]\x08"\xcc,\xc8M\xf8=\xcb\x89cE\xcaE\xebAqz|\xdb-\x0c?J\xd7|\xdf]\xfc\xff\xcb-\xccb\x07\xcd\x11j\x98Y\xdc\x9b`\x90\xa7\x90J9 \xcf!\xa0qM\x1d\xdej\x17\x82_{\x89\xea\xb2\xdfw\x0f\xc3\xdc\x1eaX\xd83\x0c\xf7\xd53\x02Z\x06g\x8b9,\xcfr\xc2\x1dl\x9d\x84\x11\xd7Y\x8e\x16\xb8\xd1)F\x8f\x1f\xba\x87\xe1\xf7>\x11X\xd2;\x02\x9ft\x08A\x8bpq\xaa\xee\xf7L\'\xbe<aC\x89\x9dG\xa4E\x83\t\xb7\x84`Y\x9f\x08\xcc\xed!\xa6\xbb\xb4w\x04\xbe\xee\x16\x86[\xa3u\x00\xcb\xe0\xdf\x1c\x17\xa6\x9f\xb4\xa1@Z\xcb\xf1\xf0[\x86\x03\xdb\xce\xb9\x00\x96\xc1\x03\r\x8c\xf8\xb9G\x18\xbe\xe8\x14\x8ai\x1dC\xf1Kb8\x9ekf\x82N\xc3\x80\xe7\x04t\xaa\xa5\xc57\xdd\xc30\xaa\x81\xf8\xe5\xe7\xe12_\x95J5\xe1\xd2Z\xfa\r\xc4\xd5\xab\xd8A\x1a\xff5\xe0x\x11\x0f\xa7K@-\x13\x8b\xaeQeS\x02k\xb3]H/\xe6\xa010x\xa0\xbe\x01\xcd\xa5\x8e\xc6\xc3\x1d\xf1z4\x08\xd5\x00\x02\xa1q\xa8\x16\rB\x94S<\x9e\x01{\xa0\xca\x92c\x17\xb0\xe2\x8c\x13n\x97\x00\xbd\x9e\xc5\xf8f&\x8ckjF\xac\x89\x95v\x15k0\xbe\x99\t3;\x87\xa1]\x94\xb6l\n\xe22xD\x9a\x9a\x99\xd19\x0cw\xd71 \xce\xacA\r=\x8b\x1aR:O56\xa1S-q\x04\x9bm\x13p\xc6&.f{\xd0\xb3\x0c"%\x1bD\x9e\xeb,Z\x16\xcb\xb3\\8V\xc8\x81\xd12\x18\xd5\xd0\x84\xdb\xbcF\xc1uL,>\xe8\x10\x82_z\x85\xe3\xfdv!\xe8\x16\xa5Cm\xa3\xf8l5\r,\xdaE\xea\xf0j\x0b3\xe2,, \x00\'\x8b9d\xdb\xcb\x99\xee\x10\x08\xad#u\xf8_\xa7P\xdcS\xd7\x80X\x13\x8bN5u\xb8;\xde\x80\x08=\x83\x1c\x87\x80\x85\x19\x0e\x14\xdb\x04\xe8\xf4,\xc665\xe1\xcd\xd6f4\x0e\xd5\xa0\xa6\x94\xef83\x8b\x87\xea\x1b1\xa9m\x08b,,@\xc0\x9a3.l?_\xf6\xf5\x91\xeb\x10\xf0\x8f4\xc5\x94\x10\xaa\xc5K-\xcdh\x1e\xae\x85I#\x8e\xe6\xe3\xcd,\xc645\xa1E\x84\xa8\xea{ \x9f\x87\x93\'\xc5\xba\x87\xca\x8dO\xa06\xadrE\\\x9f\x06T\xec&\xa4\x94p\x00Oh`\xd1\xa2i\xa8\xd8\xc9[9\xc2\xe1B\x0e\xe0\x80\xfa!\x1a\xf4\x89V~\xe2\x03\xc0\x05\x87\x80|\x07\x01,\x83\x86\xa1\x1a\xc5<\xaf\x95#\x14\xbb\xc5Q\xbbE\xc7\x88S\x02^\x9c(\xe6p\xb4P\x1ci7\t\xd3`d}#\x02\xf5!w\xc6\xeb1\xba\x91\t\x06\xad\xd8Q^\x0e\xb7Dj1 \xce\x800]\x80\x04\x00\x84\xea\x184\x0ca\x01\x16\xe0(\xf0t\x95/\xbff8\xf0c\xaa\x1d\x82\x8bpk\x94\x0e\xcf41+\x16\x90\xcdZ\x06}\xa3\xf5\xe8TS\x1b\xf4\xcd6\x0e\xd3 \xde,\nM\x1bG\xb0\xb9\x83$,\x00\xe1&\r^jiFw/\x01\xed\xcd\x91"\x1e{/\x88\xd39\xcd"\xb4x\xaa\xb1\t\x86@\x05\n\xa0w\x8c\x1e\xbd\xa3\xf5\x00\x0b\x148xl\xc9q\xcb_\x1f\x17\x9c\x02r\x1cbA\xc7\x9b\x194\x92\xb4\xc0\xbc\x892\xb2\xa8cf\x00\x86A\x96\x8dGz\xe9e\xbe\x18\x95j\xcbu\x14\x00A\x1a\xc9\xd5\xe6:%{\xb5I-\xe1\x91\\\xc2\x03,\x83\x18\x13\x03\xa9?\xc2\xa9R\x1e\'\xa5\xa9\x90F\xa1\x1a\xd4\xb5\xf8\xbf\xf2C\x05<r\x9c\x02\x18-\x83\xda>\x1a%\xe9V\x1e\xe9\xa5\x1c\xc00\xa8m`\xe5\xb9q\x0f\xe7\x9d\x84\x0bNQ@4\xb0h\xca\xd5H\x895\xb1\x92\x16\xc9\x95\xbf\x84|\x17!\xd3\xcac\xf7\x057\x16\xa49\xb00\xdd\x81_\xd2\xec8m\x15G\xfdL\x05Dq\x96\x95\xc7\x8c\x936\\(\xe6\x11f\xd1`|S3\x9a\x87\xfbw\x94\x1e8\x12\xbfxRJx\xfc\x95\xed\xc2/i\x0e\xfc\x9a\xe1\xc0\xa2\x0c\x07\x8a]b\x19\xb0b\x7f\x1a\x18\x02\x1a\x84\xb0\xe8\x16\xa4\xf3\x87\x94\'Oy\xb6\x0c\xd3\xa2\xb64\x8d\x17\x88\x10-\x83\x9e\xd1z0\x1aq\x01;\xa5\x84\x87U\x9a\xf52h<f>\x08VN\x0c\xbeX\xdd\x84|\x97\x98/\x8b\x96AH\x10\xc1\xaari\xf8\x97t\xd5%x\xed\xba\xeaT\xa4\xb2]\x85\xa2\xacH\xb2\xd5\x90RN@\xa94\xbf\x1eeba\x90F\xb1\x85.A\xd4\xf3g\x81Z\x06\x16:\x9f\x0e\x1c\x002\xac\x1c\xec.\x01\x11\x06\x06m"\x95\xd3C\xe7\xec\x02N\xdb\x04\x80\x01\x1aZ4\x08\xf5\xd1\x02\xb1s\x04\xa7\x94n\x8c\x99\x85>\xc8h\x15(\xb3;w%dZy\xfc\xef\x84\x1d\xa3\xb7\x15\xe1\x8e\xf5E\xb8kC!\x1e\xdb^\x84Q\xff\x14\xe1\xc9\x1d\xc5X\x93\xed\x12\xdd\x0c\x06\xcf\x06\x00\xc0\xc1\x03\xb3\x93\xec\xd8\x99\xe3\x064\xc0\x88\xfa\x06\x8c\xa8\x17\xd8\xb4\x06\x01\xd8}\xc1\x8d\xb7\xf7\x97b\xc4\x96"\x0c\\W\x80\xfb7\x8b\xe9>\xfcO\x11\xfe\xb3\xa7\x04\xc9\xc5\\\x85Z\x93\xceK\xf1\xc6\x17N \x1c/\xe4\xc0\xf3\xd2\xfb22\xb8X\x9f\\\xd7\xcc\xc2\xac\x15\x05@\t\'\xc8_\x00\tfqz\n,\x83\xe4b\x0eK3\x1c~k\x13\xab\xce8q\xa4@\x94\x18\xcd\xc24\xa8\x17@\xbbK\xe5\xd2\xb9\xc8+\xabR\x04\xab\x8bU\x84\xeaT\x94\xd7\x17;\'v\xc6\x00`\x14gA\x00i\x11\xb3T\xeay#\xf5\x0c|\xd4\xcc!\x90\xf8\x95\x00N\x1c\xa1\xb7\x8cP\n\x80}y\x1cr\xed\x024:\x06\x9d\xa3t0\xf9\x08\x00\xef.\xc5\xa0\xb9\xba\x15jw\x9e\x1b\xe3v\x95\xe0\xb5=%X\x99\xe6\xc0\xf1|\x17\x9c\x02!\xc6\xa4A\x9cY\x83h#\x0bc\x05\xd5\x147\x9es\xe2\xbb\x14;x7\xa1uM\x1d\x9ek\xaa\x9c\xfa\xf1\xc0\x13\xf0K\x9a\x03\x8fn+\xc2\xc7\x87J\xb1\xf5\x8c\x13\xa9\xa5<\xf4\x1a\x06\xb1&\rbM\x1a\xd46h\xa0\r X\x03A\xe5\x0ck\xdc\x04\x14\xba\x059\x06\xc3\\\xbc\x05\xb0\x8c\xa4\xa9C\x80\x9b/\xd3N\xd5\xb1\x0c\xee\xae\xa3GT\x88\x06v\'\xe1\xd3c6\xbc\xb1\xbf\x14\xf3R\x1d\xf8%\xcd\x81\xff\xee/\xc5\x87\x87\xad(\xb5s\x08\xb5\xb0\x18V\xd7\x88\xc8\xcb4\xf4\xa6R}\xb9\x9a\xedU\xe5\x1ab\xe7)\x88\xad\x9e\xb2cEnR\x08U\x1bG\x98\x95d\xc3\xf2,\'\xc0\x00!:\x16F\xaf\x8e,\xa9\x84\xc3\xaf\x19\x0e\xc0Mh\x15\xa9\x0b8u!\xc7&\xa0\xd8U\xfe\xbc\xfb\x95t/\xf9N\x01_\x1e\xb7\xe1\xcf\x0c\x078\x81\xd0-N\x8f)\x1dC\xb1\xa4w86\x0f\x88\xc4\xa6\x01\x91X{[$\x86\xd65\x00DrG\x18\x88\xd36\x01\x9f\x1f\xb7!\xb7\x94\x87\xc5\xc4\xe2\x99\xa6f\xb4\xab\xa1\x14|\x1e\xf6\xe5\xbb\xf1\xde\xa1R\x9c\xcc\xe3`\xd2\xb3\x18\xde\xc8\x84Y\x9d\xc3\xb0\xa2o\x046\x0f\x88\xc4\xe6\x81\x91X\xda\'\x1c\xad"\xb4\x97\xbd\xb6\xe1A\xc70\x88\xd41bII\x1dz\xc0W\xea\x85\x93\x07\xec\xbcxI\x98\x8eQ|\xe1Yy\x02/\xed\xfa:[*\xe0\xb3\xc3V<\xbe\xbd\x08\x8fm/\xc2G\x07K\x91Q\xc8\xa1V\x88\x06\xaf\xb40cXB\xe0\xaf\x1f\x95\x1b\x9b+\x14\x00\x17\xa9\x9d*\xd7\x0c\r\x0by\xf1\xd5\xb37\n\x00"u\xac\xb8hJ\xe2h~w\x9e\x1b\x05.qa\xf8\xcd\xfd\xa5xc_)\xce\xdb\x04@\x03\x9c\xb1\t\xd8~^<\x7f\xa2\x98\xc3\x07\x87m\xd8s\xde\x05F\x07\xdc\x97`\x90w\x14{c\xd62\xf2WA\x8e\x9d\xe0.\xa7\xc7b\xa4\xbdV\x97\xc3\xbe|\x0e\x7f\x9f\x15\xb5\\ZDj\xf1\xf9\xad\xa1x\xb3\x8d\x05\x03b\rh\x10\xa2A\x83\x10\rZ\x84\x8b\x9b\xdb\xca\x83\'\xe0\xa7T;6\x9ds\x03\x0c04\xc1\x80\x91>\xaa\x8f\x1ex\x02V\x9dv"\xa9P\x9c\xde\x19\x9a`\xc0\x8c\xce\xa1\x18\xd7\xcc\x84\xeeQ:4\x08\xd1\xa0\xbeEL\xd7\xac\xbd\xf2\xe6\xa0e\x81faZ0R\xab<g\x17\xe0.Gk\x8a\x00\x1c-\xe2\xe0\x94\x16\xe9\xe3\xcd\xac\x98\x0fiMh\xf6I;\xf2\xed<\xfa\xc5\x19\xf0J\x1b3z\xc7\x19\x10o\x16\xbfZn\xa9\xa5\xc3\xc3MM\xf8\xae{\x18^ke\xb9\xec\r^*\xd5\x9b+\x14\x00j\xa5\xa9*\x84h\x19\x84H\xd6\x1a\x8b\xdc$o\x82j\x14\xaaAkiZ\xe7X!\x87q\xbb\x8a\xf1\xd8\xb6"<\xb0\xa5\x08_\x1c\xb3"L\xc7bx\x03#\xc2\x8d\x1a\x9c)\xe1\xf1\xea\xbf\xa5xl[\x11\x1e\xd9Z\x8c\x9f\x92\xed\x00\x01\xc3\xea\x99\xf0TcS@\xed\x9eZ\x06Q\x1d\x12\x04$\x97pH\x91\xec\x08\xf9R\xec&l;\xef\x86\x8d\xbb\xbc\x8d`\xe7\x9d\x82l\x92\xbaYx\xd93y\xe3\x12\x08%nQ\xfa\x05KaK\x8e\x1b_\'\xd9\xe1v\th\x14\xa1\xc33M\xcdA\xa7>J\xdc\x84cE\xbc8\xb2g\x19\xdc\x11g@L\x80E\xd9\x02\xa7 >W%P/D\x9c\xca\x82 :\xf19Y\x14\xb8<!\xed\xd2^w\xd6\x05\x08\x04\x83\x9eEbm\xbd\xac\xc6y\xbc\x88\xc7\xde<q~\xff\xde\x04#>\xed\x18\x8a\x95}#\xb0\xeb\xce\x1a\xd8yg\r\xac\xbb-\x12?t\x0f\xc7=u\x8c\xe2\x1a\x82\xcaM\x89\x7fm\xbeFTNsQ\xf1P\xd3\xc0\xa2\xa6A\x1cbgY\xcb:\xa4\xdaF\x16\xe3\x9a\x9a\xd1\xbc\xa6\x0e\xe0\tGs\\X\x99j\xc7\xf1B\x0e\x9d\xa3\xf4\x98\xd19\x14\xd3:\x84\xe0\xeex\x03t\x1a %\xdf\x85\x95\xa7\xec\xd8\x9b\xe3\x82I\xc7\xe0\xc1F&Li\x17\x82X\x1f\xb3\n\x1e\x9a\x85I\x9d1\x03\xa4\x95\xf0\xf8&\xc5\x8e"\x9f\rIyN\xc2\x87Gl\xf8.\xd9\x0e^\x9a\xae\xb8Tt,\x03\xbd4\xbdQ\xe4$\x94\xf8n\xf3\x05\xb0,\xcb\x85\r\xe7D\x15\xca@I\xe4:\x04\xccJ\xb2!\xb3\x88\x83\xd9\xc8\xe2\xb5V\xc1M(\x00\x80\x9e\x15UK\x01\xb1\xc2\xfa\xee+\x00\x80"\x17\xe1\xbb\x14;\x92\x8a\xf9JiM-#\xb4hWC\xccSj\t\x8f\xcf\x8e\xdb\x90n\xf5\x7f\xd6R\x8e\xf0c\x8a\x1d\xff\x9ew\x03\x04\xb4\xaf\xa1E\x97Ze\xcf"\x80 Z\x0e\x02\x96g9\xf1\xc7\x19\'J\xa4:a`\x01-\xcb\xa0\xd0-\n\xd5J\x92]*\xd5\x10\xcd\xc4\x89\x13\'\xf9\x1e\xbc\x16\x04j\xa0*\x97\x86\xb7\xe9l\x8b\x96\xc1\xb6\\7\x8e\xe6q\xb0\x03H\xac\xad\x97u\xbf\x1b\x87j\xd02B\x8b\xda\x16\r\x1aE\xe8\xd0\xa9\xb6\x1e\x0f74\xe1\xf5\xd6\x16\xf4\xac\xadC\x84\x9e\xc5\xad5\xb5H\x08\xd1 >T\x8bV\xb5\xf4\xe8\x19\xab\xc7\xd3MLx\xb1\x859\xa0\xed\x1f\x0f!:\x06v\x9e\xf0w\x8e\x1b\x9c\x9b\x90T\xca#\xc7A\xd0\xb2@\x8eC\xc0\xee\x0b\x1c>:b\xc5\xb7)vp\x02\xc1\xa8e\xc0\xf1\x84\xb65u~\xf3\xceG\n9\xac<\xed\x84\x8b\'t\x8f\xd6c@\\\xd9\x9e\x05;\x0f\xac?\xe7D\x9eM@\x9e[\\k\x88\xd4kP\xca\tH.\xe1\xb10\xcd\x81\xf7\x0eY\x91Y\xcc\x01,\x03\xa3\x8e\xc1\xbd\tF\xc5\xb4\xd5O\xa7\x1c\x98~\xc2\x06\x9e\'4\x8b\xd4b@\xac\x01%n\x01\xe9V\x1e\x99RH/\xe5\x91\xef"\xd44\x88\x1b\xa6r\x1d\x02V\x9fuA\xe0\t\xa7\x1d\x02LZQ\xbd2\xcf)`\x7f>\x87O\x8f\xd9\xf0U\x92\x1dvI\xf7?\xce\xa2\xc1\xbd\tF\xc5n\xea\x12\x8e\xb08\xc3\x81\xd3%<\xe2B\xfd\xcf{\x13\xa2e\xa0e\x81\x8d\xb9n\xd8]\x02N\x94\xf0H.\xe6\xa1a\xc4u\x90,\x1b\x8f\xbdyn|v\xdc\x8e\xefR\xec\xb0:x\xd4\xb0h1\xb1m\x88\xb8\'@B\xcb08V\xc4\xe3T\x89\x18\xfe\x94\xd4V\xbfM\xb6cQ\x86\x03\x8b\xd2\x1dX\x90\xe6\xc0\x1f\xa7]8R\xc8A\xc3\x02q&\r\xbcM\xfe\xdby\xc2\xf2,\'N\x16p\x88\xb2h\xf0t\x13\x93,\x84/\x15\xd5\xcfG\xd5\xa4Z\x18\x83S\t\x8c\xaf1\xb8O\x8f\xd9\xf0\xfa\xde\x12\x08<\xe1\xf1\xa6f\xcc\xec\x12\nS\xa0y\x9bJ&\xdfIx}\x7f\t\xbeO\xb6Cp\x93h\xa8\xcc\xc8\xc2\xa2\x03\xce\xdb\t\x0e\xa7\x80\x84\x08-\x1e\xaao\xc4\xeal\'\x0e\x9duad\x0b\x0b\xe6\'*\x8d\xc1\xcdOu\xe0\xa9\x9d\xc5p8\x05\xbc\xd4&\x04\x9f\xddZf\x0c\xce)\x10\xa6\x1d\xb5\xe1\xbdCV8\x9d\x02\x18\x9d\xb8\xb9)L\xc7 \xcfI\xc8(\xe1\xa0\xd7\xb2h\x1a\xa6\xc5\x89b7\xf4\x0c\x83\x05\xbd\xc2qO\x9d2!\xf3\xd8\xb6b\xcc=i\x05\xb4,\xc2\xf4\xa2\xffc\xc1\xe7[\x94\x13\x808\xb3\x06\xd3;\x85\xa2[\x94\x0e\xa7m<\xfe\xb3\xbb\x04\xbfe8\x01\x9e`1\xb1h\x18\xc2\x82e\x18\x9c\xb6\n\xc8\xb3\xf1\xa8e\xd1"R\x0f$\x17ph_[\x87\xf9\x89\xe1\xb2Y\x07H\xc6\xe0\x86n*\xc4\x9e3.\xb4\x8f\xf5?\xef\x8bK\x00>;f\xc3\xa7\xc7J\x91W*\xaa\xe0\x1a\xf4\xa2\x89\x0e\x06@\x81\x8b`\x95v\x1b\xd7\x0e\xd5\xe0\xe5\x16\x16\xbc\xd4\xd2\xdf_\xef\x9al\x17\x9e\xdaQ\x84\xd3%\x92f\x91\xa7*x>(<\x8f\xce2\xa8\x13\xaa\xc1\xeb\xad-\x18\xdb\xc4$o\xf6\xcbw\txl[1V\xa5:\xd0,J\x87\xbd\x83j\x04\xd4\x94\xaa\x08\xaa1\xb8\xaaI\xe0a\x88J\xb5\xe4\x8e8=ZDh\x01\x1e\xf8-\xd3\x89\xaf\x93\xec\x9567]\x1e5\x0c\x0c&\xde\x12\x82\xb7\xdaX\xd0\xbc\xa6\x0e\x06-\x83|\x87\x80\xac\x12\x01f-\x83A\xf5\x8d\xf8\xbak(^ji\x16\x1d\xd2\xb0L@\xc1\xa4\xd7@\xdc\xbc\xc40~\x9d\x99\x81e\xf0l33&\xb7\x0fA\x87h\x1d\x8c\x1a\x06)\x85<\xf6\x9dw\xe3\xacC@\xbb\x9az|\xdc!\x04\xaf\xb64#L\xcbB\xc32\xf0\xed\xab\x8c\x1a\x06`Y\x80\x115\x962Jxd\xf9\x84\xb3\xc5\x1c\xd2Jyy\xe3T\x1d\xb3\x06\x1f\xb4\x0f\xc5\xd8\xe6&\xd4\x8b\xd0\xc2\xc1\x13\x0e_\xe0p0\x9f\x83\x00\xe0\xcezF\xfc\xafS\x88h>\x82a\xa0c\x18\xbf\xb5\x12F\xca?X\xf1\xaf\xefy_\xf4,\xf0RK\x13fw\t\xc3\x9d\xf5\x8d\x88\x0f\xd5\x82\'\xe0L)\x8f\xd3\xa5<\\\x02\xa1n\xb8\x16#\x1a\x1b\xf1u\xb70\xbc\xd4\xd2\xe4W^yN\xc2\xeal\x17\xf2\xec\x84\x10\x03\x83G\x9b\x981\xadS(\xa6v\n\xc5\xdb\xedC\xf1\xea-!x\xae\x8d\x05\xdd\xe2\xf4\xd0j\x81\xd3E\x1c>9j\xc5\x96\xdc2s\x12\x0c\xa4\xcdd\x8c\x98\'\x95\x1b\x0f\xf5\x0b\xa0\x1a\xe3\xfb\x05@\x00f\x9d\xb4\xe1\xff\xf6\x95\xc2f\x17P\xc3\xa2\xc1\xa3\x8d\x8c\x18\xdd\xc8\x88\xa6\xa1Z?\x1d~ox\x02\xf2\x9c\x02x\x12\xd7\x13.\xa7\xc1\xf3\x04\xec\xcf\xe7\x90^\xca\xc3\xce\x13X\x88\x9b\xd2n\x89\xd0"\xc6\xc4\x82\'\xd1\xbf@F)\x8fV\xe1Zt\xaa\xa5\x1c\x05g\xd9\x04\xec\x94\x16\x8a\xdb\xd5\xd0\xa2\xad\xcf\xa64H\xcfx\xac\x88CR\x11/n|\x93L@4\r\xd3\xa0e\xb8\x16\x17\x9c\x82\xbc\x0e\xd0\xb3\xb6\x0e\xb1^\x8b\xb6\xfb\xf39\x1c\x91\xccV\x04\x83\'q\xbfD\xff\x18\xbdbgl)G8\x90\xef\xc6Y\xbb\xa8n\xabe\x19D\x19\x18\xb4\x8a\xd0 \xce\xa4\xc1\xc1\x02\x0e\x07\xf29\xc4\x98X\xf4\xa8\xadSh\xd58x\xc2\xd6\\7N[y\xc4\x9b5\xe8\xees\xbe<\xce\xd9\x05\x9c(\xe6p\xdeApHM\xd5,\xed\xd8\xbe%R\x87\xf0\x00\x0b\xd8\x02\x01\xb3\x92\xecxeO1\\<\xf0TS\x13>\xe9\x10\x8a\xc8\x00\x1aR\xc9\x92\xc5\xd0_S\xed\x00\xc3\xe0\x9dv!x\xb7\xad\x05\x90\x9c\x0b\xed\xba\xc0!\xb9\x98Cm\x13\x8b\x81q\x06?\xa1ZQ\xd4/\x80\xaa\x89*\x00\xaa1\xbe\x02\x00\x92n\xff\'Gm\xf8\xfc\x98\x15\xc56\x01\xd01h\x1e\xa1Ebm\x1d\xba\xd4\xd2\xa1\x8e\x99\x85Y\xc3\xc8\xe6\xef\x9d<\xe1\xbc\x138\x98\xef\xc6\xee\x0bn\x18\xb4\x0c>l\x17\x12T/^\xa5\xeaS\xe8\x120z{1\x96\xa7:\x10\x1d\xa6\xc5\xfc\x1ea\xe8\xefc\xe6\xd9\x9b\xcf\x8e\xd9\xf0\xea\xde\x12\x10\x0f\xbc\xda\xc6\x82\xa9\x01\x1c\xf1\\)\xaa\x00\xa8\x9a\xa8\x02\xa0\x1a\x13H\x00@2\xe067\xd5\x81\xd9\'E/Tp\x8b\xa6\x05\xa0\x05\xc2\xf5,B\xb4\xe24\x04I\x8b\xab\xf9N\x82\xe0\x16\x007\xa1am\x1d\x16\xf5\n\xc7\xad\x92&\x8aJ\xf5#\xc7!\xe0\xc1-E\xd8t\xda\x89\xda\xa1Z\xfc\xd8=\x0cw\xc6\x07\x16\x00\x05.\xc2\x7f\xf6\x94`^\xb2\r\x1a-\x83\xff\xdd\x1a\x86\xe7$op\x95\x89*\x00\xaa&\xaa\x00\xa8\xc6\x04\x13\x00\x1e\x0e\x17rX}\xc6\x89\r\xe7\\8Y$\xba\x03,v{y\xb3\x02\xa0\xd52\x08\xd5\x8a*\xa4\xad#t\x18\\G\x8f\x11\xf5\x8ce\xea\x8f*\xd5\x0e\x1bOxyO)\xbe:n\x054@\xffX\x03\xc641\xa1^\x88\x06\xa1:qm\xc4\xce\x03\x19V\x1e\xbfg:\xb18\xdd\x01\x9b]@\xc7X=\xe6\xf6\x08C\xcbr\x16\xa8/\x17U\x00TMT\x01P\x8d\xb9\x98\x00\xf0\x90\xe3\x10\x90R\xcc\xa3\xc0-\x1a\x8csK\xea\xec\x0cD_\xba\x16-\x83\x1az\x06-\xc2\xb5j\xc7\x7f\x83\xb05\xc7\x8d\xe7\xff-\xc1\xc1\\\x17@\x80\xd6\xc0 \xd6\xc4"\\\xc7B\xcb\x88B"\xcb*\xc0\xee\x10\x00\x96\xc1-5\xb5\x98\xd4\xd6\x82a\t\x81wE_)\xaa\x00\xa8\x9a\xa8\x02\xe0*A\xd7`\xafCE\x05\x80\xca\xcd\x87@\xc0\xf6\x0bn\xfc\x92\xea\xc0\xbfyn\x9c\xb3\x0b(\xe6H\xb4\x08J\x80N\xc3 L\xcb \xc6\xc4"1Z\x8f\x87\xea\x19\xd0\xc9k#Ye\xa3\n\x80\xaa\x89*\x00\xaa1\xaa\x00P\xa9\x08\x0e^4\r\xc2\x13\xe4\xdd\xc1\x0c\xc4u \x1d+\xa9\xc7^eT\x01P5Q\x05@5F\x15\x00*\xd5\x05U\x00TM*\xaf\x07\xa9&b\xa4\x9adSEEE\xe2R\xda\xec\xa5\xc4U\xa9L\x01p\xa9_\x91\xd7\xe9M]j6UTT\xae/\x97\xd2f/%\xaeJe\n\x80KE}S***\xc1\xb8N\x03\xc4\x9b\x8d\xeb\'\x00TTTT\x82\xa1\x0e\x10\xaf\t\xaa\x00PQQQ\xb9IQ\xb5\x80T\xaa\x1c\x9e=\x14\xd9\xd9g\xb0c\xc7N\xfc\xf0\xc3\x0fX\xbdz\xb5o\xb4\x80t\xef\xde\x1d\xcf>;\x0e\xbdz\xf7A\\L\xecey\x1f\xf3\x85\xe7y\x9cL:\x89e\xbf/\xc3\x07\x1f~\x00\xbb\xcd\xee\x1b% \xff}\xf3M\x8c\xb8\xef>4o\xde\x1c\x06\x83\xeasW\xa5\xeaq\xc3\n\x80k\xb1\x11\xeb\xbaS\x95\x1f\xf2\n\xf3F\x00@\x047\xe7\xc6\xe9\xac\xd3X\xb6l\x19f\xcd\x9a\x85\xf4\xf4\xf4\x8b:\x17\x89\x8e\x8e\xc6\xa0A\x83p\xdf}\xf7\xa1g\xaf\x9e0\x9b\xcc\n\xe79\x15\x85\x88p\xfe\xfc\x05\xacY\xb3\x06\x8b\x16-\xac\x90\x10\n\r\rE\xef>\xbd\xf1\xec\xf8g\xd1\xa1cG\xd4\x88\x8c\x14\xd3f\x98+)\x8e+\x82\x88.\xeb\xf9Un|nX\x01P1\xae\xb0\x97\xba\xd1\xa9B\xc5\xe3v\xbb\xb1{\xf7n\xcc\x9d;\x17\xabW\xaf\xc6\x993g|\xa3(`\x18\x06\xf1\xf1\xf1\x18=z4F><\x12\xf5\xeb\xd5\xaf\xd0(\xdc#xJ\xad\xa5\xf8w\xcf\xbf\x985s&6l\xdc\x88\xc2\xc2B\xdf\xa8~\xf4\xee\xdd\x1b\x0f>\xf8 \x06\x0f\xbe\x1b\xd1\xd1\xd1`YM\xd5*DT\xbd\xec\xa8\\_nr\x01\xa0R\xdd\xc8\xc9\xc9\xc1\xd6\xad\xff\xe0\x97_\xe6c\xf9\xf2\xe5\xbe\xa7\x03\xd2\xa5KW\xbc\xf0\xc2\xf3H\xec\x91\x88\xb8\xb8X\xb0\x9a\xe0..\x1dN\x07\x92\x93S\xb0v\xed\x1a\xfc\xdfk\xff\xe7{:(oO\x98\x80{\x87\rC\xcb\x96-\xa0\xd7\x1b.\xfa\x95\xa2\xa2R\x15P\x05\x80J\xb5\x82\x08ps.\xe4\xe6\xe4\xe2\xb7\xa5\xbfa\xe6\xacYHM=u\xd1\x9d\xa6\xb5\xa3j\xe3\xf6\xdbo\xc7#\xa3\x1eA\xef\xde\xbd\xfd\xbe\x06\x04A@~~>V\xadZ\x85\xc5\x8b\x17a\xed\xda\xbf/\xda\x89\x1b\x8dF\x0c\x1e<\x18c\xc7\x8eE\xbbv\xed\x11\x1e\x1e\x06\x96e\xd4!\xb6J\xb5A\x15\x00*\xd7\x98\xca\x99\x83\x10\x88\xc0\xb9\xdd\xd8\xbbw/\xe6\xcf\x9f\x8f\xbf\xfe\xfa\x0biii\xbe\xd1\x140\x0c\x83\xb8\xb88<\xf5\xf4\x18\x8c|\xe8A\xd4\xad[\x07:\x9d\x016\x9b\x15\xfb\xf7\xef\xc7\xc7\x1f\x7f\x8c\xad[\xb7\xa2\xb4\xb4\xd4\xf7R?z\xf5\xea\x85a\xc3\x86a\xf8\xf0\xe1\x88\x89\x8d\x81\x86e+\xe5\xb9TT\xae%\xaa\x00P\xa9\xbe\x88\x13\xf6\xc89\x7f\x1e;w\xee\xc4\x9c\xd9\xb3\xb1v\xedZ\xdfX\x01\xe9\xdb\xb7\x1fF\x8f~\x0c-[\xb6\xc4\x86\xf5\x1b0{\xcel\xa4\xa7\xa7\xfbF\xf3#.>\x0ec\xc7\x8c\xc5\xc0\x81w\xa0]\xfbv\xd0j\xb4\x00\xd4EV\x95\xea\x89*\x00T\xaa\'\x04\x90\xe8\xaf\x1c\x02\tp;]8\x97\x9b\x8b\xe5\xcb\x96a\xe6\xcc\x99HOK\x03\xc7K\x8e\x0f\x82\x10\x1a\x1a\x8a\x8e\x1d;b\xd3\xa6M\xbe\xa7\xfc0\x9b\xcd\xe8\xd6\xad\x1b^{\xed\xff\xd0\xa5KgX\xcc\x16\xb0Z\x8d\x98\x0f"i\xeaGE\xe5ZP9_\xd1\xb8\x11\x04\x80\xaa\xe2\xa6\xe2\x8d\xd3\xe5\xc2\x91C\x87\xb0h\xf1b,Z\xb4\x08\xa7O\x9f\xf6\x8dr\xc9\xdcv\xdbm\x18:t(\xee\x1b1\x025k\xd4P\xad\xb0VI*\xafS\xbc\x99\xa8\xf6\x02@E\xc5\x17" ??\x1f\xdbwl\xc7/\xbf\xfc\x82\xc5\x8b\x16\xf9F\xa90\x93&M\xc2\x90\xa1C\xd0\xacYs\xe8t:\xb0\xea`C\xe5\x06B\x15\x00*7\x1cD\x00\x03\x02\xc7\xf3\xc8>\x9b\x8d\xa5K\x96\xe2\x87\x1f~\xc0\xf1\xe3\xc7/\xaa-\x04\x00!!!\xe8\xd7\xbf?\x9e\x7f\xfe9t\xec\xd0\x01a\xe1\xe1\xb8~\xdb\xb8TT\xae\x1e\xaa\x00P\xb9\xe1qsn\x1c;z\x0c?\xfc\xf8\x03V\xadZ\x85\xb4\xd4\xe0\xdaB\x9e\xcd\\\x83\x06\xdd\x89\xb8\xb8xu\xbaG\xe5\x86F\x15\x00*7\x05\x04BA~\x01\xb6\xef\xd8\x8e\xb9s\xe7b\xe9\x92\xa5\x8a\xf3\x16\x8b\x05o\xbc\xf1\x06\x06\r\x1a\x84\x96-[\xc2`P7s\xa9\xdc\xf8TY\x01ps.\xeeV\xeeBV\xe5\xde\xad\xf2\xb9V\xf9\xf3N\xc7\xedr#\xf7|.\xfe\xfe\xfbo\xcc\x993\x07\xe7\xcf\x9f\xc7\xc0\x81\x031\xea\xd1Qh\xd5\xb2\x15BCC\x15v{\xaeU\x1eUT\xae\x07UV\x00\xa8\xa8\\-\x88\x08\x1c\xc7!%%\x05.\xb7\x1b\r\x1b4DH\x88\xe5&\x1cp\xa8T\x98\x1bt$\xa0\n\x00\x95\x9b\x12\xcf\xf4\x0e\xc3\xb2\xe2\xd7\xa6o\x04\x95J\xe3\x06\xed;o\x08n\xd8\x15.U\xaa\xa9\x94\x0b\xc3\x88\xae\x02\xd4\xce\xff\xaa\xa3\x96o\xd5\xe5\x86\x15\x00j\xa5S)\x0f\xb1~\xa8\xb5D\xe5\xe6\xe6\x86\x15\x00*****\xe5\xa3\n\x80\xeb\x8c:U\xa5r\xe3\xa2\xd6\xee\xaa\xcee\x08\x80\xeb\xfcR\xafs\xf2\x95\x8d:\t\xa1r\xe3\xa2\xd6\xee\xaa\xcee\x08\x80\xeb\xfcR\xafs\xf2****7\n\x97!\x00Tp\xe3}\x88\xa8\xa8\xa8\xdc\x84\xa8\x02\xe02\t\xf6!\xa2\n\x86\xab\x89Z\xba**\x95\x89*\x00*\x99`\x82A\xa52PKWE\xa52Q\x05@\xa5p\x15F\xa6W\xe1\x96**7\x13j\x13\xba8\xaa\x00\xa8\x14\xae\xc2\xc8\xf4*\xdcR\xa5\x1a\xa0\xf6Z\x95\x86\xda\x84.\x8e*\x00TT\xaa\x12j\xaf\xa5r\rQ\x05\xc0%\xa0\xda\x87\x0f\xc0u(\x92\xeb\x90\xa4\xca\r\x86\xda\x96E\xfe\x1fn\xe8E\xb5\xb8e\xb5V\x00\x00\x00\x00IEND\xaeB`\x82'
        
        # Load image from bytes literal
        try:
            # Open the image from bytes
            image = Image.open(io.BytesIO(img_bytes))
            
            # Get image dimensions
            img_width, img_height = image.size
            
            # Add some padding for the loading text
            window_width = img_width
            window_height = img_height + 50  # Extra space for loading text
            
            # Position window in center of screen
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2
            self.root.geometry(f'{window_width}x{window_height}+{x}+{y}')
            
            # Convert to PhotoImage
            photo = ImageTk.PhotoImage(image)
            
            # Create label for the image
            image_label = tk.Label(self.root, image=photo, bg='white')
            image_label.image = photo  # Keep a reference
            image_label.pack()
            
        except Exception as e:
            # Fallback if image loading fails
            print(f"Could not load splash image: {e}")
            self.root.geometry('400x300+{}+{}'.format(
                (screen_width - 400) // 2, 
                (screen_height - 300) // 2
            ))
            error_label = tk.Label(self.root, 
                                 text="Gas Exchange Data Analyzer\n\nLoading...", 
                                 font=("Arial", 16, "bold"),
                                 fg='darkgreen',
                                 justify=tk.CENTER)
            error_label.pack(expand=True)
        
        # Add loading text at the bottom
        loading_frame = tk.Frame(self.root)
        loading_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        
        loading_label = tk.Label(loading_frame, 
                               text="Loading Gas Exchange Data Analyzer...", 
                               font=("Arial", 10))
        loading_label.pack()
        
        # Add a simple progress indicator
        self.progress_label = tk.Label(loading_frame, text="", font=("Arial", 8))
        self.progress_label.pack()
        
        # Start progress animation
        self.progress_dots = 0
        self.animate_progress()
        
        self.root.after(2000, self.close_splash)  # Auto-close after 2 seconds
        self.root.mainloop()
    
    def animate_progress(self):
        """Animate loading dots"""
        self.progress_dots = (self.progress_dots + 1) % 4
        dots = "." * self.progress_dots
        self.progress_label.config(text=dots)
        
        # Continue animation until splash screen closes
        if self.progress_dots < 3:  # Only continue if not about to close
            self.root.after(500, self.animate_progress)
    
    def close_splash(self):
        self.root.destroy()


def make_channel_plot(df, yvar, title, ylab, palette='tab10', highlight_darkness=False, darkness_periods=None):
    # Add this debug print
    print(f"\n--- make_channel_plot called for {yvar} ---")
    print(f"df is None: {df is None}")
    if df is not None:
        print(f"df shape: {df.shape}")
        print(f"df empty: {df.empty}")
    print(f"yvar in columns: {yvar in df.columns if df is not None else False}")
    
    # Close any existing figure to prevent accumulation
    plt.close('all')
    
    if df is None or len(df) == 0 or yvar not in df.columns:
        print(">>> CONDITION TRIGGERED: Showing 'No data available'")
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.text(0.5, 0.5, 'No data available', ha='center', va='center', fontsize=12)
        ax.set_axis_off()
        return fig
        
    print(">>> CONDITION NOT TRIGGERED: Plotting data")
    if highlight_darkness and darkness_periods:
        return make_light_dark_plot(df, yvar, title, ylab, palette, darkness_periods)
    
    fig, ax = plt.subplots(figsize=(12, 6))
    
    # Get unit for y-axis label
    unit = UNITS_MAP.get(yvar, '')
    if unit:
        ylabel = f'{ylab} ({unit})'
    else:
        ylabel = ylab
    
    # Get unique channels
    channels = df['Channel'].unique()
    num_channels = len(channels)
    
    # Get colors based on palette
    colors = []
    try:
        if palette in colorblind_friendly_palettes:
            if isinstance(colorblind_friendly_palettes[palette], list):
                color_list = colorblind_friendly_palettes[palette]
                colors = [color_list[i % len(color_list)] for i in range(num_channels)]
            else:
                cmap = colorblind_friendly_palettes[palette]
                colors = cmap(np.linspace(0, 1, num_channels))
        elif palette in plt.colormaps():
            cmap = plt.get_cmap(palette)
            colors = cmap(np.linspace(0, 1, num_channels))
        else:
            colors = plt.cm.tab10(np.linspace(0, 1, num_channels))
    except Exception as e:
        colors = plt.cm.tab10(np.linspace(0, 1, num_channels))
    
    # Ensure we have enough colors
    if len(colors) < num_channels:
        colors = list(colors) * (num_channels // len(colors) + 1)
        colors = colors[:num_channels]
    
    # Track min and max time for x-axis limits
    all_times = []
    
    # Plot each channel
    for idx, channel in enumerate(channels):
        channel_data = df[df['Channel'] == channel]
        color_idx = idx % len(colors)
        
        # Sort by time_repeated for proper plotting
        channel_data = channel_data.sort_values('time_repeated')
        
        # Collect times for this channel - ensure we use time_repeated
        times = channel_data['time_repeated'].tolist()
        all_times.extend(times)
        
        # Verify time_repeated column exists
        if 'time_repeated' not in channel_data.columns:
            print(f"WARNING: 'time_repeated' column missing for channel {channel}")
            continue
            
        # Verify yvar column exists
        if yvar not in channel_data.columns:
            print(f"WARNING: '{yvar}' column missing for channel {channel}")
            continue
        
        # Plot the data with time_repeated on x-axis
        ax.plot(channel_data['time_repeated'], channel_data[yvar], 
                label=str(channel), color=colors[color_idx], marker='o', 
                linestyle='-', linewidth=2, markersize=5)
    
    ax.set_title(title, fontsize=14, fontweight='bold')
    ax.set_xlabel('Time (minutes)', fontsize=12)
    ax.set_ylabel(ylabel, fontsize=12)
    
    # Set x-axis limits with special handling for x=0
    if all_times:
        min_time = min(all_times)
        max_time = max(all_times)
        
        # Determine if we should include x=0 in the plot
        include_zero = False
        # Check if zero is within or near the data range
        if min_time <= 0 <= max_time:
            include_zero = True
        elif min_time > 0 and min_time < (max_time - min_time) * 0.1:
            # If data starts just above zero, include zero
            include_zero = True
        elif max_time < 0 and abs(max_time) < abs(min_time) * 0.1:
            # If data ends just below zero, include zero
            include_zero = True
        
        # Adjust x-axis limits to include zero if appropriate
        if include_zero:
            x_min = min(min_time, 0)
            x_max = max(max_time, 0)
            # Add small padding
            padding = (x_max - x_min) * 0.02 if x_max > x_min else 0.5
            ax.set_xlim(x_min - padding, x_max + padding)
        else:
            # Use data limits with padding
            padding = (max_time - min_time) * 0.02 if max_time > min_time else 0.5
            ax.set_xlim(min_time - padding, max_time + padding)
        
        # Add vertical line at x=0 if it's within the plot range
        if include_zero:
            ax.axvline(x=0, color='gray', linestyle='--', alpha=0.5, linewidth=0.8)
    
    # Adjust legend
    if num_channels <= 10:
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=10)
    else:
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=8, ncol=2)
    
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.tick_params(axis='both', which='major', labelsize=10)
    
    # Ensure x=0 tick is displayed if within range
    if all_times:
        xlim = ax.get_xlim()
        if xlim[0] <= 0 <= xlim[1]:
            # Get current ticks
            ticks = list(ax.get_xticks())
            # Add 0 if not already in ticks
            if 0 not in ticks:
                ticks.append(0)
                ticks.sort()
                ax.set_xticks(ticks)
    
    fig.tight_layout()
    
    return fig

def make_dual_channel_plot(df, col_a, col_b, title=None, ylab=None, 
                           series_names=None, palette='tab10'):
    # Close any existing figure to prevent accumulation
    plt.close('all')
    
    if df is None or len(df) == 0 or not all(col in df.columns for col in [col_a, col_b]):
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.text(0.5, 0.5, 'No data available', ha='center', va='center', fontsize=12)
        ax.set_axis_off()
        return fig
    
    fig, ax = plt.subplots(figsize=(12, 6))
    
    # Get unit for y-axis label
    unit = UNITS_MAP.get(col_a, '')
    if unit:
        ylabel = f'{ylab} ({unit})' if ylab else f'Value ({unit})'
    else:
        ylabel = ylab if ylab else 'Value'
    
    # Get unique channels
    channels = df['Channel'].unique()
    num_channels = len(channels)
    
    # Get colors based on palette
    colors = []
    try:
        if palette in colorblind_friendly_palettes:
            if isinstance(colorblind_friendly_palettes[palette], list):
                color_list = colorblind_friendly_palettes[palette]
                colors = [color_list[i % len(color_list)] for i in range(num_channels)]
            else:
                cmap = colorblind_friendly_palettes[palette]
                colors = cmap(np.linspace(0, 1, num_channels))
        elif palette in plt.colormaps():
            cmap = plt.get_cmap(palette)
            colors = cmap(np.linspace(0, 1, num_channels))
        else:
            colors = plt.cm.tab10(np.linspace(0, 1, num_channels))
    except Exception as e:
        colors = plt.cm.tab10(np.linspace(0, 1, num_channels))
    
    # Ensure we have enough colors
    if len(colors) < num_channels:
        colors = list(colors) * (num_channels // len(colors) + 1)
        colors = colors[:num_channels]
    
    # Plot both series
    linestyles = ['-', '--']
    markers = ['o', 's']
    linewidths = [2, 1.5]
    if series_names is None:
        series_names = [col_a, col_b]
    
    # Track min and max time for x-axis limits
    all_times = []
    
    for idx, channel in enumerate(channels):
        channel_data = df[df['Channel'] == channel]
        color_idx = idx % len(colors)
        
        # Sort by time_repeated for proper plotting
        channel_data = channel_data.sort_values('time_repeated')
        
        # Verify time_repeated column exists
        if 'time_repeated' not in channel_data.columns:
            print(f"WARNING: 'time_repeated' column missing for channel {channel}")
            continue
        
        # Collect times - ensure we use time_repeated
        times = channel_data['time_repeated'].tolist()
        all_times.extend(times)
        
        # Plot first series with time_repeated on x-axis
        if col_a in channel_data.columns:
            ax.plot(channel_data['time_repeated'], channel_data[col_a],
                    label=f'{channel} - {series_names[0]}', 
                    color=colors[color_idx], marker=markers[0], 
                    linestyle=linestyles[0], linewidth=linewidths[0], markersize=5)
        
        # Plot second series with time_repeated on x-axis
        if col_b in channel_data.columns:
            ax.plot(channel_data['time_repeated'], channel_data[col_b],
                    label=f'{channel} - {series_names[1]}', 
                    color=colors[color_idx], marker=markers[1], 
                    linestyle=linestyles[1], linewidth=linewidths[1], 
                    markersize=5, alpha=0.8)
    
    plot_title = title if title else f'{col_a} vs {col_b}'
    ax.set_title(plot_title, fontsize=14, fontweight='bold')
    ax.set_xlabel('Time (minutes)', fontsize=12)
    ax.set_ylabel(ylabel, fontsize=12)
    
    # Set x-axis limits with special handling for x=0
    if all_times:
        min_time = min(all_times)
        max_time = max(all_times)
        
        # Determine if we should include x=0 in the plot
        include_zero = False
        # Check if zero is within or near the data range
        if min_time <= 0 <= max_time:
            include_zero = True
        elif min_time > 0 and min_time < (max_time - min_time) * 0.1:
            # If data starts just above zero, include zero
            include_zero = True
        elif max_time < 0 and abs(max_time) < abs(min_time) * 0.1:
            # If data ends just below zero, include zero
            include_zero = True
        
        # Adjust x-axis limits to include zero if appropriate
        if include_zero:
            x_min = min(min_time, 0)
            x_max = max(max_time, 0)
            # Add small padding
            padding = (x_max - x_min) * 0.02 if x_max > x_min else 0.5
            ax.set_xlim(x_min - padding, x_max + padding)
        else:
            # Use data limits with padding
            padding = (max_time - min_time) * 0.02 if max_time > min_time else 0.5
            ax.set_xlim(min_time - padding, max_time + padding)
        
        # Add vertical line at x=0 if it's within the plot range
        if include_zero:
            ax.axvline(x=0, color='gray', linestyle='--', alpha=0.5, linewidth=0.8)
    
    # Adjust legend
    if num_channels <= 8:
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=10)
    elif num_channels <= 12:
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=9, ncol=2)
    else:
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=8, ncol=2)
    
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.tick_params(axis='both', which='major', labelsize=10)
    
    # Ensure x=0 tick is displayed if within range
    if all_times:
        xlim = ax.get_xlim()
        if xlim[0] <= 0 <= xlim[1]:
            # Get current ticks
            ticks = list(ax.get_xticks())
            # Add 0 if not already in ticks
            if 0 not in ticks:
                ticks.append(0)
                ticks.sort()
                ax.set_xticks(ticks)
    
    fig.tight_layout()
    
    return fig

def create_WUEi_boxplot(df, palette='tab10'):
    # Close any existing figure to prevent accumulation
    plt.close('all')
    
    if df is None or len(df) == 0 or 'WUEi' not in df.columns:
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.text(0.5, 0.5, 'No data available', ha='center', va='center', fontsize=12)
        ax.set_axis_off()
        return fig
    
    fig, ax = plt.subplots(figsize=(12, 6))
    # Prepare data for boxplot
    channels = df['Channel'].unique()
    wuei_data = []
    channel_labels = []
    
    for channel in channels:
        channel_data = df[df['Channel'] == channel]
        wuei_values = channel_data['WUEi'].dropna()
        if len(wuei_values) > 0:
            wuei_data.append(wuei_values)
            channel_labels.append(str(channel))
    
    if not wuei_data:
        ax.text(0.5, 0.5, 'No WUEi data available', ha='center', va='center', fontsize=12)
        return fig
    
    # Create boxplot - NO TIME AXIS FOR BOXPLOTS (distribution by channel)
    positions = range(1, len(wuei_data) + 1)
    box_plot = ax.boxplot(wuei_data, positions=positions, patch_artist=True)
    
    # Apply colors
    colors = []
    if palette in plt.colormaps():
        cmap = plt.get_cmap(palette)
        colors = cmap(np.linspace(0, 1, len(wuei_data)))
    else:
        colors = plt.cm.tab10(np.linspace(0, 1, len(wuei_data)))
    
    for patch, color in zip(box_plot['boxes'], colors):
        patch.set_facecolor(color)
        patch.set_alpha(0.6)
    
    # Add individual data points
    for i, data in enumerate(wuei_data):
        x = np.random.normal(i + 1, 0.04, size=len(data))
        ax.scatter(x, data, alpha=0.6, color=colors[i], s=20)
    
    ax.set_title('WUEi Distribution by Channel', fontsize=14, fontweight='bold')
    ax.set_xlabel('Channel', fontsize=12)  # NOT Time - this is distribution plot
    ax.set_ylabel('WUEi (µmol CO₂ / mol H₂O)', fontsize=12)
    ax.set_xticks(positions)
    ax.set_xticklabels(channel_labels, rotation=45, ha='right')
    ax.grid(True, alpha=0.3, linestyle='--')
    
    return fig

def create_WUEi_comparison_plot(df, palette='tab10'):
    """Create layered violin plot (background) and boxplot (foreground) for WUEi"""
    # Close any existing figure to prevent accumulation
    plt.close('all')
    
    if df is None or len(df) == 0 or 'WUEi' not in df.columns:
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.text(0.5, 0.5, 'No data available', ha='center', va='center', fontsize=12)
        ax.set_axis_off()
        return fig
    
    fig, ax = plt.subplots(figsize=(12, 6))
    
    # Prepare data for plots
    channels = df['Channel'].unique()
    wuei_data = []
    channel_labels = []
    
    for channel in channels:
        channel_data = df[df['Channel'] == channel]
        wuei_values = channel_data['WUEi'].dropna()
        if len(wuei_values) > 0:
            wuei_data.append(wuei_values)
            channel_labels.append(str(channel))
    
    if not wuei_data:
        ax.text(0.5, 0.5, 'No WUEi data available', ha='center', va='center', fontsize=12)
        return fig
    
    # Get colors
    colors = []
    if palette in plt.colormaps():
        cmap = plt.get_cmap(palette)
        colors = cmap(np.linspace(0, 1, len(wuei_data)))
    else:
        colors = plt.cm.tab10(np.linspace(0, 1, len(wuei_data)))
    
    positions = range(1, len(wuei_data) + 1)
    
    # --- LAYER 1: VIOLIN PLOT (background) ---
    # Create violin plot with low alpha for background
    violin_parts = ax.violinplot(wuei_data, positions=positions, 
                                 showmeans=False, showmedians=False, showextrema=False)
    
    # Customize violin plot as background layer
    for i, body in enumerate(violin_parts['bodies']):
        body.set_facecolor(colors[i])
        body.set_alpha(0.3)  # Very transparent for background
        body.set_edgecolor(colors[i])
        body.set_linewidth(1)
        body.set_zorder(1)  # Low z-order = background
    
    # --- LAYER 2: BOXPLOT (foreground) ---
    # Create boxplot with higher opacity
    box_plot = ax.boxplot(wuei_data, positions=positions, patch_artist=True)
    
    # Customize boxplot boxes
    for i, (box, color) in enumerate(zip(box_plot['boxes'], colors)):
        box.set_facecolor(color)
        box.set_alpha(0.7)  # More opaque for foreground
        box.set_edgecolor('black')
        box.set_linewidth(1.5)
        box.set_zorder(3)  # Higher z-order = foreground
    
    # Customize boxplot medians (red line)
    for median in box_plot['medians']:
        median.set_color('red')
        median.set_linewidth(2.5)
        median.set_zorder(4)  # Even higher for medians
    
    # Customize whiskers and caps
    for whisker in box_plot['whiskers']:
        whisker.set(color='black', linewidth=1.5, zorder=2)
    for cap in box_plot['caps']:
        cap.set(color='black', linewidth=1.5, zorder=2)
    
    # --- LAYER 3: DATA POINTS (top layer) ---
    for i, data in enumerate(wuei_data):
        # Add jitter to x-position for better visibility
        x_jitter = np.random.normal(positions[i], 0.08, size=len(data))
        ax.scatter(x_jitter, data, alpha=0.6, color=colors[i], s=40, 
                  edgecolor='black', linewidth=0.8, zorder=5, label=f'Channel {channel_labels[i]}')
    
    # --- LAYER 4: MEAN MARKERS (optional top layer) ---
    for i, data in enumerate(wuei_data):
        mean_val = np.mean(data)
        ax.scatter(positions[i], mean_val, color='yellow', s=100, 
                  edgecolor='black', linewidth=1.5, zorder=6, marker='D')
    
    # Configure plot appearance
    ax.set_title('WUEi Distribution: Violin + Boxplot', 
                fontsize=16, fontweight='bold', pad=20)
    ax.set_xlabel('Channel', fontsize=12)  # NOT Time - this is distribution plot
    ax.set_ylabel('WUEi (µmol CO₂ / mol H₂O)', fontsize=12)
    ax.set_xticks(positions)
    ax.set_xticklabels(channel_labels, rotation=45, ha='right', fontsize=10)
    ax.grid(True, alpha=0.2, linestyle='--', axis='y', zorder=0)
    
    # Add statistics annotations
    for i, data in enumerate(wuei_data):
        median = np.median(data)
        mean = np.mean(data)
        q1 = np.percentile(data, 25)
        q3 = np.percentile(data, 75)
        
        # Annotate median inside the box
        ax.text(positions[i], median, f'{median:.1f}', 
                ha='center', va='center', fontsize=9, fontweight='bold',
                bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.8, edgecolor='none'),
                zorder=7)
        
        # Optional: add quartile annotations
        ax.text(positions[i], q1, f'Q1', ha='center', va='top', fontsize=8, alpha=0.7)
        ax.text(positions[i], q3, f'Q3', ha='center', va='bottom', fontsize=8, alpha=0.7)
    
    # Create custom legend
    
    legend_elements = [
        Patch(facecolor='gray', alpha=0.3, edgecolor='gray', 
              label='Violin (distribution shape)'),
        Rectangle((0, 0), 1, 1, facecolor='gray', alpha=0.7, edgecolor='black',
                 label='Box (quartiles)'),
        Line2D([0], [0], color='red', linewidth=2.5, label='Median'),
        Line2D([0], [0], marker='o', color='w', markerfacecolor='gray', 
               markersize=8, label='Data points'),
        Line2D([0], [0], marker='D', color='w', markerfacecolor='yellow', 
               markersize=8, label='Mean'),
    ]
    
    ax.legend(handles=legend_elements, loc='upper left', 
          bbox_to_anchor=(1.02, 1),  # Position to the right of the plot
          fontsize=9, framealpha=0.9, fancybox=True,
          borderaxespad=0.5)  # Add some padding
    
    plt.tight_layout()
    return fig


def make_light_dark_plot(df, yvar, title, ylab, palette='tab10', darkness_periods=None):
    """Create plot with yellow and black rectangles for light-dark periods with proper subscripts"""
    # Close any existing figure to prevent accumulation
    plt.close('all')
    
    if df is None or len(df) == 0 or yvar not in df.columns:
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.text(0.5, 0.5, 'No data available', ha='center', va='center', fontsize=12)
        ax.set_axis_off()
        return fig
    
    fig, ax = plt.subplots(figsize=(12, 6))
    
    # Get unit for y-axis label
    unit = UNITS_MAP.get(yvar, '')
    
    # Create y-axis label with proper formatting
    if unit:
        # Handle special formatting for units with subscripts
        if 'CO₂' in ylab:
            ylabel = f'{ylab} ({unit})'
        elif 'O₃' in ylab:
            ylabel = f'{ylab} ({unit})'
        elif 'H₂O' in ylab:
            ylabel = f'{ylab} ({unit})'
        else:
            ylabel = f'{ylab} ({unit})'
    else:
        ylabel = ylab
    
    # Get unique channels
    channels = df['Channel'].unique()
    num_channels = len(channels)

    # Get colors based on palette
    colors = []
    try:
        if palette in colorblind_friendly_palettes:
            if isinstance(colorblind_friendly_palettes[palette], list):
                color_list = colorblind_friendly_palettes[palette]
                colors = [color_list[i % len(color_list)] for i in range(num_channels)]
            else:
                cmap = colorblind_friendly_palettes[palette]
                colors = cmap(np.linspace(0, 1, num_channels))
        elif palette in plt.colormaps():
            cmap = plt.get_cmap(palette)
            colors = cmap(np.linspace(0, 1, num_channels))
        else:
            colors = plt.cm.tab10(np.linspace(0, 1, num_channels))
    except Exception as e:
        colors = plt.cm.tab10(np.linspace(0, 1, num_channels))
    
    # Ensure we have enough colors
    if len(colors) < num_channels:
        colors = list(colors) * (num_channels // len(colors) + 1)
        colors = colors[:num_channels]
    
    # Calculate typical time interval between measurements
    all_times = []
    for channel in channels:
        channel_data = df[df['Channel'] == channel]
        if len(channel_data) > 0:
            # Sort by time_repeated
            channel_data = channel_data.sort_values('time_repeated')
            times = channel_data['time_repeated'].unique()
            all_times.extend(times.tolist())
    
    if len(all_times) > 1:
        sorted_times = np.sort(np.unique(all_times))
        time_intervals = np.diff(sorted_times)
        # Get the minimum positive interval
        positive_intervals = time_intervals[time_intervals > 0]
        if len(positive_intervals) > 0:
            typical_interval = float(np.min(positive_intervals))
        else:
            typical_interval = 1.0
    else:
        typical_interval = 1.0
    
    # Calculate y-range for rectangle positioning
    all_y_vals = []
    for channel in channels:
        channel_data = df[df['Channel'] == channel]
        if yvar in channel_data.columns:
            valid_vals = channel_data[yvar].dropna().values
            all_y_vals.extend(valid_vals)
    
    if len(all_y_vals) > 0:
        y_min = np.nanmin(all_y_vals)
        y_max = np.nanmax(all_y_vals)
        y_range = y_max - y_min if y_max != y_min else 1
    else:
        y_min = 0
        y_max = 1
        y_range = 1
    
    # Handle case where all values are the same
    if len(np.unique(all_y_vals)) == 1:
        single_val = np.unique(all_y_vals)[0]
        y_min = single_val - 0.1 * abs(single_val) - 0.01
        y_max = single_val + 0.1 * abs(single_val) + 0.01
        if y_min == y_max:
            y_min = -0.1
            y_max = 0.1
        y_range = y_max - y_min
    
    # Calculate rectangle positioning - IMPORTANT: Handle negative times
    start_y = y_max + (y_range * 0.05)  # Top of data + 5% margin
    
    # Rectangle height calculations
    darkness_band_height = y_range * 0.08  # 8% of y-range for all rectangles
    bar_height = darkness_band_height / max(1, num_channels)  # Height per channel
    
    # Store per-channel data
    channel_data_dict = {}
    
    # FIRST: Plot the data
    for idx, channel in enumerate(channels):
        channel_data = df[df['Channel'] == channel]
        color_idx = idx % len(colors)
        
        # Sort by time_repeated for proper plotting
        channel_data = channel_data.sort_values('time_repeated')
        
        # Store channel data (note: idx is reversed like R's rev_indices)
        rev_idx = num_channels - idx  # Reverse index like R
        channel_str = str(channel)
        
        channel_data_dict[channel_str] = {
            'data': channel_data,
            'color': colors[color_idx],
            'rev_idx': rev_idx,  # Store reversed index
            'y_values': channel_data[yvar].values if yvar in channel_data.columns else [],
            'times': channel_data['time_repeated'].values,
            'x_min': channel_data['time_repeated'].min() if len(channel_data) > 0 else 0,
            'x_max': channel_data['time_repeated'].max() if len(channel_data) > 0 else 0,
            'sorted_times': np.sort(channel_data['time_repeated'].unique())
        }
        
        # Plot the data
        if yvar in channel_data.columns:
            ax.plot(channel_data['time_repeated'], channel_data[yvar], 
                    label=channel_str, color=colors[color_idx], marker='o', 
                    linestyle='-', linewidth=2, markersize=5, zorder=10)
    
    # SECOND: Add background rectangles - ONE PER CHANNEL
    # First, create yellow background rectangles for all channels
    for channel_str, channel_info in channel_data_dict.items():
        rev_idx = channel_info['rev_idx']
        x_min = channel_info['x_min']
        x_max = channel_info['x_max']
        
        # Calculate y position for this channel
        y_min_rect = start_y + (rev_idx - 1) * bar_height
        y_max_rect = start_y + rev_idx * bar_height
        
        # Get the sorted times for this channel
        sorted_times = channel_info['sorted_times']
        if len(sorted_times) > 1:
            # Calculate the actual span including gaps
            rect_start = x_min
            rect_end = x_max
            
            # Extend to show typical interval beyond last point
            if rect_end > rect_start:
                rect_end = rect_end + typical_interval
            else:
                rect_end = rect_start + typical_interval
        else:
            # Single point
            rect_start = x_min
            rect_end = x_min + typical_interval
        
        # Create YELLOW background rectangle for light periods
        yellow_rect = Rectangle(
            (rect_start, y_min_rect),
            (rect_end - rect_start),
            (y_max_rect - y_min_rect),
            facecolor='lightyellow',
            edgecolor='none',
            linewidth=0,
            alpha=1.0,
            zorder=5
        )
        ax.add_patch(yellow_rect)
    
    # Third, add black rectangles for darkness periods
    if darkness_periods:
        for channel_str, channel_info in channel_data_dict.items():
            rev_idx = channel_info['rev_idx']
            
            if channel_str in darkness_periods:
                dark_times = darkness_periods[channel_str]
                if dark_times:
                    # Sort and deduplicate dark times
                    sorted_dark_times = np.sort(np.unique(dark_times))
                    
                    # Group consecutive dark times
                    if len(sorted_dark_times) > 1:
                        # Calculate differences between consecutive times
                        time_diffs = np.diff(sorted_dark_times)
                        
                        # Group if difference is more than 1.5x the typical interval
                        groups = np.zeros(len(sorted_dark_times), dtype=int)
                        groups[0] = 1
                        for i, diff in enumerate(time_diffs):
                            if diff > (typical_interval * 1.5):
                                groups[i+1] = groups[i] + 1
                            else:
                                groups[i+1] = groups[i]
                        
                        # Split into groups
                        zero_groups = []
                        current_group = []
                        current_group_id = groups[0]
                        
                        for i, time in enumerate(sorted_dark_times):
                            if groups[i] == current_group_id:
                                current_group.append(time)
                            else:
                                zero_groups.append(current_group)
                                current_group = [time]
                                current_group_id = groups[i]
                        
                        if current_group:
                            zero_groups.append(current_group)
                    else:
                        zero_groups = [sorted_dark_times]
                    
                    # Create black rectangles for each group
                    for group in zero_groups:
                        if len(group) > 0:
                            # Calculate the span of the group
                            group_min = min(group)
                            group_max = max(group)
                            
                            # Calculate rectangle boundaries
                            if len(group) == 1:
                                # Single point: rectangle starts at the point, extends one interval to right
                                group_start = group_min
                                group_end = group_min + typical_interval
                            else:
                                # Check if points are consecutive with typical interval
                                diffs = np.diff(group)
                                if np.all(np.abs(diffs - typical_interval) < (typical_interval * 0.1)):
                                    # Consecutive measurements with regular spacing
                                    group_start = group_min
                                    group_end = group_max + typical_interval
                                else:
                                    # Non-consecutive or irregular spacing
                                    group_start = group_min
                                    group_end = group_max
                                    if group_end > group_start:
                                        group_end = group_end + typical_interval
                            
                            # Calculate y position for this channel
                            y_min_rect = start_y + (rev_idx - 1) * bar_height
                            y_max_rect = start_y + rev_idx * bar_height
                            
                            # Create black rectangle
                            black_rect = Rectangle(
                                (group_start, y_min_rect),
                                (group_end - group_start),
                                (y_max_rect - y_min_rect),
                                facecolor='black',
                                edgecolor='none',
                                linewidth=0,
                                alpha=1.0,
                                zorder=6
                            )
                            ax.add_patch(black_rect)
    
    # Set title with proper formatting
    ax.set_title(title, fontsize=14, fontweight='bold')
    ax.set_xlabel('Time (minutes)', fontsize=12)
    ax.set_ylabel(ylabel, fontsize=12)
    
    # Adjust legend
    if num_channels <= 8:
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=10)
    else:
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=8, ncol=2)
    
    ax.grid(True, alpha=0.3, linestyle='--', zorder=1)
    ax.tick_params(axis='both', which='major', labelsize=10)
    
    # Expand limits to show rectangles
    y_top = start_y + darkness_band_height + (y_range * 0.05)
    ax.set_ylim(y_min - (y_range * 0.05), y_top)
    
    # Expand x limits to show all data and rectangles
    x_min_all = min(all_times) if all_times else 0
    x_max_all = max(all_times) if all_times else 1
    
    # Add padding to x-axis
    x_padding = max(abs(x_max_all - x_min_all) * 0.02, typical_interval * 0.5)
    ax.set_xlim(x_min_all - x_padding, x_max_all + typical_interval + x_padding)
    
    fig.tight_layout()
    
    return fig
  

# ============================================
# MAIN APPLICATION CLASS
# ============================================


class GasExchangeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gas Exchange Jyrkki Data Processing David Lazaro-Gimeno")
    
        # Get screen dimensions
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
    
        # Calculate position to center the window
        window_width = 1600
        window_height = 950
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 4  # Position at 1/4 from top (higher than center)
    
        # Set geometry with position
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
        # Rest of your initialization code...
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Data storage
        self.raw_data = None
        self.base_data = None
        self.original_calculated_data = None
        self.modified_calculated_data = None
        self.darkness_data = None
        self.radiation_data = None

        # Parameter storage
        self.channel_names = {}
        self.leaf_area_params = {}
        self.air_flow_params = {}
        self.radiation_params = {}
        self.boundary_params = {}
        self.cutic_params = {}

        # Darkness and radiation settings
        self.darkness_settings = {}
        self.radiation_settings = {}
        self.darkness_applied = False
        self.radiation_applied = False

        # Track which parameters have been applied
        self.parameters_applied = {
            'channel_names': False,
            'leaf_area': False,
            'air_flow': False,
            'radiation': False,
            'boundary': False,
            'cutic': False
        }

        # Initialize matplotlib figure management
        self.clear_old_figures()

        # Plot storage
        self.current_plots = {}
        self.plot_tabs = {}
        self.selected_channels = []

        # Initialize darkness_widgets dictionary
        self.darkness_widgets = {}

        # Initialize radiation_widgets dictionary
        self.radiation_widgets = {}
    
        self.batch_metric_configs = {}
    
        # Batch processing attributes
        self.batch_file_paths = []
        self.batch_stats = None
        self.batch_all_processed_data = None
        self.batch_channel_mapping = {}
        self.batch_selected_channels = []
        self.batch_channel_vars = {}
        self.batch_metric_frames = {}
    
        # NEW: Batch metric plot canvas references for proper scrolling
        self.batch_metric_canvases = {}
        self.batch_metric_canvas_windows = {}
    
        # Store initial entry values for change detection
        self.initial_entry_values = {
            'remove_start': "0",
            'remove_end': "0",
            'time0_shift': "0"
        }
    
        # Store initial slider values
        self.initial_slider_values = {
            'remove_start': "0",
            'remove_end': "0",
            'time0_shift': "0"
        }
    
        # Store column widths for table persistence
        self._column_widths = {}
    
        # Sort order for table sorting
        self._sort_order = {}

        # Set up UI (THIS CREATES sidebar_frame)
        self.setup_ui()

        # NOW add batch processing UI (after sidebar_frame exists)
        self.add_batch_processing_ui()

        # Initialize UI elements with default states
        self.initialize_ui_states()
        
    def on_closing(self):
        """Clean up resources before closing"""
        print("Cleaning up resources...")
        
        # Close all matplotlib figures
        plt.close('all')
        
        # Destroy all widgets
        self.root.destroy()
        
        # Exit the application
        if sys.platform == "win32":
            # On Windows, we might need to force exit
            os._exit(0)
        
    def initialize_ui_states(self):
        """Initialize UI elements to default states"""
        # Hide processing frames until data is loaded
        if hasattr(self, 'processing_frame'):
            self.processing_frame.grid_remove()
        if hasattr(self, 'time_range_frame'):
            self.time_range_frame.grid_remove()
        if hasattr(self, 'color_frame'):
            self.color_frame.grid_remove()
        if hasattr(self, 'coef_frame'):
            self.coef_frame.grid_remove()
        if hasattr(self, 'color_coef_frame'):
            self.color_coef_frame.grid_remove()
    
        # Initialize time range text
        if hasattr(self, 'time_range_text'):
            self.time_range_text.delete(1.0, tk.END)
            self.time_range_text.insert(1.0, "Processing...\n\n")
        
    def setup_ui(self):
        # Create main container
        main_container = ttk.Frame(self.root, padding="10")
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_container.columnconfigure(1, weight=1)
        main_container.rowconfigure(0, weight=1)
        
        # Create sidebar
        sidebar = ttk.Frame(main_container, width=400)
        sidebar.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W), padx=(0, 10))
        
        # Create main content area
        content = ttk.Frame(main_container)
        content.grid(row=0, column=1, sticky=(tk.N, tk.S, tk.E, tk.W))
        
        # Setup sidebar
        self.setup_sidebar(sidebar)
        
        # Setup notebook for tabs
        self.setup_notebook(content)

    def adjust_slider(self, slider_var, delta):
        """Adjust slider value by delta"""
        current = slider_var.get()
        max_val = 15 if slider_var == self.time0_shift_var else 10
        new_val = max(0, min(max_val, current + delta))
        slider_var.set(new_val)
        
        # Enable Apply Changes button instead of auto-updating
        if hasattr(self, 'apply_changes_btn'):
            self.apply_changes_btn.config(state='normal')
            self.slider_info_label.config(text="Changes pending - click 'Apply Changes'", 
                                         foreground="orange")
        
    def setup_sidebar(self, parent):
        # Create a canvas with vertical scrollbar for the sidebar
        sidebar_canvas = tk.Canvas(parent, width=450)
        sidebar_canvas.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W))

        # Create scrollbar
        sidebar_scrollbar = ttk.Scrollbar(parent, orient="vertical", command=sidebar_canvas.yview)
        sidebar_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # Create scrollable frame inside canvas
        self.sidebar_frame = ttk.Frame(sidebar_canvas)
        sidebar_canvas.create_window((0, 0), window=self.sidebar_frame, anchor="nw")
        sidebar_canvas.configure(yscrollcommand=sidebar_scrollbar.set)
    
        # Configure canvas to resize with window
        parent.rowconfigure(0, weight=1)

        def configure_canvas(event):
            sidebar_canvas.configure(scrollregion=sidebar_canvas.bbox("all"))

        self.sidebar_frame.bind("<Configure>", configure_canvas)

        # Title - FIX: use self.sidebar_frame, not sidebar_frame
        title_label = ttk.Label(self.sidebar_frame, text="Gas Exchange Analyzer", 
                               font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 25), padx=10)

        # File selection
        file_frame = ttk.LabelFrame(self.sidebar_frame, text="File Selection", padding="15")
        file_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15), padx=10)

        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=42)
        file_entry.grid(row=0, column=0, padx=(0, 10), pady=5)

        ttk.Button(file_frame, text="Browse", command=self.browse_file).grid(
            row=0, column=1, pady=5)

        self.process_btn = ttk.Button(file_frame, text="Process File", command=self.process_file)
        self.process_btn.grid(row=1, column=0, columnspan=2, pady=(10, 5))

        # Processing controls frame (initially hidden)
        self.processing_frame = ttk.LabelFrame(self.sidebar_frame, text="Processing Controls", padding="15")
        # Don't grid it yet - will be shown after file processing

        # Channel selector
        ttk.Label(self.processing_frame, text="Select Channels:", 
                  font=("Arial", 10, "bold")).grid(
            row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))

        self.channel_select_all_btn = ttk.Button(self.processing_frame, text="Select All", 
                                                command=self.select_all_channels, width=12)
        self.channel_select_all_btn.grid(row=1, column=0, padx=(0, 8), pady=(0, 12))

        self.channel_deselect_all_btn = ttk.Button(self.processing_frame, text="Deselect All", 
                                                  command=self.deselect_all_channels, width=12)
        self.channel_deselect_all_btn.grid(row=1, column=1, padx=(0, 8), pady=(0, 12))

        self.channel_listbox_frame = ttk.Frame(self.processing_frame)
        self.channel_listbox_frame.grid(row=2, column=0, columnspan=3, 
                                       sticky=(tk.W, tk.E), pady=(0, 15))

        self.channel_listbox = tk.Listbox(self.channel_listbox_frame, selectmode=tk.MULTIPLE, 
                                         height=6, exportselection=False)
        channel_scrollbar = ttk.Scrollbar(self.channel_listbox_frame, orient="vertical", 
                                         command=self.channel_listbox.yview)
        self.channel_listbox.config(yscrollcommand=channel_scrollbar.set)

        self.channel_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        channel_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Bind selection change event
        self.channel_listbox.bind('<<ListboxSelect>>', self.on_channel_selection_change)

        # Time controls
        time_frame = ttk.Frame(self.processing_frame)
        time_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))

        # Remove Initial Points input
        ttk.Label(time_frame, text="Remove Initial Points:").grid(
            row=0, column=0, sticky=tk.W, pady=(0, 8))

        self.remove_start_var = tk.StringVar(value="0")
        self.remove_start_entry = ttk.Entry(time_frame, textvariable=self.remove_start_var, width=8)
        self.remove_start_entry.grid(row=0, column=1, padx=(5, 0), pady=(0, 8), sticky=tk.W)

        # Remove Final Points input
        ttk.Label(time_frame, text="Remove Final Points:").grid(
            row=1, column=0, sticky=tk.W, pady=(0, 8))

        self.remove_end_var = tk.StringVar(value="0")
        self.remove_end_entry = ttk.Entry(time_frame, textvariable=self.remove_end_var, width=8)
        self.remove_end_entry.grid(row=1, column=1, padx=(5, 0), pady=(0, 8), sticky=tk.W)

        # Shift Start Time input
        shift_time_frame = ttk.Frame(time_frame)
        shift_time_frame.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 8))

        ttk.Label(shift_time_frame, text="Shift Start Time:").grid(
            row=0, column=0, sticky=tk.W)

        self.time0_shift_var = tk.StringVar(value="0")
        self.time0_shift_entry = ttk.Entry(shift_time_frame, textvariable=self.time0_shift_var, width=8)
        self.time0_shift_entry.grid(row=0, column=1, padx=(5, 5), sticky=tk.W)

        self.shift_time_info_label = ttk.Label(shift_time_frame, text="", foreground="green", font=("Arial", 9))
        self.shift_time_info_label.grid(row=0, column=2, padx=(5, 0), sticky=tk.W)
        shift_time_frame.columnconfigure(2, weight=1)

        # Validation
        def validate_numeric_input(P):
            if P == "" or P == "-":
                return True
            try:
                if P.startswith('-'):
                    int(P[1:])
                else:
                    int(P)
                return True
            except ValueError:
                return False

        vcmd = (self.root.register(validate_numeric_input), '%P')
        self.remove_start_entry.config(validate="key", validatecommand=vcmd)
        self.remove_end_entry.config(validate="key", validatecommand=vcmd)
        self.time0_shift_entry.config(validate="key", validatecommand=vcmd)

        # Bind entry changes
        self.remove_start_var.trace('w', lambda *args: self.on_entry_change())
        self.remove_end_var.trace('w', lambda *args: self.on_entry_change())
        self.time0_shift_var.trace('w', lambda *args: self.on_entry_change_and_update_time_info())

        for entry_var in [self.remove_start_var, self.remove_end_var, self.time0_shift_var]:
            entry_var.trace('w', lambda *args: self.on_entry_change())

        self.slider_info_label = ttk.Label(time_frame, text="Enter values and click 'Apply Changes'", 
                                          foreground="blue", font=("Arial", 8))
        self.slider_info_label.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))

        self.apply_changes_btn = ttk.Button(self.processing_frame, text="Apply Changes", 
                                           command=self.apply_time_changes, state='disabled')
        self.apply_changes_btn.grid(row=4, column=0, columnspan=3, pady=(10, 15))

        self.initial_entry_values = {
            'remove_start': "0",
            'remove_end': "0",
            'time0_shift': "0"
        }

        # Time Range Information frame (initially hidden)
        self.time_range_frame = ttk.LabelFrame(self.sidebar_frame, text="Time Range Information", padding="15")
        # Don't grid it yet - will be shown after file processing

        self.time_range_text = tk.Text(self.time_range_frame, height=10, width=50, 
                                      font=("Courier New", 9), wrap=tk.WORD)
        self.time_range_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        time_range_scrollbar = ttk.Scrollbar(self.time_range_frame, command=self.time_range_text.yview)
        time_range_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.time_range_text.config(yscrollcommand=time_range_scrollbar.set)

        # Visualization Settings frame (initially hidden)
        self.color_frame = ttk.LabelFrame(self.sidebar_frame, text="Visualization Settings", padding="15")
        # Don't grid it yet - will be shown after file processing

        ttk.Label(self.color_frame, text="Color Palette:").grid(
            row=0, column=0, sticky=tk.W, pady=(0, 8))

        self.palette_var = tk.StringVar(value="tab10")
        palette_combo = ttk.Combobox(self.color_frame, textvariable=self.palette_var,
                                     values=["tab10", "viridis", "plasma", "magma", "cividis", 
                                             "inferno", "Set2", "Set3", "Dark2", "Paired",
                                             "colorblind1", "colorblind2", "colorblind3"],
                                     width=28)
        palette_combo.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        palette_combo.bind('<<ComboboxSelected>>', self.on_batch_palette_change)
        
        palette_combo.bind('<<ComboboxSelected>>', lambda e: self.update_all_plots())

        # Coefficient tables (initially hidden)
        self.setup_coefficient_tables(self.sidebar_frame, hidden=True)

        # Configure column weights
        self.sidebar_frame.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        # Mouse wheel scrolling
        def _on_mousewheel(event):
            sidebar_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        sidebar_canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    def finalize_slider_change(self):
        """Finalize slider change and update plots"""
        if self.original_calculated_data is None:
            return
        
        # Validate slider values
        df = self.original_calculated_data
        channels = df['Channel'].unique()
        
        # Find minimum points across channels for validation
        min_points_per_channel = {}
        for channel in channels:
            channel_data = df[df['Channel'] == channel]
            min_points_per_channel[str(channel)] = len(channel_data)
        
        if min_points_per_channel:
            min_points = min(min_points_per_channel.values())
            max_remove = max(0, min_points - 2)  # Leave at least 2 points
            
            # Validate remove_start
            remove_start = self.remove_start_var.get()
            if remove_start > max_remove:
                self.remove_start_var.set(max_remove)
                messagebox.showwarning("Warning", 
                                     f"Cannot remove more than {max_remove} initial points. Adjusted to {max_remove}.")
            
            # Validate remove_end
            remove_end = self.remove_end_var.get()
            if remove_end > max_remove:
                self.remove_end_var.set(max_remove)
                messagebox.showwarning("Warning", 
                                     f"Cannot remove more than {max_remove} final points. Adjusted to {max_remove}.")
            
            # Validate total removal
            if remove_start + remove_end > max_remove:
                # Distribute removal evenly
                half_max = max_remove // 2
                self.remove_start_var.set(min(remove_start, half_max))
                self.remove_end_var.set(min(remove_end, half_max))
                messagebox.showwarning("Warning", 
                                     f"Total removal limited to {max_remove}. Adjusted to {self.remove_start_var.get()}+{self.remove_end_var.get()}.")
        
        # Reset info label
        self.root.after(1000, lambda: self.slider_info_label.config(
            text="Drag sliders to adjust, release to update plots", 
            foreground="blue"
        ))
    
    def on_entry_change_and_update_time_info(self):
        """Handle entry changes and update time information display"""
        self.on_entry_change()  # Enable apply button
        self.update_shift_time_info()  # Update time information

    def batch_select_all_channels(self):
        """Select all channels in batch results"""
        for channel, var in self.batch_channel_vars.items():
            var.set(True)
    
        # Update selected channels list
        self.batch_selected_channels = list(self.batch_channel_vars.keys())
    
        # Refresh metric category plots
        if hasattr(self, 'batch_stats') and self.batch_stats:
            self.update_batch_metric_category_plots(self.batch_stats)
    
        # Refresh darkness plots
        if hasattr(self, 'batch_stats') and self.batch_stats:
            if hasattr(self, 'batch_combined_df'):
                self.update_batch_darkness_plots(self.batch_stats, self.batch_combined_df)
            else:
                self.update_batch_darkness_plots(self.batch_stats)
    
        # Refresh distribution plots
        self.update_batch_distribution_plots(self.batch_stats)
    
        # ============ UPDATE THE WUEi DISTRIBUTION TAB ============
        self.update_wuei_distribution_tab()

    def batch_deselect_all_channels(self):
        """Deselect all channels in batch results"""
        for channel, var in self.batch_channel_vars.items():
            var.set(False)
    
        # Update selected channels list
        self.batch_selected_channels = []
    
        # Refresh metric category plots
        if hasattr(self, 'batch_stats') and self.batch_stats:
            self.update_batch_metric_category_plots(self.batch_stats)
    
        # Refresh darkness plots
        if hasattr(self, 'batch_stats') and self.batch_stats:
            if hasattr(self, 'batch_combined_df'):
                self.update_batch_darkness_plots(self.batch_stats, self.batch_combined_df)
            else:
                self.update_batch_darkness_plots(self.batch_stats)
    
        # Refresh distribution plots
        self.update_batch_distribution_plots(self.batch_stats)
    
        # ============ UPDATE THE WUEi DISTRIBUTION TAB ============
        self.update_wuei_distribution_tab()
       
    def on_batch_palette_change(self, event=None):
        """Update batch plots when palette changes"""
        if hasattr(self, 'batch_stats') and self.batch_stats:
            # Clear the existing plots frame
            for widget in self.batch_plots_frame.winfo_children():
                widget.destroy()
        
            # Regenerate the batch plots with the new palette
            # First update the palette variable for batch plots
            palette = self.batch_palette_var.get()
            # Temporarily store the batch palette in a separate variable
            self.current_batch_palette = palette
            # Regenerate plots
            self.update_batch_plots_tab(self.batch_stats)
    
    def update_shift_time_info(self):
        """Update the time information display for Shift Start Time with integer values"""
        if not hasattr(self, 'shift_time_info_label'):
            return
        
        # Get the current time0_shift value
        try:
            time0_shift = int(self.time0_shift_var.get())
        except ValueError:
            time0_shift = 0
        
        # Get the time interval from data (integer)
        time_interval_minutes = self.get_time_interval_minutes()
        
        if time_interval_minutes is not None:
            # Calculate total shift in minutes (integer)
            total_shift_minutes = time0_shift * time_interval_minutes
            
            # Update the label with information
            if time0_shift == 0:
                self.shift_time_info_label.config(text="No time shift", foreground="green")
            elif time0_shift > 0:
                # Positive shift - time starts later (right shift)
                self.shift_time_info_label.config(
                    text=f"= +{total_shift_minutes} min (time starts later)", 
                    foreground="orange"
                )
            else:
                # Negative shift - time starts earlier (left shift)
                self.shift_time_info_label.config(
                    text=f"= {total_shift_minutes} min (time starts earlier)", 
                    foreground="blue"
                )
        else:
            self.shift_time_info_label.config(text="Load data to see time info", foreground="gray")
    
    def get_time_interval_minutes(self):
        """Get the time interval in minutes from the current data - returns integer"""
        if self.modified_calculated_data is not None and not self.modified_calculated_data.empty:
            df = self.modified_calculated_data
        elif self.original_calculated_data is not None and not self.original_calculated_data.empty:
            df = self.original_calculated_data
        else:
            return None
        
        if 'time_interval_minutes' in df.columns:
            # Get the first non-null time interval
            intervals = df['time_interval_minutes'].dropna().unique()
            if len(intervals) > 0:
                return int(intervals[0])  # Ensure integer
        
        # Fallback: calculate from time_repeated differences
        channels = df['Channel'].unique()
        for channel in channels:
            channel_data = df[df['Channel'] == channel]
            if len(channel_data) > 1:
                sorted_times = np.sort(channel_data['time_repeated'].unique())
                if len(sorted_times) > 1:
                    time_interval = sorted_times[1] - sorted_times[0]
                    return int(time_interval)  # Return as integer
        
        return None  # No data available
    
    def update_entry_ranges(self):
        """Update entry validation based on current data"""
        if self.original_calculated_data is None:
            return
        
        df = self.original_calculated_data
        if df is None or len(df) == 0:
            return
        
        # Find the channel with the minimum number of points
        channels = df['Channel'].unique()
        min_points = float('inf')
        
        for channel in channels:
            channel_data = df[df['Channel'] == channel]
            if len(channel_data) < min_points:
                min_points = len(channel_data)
        
        if min_points == float('inf') or min_points < 2:
            min_points = 10  # default fallback
        
        # Calculate safe maximum values (leave at least 2 points)
        max_remove = max(0, min_points - 2)
        
        # Update the info label with these limits
        self.update_entry_info_label(max_remove, min_points)
    
    def update_entry_info_label(self, max_remove=None, min_points=None):
        """Update the information label about entry limits"""
        if self.original_calculated_data is None:
            self.slider_info_label.config(text="Load data to see limits")
            return
        
        if max_remove is None or min_points is None:
            df = self.original_calculated_data
            channels = df['Channel'].unique()
            
            # Calculate min points across channels
            points_per_channel = {}
            for channel in channels:
                channel_data = df[df['Channel'] == channel]
                points_per_channel[str(channel)] = len(channel_data)
            
            if points_per_channel:
                min_points = min(points_per_channel.values())
                max_remove = max(0, min_points - 2)
            else:
                return
        
        info_text = f"Max removals per channel: {max_remove} (leaves 2+ points, min {min_points} points)"
        
        self.slider_info_label.config(text=info_text)
    
    def reset_entries(self):
        """Reset all entry boxes to default values"""
        self.remove_start_var.set("0")
        self.remove_end_var.set("0")
        self.time0_shift_var.set("0")
        
        # Update initial values
        self.initial_entry_values = {
            'remove_start': "0",
            'remove_end': "0", 
            'time0_shift': "0"
        }
        
        # Disable Apply Changes button
        if hasattr(self, 'apply_changes_btn'):
            self.apply_changes_btn.config(state='disabled')
        
        # Reset info label
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="Enter values, then click 'Apply Changes'", 
                foreground="blue"
            )
    def check_slider_changes(self):
        """Check if sliders have changed from their initial values"""
        if not hasattr(self, 'initial_slider_values'):
            return False
        
        current_values = {
            'remove_start': self.remove_start_var.get(),
            'remove_end': self.remove_end_var.get(),
            'time0_shift': self.time0_shift_var.get()
        }
        
        for key, current in current_values.items():
            if current != self.initial_slider_values[key]:
                return True
        return False
    
    def add_batch_processing_ui(self):
        """Add batch processing UI to sidebar"""
        # Make sure sidebar_frame exists
        if not hasattr(self, 'sidebar_frame'):
            print("Warning: sidebar_frame not yet initialized")
            return

        # Batch processing frame
        self.batch_frame = ttk.LabelFrame(self.sidebar_frame, text="Batch Processing", padding="15")
        self.batch_frame.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=(10, 0), padx=10)

        # File selection for batch processing
        ttk.Label(self.batch_frame, text="Select Multiple Files:", 
                  font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))

        self.batch_file_paths = []

        # File listbox with scrollbar
        file_frame = ttk.Frame(self.batch_frame)
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        self.batch_file_listbox = tk.Listbox(file_frame, height=5, selectmode=tk.EXTENDED, exportselection=False)
        batch_scrollbar = ttk.Scrollbar(file_frame, orient="vertical", command=self.batch_file_listbox.yview)
        self.batch_file_listbox.config(yscrollcommand=batch_scrollbar.set)

        self.batch_file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        batch_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Bind selection event to update file info display
        self.batch_file_listbox.bind('<<ListboxSelect>>', self.on_batch_file_select)

        # Buttons for file management
        btn_frame = ttk.Frame(self.batch_frame)
        btn_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Button(btn_frame, text="Add Files", command=self.add_batch_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Clear All", command=self.clear_batch_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Process Batch", command=self.process_batch).pack(side=tk.LEFT, padx=2)

        # ============ ADD TIME RANGE INFORMATION FRAME ============
        self.batch_time_range_frame = ttk.LabelFrame(self.batch_frame, text="File Information (Click to select file)", padding="10")
        self.batch_time_range_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 5))
    
        # Create text widget for file information
        self.batch_time_range_text = tk.Text(self.batch_time_range_frame, height=12, width=50, 
                                        font=("Courier New", 9), wrap=tk.WORD)
        self.batch_time_range_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
    
        # Add scrollbar
        batch_time_scrollbar = ttk.Scrollbar(self.batch_time_range_frame, command=self.batch_time_range_text.yview)
        batch_time_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.batch_time_range_text.config(yscrollcommand=batch_time_scrollbar.set)
    
        # Configure frame to expand
        self.batch_time_range_frame.columnconfigure(0, weight=1)
        self.batch_time_range_frame.rowconfigure(0, weight=1)

        # Batch palette selector
        batch_palette_frame = ttk.Frame(self.batch_frame)
        batch_palette_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 5))

        ttk.Label(batch_palette_frame, text="Batch Plots Palette:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))

        self.batch_palette_var = tk.StringVar(value="tab10")
        self.batch_palette_combo = ttk.Combobox(batch_palette_frame, textvariable=self.batch_palette_var,
                                                values=["tab10", "viridis", "plasma", "magma", "cividis", 
                                                        "inferno", "Set2", "Set3", "Dark2", "Paired",
                                                        "colorblind1", "colorblind2", "colorblind3"],
                                                width=25)
        self.batch_palette_combo.grid(row=0, column=1, sticky=tk.W)
        self.batch_palette_combo.bind('<<ComboboxSelected>>', self.on_batch_palette_change)

        # Load channel mapping file
        ttk.Label(self.batch_frame, text="Channel Mapping File (Excel):").grid(row=5, column=0, sticky=tk.W, pady=(5, 0))

        self.mapping_path_var = tk.StringVar()
        mapping_entry = ttk.Entry(self.batch_frame, textvariable=self.mapping_path_var, width=40)
        mapping_entry.grid(row=6, column=0, padx=(0, 5), pady=5, sticky=tk.W)

        ttk.Button(self.batch_frame, text="Browse Mapping", command=self.browse_mapping_file).grid(row=6, column=1, pady=5)

        # Bind mapping file change to update file info
        self.mapping_path_var.trace('w', lambda *args: self.update_batch_file_info())

        # Output options
        ttk.Label(self.batch_frame, text="Output Options:", font=("Arial", 10, "bold")).grid(row=7, column=0, sticky=tk.W, pady=(10, 5))

        self.output_dir_var = tk.StringVar()
        output_dir_entry = ttk.Entry(self.batch_frame, textvariable=self.output_dir_var, width=40)
        output_dir_entry.grid(row=8, column=0, padx=(0, 5), pady=5, sticky=tk.W)

        ttk.Button(self.batch_frame, text="Browse Output", command=self.browse_output_dir).grid(row=8, column=1, pady=5)

        # Statistics options
        self.show_stats_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.batch_frame, text="Generate statistics (mean ± SD)", 
            variable=self.show_stats_var).grid(row=9, column=0, columnspan=2, sticky=tk.W, pady=5)
    
        # Store batch file info for quick access
        self.batch_file_info_cache = {}

    def browse_mapping_file(self):
        """Browse for channel mapping Excel file"""
        filename = filedialog.askopenfilename(
            title="Select Channel Mapping Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.mapping_path_var.set(filename)

    def browse_output_dir(self):
        """Browse for output directory"""
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_dir_var.set(directory)

    def add_batch_files(self):
        """Add files to batch processing list"""
        filenames = filedialog.askopenfilenames(
            title="Select Gas Exchange Files",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        for filename in filenames:
            if filename not in self.batch_file_paths:
                self.batch_file_paths.append(filename)
                self.batch_file_listbox.insert(tk.END, os.path.basename(filename))
                # Clear cache for this file
                if filename in self.batch_file_info_cache:
                    del self.batch_file_info_cache[filename]
    
        # If this was the first file added, select it to show info
        if len(self.batch_file_paths) == 1:
            self.batch_file_listbox.selection_set(0)
            self.on_batch_file_select()

    def clear_batch_files(self):
        """Clear all files from batch list"""
        self.batch_file_paths = []
        self.batch_file_listbox.delete(0, tk.END)
        self.batch_file_info_cache = {}
        # Clear info display
        self.batch_time_range_text.delete(1.0, tk.END)
        self.batch_time_range_text.insert(1.0, "No files selected.\n\nAdd files using 'Add Files' button.")
    
    def on_batch_file_select(self, event=None):
        """Handle selection of a file in the batch list to show its information"""
        selection = self.batch_file_listbox.curselection()
        if not selection:
            return
        
        # Get the selected file path
        selected_index = selection[0]
        if selected_index < len(self.batch_file_paths):
            filepath = self.batch_file_paths[selected_index]
            self.update_batch_file_info_display(filepath)

    def update_batch_file_info(self):
        """Update all batch file information (called when mapping file changes)"""
        if not hasattr(self, 'batch_file_paths') or not self.batch_file_paths:
            return
        
        # Clear cache
        self.batch_file_info_cache = {}
        
        # Update info for currently selected file
        selection = self.batch_file_listbox.curselection()
        if selection:
            selected_index = selection[0]
            if selected_index < len(self.batch_file_paths):
                filepath = self.batch_file_paths[selected_index]
                self.update_batch_file_info_display(filepath)

    def update_batch_file_info_display(self, filepath):
        """Update the time range information display for a specific batch file"""
        filename = os.path.basename(filepath)
        
        # Check if we have cached info for this file
        if filepath in self.batch_file_info_cache:
            info_text = self.batch_file_info_cache[filepath]
            self.batch_time_range_text.delete(1.0, tk.END)
            self.batch_time_range_text.insert(1.0, info_text)
            return
        
        # Show loading message
        self.batch_time_range_text.delete(1.0, tk.END)
        self.batch_time_range_text.insert(1.0, f"Loading information for:\n{filename}\n\nPlease wait...")
        self.batch_time_range_text.update()
        
        try:
            # Load and process the file to get time range information
            info_text = self.get_batch_file_info(filepath, filename)
            
            # Cache the info
            self.batch_file_info_cache[filepath] = info_text
            
            # Update display
            self.batch_time_range_text.delete(1.0, tk.END)
            self.batch_time_range_text.insert(1.0, info_text)
            
        except Exception as e:
            error_text = f"Error loading file information:\n{str(e)}"
            self.batch_time_range_text.delete(1.0, tk.END)
            self.batch_time_range_text.insert(1.0, error_text)

    def get_batch_file_info(self, filepath, filename):
        """Extract time range and channel information from a batch file with clean formatting"""
        try:
            # Load raw data
            raw_data = pd.read_csv(filepath, delimiter='\t', skiprows=2, header=None, decimal=',')
        
            # Process headers
            headers = raw_data.iloc[0, 1:].tolist()
            data = raw_data.iloc[2:, 1:].reset_index(drop=True)
        
            column_names = [
                "Channel", "CO2_M", "CO2_R", "H2O_M", "H2O_R", "O3_M", "O3_R",
                "CO2_exch_rate", "H2O_evol_rate", "overall_O3_uptake_rate", "air_temp", 
                "transp_ind_leaf_temp_depr", "leaf_temp", "satur_hum_at_leaf_temp", 
                "leaf_air_hum_grad", "stomatal_cond", "stomatal_O3_uptake_rate", "O3_cumul_dose",
                "stom_change_rate", "normalised_stom_cond", "relstom_condchange_rate", "leaf_area", 
                "air_flow_rate", "absorbed_radiation", "boundary_layer_cond", "cutic_conductance", 
                "temp_calibr", "time"
            ]
        
            data.columns = column_names
        
            # Convert numeric columns
            for col in column_names[1:-1]:
                data[col] = pd.to_numeric(data[col].astype(str).str.replace(',', '.'), errors='coerce')
        
            # Normalize channel names
            data['Channel'] = data['Channel'].astype(str).str.strip()
            data['Channel'] = data['Channel'].str.split(',').str[0]
            data['Channel'] = data['Channel'].str.split('.').str[0]
        
            data['Channel'] = data['Channel'].apply(lambda x: re.sub(r'\D', '', str(x)))
        
            # Extract time information
            time_seconds = []
            for time_str in data['time']:
                parts = str(time_str).split(':')
                if len(parts) >= 3:
                    secs = float(parts[0]) * 3600 + float(parts[1]) * 60 + float(parts[2])
                else:
                    secs = 0.0
                time_seconds.append(secs)
        
            # Calculate continuous time
            day_increment = 0
            full_seconds = []
            prev_sec = time_seconds[0] if time_seconds else 0
        
            for sec in time_seconds:
                if sec < prev_sec:
                    day_increment += 1
                prev_sec = sec
                full_seconds.append(sec + day_increment * 24 * 3600)
        
            full_minutes = [sec / 60 for sec in full_seconds]
        
            # Get time interval
            time_interval = 1
            if len(full_minutes) > 1:
                intervals = np.diff(full_minutes)
                positive_intervals = intervals[intervals > 0]
                if len(positive_intervals) > 0:
                    time_interval = int(np.round(np.median(positive_intervals)))
        
            # Get unique channels
            channels = data['Channel'].unique()
            channels = sorted([ch for ch in channels if ch and ch != 'nan'], key=lambda x: int(x) if x.isdigit() else 0)
        
            # Load mapping if available
            channel_mapping = {}
            replicate_name = None
        
            if hasattr(self, 'mapping_path_var') and self.mapping_path_var.get():
                try:
                    mapping_df = pd.read_excel(self.mapping_path_var.get())
                    col_names = mapping_df.columns.tolist()
                
                    if len(col_names) >= 5:
                        # Identify columns by position (as per your structure)
                        replicate_col = col_names[0]  # Rep1, Rep2, etc.
                        channel_num_col = col_names[1]  # Channel 1, Channel 2, etc.
                        channel_name_col = col_names[2]  # drq1-4_nacl, etc.
                    
                        # Extract replicate name from filename
                        for rep in mapping_df[replicate_col].astype(str):
                            if rep.lower() in filename.lower():
                                replicate_name = rep
                                break
                    
                        if replicate_name:
                            rep_data = mapping_df[mapping_df[replicate_col].astype(str) == replicate_name]
                            for _, row in rep_data.iterrows():
                                channel_num = str(row[channel_num_col]).strip()
                                # Clean channel number
                                channel_num = re.sub(r'\D', '', channel_num)
                                new_name = str(row[channel_name_col]).strip()
                            
                                channel_mapping[channel_num] = {
                                    'new_name': new_name
                                }
                except Exception as e:
                    print(f"Error loading mapping for info display: {e}")
        
            # Build channel information (simplified - no leaf area or radiation)
            channel_info = []
            for channel in channels:
                channel_data = data[data['Channel'] == channel]
                if len(channel_data) > 0:
                    ch_points = len(channel_data)
                
                    # Get mapped name if available
                    if channel in channel_mapping:
                        display_name = channel_mapping[channel]['new_name']
                    else:
                        display_name = channel
                
                    channel_info.append({
                        'channel_num': int(channel) if channel.isdigit() else channel,
                        'original_channel': channel,
                        'display_name': display_name,
                        'points': ch_points
                    })
        
            # Sort by channel number
            channel_info.sort(key=lambda x: x['channel_num'] if isinstance(x['channel_num'], int) else 0)
        
            # Build info text with clean formatting (NO time columns, NO leaf area, NO radiation)
            info_text = f"FILE INFORMATION: {filename}\n"
            info_text += "=" * 50 + "\n\n"
        
            info_text += "Time Summary:\n"
            info_text += f"    Time interval: {time_interval:.1f} min\n"
            info_text += f"    Total points: {len(data)}\n"
            info_text += f"    Channels: {len(channels)}\n\n"
        
            info_text += "Channel Details:\n"
            info_text += "-" * 50 + "\n"
            info_text += f"{'Channel':<10} {'Name':<30} {'Points':<8}\n"
            info_text += "-" * 50 + "\n"
        
            for ch in channel_info:
                info_text += f"{ch['original_channel']:<10} {ch['display_name']:<30} {ch['points']:<8}\n"
        
            # Add mapping info
            if channel_mapping:
                info_text += f"\n✓ Mapping Applied: {os.path.basename(self.mapping_path_var.get())}\n"
                info_text += f"  Replicate: {replicate_name if replicate_name else 'Auto-detected'}\n"
            else:
                info_text += f"\n✗ No mapping file loaded\n"
        
            info_text += "\n" + "=" * 50 + "\n"
            info_text += "Click 'Process Batch' to analyze all files"
        
            return info_text
        
        except Exception as e:
            traceback.print_exc()
            return f"Error loading {filename}:\n{str(e)}"
    
    def process_batch(self):
        """Process multiple files in batch mode"""
        if not self.batch_file_paths:
            messagebox.showwarning("Warning", "No files selected for batch processing")
            return

        if not self.output_dir_var.get():
            messagebox.showwarning("Warning", "Please select an output directory")
            return

        # Show processing dialog
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Batch Processing")
        progress_window.transient(self.root)
    
        # ADD THIS SECTION - Center the window on screen
        window_width = 400
        window_height = 200
    
        # Get screen dimensions
        screen_width = progress_window.winfo_screenwidth()
        screen_height = progress_window.winfo_screenheight()
    
        # Calculate position to center the window
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
    
        # Set geometry with centered position
        progress_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
        # Set the icon for the progress window
        try:
            progress_window.iconbitmap("favicon32.ico")
        except:
            try:
                icon = tk.PhotoImage(file="favicon32.ico")
                progress_window.iconphoto(True, icon)
                progress_window.icon_image = icon  # Keep reference
            except:
                pass

        ttk.Label(progress_window, text="Processing files...", font=("Arial", 12, "bold")).pack(pady=10)
        progress_bar = ttk.Progressbar(progress_window, mode='determinate', maximum=len(self.batch_file_paths))
        progress_bar.pack(pady=10, padx=20, fill=tk.X)

        status_label = ttk.Label(progress_window, text="")
        status_label.pack(pady=5)

        # Load channel mapping if provided
        channel_mapping = {}
        if self.mapping_path_var.get():
            try:
                mapping_df = pd.read_excel(self.mapping_path_var.get())
            
                # Print column names for debugging
                print(f"Mapping file columns: {mapping_df.columns.tolist()}")
                print(f"Mapping file shape: {mapping_df.shape}")
            
                # Your Excel structure:
                # Column 0: Replicate identifier (Rep1, Rep2, etc.)
                # Column 1: Channel number (Channel 1, Channel 2, etc.)
                # Column 2: Channel name to rename to (drq1-4_nacl, etc.)
                # Column 3: Leaf area value
                # Column 4: Absorbed radiation value
            
                # Determine column indices based on actual column names
                col_names = mapping_df.columns.tolist()
            
                # Try to identify columns by name or position
                if len(col_names) >= 5:
                    # Use column names if they match expected patterns
                    replicate_col = None
                    channel_num_col = None
                    channel_name_col = None
                    leaf_area_col = None
                    radiation_col = None
                
                    for col in col_names:
                        col_lower = str(col).lower()
                        if 'rep' in col_lower or 'replicate' in col_lower:
                            replicate_col = col
                        elif 'channel' in col_lower and ('num' in col_lower or 'number' in col_lower):
                            channel_num_col = col
                        elif 'name' in col_lower or 'channel' in col_lower:
                            if channel_num_col is None:  # First channel-related column
                                channel_name_col = col
                        elif 'leaf' in col_lower or 'area' in col_lower:
                            leaf_area_col = col
                        elif 'radiation' in col_lower or 'absorbed' in col_lower:
                            radiation_col = col
                
                    # If we couldn't identify by name, use position
                    if replicate_col is None:
                        replicate_col = col_names[0]
                    if channel_num_col is None:
                        channel_num_col = col_names[1]
                    if channel_name_col is None:
                        channel_name_col = col_names[2]
                    if leaf_area_col is None:
                        leaf_area_col = col_names[3]
                    if radiation_col is None:
                        radiation_col = col_names[4]
                
                    print(f"Using columns: Replicate='{replicate_col}', ChannelNum='{channel_num_col}', "
                          f"ChannelName='{channel_name_col}', LeafArea='{leaf_area_col}', Radiation='{radiation_col}'")
                
                    # Process each row
                    for _, row in mapping_df.iterrows():
                        rep = str(row[replicate_col]).strip()
                        channel_num = str(row[channel_num_col]).strip()
                        new_channel_name = str(row[channel_name_col]).strip()
                    
                        # Extract base name without the bracketed number
                        # Example: "drq1-4_nacl (2)" -> "drq1-4_nacl"
                        if ' (' in new_channel_name:
                            base_name = new_channel_name.split(' (')[0]
                        else:
                            base_name = new_channel_name
                    
                        # Get leaf area and radiation values
                        leaf_area = None
                        radiation = None
                        try:
                            leaf_area = float(row[leaf_area_col])
                        except (ValueError, TypeError):
                            pass
                    
                        try:
                            radiation = float(row[radiation_col])
                        except (ValueError, TypeError):
                            pass
                    
                        # Store mapping for this replicate and channel
                        if rep not in channel_mapping:
                            channel_mapping[rep] = {}
                    
                        channel_mapping[rep][channel_num] = {
                            'original_channel_num': channel_num,
                            'new_name': base_name,
                            'leaf_area': leaf_area,
                            'radiation': radiation,
                            'full_name': new_channel_name
                        }
                
                    print(f"Loaded mapping for {len(channel_mapping)} replicates")
                    print(f"Channels per replicate: {[len(v) for v in channel_mapping.values()]}")
                
                else:
                    print(f"Expected at least 5 columns, found {len(col_names)}")
                
            except Exception as e:
                print(f"Error loading mapping: {e}")
                traceback.print_exc()

        # Process each file
        results_by_metric = {}
        all_results = []

        for i, filepath in enumerate(self.batch_file_paths):
            filename = os.path.basename(filepath)
            status_label.config(text=f"Processing: {filename}")
            progress_window.update()

            try:
                # Extract replicate identifier from filename
                # Assuming format like "Rep1_data.txt" or "data_Rep1.txt"
                replicate_name = None
                for rep in channel_mapping.keys():
                    if rep.lower() in filename.lower():
                        replicate_name = rep
                        break
            
                if replicate_name is None and channel_mapping:
                    # Try to find by matching first part
                    for rep in channel_mapping.keys():
                        if filename.startswith(rep):
                            replicate_name = rep
                            break
            
                print(f"Processing {filename} as replicate: {replicate_name}")
            
                # Process the file
                result = self.process_single_file(filepath, channel_mapping, replicate_name)
                if result is not None:
                    all_results.append(result)

                    # Collect results by metric for averaging
                    for metric in result['calculated_metrics'].columns:
                        if metric not in results_by_metric:
                            results_by_metric[metric] = []
                        results_by_metric[metric].append(result['calculated_metrics'][metric])

            except Exception as e:
                print(f"Error processing {filepath}: {e}")
                traceback.print_exc()
                status_label.config(text=f"Error: {os.path.basename(filepath)}")
                progress_window.update()

            progress_bar['value'] = i + 1
            progress_window.update()

        status_label.config(text="Generating statistics...")
        progress_window.update()

        # Generate statistics
        if all_results:
            self.generate_batch_statistics(all_results, results_by_metric)
            status_label.config(text="Batch processing complete!")
            messagebox.showinfo("Success", 
                f"Processed {len(all_results)} files successfully!\n\n"
                f"Results saved to: {self.output_dir_var.get()}")
        else:
            status_label.config(text="No files processed successfully")
            messagebox.showerror("Error", "No files were processed successfully")

        progress_window.destroy()

    def process_single_file(self, filepath, channel_mapping, replicate_name=None):
        """Process a single file for batch processing with channel mapping"""
        try:
            # Load raw data
            raw_data = pd.read_csv(filepath, delimiter='\t', skiprows=2, header=None, decimal=',')
        
            # Process raw data similarly to existing method
            headers = raw_data.iloc[0, 1:].tolist()
            data = raw_data.iloc[2:, 1:].reset_index(drop=True)

            column_names = [
                "Channel", "CO2_M", "CO2_R", "H2O_M", "H2O_R", "O3_M", "O3_R",
                "CO2_exch_rate", "H2O_evol_rate", "overall_O3_uptake_rate", "air_temp", 
                "transp_ind_leaf_temp_depr", "leaf_temp", "satur_hum_at_leaf_temp", 
                "leaf_air_hum_grad", "stomatal_cond", "stomatal_O3_uptake_rate", "O3_cumul_dose",
                "stom_change_rate", "normalised_stom_cond", "relstom_condchange_rate", "leaf_area", 
                "air_flow_rate", "absorbed_radiation", "boundary_layer_cond", "cutic_conductance", 
                "temp_calibr", "time"
            ]

            data.columns = column_names

            # Convert numeric columns
            for col in column_names[1:-1]:  # Skip Channel and time
                data[col] = pd.to_numeric(data[col].astype(str).str.replace(',', '.'), errors='coerce')
        
            # ============ CHANNEL NORMALIZATION ============
            # Convert to string and extract just the channel number
            data['Channel'] = data['Channel'].astype(str).str.strip()
        
            # Split on comma and take first part, then split on decimal
            data['Channel'] = data['Channel'].str.split(',').str[0]
            data['Channel'] = data['Channel'].str.split('.').str[0]
        
            # Extract digits only
            def extract_channel_number(x):
                digits = re.sub(r'\D', '', str(x))
                if digits:
                 def extract_channel_number(x):return str(int(digits))
                return x
        
            data['Channel'] = data['Channel'].apply(extract_channel_number)
        
            print(f"Normalized channel names: {data['Channel'].unique().tolist()}")
        
            # ============ APPLY CHANNEL MAPPING ============
            channel_rename_map = {}
            if replicate_name and replicate_name in channel_mapping:
                mapping = channel_mapping[replicate_name]
                print(f"Applying mapping for replicate {replicate_name}: {list(mapping.keys())}")
            
                leaf_area_map = {}
                radiation_map = {}
            
                for channel_num, info in mapping.items():
                    # Clean the channel number from mapping
                    channel_num_clean = channel_num.replace('Channel ', '').strip()
                    channel_num_clean = channel_num_clean.split(',')[0].split('.')[0]
                    channel_num_clean = re.sub(r'\D', '', channel_num_clean)
                
                    # Store the FULL name (with bracket) for this replicate
                    # This preserves the replicate-specific naming
                    full_name = info['full_name']  # This already has the bracketed number
                    print(f"  Mapping {channel_num_clean} -> {full_name}")
                
                    channel_rename_map[channel_num_clean] = full_name
                
                    if info['leaf_area'] is not None:
                        leaf_area_map[channel_num_clean] = info['leaf_area']
                    if info['radiation'] is not None:
                        radiation_map[channel_num_clean] = info['radiation']
            
                # Apply channel renaming
                for old_name, new_name in channel_rename_map.items():
                    mask = data['Channel'] == old_name
                    if not mask.any():
                        try:
                            old_int = int(old_name)
                            data_int = data['Channel'].astype(int)
                            mask = data_int == old_int
                        except:
                            pass
                
                    if mask.any():
                        data.loc[mask, 'Channel'] = new_name
                        print(f"  Renamed {mask.sum()} rows from '{old_name}' to '{new_name}'")
                    else:
                        print(f"  Warning: Could not find channel '{old_name}' in data")
                        print(f"  Available channels: {data['Channel'].unique().tolist()}")
            
                # Apply leaf area and radiation updates
                for channel, leaf_value in leaf_area_map.items():
                    if channel in channel_rename_map:
                        channel_key = channel_rename_map[channel]
                    else:
                        channel_key = channel
                
                    mask = data['Channel'].str.strip() == channel_key
                    if mask.any():
                        data.loc[mask, 'leaf_area'] = leaf_value
                        print(f"  Set leaf_area for '{channel_key}' to {leaf_value}")
            
                for channel, rad_value in radiation_map.items():
                    if channel in channel_rename_map:
                        channel_key = channel_rename_map[channel]
                    else:
                        channel_key = channel
                
                    mask = data['Channel'].str.strip() == channel_key
                    if mask.any():
                        data.loc[mask, 'absorbed_radiation'] = rad_value
                        print(f"  Set absorbed_radiation for '{channel_key}' to {rad_value}")

            # Verify channel names after renaming
            print(f"Final channel names: {data['Channel'].unique().tolist()}")

            # Calculate metrics
            df = calculate_common_metrics(data)
            df = calculate_conductance_metrics(df)
            df = calculate_stomatal_metrics(df)
            df = calculate_ozone_metrics(df)
            df = calculate_co2_metrics(df)
            df = calculate_additional_metrics(df)

            # Ensure Channel column remains string
            df['Channel'] = df['Channel'].astype(str)
        
            # Process time columns
            df = process_time_columns(df, 0, 0, 0)

            # Add replicate identifier
            if replicate_name is None:
                replicate_name = os.path.basename(filepath).split('_')[0]
            df['replicate'] = replicate_name

            return {
                'replicate': replicate_name,
                'filepath': filepath,
                'calculated_metrics': df,
                'raw_data': data,
                'channel_mapping': channel_rename_map,
                'renamed_channels': list(channel_rename_map.values()) if channel_rename_map else df['Channel'].unique().tolist()
            }

        except Exception as e:
            print(f"Error in process_single_file: {e}")
            traceback.print_exc()
            return None


    def generate_wuei_distribution_plot(self, all_processed_data, output_dir):
        """Generate violin + boxplot for WUEi distribution by channel (all time points combined)"""
    
        # Combine all processed data into a single dataframe
        combined_df = pd.concat(all_processed_data, ignore_index=True)
    
        # Apply base name extraction to Channel column
        combined_df['Channel_original'] = combined_df['Channel']
        combined_df['Channel'] = combined_df['Channel'].astype(str).str.replace(r'\s*\(\d+\)$', '', regex=True)
    
        # Filter to only include WUEi
        if 'WUEi' not in combined_df.columns:
            print("WUEi column not found, skipping distribution plot")
            return
    
        # Close any existing figures to prevent memory issues
        plt.close('all')
    
        # Create the figure
        fig, ax = plt.subplots(figsize=(12, 8))
    
        # Prepare data for each base channel (ALL time points combined)
        channels = combined_df['Channel'].unique()
        wuei_data = []
        channel_labels = []
    
        for channel in sorted(channels):
            channel_data = combined_df[combined_df['Channel'] == channel]
            wuei_values = channel_data['WUEi'].dropna().values
            if len(wuei_values) > 0:
                wuei_data.append(wuei_values)
                channel_labels.append(channel)
    
        if not wuei_data:
            ax.text(0.5, 0.5, 'No WUEi data available', ha='center', va='center', fontsize=12)
            ax.set_axis_off()
            plt.savefig(os.path.join(output_dir, "WUEi_Distribution.png"), dpi=300, bbox_inches='tight')
            plt.close(fig)
            return
    
        # Get colors for channels
        colors = plt.cm.tab10(np.linspace(0, 1, len(wuei_data)))
        positions = range(1, len(wuei_data) + 1)
    
        # --- LAYER 1: VIOLIN PLOT (background) ---
        violin_parts = ax.violinplot(wuei_data, positions=positions, 
                                     showmeans=False, showmedians=False, showextrema=False)
    
        for i, body in enumerate(violin_parts['bodies']):
            body.set_facecolor(colors[i])
            body.set_alpha(0.3)
            body.set_edgecolor(colors[i])
            body.set_linewidth(1)
            body.set_zorder(1)
    
        # --- LAYER 2: BOXPLOT (foreground) ---
        box_plot = ax.boxplot(wuei_data, positions=positions, patch_artist=True, tick_labels=channel_labels)
    
        for i, (box, color) in enumerate(zip(box_plot['boxes'], colors)):
            box.set_facecolor(color)
            box.set_alpha(0.7)
            box.set_edgecolor('black')
            box.set_linewidth(1.5)
            box.set_zorder(3)
    
        for median in box_plot['medians']:
            median.set_color('red')
            median.set_linewidth(2.5)
            median.set_zorder(4)
    
        for whisker in box_plot['whiskers']:
            whisker.set(color='black', linewidth=1.5, zorder=2)
        for cap in box_plot['caps']:
            cap.set(color='black', linewidth=1.5, zorder=2)
    
        # --- LAYER 3: DATA POINTS (top layer) ---
        for i, data in enumerate(wuei_data):
            x_jitter = np.random.normal(positions[i], 0.08, size=len(data))
            ax.scatter(x_jitter, data, alpha=0.5, color=colors[i], s=20, 
                      edgecolor='black', linewidth=0.5, zorder=5)
    
        # --- LAYER 4: MEAN MARKERS ---
        for i, data in enumerate(wuei_data):
            mean_val = np.mean(data)
            ax.scatter(positions[i], mean_val, color='yellow', s=100, 
                      edgecolor='black', linewidth=1.5, zorder=6, marker='D')
    
        # Configure plot
        ax.set_title('WUEi Distribution by Channel (All Time Points Combined)', 
                    fontsize=14, fontweight='bold')
        ax.set_xlabel('Channel', fontsize=12)
        ax.set_ylabel('WUEi (μmol CO₂ / mol H₂O)', fontsize=12)
        ax.set_xticks(positions)
        ax.set_xticklabels(channel_labels, rotation=45, ha='right', fontsize=10)
        ax.grid(True, alpha=0.2, linestyle='--', axis='y', zorder=0)
    
        # Add statistics annotations
        for i, data in enumerate(wuei_data):
            median = np.median(data)
            q1 = np.percentile(data, 25)
            q3 = np.percentile(data, 75)
        
            ax.text(positions[i], median, f'{median:.1f}', 
                    ha='center', va='center', fontsize=9, fontweight='bold',
                    bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.8, edgecolor='none'),
                    zorder=7)
        
            ax.text(positions[i], q1, f'Q1', ha='center', va='top', fontsize=8, alpha=0.7)
            ax.text(positions[i], q3, f'Q3', ha='center', va='bottom', fontsize=8, alpha=0.7)
    
        # Legend
        legend_elements = [
            Patch(facecolor='gray', alpha=0.3, edgecolor='gray', label='Violin (distribution shape)'),
            Rectangle((0, 0), 1, 1, facecolor='gray', alpha=0.7, edgecolor='black', label='Box (quartiles)'),
            Line2D([0], [0], color='red', linewidth=2.5, label='Median'),
            Line2D([0], [0], marker='o', color='w', markerfacecolor='gray', markersize=8, label='Data points'),
            Line2D([0], [0], marker='D', color='w', markerfacecolor='yellow', markersize=8, label='Mean'),
        ]
    
        ax.legend(handles=legend_elements, loc='upper left', 
                  bbox_to_anchor=(1.02, 1), fontsize=9, framealpha=0.9)
    
        plt.tight_layout()
        plt.subplots_adjust(right=0.85)
    
        # Save the plot
        output_path = os.path.join(output_dir, "WUEi_Distribution.png")
        plt.savefig(output_path, dpi=300, bbox_inches='tight')
        plt.close(fig)
    
        print(f"WUEi distribution plot saved to {output_path}")
    
    
    def save_statistics_to_excel(self, stats_by_channel, output_dir):
        """Save statistics to Excel file with proper channel names"""
        filename = os.path.join(output_dir, "batch_statistics.xlsx")
    
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            for channel, stats in stats_by_channel.items():
                # Prepare data for DataFrame
                data = []
                time_points = sorted(stats.keys())
            
                for time_point in time_points:
                    row = {'time_repeated': time_point}
                    for metric, values in stats[time_point].items():
                        row[f'{metric}_mean'] = values['mean']
                        row[f'{metric}_std'] = values['std']
                        row[f'{metric}_n'] = values['n']
                        if 'sem' in values:
                            row[f'{metric}_sem'] = values['sem']
                    data.append(row)
            
                if data:
                    df = pd.DataFrame(data)
                    # Clean sheet name (Excel has 31 char limit, remove invalid chars)
                    sheet_name = str(channel).replace('/', '_').replace('\\', '_').replace(':', '_')
                    sheet_name = sheet_name[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
    
        print(f"Statistics saved to {filename}")

    def generate_averaged_plots(self, stats_by_channel, output_dir):
        """Generate averaged plots with error bars for all metrics using renamed channels"""
    
        # Define metrics to plot (same as main plots)
        plot_categories = {
            '01 Environment': [
                ("air_temp", "Air Temperature", "Temperature (ºC)"),
                ("saturating_air_humidity", "Saturation Humidity at Air Temperature", "Saturation Humidity (mmol/mol)"),
                ("absorbed_radiation", "Absorbed Radiation", "Radiation (cal/cm²s)"),
                ("relative_air_humidity", "Relative Air Humidity", "Humidity (%)"),
                ("VPD", "Vapor Pressure Deficit", "VPD (kPa)"),
                ("rad_ind_leaf_temp_increase", "Radiation-induced Leaf Temperature Increase", "Temperature Increase (ºC)")
            ],
            '02 Base_Fluxes': [
                ("CO2_exchange_rate", "CO₂ Exchange Rate", "CO₂ Exchange (mlmol/m²s)"),
                ("Transpiration_H2O_evol_rate", "Transpiration Rate", "Transpiration (mmol/m²s)"),
                ("transp_ind_leaf_temp_depr", "Transpiration-induced Leaf Temperature Depression", "Temperature Depression (ºC)"),
                ("leaf_temp_c", "Leaf Temperature", "Leaf Temperature (ºC)")
            ],
            '03 Humidity_Vapor': [
                ("leaf_air_hum_grad", "Leaf Air Humidity Gradient", "Gradient (mmol/mol)"),
                ("satur_hum_at_leaf_temp_c", "Saturation Humidity at Leaf Temperature", "Saturation Humidity (mmol/mol)"),
                ("mean_vapor_pressure", "Mean Vapor Pressure", "Pressure (mbar)")
            ],
            '04 Conductances': [
                ("overall_conductance", "Overall Conductance", "Conductance (mmol/m²s)"),
                ("corr_d_overall_conductance", "Corrected Overall Conductance", "Conductance (cm/s)"),
                ("massflow_correction_for_overall_resistance", "Massflow Correction for Overall Resistance", "Correction Factor"),
                ("b_layer_conductance", "Boundary Layer Conductance", "Conductance (cm/s)"),
                ("corr_leaf_conductance", "Corrected Leaf Conductance", "Conductance (cm/s)"),
                ("corr_stomatal_conductance", "Corrected Stomatal Conductance", "Conductance (cm/s)"),
                ("Stomatal_conductance_corrected", "Stomatal Conductance Corrected", "Conductance (mmol/m²s)"),
                ("normalised_corr_stomatal_conductance", "Normalised Stomatal Conductance", "Normalised Conductance"),
                ("corr_uncorr_stom_cond", "Corrected/Uncorrected Stomatal Conductance Ratio", "Ratio"),
                ("corr_stomatal_change_rate", "Corrected Stomatal Change Rate", "Change Rate (mmol/m²s)"),
                ("relat_stom_cond_change_rate", "Relative Stomatal Conductance Change Rate", "Change Rate")
            ],
            '05 Ozone': [
                ("overall_O3_uptake_rate_c", "Overall O₃ Uptake Rate", "Uptake Rate (nmol/m²s)"),
                ("corr_conductance_for_O3_c", "Corrected Conductance for O₃", "Conductance (mmol/m²s)"),
                ("corr_stomatal_O3_uptake_rate", "Stomatal O₃ Uptake", "O₃ Uptake (nmol O₃/m²s)"),
                ("corr_cumul_O3_dose", "Cumulative O₃ Dose", "O₃ Dose (μmol/m²)")
            ],
            '06 CO2': [
                ("massflow_corr_d_ca", "Massflow Corrected Ambient CO₂", "CO₂ (μmol CO₂/mol air)"),
                ("massflow_corr_for_CO2", "Massflow Correction for CO₂", "Correction Factor"),
                ("mflowcorr_Ca_gradient", "Massflow Correction for CO₂ Gradient", "Correction Factor"),
                ("corr_uncorr_ca", "Corrected/Uncorrected Ambient CO₂ Ratio", "Ratio"),
                ("stom_resist_to_CO2", "Stomatal Resistance to CO₂", "Resistance (m²s/mmol)"),
                ("b_layer_resist_to_CO2", "Boundary Layer Resistance to CO₂", "Resistance (m²s/mmol)"),
                ("corr_ted_CO2_grad", "Corrected CO₂ Gradient", "CO₂ Gradient (μmol CO₂/mol air)"),
                ("intercell_CO2_conc_in_gas_phase", "Intercellular CO₂", "CO₂ (μmol CO₂/mol air)"),
                ("uncorr_ci", "Uncorrected Intercellular CO₂", "CO₂ (μmol CO₂/mol air)"),
                ("corr_uncorr_ci", "Corrected/Uncorrected Intercellular CO₂ Ratio", "Ratio")
            ],
            '07 Mesophyll': [
                ("mesophyll_conductance_for_CO2", "Mesophyll Conductance for CO₂", "Conductance (mmol/m²s)"),
                ("uncorr_gm_prima", "Uncorrected Mesophyll Conductance", "Conductance (mmol/m²s)"),
                ("corr_uncorr_gm_prima", "Corrected/Uncorrected Mesophyll Conductance Ratio", "Ratio"),
                ("CO2_comp_point", "CO₂ Compensation Point", "CO₂ (μmol CO₂/mol air)"),
                ("uncorr_gamma", "Uncorrected CO₂ Compensation Point", "CO₂ (μmol CO₂/mol air)"),
                ("corr_uncorr_gamma", "Corrected/Uncorrected Compensation Point Ratio", "Ratio")
            ],
            '08 Efficiency': [
                ("WUEi", "Intrinsic Water Use Efficiency", "WUEi (μmol CO₂ / mol H₂O)"),
                ("WUE", "Water Use Efficiency", "WUE (μmol CO₂ / mmol H₂O)")
            ]
        }
    
        # Create subdirectory for averaged plots
        plots_dir = os.path.join(output_dir, "averaged_plots")
        os.makedirs(plots_dir, exist_ok=True)
    
        # Get palette for colors
        palette = self.palette_var.get() if hasattr(self, 'palette_var') else 'tab10'
    
        # Generate plots for each category
        for category, plots in plot_categories.items():
            n_plots = len(plots)
            fig, axes = plt.subplots(n_plots, 1, figsize=(12, 4 * n_plots))
            if n_plots == 1:
                axes = [axes]
        
            for idx, (metric, title, ylabel) in enumerate(plots):
                ax = axes[idx]
            
                # Check if metric exists in any channel
                metric_exists = False
                for channel, stats in stats_by_channel.items():
                    for time_point in stats.values():
                        if metric in time_point:
                            metric_exists = True
                            break
                    if metric_exists:
                        break
            
                if not metric_exists:
                    ax.text(0.5, 0.5, f'No data available for {metric}', 
                           ha='center', va='center', fontsize=12, transform=ax.transAxes)
                    ax.set_axis_off()
                    continue
            
                # Sort channels alphabetically for consistent legend
                sorted_channels = sorted(stats_by_channel.keys())
            
                # Generate colors for channels
                colors = plt.cm.tab10(np.linspace(0, 1, len(sorted_channels)))
            
                # Plot data for each channel with error bars
                for i, channel in enumerate(sorted_channels):
                    stats = stats_by_channel[channel]
                    # Sort time points numerically
                    time_points = sorted(stats.keys())
                    means = []
                    errors = []
                    valid_times = []
                
                    for t in time_points:
                        if metric in stats[t]:
                            means.append(stats[t][metric]['mean'])
                            errors.append(stats[t][metric]['sem'])
                            valid_times.append(t)
                
                    if means:
                        # Print debug info for CO2 exchange rate
                        if metric == 'CO2_exchange_rate':
                            print(f"\nChannel {channel} - CO2_exchange_rate:")
                            print(f"  Time points: {valid_times}")
                            print(f"  Means: {means}")
                            print(f"  Errors: {errors}")
                            print(f"  Number of points: {len(valid_times)}")
                    
                        # Plot with error bars - one point per time point
                        ax.errorbar(valid_times, means, yerr=errors, 
                                   label=str(channel),
                                   marker='o', capsize=3, linewidth=1.5, markersize=4,
                                   color=colors[i], alpha=0.8)
            
                ax.set_title(title, fontsize=12, fontweight='bold')
                ax.set_xlabel('Time (minutes)', fontsize=10)
                ax.set_ylabel(ylabel, fontsize=10)
            
                # Position legend on the right side (outside plot)
                ax.legend(bbox_to_anchor=(1.02, 1), loc='upper left', 
                         fontsize=8, framealpha=0.9, borderaxespad=0)
            
                ax.grid(True, alpha=0.3)
        
            plt.tight_layout()
            # Adjust layout to make room for legend on right
            plt.subplots_adjust(right=0.85)
        
            # Save plot
            output_path = os.path.join(plots_dir, f"{category}.png")
            plt.savefig(output_path, dpi=300, bbox_inches='tight')
            plt.close(fig)
        
            print(f"Saved {category}.png")
    
        print(f"Averaged plots saved to {plots_dir}")

    
    def create_combined_efficiency_plot(self, stats_by_channel, output_dir):
        """Create a combined plot with both WUEi and WUE"""
    
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
    
        # Sort channels alphabetically
        sorted_channels = sorted(stats_by_channel.keys())
        colors = plt.cm.tab10(np.linspace(0, 1, len(sorted_channels)))
    
        # Plot WUEi (left subplot)
        for i, channel in enumerate(sorted_channels):
            stats = stats_by_channel[channel]
            time_points = sorted(stats.keys())
            means = []
            errors = []
            valid_times = []
        
            for t in time_points:
                if 'WUEi' in stats[t]:
                    means.append(stats[t]['WUEi']['mean'])
                    errors.append(stats[t]['WUEi']['sem'])
                    valid_times.append(t)
        
            if means:
                ax1.errorbar(valid_times, means, yerr=errors, 
                            label=str(channel), marker='o', capsize=3,
                            linewidth=2, markersize=5, color=colors[i], alpha=0.7)
    
        ax1.set_title('Intrinsic Water Use Efficiency (WUEi)', fontsize=12, fontweight='bold')
        ax1.set_xlabel('Time (minutes)', fontsize=10)
        ax1.set_ylabel('WUEi (μmol CO₂ / mol H₂O)', fontsize=10)
        ax1.legend(bbox_to_anchor=(1.02, 1), loc='upper left', fontsize=8)
        ax1.grid(True, alpha=0.3)
    
        # Plot WUE (right subplot)
        for i, channel in enumerate(sorted_channels):
            stats = stats_by_channel[channel]
            time_points = sorted(stats.keys())
            means = []
            errors = []
            valid_times = []
        
            for t in time_points:
                if 'WUE' in stats[t]:
                    means.append(stats[t]['WUE']['mean'])
                    errors.append(stats[t]['WUE']['sem'])
                    valid_times.append(t)
        
            if means:
                ax2.errorbar(valid_times, means, yerr=errors, 
                            label=str(channel), marker='s', capsize=3,
                            linewidth=2, markersize=5, color=colors[i], alpha=0.7)
    
        ax2.set_title('Water Use Efficiency (WUE)', fontsize=12, fontweight='bold')
        ax2.set_xlabel('Time (minutes)', fontsize=10)
        ax2.set_ylabel('WUE (μmol CO₂ / mmol H₂O)', fontsize=10)
        ax2.legend(bbox_to_anchor=(1.02, 1), loc='upper left', fontsize=8)
        ax2.grid(True, alpha=0.3)
    
        plt.tight_layout()
        plt.subplots_adjust(right=0.85)
    
        # Save combined plot
        output_path = os.path.join(output_dir, "Efficiency_Combined.png")
        plt.savefig(output_path, dpi=300, bbox_inches='tight')
        plt.close()
    
        print(f"Saved Efficiency_Combined.png")
    
    def generate_batch_report(self, all_results, stats_by_channel, output_dir):
        """Generate a comprehensive batch processing report"""
        report_path = os.path.join(output_dir, "batch_report.txt")
    
        with open(report_path, 'w') as f:
            f.write("=" * 80 + "\n")
            f.write("GAS EXCHANGE DATA ANALYZER - BATCH PROCESSING REPORT\n")
            f.write("=" * 80 + "\n\n")
        
            f.write(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Output Directory: {output_dir}\n\n")
        
            f.write("PROCESSED FILES:\n")
            f.write("-" * 40 + "\n")
            for result in all_results:
                f.write(f"  - {os.path.basename(result['filepath'])} (Replicate: {result['replicate']})\n")
            f.write(f"\nTotal: {len(all_results)} files\n\n")
        
            f.write("STATISTICS SUMMARY:\n")
            f.write("-" * 40 + "\n")
            for channel, stats in stats_by_channel.items():
                f.write(f"\nChannel: {channel}\n")
                if stats:
                    first_time = min(stats.keys())
                    f.write(f"  Time points: {len(stats)} (from {first_time:.1f} min)\n")
                    f.write(f"  Metrics: {list(stats[first_time].keys())}\n")
        
            f.write("\n" + "=" * 80 + "\n")
            f.write("End of Report\n")
    
        print(f"Batch report saved to {report_path}")
    
    def setup_batch_results_tabs(self):
        """Setup batch results tabs - WITHOUT Distribution Plots tab"""
        # Create a notebook for batch results (will be added to main notebook)
        self.batch_results_notebook = ttk.Notebook(self.main_notebook)
        self.main_notebook.add(self.batch_results_notebook, text="Batch Results")

        # Configure the main notebook to expand
        self.main_notebook.columnconfigure(0, weight=1)
        self.main_notebook.rowconfigure(0, weight=1)

        # ============ TAB 1: Statistics Table ============
        stats_tab = ttk.Frame(self.batch_results_notebook)
        self.batch_results_notebook.add(stats_tab, text="Statistics Table")

        stats_tab.columnconfigure(0, weight=1)
        stats_tab.rowconfigure(0, weight=1)

        # Treeview for statistics table
        self.stats_tree = ttk.Treeview(stats_tab)
        vsb = ttk.Scrollbar(stats_tab, orient="vertical", command=self.stats_tree.yview)
        hsb = ttk.Scrollbar(stats_tab, orient="horizontal", command=self.stats_tree.xview)
        self.stats_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.stats_tree.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.E, tk.W))

        # ============ TAB 2: Metric Plots ============
        self.metric_plots_tab = ttk.Frame(self.batch_results_notebook)
        self.batch_results_notebook.add(self.metric_plots_tab, text="Metric Plots")

        # CRITICAL: Configure metric_plots_tab to expand
        self.metric_plots_tab.columnconfigure(0, weight=1)
        self.metric_plots_tab.rowconfigure(0, weight=0)  # Row 0 is channel selector (fixed height)
        self.metric_plots_tab.rowconfigure(1, weight=1)  # Row 1 is the metric category notebook (expands)
        self.metric_plots_tab.rowconfigure(2, weight=0)  # Row 2 would be for batch_plots_frame if used

        # Create channel selector frame (top)
        self.setup_batch_channel_selector()

        # Create notebook for metric categories (Environment, Base Fluxes, etc.)
        self.metric_category_notebook = ttk.Notebook(self.metric_plots_tab)
        self.metric_category_notebook.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), pady=(10, 0))

        # CRITICAL: Configure the notebook to expand
        self.metric_category_notebook.columnconfigure(0, weight=1)
        self.metric_category_notebook.rowconfigure(0, weight=1)

        # Create scrollable frames for each category
        self.metric_category_frames = {}
        self.metric_category_configs = {}

        # Define plot categories (same as before)
        plot_categories = {
           'Environment': [
               ("air_temp", "Air Temperature (Mean ± SEM)", "Temperature (ºC)"),
               ("saturating_air_humidity", "Saturation Humidity (Mean ± SEM)", "Saturation Humidity (mmol/mol)"),
               ("absorbed_radiation", "Absorbed Radiation (Mean ± SEM)", "Radiation (cal/cm²s)"),
               ("relative_air_humidity", "Relative Air Humidity (Mean ± SEM)", "Humidity (%)"),
               ("VPD", "Vapor Pressure Deficit (Mean ± SEM)", "VPD (kPa)"),
               ("rad_ind_leaf_temp_increase", "Radiation-induced Leaf Temperature Increase (Mean ± SEM)", "Temperature Increase (ºC)")
           ],
           'Base Fluxes': [
               ("CO2_exchange_rate", "CO₂ Exchange Rate (Mean ± SEM)", "CO₂ Exchange (μmol/m²s)"),
               ("Transpiration_H2O_evol_rate", "Transpiration Rate (Mean ± SEM)", "Transpiration (mmol/m²s)"),
               ("transp_ind_leaf_temp_depr", "Transpiration-induced Leaf Temperature Depression (Mean ± SEM)", "Temperature Depression (ºC)"),
               ("leaf_temp_c", "Leaf Temperature (Mean ± SEM)", "Leaf Temperature (ºC)")
           ],
           'Humidity Vapor': [
               ("leaf_air_hum_grad", "Leaf Air Humidity Gradient (Mean ± SEM)", "Gradient (mmol/mol)"),
               ("satur_hum_at_leaf_temp_c", "Saturation Humidity at Leaf Temperature (Mean ± SEM)", "Saturation Humidity (mmol/mol)"),
               ("mean_vapor_pressure", "Mean Vapor Pressure (Mean ± SEM)", "Pressure (mbar)")
           ],
           'Conductances': [
               ("overall_conductance", "Overall Conductance (Mean ± SEM)", "Conductance (mmol/m²s)"),
               ("corr_d_overall_conductance", "Corrected Overall Conductance (Mean ± SEM)", "Conductance (cm/s)"),
               ("massflow_correction_for_overall_resistance", "Massflow Correction (Mean ± SEM)", "Correction Factor"),
               ("b_layer_conductance", "Boundary Layer Conductance (Mean ± SEM)", "Conductance (cm/s)"),
               ("corr_leaf_conductance", "Corrected Leaf Conductance (Mean ± SEM)", "Conductance (cm/s)"),
               ("corr_stomatal_conductance", "Corrected Stomatal Conductance (Mean ± SEM)", "Conductance (cm/s)"),
               ("Stomatal_conductance_corrected", "Stomatal Conductance Corrected (Mean ± SEM)", "Conductance (mmol/m²s)"),
               ("normalised_corr_stomatal_conductance", "Normalised Stomatal Conductance (Mean ± SEM)", "Normalised Conductance"),
               ("corr_uncorr_stom_cond", "Corr/Uncorr Stomatal Conductance Ratio (Mean ± SEM)", "Ratio"),
               ("corr_stomatal_change_rate", "Corrected Stomatal Change Rate (Mean ± SEM)", "Change Rate (mmol/m²s)"),
               ("relat_stom_cond_change_rate", "Relative Stomatal Conductance Change Rate (Mean ± SEM)", "Change Rate")
           ],
           'Ozone': [
               ("overall_O3_uptake_rate_c", "Overall O₃ Uptake Rate (Mean ± SEM)", "Uptake Rate (nmol/m²s)"),
               ("corr_conductance_for_O3_c", "Corrected Conductance for O₃ (Mean ± SEM)", "Conductance (mmol/m²s)"),
               ("corr_stomatal_O3_uptake_rate", "Stomatal O₃ Uptake (Mean ± SEM)", "O₃ Uptake (nmol O₃/m²s)"),
               ("corr_cumul_O3_dose", "Cumulative O₃ Dose (Mean ± SEM)", "O₃ Dose (μmol/m²)")
           ],
           'CO2': [
               ("massflow_corr_d_ca", "Massflow Corrected Ambient CO₂ (Mean ± SEM)", "CO₂ (μmol CO₂/mol air)"),
               ("massflow_corr_for_CO2", "Massflow Correction for CO₂ (Mean ± SEM)", "Correction Factor"),
               ("mflowcorr_Ca_gradient", "Massflow Correction for CO₂ Gradient (Mean ± SEM)", "Correction Factor"),
               ("corr_uncorr_ca", "Corr/Uncorr Ambient CO₂ Ratio (Mean ± SEM)", "Ratio"),
               ("stom_resist_to_CO2", "Stomatal Resistance to CO₂ (Mean ± SEM)", "Resistance (m²s/mmol)"),
               ("b_layer_resist_to_CO2", "Boundary Layer Resistance to CO₂ (Mean ± SEM)", "Resistance (m²s/mmol)"),
               ("corr_ted_CO2_grad", "Corrected CO₂ Gradient (Mean ± SEM)", "CO₂ Gradient (μmol CO₂/mol air)"),
               ("intercell_CO2_conc_in_gas_phase", "Intercellular CO₂ (Mean ± SEM)", "CO₂ (μmol CO₂/mol air)"),
               ("uncorr_ci", "Uncorrected Intercellular CO₂ (Mean ± SEM)", "CO₂ (μmol CO₂/mol air)"),
               ("corr_uncorr_ci", "Corr/Uncorr Intercellular CO₂ Ratio (Mean ± SEM)", "Ratio")
           ],
           'Mesophyll': [
               ("mesophyll_conductance_for_CO2", "Mesophyll Conductance (Mean ± SEM)", "Conductance (mmol/m²s)"),
               ("uncorr_gm_prima", "Uncorrected Mesophyll Conductance (Mean ± SEM)", "Conductance (mmol/m²s)"),
               ("corr_uncorr_gm_prima", "Corr/Uncorr Mesophyll Conductance Ratio (Mean ± SEM)", "Ratio"),
               ("CO2_comp_point", "CO₂ Compensation Point (Mean ± SEM)", "CO₂ (μmol CO₂/mol air)"),
               ("uncorr_gamma", "Uncorrected CO₂ Compensation Point (Mean ± SEM)", "CO₂ (μmol CO₂/mol air)"),
               ("corr_uncorr_gamma", "Corr/Uncorr Compensation Point Ratio (Mean ± SEM)", "Ratio")
           ],
           'Efficiency': [
               ("WUEi", "Intrinsic Water Use Efficiency (Mean ± SEM)", "WUEi (μmol CO₂ / mol H₂O)"),
               ("WUE", "Water Use Efficiency (Mean ± SEM)", "WUE (μmol CO₂ / mmol H₂O)")
           ]
       }

        for category_name, metrics_list in plot_categories.items():
            # Create a scrollable frame for each category
            tab = ttk.Frame(self.metric_category_notebook)
            self.metric_category_notebook.add(tab, text=category_name)

            # CRITICAL: Configure tab to expand BOTH directions
            tab.columnconfigure(0, weight=1)
            tab.rowconfigure(0, weight=1)

            # Create scrollable container
            container = ttk.Frame(tab)
            container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
            container.columnconfigure(0, weight=1)
            container.rowconfigure(0, weight=1)

            canvas = tk.Canvas(container, highlightthickness=0)
            v_scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
            h_scrollbar = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)

            scrollable_frame = ttk.Frame(canvas)

            scrollable_frame.bind(
                "<Configure>",
                lambda e, c=canvas: c.configure(scrollregion=c.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

            canvas.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
            v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
            h_scrollbar.grid(row=1, column=0, sticky=(tk.E, tk.W))

            # Store the scrollable frame and configs for later updates
            self.metric_category_frames[category_name] = scrollable_frame
            self.metric_category_configs[category_name] = metrics_list

            # Configure container to expand
            container.columnconfigure(0, weight=1)
            container.rowconfigure(0, weight=1)

        # ============ TAB 3: WUEi Distribution Plot ============
        self.wuei_dist_tab = ttk.Frame(self.batch_results_notebook)
        self.batch_results_notebook.add(self.wuei_dist_tab, text="WUEi Distribution")

        # Configure WUEi distribution tab to expand
        self.wuei_dist_tab.columnconfigure(0, weight=1)
        self.wuei_dist_tab.rowconfigure(0, weight=1)

        # Create scrollable frame for this tab
        canvas = tk.Canvas(self.wuei_dist_tab, highlightthickness=0)
        v_scrollbar = ttk.Scrollbar(self.wuei_dist_tab, orient="vertical", command=canvas.yview)
        h_scrollbar = ttk.Scrollbar(self.wuei_dist_tab, orient="horizontal", command=canvas.xview)

        self.wuei_dist_scrollable_frame = ttk.Frame(canvas)
        self.wuei_dist_scrollable_frame.bind("<Configure>", 
           lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=self.wuei_dist_scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        canvas.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.E, tk.W))

        self.wuei_dist_tab.columnconfigure(0, weight=1)
        self.wuei_dist_tab.rowconfigure(0, weight=1)

        # Show initial message
        self.show_wuei_dist_placeholder_message()

    def show_wuei_dist_placeholder_message(self):
        """Show placeholder message in WUEi distribution tab before batch processing"""
        for widget in self.wuei_dist_scrollable_frame.winfo_children():
            widget.destroy()
    
        ttk.Label(self.wuei_dist_scrollable_frame, 
             text="Run batch processing to generate WUEi distribution plot\n\n"
                  "Click 'Process Batch' after selecting your files.",
             font=("Arial", 12), foreground="gray", justify=tk.CENTER).pack(pady=50, expand=True)
    def setup_batch_plots_scrollable_frame(self):
        """Setup the scrollable frame for batch plots (for compatibility with update_batch_plots_tab)"""
        # Create a frame inside the metric_plots_tab to hold all plots
        self.batch_plots_frame = ttk.Frame(self.metric_plots_tab)
        self.batch_plots_frame.grid(row=2, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), padx=10, pady=10)
    
        # CRITICAL: Configure the batch_plots_frame to expand in BOTH directions
        self.batch_plots_frame.columnconfigure(0, weight=1)
        self.batch_plots_frame.rowconfigure(0, weight=1)
    
        # Configure grid weights for the metric_plots_tab to allow expansion
        self.metric_plots_tab.columnconfigure(0, weight=1)
        self.metric_plots_tab.rowconfigure(1, weight=0)  # Row 1 (metric_category_notebook) gets weight
        self.metric_plots_tab.rowconfigure(2, weight=1)  # Row 2 is the batch_plots_frame row

    def setup_batch_channel_selector(self):
        """Setup channel selector at the top of the Metric Plots tab"""
        self.batch_selector_frame = ttk.LabelFrame(self.metric_plots_tab, text="Channel Selection", padding="10")
        self.batch_selector_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=5)
    
        # Instructions
        ttk.Label(self.batch_selector_frame, text="Select channels to display in metric plots:", 
                  font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))
    
        # Buttons for select/deselect all
        btn_frame = ttk.Frame(self.batch_selector_frame)
        btn_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))
    
        self.batch_select_all_btn = ttk.Button(btn_frame, text="Select All", 
                                               command=self.batch_select_all_channels, width=12)
        self.batch_select_all_btn.pack(side=tk.LEFT, padx=(0, 10))
    
        self.batch_deselect_all_btn = ttk.Button(btn_frame, text="Deselect All", 
                                                 command=self.batch_deselect_all_channels, width=12)
        self.batch_deselect_all_btn.pack(side=tk.LEFT)
    
    # Frame for checkboxes (will be populated when batch results are loaded)
        self.batch_checkbox_frame = ttk.Frame(self.batch_selector_frame)
        self.batch_checkbox_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E))
    
        # Store channel checkboxes and variables
        self.batch_channel_vars = {}
        self.batch_selected_channels = []  # List of selected channel names
    
        # Add a separator for visual clarity
        separator = ttk.Separator(self.batch_selector_frame, orient='horizontal')
        separator.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
    
        # Configure grid
        self.batch_selector_frame.columnconfigure(0, weight=1)
    
        # Initially show a message that no channels are loaded
        ttk.Label(self.batch_checkbox_frame, text="No channels loaded yet. Run batch processing first.", 
                  foreground="gray").grid(row=0, column=0, pady=10)

    def setup_batch_metric_tabs(self):
        """Setup individual metric tabs for batch results matching Interactive Plots structure"""
        # Create a notebook for batch metric plots (similar to plots_notebook)
        self.batch_metric_notebook = ttk.Notebook(self.batch_results_notebook)
        self.batch_results_notebook.add(self.batch_metric_notebook, text="Metric Plots")

        # Define the same categories as in Interactive Plots
        plot_categories = {
            'Environment': [
                ("air_temp", "Air Temperature", "Temperature (ºC)"),
                ("saturating_air_humidity", "Saturation Humidity", "Saturation Humidity (mmol/mol)"),
                ("absorbed_radiation", "Absorbed Radiation", "Radiation (cal/cm²s)"),
                ("relative_air_humidity", "Relative Air Humidity", "Humidity (%)"),
                ("VPD", "Vapor Pressure Deficit", "VPD (kPa)"),
                ("rad_ind_leaf_temp_increase", "Radiation-induced Leaf Temperature Increase", "Temperature Increase (ºC)")
            ],
            'Base_Fluxes': [
                ("CO2_exchange_rate", "CO₂ Exchange Rate", "CO₂ Exchange (μmol/m²s)"),
                ("Transpiration_H2O_evol_rate", "Transpiration Rate", "Transpiration (mmol/m²s)"),
                ("transp_ind_leaf_temp_depr", "Transpiration-induced Leaf Temperature Depression", "Temperature Depression (ºC)"),
                ("leaf_temp_c", "Leaf Temperature", "Leaf Temperature (ºC)")
            ],
            'Humidity_Vapor': [
                ("leaf_air_hum_grad", "Leaf Air Humidity Gradient", "Gradient (mmol/mol)"),
                ("satur_hum_at_leaf_temp_c", "Saturation Humidity at Leaf Temperature", "Saturation Humidity (mmol/mol)"),
                ("mean_vapor_pressure", "Mean Vapor Pressure", "Pressure (mbar)")
            ],
            'Conductances': [
                ("overall_conductance", "Overall Conductance", "Conductance (mmol/m²s)"),
                ("corr_d_overall_conductance", "Corrected Overall Conductance", "Conductance (cm/s)"),
                ("massflow_correction_for_overall_resistance", "Massflow Correction", "Correction Factor"),
                ("b_layer_conductance", "Boundary Layer Conductance", "Conductance (cm/s)"),
                ("corr_leaf_conductance", "Corrected Leaf Conductance", "Conductance (cm/s)"),
                ("corr_stomatal_conductance", "Corrected Stomatal Conductance", "Conductance (cm/s)"),
                ("Stomatal_conductance_corrected", "Stomatal Conductance Corrected", "Conductance (mmol/m²s)"),
                ("normalised_corr_stomatal_conductance", "Normalised Stomatal Conductance", "Normalised Conductance"),
                ("corr_uncorr_stom_cond", "Corr/Uncorr Stomatal Conductance Ratio", "Ratio"),
                ("corr_stomatal_change_rate", "Corrected Stomatal Change Rate", "Change Rate (mmol/m²s)"),
                ("relat_stom_cond_change_rate", "Relative Stomatal Conductance Change Rate", "Change Rate")
            ],
            'Ozone': [
                ("overall_O3_uptake_rate_c", "Overall O₃ Uptake Rate", "Uptake Rate (nmol/m²s)"),
                ("corr_conductance_for_O3_c", "Corrected Conductance for O₃", "Conductance (mmol/m²s)"),
                ("corr_stomatal_O3_uptake_rate", "Stomatal O₃ Uptake", "O₃ Uptake (nmol O₃/m²s)"),
                ("corr_cumul_O3_dose", "Cumulative O₃ Dose", "O₃ Dose (μmol/m²)")
            ],
            'CO2': [
                ("massflow_corr_d_ca", "Massflow Corrected Ambient CO₂", "CO₂ (μmol CO₂/mol air)"),
                ("massflow_corr_for_CO2", "Massflow Correction for CO₂", "Correction Factor"),
                ("mflowcorr_Ca_gradient", "Massflow Correction for CO₂ Gradient", "Correction Factor"),
                ("corr_uncorr_ca", "Corr/Uncorr Ambient CO₂ Ratio", "Ratio"),
                ("stom_resist_to_CO2", "Stomatal Resistance to CO₂", "Resistance (m²s/mmol)"),
                ("b_layer_resist_to_CO2", "Boundary Layer Resistance to CO₂", "Resistance (m²s/mmol)"),
                ("corr_ted_CO2_grad", "Corrected CO₂ Gradient", "CO₂ Gradient (μmol CO₂/mol air)"),
                ("intercell_CO2_conc_in_gas_phase", "Intercellular CO₂", "CO₂ (μmol CO₂/mol air)"),
                ("uncorr_ci", "Uncorrected Intercellular CO₂", "CO₂ (μmol CO₂/mol air)"),
                ("corr_uncorr_ci", "Corr/Uncorr Intercellular CO₂ Ratio", "Ratio")
            ],
            'Mesophyll': [
                ("mesophyll_conductance_for_CO2", "Mesophyll Conductance", "Conductance (mmol/m²s)"),
                ("uncorr_gm_prima", "Uncorrected Mesophyll Conductance", "Conductance (mmol/m²s)"),
                ("corr_uncorr_gm_prima", "Corr/Uncorr Mesophyll Conductance Ratio", "Ratio"),
                ("CO2_comp_point", "CO₂ Compensation Point", "CO₂ (μmol CO₂/mol air)"),
                ("uncorr_gamma", "Uncorrected CO₂ Compensation Point", "CO₂ (μmol CO₂/mol air)"),
                ("corr_uncorr_gamma", "Corr/Uncorr Compensation Point Ratio", "Ratio")
            ],
            'Efficiency': [
                ("WUEi", "Intrinsic Water Use Efficiency", "WUEi (μmol CO₂ / mol H₂O)"),
                ("WUE", "Water Use Efficiency", "WUE (μmol CO₂ / mmol H₂O)")
            ]
        }

        # Store references to each tab's frame for later updates
        self.batch_metric_frames = {}
        self.batch_metric_configs = plot_categories
        self.batch_metric_canvases = {}
        self.batch_metric_canvas_windows = {}

        for category_name, metrics_list in plot_categories.items():
            # Create a tab
            tab = ttk.Frame(self.batch_metric_notebook)
            self.batch_metric_notebook.add(tab, text=category_name)

            # CRITICAL: Configure tab to expand BOTH directions
            tab.columnconfigure(0, weight=1)
            tab.rowconfigure(0, weight=1)

            # Create container that fills the entire tab
            container = ttk.Frame(tab)
            container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
            container.columnconfigure(0, weight=1)
            container.rowconfigure(0, weight=1)

            # Create canvas that fills the container
            canvas = tk.Canvas(container, highlightthickness=0)
            canvas.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

            # Add scrollbars
            v_scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
            v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

            h_scrollbar = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
            h_scrollbar.grid(row=1, column=0, sticky=(tk.E, tk.W))

            # Create scrollable frame inside canvas
            scrollable_frame = ttk.Frame(canvas)

            # Make scrollable frame expand horizontally only
            scrollable_frame.columnconfigure(0, weight=1)

            # ============ CRITICAL FIX: Function to update scroll region when content changes ============
            def update_scroll_region(event, c=canvas, sf=scrollable_frame, cat=category_name):
                # Update scroll region to include all content
                c.configure(scrollregion=c.bbox("all"))
                # Update canvas window height to match content
                if cat in self.batch_metric_canvas_windows:
                    window_id = self.batch_metric_canvas_windows[cat]
                    c.itemconfig(window_id, height=sf.winfo_reqheight())

            scrollable_frame.bind("<Configure>", update_scroll_region)

            # Create window in canvas - initially with width only, height will be set by content
            canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", 
                                                 width=canvas.winfo_width())

            # ============ CRITICAL FIX: Update canvas window when canvas resizes ============
            def on_canvas_resize(event, c=canvas, w_id=canvas_window, sf=scrollable_frame, cat=category_name):
                # Update the canvas window width when canvas is resized
                c.itemconfig(w_id, width=c.winfo_width())
                # Update scroll region
                c.configure(scrollregion=c.bbox("all"))
                # Update height to match content
                c.itemconfig(w_id, height=sf.winfo_reqheight())

            canvas.bind("<Configure>", on_canvas_resize)

            # Store references
            self.batch_metric_canvas_windows[category_name] = canvas_window
            canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

            self.batch_metric_frames[category_name] = scrollable_frame
            self.batch_metric_canvases[category_name] = canvas

    def update_batch_metric_canvas_height(self, category_name):
        """Force update of canvas height for a specific metric category"""
        if category_name in self.batch_metric_canvases:
            canvas = self.batch_metric_canvases[category_name]
            frame = self.batch_metric_frames[category_name]
            window_id = self.batch_metric_canvas_windows.get(category_name)
        
            if window_id:
                # Force the frame to update its layout
                frame.update_idletasks()
                # Update the canvas window height to match the frame's content
                canvas.itemconfig(window_id, height=frame.winfo_reqheight())
                # Update scroll region
                canvas.configure(scrollregion=canvas.bbox("all"))
    
    def update_batch_distribution_plots(self, all_processed_data):
        """Update the distribution plots tab with violin + boxplot for multiple metrics"""
        if not all_processed_data:
            return
    
        # Clear existing widgets
        for widget in self.batch_distribution_frame.winfo_children():
            widget.destroy()
    
        # Combine all processed data
        combined_df = pd.concat(all_processed_data, ignore_index=True)
    
        # Apply base name extraction to Channel column
        combined_df['Channel'] = combined_df['Channel'].astype(str).str.replace(r'\s*\(\d+\)$', '', regex=True)
    
        # Define metrics to plot distributions for
        distribution_metrics = [
            ("WUEi", "Intrinsic Water Use Efficiency", "WUEi (μmol CO₂ / mol H₂O)"),
            ("WUE", "Water Use Efficiency", "WUE (μmol CO₂ / mmol H₂O)"),
            ("CO2_exchange_rate", "CO₂ Exchange Rate", "CO₂ Exchange (μmol/m²s)"),
            ("Stomatal_conductance_corrected", "Stomatal Conductance", "Conductance (mmol/m²s)"),
            ("Transpiration_H2O_evol_rate", "Transpiration Rate", "Transpiration (mmol/m²s)"),
            ("VPD", "Vapor Pressure Deficit", "VPD (kPa)")
        ]
    
        # Get palette for colors
        palette = self.palette_var.get() if hasattr(self, 'palette_var') else 'tab10'
    
        row = 0
        for metric, title, ylabel in distribution_metrics:
            if metric not in combined_df.columns:
                print(f"Warning: {metric} not found in data, skipping distribution plot")
                continue
        
            # Create the combined violin + boxplot
            fig = self.create_batch_distribution_plot(combined_df, metric, title, ylabel, palette)
            fig.set_size_inches(12, 8)
        
            # Create container frame for this plot
            plot_container = ttk.Frame(self.batch_distribution_frame)
            plot_container.grid(row=row, column=0, ticky=(tk.N, tk.S, tk.E, tk.W), padx=10, pady=(20, 0), columnspan=2)
        
            # Add title label
            ttk.Label(plot_container, text=title, font=("Arial", 11, "bold")).grid(
                row=0, column=0, sticky=tk.W)
        
            # Create canvas for the plot
            canvas = FigureCanvasTkAgg(fig, plot_container)
            canvas.draw()
            canvas.get_tk_widget().grid(row=1, column=0, ticky=(tk.N, tk.S, tk.E, tk.W), pady=(5, 0))
        
            # Add download button
            btn = ttk.Button(plot_container, text=f"Download {title} Distribution",
                            command=lambda f=fig, t=title: self.download_plot(f, t))
            btn.grid(row=2, column=0, sticky=tk.W, pady=(5, 20))
        
            plot_container.columnconfigure(0, weight=1)
            row += 1
    
        self.batch_distribution_frame.columnconfigure(0, weight=1)

    def create_batch_distribution_plot(self, df, metric, title, ylabel, palette='tab10'):
        """Create a combined violin + boxplot for distribution analysis"""
    
        # Close any existing figure
        plt.close('all')
    
        if df is None or len(df) == 0 or metric not in df.columns:
            fig, ax = plt.subplots(figsize=(12, 6))
            ax.text(0.5, 0.5, 'No data available', ha='center', va='center', fontsize=12)
            ax.set_axis_off()
            return fig
    
        fig, ax = plt.subplots(figsize=(12, 8))
    
        # Prepare data for each channel
        channels = df['Channel'].unique()
        metric_data = []
        channel_labels = []
    
        for channel in sorted(channels):
            channel_data = df[df['Channel'] == channel]
            values = channel_data[metric].dropna().values
            if len(values) > 0:
                metric_data.append(values)
                channel_labels.append(str(channel))
    
        if not metric_data:
            ax.text(0.5, 0.5, f'No {metric} data available', ha='center', va='center', fontsize=12)
            ax.set_axis_off()
            return fig
    
        # Get colors
        colors = plt.cm.tab10(np.linspace(0, 1, len(metric_data)))
        positions = range(1, len(metric_data) + 1)
    
        # --- LAYER 1: VIOLIN PLOT (background) ---
        violin_parts = ax.violinplot(metric_data, positions=positions, 
                                     showmeans=False, showmedians=False, showextrema=False)
    
        for i, body in enumerate(violin_parts['bodies']):
            body.set_facecolor(colors[i])
            body.set_alpha(0.3)
            body.set_edgecolor(colors[i])
            body.set_linewidth(1)
            body.set_zorder(1)
    
        # --- LAYER 2: BOXPLOT (foreground) ---
        box_plot = ax.boxplot(metric_data, positions=positions, patch_artist=True)
    
        for i, (box, color) in enumerate(zip(box_plot['boxes'], colors)):
            box.set_facecolor(color)
            box.set_alpha(0.7)
            box.set_edgecolor('black')
            box.set_linewidth(1.5)
            box.set_zorder(3)
    
        for median in box_plot['medians']:
            median.set_color('red')
            median.set_linewidth(2.5)
            median.set_zorder(4)
    
        for whisker in box_plot['whiskers']:
            whisker.set(color='black', linewidth=1.5, zorder=2)
        for cap in box_plot['caps']:
            cap.set(color='black', linewidth=1.5, zorder=2)
    
        # --- LAYER 3: DATA POINTS (top layer) ---
        for i, data in enumerate(metric_data):
            x_jitter = np.random.normal(positions[i], 0.08, size=len(data))
            ax.scatter(x_jitter, data, alpha=0.5, color=colors[i], s=20, 
                      edgecolor='black', linewidth=0.5, zorder=5)
    
        # --- LAYER 4: MEAN MARKERS ---
        for i, data in enumerate(metric_data):
            mean_val = np.mean(data)
            ax.scatter(positions[i], mean_val, color='yellow', s=100, 
                      edgecolor='black', linewidth=1.5, zorder=6, marker='D')
    
        # Configure plot
        ax.set_title(f'{title} Distribution by Channel', fontsize=14, fontweight='bold')
        ax.set_xlabel('Channel', fontsize=12)
        ax.set_ylabel(ylabel, fontsize=12)
        ax.set_xticks(positions)
        ax.set_xticklabels(channel_labels, rotation=45, ha='right', fontsize=10)
        ax.grid(True, alpha=0.2, linestyle='--', axis='y', zorder=0)
    
        # Add statistics annotations
        for i, data in enumerate(metric_data):
            median = np.median(data)
            q1 = np.percentile(data, 25)
            q3 = np.percentile(data, 75)
        
            ax.text(positions[i], median, f'{median:.1f}', 
                    ha='center', va='center', fontsize=9, fontweight='bold',
                    bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.8, edgecolor='none'),
                    zorder=7)
        
            ax.text(positions[i], q1, f'Q1', ha='center', va='top', fontsize=8, alpha=0.7)
            ax.text(positions[i], q3, f'Q3', ha='center', va='bottom', fontsize=8, alpha=0.7)
    
        # Legend
    
        legend_elements = [
            Patch(facecolor='gray', alpha=0.3, edgecolor='gray', label='Violin (distribution shape)'),
            Rectangle((0, 0), 1, 1, facecolor='gray', alpha=0.7, edgecolor='black', label='Box (quartiles)'),
            Line2D([0], [0], color='red', linewidth=2.5, label='Median'),
            Line2D([0], [0], marker='o', color='w', markerfacecolor='gray', markersize=8, label='Data points'),
            Line2D([0], [0], marker='D', color='w', markerfacecolor='yellow', markersize=8, label='Mean'),
        ]
    
        ax.legend(handles=legend_elements, loc='upper left', 
                  bbox_to_anchor=(1.02, 1), fontsize=9, framealpha=0.9)
    
        plt.tight_layout()
        plt.subplots_adjust(right=0.85)
    
        return fig

    def update_batch_metric_plots(self, stats_by_channel):
        """Update all metric plots in the batch results tabs"""
        if not stats_by_channel:
            return

        # Filter by selected channels
        filtered_stats = {}
        for channel in self.batch_selected_channels:
            if channel in stats_by_channel:
                filtered_stats[channel] = stats_by_channel[channel]

        if not filtered_stats:
            # Show message in all metric tabs
            for category_name, frame in self.batch_metric_frames.items():
                for widget in frame.winfo_children():
                    widget.destroy()
                fig, ax = plt.subplots(figsize=(12, 6))
                ax.text(0.5, 0.5, 'No channels selected\n\nUse the channel selector above to choose channels', 
                    ha='center', va='center', fontsize=12, fontweight='bold')
                ax.set_axis_off()
                canvas = FigureCanvasTkAgg(fig, frame)
                canvas.draw()
                canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.W, tk.E))
            
                # Update canvas scroll region after adding content
                if category_name in self.batch_metric_canvases:
                    canvas_ref = self.batch_metric_canvases[category_name]
                    canvas_ref.configure(scrollregion=canvas_ref.bbox("all"))
                    # Update canvas window width
                    if category_name in self.batch_metric_canvas_windows:
                        canvas_ref.itemconfig(
                            self.batch_metric_canvas_windows[category_name], 
                            width=canvas_ref.winfo_width()
                        )
            return

        # Get palette
        palette = self.palette_var.get() if hasattr(self, 'palette_var') else 'tab10'

        # Update each category tab
        for category_name, metrics_list in self.batch_metric_configs.items():
            frame = self.batch_metric_frames[category_name]
        
            # Clear existing widgets
            for widget in frame.winfo_children():
                widget.destroy()
        
            row = 0
            for metric, title, ylabel in metrics_list:
                # Check if metric exists in any channel
                metric_exists = False
                for channel, stats in filtered_stats.items():
                    for time_point in stats.values():
                        if metric in time_point:
                            metric_exists = True
                            break
                    if metric_exists:
                        break
            
                if not metric_exists:
                    continue
            
                # Create figure
                fig, ax = plt.subplots(figsize=(14, 6))
            
                # Sort channels
                sorted_channels = sorted(filtered_stats.keys())
                num_channels = len(sorted_channels)
            
                # Get colors
                colors = plt.cm.tab10(np.linspace(0, 1, num_channels))
            
                # Plot data for each channel with error bars
                for i, channel in enumerate(sorted_channels):
                    stats = filtered_stats[channel]
                    time_points = sorted(stats.keys())
                    means = []
                    errors = []
                    valid_times = []
                
                    for t in time_points:
                        if metric in stats[t]:
                            means.append(stats[t][metric]['mean'])
                            errors.append(stats[t][metric]['sem'])
                            valid_times.append(t)
                
                    if means:
                        ax.errorbar(valid_times, means, yerr=errors, 
                                   label=channel, marker='o', capsize=3, 
                                   linewidth=2, markersize=5, color=colors[i], alpha=0.7)
            
                ax.set_title(title, fontsize=14, fontweight='bold')
                ax.set_xlabel('Time (minutes)', fontsize=12)
                ax.set_ylabel(ylabel, fontsize=12)
                ax.legend(bbox_to_anchor=(1.02, 1), loc='upper left', fontsize=9)
                ax.grid(True, alpha=0.3, linestyle='--')
            
                plt.tight_layout()
                plt.subplots_adjust(right=0.85)
            
                # Embed plot
                plot_container = ttk.Frame(frame)
                # Use sticky=(tk.W, tk.E, tk.N, tk.S) for full expansion
                plot_container.grid(row=row, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), 
                                   padx=10, pady=10, columnspan=2)
            
                ttk.Label(plot_container, text=title, font=("Arial", 10, "bold")).grid(
                    row=0, column=0, sticky=tk.W)
            
                canvas = FigureCanvasTkAgg(fig, plot_container)
                canvas.draw()
                canvas.get_tk_widget().grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5, 0))
            
                btn = ttk.Button(plot_container, text=f"Download {title}",
                                command=lambda f=fig, t=title: self.download_plot(f, t))
                btn.grid(row=2, column=0, sticky=tk.W, pady=(5, 20))
            
                plot_container.columnconfigure(0, weight=1)
                plot_container.rowconfigure(1, weight=1)  # Allow the canvas row to expand
                row += 1
        
            # After adding all plots, configure the frame to expand and update scroll region
            frame.columnconfigure(0, weight=1)
            frame.rowconfigure(0, weight=1)  # Add row weight for expansion

            # Force the frame to update its geometry
            frame.update_idletasks()

            # Update canvas scroll region and window size
            if category_name in self.batch_metric_canvases:
                canvas_ref = self.batch_metric_canvases[category_name]
                canvas_ref.configure(scrollregion=canvas_ref.bbox("all"))
    
                # Update the canvas window to fill the canvas width
                window_id = self.batch_metric_canvas_windows.get(category_name)
                if window_id:
                    # Set the canvas window width to match the canvas width
                    canvas_ref.itemconfig(window_id, width=canvas_ref.winfo_width())
                    # Also set height to match content
                    canvas_ref.itemconfig(window_id, height=frame.winfo_reqheight())

    def generate_batch_statistics(self, all_results, results_by_metric):
        """Generate statistics from batch results with proper channel naming - group by base name"""
        output_dir = self.output_dir_var.get()

        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)

        # ============ KEY FIX: Map each result to base channel names ============
        all_processed_data = []

        for result in all_results:
            df = result['calculated_metrics'].copy()

            # Apply base name extraction to Channel column
            df['Channel_original'] = df['Channel']  # Keep original for reference
            df['Channel'] = df['Channel'].astype(str).str.replace(r'\s*\(\d+\)$', '', regex=True)

            # Also add replicate identifier to track per-replicate data if needed
            df['replicate'] = result['replicate']

            # IMPORTANT: Round time_repeated to avoid floating point precision issues
            if 'time_repeated' in df.columns:
                # Round to 6 decimal places to ensure times that should be equal are treated as equal
                df['time_repeated'] = df['time_repeated'].round(6)

            all_processed_data.append(df)

        # Combine all results into a single dataframe with base channel names
        combined_df = pd.concat(all_processed_data, ignore_index=True)

        # IMPORTANT: Ensure Channel column is string and clean
        combined_df['Channel'] = combined_df['Channel'].astype(str).str.strip()

        # Print debug info for verification
        print(f"\n=== BATCH STATISTICS DEBUG ===")
        print(f"Base channel names: {combined_df['Channel'].unique().tolist()}")

        # Show what's happening for a specific metric (WUEi)
        for channel in combined_df['Channel'].unique():
            channel_data = combined_df[combined_df['Channel'] == channel]
            print(f"\nChannel {channel}:")
            print(f"  Replicates: {channel_data['replicate'].unique().tolist()}")
            print(f"  Time points: {sorted(channel_data['time_repeated'].unique())}")
            if 'WUEi' in channel_data.columns:
                # Show WUEi values for first few time points
                for time_point in sorted(channel_data['time_repeated'].unique())[:3]:
                    values = channel_data[channel_data['time_repeated'] == time_point]['WUEi'].values
                    print(f"    Time {time_point:.1f}: WUEi values = {values}")

        # Store channel renaming info for plots (base names)
        self.batch_channel_mapping = {}
        for result in all_results:
            if result.get('channel_mapping'):
                # Map to base names
                for orig, new in result['channel_mapping'].items():
                    base_name = re.sub(r'\s*\(\d+\)$', '', new)
                    self.batch_channel_mapping[orig] = base_name

        # Calculate statistics for each metric by BASE CHANNEL NAME and time_repeated
        numeric_columns = combined_df.select_dtypes(include=[np.number]).columns.tolist()
        exclude_columns = ['Channel', 'replicate', 'Channel_original', 'time_sec', 'time_full_sec', 'time_full_minutes']
        metric_columns = [col for col in numeric_columns if col not in exclude_columns]

        # Group by Base Channel and time_repeated
        stats_by_base_channel = {}
        channels = combined_df['Channel'].unique()

        # Sort channels alphabetically for consistent ordering
        for channel in sorted(channels):
            channel_data = combined_df[combined_df['Channel'] == channel]
            if len(channel_data) == 0:
                continue

            # Group by time_repeated (now with proper rounding)
            grouped = channel_data.groupby('time_repeated')

            stats_dict = {}
            for time_point, group in grouped:
                stats_dict[time_point] = {}
                for metric in metric_columns:
                    if metric in group.columns:
                        values = group[metric].dropna()
                        if len(values) > 0:
                            stats_dict[time_point][metric] = {
                                'mean': values.mean(),
                                'std': values.std(),
                                'n': len(values),
                                'sem': values.std() / np.sqrt(len(values)) if len(values) > 1 else np.nan
                            }

            stats_by_base_channel[channel] = stats_dict

            # Print debug for WUEi
            if 'WUEi' in metric_columns:
                print(f"\nChannel {channel} - WUEi statistics:")
                for time_point in sorted(stats_dict.keys())[:5]:
                    if 'WUEi' in stats_dict[time_point]:
                        s = stats_dict[time_point]['WUEi']
                        print(f"  Time {time_point:.1f}: mean={s['mean']:.2f}, sem={s['sem']:.2f}, n={s['n']}")

        # Save statistics to Excel with base channel names
        self.save_statistics_to_excel(stats_by_base_channel, output_dir)

        # Generate averaged plots with error bars using base channel names
        self.generate_averaged_plots(stats_by_base_channel, output_dir)

        # Generate batch report with base channel renaming info
        self.generate_batch_report(all_results, stats_by_base_channel, output_dir)

        # CRITICAL: Store stats for later use and update the UI
        self.batch_stats = stats_by_base_channel

        # Update batch results tabs with base channel names
        self.update_batch_results_tabs(stats_by_base_channel)

        # ============ Generate WUEi distribution plot (also export) ============
        self.generate_wuei_distribution_plot(all_processed_data, output_dir)

        # ============ NEW: Store combined data for darkness detection ============
        self.batch_all_processed_data = all_processed_data
        self.batch_combined_df = combined_df

        # ============ NEW: Generate darkness detection plots for batch results ============
        self.update_batch_darkness_plots(stats_by_base_channel, combined_df)

        # ============ FIX: Force update WUEi distribution tab after everything is loaded ============
        # This ensures the tab updates with all channels selected by default
        self.root.after(500, self.update_wuei_distribution_tab)

        # Close all figures to free memory
        plt.close('all')

        return stats_by_base_channel


    def show_empty_batch_message(self):
        """Show empty message in batch result tabs"""
        # Clear channel selector
        if hasattr(self, 'batch_checkbox_frame'):
            for widget in self.batch_checkbox_frame.winfo_children():
                widget.destroy()
            ttk.Label(self.batch_checkbox_frame, text="No channels loaded yet. Run batch processing first.", 
                      foreground="gray").grid(row=0, column=0, pady=10)
    
        # Clear metric category tabs
        for category_name, frame in self.metric_category_frames.items():
            for widget in frame.winfo_children():
                widget.destroy()
            fig, ax = plt.subplots(figsize=(12, 6))
            ax.text(0.5, 0.5, 'No data available - run batch processing first', 
                    ha='center', va='center', fontsize=14, fontweight='bold')
            ax.set_axis_off()
            canvas = FigureCanvasTkAgg(fig, frame)
            canvas.draw()
            canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.W, tk.E))
    
        # Clear distribution tab
        if hasattr(self, 'batch_distribution_frame'):
            for widget in self.batch_distribution_frame.winfo_children():
                widget.destroy()
            fig, ax = plt.subplots(figsize=(12, 6))
            ax.text(0.5, 0.5, 'No data available - run batch processing first', 
                    ha='center', va='center', fontsize=14, fontweight='bold')
            ax.set_axis_off()
            canvas = FigureCanvasTkAgg(fig, self.batch_distribution_frame)
            canvas.draw()
            canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.W, tk.E))

    def populate_batch_channel_selector(self, stats_by_channel):
        """Populate the channel selector with checkboxes for each channel"""
        # Clear existing checkboxes
        for widget in self.batch_checkbox_frame.winfo_children():
            widget.destroy()

        # Clear stored variables
        self.batch_channel_vars = {}

        # Get sorted channels
        channels = sorted(stats_by_channel.keys())

        if not channels:
            ttk.Label(self.batch_checkbox_frame, text="No channels available", 
                      foreground="gray").grid(row=0, column=0, pady=10)
            self.batch_selected_channels = []
            return

        # Create checkboxes in a grid layout
        max_per_row = 6
        for idx, channel in enumerate(channels):
            row = idx // max_per_row
            col = idx % max_per_row
        
            # FIX: Set all channels to selected by default
            var = tk.BooleanVar(value=True)
            self.batch_channel_vars[channel] = var
        
            cb = ttk.Checkbutton(self.batch_checkbox_frame, text=channel, variable=var,
                                command=lambda ch=channel: self.on_batch_channel_toggle(ch))
            cb.grid(row=row, column=col, sticky=tk.W, padx=5, pady=2)

        # Update selected channels list - FIX: Include all channels
        self.batch_selected_channels = channels.copy()

        # Configure grid columns
        for i in range(max_per_row):
            self.batch_checkbox_frame.columnconfigure(i, weight=1)
    
        # FIX: Force update of WUEi distribution tab after populating channels
        self.update_wuei_distribution_tab()


    def on_batch_channel_toggle(self, channel):
        """Handle when a channel checkbox is toggled"""
        # Update the selected channels list
        self.batch_selected_channels = [ch for ch, var in self.batch_channel_vars.items() if var.get()]
    
        print(f"Channel toggled: {channel}")
        print(f"Selected channels: {self.batch_selected_channels}")
    
        # Refresh the metric category plots
        if hasattr(self, 'batch_stats') and self.batch_stats:
            self.update_batch_metric_category_plots(self.batch_stats)
    
        # Refresh the darkness plots
        if hasattr(self, 'batch_stats') and self.batch_stats:
            if hasattr(self, 'batch_combined_df'):
                self.update_batch_darkness_plots(self.batch_stats, self.batch_combined_df)
            else:
                self.update_batch_darkness_plots(self.batch_stats)
    
        # ============ UPDATE THE WUEi DISTRIBUTION TAB ============
        # Force update of the WUEi distribution tab
        self.update_wuei_distribution_tab()
    
    def update_batch_metric_category_plots(self, stats_by_channel):
        """Update all metric category plots with selected channels"""
        if not stats_by_channel:
            return

        # Filter stats_by_channel to only include selected channels
        filtered_stats = {}
        for channel in self.batch_selected_channels:
            if channel in stats_by_channel:
                filtered_stats[channel] = stats_by_channel[channel]

        # If no channels selected, show message in all categories
        if not filtered_stats:
            for category_name, frame in self.metric_category_frames.items():
                for widget in frame.winfo_children():
                    widget.destroy()
        
                    # Create a frame that expands to fill all space
                    msg_container = ttk.Frame(frame)
                    msg_container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
                    msg_container.columnconfigure(0, weight=1)
                    msg_container.rowconfigure(0, weight=1)
        
                    fig, ax = plt.subplots(figsize=(12, 6))
                    ax.text(0.5, 0.5, 'No channels selected\n\nUse the channel selector above to choose channels', 
                            ha='center', va='center', fontsize=12, fontweight='bold')
                    ax.set_axis_off()
                    canvas = FigureCanvasTkAgg(fig, msg_container)
                    canvas.draw()
                    canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        
                    # Update canvas scroll region after adding content
                    if category_name in self.batch_metric_canvases:
                        canvas_ref = self.batch_metric_canvases[category_name]
                        canvas_ref.configure(scrollregion=canvas_ref.bbox("all"))
                        if category_name in self.batch_metric_canvas_windows:
                            canvas_ref.itemconfig(
                                self.batch_metric_canvas_windows[category_name], 
                                width=canvas_ref.winfo_width()
                            )
            return

        # Get the current batch palette
        if hasattr(self, 'batch_palette_var'):
            palette = self.batch_palette_var.get()
        else:
            palette = 'tab10'

        # Helper function to get colors based on palette
        def get_palette_colors(n_colors):
            try:
                if palette in colorblind_friendly_palettes:
                    if isinstance(colorblind_friendly_palettes[palette], list):
                        color_list = colorblind_friendly_palettes[palette]
                        colors = [color_list[i % len(color_list)] for i in range(n_colors)]
                    else:
                        cmap = colorblind_friendly_palettes[palette]
                        colors = cmap(np.linspace(0, 1, n_colors))
                elif palette in plt.colormaps():
                    cmap = plt.get_cmap(palette)
                    colors = cmap(np.linspace(0, 1, n_colors))
                else:
                    colors = plt.cm.tab10(np.linspace(0, 1, n_colors))
                return colors
            except Exception as e:
                return plt.cm.tab10(np.linspace(0, 1, n_colors))

        # Sort selected channels alphabetically
        sorted_channels = sorted(filtered_stats.keys())
        num_channels = len(sorted_channels)
        colors = get_palette_colors(num_channels)

        # Update each category
        for category_name, metrics_list in self.metric_category_configs.items():
            frame = self.metric_category_frames[category_name]

            # Clear existing widgets
            for widget in frame.winfo_children():
                widget.destroy()

            # ============ CRITICAL FIX: Create container that expands BOTH directions ============
            container = ttk.Frame(frame)
            container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
            container.columnconfigure(0, weight=1)
            container.rowconfigure(0, weight=1)  # <-- ADD THIS - allows container to expand vertically

            # Create an inner frame to hold the plots - this is what will push the height
            plots_frame = ttk.Frame(container)
            plots_frame.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E))
            plots_frame.columnconfigure(0, weight=1)
            # DO NOT set row weight on plots_frame - let it determine height from content

            row = 0
        
            # ============ FOR EFFICIENCY TAB: ONLY time-series plots (NO distribution plot here) ============
            if category_name == 'Efficiency':
                # Only add the time-series plots for WUEi and WUE
                for metric, title, ylabel in metrics_list:
                    # Check if metric exists in any selected channel
                    metric_exists = False
                    for channel, stats in filtered_stats.items():
                        for time_point in stats.values():
                            if metric in time_point:
                                metric_exists = True
                                break
                        if metric_exists:
                            break

                    if metric_exists:
                        # Create figure for time-series
                        fig, ax = plt.subplots(figsize=(14, 6))

                        # Plot data for each selected channel with error bars
                        for i, channel in enumerate(sorted_channels):
                            stats = filtered_stats[channel]
                            time_points = sorted(stats.keys())
                            means = []
                            errors = []
                            valid_times = []

                            for t in time_points:
                                if metric in stats[t]:
                                    means.append(stats[t][metric]['mean'])
                                    errors.append(stats[t][metric]['sem'])
                                    valid_times.append(t)

                            if means:
                                ax.errorbar(valid_times, means, yerr=errors, 
                                           label=channel, marker='o', capsize=3, 
                                           linewidth=2, markersize=5, color=colors[i], alpha=0.7)

                        ax.set_title(title, fontsize=14, fontweight='bold')
                        ax.set_xlabel('Time (minutes)', fontsize=12)
                        ax.set_ylabel(ylabel, fontsize=12)
                        ax.legend(bbox_to_anchor=(1.02, 1), loc='upper left', 
                                 fontsize=9, framealpha=0.9, borderaxespad=0)
                        ax.grid(True, alpha=0.3, linestyle='--')
                        plt.tight_layout()
                        plt.subplots_adjust(right=0.85)

                        # Embed plot - use plots_frame instead of container
                        plot_container = ttk.Frame(plots_frame)
                        plot_container.grid(row=row, column=0, sticky=(tk.W, tk.E, tk.N), 
                                           padx=10, pady=10)
                        plot_container.columnconfigure(0, weight=1)
                        # Let the canvas expand, but row should NOT have weight
                        plot_container.rowconfigure(1, weight=0)  # No vertical expansion for individual plots

                        ttk.Label(plot_container, text=title, font=("Arial", 10, "bold")).grid(
                            row=0, column=0, sticky=tk.W)

                        canvas = FigureCanvasTkAgg(fig, plot_container)
                        canvas.draw()
                        canvas.get_tk_widget().grid(row=1, column=0, sticky=(tk.W, tk.E))

                        btn = ttk.Button(plot_container, text=f"Download {title}",
                                        command=lambda f=fig, t=title: self.download_plot(f, t))
                        btn.grid(row=2, column=0, sticky=tk.W, pady=(5, 0))

                        row += 1
        
            else:
                # For other categories, just plot the time-series plots
                for metric, title, ylabel in metrics_list:
                    metric_exists = False
                    for channel, stats in filtered_stats.items():
                        for time_point in stats.values():
                            if metric in time_point:
                                metric_exists = True
                                break
                        if metric_exists:
                            break

                    if not metric_exists:
                        continue

                    fig, ax = plt.subplots(figsize=(14, 6))

                    for i, channel in enumerate(sorted_channels):
                        stats = filtered_stats[channel]
                        time_points = sorted(stats.keys())
                        means = []
                        errors = []
                        valid_times = []
    
                        for t in time_points:
                            if metric in stats[t]:
                                means.append(stats[t][metric]['mean'])
                                errors.append(stats[t][metric]['sem'])
                                valid_times.append(t)

                        if means:
                            ax.errorbar(valid_times, means, yerr=errors, 
                                       label=channel, marker='o', capsize=3, 
                                       linewidth=2, markersize=5, color=colors[i], alpha=0.7)

                    ax.set_title(title, fontsize=14, fontweight='bold')
                    ax.set_xlabel('Time (minutes)', fontsize=12)
                    ax.set_ylabel(ylabel, fontsize=12)
                    ax.legend(bbox_to_anchor=(1.02, 1), loc='upper left', 
                             fontsize=9, framealpha=0.9, borderaxespad=0)
                    ax.grid(True, alpha=0.3, linestyle='--')
                    plt.tight_layout()
                    plt.subplots_adjust(right=0.85)

                    plot_container = ttk.Frame(plots_frame)
                    plot_container.grid(row=row, column=0, sticky=(tk.W, tk.E, tk.N), 
                                       padx=10, pady=10)
                    plot_container.columnconfigure(0, weight=1)
                    plot_container.rowconfigure(1, weight=0)

                    ttk.Label(plot_container, text=title, font=("Arial", 10, "bold")).grid(
                        row=0, column=0, sticky=tk.W)

                    canvas = FigureCanvasTkAgg(fig, plot_container)
                    canvas.draw()
                    canvas.get_tk_widget().grid(row=1, column=0, sticky=(tk.W, tk.E))

                    btn = ttk.Button(plot_container, text=f"Download {title}",
                                    command=lambda f=fig, t=title: self.download_plot(f, t))
                    btn.grid(row=2, column=0, sticky=tk.W, pady=(5, 0))

                    row += 1

                # If no plots were added, show message
                if row == 0:
                    msg_container = ttk.Frame(plots_frame)
                    msg_container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
                    msg_container.columnconfigure(0, weight=1)
                    msg_container.rowconfigure(0, weight=1)
            
                    fig, ax = plt.subplots(figsize=(12, 6))
                    ax.text(0.5, 0.5, f'No data available for {category_name}\n\nRun batch processing with the correct file format to see plots', 
                            ha='center', va='center', fontsize=12, fontweight='bold')
                    ax.set_axis_off()
                    canvas = FigureCanvasTkAgg(fig, msg_container)
                    canvas.draw()
                    canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

            # ============ CRITICAL: Update canvas height after adding all plots ============
            # Force the plots_frame to update its layout
            plots_frame.update_idletasks()
            frame.update_idletasks()
        
            # Update canvas scroll region and window size
            if category_name in self.batch_metric_canvases:
                canvas_ref = self.batch_metric_canvases[category_name]
                canvas_ref.configure(scrollregion=canvas_ref.bbox("all"))
                window_id = self.batch_metric_canvas_windows.get(category_name)
                if window_id:
                    canvas_ref.itemconfig(window_id, width=canvas_ref.winfo_width())
                    canvas_ref.itemconfig(window_id, height=plots_frame.winfo_reqheight())

    
    def detect_darkness_periods_in_batch(self, df):
        """
        Automatically detect darkness periods in batch data
        Returns a dictionary with channel names as keys and lists of (start, end) time ranges
        """
        darkness_periods = {}
    
        if df is None or len(df) == 0:
            return darkness_periods
    
        # Ensure we have the required columns
        if 'absorbed_radiation' not in df.columns or 'time_repeated' not in df.columns:
            print("Required columns for darkness detection not found")
            return darkness_periods
    
        # Process each channel separately
        for channel in df['Channel'].unique():
            channel_data = df[df['Channel'] == channel].copy()
            channel_data = channel_data.sort_values('time_repeated')
        
            # Find points where absorbed_radiation is zero (or very close to zero)
            darkness_mask = channel_data['absorbed_radiation'] <= 1e-6
        
            if not darkness_mask.any():
                continue
        
            # Find continuous blocks of darkness - return FULL periods (start AND end)
            darkness_periods_for_channel = []
            in_darkness = False
            period_start = None
        
            for idx, row in channel_data.iterrows():
                is_dark = row['absorbed_radiation'] <= 1e-6
            
                if is_dark and not in_darkness:
                    # Start of a darkness period
                    in_darkness = True
                    period_start = row['time_repeated']
                elif not is_dark and in_darkness:
                    # End of darkness period - store as dict with start and end
                    in_darkness = False
                    darkness_periods_for_channel.append({
                        'start': period_start,
                        'end': row['time_repeated']
                    })
                    period_start = None
        
            # If still in darkness at the end of data
            if in_darkness and period_start is not None:
                darkness_periods_for_channel.append({
                    'start': period_start,
                    'end': channel_data['time_repeated'].iloc[-1]
                })
        
            if darkness_periods_for_channel:
                darkness_periods[str(channel)] = darkness_periods_for_channel
                print(f"Channel {channel}: Darkness periods: {darkness_periods_for_channel}")
    
        return darkness_periods

    def create_batch_darkness_plot(self, df, yvar, title, ylab, palette='tab10', darkness_periods=None):
        """
        Create a plot for batch results with yellow/black rectangles for darkness periods
        """
        # Close any existing figure to prevent accumulation
        plt.close('all')
    
        if df is None or len(df) == 0 or yvar not in df.columns:
            fig, ax = plt.subplots(figsize=(12, 6))
            ax.text(0.5, 0.5, 'No data available', ha='center', va='center', fontsize=12)
            ax.set_axis_off()
            return fig
    
        fig, ax = plt.subplots(figsize=(12, 6))
    
        # Get unit for y-axis label
        unit = UNITS_MAP.get(yvar, '')
    
        # Create y-axis label with proper formatting
        if unit:
            ylabel = f'{ylab} ({unit})'
        else:
            ylabel = ylab
    
        # Get unique channels
        channels = df['Channel'].unique()
        num_channels = len(channels)

        # Get colors based on palette
        colors = []
        try:
            if palette in colorblind_friendly_palettes:
                if isinstance(colorblind_friendly_palettes[palette], list):
                    color_list = colorblind_friendly_palettes[palette]
                    colors = [color_list[i % len(color_list)] for i in range(num_channels)]
                else:
                    cmap = colorblind_friendly_palettes[palette]
                    colors = cmap(np.linspace(0, 1, num_channels))
            elif palette in plt.colormaps():
                cmap = plt.get_cmap(palette)
                colors = cmap(np.linspace(0, 1, num_channels))
            else:
                colors = plt.cm.tab10(np.linspace(0, 1, num_channels))
        except Exception as e:
            colors = plt.cm.tab10(np.linspace(0, 1, num_channels))
    
        # Ensure we have enough colors
        if len(colors) < num_channels:
            colors = list(colors) * (num_channels // len(colors) + 1)
            colors = colors[:num_channels]
    
        # Calculate typical time interval between measurements
        all_times = []
        for channel in channels:
            channel_data = df[df['Channel'] == channel]
            if len(channel_data) > 0:
                channel_data = channel_data.sort_values('time_repeated')
                times = channel_data['time_repeated'].unique()
                all_times.extend(times.tolist())
    
        if len(all_times) > 1:
            sorted_times = np.sort(np.unique(all_times))
            time_intervals = np.diff(sorted_times)
            positive_intervals = time_intervals[time_intervals > 0]
            if len(positive_intervals) > 0:
                typical_interval = float(np.min(positive_intervals))
            else:
                typical_interval = 1.0
        else:
            typical_interval = 1.0
    
        # Calculate y-range for rectangle positioning
        all_y_vals = []
        for channel in channels:
            channel_data = df[df['Channel'] == channel]
            if yvar in channel_data.columns:
                valid_vals = channel_data[yvar].dropna().values
                all_y_vals.extend(valid_vals)
    
        if len(all_y_vals) > 0:
            y_min = np.nanmin(all_y_vals)
            y_max = np.nanmax(all_y_vals)
            y_range = y_max - y_min if y_max != y_min else 1
        else:
            y_min = 0
            y_max = 1
            y_range = 1
    
        # Handle case where all values are the same
        if len(np.unique(all_y_vals)) == 1:
            single_val = np.unique(all_y_vals)[0]
            y_min = single_val - 0.1 * abs(single_val) - 0.01
            y_max = single_val + 0.1 * abs(single_val) + 0.01
            if y_min == y_max:
                y_min = -0.1
                y_max = 0.1
            y_range = y_max - y_min
    
        # Calculate rectangle positioning
        start_y = y_max + (y_range * 0.05)
        darkness_band_height = y_range * 0.08
        bar_height = darkness_band_height / max(1, num_channels)
    
        # Build lookup for darkness periods
        darkness_lookup = {}
        if darkness_periods:
            for dark_channel, periods in darkness_periods.items():
                dark_channel_str = str(dark_channel).strip()
                darkness_lookup[dark_channel_str] = periods
    
        # Store per-channel data
        channel_data_dict = {}
    
        # FIRST: Plot the data
        for idx, channel in enumerate(channels):
            channel_data = df[df['Channel'] == channel]
            color_idx = idx % len(colors)
        
            # Sort by time_repeated for proper plotting
            channel_data = channel_data.sort_values('time_repeated')
        
            # Store channel data
            rev_idx = num_channels - idx
            channel_str = str(channel).strip()
        
            channel_data_dict[channel_str] = {
                'data': channel_data,
                'color': colors[color_idx],
                'rev_idx': rev_idx,
                'x_min': channel_data['time_repeated'].min() if len(channel_data) > 0 else 0,
                'x_max': channel_data['time_repeated'].max() if len(channel_data) > 0 else 0,
                'sorted_times': np.sort(channel_data['time_repeated'].unique())
            }
        
            # Plot the data
            if yvar in channel_data.columns:
                ax.plot(channel_data['time_repeated'], channel_data[yvar], 
                        label=channel_str, color=colors[color_idx], marker='o', 
                        linestyle='-', linewidth=2, markersize=5, zorder=10)
    
        # SECOND: Add yellow background rectangles for ALL channels
        for channel_str, channel_info in channel_data_dict.items():
            rev_idx = channel_info['rev_idx']
            x_min = channel_info['x_min']
            x_max = channel_info['x_max']
        
            # Calculate y position for this channel
            y_min_rect = start_y + (rev_idx - 1) * bar_height
            y_max_rect = start_y + rev_idx * bar_height
        
            # Get the sorted times for this channel
            sorted_times = channel_info['sorted_times']
            if len(sorted_times) > 1:
                rect_start = x_min
                rect_end = x_max
            
                # Extend to show typical interval beyond last point
                if rect_end > rect_start:
                    rect_end = rect_end + typical_interval
                else:
                    rect_end = rect_start + typical_interval
            else:
                rect_start = x_min
                rect_end = x_min + typical_interval
        
            # Create YELLOW background rectangle for light periods
            yellow_rect = Rectangle(
                (rect_start, y_min_rect),
                (rect_end - rect_start),
                (y_max_rect - y_min_rect),
                facecolor='lightyellow',
                edgecolor='none',
                linewidth=0,
                alpha=1.0,
                zorder=5
            )
            ax.add_patch(yellow_rect)
    
        # THIRD: Add black rectangles for darkness periods using the full periods
        for channel_str, channel_info in channel_data_dict.items():
            rev_idx = channel_info['rev_idx']
        
            # Get darkness periods for this channel
            dark_periods = darkness_lookup.get(channel_str, [])
        
            for period in dark_periods:
                dark_start = period['start']
                dark_end = period['end']
            
                if dark_end > dark_start:
                    # Calculate y position for this channel
                    y_min_rect = start_y + (rev_idx - 1) * bar_height
                    y_max_rect = start_y + rev_idx * bar_height
                
                    # Create black rectangle spanning the entire darkness period
                    black_rect = Rectangle(
                        (dark_start, y_min_rect),
                        (dark_end - dark_start),
                        (y_max_rect - y_min_rect),
                        facecolor='black',
                        edgecolor='none',
                        linewidth=0,
                        alpha=1.0,
                        zorder=6
                    )
                    ax.add_patch(black_rect)
                    print(f"Added black rectangle for channel {channel_str} from {dark_start} to {dark_end}")
    
        # Set title with proper formatting
        ax.set_title(title, fontsize=14, fontweight='bold')
        ax.set_xlabel('Time (minutes)', fontsize=12)
        ax.set_ylabel(ylabel, fontsize=12)
    
        # Adjust legend
        if num_channels <= 8:
            ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=10)
        else:
            ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=8, ncol=2)
    
        ax.grid(True, alpha=0.3, linestyle='--', zorder=1)
        ax.tick_params(axis='both', which='major', labelsize=10)
    
        # Expand limits to show rectangles
        y_top = start_y + darkness_band_height + (y_range * 0.05)
        ax.set_ylim(y_min - (y_range * 0.05), y_top)
    
        # Expand x limits to show all data and rectangles
        x_min_all = min(all_times) if all_times else 0
        x_max_all = max(all_times) if all_times else 1
    
        x_padding = max(abs(x_max_all - x_min_all) * 0.02, typical_interval * 0.5)
        ax.set_xlim(x_min_all - x_padding, x_max_all + typical_interval + x_padding)
    
        fig.tight_layout()
    
        return fig

    def update_batch_darkness_plots(self, stats_by_channel, combined_df=None):
        """
        Generate darkness plots for batch results with auto-detected darkness periods
        Uses the same channel selection as Metric Plots tab
        """
        if not stats_by_channel:
            return
    
        # Create a separate notebook tab for darkness plots if it doesn't exist
        if not hasattr(self, 'batch_darkness_notebook'):
            self.batch_darkness_notebook = ttk.Notebook(self.batch_results_notebook)
            self.batch_results_notebook.add(self.batch_darkness_notebook, text="Darkness Detection")
    
        # Clear existing tabs
        for tab in self.batch_darkness_notebook.tabs():
            self.batch_darkness_notebook.forget(tab)
    
        # Get combined data for darkness period detection only (not for plotting)
        if combined_df is None and hasattr(self, 'batch_all_processed_data'):
            combined_df = pd.concat(self.batch_all_processed_data, ignore_index=True)
            # Keep original channel names for darkness detection
            combined_df['Channel_original'] = combined_df['Channel']
            combined_df['Channel'] = combined_df['Channel'].astype(str).str.replace(r'\s*\(\d+\)$', '', regex=True)
    
        if combined_df is None or len(combined_df) == 0:
            # Create empty tab with message
            empty_tab = ttk.Frame(self.batch_darkness_notebook)
            self.batch_darkness_notebook.add(empty_tab, text="No Data")
            ttk.Label(empty_tab, text="No data available for darkness detection", 
                      font=("Arial", 12)).pack(pady=50)
            return
    
        # Auto-detect darkness periods using raw data
        darkness_periods = self.detect_darkness_periods_in_batch(combined_df)
    
        if not darkness_periods:
            # Create tab showing no darkness detected
            info_tab = ttk.Frame(self.batch_darkness_notebook)
            self.batch_darkness_notebook.add(info_tab, text="No Darkness")
            ttk.Label(info_tab, text="No darkness periods detected (absorbed_radiation never zero)", 
                      font=("Arial", 12)).pack(pady=50)
            return
    
        # Print debug info
        print(f"\n=== DARKNESS DETECTION ===")
        print(f"Darkness periods detected: {list(darkness_periods.keys())}")
        for ch, periods in darkness_periods.items():
            print(f"  Channel {ch}: {periods}")
    
        # Define the same darkness plots as in the main app
        darkness_plots = [
            ("absorbed_radiation", "Absorbed Radiation with Auto-Detected Darkness", "Radiation"),
            ("CO2_exchange_rate", "CO₂ Exchange Rate - Light/Dark", "CO₂ Exchange Rate"),
            ("Transpiration_H2O_evol_rate", "Transpiration H₂O Evolution Rate - Light/Dark", "Transpiration Rate"),
            ("leaf_temp_c", "Leaf Temperature - Light/Dark", "Leaf Temperature"),
            ("VPD", "Vapor Pressure Deficit (VPD) - Light/Dark", "VPD"),
            ("Stomatal_conductance_corrected", "Stomatal Conductance (Corrected) - Light/Dark", "Conductance"),
            ("overall_conductance", "Overall Conductance - Light/Dark", "Conductance"),
            ("b_layer_conductance", "Boundary Layer Conductance - Light/Dark", "Conductance"),
            ("corr_leaf_conductance", "Corrected Leaf Conductance - Light/Dark", "Conductance"),
            ("corr_stomatal_conductance", "Corrected Stomatal Conductance - Light/Dark", "Conductance"),
            ("stom_resist_to_CO2", "Stomatal Resistance to CO₂ - Light/Dark", "Resistance"),
            ("b_layer_resist_to_CO2", "Boundary Layer Resistance to CO₂ - Light/Dark", "Resistance"),
            ("vapor_pressure_around_leaves", "Vapor Pressure Around Leaves - Light/Dark", "Vapor Pressure"),
            ("saturation_vapor_pressure_inside_leaves", "Saturation Vapor Pressure Inside Leaves - Light/Dark", "Vapor Pressure"),
            ("normalised_corr_stomatal_conductance", "Normalised Corrected Stomatal Conductance - Light/Dark", "Normalised Conductance"),
            ("intercell_CO2_conc_in_gas_phase", "Intercellular CO₂ Concentration - Light/Dark", "CO₂ Concentration"),
            ("mesophyll_conductance_for_CO2", "Mesophyll Conductance for CO₂ - Light/Dark", "Conductance"),
            ("corr_ted_CO2_grad", "Corrected CO₂ Gradient - Light/Dark", "CO₂ Gradient"),
            ("corr_conductance_for_O3_c", "Corrected Conductance for O₃ - Light/Dark", "Conductance"),
            ("corr_stomatal_O3_uptake_rate", "Corrected Stomatal O₃ Uptake Rate - Light/Dark", "O₃ Uptake Rate"),
            ("corr_cumul_O3_dose", "Corrected Cumulative O₃ Dose - Light/Dark", "O₃ Dose"),
            ("corr_stomatal_change_rate", "Corrected Stomatal Change Rate - Light/Dark", "Change Rate"),
            ("relat_stom_cond_change_rate", "Relative Stomatal Conductance Change Rate - Light/Dark", "Change Rate")
        ]
    
        # Filter to only plots that exist in the statistics
        available_plots = []
        # Check which metrics exist in stats_by_channel
        for channel, stats in stats_by_channel.items():
            for time_point in stats.values():
                for metric in darkness_plots:
                    metric_name = metric[0]
                    if metric_name in time_point:
                        if metric_name not in [p[0] for p in available_plots]:
                            available_plots.append(metric)
                break  # Only need to check first time point
            break  # Only need to check first channel
    
        # Get palette
        palette = self.batch_palette_var.get() if hasattr(self, 'batch_palette_var') else 'tab10'
    
        # Create a single scrollable tab with all plots
        all_plots_tab = ttk.Frame(self.batch_darkness_notebook)
        self.batch_darkness_notebook.add(all_plots_tab, text="Darkness Plots")
    
        # Create scrollable container
        container = ttk.Frame(all_plots_tab)
        container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
    
        canvas = tk.Canvas(container, highlightthickness=0)
        v_scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        h_scrollbar = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
    
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
    
        canvas.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.E, tk.W))
    
        # Add plots using statistics data (mean ± SEM)
        row = 0
        for yvar, title, ylab in available_plots:
            # Create figure with error bars using statistics
            fig, ax = plt.subplots(figsize=(12, 6))
        
            # ============ USE THE SAME CHANNEL SELECTION AS METRIC PLOTS ============
            # Filter stats_by_channel to selected channels from batch_channel_vars
            if hasattr(self, 'batch_channel_vars') and self.batch_channel_vars:
                selected_channels = [ch for ch, var in self.batch_channel_vars.items() if var.get()]
                filtered_stats = {ch: stats_by_channel[ch] for ch in selected_channels if ch in stats_by_channel}
            else:
                filtered_stats = stats_by_channel
        
            if not filtered_stats:
                ax.text(0.5, 0.5, 'No channels selected\n\nUse the channel selector in Metric Plots tab', 
                        ha='center', va='center', fontsize=12, fontweight='bold')
                ax.set_axis_off()
            else:
                # Sort channels alphabetically
                sorted_channels = sorted(filtered_stats.keys())
                num_channels = len(sorted_channels)
            
                # Get colors based on palette
                colors = []
                try:
                    if palette in colorblind_friendly_palettes:
                        if isinstance(colorblind_friendly_palettes[palette], list):
                            color_list = colorblind_friendly_palettes[palette]
                            colors = [color_list[i % len(color_list)] for i in range(num_channels)]
                        else:
                            cmap = colorblind_friendly_palettes[palette]
                            colors = cmap(np.linspace(0, 1, num_channels))
                    elif palette in plt.colormaps():
                        cmap = plt.get_cmap(palette)
                        colors = cmap(np.linspace(0, 1, num_channels))
                    else:
                        colors = plt.cm.tab10(np.linspace(0, 1, num_channels))
                except Exception as e:
                    colors = plt.cm.tab10(np.linspace(0, 1, num_channels))
            
                # Plot data for each channel with error bars (mean ± SEM)
                for i, channel in enumerate(sorted_channels):
                    stats = filtered_stats[channel]
                    time_points = sorted(stats.keys())
                    means = []
                    errors = []
                    valid_times = []
                
                    for t in time_points:
                        if yvar in stats[t]:
                            means.append(stats[t][yvar]['mean'])
                            errors.append(stats[t][yvar]['sem'])
                            valid_times.append(t)
                
                    if means:
                        ax.errorbar(valid_times, means, yerr=errors, 
                                   label=channel, marker='o', capsize=3, 
                                   linewidth=2, markersize=5, color=colors[i], alpha=0.7)
            
                # Get unit for y-axis label
                unit = UNITS_MAP.get(yvar, '')
                if unit:
                    ylabel_with_unit = f'{ylab} ({unit})'
                else:
                    ylabel_with_unit = ylab
            
                ax.set_title(title, fontsize=14, fontweight='bold')
                ax.set_xlabel('Time (minutes)', fontsize=12)
                ax.set_ylabel(ylabel_with_unit, fontsize=12)
                ax.legend(bbox_to_anchor=(1.02, 1), loc='upper left', fontsize=9)
                ax.grid(True, alpha=0.3, linestyle='--')
            
                # Add yellow background for light periods and black for darkness
                # Get y-range for rectangle positioning
                all_y_vals = []
                for channel, stats in filtered_stats.items():
                    for t, metrics in stats.items():
                        if yvar in metrics:
                            all_y_vals.append(metrics[yvar]['mean'])
                            all_y_vals.append(metrics[yvar]['mean'] - metrics[yvar]['sem'])
                            all_y_vals.append(metrics[yvar]['mean'] + metrics[yvar]['sem'])
            
                if all_y_vals:
                    y_min = np.nanmin(all_y_vals)
                    y_max = np.nanmax(all_y_vals)
                    y_range = y_max - y_min if y_max != y_min else 1
                else:
                    y_min = 0
                    y_max = 1
                    y_range = 1
            
                # Calculate rectangle positioning
                start_y = y_max + (y_range * 0.05)
                darkness_band_height = y_range * 0.08
                bar_height = darkness_band_height / max(1, num_channels)
            
                # Get time range
                all_times = []
                for channel, stats in filtered_stats.items():
                    all_times.extend(stats.keys())
            
                if all_times:
                    x_min = min(all_times)
                    x_max = max(all_times)
                
                    # Add yellow background for all channels
                    for rev_idx, channel in enumerate(sorted_channels):
                        y_min_rect = start_y + rev_idx * bar_height
                        y_max_rect = start_y + (rev_idx + 1) * bar_height
                    
                        # Yellow background rectangle
                        yellow_rect = Rectangle(
                            (x_min, y_min_rect),
                            (x_max - x_min),
                            (y_max_rect - y_min_rect),
                            facecolor='lightyellow',
                            edgecolor='none',
                            linewidth=0,
                            alpha=1.0,
                            zorder=1
                        )
                        ax.add_patch(yellow_rect)
                
                    # Add black rectangles for darkness periods (only for selected channels)
                    for rev_idx, channel in enumerate(sorted_channels):
                        y_min_rect = start_y + rev_idx * bar_height
                        y_max_rect = start_y + (rev_idx + 1) * bar_height
                    
                        if channel in darkness_periods:
                            for period in darkness_periods[channel]:
                                dark_start = period['start']
                                dark_end = period['end']
                            
                                if dark_end > dark_start:
                                    black_rect = Rectangle(
                                        (dark_start, y_min_rect),
                                        (dark_end - dark_start),
                                        (y_max_rect - y_min_rect),
                                        facecolor='black',
                                        edgecolor='none',
                                        linewidth=0,
                                        alpha=1.0,
                                        zorder=2
                                    )
                                    ax.add_patch(black_rect)
                
                    # Adjust y-limits to show rectangles
                    y_top = start_y + darkness_band_height + (y_range * 0.05)
                    ax.set_ylim(y_min - (y_range * 0.05), y_top)
        
            plt.tight_layout()
            plt.subplots_adjust(right=0.85)
        
            # Create container for this plot
            plot_container = ttk.Frame(scrollable_frame)
            plot_container.grid(row=row, column=0, sticky=(tk.W, tk.E), padx=10, pady=(20, 0))
        
            ttk.Label(plot_container, text=title, font=("Arial", 11, "bold")).grid(
                row=0, column=0, sticky=tk.W)
        
            canvas_widget = FigureCanvasTkAgg(fig, plot_container)
            canvas_widget.draw()
            canvas_widget.get_tk_widget().grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
            btn = ttk.Button(plot_container, text=f"Download {title}",
                            command=lambda f=fig, t=title: self.download_plot(f, t))
            btn.grid(row=2, column=0, sticky=tk.W, pady=(5, 20))
        
            plot_container.columnconfigure(0, weight=1)
            row += 1
    
        # Add a summary tab with darkness period information (only for selected channels)
        summary_tab = ttk.Frame(self.batch_darkness_notebook)
        self.batch_darkness_notebook.add(summary_tab, text="Darkness Summary")
    
        # Create text widget for summary
        summary_text = tk.Text(summary_tab, wrap=tk.WORD, font=("Courier New", 10))
        summary_scrollbar = ttk.Scrollbar(summary_tab, orient="vertical", command=summary_text.yview)
        summary_text.configure(yscrollcommand=summary_scrollbar.set)
    
        summary_text.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        summary_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
    
        # Build summary content with full period information (only for selected channels)
        summary_content = "AUTO-DETECTED DARKNESS PERIODS SUMMARY\n"
        summary_content += "=" * 60 + "\n\n"
    
        # Get selected channels
        if hasattr(self, 'batch_channel_vars') and self.batch_channel_vars:
            selected_channels = [ch for ch, var in self.batch_channel_vars.items() if var.get()]
        else:
            selected_channels = list(darkness_periods.keys())
    
        for channel in selected_channels:
            if channel in darkness_periods:
                periods = darkness_periods[channel]
                summary_content += f"Channel: {channel}\n"
                summary_content += f"  Number of darkness periods: {len(periods)}\n"
                for i, period in enumerate(periods):
                    summary_content += f"    Period {i+1}: {period['start']:.1f} - {period['end']:.1f} min\n"
                summary_content += "\n"
            else:
                summary_content += f"Channel: {channel}\n"
                summary_content += f"  No darkness periods detected\n\n"
    
        summary_content += "=" * 60 + "\n"
        summary_content += "Note: Darkness is detected when absorbed_radiation = 0\n"
        summary_content += "Each period shows the continuous block of darkness (start to end)\n"
        summary_content += "Plots show mean ± SEM across replicates\n"
    
        summary_text.insert(tk.END, summary_content)
        summary_text.configure(state='disabled')
    
        summary_tab.columnconfigure(0, weight=1)
        summary_tab.rowconfigure(0, weight=1)
    
        # Configure weights
        all_plots_tab.columnconfigure(0, weight=1)
        all_plots_tab.rowconfigure(0, weight=1)
    
        print(f"Darkness detection complete. Found darkness in {len(darkness_periods)} channels")
    
    def _create_fallback_distribution_plot(self, frame, row, filtered_stats, get_palette_colors, palette):
        """Create a fallback distribution plot using statistics when raw data is not available"""
        try:
            # Try to reconstruct data for distribution plot from stats
            dist_data = []
            dist_labels = []
            sorted_channels = sorted(filtered_stats.keys())
            for channel in sorted_channels:
                channel_stats = filtered_stats[channel]
                # Collect all means across time points
                means_list = []
                for time_point in channel_stats.values():
                    if 'WUEi' in time_point:
                        means_list.append(time_point['WUEi']['mean'])
                if means_list:
                    dist_data.append(means_list)
                    dist_labels.append(channel)
        
            if dist_data:
                fig_dist, ax_dist = plt.subplots(figsize=(12, 8))
            
                # Create boxplot from the means - use tick_labels parameter
                box_plot = ax_dist.boxplot(dist_data, tick_labels=dist_labels, patch_artist=True)
            
                # Get colors
                dist_colors = get_palette_colors(len(dist_data))
                for patch, color in zip(box_plot['boxes'], dist_colors):
                    patch.set_facecolor(color)
                    patch.set_alpha(0.7)
            
                ax_dist.set_title('WUEi Distribution by Channel (Mean values over time)', 
                                fontsize=14, fontweight='bold')
                ax_dist.set_xlabel('Channel', fontsize=12)
                ax_dist.set_ylabel('WUEi (μmol CO₂ / mol H₂O)', fontsize=12)
                ax_dist.grid(True, alpha=0.3)
                plt.xticks(rotation=45, ha='right')
                plt.tight_layout()
            
                # Embed plot
                plot_container = ttk.Frame(frame)
                plot_container.grid(row=row, column=0, sticky=(tk.W, tk.E), 
                                   padx=10, pady=10, columnspan=2)
            
                ttk.Label(plot_container, text="WUEi Distribution by Channel (Mean values)", 
                         font=("Arial", 11, "bold")).grid(row=0, column=0, sticky=tk.W)
            
                canvas = FigureCanvasTkAgg(fig_dist, plot_container)
                canvas.draw()
                canvas.get_tk_widget().grid(row=1, column=0, sticky=(tk.W, tk.E))
            
                btn = ttk.Button(plot_container, text="Download WUEi Distribution",
                                command=lambda f=fig_dist, t="WUEi_Distribution": self.download_plot(f, t))
                btn.grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
                return True
            else:
                return False
        except Exception as e:
            print(f"Error creating fallback distribution plot: {e}")
            return False

    def create_full_distribution_plot(self, df, metric, title, ylabel, palette='tab10'):
        """Create a full distribution plot with violin, boxplot, and scatter points"""
    
        # Close any existing figure
        plt.close('all')
    
        if df is None or len(df) == 0 or metric not in df.columns:
            fig, ax = plt.subplots(figsize=(12, 6))
            ax.text(0.5, 0.5, 'No data available', ha='center', va='center', fontsize=12)
            ax.set_axis_off()
            return fig
    
        fig, ax = plt.subplots(figsize=(12, 8))
    
        # Prepare data for each channel
        channels = df['Channel'].unique()
        metric_data = []
        channel_labels = []
    
        for channel in sorted(channels):
            channel_data = df[df['Channel'] == channel]
            values = channel_data[metric].dropna().values
            if len(values) > 0:
                metric_data.append(values)
                channel_labels.append(str(channel))
    
        if not metric_data:
            ax.text(0.5, 0.5, f'No {metric} data available', ha='center', va='center', fontsize=12)
            ax.set_axis_off()
            return fig
    
        # Get colors
        colors = plt.cm.tab10(np.linspace(0, 1, len(metric_data)))
        positions = range(1, len(metric_data) + 1)
    
        # --- LAYER 1: VIOLIN PLOT (background) ---
        violin_parts = ax.violinplot(metric_data, positions=positions, 
                                     showmeans=False, showmedians=False, showextrema=False)
    
        for i, body in enumerate(violin_parts['bodies']):
            body.set_facecolor(colors[i])
            body.set_alpha(0.3)
            body.set_edgecolor(colors[i])
            body.set_linewidth(1)
            body.set_zorder(1)
    
        # --- LAYER 2: BOXPLOT (foreground) ---
        # Use tick_labels instead of labels (fixes deprecation warning)
        box_plot = ax.boxplot(metric_data, positions=positions, patch_artist=True, tick_labels=channel_labels)
    
        for i, (box, color) in enumerate(zip(box_plot['boxes'], colors)):
            box.set_facecolor(color)
            box.set_alpha(0.7)
            box.set_edgecolor('black')
            box.set_linewidth(1.5)
            box.set_zorder(3)
    
        for median in box_plot['medians']:
            median.set_color('red')
            median.set_linewidth(2.5)
            median.set_zorder(4)
    
        for whisker in box_plot['whiskers']:
            whisker.set(color='black', linewidth=1.5, zorder=2)
        for cap in box_plot['caps']:
            cap.set(color='black', linewidth=1.5, zorder=2)
    
        # --- LAYER 3: DATA POINTS (top layer) ---
        for i, data in enumerate(metric_data):
            # Add jitter to x-position for better visibility
            x_jitter = np.random.normal(positions[i], 0.08, size=len(data))
            ax.scatter(x_jitter, data, alpha=0.5, color=colors[i], s=20, 
                      edgecolor='black', linewidth=0.5, zorder=5)
    
        # --- LAYER 4: MEAN MARKERS ---
        for i, data in enumerate(metric_data):
            mean_val = np.mean(data)
            ax.scatter(positions[i], mean_val, color='yellow', s=100, 
                      edgecolor='black', linewidth=1.5, zorder=6, marker='D')
    
        # Configure plot
        ax.set_title(title, fontsize=14, fontweight='bold')
        ax.set_xlabel('Channel', fontsize=12)
        ax.set_ylabel(ylabel, fontsize=12)
        ax.set_xticks(positions)
        ax.set_xticklabels(channel_labels, rotation=45, ha='right', fontsize=10)
        ax.grid(True, alpha=0.2, linestyle='--', axis='y', zorder=0)
    
        # Add statistics annotations
        for i, data in enumerate(metric_data):
            median = np.median(data)
            q1 = np.percentile(data, 25)
            q3 = np.percentile(data, 75)
        
            ax.text(positions[i], median, f'{median:.1f}', 
                    ha='center', va='center', fontsize=9, fontweight='bold',
                    bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.8, edgecolor='none'),
                    zorder=7)
        
            ax.text(positions[i], q1, f'Q1', ha='center', va='top', fontsize=8, alpha=0.7)
            ax.text(positions[i], q3, f'Q3', ha='center', va='bottom', fontsize=8, alpha=0.7)
    
        # Legend
        legend_elements = [
            Patch(facecolor='gray', alpha=0.3, edgecolor='gray', label='Violin (distribution shape)'),
            Rectangle((0, 0), 1, 1, facecolor='gray', alpha=0.7, edgecolor='black', label='Box (quartiles)'),
            Line2D([0], [0], color='red', linewidth=2.5, label='Median'),
            Line2D([0], [0], marker='o', color='w', markerfacecolor='gray', markersize=8, label='Data points'),
            Line2D([0], [0], marker='D', color='w', markerfacecolor='yellow', markersize=8, label='Mean'),
        ]
    
        ax.legend(handles=legend_elements, loc='upper left', 
                  bbox_to_anchor=(1.02, 1), fontsize=9, framealpha=0.9)
    
        plt.tight_layout()
        plt.subplots_adjust(right=0.85)
    
        return fig


    
    def update_batch_distribution_plots(self, stats_by_channel):
        """Update the distribution plots tab with violin + boxplot for WUEi (grouped in Efficiency tab)"""
        if not stats_by_channel:
            return
    
        # Clear existing widgets
        for widget in self.batch_distribution_frame.winfo_children():
            widget.destroy()
    
        # Get the original data from batch results (if available)
        # For distribution plots, we need the raw data, not just statistics
        if not hasattr(self, 'all_processed_data') or not self.all_processed_data:
            # Try to load from saved data
            output_dir = self.output_dir_var.get()
            data_path = os.path.join(output_dir, "combined_data.csv")
            if os.path.exists(data_path):
                try:
                    combined_df = pd.read_csv(data_path)
                except:
                    combined_df = None
            else:
                combined_df = None
        else:
            combined_df = pd.concat(self.all_processed_data, ignore_index=True)
    
        if combined_df is None or len(combined_df) == 0:
            fig, ax = plt.subplots(figsize=(12, 6))
            ax.text(0.5, 0.5, 'No data available for distribution plots', 
                    ha='center', va='center', fontsize=12)
            ax.set_axis_off()
            canvas = FigureCanvasTkAgg(fig, self.batch_distribution_frame)
            canvas.draw()
            canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.W, tk.E))
            return
    
        # Apply base name extraction to Channel column
        combined_df['Channel'] = combined_df['Channel'].astype(str).str.replace(r'\s*\(\d+\)$', '', regex=True)
    
        # Filter by selected channels
        if self.batch_selected_channels:
            combined_df = combined_df[combined_df['Channel'].isin(self.batch_selected_channels)]
    
        if len(combined_df) == 0:
            fig, ax = plt.subplots(figsize=(12, 6))
            ax.text(0.5, 0.5, 'No channels selected\n\nUse the channel selector to choose channels', 
                    ha='center', va='center', fontsize=12, fontweight='bold')
            ax.set_axis_off()
            canvas = FigureCanvasTkAgg(fig, self.batch_distribution_frame)
            canvas.draw()
            canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.W, tk.E))
            return
    
        # Get palette
        palette = self.palette_var.get() if hasattr(self, 'palette_var') else 'tab10'
    
        # Create WUEi distribution plot (main distribution plot)
        if 'WUEi' in combined_df.columns:
            fig = self.create_batch_distribution_plot(combined_df, 'WUEi', 
                                                       'Intrinsic Water Use Efficiency (WUEi)', 
                                                       'WUEi (μmol CO₂ / mol H₂O)', palette)
            fig.set_size_inches(12, 8)
        
            plot_container = ttk.Frame(self.batch_distribution_frame)
            plot_container.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=(20, 0), columnspan=2)
        
            ttk.Label(plot_container, text="WUEi Distribution by Channel", 
                      font=("Arial", 11, "bold")).grid(row=0, column=0, sticky=tk.W)
        
            canvas = FigureCanvasTkAgg(fig, plot_container)
            canvas.draw()
            canvas.get_tk_widget().grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
            btn = ttk.Button(plot_container, text="Download WUEi Distribution",
                            command=lambda f=fig, t="WUEi_Distribution": self.download_plot(f, t))
            btn.grid(row=2, column=0, sticky=tk.W, pady=(5, 20))
        
            plot_container.columnconfigure(0, weight=1)
    
        self.batch_distribution_frame.columnconfigure(0, weight=1)

    
    def update_batch_results_tabs(self, stats_by_channel):
        """Update batch results tabs with statistics"""
        if not stats_by_channel:
            # Show message in tabs
            if hasattr(self, 'batch_results_notebook'):
                for tab_id in self.batch_results_notebook.tabs():
                    tab = self.batch_results_notebook.nametowidget(tab_id)
                    for widget in tab.winfo_children():
                        widget.destroy()
                    fig, ax = plt.subplots(figsize=(12, 6))
                    ax.text(0.5, 0.5, 'No data available - batch processing may have failed', 
                           ha='center', va='center', fontsize=14)
                    ax.set_axis_off()
                    canvas = FigureCanvasTkAgg(fig, tab)
                    canvas.draw()
                    canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.W, tk.E))
            return

        # Print channel names for debugging
        print(f"\nUpdating batch results tabs with channels: {list(stats_by_channel.keys())}")

        # Update statistics table
        self.update_statistics_table(stats_by_channel)

        # Populate channel selector
        self.populate_batch_channel_selector(stats_by_channel)

        # Update metric category plots (now includes WUEi distribution in Efficiency tab)
        self.update_batch_metric_category_plots(stats_by_channel)
    
        # ============ FIX: Force update WUEi distribution tab ============
        self.update_wuei_distribution_tab()

    def update_distribution_plot_tab(self):
        """Update the distribution plot tab (violin + boxplot)"""
        output_dir = self.output_dir_var.get()
        plot_path = os.path.join(output_dir, "WUEi_Distribution.png")
    
        if not os.path.exists(plot_path):
            return
    
        # Check if we already have a distribution tab
        if not hasattr(self, 'distribution_tab'):
            # Create the tab - FIX: use batch_results_notebook instead of batch_notebook
            self.distribution_tab = ttk.Frame(self.batch_results_notebook)
            self.batch_results_notebook.add(self.distribution_tab, text="WUEi Distribution")
        
            # Create scrollable frame
            canvas = tk.Canvas(self.distribution_tab)
            v_scrollbar = ttk.Scrollbar(self.distribution_tab, orient="vertical", command=canvas.yview)
            h_scrollbar = ttk.Scrollbar(self.distribution_tab, orient="horizontal", command=canvas.xview)
        
            scrollable_frame = ttk.Frame(canvas)
            scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
            canvas.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
            v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
            h_scrollbar.grid(row=1, column=0, sticky=(tk.E, tk.W))
        
            self.distribution_tab_frame = scrollable_frame
            self.distribution_tab.columnconfigure(0, weight=1)
            self.distribution_tab.rowconfigure(0, weight=1)
        else:
            # Clear existing widgets
            for widget in self.distribution_tab_frame.winfo_children():
                widget.destroy()

    
        try:
            img = Image.open(plot_path)
            # Scale to fit (max width 1000, maintain aspect ratio)
            max_width = 1000
            if img.width > max_width:
                ratio = max_width / img.width
                new_size = (max_width, int(img.height * ratio))
                img = img.resize(new_size, Image.Resampling.LANCZOS)
        
            photo = ImageTk.PhotoImage(img)
        
            # Create label for the image
            img_label = ttk.Label(self.distribution_tab_frame, image=photo)
            img_label.image = photo  # Keep reference
            img_label.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=10)
        
            # Add download button
            btn = ttk.Button(self.distribution_tab_frame, text="Download Distribution Plot",
                            command=lambda: self.download_file(plot_path))
            btn.grid(row=1, column=0, pady=5)
        
        except Exception as e:
            error_label = ttk.Label(self.distribution_tab_frame, 
                                    text=f"Could not load plot: {str(e)}")
            error_label.grid(row=0, column=0, pady=20)
    
    def update_wuei_distribution_tab(self):
        """Update the WUEi distribution plot tab dynamically when channels are selected/deselected"""

        # Clear existing widgets
        for widget in self.wuei_dist_scrollable_frame.winfo_children():
            widget.destroy()

        # Check if we have batch data available
        if not hasattr(self, 'batch_all_processed_data') or not self.batch_all_processed_data:
            ttk.Label(self.wuei_dist_scrollable_frame, 
                     text="No batch data available.\n\nRun batch processing first to generate distribution plots.",
                     font=("Arial", 12), foreground="gray").pack(pady=50)
            return

        # Get selected channels - FIX: If no channels selected in batch_channel_vars but channels exist, select all
        if hasattr(self, 'batch_channel_vars') and self.batch_channel_vars:
            selected_channels = [ch for ch, var in self.batch_channel_vars.items() if var.get()]
        else:
            selected_channels = None

        # FIX: If selected_channels is empty but we have stats, that means no channels are selected
        # But if we have batch_all_processed_data, we should show ALL channels by default
        if (not selected_channels or len(selected_channels) == 0) and hasattr(self, 'batch_stats') and self.batch_stats:
            # No channels selected, but we have stats - select all channels
            selected_channels = list(self.batch_stats.keys())
            print(f"DEBUG - No channels selected, defaulting to all: {selected_channels}")
        
            # Also update the checkboxes to reflect this selection
            if hasattr(self, 'batch_channel_vars'):
                for ch in self.batch_channel_vars:
                    self.batch_channel_vars[ch].set(True)
                self.batch_selected_channels = selected_channels.copy()

        print(f"DEBUG - Selected channels: {selected_channels}")

        # Combine all processed data
        try:
            combined_df = pd.concat(self.batch_all_processed_data, ignore_index=True)

            print(f"DEBUG - Original channels in data: {combined_df['Channel'].unique().tolist()}")

            # CRITICAL: Extract base channel names (remove bracketed numbers like (2), (3), etc.)
            # This must match exactly what was done in generate_batch_statistics
            combined_df['Channel_base'] = combined_df['Channel'].astype(str).str.replace(r'\s*\(\d+\)$', '', regex=True)

            print(f"DEBUG - Base channels in data: {combined_df['Channel_base'].unique().tolist()}")

            # Filter by selected channels (using base names)
            # IMPORTANT FIX: Only filter if selected_channels is not None AND has items
            if selected_channels and len(selected_channels) > 0:
                combined_df = combined_df[combined_df['Channel_base'].isin(selected_channels)]

            print(f"DEBUG - After filtering, rows: {len(combined_df)}")

            # Get unique base channels after filtering
            unique_channels = combined_df['Channel_base'].nunique() if len(combined_df) > 0 else 0

            # Check if we have data after filtering
            if len(combined_df) == 0 or unique_channels == 0:
                # Show appropriate message based on selection state
                if selected_channels and len(selected_channels) > 0:
                    message = "No data available for selected channels.\n\nSelect different channels or run batch processing."
                else:
                    message = "No channels selected.\n\nUse the channel selector above to choose channels."
            
                ttk.Label(self.wuei_dist_scrollable_frame, 
                         text=message,
                         font=("Arial", 12), foreground="gray").pack(pady=50)
                return

            # Check if WUEi column exists
            if 'WUEi' not in combined_df.columns:
                ttk.Label(self.wuei_dist_scrollable_frame, 
                         text="WUEi column not found in data.\n\nMake sure your data contains WUEi values.",
                         font=("Arial", 12), foreground="gray").pack(pady=50)
                return

            # Get palette
            if hasattr(self, 'batch_palette_var'):
                palette = self.batch_palette_var.get()
            else:
                palette = 'tab10'

            # Create the distribution plot - use Channel_base for grouping
            fig = self.create_full_distribution_plot_with_base_names(
                combined_df, 
                'WUEi', 
                'Intrinsic Water Use Efficiency (WUEi) Distribution', 
                'WUEi (μmol CO₂ / mol H₂O)', 
                palette
            )
            fig.set_size_inches(12, 8)

            # Create a frame for the plot
            img_frame = ttk.Frame(self.wuei_dist_scrollable_frame)
            img_frame.pack(expand=True, fill=tk.BOTH, pady=20)

            # Embed the plot
            canvas = FigureCanvasTkAgg(fig, img_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(expand=True, fill=tk.BOTH)

            # Add info text
            total_channels = unique_channels
            if selected_channels and len(selected_channels) > 0 and hasattr(self, 'batch_channel_vars'):
                info_text = f"Showing {total_channels} channel(s)"
                if len(selected_channels) < len(self.batch_channel_vars):
                    info_text += f" (filtered from {len(self.batch_channel_vars)} total)"
            else:
                info_text = f"Showing {total_channels} channel(s)"

            # Add download button
            btn_frame = ttk.Frame(self.wuei_dist_scrollable_frame)
            btn_frame.pack(pady=10)

            btn = ttk.Button(btn_frame, 
                           text="Download WUEi Distribution Plot",
                           command=lambda f=fig: self.download_plot(f, "WUEi_Distribution"))
            btn.pack()

            # Add info label
            info_label = ttk.Label(self.wuei_dist_scrollable_frame, 
                                 text=info_text + "\nThis plot shows the distribution of WUEi values across all time points for each channel.\n"
                                      "Violin plot (background) shows the probability density, boxplot (foreground) shows quartiles.\n"
                                      "Red line = median, yellow diamond = mean, dots = individual data points.",
                                 font=("Arial", 9), foreground="gray", justify=tk.CENTER)
            info_label.pack(pady=10)

            # Store reference to the figure
            self.current_wuei_fig = fig

            print(f"DEBUG - WUEi distribution plot created successfully with {unique_channels} channels")

        except Exception as e:
            print(f"Error creating WUEi distribution plot: {e}")
            traceback.print_exc()
            ttk.Label(self.wuei_dist_scrollable_frame, 
                     text=f"Error creating plot: {str(e)}",
                     font=("Arial", 12), foreground="red").pack(pady=50)


    def create_full_distribution_plot_with_base_names(self, df, metric, title, ylabel, palette='tab10'):
        """Create a full distribution plot with violin, boxplot, and scatter points using base channel names"""
    
        # Close any existing figure
        plt.close('all')
    
        if df is None or len(df) == 0 or metric not in df.columns:
            fig, ax = plt.subplots(figsize=(12, 6))
            ax.text(0.5, 0.5, 'No data available', ha='center', va='center', fontsize=12)
            ax.set_axis_off()
            return fig
    
        fig, ax = plt.subplots(figsize=(12, 8))
    
        # Use Channel_base for grouping (this has bracketed numbers removed)
        # This ensures consistency with the channel selector
        if 'Channel_base' in df.columns:
            channel_col = 'Channel_base'
        else:
            channel_col = 'Channel'
    
        # Prepare data for each channel
        channels = df[channel_col].unique()
        metric_data = []
        channel_labels = []
    
        for channel in sorted(channels):
            channel_data = df[df[channel_col] == channel]
            values = channel_data[metric].dropna().values
            if len(values) > 0:
                metric_data.append(values)
                channel_labels.append(str(channel))
    
        if not metric_data:
            ax.text(0.5, 0.5, f'No {metric} data available', ha='center', va='center', fontsize=12)
            ax.set_axis_off()
            return fig
    
        # Get colors
        try:
            colors = plt.cm.tab10(np.linspace(0, 1, len(metric_data)))
        except:
            colors = plt.cm.Set2(np.linspace(0, 1, len(metric_data)))
    
        positions = range(1, len(metric_data) + 1)
    
        # --- LAYER 1: VIOLIN PLOT (background) ---
        try:
            violin_parts = ax.violinplot(metric_data, positions=positions, 
                                         showmeans=False, showmedians=False, showextrema=False)
        
            for i, body in enumerate(violin_parts['bodies']):
                body.set_facecolor(colors[i])
                body.set_alpha(0.3)
                body.set_edgecolor(colors[i])
                body.set_linewidth(1)
                body.set_zorder(1)
        except Exception as e:
            print(f"Warning: Violin plot failed: {e}")
            # Continue without violin plot
    
        # --- LAYER 2: BOXPLOT (foreground) ---
        try:
            # Use tick_labels parameter (for newer matplotlib)
            box_plot = ax.boxplot(metric_data, positions=positions, patch_artist=True, tick_labels=channel_labels)
        except TypeError:
            # Fallback for older matplotlib
            box_plot = ax.boxplot(metric_data, positions=positions, patch_artist=True)
            ax.set_xticklabels(channel_labels)
    
        for i, (box, color) in enumerate(zip(box_plot['boxes'], colors)):
            box.set_facecolor(color)
            box.set_alpha(0.7)
            box.set_edgecolor('black')
            box.set_linewidth(1.5)
            box.set_zorder(3)
    
        for median in box_plot['medians']:
            median.set_color('red')
            median.set_linewidth(2.5)
            median.set_zorder(4)
    
        for whisker in box_plot['whiskers']:
            whisker.set(color='black', linewidth=1.5, zorder=2)
        for cap in box_plot['caps']:
            cap.set(color='black', linewidth=1.5, zorder=2)
    
        # --- LAYER 3: DATA POINTS (top layer) ---
        for i, data in enumerate(metric_data):
            if len(data) > 0:
                x_jitter = np.random.normal(positions[i], 0.08, size=len(data))
                ax.scatter(x_jitter, data, alpha=0.5, color=colors[i], s=20, 
                          edgecolor='black', linewidth=0.5, zorder=5)
    
        # --- LAYER 4: MEAN MARKERS ---
        for i, data in enumerate(metric_data):
            if len(data) > 0:
                mean_val = np.mean(data)
                ax.scatter(positions[i], mean_val, color='yellow', s=100, 
                          edgecolor='black', linewidth=1.5, zorder=6, marker='D')
    
        # Configure plot
        ax.set_title(title, fontsize=14, fontweight='bold')
        ax.set_xlabel('Channel', fontsize=12)
        ax.set_ylabel(ylabel, fontsize=12)
        ax.set_xticks(positions)
        ax.set_xticklabels(channel_labels, rotation=45, ha='right', fontsize=10)
        ax.grid(True, alpha=0.2, linestyle='--', axis='y', zorder=0)
    
        # Add statistics annotations
        for i, data in enumerate(metric_data):
            if len(data) > 0:
                median = np.median(data)
                q1 = np.percentile(data, 25)
                q3 = np.percentile(data, 75)
            
                ax.text(positions[i], median, f'{median:.1f}', 
                        ha='center', va='center', fontsize=9, fontweight='bold',
                        bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.8, edgecolor='none'),
                        zorder=7)
            
                # Only show Q1/Q3 labels if there's enough space
                ylim = ax.get_ylim()
                if q1 > ylim[0] + (ylim[1] - ylim[0]) * 0.1:
                    ax.text(positions[i], q1, f'Q1', ha='center', va='top', fontsize=8, alpha=0.7)
                    ax.text(positions[i], q3, f'Q3', ha='center', va='bottom', fontsize=8, alpha=0.7)
    
        # Legend
        legend_elements = [
            Patch(facecolor='gray', alpha=0.3, edgecolor='gray', label='Violin (distribution shape)'),
            Rectangle((0, 0), 1, 1, facecolor='gray', alpha=0.7, edgecolor='black', label='Box (quartiles)'),
            Line2D([0], [0], color='red', linewidth=2.5, label='Median'),
            Line2D([0], [0], marker='o', color='w', markerfacecolor='gray', markersize=8, label='Data points'),
            Line2D([0], [0], marker='D', color='w', markerfacecolor='yellow', markersize=8, label='Mean'),
        ]
    
        ax.legend(handles=legend_elements, loc='upper left', 
                  bbox_to_anchor=(1.02, 1), fontsize=9, framealpha=0.9)
    
        plt.tight_layout()
        plt.subplots_adjust(right=0.85)
    
        return fig
    
    def create_filtered_wuei_plot(self, selected_channels, output_dir):
        """Create a filtered WUEi distribution plot for selected channels only"""
        if not hasattr(self, 'batch_all_processed_data') or not self.batch_all_processed_data:
            print("No batch data available for filtered plot")
            return
    
        try:
            # Combine all processed data
            combined_df = pd.concat(self.batch_all_processed_data, ignore_index=True)
        
            # Apply base name extraction to Channel column
            combined_df['Channel'] = combined_df['Channel'].astype(str).str.replace(r'\s*\(\d+\)$', '', regex=True)
        
            # Filter by selected channels
            combined_df = combined_df[combined_df['Channel'].isin(selected_channels)]
        
            if len(combined_df) == 0:
                print("No data available for selected channels")
                return
        
            # Create the distribution plot
            fig, ax = plt.subplots(figsize=(12, 8))
        
            # Prepare data for each selected channel
            channels = combined_df['Channel'].unique()
            wuei_data = []
            channel_labels = []
        
            for channel in sorted(channels):
                channel_data = combined_df[combined_df['Channel'] == channel]
                wuei_values = channel_data['WUEi'].dropna().values
                if len(wuei_values) > 0:
                    wuei_data.append(wuei_values)
                    channel_labels.append(channel)
        
            if not wuei_data:
                return
        
            # Get colors for channels
            colors = plt.cm.tab10(np.linspace(0, 1, len(wuei_data)))
            positions = range(1, len(wuei_data) + 1)
        
            # --- LAYER 1: VIOLIN PLOT (background) ---
            violin_parts = ax.violinplot(wuei_data, positions=positions, 
                                         showmeans=False, showmedians=False, showextrema=False)
        
            for i, body in enumerate(violin_parts['bodies']):
                body.set_facecolor(colors[i])
                body.set_alpha(0.3)
                body.set_edgecolor(colors[i])
                body.set_linewidth(1)
                body.set_zorder(1)
        
            # --- LAYER 2: BOXPLOT (foreground) ---
            # Use tick_labels parameter (for newer matplotlib) or labels for older
            try:
                box_plot = ax.boxplot(wuei_data, positions=positions, patch_artist=True, tick_labels=channel_labels)
            except TypeError:
                # Fallback for older matplotlib
                box_plot = ax.boxplot(wuei_data, positions=positions, patch_artist=True)
                ax.set_xticklabels(channel_labels)
        
            for i, (box, color) in enumerate(zip(box_plot['boxes'], colors)):
                box.set_facecolor(color)
                box.set_alpha(0.7)
                box.set_edgecolor('black')
                box.set_linewidth(1.5)
                box.set_zorder(3)
        
            for median in box_plot['medians']:
                median.set_color('red')
                median.set_linewidth(2.5)
                median.set_zorder(4)
        
            for whisker in box_plot['whiskers']:
                whisker.set(color='black', linewidth=1.5, zorder=2)
            for cap in box_plot['caps']:
                cap.set(color='black', linewidth=1.5, zorder=2)
        
            # --- LAYER 3: DATA POINTS (top layer) ---
            for i, data in enumerate(wuei_data):
                x_jitter = np.random.normal(positions[i], 0.08, size=len(data))
                ax.scatter(x_jitter, data, alpha=0.5, color=colors[i], s=20, 
                          edgecolor='black', linewidth=0.5, zorder=5)
        
            # --- LAYER 4: MEAN MARKERS ---
            for i, data in enumerate(wuei_data):
                mean_val = np.mean(data)
                ax.scatter(positions[i], mean_val, color='yellow', s=100, 
                          edgecolor='black', linewidth=1.5, zorder=6, marker='D')
        
            # Configure plot
            if len(selected_channels) == 1:
                title = f'WUEi Distribution - Channel: {selected_channels[0]}'
            else:
                title = f'WUEi Distribution by Channel ({len(selected_channels)} channels selected)'
        
            ax.set_title(title, fontsize=14, fontweight='bold')
            ax.set_xlabel('Channel', fontsize=12)
            ax.set_ylabel('WUEi (μmol CO₂ / mol H₂O)', fontsize=12)
            ax.set_xticks(positions)
            ax.set_xticklabels(channel_labels, rotation=45, ha='right', fontsize=10)
            ax.grid(True, alpha=0.2, linestyle='--', axis='y', zorder=0)
        
            # Add statistics annotations
            for i, data in enumerate(wuei_data):
                median = np.median(data)
                q1 = np.percentile(data, 25)
                q3 = np.percentile(data, 75)
            
                ax.text(positions[i], median, f'{median:.1f}', 
                        ha='center', va='center', fontsize=9, fontweight='bold',
                        bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.8, edgecolor='none'),
                        zorder=7)
        
            # Legend
            legend_elements = [
                Patch(facecolor='gray', alpha=0.3, edgecolor='gray', label='Violin (distribution shape)'),
                Rectangle((0, 0), 1, 1, facecolor='gray', alpha=0.7, edgecolor='black', label='Box (quartiles)'),
                Line2D([0], [0], color='red', linewidth=2.5, label='Median'),
                Line2D([0], [0], marker='o', color='w', markerfacecolor='gray', markersize=8, label='Data points'),
                Line2D([0], [0], marker='D', color='w', markerfacecolor='yellow', markersize=8, label='Mean'),
            ]
        
            ax.legend(handles=legend_elements, loc='upper left', 
                      bbox_to_anchor=(1.02, 1), fontsize=9, framealpha=0.9)
        
            plt.tight_layout()
            plt.subplots_adjust(right=0.85)
        
            # Save the filtered plot
            output_path = os.path.join(output_dir, "WUEi_Distribution_filtered.png")
            plt.savefig(output_path, dpi=300, bbox_inches='tight')
            plt.close(fig)
        
            print(f"Filtered WUEi distribution plot saved to {output_path}")
        
        except Exception as e:
            print(f"Error creating filtered WUEi plot: {e}")
            traceback.print_exc()
    
    def download_file(self, filepath):
        """Download a file to user-selected location"""
        if not os.path.exists(filepath):
            messagebox.showerror("Error", f"File not found: {filepath}")
            return
    
        save_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            initialfile=os.path.basename(filepath),
            filetypes=[("PNG files", "*.png"), ("All files", "*.*")]
        )
    
        if save_path:
            try:
                shutil.copy2(filepath, save_path)
                messagebox.showinfo("Success", f"File saved to {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {str(e)}")
    
    def test_batch_channels(self):
        """Test function to verify channel names in batch results"""
        if not hasattr(self, 'batch_results') or not self.batch_results:
            print("No batch results available")
            return
    
        print("\n" + "="*60)
        print("TEST BATCH CHANNELS")
        print("="*60)
    
        for i, result in enumerate(self.batch_results):
            print(f"\nFile {i+1}: {result.get('filepath', 'Unknown')}")
            print(f"  Replicate: {result.get('replicate', 'Unknown')}")
        
            df = result.get('calculated_metrics')
            if df is not None:
                channels = df['Channel'].unique()
                print(f"  Channels ({len(channels)}): {channels.tolist()}")
            
                # Check data types
                print(f"  Channel type: {df['Channel'].dtype}")
            
                # Check first few rows
                print("  First 5 rows of Channel column:")
                for idx, val in enumerate(df['Channel'].head(5)):
                    print(f"    Row {idx}: '{val}'")
            else:
                print("  No calculated metrics")
    
        print("="*60 + "\n")
    
    def update_statistics_table(self, stats_by_channel):
        """Update the statistics table with batch results using renamed channels"""
        # Clear existing
        for item in self.stats_tree.get_children():
            self.stats_tree.delete(item)
    
        if not stats_by_channel:
            return
    
        # Build columns
        first_channel = list(stats_by_channel.keys())[0]
        first_time = list(stats_by_channel[first_channel].keys())[0]
        metrics = list(stats_by_channel[first_channel][first_time].keys())
    
        columns = ['Channel', 'time_repeated'] + [f"{m}_mean" for m in metrics] + [f"{m}_std" for m in metrics]
        self.stats_tree['columns'] = columns
        self.stats_tree['show'] = 'headings'
    
        # Configure columns with appropriate widths
        for col in columns:
            self.stats_tree.heading(col, text=col)
            if col == 'Channel':
                self.stats_tree.column(col, width=150, anchor=tk.CENTER)
            elif col == 'time_repeated':
                self.stats_tree.column(col, width=100, anchor=tk.CENTER)
            else:
                self.stats_tree.column(col, width=120, anchor=tk.CENTER)
    
        # Insert data with renamed channel names
        for channel in sorted(stats_by_channel.keys()):
            stats = stats_by_channel[channel]
            for time_point in sorted(stats.keys()):
                row = [channel, f"{time_point:.1f}"]
                for metric in metrics:
                    if metric in stats[time_point]:
                        row.append(f"{stats[time_point][metric]['mean']:.3f}")
                    else:
                        row.append("")
                for metric in metrics:
                    if metric in stats[time_point]:
                        row.append(f"{stats[time_point][metric]['std']:.3f}")
                    else:
                        row.append("")
                self.stats_tree.insert('', 'end', values=row)

    def update_batch_plots_tab(self, stats_by_channel):
        """Update the batch plots tab with averaged plots - using selected channels only"""
        print(f"\n=== UPDATE BATCH PLOTS TAB ===")
        print(f"stats_by_channel keys: {list(stats_by_channel.keys()) if stats_by_channel else 'None'}")
        print(f"Selected channels: {self.batch_selected_channels}")

        # Clear existing
        for widget in self.batch_plots_frame.winfo_children():
            widget.destroy()

        if not stats_by_channel:
            fig, ax = plt.subplots(figsize=(12, 6))
            ax.text(0.5, 0.5, 'No data available', ha='center', va='center', fontsize=14)
            ax.set_axis_off()
            canvas = FigureCanvasTkAgg(fig, self.batch_plots_frame)
            canvas.draw()
            canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.W, tk.E))
            return

        # Filter stats_by_channel to only include selected channels
        filtered_stats = {}
        for channel in self.batch_selected_channels:
            if channel in stats_by_channel:
                filtered_stats[channel] = stats_by_channel[channel]
    
        # If no channels selected, show message
        if not filtered_stats:
            fig, ax = plt.subplots(figsize=(12, 6))
            ax.text(0.5, 0.5, 'No channels selected\n\nUse the channel selector above to choose channels', 
                    ha='center', va='center', fontsize=12, fontweight='bold')
            ax.set_axis_off()
            canvas = FigureCanvasTkAgg(fig, self.batch_plots_frame)
            canvas.draw()
            canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.W, tk.E))
            return

        # Get the current batch palette
        if hasattr(self, 'batch_palette_var'):
            palette = self.batch_palette_var.get()
        else:
            palette = 'tab10'
    
        print(f"Using batch palette: {palette}")
        print(f"Filtered channels: {list(filtered_stats.keys())}")
    
        # Helper function to get colors based on palette
        def get_palette_colors(n_colors):
            try:
                if palette in colorblind_friendly_palettes:
                    if isinstance(colorblind_friendly_palettes[palette], list):
                        color_list = colorblind_friendly_palettes[palette]
                        colors = [color_list[i % len(color_list)] for i in range(n_colors)]
                    else:
                        cmap = colorblind_friendly_palettes[palette]
                        colors = cmap(np.linspace(0, 1, n_colors))
                elif palette in plt.colormaps():
                    cmap = plt.get_cmap(palette)
                    colors = cmap(np.linspace(0, 1, n_colors))
                else:
                    colors = plt.cm.tab10(np.linspace(0, 1, n_colors))
                return colors
            except Exception as e:
                return plt.cm.tab10(np.linspace(0, 1, n_colors))
    
        # Define ALL plots to display - same as before
        plot_categories = {
            'Environment': [
                ("air_temp", "Air Temperature (Mean ± SEM)", "Temperature (ºC)"),
                ("saturating_air_humidity", "Saturation Humidity (Mean ± SEM)", "Saturation Humidity (mmol/mol)"),
                ("absorbed_radiation", "Absorbed Radiation (Mean ± SEM)", "Radiation (cal/cm²s)"),
                ("relative_air_humidity", "Relative Air Humidity (Mean ± SEM)", "Humidity (%)"),
                ("VPD", "Vapor Pressure Deficit (Mean ± SEM)", "VPD (kPa)"),
                ("rad_ind_leaf_temp_increase", "Radiation-induced Leaf Temperature Increase (Mean ± SEM)", "Temperature Increase (ºC)")
            ],
            'Base_Fluxes': [
                ("CO2_exchange_rate", "CO₂ Exchange Rate (Mean ± SEM)", "CO₂ Exchange (μmol/m²s)"),
                ("Transpiration_H2O_evol_rate", "Transpiration Rate (Mean ± SEM)", "Transpiration (mmol/m²s)"),
                ("transp_ind_leaf_temp_depr", "Transpiration-induced Leaf Temperature Depression (Mean ± SEM)", "Temperature Depression (ºC)"),
                ("leaf_temp_c", "Leaf Temperature (Mean ± SEM)", "Leaf Temperature (ºC)")
            ],
            'Humidity_Vapor': [
                ("leaf_air_hum_grad", "Leaf Air Humidity Gradient (Mean ± SEM)", "Gradient (mmol/mol)"),
                ("satur_hum_at_leaf_temp_c", "Saturation Humidity at Leaf Temperature (Mean ± SEM)", "Saturation Humidity (mmol/mol)"),
                ("mean_vapor_pressure", "Mean Vapor Pressure (Mean ± SEM)", "Pressure (mbar)")
            ],
            'Conductances': [
                ("overall_conductance", "Overall Conductance (Mean ± SEM)", "Conductance (mmol/m²s)"),
                ("corr_d_overall_conductance", "Corrected Overall Conductance (Mean ± SEM)", "Conductance (cm/s)"),
                ("massflow_correction_for_overall_resistance", "Massflow Correction (Mean ± SEM)", "Correction Factor"),
                ("b_layer_conductance", "Boundary Layer Conductance (Mean ± SEM)", "Conductance (cm/s)"),
                ("corr_leaf_conductance", "Corrected Leaf Conductance (Mean ± SEM)", "Conductance (cm/s)"),
                ("corr_stomatal_conductance", "Corrected Stomatal Conductance (Mean ± SEM)", "Conductance (cm/s)"),
                ("Stomatal_conductance_corrected", "Stomatal Conductance Corrected (Mean ± SEM)", "Conductance (mmol/m²s)"),
                ("normalised_corr_stomatal_conductance", "Normalised Stomatal Conductance (Mean ± SEM)", "Normalised Conductance"),
                ("corr_uncorr_stom_cond", "Corr/Uncorr Stomatal Conductance Ratio (Mean ± SEM)", "Ratio"),
                ("corr_stomatal_change_rate", "Corrected Stomatal Change Rate (Mean ± SEM)", "Change Rate (mmol/m²s)"),
                ("relat_stom_cond_change_rate", "Relative Stomatal Conductance Change Rate (Mean ± SEM)", "Change Rate")
            ],
            'Ozone': [
                ("overall_O3_uptake_rate_c", "Overall O₃ Uptake Rate (Mean ± SEM)", "Uptake Rate (nmol/m²s)"),
                ("corr_conductance_for_O3_c", "Corrected Conductance for O₃ (Mean ± SEM)", "Conductance (mmol/m²s)"),
                ("corr_stomatal_O3_uptake_rate", "Stomatal O₃ Uptake (Mean ± SEM)", "O₃ Uptake (nmol O₃/m²s)"),
                ("corr_cumul_O3_dose", "Cumulative O₃ Dose (Mean ± SEM)", "O₃ Dose (μmol/m²)")
            ],
            'CO2': [
                ("massflow_corr_d_ca", "Massflow Corrected Ambient CO₂ (Mean ± SEM)", "CO₂ (μmol CO₂/mol air)"),
                ("massflow_corr_for_CO2", "Massflow Correction for CO₂ (Mean ± SEM)", "Correction Factor"),
                ("mflowcorr_Ca_gradient", "Massflow Correction for CO₂ Gradient (Mean ± SEM)", "Correction Factor"),
                ("corr_uncorr_ca", "Corr/Uncorr Ambient CO₂ Ratio (Mean ± SEM)", "Ratio"),
                ("stom_resist_to_CO2", "Stomatal Resistance to CO₂ (Mean ± SEM)", "Resistance (m²s/mmol)"),
                ("b_layer_resist_to_CO2", "Boundary Layer Resistance to CO₂ (Mean ± SEM)", "Resistance (m²s/mmol)"),
                ("corr_ted_CO2_grad", "Corrected CO₂ Gradient (Mean ± SEM)", "CO₂ Gradient (μmol CO₂/mol air)"),
                ("intercell_CO2_conc_in_gas_phase", "Intercellular CO₂ (Mean ± SEM)", "CO₂ (μmol CO₂/mol air)"),
                ("uncorr_ci", "Uncorrected Intercellular CO₂ (Mean ± SEM)", "CO₂ (μmol CO₂/mol air)"),
                ("corr_uncorr_ci", "Corr/Uncorr Intercellular CO₂ Ratio (Mean ± SEM)", "Ratio")
            ],
            'Mesophyll': [
                ("mesophyll_conductance_for_CO2", "Mesophyll Conductance (Mean ± SEM)", "Conductance (mmol/m²s)"),
                ("uncorr_gm_prima", "Uncorrected Mesophyll Conductance (Mean ± SEM)", "Conductance (mmol/m²s)"),
                ("corr_uncorr_gm_prima", "Corr/Uncorr Mesophyll Conductance Ratio (Mean ± SEM)", "Ratio"),
                ("CO2_comp_point", "CO₂ Compensation Point (Mean ± SEM)", "CO₂ (μmol CO₂/mol air)"),
                ("uncorr_gamma", "Uncorrected CO₂ Compensation Point (Mean ± SEM)", "CO₂ (μmol CO₂/mol air)"),
                ("corr_uncorr_gamma", "Corr/Uncorr Compensation Point Ratio (Mean ± SEM)", "Ratio")
            ],
            'Efficiency': [
                ("WUEi", "Intrinsic Water Use Efficiency (Mean ± SEM)", "WUEi (μmol CO₂ / mol H₂O)"),
                ("WUE", "Water Use Efficiency (Mean ± SEM)", "WUE (μmol CO₂ / mmol H₂O)")
            ]
        }
    
        row = 0
        for category_name, metrics_list in plot_categories.items():
            # Add category header
            header_frame = ttk.Frame(self.batch_plots_frame)
            header_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), padx=10, pady=(20, 5), columnspan=2)
            ttk.Label(header_frame, text=category_name, 
                      font=("Arial", 12, "bold")).grid(row=0, column=0, sticky=tk.W)
            row += 1
    
            for metric, title, ylabel in metrics_list:
                # Check if metric exists in any selected channel
                metric_exists = False
                for channel, stats in filtered_stats.items():
                    for time_point in stats.values():
                        if metric in time_point:
                            metric_exists = True
                            break
                    if metric_exists:
                        break
        
                if not metric_exists:
                    print(f"Warning: {metric} not found in batch results, skipping")
                    continue
        
                # Create figure
                fig, ax = plt.subplots(figsize=(14, 6))
        
                # Sort selected channels alphabetically
                sorted_channels = sorted(filtered_stats.keys())
                num_channels = len(sorted_channels)
        
                # Get colors using the palette
                colors = get_palette_colors(num_channels)
        
                # Plot data for each selected channel with error bars
                for i, channel in enumerate(sorted_channels):
                    stats = filtered_stats[channel]
                    # Sort time points numerically
                    time_points = sorted(stats.keys())
                    means = []
                    errors = []
                    valid_times = []
            
                    for t in time_points:
                        if metric in stats[t]:
                            means.append(stats[t][metric]['mean'])
                            errors.append(stats[t][metric]['sem'])
                            valid_times.append(t)
            
                    if means:
                        # Plot with error bars - one point per time point
                        ax.errorbar(valid_times, means, yerr=errors, 
                                   label=channel, marker='o', capsize=3, 
                                   linewidth=2, markersize=5, color=colors[i], alpha=0.7)
        
                ax.set_title(title, fontsize=14, fontweight='bold')
                ax.set_xlabel('Time (minutes)', fontsize=12)
                ax.set_ylabel(ylabel, fontsize=12)
        
                # Place legend on right side
                ax.legend(bbox_to_anchor=(1.02, 1), loc='upper left', 
                         fontsize=9, framealpha=0.9, borderaxespad=0)
        
                ax.grid(True, alpha=0.3, linestyle='--')
        
                # Adjust layout for right-side legend
                plt.tight_layout()
                plt.subplots_adjust(right=0.85)
        
                # Embed plot
                plot_container = ttk.Frame(self.batch_plots_frame)
                plot_container.grid(row=row, column=0, sticky=(tk.W, tk.E), 
                                   padx=10, pady=10, columnspan=2)
        
                # Add metric label
                ttk.Label(plot_container, text=title, font=("Arial", 10, "bold")).grid(
                    row=0, column=0, sticky=tk.W)
        
                canvas = FigureCanvasTkAgg(fig, plot_container)
                canvas.draw()
                canvas.get_tk_widget().grid(row=1, column=0, sticky=(tk.W, tk.E))
        
                # Download button
                btn = ttk.Button(plot_container, text=f"Download {title}",
                                command=lambda f=fig, t=title: self.download_plot(f, t))
                btn.grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
        
                row += 1
    
            # Add a separator between categories
            if category_name != list(plot_categories.keys())[-1]:
                separator = ttk.Separator(self.batch_plots_frame, orient='horizontal')
                separator.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=10, padx=10, columnspan=2)
                row += 1
    
        # Add dual plot for vapor pressure comparison
        self.add_dual_plot_to_batch_results(filtered_stats, row)
    
        # Configure grid
        self.batch_plots_frame.columnconfigure(0, weight=1)

    def add_dual_plot_to_batch_results(self, stats_by_channel, current_row):
        """Add dual plot for vapor pressure comparison to batch results"""
    
        # Check if both vapor pressure metrics exist
        metric1 = "saturation_vapor_pressure_inside_leaves"
        metric2 = "vapor_pressure_around_leaves"
    
        # Check if both metrics exist in any channel
        metrics_exist = False
        for channel, stats in stats_by_channel.items():
            for time_point in stats.values():
                if metric1 in time_point and metric2 in time_point:
                    metrics_exist = True
                    break
            if metrics_exist:
                break
    
        if not metrics_exist:
            print("Vapor pressure metrics not found, skipping dual plot")
            return
    
        # Get the current palette
        palette = self.palette_var.get() if hasattr(self, 'palette_var') else 'tab10'
    
        # Helper function to get colors based on palette
        def get_palette_colors(n_colors):
            try:
                if palette in colorblind_friendly_palettes:
                    if isinstance(colorblind_friendly_palettes[palette], list):
                        color_list = colorblind_friendly_palettes[palette]
                        colors = [color_list[i % len(color_list)] for i in range(n_colors)]
                    else:
                        cmap = colorblind_friendly_palettes[palette]
                        colors = cmap(np.linspace(0, 1, n_colors))
                elif palette in plt.colormaps():
                    cmap = plt.get_cmap(palette)
                    colors = cmap(np.linspace(0, 1, n_colors))
                else:
                    colors = plt.cm.tab10(np.linspace(0, 1, n_colors))
                return colors
            except Exception as e:
                return plt.cm.tab10(np.linspace(0, 1, n_colors))
    
        # Create dual plot
        fig, ax = plt.subplots(figsize=(14, 6))
    
        # Sort base channels alphabetically
        sorted_channels = sorted(stats_by_channel.keys())
        num_channels = len(sorted_channels)
    
        # Get colors using the palette
        colors = get_palette_colors(num_channels)
    
        # Plot data for each base channel with error bars for both metrics
        for i, channel in enumerate(sorted_channels):
            stats = stats_by_channel[channel]
            time_points = sorted(stats.keys())
        
            # Get data for first metric
            means1 = []
            errors1 = []
            valid_times1 = []
        
            for t in time_points:
                if metric1 in stats[t]:
                    means1.append(stats[t][metric1]['mean'])
                    errors1.append(stats[t][metric1]['sem'])
                    valid_times1.append(t)
        
            # Get data for second metric
            means2 = []
            errors2 = []
            valid_times2 = []
        
            for t in time_points:
                if metric2 in stats[t]:
                    means2.append(stats[t][metric2]['mean'])
                    errors2.append(stats[t][metric2]['sem'])
                    valid_times2.append(t)
        
            # Plot both metrics for this channel
            if means1:
                ax.errorbar(valid_times1, means1, yerr=errors1, 
                           label=f'{channel} - Inside Leaves', marker='o', 
                           linestyle='-', linewidth=2, markersize=5, 
                           color=colors[i], alpha=0.8)
        
            if means2:
                ax.errorbar(valid_times2, means2, yerr=errors2, 
                           label=f'{channel} - Around Leaves', marker='s', 
                           linestyle='--', linewidth=2, markersize=5, 
                           color=colors[i], alpha=0.6)
    
        ax.set_title('Vapor Pressure: Inside vs Around Leaves (Mean ± SEM)', 
                    fontsize=14, fontweight='bold')
        ax.set_xlabel('Time (minutes)', fontsize=12)
        ax.set_ylabel('Vapor Pressure (mbar)', fontsize=12)
    
        # Place legend on right side
        ax.legend(bbox_to_anchor=(1.02, 1), loc='upper left', 
                 fontsize=8, framealpha=0.9, borderaxespad=0, ncol=2)
    
        ax.grid(True, alpha=0.3, linestyle='--')
    
        # Adjust layout for right-side legend
        plt.tight_layout()
        plt.subplots_adjust(right=0.85)
    
        # Embed plot
        plot_container = ttk.Frame(self.batch_plots_frame)
        plot_container.grid(row=current_row, column=0, sticky=(tk.W, tk.E), 
                           padx=10, pady=10, columnspan=2)
    
        # Add metric label
        ttk.Label(plot_container, text="Vapor Pressure Comparison", 
                  font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W)
    
        canvas = FigureCanvasTkAgg(fig, plot_container)
        canvas.draw()
        canvas.get_tk_widget().grid(row=1, column=0, sticky=(tk.W, tk.E))
    
        # Download button
        btn = ttk.Button(plot_container, text="Download Vapor Pressure Comparison",
                        command=lambda f=fig, t="Vapor_Pressure_Comparison": self.download_plot(f, t))
        btn.grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
    
    
    
    def enable_apply_button(self):
        """Enable apply button whenever sliders are changed from initial values"""
        if not hasattr(self, 'apply_changes_btn'):
            return
        
        # Always enable if the button exists and we have data
        if self.original_calculated_data is not None:
            self.apply_changes_btn.config(state='normal')
            if hasattr(self, 'slider_info_label'):
                self.slider_info_label.config(
                    text="Changes pending - click 'Apply Changes'", 
                    foreground="orange"
                )
    
    def apply_time_changes(self):
        # Get current entry values and convert to integers
        try:
            remove_start = int(self.remove_start_var.get())
            self.remove_start_var.set(str(remove_start))
        except ValueError:
            remove_start = 0
            self.remove_start_var.set("0")
    
        try:
            remove_end = int(self.remove_end_var.get())
            self.remove_end_var.set(str(remove_end))
        except ValueError:
            remove_end = 0
            self.remove_end_var.set("0")
    
        try:
            time0_shift = int(self.time0_shift_var.get())
            self.time0_shift_var.set(str(time0_shift))
        except ValueError:
            time0_shift = 0
            self.time0_shift_var.set("0")

        # Show processing message
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="ADJUSTING TIME FRAME,\nIT WILL NOT BE PROPAGATED TO EXTRA TABS AUTOMATICALLY,\nPLEASE WAIT...", 
                foreground="black",
                background="orange",
                font=("Arial", 9, "bold")
            )
            self.root.update()

        # Recalculate ORIGINAL data with new time parameters
        print(f"\nApplying time changes - remove_start={remove_start}, remove_end={remove_end}, time0_shift={time0_shift}")
        self.calculate_original_data()

        # IMPORTANT: Recalculate modified data with all current parameters
        if any(self.parameters_applied.values()):
            print(f"Parameters have been applied, recalculating modified data")
            self.recalculate_modified_data()
        else:
            print(f"No parameters applied, using original data")
            self.modified_calculated_data = self.original_calculated_data.copy()
    
        # Update displays
        self.update_calculated_data_display()
        self.update_time_range_info()
        self.update_all_plots()

        # Disable the Apply Changes button after applying
        self.apply_changes_btn.config(state='disabled')

        # Store current values as new baseline
        self.initial_slider_values = {
            'remove_start': self.remove_start_var.get(),
            'remove_end': self.remove_end_var.get(),
            'time0_shift': self.time0_shift_var.get()
        }

        # Update time information display
        self.update_shift_time_info()

        # Update info label
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(text="Changes applied successfully!", foreground="green")

        # Reset info label after delay - ADD THIS SECTION
        self.root.after(3000, lambda: 
            self.slider_info_label.config(
                text="Enter values, then click 'Apply Changes'", 
                foreground="blue",
                background="",  # Reset to default background
                font=("Arial", 8)  # Reset to normal font
            ) if hasattr(self, 'slider_info_label') else None
        )

        # Debug
        self.debug_data_flow()
    
    def sync_darkness_with_time_changes(self):
        """Sync darkness times with current time parameters"""
        if not self.darkness_settings:
            return
        
        new_darkness_settings = {}
        
        for channel_str, dark_info in self.darkness_settings.items():
            if dark_info['type'] == 'manual':
                # For manual times, we need to adjust them based on time parameters
                adjusted_times = []
                
                # Get current data for this channel
                df = self.get_data_for_plots()
                if df is not None:
                    channel_data = df[df['Channel'].astype(str) == channel_str]
                    
                    for original_time in dark_info['times']:
                        # Find the closest matching time in current data
                        if not channel_data.empty:
                            time_diffs = abs(channel_data['time_repeated'] - original_time)
                            if not time_diffs.empty and time_diffs.min() < 0.01:
                                closest_idx = time_diffs.idxmin()
                                adjusted_times.append(channel_data.loc[closest_idx, 'time_repeated'])
                
                if adjusted_times:
                    new_darkness_settings[channel_str] = {
                        'type': 'manual',
                        'times': adjusted_times
                    }
            else:
                # For auto-detected, keep as is (will be re-detected)
                new_darkness_settings[channel_str] = dark_info
        
        self.darkness_settings = new_darkness_settings
    
    def update_time_range_info(self):
        df = self.get_data_for_plots()
        if df is None or len(df) == 0:
            self.time_range_text.delete(1.0, tk.END)
            self.time_range_text.insert(1.0, "No data available.")
            return
            
        min_time = df['time_repeated'].min()
        max_time = df['time_repeated'].max()
        total_points = len(df)
        channels = df['Channel'].unique()
        
        # Get time interval information
        time_interval = self.get_time_interval_minutes()
        try:
            time0_shift = int(self.time0_shift_var.get())
        except ValueError:
            time0_shift = 0
        
        # Check if we have negative values
        has_negative = min_time < 0
        
        info_text = f"""TIME FRAME SUMMARY:
        {'-' * 40}
        Overall range: {min_time:.1f} - {max_time:.1f} min
        Duration: {max_time - min_time:.1f} min
        Total points: {total_points}
        Channels: {len(channels)}
        """
        
        if time_interval is not None:
            info_text += f"Time interval: {time_interval:.2f} min\n"
        
        if time0_shift != 0 and time_interval is not None:
            total_shift = time0_shift * time_interval
            info_text += f"Time shift applied: {time0_shift} intervals = {total_shift:.2f} min\n"
        
        if has_negative:
            info_text += f"Note: Negative values indicate time points before 'time zero'\n"
        
        info_text += "\n"
        
        if len(channels) <= 8:
            for channel in channels:
                channel_data = df[df['Channel'] == channel]
                ch_min = channel_data['time_repeated'].min()
                ch_max = channel_data['time_repeated'].max()
                ch_points = len(channel_data)
                info_text += f"{channel:15}: {ch_min:6.1f}-{ch_max:6.1f} min ({ch_points:3d} pts)\n"
        
        self.time_range_text.delete(1.0, tk.END)
        self.time_range_text.insert(1.0, info_text)
    
    def on_channel_selection_change(self, event=None):
        """Handle channel selection changes in the listbox"""
        # Get selected channels from listbox (these are display names)
        selected_indices = self.channel_listbox.curselection()
        old_selection = self.selected_channels.copy() if hasattr(self, 'selected_channels') else []
        new_selection = [self.channel_listbox.get(i) for i in selected_indices]
    
        # Determine which channels were added or removed
        removed_channels = [ch for ch in old_selection if ch not in new_selection]
        added_channels = [ch for ch in new_selection if ch not in old_selection]
    
        # Update selected channels
        self.selected_channels = new_selection
    
        # Show status message about channel visibility BEFORE updates
        if removed_channels or added_channels:
            if hasattr(self, 'slider_info_label'):
                if removed_channels:
                    channel_text = ', '.join(removed_channels)
                    message = f"Hiding channel{'s' if len(removed_channels) > 1 else ''}: {channel_text}"
                    self.slider_info_label.config(
                        text=message,
                        foreground="white",
                        background="red",
                        font=("Arial", 9, "bold")
                    )
                elif added_channels:
                    channel_text = ', '.join(added_channels)
                    message = f"Showing channel{'s' if len(added_channels) > 1 else ''}: {channel_text}"
                    self.slider_info_label.config(
                        text=message,
                        foreground="white",
                        background="green",
                        font=("Arial", 9, "bold")
                    )
                # Force immediate update
                self.root.update()
    
        # Update all displays with current selection
        self.update_calculated_data_display()
        self.update_all_plots()
    
        # Also update darkness and radiation tables if they exist
        if hasattr(self, 'darkness_data') and self.darkness_data is not None:
            self.update_darkness_data_table()
        if hasattr(self, 'radiation_data') and self.radiation_data is not None:
            self.update_radiation_data_table()
    
        # After updates, schedule clearing the message (but with longer delay)
        if removed_channels or added_channels:
            self.root.after(2000, lambda: 
                self.slider_info_label.config(
                    text="Enter values, then click 'Apply Changes'", 
                    foreground="blue",
                    background="",
                    font=("Arial", 8)
                ) if hasattr(self, 'slider_info_label') else None
            )
                    
    def on_time_param_change(self, event=None):
        """Handle changes to time-related parameters - only called manually"""
        if self.original_calculated_data is None:
            return
        
        # Get current slider values
        remove_start = self.remove_start_var.get()
        remove_end = self.remove_end_var.get()
        time0_shift = self.time0_shift_var.get()
        
        # Recalculate data with new time parameters
        self.calculate_original_data()
        
        # Update modified data if parameters were applied
        if any(self.parameters_applied.values()):
            self.modified_calculated_data = self.create_modified_data()
        else:
            self.modified_calculated_data = self.original_calculated_data.copy()
        
        # Update darkness data if applied
        if self.darkness_applied:
            self.update_darkness_data()
        
        # Update radiation data if applied
        if self.radiation_applied:
            self.update_radiation_data()
        
        # Update all displays
        self.update_calculated_data_display()
        self.update_time_range_info()
        self.update_all_plots()
        
        # Update slider ranges based on new data
        self.update_slider_ranges()
    
    def update_slider_ranges(self):
        """Update slider ranges based on current data"""
        if self.original_calculated_data is None:
            return
        
        df = self.original_calculated_data
        if df is None or len(df) == 0:
            return
        
        # Find the channel with the minimum number of points
        channels = df['Channel'].unique()
        min_points = float('inf')
        
        for channel in channels:
            channel_data = df[df['Channel'] == channel]
            if len(channel_data) < min_points:
                min_points = len(channel_data)
        
        if min_points == float('inf') or min_points < 2:
            min_points = 10  # default fallback
        
        # Calculate safe maximum values
        max_remove = max(0, min_points - 1)  # Leave at least 1 point
        
        # Update info label
        self.update_slider_info_label()
    
    def update_slider_info_label(self):
        """Update the information label about values ranges"""
        if self.original_calculated_data is None:
            self.slider_info_label.config(text="Load data to see values limits")
            return
        
        df = self.original_calculated_data
        channels = df['Channel'].unique()
        
        # Calculate min points across channels
        points_per_channel = {}
        for channel in channels:
            channel_data = df[df['Channel'] == channel]
            points_per_channel[str(channel)] = len(channel_data)
        
        if points_per_channel:
            min_points = min(points_per_channel.values())
            max_remove = max(0, min_points - 2)
            
            info_text = f"Max removals: {max_remove} (based on channel with {min_points} points)"
            
            if len(points_per_channel) > 1:
                points_info = ", ".join([f"{ch}: {pts}" for ch, pts in points_per_channel.items()])
                if len(points_info) > 80:  # Truncate if too long
                    points_info = f"{len(points_per_channel)} channels, {min_points}-{max(points_per_channel.values())} points each"
                info_text += f" | Points per channel: {points_info}"
            
            self.slider_info_label.config(text=info_text)
    
    def on_entry_change(self):
        """Handle entry changes - enable Apply Changes button"""
        if not hasattr(self, 'apply_changes_btn'):
            return
        
        # Check if entries have changed from initial values
        if self.check_entry_changes():
            self.apply_changes_btn.config(state='normal')
            if hasattr(self, 'slider_info_label'):
                self.slider_info_label.config(
                    text="Changes pending - click 'Apply Changes'", 
                    foreground="orange"
                )
    
    def check_entry_changes(self):
        """Check if entry values have changed from their initial values"""
        if not hasattr(self, 'initial_entry_values'):
            return False
        
        current_values = {
            'remove_start': self.remove_start_var.get(),
            'remove_end': self.remove_end_var.get(),
            'time0_shift': self.time0_shift_var.get()
        }
        
        for key, current in current_values.items():
            if current != self.initial_entry_values[key]:
                return True
        return False
    
    def update_entry_info_label(self, max_remove=None, min_points=None):
        """Update the information label about entry limits"""
        if self.original_calculated_data is None:
            self.slider_info_label.config(text="Load data to see limits")
            return
        
        if max_remove is None or min_points is None:
            df = self.original_calculated_data
            channels = df['Channel'].unique()
            
            # Calculate min points across channels
            points_per_channel = {}
            for channel in channels:
                channel_data = df[df['Channel'] == channel]
                points_per_channel[str(channel)] = len(channel_data)
            
            if points_per_channel:
                min_points = min(points_per_channel.values())
                max_remove = max(0, min_points - 2)
            else:
                return
        
        info_text = f"Max removals per channel: {max_remove} (leaves 2+ points, min {min_points} points)"
        
        self.slider_info_label.config(text=info_text)
        
    def setup_coefficient_tables(self, parent, hidden=True):
        # Coefficient table
        coef_data = get_coefficient_table_data()

        self.coef_frame = ttk.LabelFrame(parent, text="Absorbed Radiation Coefficients\nMaximum light intensity values at different heights (µmol m⁻² s⁻¹)", padding="5")
        # Only grid if not hidden
        if not hidden:
            self.coef_frame.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(10, 5), padx=10)
        else:
            self.coef_frame.grid_remove()

        # Create treeview for coefficient table
        columns = list(coef_data.columns)
        self.coef_tree = ttk.Treeview(self.coef_frame, columns=columns, show='headings', height=5)

        for col in columns:
            self.coef_tree.heading(col, text=col)
            self.coef_tree.column(col, width=80, anchor=tk.CENTER)

        for _, row in coef_data.iterrows():
            self.coef_tree.insert('', 'end', values=list(row))

        self.coef_tree.grid(row=0, column=0, sticky=(tk.W, tk.E))

        # Color table
        color_data = get_color_table_data()

        self.color_coef_frame = ttk.LabelFrame(parent, text="Light Intensity Coefficients (µmol m⁻² s⁻¹)", padding="5")
        if not hidden:
            self.color_coef_frame.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=(5, 0), padx=10)
        else:
            self.color_coef_frame.grid_remove()

        # Configure frame to expand properly
        self.color_coef_frame.columnconfigure(0, weight=1)
        self.color_coef_frame.rowconfigure(0, weight=1)

        # Create container for treeview and scrollbar
        tree_container = ttk.Frame(self.color_coef_frame)
        tree_container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        # Create treeview inside container
        columns = list(color_data.columns)
        self.color_tree = ttk.Treeview(tree_container, columns=columns, show='headings', height=8)

        for col in columns:
            self.color_tree.heading(col, text=col)
            self.color_tree.column(col, width=100, anchor=tk.CENTER)

        for _, row in color_data.iterrows():
            self.color_tree.insert('', 'end', values=list(row))

        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.color_tree.yview)
        self.color_tree.configure(yscrollcommand=v_scrollbar.set)

        h_scrollbar = ttk.Scrollbar(tree_container, orient="horizontal", command=self.color_tree.xview)
        self.color_tree.configure(xscrollcommand=h_scrollbar.set)

        # Grid layout
        self.color_tree.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.E, tk.W))

    
    def setup_notebook(self, parent):
        # Create main notebook for tabs
        self.main_notebook = ttk.Notebook(parent)
        self.main_notebook.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        
        # Create main tabs
        self.setup_introduction_tab()
        self.setup_raw_data_tab()
        self.setup_calculated_data_tab()
        self.setup_plots_tab()
        self.setup_darkness_tab()
        self.setup_radiation_tab()
        self.setup_batch_results_tabs()
        self.setup_notes_tab()
        
    def setup_introduction_tab(self):
        tab = ttk.Frame(self.main_notebook)
        self.main_notebook.add(tab, text="Introduction")
        
        intro_text = """
Welcome to the Plant-Invent-Gas-Exchange-Data-Analyzer

Research License (CC BY-NC-SA 4.0)

https://creativecommons.org/licenses/by-nc-sa/4.0


For commercial use, contact: david.lazaro.gimeno@gmail.com

This application allows you to upload and process gas exchange data files from Plant Invent Jyrkki systems in TXT format.

  * Instantly view the raw data in tabular format.

  * Calculate and observe processed data, including various metrics.

  * Custimize channel names, values of clurical columns and visualize easily your initial results.

Obtain the processed data and download your plot at 300 dpi for your publication if liked.

Use the sidebar to upload files, ajust timepoints, select different blind color palettes and initiate data processing.

Timepoint removal will apply to all channels at the same time.

Removing undesired initial points generated while stabilyzing the system, and/or end points , will provide an improved plot.


*********************************************************************************
Batch Processing allows you to include several replicates of the same experiment.
*********************************************************************************

To that aim each replicate file must be labeled as _Rep1.txt, _Rep2.txt

Using an Excel file with the experimental design for each replicate will allow to discriminate randomization (not mandatory). An inspiring example. Leaf area and absorbed radiation are not mandatory columns, but if included, they apply in the batch process calculations, replacing default inputs:
                Channel name          Leaf area  absorbed radiation
Rep1	Channel 1	no_cond1-4_nacl     0.0013478  0.000591
Rep1	Channel 2	no_cond1-6          0.0022008  0.000591
Rep1	Channel 3	no_cond1-6_nacl     0.001985   0.000591
Rep1	Channel 4	no_cond1-6 (2)      0.001684   0.000591
Rep1	Channel 5	cond1_nacl          0.0013812  0.000591
Rep1	Channel 6	WT_nacl             0.0015385  0.000591
Rep1	Channel 7	WT                  0.0018776  0.000591
Rep1	Channel 8	no_cond1-4          0.0024069  0.000591
Rep2	Channel 1	no_cond1-4_nacl (2) 0.0011749  0.000591
Rep2	Channel 2	WT_nacl (2)         0.0011914  0.000591
Rep2	Channel 3	no_cond1-6 (3)      0.0016018  0.000591
Rep2	Channel 4	WT (2)              0.0017024  0.000591
Rep2	Channel 5	cond1_nacl (2)      0.0012284  0.000591
Rep2	Channel 6	no_cond1-6_nacl (2) 0.0017947  0.000591
Rep2	Channel 7	cond1               0.0016419  0.000591
Rep2	Channel 8	no_cond1-4 (2)      0.001884   0.000591
Rep3	Channel 1	no_cond1-6_nacl (3) 0.0020586  0.000591
Rep3	Channel 2	no_cond1-6_nacl (4) 0.0015655  0.000591
Rep3	Channel 3	WT_nacl (3)         0.001388   0.000591
Rep3	Channel 4	no_cond1-4_nacl (3) 0.0018305  0.000591
Rep3	Channel 5	no_cond1-6 (4)      0.0019324  0.000591
Rep3	Channel 6	cond1 (2)           0.0018989  0.000591
Rep3	Channel 7	WT (3)              0.0016374  0.000591
Rep3	Channel 8	no_cond1-4 (3)      0.0023528  0.000591
Rep4	Channel 1	cond1_nacl (3)      0.0008682  0.000591
Rep4	Channel 2	WT_nacl (4)         0.0008682  0.000591
Rep4	Channel 3	no_cond1-4_nacl (4) 0.0010081  0.000591
Rep4	Channel 4	cond1_nacl (4)      0.0010232  0.000591
Rep4	Channel 5	WT (4)              0.0013417  0.000591
Rep4	Channel 6	no_cond1-4 (4)      0.0013368  0.000591
Rep4	Channel 7	cond1 (3)           0.0017162  0.000591
Rep4	Channel 8	no_cond1-4 (5)      0.0018093  0.000591
Rep5	Channel 1	no_cond1-4_nacl (5) 0.0016058  0.000591
Rep5	Channel 2	cond1 (4)           0.0013253  0.000591
Rep5	Channel 3	cond1_nacl (5)      0.0010997  0.000591
Rep5	Channel 4	no_cond1-6 (5)      0.001486   0.000591
Rep5	Channel 5	WT_nacl (5)         0.00088    0.000591
Rep5	Channel 6	WT (5)              0.0011167  0.000591
Rep5	Channel 7	no_cond1-6_nacl (5) 0.001421   0.000591
Rep5	Channel 8	cond1 (5)           0.0017382  0.000591
"""
        
        text_widget = scrolledtext.ScrolledText(tab, wrap=tk.WORD, width=100, height=30)
        text_widget.insert(tk.END, intro_text)
        text_widget.configure(state='disabled')
        text_widget.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), padx=10, pady=10)
        
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        
    def setup_raw_data_tab(self):
        tab = ttk.Frame(self.main_notebook)
        self.main_notebook.add(tab, text="Raw Data")
        
        frame = ttk.Frame(tab)
        frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), padx=10, pady=10)
        
        # Initialize the treeview
        self.raw_tree = ttk.Treeview(frame)
        
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.raw_tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.raw_tree.xview)
        self.raw_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.raw_tree.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.E, tk.W))
        
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        
    def setup_calculated_data_tab(self):
        tab = ttk.Frame(self.main_notebook)
        self.main_notebook.add(tab, text="Interactive Calculated Data")
        
        main_container = ttk.Frame(tab)
        main_container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), padx=10, pady=10)
        
        tree_frame = ttk.Frame(main_container)
        tree_frame.grid(row=0, column=0, columnspan=6, sticky=(tk.N, tk.S, tk.E, tk.W), pady=(0, 10))
        
        # Create a Treeview with custom heading style
        self.calc_tree = ttk.Treeview(tree_frame)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.calc_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.calc_tree.xview)
        self.calc_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.calc_tree.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.E, tk.W))
        
        button_frame = ttk.Frame(main_container)
        button_frame.grid(row=1, column=0, columnspan=6, sticky=(tk.W, tk.E), pady=(0, 10))
        
        button_frame = ttk.Frame(main_container)
        button_frame.grid(row=1, column=0, columnspan=6, sticky=(tk.W, tk.E), pady=(0, 10))
        
        buttons = [
            ("Rename Channels", self.update_channel_names),
            ("Update Leaf Area", self.update_leaf_area),
            ("Update Air Flow Rate", self.update_air_flow),
            ("Update Absorbed Radiation", self.update_radiation),
            ("Update Boundary Layer Cond", self.update_boundary),
            ("Update Cutic Conductance", self.update_cutic)
        ]
        
        for i, (text, command) in enumerate(buttons):
            btn = ttk.Button(button_frame, text=text, command=command)
            btn.grid(row=0 if i < 4 else 1, column=i % 4, padx=2, pady=2, sticky=(tk.W, tk.E))
        
        ttk.Button(main_container, text="Download Data", command=self.download_data).grid(
            row=2, column=0, columnspan=6, pady=(10, 0))
        
        self.channel_params_frame = ttk.LabelFrame(main_container, text="Channel Parameters")
        self.channel_params_frame.grid(row=3, column=0, columnspan=6, sticky=(tk.W, tk.E), pady=(10, 0))
        
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        for i in range(6):
            button_frame.columnconfigure(i, weight=1)
        
    def setup_plot_tab_structure(self, notebook, tab_name, tab_text):
        """Helper method to create a plot tab with both vertical and horizontal scrollbars"""
        tab = ttk.Frame(notebook)
        notebook.add(tab, text=tab_text)
        
        # Create container for scrollbars and canvas
        container = ttk.Frame(tab)
        container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        
        # Create canvas
        canvas = tk.Canvas(container)
        
        # Create scrollbars
        v_scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        h_scrollbar = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        
        # Create scrollable frame
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout
        canvas.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.E, tk.W))
        
        # Store reference
        self.plot_tabs[tab_name] = scrollable_frame
        
        # Configure grid weights
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        
        return tab
    
    def setup_plots_tab(self):
        self.plots_notebook = ttk.Notebook(self.main_notebook)
        self.main_notebook.add(self.plots_notebook, text="Interactive Plots")
        
        self.setup_environment_plots_tab()
        self.setup_base_fluxes_tab()
        self.setup_humidity_vapor_tab()
        self.setup_conductances_tab()
        self.setup_ozone_tab()
        self.setup_co2_tab()
        self.setup_mesophyll_tab()
        self.setup_efficiency_tab()
        
    def setup_environment_plots_tab(self):
        self.setup_plot_tab_structure(self.plots_notebook, 'environment', "Environment")

        
    def setup_base_fluxes_tab(self):
        self.setup_plot_tab_structure(self.plots_notebook, 'base_fluxes', "Base Fluxes")
        
    def setup_humidity_vapor_tab(self):
        self.setup_plot_tab_structure(self.plots_notebook, 'humidity_vapor', "Humidity/Vapor")
        
    def setup_conductances_tab(self):
        self.setup_plot_tab_structure(self.plots_notebook, 'conductances', "Conductances")
        
    def setup_ozone_tab(self):
        self.setup_plot_tab_structure(self.plots_notebook, 'ozone', "Ozone")
        
    def setup_co2_tab(self):
        self.setup_plot_tab_structure(self.plots_notebook, 'co2', "CO2")

        
    def setup_mesophyll_tab(self):
        self.setup_plot_tab_structure(self.plots_notebook, 'mesophyll', "Mesophyll")
        
    def setup_efficiency_tab(self):
        self.setup_plot_tab_structure(self.plots_notebook, 'efficiency', "Efficiency")
        
    def setup_darkness_tab(self):
        """Setup the Light-Dark tab with darkness period input boxes"""
        tab = ttk.Frame(self.main_notebook)
        self.main_notebook.add(tab, text="Light-Dark")
        
        # Create main container with proper grid weights
        main_container = ttk.Frame(tab)
        main_container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), padx=10, pady=10)
        
        # Configure main container to expand
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(1, weight=1)  # Row 1 gets weight for notebook
        
        # Darkness controls - sticky to expand horizontally
        control_frame = ttk.LabelFrame(main_container, text="Darkness Settings", padding="10")
        control_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        control_frame.columnconfigure(0, weight=1)
        
        ttk.Label(control_frame, text="Enter darkness time points (comma-separated):").grid(
            row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Create a frame for darkness inputs with scrollbar
        darkness_inputs_container = ttk.Frame(control_frame)
        darkness_inputs_container.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Create canvas and scrollbar for darkness inputs
        darkness_canvas = tk.Canvas(darkness_inputs_container, height=150)
        darkness_scrollbar = ttk.Scrollbar(darkness_inputs_container, orient="vertical", command=darkness_canvas.yview)
        self.darkness_inputs_frame = ttk.Frame(darkness_canvas)
        
        self.darkness_inputs_frame.bind(
            "<Configure>",
            lambda e: darkness_canvas.configure(scrollregion=darkness_canvas.bbox("all"))
        )
        
        darkness_canvas.create_window((0, 0), window=self.darkness_inputs_frame, anchor="nw")
        darkness_canvas.configure(yscrollcommand=darkness_scrollbar.set)
        
        darkness_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E))
        darkness_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configure columns
        darkness_inputs_container.columnconfigure(0, weight=1)
        self.darkness_inputs_frame.columnconfigure(0, weight=1)
        self.darkness_inputs_frame.columnconfigure(1, weight=1)
        self.darkness_inputs_frame.columnconfigure(2, weight=1)
        self.darkness_inputs_frame.columnconfigure(3, weight=1)
        
        # Buttons
        button_frame = ttk.Frame(control_frame)
        button_frame.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        ttk.Button(button_frame, text="Apply Darkness", command=self.apply_darkness).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        # Add download button for darkness data
        ttk.Button(button_frame, text="Download Darkness Data", command=self.download_darkness_data).grid(
            row=0, column=2, sticky=tk.W, padx=(0, 10))
        
        # Darkness notebook - make it expand
        self.darkness_notebook = ttk.Notebook(main_container)
        self.darkness_notebook.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        
        # Create only TWO tabs: one for all plots and one for data table
        self.setup_darkness_plots_tab()  # Single tab for ALL plots
        self.setup_darkness_data_table_tab()  # Separate tab for data table
    
    def setup_darkness_plots_tab(self):
        """Setup a single tab for ALL darkness plots"""
        tab = ttk.Frame(self.darkness_notebook)
        self.darkness_notebook.add(tab, text="Darkness Plots")
        
        # Create a canvas with both vertical and horizontal scrollbars
        container = ttk.Frame(tab)
        container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        
        canvas = tk.Canvas(container)
        
        v_scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        h_scrollbar = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        canvas.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.E, tk.W))
        
        # Store reference
        self.plot_tabs['darkness_plots'] = scrollable_frame
        
        # Configure grid weights
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
    
    def setup_darkness_data_table_tab(self):
        """Setup a tab to display the darkness data table"""
        tab = ttk.Frame(self.darkness_notebook)  # CHANGED: darkness_notebook instead of darkness_plots_notebook
        self.darkness_notebook.add(tab, text="Darkness Data Table")  # CHANGED: self.darkness_notebook
        
        main_container = ttk.Frame(tab)
        main_container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), padx=10, pady=10)
        
        # Configure main container to expand
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(0, weight=1)
        
        # Create treeview for darkness data
        frame = ttk.Frame(main_container)
        frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        
        self.darkness_tree = ttk.Treeview(frame)
        
        # Configure tags for special rows
        self.darkness_tree.tag_configure('info', background='lightyellow')
        
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.darkness_tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.darkness_tree.xview)
        self.darkness_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.darkness_tree.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.E, tk.W))
        
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        
        # Store reference to the tab frame
        self.plot_tabs['darkness_data_table'] = tab
        
    def update_darkness_data_table(self):
        """Update the darkness data table display with proper column widths"""
        if self.darkness_data is None or self.darkness_tree is None:
            return
    
        # Clear existing items
        for item in self.darkness_tree.get_children():
            self.darkness_tree.delete(item)
    
        # Get the darkness data
        darkness_df = self.darkness_data.copy()
    
        # Apply channel filtering
        if self.selected_channels:
            # Filter logic remains the same...
            valid_channels = []
            for display_channel in self.selected_channels:
                actual_channel = display_channel
                if hasattr(self, 'channel_names') and self.channel_names:
                    if display_channel in self.channel_names.values():
                        for orig_name, new_name in self.channel_names.items():
                            if new_name == display_channel:
                                actual_channel = orig_name
                                break
            
                if actual_channel in darkness_df['Channel'].astype(str).values:
                    valid_channels.append(actual_channel)
                elif display_channel in darkness_df['Channel'].astype(str).values:
                    valid_channels.append(display_channel)
        
            if valid_channels:
                darkness_df = darkness_df[darkness_df['Channel'].astype(str).isin(valid_channels)]
    
        # Check if we have data to display
        if darkness_df is None or len(darkness_df) == 0:
            columns = ['Info']
            self.darkness_tree['columns'] = columns
            self.darkness_tree['show'] = 'headings'
            for col in columns:
                self.darkness_tree.heading(col, text=col)
                self.darkness_tree.column(col, width=300, anchor=tk.CENTER)
            self.darkness_tree.insert('', 'end', values=['No darkness data available'])
            return
    
        # Set up columns with units
        columns = list(darkness_df.columns)
        self.darkness_tree['columns'] = columns
        self.darkness_tree['show'] = 'headings'
    
        for col in columns:
            # Get unit for this column
            unit = UNITS_MAP.get(col, "")
        
            # Create single-line header with unit in parentheses
            if unit:
                header_text = f"{col} ({unit})"
            else:
                header_text = col
            
            self.darkness_tree.heading(col, text=header_text)
        
            # Calculate optimal width and then MAXIMIZE it
            width = self.calculate_optimal_width(darkness_df, col, header_text, is_raw=False)
        
            # MAXIMIZE THE WIDTH - multiply by 1.5 to make it 50% wider
            width = int(width * 1)  # If required, 30% wider for maximum readability
        
            # Set a minimum width to ensure visibility
            min_width = max(150, width)
        
            self.darkness_tree.column(col, width=width, anchor=tk.CENTER, minwidth=min_width)
    
        # Insert data with formatting
        for _, row in darkness_df.iterrows():
            values = []
            for val in row:
                values.append(self.format_value_for_display(val))
            self.darkness_tree.insert('', 'end', values=values)
    
        # Add info about number of rows
        total_rows = len(darkness_df)
        info_row = [''] * len(columns)
        info_row[0] = f"Total rows: {total_rows}"
        self.darkness_tree.insert('', 'end', values=info_row, tags=('info',))
        self.darkness_tree.tag_configure('info', background='lightyellow')
    
    def download_darkness_data(self):
        """Download the darkness data table with two export options"""
        if self.darkness_data is None:
            messagebox.showwarning("Warning", "No darkness data available to download.")
            return

        # Create custom dialog for export options
        dialog = tk.Toplevel(self.root)
        dialog.title("Export Options")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Set icon for the dialog
        self._set_dialog_icon(dialog)
    
        # Center the dialog
        dialog_width = 400
        dialog_height = 250
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        x = (screen_width - dialog_width) // 2
        y = (screen_height - dialog_height) // 2
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
    
        # Add instruction label
        ttk.Label(dialog, text="Select export format:", font=("Arial", 12, "bold")).pack(pady=20)
    
        # Add description
        desc_frame = ttk.Frame(dialog)
        desc_frame.pack(pady=10)
        ttk.Label(desc_frame, text="Choose the format for your exported data:\n\n- RE-IMPORT (28 columns compatible)\n- ALL columns (all calculated columns)", font=("Arial", 10)).pack()
    
        # Variable to store the result
        result = [None]
    
        def on_reimport():
            result[0] = "reimport"
            dialog.destroy()
    
        def on_all_columns():
            result[0] = "all_columns"
            dialog.destroy()
    
        # Button frame
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
    
        ttk.Button(button_frame, text="RE-IMPORT", command=on_reimport, 
                   width=20).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="ALL COLUMNS", command=on_all_columns, 
                   width=20).pack(side=tk.LEFT, padx=10)
    
        # Wait for dialog to close
        self.root.wait_window(dialog)
    
        # Check if user closed the dialog without choosing
        if result[0] is None:
            return
    
        export_mode = result[0]

        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[
                ("TSV/TXT files", "*.txt"),
                ("All files", "*.*")
            ],
            title="Save Darkness Data As"
        )

        if not filename:
            return

        # Ensure .txt extension
        if not filename.endswith('.txt'):
            filename = filename.rsplit('.', 1)[0] + '.txt' if '.' in filename else filename + '.txt'

        try:
        
            # Get current date and time
            now = datetime.now()
            date_str = now.strftime("%Y-%m-%d")
            time_str = now.strftime("%H:%M:%S")
        
            # Get the original filename if available
            original_filename = self.file_path_var.get() if hasattr(self, 'file_path_var') else "Unknown"
            modified_text = f"Modified file: {os.path.basename(original_filename)}"
        
            # Define the original 28 columns (exactly as from raw data)
            original_28_columns = [
                "Channel", "CO2_M", "CO2_R", "H2O_M", "H2O_R", "O3_M", "O3_R",
                "CO2_exch_rate", "H2O_evol_rate", "overall_O3_uptake_rate", "air_temp", 
                "transp_ind_leaf_temp_depr", "leaf_temp", "satur_hum_at_leaf_temp", 
                "leaf_air_hum_grad", "stomatal_cond", "stomatal_O3_uptake_rate", "O3_cumul_dose",
                "stom_change_rate", "normalised_stom_cond", "relstom_condchange_rate", "leaf_area", 
                "air_flow_rate", "absorbed_radiation", "boundary_layer_cond", "cutic_conductance", 
                "temp_calibr", "time"
            ]
        
            # Create a copy for export
            df_export = self.darkness_data.copy()
        
            # Apply column selection based on user choice
            if export_mode == "reimport":
                # Only keep columns that actually exist in the dataframe
                export_columns = [col for col in original_28_columns if col in df_export.columns]
                df_export = df_export[export_columns].copy()
                mode_text = "Re-import compatible (original 28 columns only)"
            else:  # all_columns
                mode_text = "Complete data (all calculated columns)"
        
            # Sort by continuous time (handle day rollover)
            if 'time' in df_export.columns:
                time_seconds = []
                for time_val in df_export['time']:
                    parts = str(time_val).split(':')
                    if len(parts) >= 3:
                        secs = float(parts[0]) * 3600 + float(parts[1]) * 60 + float(parts[2])
                    else:
                        secs = 0.0
                    time_seconds.append(secs)
            
                # Detect day rollover
                day_increment = 0
                full_seconds = []
                prev_sec = time_seconds[0] if time_seconds else 0
                max_day_increment = 0
            
                for sec in time_seconds:
                    if sec < prev_sec:
                        day_increment += 1
                        if day_increment > max_day_increment:
                            max_day_increment = day_increment
                    prev_sec = sec
                    full_seconds.append(sec + day_increment * 24 * 3600)
            
                # Add temporary sort column
                df_export['_sort_time'] = full_seconds
                df_export = df_export.sort_values(['_sort_time', 'Channel'])
                df_export = df_export.drop(columns=['_sort_time'])
        
            # Convert numeric values to string with comma as decimal separator
            for col in df_export.columns:
                if df_export[col].dtype in ['float64', 'float32', 'int64', 'int32']:
                    df_export[col] = df_export[col].apply(
                        lambda x: str(x).replace('.', ',') if pd.notna(x) else ''
                    )
        
            # Create units row
            units_row = []
            for col in df_export.columns:
                unit = UNITS_MAP.get(col, "")
                # Replace special characters for plain text
                unit = unit.replace('₂', '2').replace('₃', '3').replace('ₒ', 'o')
                unit = unit.replace('µ', 'u').replace('μ', 'u')
                units_row.append(unit)
        
            # Add empty column at the beginning
            headers = df_export.columns.tolist()
            headers.insert(0, "")
            units_row.insert(0, "")
        
            # Add empty column to dataframe
            df_export.insert(0, "", "")
        
            # Write as TSV (tab-separated) with .txt extension
            with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f, delimiter='\t')
            
                # Row 1: Metadata
                writer.writerow(["", f"Date: {date_str}", f"Time: {time_str}", modified_text, f"Mode: {mode_text}"])
            
                # Row 2: Empty row
                writer.writerow([])
            
                # Row 3: Headers
                writer.writerow(headers)
            
                # Row 4: Units row
                writer.writerow(units_row)
            
                # Write data starting from row 5
                for _, row in df_export.iterrows():
                    writer.writerow(row)
        
            total_darkness_periods = 0
            if self.darkness_settings:
                for channel, dark_info in self.darkness_settings.items():
                    total_darkness_periods += len(dark_info.get('times', []))
        
            # Count columns exported
            num_cols_exported = len(df_export.columns) - 1
        
            # Check if multi-day experiment was detected
            multi_day_note = ""
            if 'max_day_increment' in locals() and max_day_increment > 0:
                multi_day_note = f"\nMulti-day experiment detected ({max_day_increment + 1} days) - Sorted correctly by continuous time"
        
            messagebox.showinfo("Success", 
                f"Darkness data saved to {filename}\n\n"
                f"Export mode: {mode_text}\n"
                f"Format: TSV (tab-separated) with .txt extension\n"
                f"Decimal separator: Comma (,)\n"
                f"Sorting: By continuous time (handles day rollover){multi_day_note}\n"
                f"Total rows: {len(self.darkness_data)}\n"
                f"Channels: {len(self.darkness_data['Channel'].unique())}\n"
                f"Columns exported: {num_cols_exported}\n"
                f"Darkness periods: {total_darkness_periods}\n\n"
                f"File structure:\n"
                f"  Row 1: Metadata (Date, Time, Source file, Export mode)\n"
                f"  Row 2: Empty\n"
                f"  Row 3: Column headers (starting with empty column A)\n"
                f"  Row 4: Units\n"
                f"  Row 5+: Data with empty first column")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {str(e)}")
            traceback.print_exc()

        
    def update_darkness_for_slider_changes(self):
        """Update darkness settings when time parameters change"""
        if not self.darkness_applied or not self.darkness_settings:
            return
        
        # Get current data to understand new time mapping
        df = self.get_data_for_plots()
        if df is None:
            return
        
        new_darkness_settings = {}
        
        # For each channel, map old darkness times to new time_repeated values
        for channel_str, dark_info in self.darkness_settings.items():
            # Try to find the channel in current data
            channel_data = None
            
            # First try the display name
            channel_data = df[df['Channel'].astype(str) == channel_str]
            
            # If not found, check if it's a renamed channel
            if channel_data.empty and channel_str in self.channel_names.values():
                # Find the original name for this renamed channel
                for orig_name, new_name in self.channel_names.items():
                    if new_name == channel_str:
                        channel_data = df[df['Channel'].astype(str) == orig_name]
                        break
            
            if channel_data.empty:
                continue
            
            if dark_info['type'] == 'manual':
                # For manual times, find closest points in current data
                adjusted_times = []
                for old_time in dark_info['times']:
                    if not channel_data.empty:
                        # Find time point in current data
                        closest_idx = None
                        min_diff = float('inf')
                        for idx, row in channel_data.iterrows():
                            diff = abs(row['time_repeated'] - old_time)
                            if diff < min_diff:
                                min_diff = diff
                                closest_idx = idx
                        
                        if closest_idx is not None and min_diff < 0.01:
                            adjusted_times.append(channel_data.loc[closest_idx, 'time_repeated'])
                
                if adjusted_times:
                    new_darkness_settings[channel_str] = {
                        'type': 'manual',
                        'times': adjusted_times
                    }
            else:
                # For auto-detected, keep as is (will be re-detected)
                new_darkness_settings[channel_str] = dark_info
        
        self.darkness_settings = new_darkness_settings
    
    def setup_radiation_environment_tab(self):
        """Setup the environment tab for radiation plots"""
        self.setup_plot_tab_structure(self.radiation_plots_notebook, 'radiation_environment', "Environment")

    def setup_radiation_fluxes_tab(self):
        """Setup the fluxes tab for radiation plots"""
        self.setup_plot_tab_structure(self.radiation_plots_notebook, 'radiation_fluxes', "Fluxes")
        
    def setup_radiation_tab(self):
        """Redesign Radiation tab to match Darkness tab structure"""
        tab = ttk.Frame(self.main_notebook)
        self.main_notebook.add(tab, text="Red+Blue Light")
        
        main_container = ttk.Frame(tab)
        main_container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), padx=10, pady=10)
        
        # Configure main container to expand
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(1, weight=1)  # Row 1 gets weight for notebook
        
        # Radiation controls
        control_frame = ttk.LabelFrame(main_container, text="Radiation Interval Settings", padding="10")
        control_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        control_frame.columnconfigure(0, weight=1)
        
        ttk.Label(control_frame, text="Enter intervals (format: start-end:value; multiple separated by semicolon):").grid(
            row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # Create a frame for radiation inputs
        radiation_inputs_container = ttk.Frame(control_frame)
        radiation_inputs_container.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Create canvas and scrollbar for radiation inputs
        radiation_canvas = tk.Canvas(radiation_inputs_container, height=150)
        radiation_scrollbar = ttk.Scrollbar(radiation_inputs_container, orient="vertical", command=radiation_canvas.yview)
        self.radiation_inputs_frame = ttk.Frame(radiation_canvas)
        
        self.radiation_inputs_frame.bind(
            "<Configure>",
            lambda e: radiation_canvas.configure(scrollregion=radiation_canvas.bbox("all"))
        )
        
        radiation_canvas.create_window((0, 0), window=self.radiation_inputs_frame, anchor="nw")
        radiation_canvas.configure(yscrollcommand=radiation_scrollbar.set)
        
        radiation_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E))
        radiation_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configure columns
        radiation_inputs_container.columnconfigure(0, weight=1)
        self.radiation_inputs_frame.columnconfigure(0, weight=1)
        self.radiation_inputs_frame.columnconfigure(1, weight=1)
        self.radiation_inputs_frame.columnconfigure(2, weight=1)
        self.radiation_inputs_frame.columnconfigure(3, weight=1)
        
        # Buttons
        button_frame = ttk.Frame(control_frame)
        button_frame.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        ttk.Button(button_frame, text="Apply Radiation Intervals", command=self.apply_radiation).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10))

        
        # ADD DOWNLOAD BUTTON FOR RADIATION DATA
        ttk.Button(button_frame, text="Download Radiation Data", command=self.download_radiation_data).grid(
            row=0, column=2, sticky=tk.W, padx=(0, 10))
        
        # Radiation notebook - make it expand
        self.radiation_notebook = ttk.Notebook(main_container)  # CHANGED NAME: radiation_notebook instead of radiation_plots_notebook
        self.radiation_notebook.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        
        # Create only TWO tabs: one for all plots and one for data table (EXACTLY LIKE DARKNESS TAB)
        self.setup_radiation_plots_tab()  # Single tab for ALL plots
        self.setup_radiation_data_table_tab()  # Separate tab for data table
         
    def setup_radiation_data_table_tab(self):
        """Setup a tab to display the radiation data table"""
        tab = ttk.Frame(self.radiation_notebook)  # CHANGED: radiation_notebook
        self.radiation_notebook.add(tab, text="Radiation Data Table")  # CHANGED: self.radiation_notebook
        
        main_container = ttk.Frame(tab)
        main_container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), padx=10, pady=10)
        
        # Configure main container to expand
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(0, weight=1)
        
        # Create treeview for radiation data
        frame = ttk.Frame(main_container)
        frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        
        self.radiation_tree = ttk.Treeview(frame)
        
        # Configure tags for special rows
        self.radiation_tree.tag_configure('info', background='lightyellow')
        
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.radiation_tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.radiation_tree.xview)
        self.radiation_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.radiation_tree.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.E, tk.W))
        
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        
        # Store reference to the tab frame
        self.plot_tabs['radiation_data_table'] = tab
    
    def update_radiation_data_table(self):
        """Update the radiation data table display"""
        if self.radiation_data is None or self.radiation_tree is None:
            return
    
        # Clear existing items
        for item in self.radiation_tree.get_children():
            self.radiation_tree.delete(item)
    
        # Get the radiation data
        radiation_df = self.radiation_data.copy()
    
        # Apply channel filtering
        if self.selected_channels:
            # Filter by selected channels
            valid_channels = []
            for display_channel in self.selected_channels:
                # Check if this is a renamed channel
                actual_channel = display_channel
                if hasattr(self, 'channel_names') and self.channel_names:
                    if display_channel in self.channel_names.values():
                        # Find the original name
                        for orig_name, new_name in self.channel_names.items():
                            if new_name == display_channel:
                                actual_channel = orig_name
                                break
            
                # Check if channel exists in radiation data
                if actual_channel in radiation_df['Channel'].astype(str).values:
                    valid_channels.append(actual_channel)
                elif display_channel in radiation_df['Channel'].astype(str).values:
                    valid_channels.append(display_channel)
        
            if valid_channels:
                radiation_df = radiation_df[radiation_df['Channel'].astype(str).isin(valid_channels)]
    
        # Check if we have data to display
        if radiation_df is None or len(radiation_df) == 0:
            # Show empty table with message
            columns = ['Info']
            self.radiation_tree['columns'] = columns
            self.radiation_tree['show'] = 'headings'
            for col in columns:
                self.radiation_tree.heading(col, text=col)
                self.radiation_tree.column(col, width=300, anchor=tk.CENTER)
                self.radiation_tree.insert('', 'end', values=['No radiation data available'])
            return
    
        # Set up columns with units
        columns = list(radiation_df.columns)
        self.radiation_tree['columns'] = columns
        self.radiation_tree['show'] = 'headings'
    
        for col in columns:
            # Get unit for this column
            unit = UNITS_MAP.get(col, "")
        
            # Create single-line header with unit in parentheses
            if unit:
                header_text = f"{col} ({unit})"
            else:
                header_text = col
            
            self.radiation_tree.heading(col, text=header_text)
        
            # Calculate optimal width and then MAXIMIZE it
            width = self.calculate_optimal_width(radiation_df, col, header_text, is_raw=False)
        
            # MAXIMIZE THE WIDTH - make it 80% wider for maximum readability if needed
            width = int(width * 1)
        
            # Set column-specific minimum widths for key columns
            if col == 'Channel':
                min_width = max(200, width)
            elif col in ['time_repeated', 'time']:
                min_width = max(150, width)
            elif col in ['absorbed_radiation', 'CO2_exchange_rate', 'Stomatal_conductance_corrected']:
                min_width = max(180, width)
            else:
                min_width = max(150, width)
            
            self.radiation_tree.column(col, width=width, anchor=tk.CENTER, minwidth=min_width)
    
        # Insert data with proper formatting
        for _, row in radiation_df.iterrows():
            values = []
            for val in row:
                if pd.isna(val):
                    values.append('')
                else:
                    # Format numbers nicely
                    if isinstance(val, float):
                        # Format floats to reasonable precision
                        if abs(val) < 0.0001:
                            values.append(f"{val:.2e}")
                        elif abs(val) < 0.01:
                            values.append(f"{val:.6f}")
                        elif abs(val) < 1:
                            values.append(f"{val:.4f}")
                        elif abs(val) < 100:
                            values.append(f"{val:.2f}")
                        else:
                            values.append(f"{val:.1f}")
                    else:
                        values.append(str(val))
            self.radiation_tree.insert('', 'end', values=values)
    
        # Add info about number of rows
        total_rows = len(radiation_df)
        info_row = [''] * len(columns)
        info_row[0] = f"Total rows: {total_rows}"
        self.radiation_tree.insert('', 'end', values=info_row, tags=('info',))
        self.radiation_tree.tag_configure('info', background='lightyellow')
         
    def download_radiation_data(self):
        """Download the radiation data table with two export options (TSV format)"""
        if self.radiation_data is None:
            messagebox.showwarning("Warning", "No radiation data available to download.")
            return

        # Create custom dialog for export options
        dialog = tk.Toplevel(self.root)
        dialog.title("Export Options")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Set icon for the dialog
        self._set_dialog_icon(dialog)
    
        # Center the dialog
        dialog_width = 400
        dialog_height = 250
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        x = (screen_width - dialog_width) // 2
        y = (screen_height - dialog_height) // 2
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
    
        # Add instruction label
        ttk.Label(dialog, text="Select export format:", font=("Arial", 12, "bold")).pack(pady=20)
    
        # Add description
        desc_frame = ttk.Frame(dialog)
        desc_frame.pack(pady=10)
        ttk.Label(desc_frame, text="Choose the format for your exported data:\n\n- RE-IMPORT (28 columns compatible)\n- ALL columns (all calculated columns)", font=("Arial", 10)).pack()
    
        # Variable to store the result
        result = [None]
    
        def on_reimport():
            result[0] = "reimport"
            dialog.destroy()
    
        def on_all_columns():
            result[0] = "all_columns"
            dialog.destroy()
    
        # Button frame
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
    
        ttk.Button(button_frame, text="RE-IMPORT", command=on_reimport, 
                   width=20).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="ALL COLUMNS", command=on_all_columns, 
                   width=20).pack(side=tk.LEFT, padx=10)
    
        # Wait for dialog to close
        self.root.wait_window(dialog)
    
        # Check if user closed the dialog without choosing
        if result[0] is None:
            return
    
        export_mode = result[0]

        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[
                ("TSV/TXT files", "*.txt"),
                ("All files", "*.*")
            ],
            title="Save Radiation Data As"
        )

        if not filename:
            return

        # Ensure .txt extension
        if not filename.endswith('.txt'):
            filename = filename.rsplit('.', 1)[0] + '.txt' if '.' in filename else filename + '.txt'

        try:
            # Get current date and time
            now = datetime.now()
            date_str = now.strftime("%Y-%m-%d")
            time_str = now.strftime("%H:%M:%S")
        
            # Get the original filename if available
            original_filename = self.file_path_var.get() if hasattr(self, 'file_path_var') else "Unknown"
            modified_text = f"Modified file: {os.path.basename(original_filename)}"
        
            # Define the original 28 columns (exactly as from raw data)
            original_28_columns = [
                "Channel", "CO2_M", "CO2_R", "H2O_M", "H2O_R", "O3_M", "O3_R",
                "CO2_exch_rate", "H2O_evol_rate", "overall_O3_uptake_rate", "air_temp", 
                "transp_ind_leaf_temp_depr", "leaf_temp", "satur_hum_at_leaf_temp", 
                "leaf_air_hum_grad", "stomatal_cond", "stomatal_O3_uptake_rate", "O3_cumul_dose",
                "stom_change_rate", "normalised_stom_cond", "relstom_condchange_rate", "leaf_area", 
                "air_flow_rate", "absorbed_radiation", "boundary_layer_cond", "cutic_conductance", 
                "temp_calibr", "time"
            ]
        
            # Create a copy for export
            df_export = self.radiation_data.copy()
        
            # Apply column selection based on user choice
            if export_mode == "reimport":
                # Only keep columns that actually exist in the dataframe
                export_columns = [col for col in original_28_columns if col in df_export.columns]
                df_export = df_export[export_columns].copy()
                mode_text = "Re-import compatible (original 28 columns only)"
            else:  # all_columns
                mode_text = "Complete data (all calculated columns)"
        
            # Sort by continuous time (handle day rollover)
            if 'time' in df_export.columns:
                time_seconds = []
                for time_val in df_export['time']:
                    parts = str(time_val).split(':')
                    if len(parts) >= 3:
                        secs = float(parts[0]) * 3600 + float(parts[1]) * 60 + float(parts[2])
                    else:
                        secs = 0.0
                    time_seconds.append(secs)
            
                # Detect day rollover
                day_increment = 0
                full_seconds = []
                prev_sec = time_seconds[0] if time_seconds else 0
                max_day_increment = 0
            
                for sec in time_seconds:
                    if sec < prev_sec:
                        day_increment += 1
                        if day_increment > max_day_increment:
                            max_day_increment = day_increment
                    prev_sec = sec
                    full_seconds.append(sec + day_increment * 24 * 3600)
            
                # Add temporary sort column
                df_export['_sort_time'] = full_seconds
                df_export = df_export.sort_values(['_sort_time', 'Channel'])
                df_export = df_export.drop(columns=['_sort_time'])
        
            # Convert numeric values to string with comma as decimal separator
            for col in df_export.columns:
                if df_export[col].dtype in ['float64', 'float32', 'int64', 'int32']:
                    df_export[col] = df_export[col].apply(
                        lambda x: str(x).replace('.', ',') if pd.notna(x) else ''
                    )
        
            # Create units row
            units_row = []
            for col in df_export.columns:
                unit = UNITS_MAP.get(col, "")
                # Replace special characters for plain text
                unit = unit.replace('₂', '2').replace('₃', '3').replace('ₒ', 'o')
                unit = unit.replace('µ', 'u').replace('μ', 'u')
                units_row.append(unit)
        
            # Add empty column at the beginning
            headers = df_export.columns.tolist()
            headers.insert(0, "")
            units_row.insert(0, "")
        
            # Add empty column to dataframe
            df_export.insert(0, "", "")
            
            # Write as TSV (tab-separated) with .txt extension
            with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f, delimiter='\t')
            
                # Row 1: Metadata
                writer.writerow(["", f"Date: {date_str}", f"Time: {time_str}", modified_text, f"Mode: {mode_text}"])
            
                # Row 2: Empty row
                writer.writerow([])
            
                # Row 3: Headers
                writer.writerow(headers)
            
                # Row 4: Units row
                writer.writerow(units_row)
            
                # Write data starting from row 5
                for _, row in df_export.iterrows():
                    writer.writerow(row)
        
            total_radiation_intervals = 0
            if self.radiation_settings:
                for channel, intervals in self.radiation_settings.items():
                    total_radiation_intervals += len(intervals)
        
            # Count columns exported
            num_cols_exported = len(df_export.columns) - 1
        
            # Check if multi-day experiment was detected
            multi_day_note = ""
            if 'max_day_increment' in locals() and max_day_increment > 0:
                multi_day_note = f"\nMulti-day experiment detected ({max_day_increment + 1} days) - Sorted correctly by continuous time"
        
            messagebox.showinfo("Success", 
                f"Radiation data saved to {filename}\n\n"
                f"Export mode: {mode_text}\n"
                f"Format: TSV (tab-separated) with .txt extension\n"
                f"Decimal separator: Comma (,)\n"
                f"Sorting: By continuous time (handles day rollover){multi_day_note}\n"
                f"Total rows: {len(self.radiation_data)}\n"
                f"Channels: {len(self.radiation_data['Channel'].unique())}\n"
                f"Columns exported: {num_cols_exported}\n"
                f"Radiation intervals: {total_radiation_intervals}\n\n"
                f"File structure:\n"
                f"  Row 1: Metadata (Date, Time, Source file, Export mode)\n"
                f"  Row 2: Empty\n"
                f"  Row 3: Column headers (starting with empty column A)\n"
                f"  Row 4: Units\n"
                f"  Row 5+: Data with empty first column")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {str(e)}")
            traceback.print_exc()
    
    def create_WUEi_comparison_plot(df, palette='tab10'):
        """Create combined boxplot and violin plot for WUEi"""
        # Close any existing figure to prevent accumulation
        plt.close('all')
        
        if df is None or len(df) == 0 or 'WUEi' not in df.columns:
            fig, ax = plt.subplots(figsize=(12, 6))
            ax.text(0.5, 0.5, 'No data available', ha='center', va='center', fontsize=12)
            ax.set_axis_off()
            return fig
        
        # Create figure with two subplots
        fig, axes = plt.subplots(1, 2, figsize=(12, 6), sharey=True)
        
        # Prepare data for plots
        channels = df['Channel'].unique()
        wuei_data = []
        channel_labels = []
        
        for channel in channels:
            channel_data = df[df['Channel'] == channel]
            wuei_values = channel_data['WUEi'].dropna()
            if len(wuei_values) > 0:
                wuei_data.append(wuei_values)
                channel_labels.append(str(channel))
        
        if not wuei_data:
            ax.text(0.5, 0.5, 'No WUEi data available', ha='center', va='center', fontsize=12)
            return fig
        
        # Get colors
        colors = []
        if palette in plt.colormaps():
            cmap = plt.get_cmap(palette)
            colors = cmap(np.linspace(0, 1, len(wuei_data)))
        else:
            colors = plt.cm.tab10(np.linspace(0, 1, len(wuei_data)))
        
        positions = range(1, len(wuei_data) + 1)
        
        # --- LEFT SUBPLOT: Boxplot ---
        ax1 = axes[0]
        box_plot = ax1.boxplot(wuei_data, positions=positions, patch_artist=True)
        
        # Apply colors to boxplot
        for patch, color in zip(box_plot['boxes'], colors):
            patch.set_facecolor(color)
            patch.set_alpha(0.6)
        
        # Add individual data points to boxplot
        for i, data in enumerate(wuei_data):
            x = np.random.normal(i + 1, 0.04, size=len(data))
            ax1.scatter(x, data, alpha=0.5, color=colors[i], s=20, edgecolor='black', linewidth=0.5)
        
        ax1.set_title('WUEi Distribution - Boxplot', fontsize=14, fontweight='bold')
        ax1.set_xlabel('Channel', fontsize=12)
        ax1.set_ylabel('WUEi (µmol CO₂ / mol H₂O)', fontsize=12)
        ax1.set_xticks(positions)
        ax1.set_xticklabels(channel_labels, rotation=45, ha='right')
        ax1.grid(True, alpha=0.3, linestyle='--')
        
        # Calculate statistics for annotation
        median_values = [np.median(data) for data in wuei_data]
        for i, median in enumerate(median_values):
            ax1.text(positions[i], median, f'{median:.1f}', 
                    ha='center', va='bottom', fontsize=9, fontweight='bold')
        
        # --- RIGHT SUBPLOT: Violin plot ---
        ax2 = axes[1]
        
        # Create violin plot
        violin_parts = ax2.violinplot(wuei_data, positions=positions, showmeans=True, showmedians=True)
        
        # Customize violin plot colors
        for i, pc in enumerate(violin_parts['bodies']):
            pc.set_facecolor(colors[i])
            pc.set_alpha(0.6)
            pc.set_edgecolor('black')
        
        # Customize median and mean lines
        violin_parts['cmedians'].set_color('red')
        violin_parts['cmedians'].set_linewidth(2)
        violin_parts['cmeans'].set_color('green')
        violin_parts['cmeans'].set_linewidth(2)
        
        # Add individual data points to violin plot
        for i, data in enumerate(wuei_data):
            x = np.random.normal(i + 1, 0.1, size=len(data))
            ax2.scatter(x, data, alpha=0.4, color=colors[i], s=15, edgecolor='black', linewidth=0.5)
        
        ax2.set_title('WUEi Distribution - Violin Plot', fontsize=14, fontweight='bold')
        ax2.set_xlabel('Channel', fontsize=12)
        ax2.set_xticks(positions)
        ax2.set_xticklabels(channel_labels, rotation=45, ha='right')
        ax2.grid(True, alpha=0.3, linestyle='--')
        
        # Add quartile annotations to violin plot
        for i, data in enumerate(wuei_data):
            q1 = np.percentile(data, 25)
            q3 = np.percentile(data, 75)
            ax2.text(positions[i], q1, f'Q1: {q1:.1f}', 
                    ha='center', va='top', fontsize=8, alpha=0.8)
            ax2.text(positions[i], q3, f'Q3: {q3:.1f}', 
                    ha='center', va='bottom', fontsize=8, alpha=0.8)
        
        # Add overall statistics to the figure
        fig.suptitle('WUEi Analysis: Boxplot vs Violin Plot Comparison', 
                    fontsize=16, fontweight='bold', y=0.98)
        
        plt.tight_layout()
        return fig    
    def setup_notes_tab(self):
        tab = ttk.Frame(self.main_notebook)
        self.main_notebook.add(tab, text="Notes/Biblio")
        
        notes_text = """Plant Physiology Gas Exchange Formulas:

Comprehensive explanation of calculations for photosynthesis, transpiration, stomatal conductance, ozone uptake, and derived physiological metrics.

1. Common Metrics Calculations

Saturating Air Humidity (Tetens Equation)

  saturating_air_humidity = 6.04 × exp(17.569 × air_temp ÷ (241.9 + air_temp)) [mmol/mol]
  
  Maximum water vapor air can hold at given temperature (saturation vapor pressure). Based on Tetens equation with empirical constants optimized for plant physiology applications.

Key Metric: Used to calculate humidity deficits and vapor pressure gradients driving transpiration.


Relative Air Humidity

  relative_air_humidity = (H2O_M ÷ saturating_air_humidity) × 100 [%]

Relative humidity as percentage comparing measured water vapor to maximum possible.

CO₂ Exchange Rate (Net Photosynthesis, A)

  CO2_exchange_rate = air_flow_rate × (CO2_R - CO2_M) ÷ leaf_area [µmol/m²s]

  Net CO₂ assimilation rate (photosynthesis minus respiration). Positive values indicate net CO₂ uptake.

Physiological Significance: Primary measure of photosynthetic performance. Values typically range from -5 (respiration) to 30+ µmol/m²s for C3 plants under optimal conditions.


Transpiration Rate (E)

  Transpiration_H2O_evol_rate = air_flow_rate × (H2O_M - H2O_R) ÷ leaf_area [mmol/m²s]

Water loss through transpiration. Typical values: 1-10 mmol/m²s depending on VPD and stomatal opening.

Vapor Pressure Deficit (VPD)

  VPD = saturating_air_humidity × (1 - relative_chamber_air_humidity) ÷ 10 [kPa]

Driving force for transpiration. Critical range for plants: 0.8-2.0 kPa optimal for most crops; >2.5 kPa causes stomatal closure.

Growth Impact: VPD affects both transpiration rate and stomatal behavior, influencing water use efficiency.


Leaf Temperature Estimation

  leaf_temp_c = air_temp - (0.773 × Transpiration_H2O_evol_rate) + (absorbed_radiation ÷ 0.00137) [ºC]

Empirical estimation accounting for transpirational cooling (0.773 coefficient) and radiative heating.

Note: Leaf temperature affects all gas exchange processes: +10°C typically doubles enzymatic reaction rates until optimal temperature is exceeded.


2. Conductance Metrics

Leaf-Air Humidity Gradient

  leaf_air_hum_grad = satur_hum_at_leaf_temp_c - H2O_M [mmol/mol]

Driving force for transpiration based on leaf temperature rather than air temperature.


Overall Conductance to Water Vapor (g_tw)

  overall_conductance = 1000 × Transpiration_H2O_evol_rate ÷ leaf_air_hum_grad [mmol/m²s]

Total conductance for water vapor from leaf interior to bulk air. Combines stomatal and boundary layer conductances in series.


Mass Flow Correction for Overall Resistance

  massflow_correction_for_overall_resistance = 1 + (mean_vapor_pressure ÷ (1013 - mean_vapor_pressure))

Corrects for water vapor mass flow effects on other gas movements.


Corrected Overall Conductance

  corr_d_overall_conductance = (1000 × Transpiration_H2O_evol_rate × (273 + leaf_temp_c)) ÷ (leaf_air_hum_grad × 446 × 273 × massflow_correction_for_overall_resistance) [cm/s]

Temperature- and mass flow-corrected total conductance normalized to standard conditions.


Corrected Stomatal Conductance (g_sw)

  corr_stomatal_conductance = corr_leaf_conductance - cutic_conductance [cm/s]

Stomatal component isolated by subtracting cuticular conductance. Typical range: 0-2 cm/s (0-200 mmol/m²s).

Stomatal Regulation: Primary control point for plant water loss and CO₂ uptake. Affected by light, CO₂, humidity, ABA, and circadian rhythms.


3. Stomatal Metrics

Temperature-Corrected Stomatal Conductance

  Stomatal_conductance_corrected = (corr_stomatal_conductance × 446 × 273) ÷ (273 + leaf_temp_c) [mmol/m²s]

Standardized to 25°C for cross-comparison. Conversion factor 446 represents molar volume at STP.


Normalized Stomatal Conductance

  normalised_corr_stomatal_conductance = Stomatal_conductance_corrected ÷ Stomatal_conductance_corrected[1]

Relative change from initial measurement. Useful for tracking dynamic responses to treatments.


Stomatal Change Rate

  corr_stomatal_change_rate[i] = (Stomatal_conductance_corrected[i+1] - Stomatal_conductance_corrected[i]) ÷ 2 [mmol/m²s]

Rate of stomatal movement. Fast responses (minutes) vs slow responses (hours) indicate different regulatory mechanisms.

4. Ozone Metrics

Overall O₃ Uptake Rate (F_O3)

  overall_O3_uptake_rate_c = air_flow_rate × (O3_M - O3_R) ÷ leaf_area [nmol/m²s]

Total ozone flux to leaf surface. Critical levels: >1 nmol/m²s causes visible injury in sensitive species.


Corrected Conductance for O₃ (g_O3)

  corr_conductance_for_O3_c = (b_layer_conductance × Stomatal_conductance_corrected × 0.62) ÷ (b_layer_conductance + Stomatal_conductance_corrected) [mmol/m²s]

Ozone-specific conductance (0.62×g_sw) due to lower diffusivity in water than CO₂/H₂O.


Stomatal O₃ Uptake Rate

  corr_stomatal_O3_uptake_rate = 0.001 × corr_conductance_for_O3_c × O3_M [nmol O₃/m²s]

Ozone flux through stomata. More biologically relevant than external concentration for risk assessment.


Cumulative O₃ Dose (AF_st x t)

  corr_cumul_O3_dose[i] = corr_cumul_O3_dose[i-1] + (corr_stomatal_O3_uptake_rate[i] × 0.001 × 960) [µmol/m²]

Integrated ozone exposure over time. Factor 960 converts seconds to 16-min intervals (standard for flux-based standards).

Risk Assessment: Critical levels: 4-12 mmol/m² POD₁ (Phytotoxic Ozone Dose above 1 nmol/m²s threshold) for crop yield loss.


5. CO₂ Metrics

Stomatal Resistance to CO₂ (r_s,CO2)

stom_resist_to_CO2 = 1.62 × 2.24 × (273 + leaf_temp_c) ÷ (273 × corr_stomatal_conductance) [m²s/mmol]

CO₂-specific resistance (1.62× higher than H₂O) due to different diffusion coefficients in air and water.


Boundary Layer Resistance to CO₂

  b_layer_resist_to_CO2 = 1.62 × 2.24 × (273 + air_temp) ÷ (273 × boundary_layer_cond) [m²s/mmol]

Air-side resistance component.


Intercellular CO₂ Concentration (C_i)

  intercell_CO2_conc_in_gas_phase = massflow_corr_d_ca - corr_ted_CO2_grad [µmol CO₂/mol air]

Estimated CO₂ concentration at carboxylation sites. Critical parameter linking stomatal and biochemical limitations.

Photosynthetic Limitation Analysis: C_i/C_a ratio typically 0.7-0.8 for well-watered C3 plants. Lower values indicate stomatal limitation; higher values indicate biochemical limitation.


Mesophyll Conductance (g_m)

  mesophyll_conductance_for_CO2 = ΔCO2_exchange_rate × 100 ÷ Δintercell_CO2_conc [mmol/m²s]

Conductance from intercellular air space to chloroplasts. Often limiting factor: 0.1-0.4 mol/m²s for many species.

Methodological Note: Calculated from response curve slope. Values capped at ±600 mmol/m²s to exclude physiologically unrealistic extremes.


6. Water Use Efficiency & Compensation Point

CO₂ Compensation Point (Γ)

  CO2_comp_point = (mesophyll_conductance_for_CO2 × intercell_CO2_conc_in_gas_phase × 0.001 - CO2_exchange_rate) ÷ (mesophyll_conductance_for_CO2 × 0.001) [µmol CO₂/mol air]

CO₂ concentration where photosynthesis equals respiration. Typically 30-50 µmol/mol for C3 plants at 25°C.


Instantaneous Water Use Efficiency (WUE)

  WUE = CO2_exchange_rate ÷ Transpiration_H2O_evol_rate [µmol CO₂/mmol H₂O]

Carbon gain per water loss. Range: 1-10 µmol CO₂/mmol H₂O depending on species and conditions.


Intrinsic Water Use Efficiency (WUEᵢ)

  WUEi = CO2_exchange_rate × 1000 ÷ Stomatal_conductance_corrected [µmol CO₂/mol H₂O]

Photosynthesis per unit stomatal opening. Less sensitive to VPD than WUE.

Adaptation Significance: High WUE_i indicates conservative water use strategy; low WUE_i indicates acquisitive strategy.


7. Correction Ratios

Stomatal Conductance Correction Ratio

  corr_uncorr_stom_cond = Stomatal_conductance_corrected ÷ stomatal_cond

Quantifies temperature correction magnitude. Typically 0.8-1.2 depending on leaf-air temperature difference.


CO₂ Concentration Correction Ratios

  corr_uncorr_ca = massflow_corr_d_ca ÷ CO2_M

  corr_uncorr_ci = intercell_CO2_conc_in_gas_phase ÷ uncorr_ci

Mass flow and temperature correction impacts on CO₂ concentrations.


Mesophyll Conductance Correction Ratio

  corr_uncorr_gm_prima = mesophyll_conductance_for_CO2 ÷ uncorr_gm_prima

Effect of corrections on estimated mesophyll conductance.


Compensation Point Correction Ratio

  corr_uncorr_gamma = CO2_comp_point ÷ uncorr_gamma

Impact of corrections on Γ estimation.


8. Physiological Interpretation Framework

Integrated Analysis Approach

1. Limitation Analysis: Compare C_i/C_a ratio, g_s, and g_m to identify primary photosynthetic limitations.
2. Stress Detection: Monitor changes in WUE_i, stomatal dynamics, and Γ for early stress indicators.
3. Ozone Risk Assessment: Use cumulative stomatal flux rather than atmospheric concentration.
4. Temperature Acclimation: Analyze temperature-corrected vs. uncorrected values.

9. Data Flow Diagram

Calculation Pipeline

The calculations follow a specific sequence where outputs from earlier functions become inputs to later ones:

Key Physical Principles: These calculations are based on: (1) Fick's law of diffusion (gas fluxes proportional to concentration gradients), (2) Ideal gas law (temperature corrections), (3) Series resistance model (boundary layer + stomatal resistances), and (4) Mass flow corrections (evaporation-driven gas movement).

References

Key Methodology References:
  
1. Long, S.P., & Bernacchi, C.J. (2003). Gas exchange measurements, what can they tell us about the underlying limitations to photosynthesis? Procedures and sources of error. Journal of Experimental Botany, 54(392), 2393-2401.

2. von Caemmerer, S., & Farquhar, G.D. (1981). Some relationships between the biochemistry of photosynthesis and the gas exchange of leaves. Planta, 153(4), 376-387.

3. Leuning, R. (1995). A critical appraisal of a combined stomatal-photosynthesis model for C3 plants. Plant, Cell & Environment, 18(4), 339-355.

4. Monteith, J.L., & Unsworth, M.H. (2013). Principles of Environmental Physics (4th ed.). Academic Press.

Ozone Uptake Specific:
  
5. Emberson, L.D., Büker, P., Ashmore, M.R., et al. (2000). Modelling stomatal ozone flux across Europe. Environmental Pollution, 109(3), 403-413.

6. Paoletti, E., & Manning, W.J. (2007). Toward a biologically significant and usable standard for ozone that will also protect plants. Environmental Pollution, 150(1), 85-95.

Conductance Calculations:
  
7. Jones, H.G. (2014). Plants and Microclimate: A Quantitative Approach to Environmental Plant Physiology (3rd ed.). Cambridge University Press.

8. Ball, J.T., Woodrow, I.E., & Berry, J.A. (1987). A model predicting stomatal conductance and its contribution to the control of photosynthesis under different environmental conditions. In: Progress in Photosynthesis Research (pp. 221-224). Springer.

Calibration & Corrections:
  
9. Flexas, J., Díaz-Espejo, A., Berry, J.A., et al. (2007). Analysis of leakage in IRGA's leaf chambers of open gas exchange systems: quantification and its effects in photosynthesis parameterization. Journal of Experimental Botany, 58(6), 1533-1543.

Temperature Corrections:
  
10. Bernacchi, C.J., Pimentel, C., & Long, S.P. (2003). In vivo temperature response functions of parameters required to model RuBP-limited photosynthesis. Plant, Cell & Environment, 26(9), 1419-1430.

11. Buckley, T.N., & Diaz-Espejo, A. (2015). Reporting estimates of maximum potential electron transport rate. New Phytologist, 205(1), 14-17.

Mass Flow Corrections:
  
12. Jarman, P.D. (1974). The diffusion of carbon dioxide and water vapour through stomata. Journal of Experimental Botany, 25(5), 927-936.

13. Farquhar, G.D., & Sharkey, T.D. (1982). Stomatal conductance and photosynthesis. Annual Review of Plant Physiology, 33(1), 317-345.

Essential References

1. Farquhar, G.D., von Caemmerer, S., & Berry, J.A. (1980). A biochemical model of photosynthetic CO₂ assimilation in leaves of C3 species. Planta, 149(1), 78-90.

2. von Caemmerer, S. (2000). Biochemical Models of Leaf Photosynthesis. CSIRO Publishing.

3. Jones, H.G. (1992). Plants and Microclimate (2nd ed.). Cambridge University Press.

4. Monteith, J.L., & Unsworth, M.H. (1990). Principles of Environmental Physics (2nd ed.). Edward Arnold.

5. Emberson, L.D., et al. (2000). Modelling stomatal ozone flux across Europe. Environmental Pollution, 109(3), 403-413.

Gas Exchange Analysis Formulas | Based on established plant physiology principles and standard gas exchange measurement protocols"""
        
        text_widget = scrolledtext.ScrolledText(tab, wrap=tk.WORD, width=100, height=30)
        text_widget.insert(tk.END, notes_text)
        text_widget.configure(state='disabled')
        text_widget.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), padx=10, pady=10)
        
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select TXT File",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if filename:
            self.file_path_var.set(filename)
            
    def process_file(self):
        filename = self.file_path_var.get()
        if not filename:
            messagebox.showerror("Error", "Please select a file first")
            return

        # Disable the process button and show loading message
        self.process_btn.config(state='disabled', text="Processing...")
        self.root.update()  # Force UI update

        # Show a prominent loading message in the status area
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="PROCESSING FILE, PLEASE WAIT UNTIL COMPLETED...", 
                foreground="black",
                background="yellow",
                font=("Arial", 9, "bold")
            )
            self.slider_info_label.grid()  # Ensure it's visible
            self.root.update()  # Force UI update again

        try:
            # Clear previous data
            self.clear_old_figures()
    
            # Load new data
            self.raw_data = pd.read_csv(filename, delimiter='\t', skiprows=2, 
                                        header=None, decimal=',')
    
            # Show processing frames
            if hasattr(self, 'processing_frame'):
                self.processing_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
            if hasattr(self, 'time_range_frame'):
                self.time_range_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
    
            # Process the data
            self.process_raw_data()
    
            # Reset entries to defaults
            self.reset_entries()
    
            # Update entry ranges based on new data
            self.update_entry_ranges()
    
            # Update time information display
            self.update_shift_time_info()
    
            # Initialize Apply Changes button as disabled
            if hasattr(self, 'apply_changes_btn'):
                self.apply_changes_btn.config(state='disabled')
    
            # Store initial entry values for comparison
            self.initial_entry_values = {
                'remove_start': self.remove_start_var.get(),
                'remove_end': self.remove_end_var.get(),
                'time0_shift': self.time0_shift_var.get()
            }
    
            # Update info label to normal state - DO THIS ONLY ONCE
            if hasattr(self, 'slider_info_label'):
                self.slider_info_label.config(
                    text="Enter values, then click 'Apply Changes'", 
                    foreground="blue",
                    background="",  # Reset to default background
                    font=("Arial", 8)  # Reset to normal font
                )
    
            # Initialize Apply Changes button as disabled
            if hasattr(self, 'apply_changes_btn'):
                self.apply_changes_btn.config(state='disabled')
    
            # Store initial slider values for comparison
            self.initial_slider_values = {
                'remove_start': self.remove_start_var.get(),
                'remove_end': self.remove_end_var.get(),
                'time0_shift': self.time0_shift_var.get()
            }
    
            # Show success message
            messagebox.showinfo("Success", "File processed successfully!")
    
            # Re-enable process button
            self.process_btn.config(state='normal', text="Process File")
    
        except Exception as e:
            # Re-enable process button on error
            self.process_btn.config(state='normal', text="Process File")
            if hasattr(self, 'slider_info_label'):
                self.slider_info_label.config(
                    text="Error processing file", 
                    foreground="red"
                )
    
            messagebox.showerror("Error", f"Failed to process file: {str(e)}")
            traceback.print_exc()
            
    def process_raw_data(self):
        headers = self.raw_data.iloc[0, 1:].tolist()
        data = self.raw_data.iloc[2:, 1:].reset_index(drop=True)
    
        column_names = [
            "Channel", "CO2_M", "CO2_R", "H2O_M", "H2O_R", "O3_M", "O3_R",
            "CO2_exch_rate", "H2O_evol_rate", "overall_O3_uptake_rate", "air_temp", 
            "transp_ind_leaf_temp_depr", "leaf_temp", "satur_hum_at_leaf_temp", 
            "leaf_air_hum_grad", "stomatal_cond", "stomatal_O3_uptake_rate", "O3_cumul_dose",
            "stom_change_rate", "normalised_stom_cond", "relstom_condchange_rate", "leaf_area", 
            "air_flow_rate", "absorbed_radiation", "boundary_layer_cond", "cutic_conductance", 
            "temp_calibr", "time"
        ]
    
        data.columns = column_names
    
        # Update progress message
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="Processing data, this process will take a bit...", 
                foreground="white",
                background="red",
                font=("Arial", 12, "bold")
            )
            self.root.update()
    
        for col in column_names[:-1]:
            data[col] = pd.to_numeric(data[col].astype(str).str.replace(',', '.'), errors='coerce')
    
        self.base_data = data
    
        # Calculate all data
        self.calculate_original_data()
    
        # Initialize modified data
        self.modified_calculated_data = self.original_calculated_data.copy()
    
        # Update displays
        self.update_raw_data_display()
        self.update_calculated_data_display()
        self.create_channel_params_ui()
        self.create_channel_selector()
    
        # Update slider ranges
        self.update_slider_ranges()
    
        self.update_time_range_info()
        self.update_all_plots()
    
        # Initialize entry values
        self.initial_entry_values = {
            'remove_start': "0",
            'remove_end': "0", 
            'time0_shift': "0"
        }
    
        # Enable Apply Changes button
        if hasattr(self, 'apply_changes_btn'):
            self.apply_changes_btn.config(state='disabled')
    
        # ============ SHOW ALL HIDDEN FRAMES ============
        # Show processing controls frame
        if hasattr(self, 'processing_frame'):
            self.processing_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 15), padx=10)
    
        # Show time range info frame
        if hasattr(self, 'time_range_frame'):
            self.time_range_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 15), padx=10)
    
        # Show visualization settings frame
        if hasattr(self, 'color_frame'):
            self.color_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 15), padx=10)
    
        # Show coefficient tables
        if hasattr(self, 'coef_frame'):
            self.coef_frame.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(10, 5), padx=10)
    
        if hasattr(self, 'color_coef_frame'):
            self.color_coef_frame.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=(5, 0), padx=10)
    
    def debug_data_flow(self):
        """Debug method to see what's happening with data"""
        print("\n=== DEBUG DATA FLOW ===")
        print(f"Original calculated data exists: {self.original_calculated_data is not None}")
        print(f"Modified calculated data exists: {self.modified_calculated_data is not None}")
        
        if self.original_calculated_data is not None:
            print(f"Original data shape: {self.original_calculated_data.shape}")
            channels = self.original_calculated_data['Channel'].unique()
            print(f"Original channels: {channels}")
        
        if self.modified_calculated_data is not None:
            print(f"Modified data shape: {self.modified_calculated_data.shape}")
            channels = self.modified_calculated_data['Channel'].unique()
            print(f"Modified channels: {channels}")
        
        print(f"Parameters applied: {self.parameters_applied}")
        print("=====================\n")        
            
    def calculate_original_data(self):
        df = self.base_data.copy()
    
        df = calculate_common_metrics(df)
        df = calculate_conductance_metrics(df)
        df = calculate_stomatal_metrics(df)
        df = calculate_ozone_metrics(df)
        df = calculate_co2_metrics(df)
        df = calculate_additional_metrics(df)
        
        # Convert time parameters to integers
        try:
            time0_shift = int(self.time0_shift_var.get())
        except ValueError:
            time0_shift = 0
            
        try:
            remove_start = int(self.remove_start_var.get())
        except ValueError:
            remove_start = 0
            
        try:
            remove_end = int(self.remove_end_var.get())
        except ValueError:
            remove_end = 0
        
        df = process_time_columns(df, 
                                  time0_shift,
                                  remove_start,
                                  remove_end)
        
        self.original_calculated_data = df
        # DO NOT update modified_calculated_data here
        
    def create_modified_data(self):
        if self.original_calculated_data is None:
            return None
            
        # Make a deep copy to avoid modifying the original
        modified_df = self.original_calculated_data.copy()
        
        # First, convert Channel column to string type to handle mixed types
        modified_df['Channel'] = modified_df['Channel'].astype(str)
        
        channels = modified_df['Channel'].unique()
        
        # Apply all parameter changes first (before renaming)
        for channel in channels:
            if self.parameters_applied['leaf_area'] and channel in self.leaf_area_params:
                modified_df.loc[modified_df['Channel'] == channel, 'leaf_area'] = self.leaf_area_params[channel]
            
            if self.parameters_applied['air_flow'] and channel in self.air_flow_params:
                modified_df.loc[modified_df['Channel'] == channel, 'air_flow_rate'] = self.air_flow_params[channel]
            
            if self.parameters_applied['radiation'] and channel in self.radiation_params:
                modified_df.loc[modified_df['Channel'] == channel, 'absorbed_radiation'] = self.radiation_params[channel]
            
            if self.parameters_applied['boundary'] and channel in self.boundary_params:
                modified_df.loc[modified_df['Channel'] == channel, 'boundary_layer_cond'] = self.boundary_params[channel]
            
            if self.parameters_applied['cutic'] and channel in self.cutic_params:
                modified_df.loc[modified_df['Channel'] == channel, 'cutic_conductance'] = self.cutic_params[channel]
        
        # Apply channel name changes LAST
        if self.parameters_applied['channel_names']:
            for old_name, new_name in self.channel_names.items():
                # Check if old_name exists in the dataframe
                if old_name in modified_df['Channel'].values:
                    modified_df.loc[modified_df['Channel'] == old_name, 'Channel'] = new_name
        
        # Recalculate all metrics with the updated data
        modified_df = calculate_common_metrics(modified_df)
        modified_df = calculate_conductance_metrics(modified_df)
        modified_df = calculate_stomatal_metrics(modified_df)
        modified_df = calculate_ozone_metrics(modified_df)
        modified_df = calculate_co2_metrics(modified_df)
        modified_df = calculate_additional_metrics(modified_df)
        
        # Process time columns
        modified_df = process_time_columns(modified_df,
                                          int(self.time0_shift_var.get() or 0),
                                          int(self.remove_start_var.get() or 0),
                                          int(self.remove_end_var.get() or 0))
        
        return modified_df
    
    def recalculate_modified_data(self):
        """Recalculate modified data with all current parameters"""
        if self.original_calculated_data is None:
            return
        
        # Get current time parameters
        try:
            time0_shift = int(self.time0_shift_var.get())
        except ValueError:
            time0_shift = 0
            
        try:
            remove_start = int(self.remove_start_var.get())
        except ValueError:
            remove_start = 0
            
        try:
            remove_end = int(self.remove_end_var.get())
        except ValueError:
            remove_end = 0
        
        # Create a fresh copy of ORIGINAL base data (not original_calculated_data)
        # We need to start from base_data and apply ALL calculations including time parameters
        if self.base_data is None:
            print(f"No base data available")
            return
        
        df = self.base_data.copy()
        
        # Step 1: Apply ALL basic calculations first
        print(f"Step 1 - Applying basic calculations...")
        df = calculate_common_metrics(df)
        df = calculate_conductance_metrics(df)
        df = calculate_stomatal_metrics(df)
        df = calculate_ozone_metrics(df)
        df = calculate_co2_metrics(df)
        df = calculate_additional_metrics(df)
        
        # Step 2: Apply ALL parameter changes BEFORE time processing
        print(f"Step 2 - Applying parameter changes...")
        channels = df['Channel'].unique()
        
        # Apply parameter changes
        for channel in channels:
            channel_str = str(channel)
            
            # Apply leaf area
            if self.parameters_applied['leaf_area'] and channel_str in self.leaf_area_params:
                new_value = self.leaf_area_params[channel_str]
                df.loc[df['Channel'].astype(str) == channel_str, 'leaf_area'] = new_value
                print(f"Applied leaf area {new_value} to channel {channel_str}")
            
            # Apply air flow
            if self.parameters_applied['air_flow'] and channel_str in self.air_flow_params:
                new_value = self.air_flow_params[channel_str]
                df.loc[df['Channel'].astype(str) == channel_str, 'air_flow_rate'] = new_value
            
            # Apply radiation
            if self.parameters_applied['radiation'] and channel_str in self.radiation_params:
                new_value = self.radiation_params[channel_str]
                df.loc[df['Channel'].astype(str) == channel_str, 'absorbed_radiation'] = new_value
                print(f"Applied radiation {new_value} to channel {channel_str}")
            
            # Apply boundary
            if self.parameters_applied['boundary'] and channel_str in self.boundary_params:
                new_value = self.boundary_params[channel_str]
                df.loc[df['Channel'].astype(str) == channel_str, 'boundary_layer_cond'] = new_value
            
            # Apply cutic
            if self.parameters_applied['cutic'] and channel_str in self.cutic_params:
                new_value = self.cutic_params[channel_str]
                df.loc[df['Channel'].astype(str) == channel_str, 'cutic_conductance'] = new_value
        
        # Step 3: Recalculate ALL metrics with updated parameters
        print(f"Step 3 - Recalculating metrics with updated parameters...")
        df = calculate_common_metrics(df)
        df = calculate_conductance_metrics(df)
        df = calculate_stomatal_metrics(df)
        df = calculate_ozone_metrics(df)
        df = calculate_co2_metrics(df)
        df = calculate_additional_metrics(df)
        
        # Step 4: Apply time parameters (remove_start, remove_end, time0_shift)
        print(f"Step 4 - Applying time parameters...")
        df = process_time_columns(df, time0_shift, remove_start, remove_end)
        
        # Step 5: Apply channel name changes LAST (after all calculations)
        if self.parameters_applied['channel_names']:
            print(f"Step 5 - Applying channel name changes...")
            for old_name, new_name in self.channel_names.items():
                old_str = str(old_name).strip()
                new_str = str(new_name).strip()
                # Convert Channel column to string for comparison
                df['Channel'] = df['Channel'].astype(str)
                mask = df['Channel'].str.strip() == old_str
                if mask.any():
                    df.loc[mask, 'Channel'] = new_str
                    print(f"Renamed {mask.sum()} rows from {old_str} to {new_str}")
        
        # Store the updated data
        self.modified_calculated_data = df
        print(f"Modified data updated. Shape: {df.shape}")
        print(f"Modified data time range: {df['time_repeated'].min():.1f} to {df['time_repeated'].max():.1f}")
        print("=== END ===\n")
    
    def create_channel_selector(self):
        print("\n=== CREATE CHANNEL SELECTOR CALLED ===")  # Debug
        if self.original_calculated_data is None:
            print("No original data, returning")
            return
    
        # Clear the listbox
        self.channel_listbox.delete(0, tk.END)

        # Get the data that will be displayed (with any renaming applied)
        df_for_display = self.get_data_for_calculated_table()
        if df_for_display is None:
            df_for_display = self.original_calculated_data

        # Get unique channels from the display data
        display_channels = df_for_display['Channel'].unique()
        print(f"Display channels: {display_channels}")  # Debug

        # Clear and rebuild selected_channels based on display channel names
        self.selected_channels = [str(channel) for channel in display_channels]
        print(f"Setting selected_channels to: {self.selected_channels}")  # Debug

        # Add each channel to the listbox and select it
        for i, channel in enumerate(display_channels):
            channel_str = str(channel)
            self.channel_listbox.insert(tk.END, channel_str)
            self.channel_listbox.selection_set(i)  # Select by index

        # Update plots immediately with all channels selected
        self.update_all_plots()
        print("=== CREATE CHANNEL SELECTOR COMPLETED ===\n")  # Debug
    
    def select_all_channels(self):
        """Select all channels in the listbox"""
        # Store old selection for comparison
        old_selection = self.selected_channels.copy() if hasattr(self, 'selected_channels') else []
    
        # Select all in listbox
        self.channel_listbox.selection_set(0, tk.END)
        self.selected_channels = [self.channel_listbox.get(i) for i in range(self.channel_listbox.size())]
    
        # Determine which channels were added
        added_channels = [ch for ch in self.selected_channels if ch not in old_selection]
    
        # Show status message BEFORE updates
        if added_channels and hasattr(self, 'slider_info_label'):
            channel_text = ', '.join(added_channels[:3])  # Show first 3 channels
            if len(added_channels) > 3:
                channel_text += f" and {len(added_channels) - 3} more"
        
            self.slider_info_label.config(
                text=f"✓ SELECT ALL: Showing {len(added_channels)} channel{'s' if len(added_channels) > 1 else ''} ({channel_text})",
                foreground="white",
                background="green",
                font=("Arial", 9, "bold")
            )
            self.root.update()  # Force immediate display
    
        # Update all displays with current selection
        self.update_calculated_data_display()
        self.update_all_plots()
    
        # Also update darkness and radiation tables if they exist
        if hasattr(self, 'darkness_data') and self.darkness_data is not None:
            self.update_darkness_data_table()
        if hasattr(self, 'radiation_data') and self.radiation_data is not None:
            self.update_radiation_data_table()
    
        # Clear message after updates
        if added_channels:
            self.root.after(2500, lambda: 
                self.slider_info_label.config(
                    text="Enter values, then click 'Apply Changes'", 
                    foreground="blue",
                    background="",
                    font=("Arial", 8)
                ) if hasattr(self, 'slider_info_label') else None
            )

    def deselect_all_channels(self):
        """Deselect all channels in the listbox"""
        print("\n=== DESELECT ALL CALLED ===")  # Debug
    
        # Store old selection for comparison
        old_selection = self.selected_channels.copy() if hasattr(self, 'selected_channels') else []
        print(f"Old selection: {old_selection}")
    
        # Deselect all in listbox
        self.channel_listbox.selection_clear(0, tk.END)
    
        # IMPORTANT: Explicitly set selected_channels to empty list
        self.selected_channels = []
        print(f"New selection (should be empty): {self.selected_channels}")
    
        # Force listbox to update
        self.root.update_idletasks()
    
        # Determine which channels were removed
        removed_channels = old_selection
    
        # Show status message BEFORE updates
        if removed_channels and hasattr(self, 'slider_info_label'):
            channel_text = ', '.join(removed_channels[:3])
            if len(removed_channels) > 3:
                channel_text += f" and {len(removed_channels) - 3} more"
        
            self.slider_info_label.config(
                text=f"✗ DESELECT ALL: Hiding {len(removed_channels)} channel{'s' if len(removed_channels) > 1 else ''} ({channel_text})",
                foreground="white",
                background="red",
                font=("Arial", 9, "bold")
            )
            self.root.update()
    
        # Update all displays with current selection
        print("Calling update_calculated_data_display()")
        self.update_calculated_data_display()
    
        print("Calling update_all_plots()")
        self.update_all_plots()
    
        # Also update darkness and radiation tables if they exist
        if hasattr(self, 'darkness_data') and self.darkness_data is not None:
            print("Calling update_darkness_data_table()")
            self.update_darkness_data_table()
        if hasattr(self, 'radiation_data') and self.radiation_data is not None:
            print("Calling update_radiation_data_table()")
            self.update_radiation_data_table()
    
        # Clear message after updates
        if removed_channels:
            self.root.after(2500, lambda: 
                self.slider_info_label.config(
                    text="Enter values, then click 'Apply Changes'", 
                    foreground="blue",
                    background="",
                    font=("Arial", 8)
                ) if hasattr(self, 'slider_info_label') else None
            )
    
        print("=== DESELECT ALL COMPLETED ===\n")
    
    def get_filtered_data(self, df):
        """
        Universal method to filter data based on selected channels
        Works for both original and renamed channels
        """
        print(f"\n--- get_filtered_data called ---")
        print(f"df shape: {df.shape if df is not None else 'None'}")
        print(f"self.selected_channels: {self.selected_channels}")
    
        if df is None or len(df) == 0:
            print("df is None or empty, returning as-is")
            return df

        # If no channels selected, return an empty dataframe with same structure
        if not self.selected_channels:
            print("NO CHANNELS SELECTED - returning empty dataframe")
            empty_df = pd.DataFrame(columns=df.columns)
            print(f"Empty df shape: {empty_df.shape}")
            return empty_df

        # Ensure Channel column is string
        df['Channel'] = df['Channel'].astype(str).str.strip()

        # Get unique channels in dataframe
        df_channels = set(df['Channel'].unique())
        print(f"Available channels in df: {df_channels}")

        # Track valid channels to filter
        valid_channels = []

        for display_channel in self.selected_channels:
            display_str = str(display_channel).strip()
            print(f"Checking channel: '{display_str}'")
    
            # Direct match
            if display_str in df_channels:
                print(f"  Direct match found: {display_str}")
                valid_channels.append(display_str)
                continue
    
            # Check if this is a renamed channel
            if hasattr(self, 'channel_names') and self.channel_names:
                # If display_channel is an original name that was renamed
                if display_str in self.channel_names:
                    new_name = self.channel_names[display_str].strip()
                    print(f"  {display_str} was renamed to {new_name}")
                    if new_name in df_channels:
                        print(f"  Found renamed channel: {new_name}")
                        valid_channels.append(new_name)
                        continue
            
                # If display_channel is a new name (renamed channel)
                for orig_name, new_name in self.channel_names.items():
                    new_name_clean = new_name.strip()
                    if display_str == new_name_clean and orig_name in df_channels:
                        print(f"  Found original name {orig_name} for display name {display_str}")
                        valid_channels.append(orig_name)
                        break

        print(f"Valid channels found: {valid_channels}")

        if not valid_channels:
            print("No valid channels found, returning empty dataframe")
            return pd.DataFrame(columns=df.columns)

        # Apply filter
        filtered_df = df[df['Channel'].isin(valid_channels)]
        print(f"Filtered df shape: {filtered_df.shape}")
        print(f"--- get_filtered_data completed ---\n")
        return filtered_df
    
    def get_data_for_plots(self):
        """Get filtered data for plotting - includes time trimming and channel selection"""
        print("\n--- get_data_for_plots called ---")
    
        # Start with the appropriate base data
        if self.modified_calculated_data is not None:
            df = self.modified_calculated_data.copy()
            print(f"Using modified_calculated_data, shape: {df.shape}")
        else:
            df = self.original_calculated_data.copy() if self.original_calculated_data is not None else None
            print(f"Using original_calculated_data, shape: {df.shape if df is not None else 'None'}")
    
        if df is None or len(df) == 0:
            print(f"No data available, returning None")
            return None
    
        # Ensure Channel column is string for consistent comparison
        df['Channel'] = df['Channel'].astype(str)
    
        # Apply channel filtering
        filtered_df = self.get_filtered_data(df)
    
        print(f"filtered_df shape: {filtered_df.shape if filtered_df is not None else 'None'}")
        print(f"--- get_data_for_plots completed ---\n")
    
        return filtered_df
    
    def update_channel_names(self):
        if self.original_calculated_data is None:
            messagebox.showwarning("Warning", "No data available.")
            return

        # Show processing message
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="PROCESSING CHANNEL RENAMES, PLEASE WAIT...", 
                foreground="black",
                background="yellow",
                font=("Arial", 9, "bold")
            )
            self.root.update()

        channels = self.original_calculated_data['Channel'].unique()
    
        # Store old darkness settings before clearing
        old_darkness_settings = {}
        if hasattr(self, 'darkness_settings') and self.darkness_settings:
            old_darkness_settings = self.darkness_settings.copy()

        # Clear existing channel names dictionary
        self.channel_names = {}

        # Collect new channel names
        name_changes = {}  # Store mapping from old to new names
        for channel in channels:
            widget_name = f"channel_name_{channel}"
            if hasattr(self, widget_name):
                widget = getattr(self, widget_name)
                new_name = widget.get().strip()
                channel_str = str(channel)
            
                # Store the mapping from old to new name
                if new_name and new_name != channel_str:
                    name_changes[channel_str] = new_name

        # Only proceed if there are actual changes
        if name_changes:
            # Update the channel_names dictionary
            self.channel_names = name_changes.copy()
            self.parameters_applied['channel_names'] = True

            # Update darkness settings with new channel names while preserving ALL data
            if old_darkness_settings:
                new_darkness_settings = {}
            
                # First, process each old setting
                for old_name, settings in old_darkness_settings.items():
                    # Try to find where this setting should go
                    target_name = None
                
                    # Case 1: The old_name was renamed to something else
                    if old_name in name_changes:
                        target_name = name_changes[old_name]
                    # Case 2: The old_name is actually a new display name that exists
                    elif old_name in [str(c) for c in channels]:
                        target_name = old_name
                    # Case 3: The old_name was a display name that corresponds to a renamed channel
                    else:
                        for orig, new in name_changes.items():
                            if old_name == new:
                                target_name = new
                                break
                
                    if target_name:
                        # Preserve the full settings including type, times, and manual_times
                        new_darkness_settings[target_name] = settings.copy()
            
                self.darkness_settings = new_darkness_settings

            # Recalculate modified data with new names
            self.recalculate_modified_data()

            # Update displays
            self.update_calculated_data_display()
            self.create_channel_selector()

            # Recreate darkness inputs AFTER renaming - this preserves manual entries
            # while keeping auto-detected values in the background for plotting
            self.create_darkness_inputs()
        
            # Also recreate radiation inputs
            self.create_radiation_inputs()

            # Update all plots - this will use the FULL darkness_settings 
            # (auto + manual combined) for the yellow/black rectangles
            self.update_all_plots()

            # Clear processing message and show success
            if hasattr(self, 'slider_info_label'):
                self.slider_info_label.config(
                    text=f"Channel names updated: {', '.join(self.channel_names.values())}", 
                    foreground="green",
                    background="",
                    font=("Arial", 8)
                )
                # Reset after delay
                self.root.after(3000, lambda: 
                    self.slider_info_label.config(
                        text="Enter values, then click 'Apply Changes'", 
                            foreground="blue"
                    ) if hasattr(self, 'slider_info_label') else None
                )
    
            messagebox.showinfo("Success", f"Channel names updated: {', '.join(self.channel_names.values())}")
        else:
            # Clear processing message if no changes
            if hasattr(self, 'slider_info_label'):
                self.slider_info_label.config(
                    text="No channel name changes detected", 
                    foreground="orange",
                    background="",
                    font=("Arial", 8)
                )
                # Reset after delay
                self.root.after(2000, lambda: 
                    self.slider_info_label.config(
                        text="Enter values, then click 'Apply Changes'", 
                        foreground="blue"
                    ) if hasattr(self, 'slider_info_label') else None
                )
        
            messagebox.showinfo("Info", "No channel name changes detected.")
    
    def update_leaf_area(self):
        if self.original_calculated_data is None:
            return
          
        # Show processing message
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="PROCESSING LEAF AREA UPDATES, PLEASE WAIT...", 
                foreground="black",
                background="yellow",
                font=("Arial", 9, "bold")
            )
            self.root.update()
        
        channels = self.original_calculated_data['Channel'].unique()
        for channel in channels:
            widget_name = f"leaf_area_{channel}"
            if hasattr(self, widget_name):
                widget = getattr(self, widget_name)
                try:
                    # Store as string key
                    channel_str = str(channel)
                    new_value = float(widget.get())
                    self.leaf_area_params[channel_str] = new_value
                    print(f"Stored leaf area {new_value} for channel {channel_str}")
                except ValueError:
                    print(f"Invalid value for channel {channel}")
                    pass
        
        self.parameters_applied['leaf_area'] = True
        self.recalculate_modified_data()
        
        # Update the parameter table text boxes to show new values
        if self.modified_calculated_data is not None:
            for channel in channels:
                channel_str = str(channel)
                widget_name = f"leaf_area_{channel}"
                
                if hasattr(self, widget_name):
                    widget = getattr(self, widget_name)
                    # Get the updated value from the modified data
                    if channel_str in self.modified_calculated_data['Channel'].astype(str).values:
                        channel_data = self.modified_calculated_data[
                            self.modified_calculated_data['Channel'].astype(str) == channel_str
                        ]
                        if not channel_data.empty:
                            new_value = channel_data['leaf_area'].iloc[0]
                            # Update the text entry with the new calculated value
                            widget.delete(0, tk.END)
                            widget.insert(0, str(new_value))
                            print(f"Updated widget for channel {channel_str} to {new_value}")
        
        self.update_calculated_data_display()
        self.update_all_plots()
        
        # Clear processing message and show success
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="✓ Leaf area updated successfully!", 
                foreground="green",
                background="",
                font=("Arial", 8)
            )
            # Reset after delay
            self.root.after(2000, lambda: 
                self.slider_info_label.config(
                    text="Enter values, then click 'Apply Changes'", 
                    foreground="blue"
                ) if hasattr(self, 'slider_info_label') else None
            )
        
        messagebox.showinfo("Success", "Leaf area updated!")
        
        # Debug
        self.debug_data_flow()
    
    def update_air_flow(self):
        if self.original_calculated_data is None:
            return
        
        # Show processing message
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="PROCESSING AIR FLOW RATE UPDATES, PLEASE WAIT...", 
                foreground="black",
                background="yellow",
                font=("Arial", 9, "bold")
            )
            self.root.update()
            
        channels = self.original_calculated_data['Channel'].unique()
        for channel in channels:
            widget_name = f"air_flow_rate_{channel}"
            if hasattr(self, widget_name):
                widget = getattr(self, widget_name)
                try:
                    # Store as string key
                    channel_str = str(channel)
                    new_value = float(widget.get())
                    self.air_flow_params[channel_str] = new_value
                    print(f"Stored air flow {new_value} for channel {channel_str}")
                except ValueError:
                    print(f"Invalid value for channel {channel}")
                    pass
        
        self.parameters_applied['air_flow'] = True
        self.recalculate_modified_data()
        
        # Update the parameter table text boxes to show new values
        if self.modified_calculated_data is not None:
            for channel in channels:
                channel_str = str(channel)
                widget_name = f"air_flow_rate_{channel}"
                
                if hasattr(self, widget_name):
                    widget = getattr(self, widget_name)
                    # Get the updated value from the modified data
                    if channel_str in self.modified_calculated_data['Channel'].astype(str).values:
                        channel_data = self.modified_calculated_data[
                            self.modified_calculated_data['Channel'].astype(str) == channel_str
                        ]
                        if not channel_data.empty:
                            new_value = channel_data['air_flow_rate'].iloc[0]
                            # Update the text entry with the new calculated value
                            widget.delete(0, tk.END)
                            widget.insert(0, str(new_value))
                            print(f"Updated widget for channel {channel_str} to {new_value}")
        
        self.update_calculated_data_display()
        self.update_all_plots()
        
        # Clear processing message and show success
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="✓ Air flow rate updated successfully!", 
                foreground="green",
                background="",
                font=("Arial", 8)
            )
            # Reset after delay
            self.root.after(2000, lambda: 
                self.slider_info_label.config(
                    text="Enter values, then click 'Apply Changes'", 
                    foreground="blue"
                ) if hasattr(self, 'slider_info_label') else None
            )
    
        messagebox.showinfo("Success", "Air flow rate updated!")
        
        # Debug
        self.debug_data_flow()
    
    def update_radiation(self):
        if self.original_calculated_data is None:
            return
        
        
        # Show processing message
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="PROCESSING RADIATION UPDATES, PLEASE WAIT... ", 
                foreground="black",
                background="yellow",
                font=("Arial", 9, "bold")
                )
            self.root.update()    
        channels = self.original_calculated_data['Channel'].unique()
        
        # FIRST: Store the new values
        for channel in channels:
            widget_name = f"absorbed_radiation_{channel}"
            if hasattr(self, widget_name):
                widget = getattr(self, widget_name)
                try:
                    # CRITICAL: Store as STRING key to match recalculate_modified_data()
                    channel_str = str(channel)
                    new_value = float(widget.get())
                    self.radiation_params[channel_str] = new_value
                    print(f"Stored radiation {new_value} for channel {channel_str}")
                except ValueError:
                    print(f"Invalid value for channel {channel}")
                    pass
        
        self.parameters_applied['radiation'] = True
        self.recalculate_modified_data()  # Use the new method
        
        # Update the parameter table text boxes to show new values
        if self.modified_calculated_data is not None:
            for channel in channels:
                channel_str = str(channel)
                widget_name = f"absorbed_radiation_{channel}"
                
                if hasattr(self, widget_name):
                    widget = getattr(self, widget_name)
                    # Get the updated value from the modified data
                    if channel_str in self.modified_calculated_data['Channel'].astype(str).values:
                        channel_data = self.modified_calculated_data[
                            self.modified_calculated_data['Channel'].astype(str) == channel_str
                        ]
                        if not channel_data.empty:
                            new_value = channel_data['absorbed_radiation'].iloc[0]
                            # Update the text entry with the new calculated value
                            widget.delete(0, tk.END)
                            widget.insert(0, str(new_value))
                            print(f"Updated widget for channel {channel_str} to {new_value}")
        
        self.update_calculated_data_display()
        self.update_all_plots()
        
        # Clear processing message and show success
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="✓ Absorbed radiation updated successfully!", 
                foreground="green",
                background="",
                font=("Arial", 8)
            )
            # Reset after delay
            self.root.after(2000, lambda: 
                self.slider_info_label.config(
                    text="Enter values, then click 'Apply Changes'", 
                    foreground="blue"
                ) if hasattr(self, 'slider_info_label') else None
            )
        messagebox.showinfo("Success", "Absorbed radiation updated!")
        
        # Debug
        self.debug_data_flow()
    
    def update_boundary(self):
        if self.original_calculated_data is None:
            return
        
        # Show processing message
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="PROCESSING BOUNDARY LAYER UPDATES, PLEASE WAIT...", 
                foreground="black",
                background="yellow",
                font=("Arial", 9, "bold")
            )
            self.root.update()
        
        channels = self.original_calculated_data['Channel'].unique()
        for channel in channels:
            widget_name = f"boundary_layer_cond_{channel}"
            if hasattr(self, widget_name):
                widget = getattr(self, widget_name)
                try:
                    # Store as string key
                    channel_str = str(channel)
                    new_value = float(widget.get())
                    self.boundary_params[channel_str] = new_value
                    print(f"Stored boundary cond {new_value} for channel {channel_str}")
                except ValueError:
                    print(f"Invalid value for channel {channel}")
                    pass
        
        self.parameters_applied['boundary'] = True
        self.recalculate_modified_data()
        
        # Update the parameter table text boxes to show new values
        if self.modified_calculated_data is not None:
            for channel in channels:
                channel_str = str(channel)
                widget_name = f"boundary_layer_cond_{channel}"
                
                if hasattr(self, widget_name):
                    widget = getattr(self, widget_name)
                    # Get the updated value from the modified data
                    if channel_str in self.modified_calculated_data['Channel'].astype(str).values:
                        channel_data = self.modified_calculated_data[
                            self.modified_calculated_data['Channel'].astype(str) == channel_str
                        ]
                        if not channel_data.empty:
                            new_value = channel_data['boundary_layer_cond'].iloc[0]
                            # Update the text entry with the new calculated value
                            widget.delete(0, tk.END)
                            widget.insert(0, str(new_value))
                            print(f"Updated widget for channel {channel_str} to {new_value}")
        
        self.update_calculated_data_display()
        self.update_all_plots()
        
        # Clear processing message and show success
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="✓ Boundary layer conductance updated successfully!", 
                foreground="green",
                background="",
                font=("Arial", 8)
            )
            # Reset after delay
            self.root.after(2000, lambda: 
                self.slider_info_label.config(
                    text="Enter values, then click 'Apply Changes'", 
                    foreground="blue"
                ) if hasattr(self, 'slider_info_label') else None
            )
        
        messagebox.showinfo("Success", "Boundary layer conductance updated!")
        
        # Debug
        self.debug_data_flow()
    
    def update_cutic(self):
        if self.original_calculated_data is None:
            return
        
        # Show processing message
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="PROCESSING CUTIC CONDUCTANCE UPDATES, PLEASE WAIT...", 
                foreground="black",
                background="yellow",
                font=("Arial", 9, "bold")
            )
            self.root.update()
        
        
        channels = self.original_calculated_data['Channel'].unique()
        for channel in channels:
            widget_name = f"cutic_conductance_{channel}"
            if hasattr(self, widget_name):
                widget = getattr(self, widget_name)
                try:
                    # Store as string key
                    channel_str = str(channel)
                    new_value = float(widget.get())
                    self.cutic_params[channel_str] = new_value
                    print(f"Stored cutic cond {new_value} for channel {channel_str}")
                except ValueError:
                    print(f"Invalid value for channel {channel}")
                    pass
        
        self.parameters_applied['cutic'] = True
        self.recalculate_modified_data()
        
        # Update the parameter table text boxes to show new values
        if self.modified_calculated_data is not None:
            for channel in channels:
                channel_str = str(channel)
                widget_name = f"cutic_conductance_{channel}"
                
                if hasattr(self, widget_name):
                    widget = getattr(self, widget_name)
                    # Get the updated value from the modified data
                    if channel_str in self.modified_calculated_data['Channel'].astype(str).values:
                        channel_data = self.modified_calculated_data[
                            self.modified_calculated_data['Channel'].astype(str) == channel_str
                        ]
                        if not channel_data.empty:
                            new_value = channel_data['cutic_conductance'].iloc[0]
                            # Update the text entry with the new calculated value
                            widget.delete(0, tk.END)
                            widget.insert(0, str(new_value))
                            print(f"Updated widget for channel {channel_str} to {new_value}")
        
        self.update_calculated_data_display()
        self.update_all_plots()
        
        # Clear processing message and show success
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="✓ Cuticular conductance updated successfully!", 
                foreground="green",
                background="",
                font=("Arial", 8)
            )
            # Reset after delay
            self.root.after(2000, lambda: 
                self.slider_info_label.config(
                    text="Enter values, then click 'Apply Changes'", 
                    foreground="blue"
                ) if hasattr(self, 'slider_info_label') else None
            )
        
        messagebox.showinfo("Success", "Cutic conductance updated!")
        
        # Debug
        self.debug_data_flow()
    
    def on_palette_change(self, event=None):
        """Handle palette selection changes"""
        self.update_all_plots()
    
    def apply_darkness(self):
        if self.modified_calculated_data is None:
            messagebox.showwarning("Warning", "No data available.")
            return

        # Show processing message
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="PROCESSING DARKNESS SETTINGS, PLEASE WAIT...", 
                foreground="yellow",
                background="black",
                font=("Arial", 9, "bold")
            )
            self.root.update()

        # Get current data for darkness inputs
        df = self.get_data_for_plots()
        if df is None:
            messagebox.showwarning("Warning", "No data available for darkness settings.")
            return

        channels = df['Channel'].unique()

        # Flag to track if any manual entries were made
        manual_entries = False

        # Create a NEW dictionary that will store ALL settings (auto + manual)
        new_darkness_settings = {}

        # FIRST: For each channel, get AUTO-detected values from current data
        # These will be used for plots
        for channel in channels:
            channel_str = str(channel)
    
            # Auto-detect zero absorbance from current data
            channel_data = df[df['Channel'].astype(str) == channel_str]
            if not channel_data.empty and 'absorbed_radiation' in channel_data.columns:
                zero_times = channel_data[channel_data['absorbed_radiation'] == 0]['time_repeated'].tolist()
            
                # Store as AUTO type (for plots only)
                new_darkness_settings[channel_str] = {
                    'type': 'auto',
                    'times': zero_times.copy()  # Make a copy to avoid reference issues
                }
                print(f"Auto-detected times for channel {channel_str}: {zero_times}")

        # SECOND: Process NEW manual entries from input boxes
        for channel in channels:
            channel_str = str(channel)

            if channel_str in self.darkness_widgets:
                entry_widget = self.darkness_widgets[channel_str]
                times_str = entry_widget.get()

                if times_str and times_str.strip():  # Non-empty input
                    # Process manual entries
                    manual_entries = True
                    try:
                        # Parse the times from input - ALWAYS AS INTEGERS
                        new_manual_times = []
                        for t_str in times_str.split(','):
                            t_str = t_str.strip()
                            if t_str:
                                # Convert to integer directly
                                new_manual_times.append(int(t_str))
                    
                        # Get the auto-detected times for this channel
                        if channel_str in new_darkness_settings:
                            auto_times = new_darkness_settings[channel_str].get('times', [])
                        else:
                            auto_times = []
                    
                        # Find matching time points in current data for manual entries
                        valid_manual_times = []
                        channel_data = df[df['Channel'].astype(str) == channel_str]
                    
                        for t in new_manual_times:
                            if not channel_data.empty:
                                time_diffs = abs(channel_data['time_repeated'] - t)
                                if not time_diffs.empty and time_diffs.min() < 0.01:
                                    closest_idx = time_diffs.idxmin()
                                    valid_manual_times.append(channel_data.loc[closest_idx, 'time_repeated'])
                    
                        # Combine auto and manual times (avoid duplicates)
                        all_times = list(set(auto_times + valid_manual_times))
                        all_times.sort()
                    
                        # Update the settings for this channel - MARK AS MANUAL TYPE
                        new_darkness_settings[channel_str] = {
                            'type': 'manual',
                            'times': all_times,  # Combined auto + manual for plotting
                            'manual_times': valid_manual_times  # Just manual for input boxes
                        }
                    
                        print(f"Channel {channel_str}: Auto times: {auto_times}")
                        print(f"Channel {channel_str}: Manual times: {valid_manual_times}")
                        print(f"Channel {channel_str}: Combined times: {all_times}")
                    
                    except ValueError:
                        if hasattr(self, 'slider_info_label'):
                            self.slider_info_label.config(
                                text="Error in darkness settings - please use integers only", 
                                foreground="red"
                            )
                        messagebox.showwarning("Warning", 
                                             f"Invalid time format for channel {channel_str}. Please use integers only (e.g., 5, 10, 15)")
                        return

        # Update the main darkness_settings with the new combined dictionary
        self.darkness_settings = new_darkness_settings
        self.darkness_applied = True

        # CRITICAL: Update darkness data and plots
        self.update_darkness_data()

        # Update input boxes to show ONLY the manual entries (as integers)
        self.update_darkness_inputs_from_settings()

        # Clear processing message and show success
        if hasattr(self, 'slider_info_label'):
            manual_count = sum(1 for v in self.darkness_settings.values() 
                              if v.get('type') == 'manual' and v.get('manual_times'))
            auto_count = sum(1 for v in self.darkness_settings.values() 
                            if v.get('type') == 'auto' and v.get('times'))
    
            if manual_entries:
                self.slider_info_label.config(
                    text=f"✓ Added manual entries to {manual_count} channels, keeping auto-detected", 
                    foreground="green",
                    background=""
                )
            elif auto_count > 0:
                self.slider_info_label.config(
                    text=f"✓ Auto-detected {auto_count} channels with zero absorbance", 
                    foreground="green",
                    background=""
                )
            else:
                self.slider_info_label.config(
                    text="No darkness times detected", 
                    foreground="orange"
                )
            self.root.after(3000, lambda: 
                self.slider_info_label.config(
                    text="Enter values, then click 'Apply Changes'", 
                    foreground="blue",
                    background=""
                ) if hasattr(self, 'slider_info_label') else None
            )

        total_times = sum(len(v['times']) for v in self.darkness_settings.values())
        if manual_entries:
            messagebox.showinfo("Success", 
                f"Manual entries added to {manual_count} channels\n"
                f"Total darkness periods across all channels: {total_times}")
        elif auto_count > 0:
            messagebox.showinfo("Success", 
                f"Auto-detected {auto_count} channels with zero absorbance\n"
                f"Total darkness periods: {total_times}")
        else:
            messagebox.showinfo("Info", "No darkness times detected")

    def update_darkness_inputs_from_settings(self):
        """Update input boxes to show ONLY manual entries as integers"""
        if not hasattr(self, 'darkness_widgets'):
            return
    
        for channel_str, entry_widget in self.darkness_widgets.items():
            # Clear current content
            entry_widget.delete(0, tk.END)
        
            # Check if there's a MANUAL setting for this channel
            if hasattr(self, 'darkness_settings') and self.darkness_settings:
                if channel_str in self.darkness_settings:
                    settings = self.darkness_settings[channel_str]
                    # Only show manual times if they exist and it's a manual type
                    if settings.get('type') == 'manual' and 'manual_times' in settings:
                        times = settings['manual_times']
                        if times:
                            # Format times as integers (no decimal places)
                            formatted_times = ', '.join(str(t) for t in times)
                            entry_widget.insert(0, formatted_times)
                    # If it's auto type, leave empty (don't show auto-detected values)
    
    def clear_darkness(self):
        self.darkness_settings = {}
        self.darkness_applied = False
        self.darkness_data = None
        
        # Clear input fields
        if hasattr(self, 'darkness_widgets'):
            for channel_str, entry_widget in self.darkness_widgets.items():
                entry_widget.delete(0, tk.END)
        
        # Refresh plots to show original data
        self.update_all_plots()
        messagebox.showinfo("Success", "Darkness settings cleared!")
    
    def apply_radiation(self):
        if self.modified_calculated_data is None:  # CHANGED: Check modified data
            messagebox.showwarning("Warning", "No data available.")
            return
    
        # Show processing message
        if hasattr(self, 'slider_info_label'):
            self.slider_info_label.config(
                text="PROCESSING RADIATION INTERVALS, PLEASE WAIT...", 
                foreground="black",
                background="pink",
                font=("Arial", 9, "bold")
            )
            self.root.update()

        self.radiation_settings = {}
    
        # Get current data for radiation inputs
        df = self.get_data_for_plots()
        if df is None:
            return
    
        channels = df['Channel'].unique()
    
        has_entries = False
        for channel in channels:
            channel_str = str(channel)
        
            # Use the dictionary instead of attribute
            if channel_str in self.radiation_widgets:
                entry = self.radiation_widgets[channel_str]
                intervals_str = entry.get()
                if intervals_str:
                    has_entries = True
                    intervals = []
                    for interval in intervals_str.split(';'):
                        interval = interval.strip()
                        if ':' in interval:
                            time_range, value = interval.split(':', 1)
                            if '-' in time_range:
                                start, end = time_range.split('-', 1)
                                try:
                                    # Parse start and end as floats
                                    start_val = float(start.strip())
                                    end_val = float(end.strip())
                                
                                    # FIX: Ensure we capture the full range including the endpoint
                                    # Add a small epsilon to ensure the endpoint is included
                                    # This is particularly important when intervals are adjacent
                                    intervals.append({
                                        'start': start_val,
                                        'end': end_val + 1e-9,  # Add tiny epsilon to include endpoint
                                        'value': float(value.strip())
                                    })
                                    print(f"Parsed interval: {start_val} to {end_val} = {float(value.strip())}")
                                except ValueError:
                                    print(f"Invalid number in interval: {interval}")
                                    pass
                    if intervals:
                        self.radiation_settings[channel_str] = intervals
                        print(f"Stored radiation intervals for channel {channel_str}: {intervals}")
    
        self.radiation_applied = True
    
        # ALWAYS update radiation data
        self.update_radiation_data()  
    
        # Clear processing message and show success
        if hasattr(self, 'slider_info_label'):
            if has_entries:
                total_intervals = sum(len(intervals) for intervals in self.radiation_settings.values())
                self.slider_info_label.config(
                    text=f"✓ Radiation intervals applied to {len(self.radiation_settings)} channels ({total_intervals} intervals)", 
                    foreground="green",
                    background="",
                    font=("Arial", 8)
                )
            else:
                self.slider_info_label.config(
                    text="Displaying original data without radiation intervals", 
                    foreground="orange",
                    background="",
                    font=("Arial", 8)
                )
            # Reset after delay
            self.root.after(3000, lambda: 
                self.slider_info_label.config(
                    text="Enter values, then click 'Apply Changes'", 
                    foreground="blue",
                ) if hasattr(self, 'slider_info_label') else None
            )
    
        if has_entries:
            messagebox.showinfo("Success", f"Radiation intervals applied to {len(self.radiation_settings)} channels!")
        else:
            messagebox.showinfo("Info", "Displaying original data without radiation intervals")
    
    def setup_radiation_plots_tab(self):
        """Setup a single tab for ALL radiation plots - EXACTLY LIKE DARKNESS TAB"""
        tab = ttk.Frame(self.radiation_notebook)  # CHANGED: Use radiation_notebook
        self.radiation_notebook.add(tab, text="Radiation Plots")  # CHANGED: Use radiation_notebook
        
        # Create a canvas with both vertical and horizontal scrollbars
        container = ttk.Frame(tab)
        container.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        
        canvas = tk.Canvas(container)
        
        v_scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        h_scrollbar = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        canvas.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.E, tk.W))
        
        # Store reference
        self.plot_tabs['radiation_plots'] = scrollable_frame  # CHANGED KEY: radiation_plots instead of radiation_environment
        
        # Configure grid weights
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
    
    def update_radiation_plots(self):
        """Update radiation plots with modified data"""
        if self.radiation_data is None or len(self.radiation_data) == 0:
            return
        
        # Get the current palette
        palette = self.palette_var.get()
        
        # Get filtered data for plots
        df = self.get_filtered_data(self.radiation_data)
        if df is None or len(df) == 0:
            return
        
        # Define radiation plot categories
        radiation_plot_categories = {
            'radiation_environment': [
                ("absorbed_radiation", "Absorbed Radiation (Intervals)", "Radiation"),
                ("leaf_temp_c", "Leaf Temperature (Intervals)", "Leaf Temperature"),
                ("VPD", "Vapor Pressure Deficit (Intervals)", "VPD")
            ],
            'radiation_fluxes': [
                ("CO2_exchange_rate", "CO₂ Exchange Rate (Intervals)", "CO₂ Exchange Rate"),
                ("Transpiration_H2O_evol_rate", "Transpiration Rate (Intervals)", "Transpiration Rate"),
                ("Stomatal_conductance_corrected", "Stomatal Conductance (Intervals)", "Conductance")
            ]
        }
        
        # Update each radiation tab
        for category, plots in radiation_plot_categories.items():
            if category in self.plot_tabs:
                self.update_radiation_plot_tab(category, plots, df, palette)
    
    def update_radiation_plot_tab(self, category, plot_configs, df, palette):
        """Update the radiation plot tab with proper empty handling"""
        tab_frame = self.plot_tabs[category]

        # Destroy old canvas widgets and close figures
        for widget in tab_frame.winfo_children():
            if hasattr(widget, 'destroy'):
                widget.destroy()
            elif hasattr(widget, 'get_tk_widget'):
                widget.get_tk_widget().destroy()
    
        # If no data or empty dataframe, show message
        if df is None or len(df) == 0:
            print(f"Radiation plot tab {category} - no data, showing empty message")
            fig, ax = plt.subplots(figsize=(12, 6))
            ax.text(0.5, 0.5, 'No data available', 
                    ha='center', va='center', fontsize=14, fontweight='bold')
            ax.set_axis_off()
        
            plot_container = ttk.Frame(tab_frame)
            plot_container.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=20)
        
            canvas = FigureCanvasTkAgg(fig, plot_container)
            canvas.draw()
            canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.W, tk.E))
        
            tab_frame.columnconfigure(0, weight=1)
            return
    
        row = 0
        for yvar, title, ylab in plot_configs:
            if yvar in df.columns:
                # Use standard channel plot for radiation (no yellow/black rectangles)
                fig = make_channel_plot(df, yvar, title, ylab, palette)
            
                # Set consistent figure size
                fig.set_size_inches(12, 6)
            
                # Create a container frame for each plot to ensure proper layout
                plot_container = ttk.Frame(tab_frame)
                plot_container.grid(row=row*3, column=0, sticky=(tk.W, tk.E), 
                                  padx=10, pady=(20, 0), columnspan=2)
            
                # Add title label
                ttk.Label(plot_container, text=title, font=("Arial", 11, "bold")).grid(
                    row=0, column=0, sticky=tk.W)
            
                # Create canvas for the plot
                canvas = FigureCanvasTkAgg(fig, plot_container)
                canvas.draw()
                canvas.get_tk_widget().grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
            
                # Add download button
                btn = ttk.Button(plot_container, text=f"Download {title}",
                          command=lambda f=fig, t=title: self.download_plot(f, t))
                btn.grid(row=2, column=0, sticky=tk.W, pady=(5, 20))
            
                # Configure plot container column to expand
                plot_container.columnconfigure(0, weight=1)
            
                row += 1
            else:
                print(f"Warning: {yvar} not found in radiation data, skipping plot")
    
        # Make sure the tab frame expands properly
        tab_frame.columnconfigure(0, weight=1)
    
    def clear_radiation(self):
        """Clear radiation settings and inputs"""
        self.radiation_settings = {}
        self.radiation_applied = False
        self.radiation_data = None
        
        # Clear input fields using dictionary
        if hasattr(self, 'radiation_widgets'):
            for channel_str, entry_widget in self.radiation_widgets.items():
                entry_widget.delete(0, tk.END)
        
        # Refresh plots to show original data
        self.update_all_plots()
        messagebox.showinfo("Success", "Radiation settings cleared!")
    
    def update_darkness_data(self):
        # Start from CURRENT modified data (not original)
        if self.modified_calculated_data is None:
            return
        
        df = self.modified_calculated_data.copy()
    
        print(f"\nupdate_darkness_data() starting")
        print(f"Starting from modified data shape: {df.shape}")
        print(f"Modified data channels: {df['Channel'].unique()}")
    
        # Apply darkness settings if they exist
        if self.darkness_applied and self.darkness_settings:
            print(f"Applying darkness settings: {self.darkness_settings}")
            # Apply darkness to the filtered data
            for channel, dark_info in self.darkness_settings.items():
                print(f"Processing channel {channel} with dark_info: {dark_info}")
            
                # Use ALL times (auto + manual combined) for setting absorbed_radiation to 0
                for time in dark_info['times']:  # Use 'times' which contains combined auto+manual
                    # Find points at or near the specified darkness time
                    mask = (df['Channel'].astype(str) == channel)
                    if mask.any():
                        # Find the closest time point
                        time_diffs = abs(df.loc[mask, 'time_repeated'] - time)
                        if not time_diffs.empty and time_diffs.min() < 0.01:
                            closest_idx = time_diffs.idxmin()
                            old_value = df.loc[closest_idx, 'absorbed_radiation']
                            df.loc[closest_idx, 'absorbed_radiation'] = 0
                            print(f"  Set time {time} to 0 (was {old_value})")
                    else:
                        print(f"Channel {channel} not found in data")
        
            # IMPORTANT: Recalculate metrics with darkness applied
            print(f"Recalculating metrics with darkness applied")
            df = calculate_common_metrics(df)
            df = calculate_conductance_metrics(df)
            df = calculate_stomatal_metrics(df)
            df = calculate_ozone_metrics(df)
            df = calculate_co2_metrics(df)
            df = calculate_additional_metrics(df)
        else:
            print(f"No darkness settings to apply")
    
        # Apply channel renaming to the darkness data for display
        if hasattr(self, 'channel_names') and self.channel_names:
            # Create a copy to avoid modifying original
            df_display = df.copy()
        
            # Convert Channel column to string type BEFORE renaming
            df_display['Channel'] = df_display['Channel'].astype(str)
        
            # Apply renaming
            for old_name, new_name in self.channel_names.items():
                old_str = str(old_name)
                new_str = str(new_name)
                mask = df_display['Channel'] == old_str
                if mask.any():
                    df_display.loc[mask, 'Channel'] = new_str
                    print(f"Renamed channel {old_str} to {new_str} in darkness data")
        
            # Store the display version
            self.darkness_data = df_display
        else:
            # Store original version
            self.darkness_data = df
    
        print(f"Darkness data shape: {self.darkness_data.shape}")
        print(f"Darkness data channels: {self.darkness_data['Channel'].unique()}")
    
        # Update displays
        self.update_darkness_data_table()
        palette = self.palette_var.get()
        self.update_darkness_plots(palette)

    def filter_data_to_display_range(self, df):
        """
        IMPORTANT: This function should NOT be used for darkness data
        because darkness data is already filtered through process_time_columns()
        """
        # Simply return the data as-is - filtering already done
        return df
    
    
    def update_radiation_data(self):
        # Start from CURRENT modified data
        if self.modified_calculated_data is None:
            return
        
        df = self.modified_calculated_data.copy()
    
        print(f"\nUpdate_radiation_data() starting")
        print(f"Starting from modified data shape: {df.shape}")
        print(f"Modified data channels: {df['Channel'].unique()}")
    
        # Apply radiation intervals if they exist
        if self.radiation_settings:
            print(f"Applying radiation settings: {self.radiation_settings}")
            for channel, intervals in self.radiation_settings.items():
                print(f"Processing channel {channel} with intervals: {intervals}")
                for interval in intervals:
                    # FIX: Use the exact interval values including the tiny epsilon for the endpoint
                    mask = (df['Channel'].astype(str) == channel) & \
                           (df['time_repeated'] >= interval['start']) & \
                           (df['time_repeated'] <= interval['end'])  # This now includes the endpoint
                
                    if mask.any():
                        # Get the actual time points being modified for verification
                        modified_times = df.loc[mask, 'time_repeated'].tolist()
                        old_values = df.loc[mask, 'absorbed_radiation'].tolist()
                        df.loc[mask, 'absorbed_radiation'] = interval['value']
                    
                        print(f"Set absorbed_radiation to {interval['value']} for channel {channel}")
                        print(f"  Time range: {interval['start']} to {interval['end']}")
                        print(f"  Modified {len(modified_times)} points")
                        print(f"  Time points modified: {modified_times[:5]}...")  # Show first 5
                        print(f"  Old values were: {old_values[:3]}...")  # Show first 3 values
                    else:
                        print(f"No data found for channel {channel} in interval {interval['start']}-{interval['end']}")
        
            # Recalculate with radiation intervals applied
            print(f"Recalculating metrics with radiation intervals")
            df = calculate_common_metrics(df)
            df = calculate_conductance_metrics(df)
            df = calculate_stomatal_metrics(df)
            df = calculate_ozone_metrics(df)
            df = calculate_co2_metrics(df)
            df = calculate_additional_metrics(df)
        
            # Apply time columns processing
            try:
                time0_shift = int(self.time0_shift_var.get())
            except ValueError:
                time0_shift = 0
            
            try:
                remove_start = int(self.remove_start_var.get())
            except ValueError:
                remove_start = 0
            
            try:
                remove_end = int(self.remove_end_var.get())
            except ValueError:
                remove_end = 0
        
            df = process_time_columns(df,
                                      time0_shift,
                                      remove_start,
                                      remove_end)
        else:
            print(f"No radiation settings to apply")
    
        self.radiation_data = df
        print(f"Radiation data shape: {self.radiation_data.shape}")
        print(f"Radiation data channels: {self.radiation_data['Channel'].unique()}")
    
        self.update_radiation_plots()
        self.update_radiation_data_table()
    
    def update_raw_data_display(self):
        if self.base_data is None:
            return
            
        for item in self.raw_tree.get_children():
            self.raw_tree.delete(item)
        
        # Define units for raw data columns
        raw_units = {
            "Channel": "",
            "CO2_M": "µmol/mol",
            "CO2_R": "µmol/mol",
            "H2O_M": "mmol/mol",
            "H2O_R": "mmol/mol",
            "O3_M": "nmol/mol",
            "O3_R": "nmol/mol",
            "CO2_exch_rate": "µmol/m²s",
            "H2O_evol_rate": "mmol/m²s",
            "overall_O3_uptake_rate": "nmol/m²s",
            "air_temp": "°C",
            "transp_ind_leaf_temp_depr": "°C",
            "leaf_temp": "°C",
            "satur_hum_at_leaf_temp": "mmol/mol",
            "leaf_air_hum_grad": "mmol/mol",
            "stomatal_cond": "mmol/m²s",
            "stomatal_O3_uptake_rate": "nmol/m²s",
            "O3_cumul_dose": "µmol/m²",
            "stom_change_rate": "mmol/m²s",
            "normalised_stom_cond": "",
            "relstom_condchange_rate": "%",
            "leaf_area": "cm²",
            "air_flow_rate": "cm³/s",
            "absorbed_radiation": "cal/cm²s",
            "boundary_layer_cond": "mmol/m²s",
            "cutic_conductance": "mmol/m²s",
            "temp_calibr": "°C",
            "time": ""
        }
        
        # Set up columns with units in the header
        columns = list(self.base_data.columns)
        self.raw_tree['columns'] = columns
        self.raw_tree['show'] = 'headings'
        
        style = ttk.Style()
        style.configure("Treeview", rowheight=25)
    
        # Calculate optimal widths for all columns
        for col in columns:
            # Get unit for this column
            unit = raw_units.get(col, "")
        
            # Create single-line header with unit in parentheses
            if unit:
                header_text = f"{col} ({unit})"
            else:
                header_text = col
            
            self.raw_tree.heading(col, text=header_text)
        
            # Calculate optimal width for this column
            width = self.calculate_optimal_width(self.base_data, col, header_text, is_raw=True)
            self.raw_tree.column(col, width=width, anchor=tk.CENTER, minwidth=width)
        
        # Insert data
        for _, row in self.base_data.iterrows():
            self.raw_tree.insert('', 'end', values=list(row))
    
    def calculate_optimal_width(self, dataframe, column_name, header_text, is_raw=False):
        """
        Calculate optimal column width based on header and data content
        """
         # Base width for header
        header_width = len(header_text) * 7  # Decreased from 9 to 7 pixels per character
    
        # Add extra padding for units
        if '(' in header_text and ')' in header_text:
            header_width += 20  # Increased padding for units
    
        # Calculate maximum data width
        max_data_width = 0
        if len(dataframe) > 0:
            # Sample more data for width calculation
            sample_size = min(40, len(dataframe))  # Increased from 20 to 50 rows
            sample_data = dataframe[column_name].iloc[:sample_size]
        
            for val in sample_data:
                if pd.notna(val):
                    # Format for display
                    if isinstance(val, float):
                        # More comprehensive formatting
                        abs_val = abs(val)
                        if abs_val == 0:
                            text = "0"
                        elif abs_val < 0.0001:
                            text = f"{val:.2e}"
                        elif abs_val < 0.01:
                            text = f"{val:.6f}"
                        elif abs_val < 1:
                            text = f"{val:.4f}"
                        elif abs_val < 10:
                            text = f"{val:.3f}"
                        elif abs_val < 100:
                            text = f"{val:.2f}"
                        elif abs_val < 1000:
                            text = f"{val:.1f}"
                        else:
                            text = f"{val:.0f}"
                    else:
                        text = str(val)
                
                    text_width = len(text) * 8  # Increased from 7 to 8 pixels per character
                    max_data_width = max(max_data_width, text_width)
    
        # Use the larger of header or data width, with extra padding
        optimal_width = max(header_width, max_data_width) + 15  # Increased padding from 25 to 35
    
        # Apply column-specific constraints with higher minimums
        if column_name == 'Channel':
            optimal_width = max(200, min(optimal_width, 400))  # Increased from 150-300 to 200-400
        elif column_name in ['time', 'time_repeated', 'time_full_minutes', 'time_sec']:
            optimal_width = max(150, min(optimal_width, 250))  # Increased from 120-180 to 150-250
        elif column_name in ['absorbed_radiation', 'CO2_exchange_rate', 'Transpiration_H2O_evol_rate', 
                             'Stomatal_conductance_corrected', 'overall_O3_uptake_rate_c']:
            optimal_width = max(180, min(optimal_width, 300))  # Increased from 140-220 to 180-300
        elif 'cond' in column_name.lower() or 'conductance' in column_name.lower():
            optimal_width = max(170, min(optimal_width, 280))  # Increased from 130-200 to 170-280
        elif is_raw and column_name in ['CO2_exch_rate', 'H2O_evol_rate', 'overall_O3_uptake_rate']:
            optimal_width = max(160, min(optimal_width, 250))  # Increased from 120-180 to 160-250
        else:
            optimal_width = max(140, min(optimal_width, 250))  # Increased from 100-180 to 140-250
    
        # FINAL ADJUSTMENT - MAKE ALL COLUMNS EVEN WIDER
        optimal_width = int(optimal_width * 1.5)  # Increase width by 50%
    
        return optimal_width
    
    def update_calculated_data_display(self):
        df = self.get_data_for_calculated_table()
        if df is None:
            return
        
        # Store current column widths and positions before clearing
        if hasattr(self, '_column_widths'):
            saved_widths = self._column_widths.copy()
        else:
            saved_widths = {}
    
        # Clear the tree
        for item in self.calc_tree.get_children():
            self.calc_tree.delete(item)
    
        # Set up columns
        columns = list(df.columns)
        self.calc_tree['columns'] = columns
        self.calc_tree['show'] = 'headings'
    
        # Configure columns with optimal widths
        for col in columns:
            # Determine header text
            display_col_name = col
            if col == 'Channel' and self.parameters_applied['channel_names']:
                display_col_name = 'Channel (Modified)'
        
            # Get unit for this column
            unit = UNITS_MAP.get(col, "")
        
            # Create single-line header with unit in parentheses
            if unit:
                header_text = f"{display_col_name} ({unit})"
            else:
                header_text = display_col_name
            
            self.calc_tree.heading(col, text=header_text)
        
            # Calculate optimal width
            if col in saved_widths and saved_widths[col] > 50:  # Only reuse if reasonable
                width = saved_widths[col]
            else:
                width = self.calculate_optimal_width(df, col, header_text, is_raw=False)
        
            # Apply style for better readability
            style = ttk.Style()
            style.configure("Treeview.Heading", font=('Arial', 9, 'bold'))
        
            self.calc_tree.column(col, width=width, anchor=tk.CENTER, minwidth=80)
            saved_widths[col] = width
    
        # Insert data with proper formatting
        for _, row in df.iterrows():
            formatted_values = []
            for val in row:
                if pd.isna(val):
                    formatted_values.append('')
                else:
                    formatted_values.append(self.format_value_for_display(val))
            self.calc_tree.insert('', 'end', values=formatted_values)
    
        # Save column widths for future updates
        self._column_widths = saved_widths
    
        # Bind column resize events
        self.bind_column_resize_events()
    
    def format_value_for_display(self, value):
        """Format a value for display in tables"""
        if pd.isna(value):
            return ''
        
        if isinstance(value, float):
            abs_val = abs(value)
            if abs_val == 0:
                return "0"
            elif abs_val < 0.0001:
                return f"{value:.2e}"
            elif abs_val < 0.01:
                return f"{value:.6f}"
            elif abs_val < 1:
                return f"{value:.4f}"
            elif abs_val < 10:
                return f"{value:.3f}"
            elif abs_val < 100:
                return f"{value:.2f}"
            elif abs_val < 1000:
                return f"{value:.1f}"
            else:
                return f"{value:.0f}"
        else:
            return str(value)
    
    def calculate_initial_column_width(self, col, df, header_text):
        """Calculate initial column width based on content"""
        # Base width on header text
        header_width = len(header_text) * 7  # Approximate pixels per character
        
        # Check data in column for maximum width
        max_data_width = 0
        if len(df) > 0:
            # Sample first 10 rows to get approximate width
            sample_size = min(10, len(df))
            sample_data = df[col].iloc[:sample_size]
            
            for val in sample_data:
                if pd.notna(val):
                    if isinstance(val, float):
                        # Format similar to display logic
                        if abs(val) < 0.0001:
                            text = f"{val:.2e}"
                        elif abs(val) < 0.01:
                            text = f"{val:.6f}"
                        elif abs(val) < 1:
                            text = f"{val:.4f}"
                        elif abs(val) < 100:
                            text = f"{val:.2f}"
                        else:
                            text = f"{val:.1f}"
                    else:
                        text = str(val)
                    
                    text_width = len(text) * 7
                    max_data_width = max(max_data_width, text_width)
        
        # Use the larger of header or data width, with some padding
        width = max(header_width, max_data_width) + 20
        
        # Set minimum and maximum bounds
        if col == 'Channel':
            width = max(150, min(width, 300))  # Channel column wider
        elif col in ['time', 'time_repeated']:
            width = max(100, min(width, 150))
        elif col in ['CO2_exchange_rate', 'Transpiration_H2O_evol_rate', 'Stomatal_conductance_corrected']:
            width = max(120, min(width, 200))
        else:
            width = max(100, min(width, 200))
        
        return width
    
    def bind_column_resize_events(self):
        """Bind events to save column widths when user resizes them"""
        def on_column_resize(event):
            # Get the current column widths
            if hasattr(self, '_column_widths'):
                for col in self.calc_tree['columns']:
                    self._column_widths[col] = self.calc_tree.column(col, 'width')
        
        # Bind to treeview resize events
        self.calc_tree.bind('<ButtonRelease-1>', on_column_resize)
        self.calc_tree.bind('<B1-Motion>', on_column_resize)
    
    def sort_by_column(self, col):
        """Sort treeview by column when header is clicked"""
        data = []
        for child in self.calc_tree.get_children():
            data.append((self.calc_tree.item(child)['values'], child))
        
        # Get current sort order
        if not hasattr(self, '_sort_order'):
            self._sort_order = {}
        if col not in self._sort_order:
            self._sort_order[col] = 'asc'
        
        # Sort data
        try:
            data.sort(key=lambda x: x[0][list(self.calc_tree['columns']).index(col)], 
                      reverse=(self._sort_order[col] == 'desc'))
        except:
            # If sorting fails (e.g., mixed types), use string comparison
            data.sort(key=lambda x: str(x[0][list(self.calc_tree['columns']).index(col)]), 
                      reverse=(self._sort_order[col] == 'desc'))
        
        # Reorder items in treeview
        for index, (values, child) in enumerate(data):
            self.calc_tree.move(child, '', index)
        
        # Toggle sort order
        self._sort_order[col] = 'desc' if self._sort_order[col] == 'asc' else 'asc'
    
    def get_data_for_calculated_table(self):
        """Get data for the calculated table, ensuring consistency with channel renaming"""
        # Always use modified data if it exists
        if self.modified_calculated_data is not None:
            df = self.modified_calculated_data.copy()
        elif self.original_calculated_data is not None:
            df = self.original_calculated_data.copy()
        else:
            print(f"No data available")
            return None
        
        # Apply channel filtering
        filtered_df = self.get_filtered_data_with_renaming(df)
        
        return filtered_df if filtered_df is not None and len(filtered_df) > 0 else df
    
    def get_filtered_data_with_renaming(self, df):
        """
        Get filtered data that properly handles renamed channels
        """
        if df is None or len(df) == 0:
            return df
        
        if not self.selected_channels:
            return df
        
        # Get current display names from the dataframe (already renamed if applicable)
        df_channels = df['Channel'].astype(str).unique()
        
        # Build a mapping for easy lookup
        df_channel_set = set(df_channels)
        
        # Filter channels that exist in the dataframe
        valid_channels = []
        for display_channel in self.selected_channels:
            # Check if this channel exists in the dataframe (with its current name)
            if display_channel in df_channel_set:
                valid_channels.append(display_channel)
            else:
                # If not found, check if it's an original name that was renamed
                if hasattr(self, 'channel_names'):
                    # Check if this display_channel is a new name in our mapping
                    for orig_name, new_name in self.channel_names.items():
                        orig_str = str(orig_name)
                        new_str = str(new_name)
                        
                        # Case 1: display_channel is the new name
                        if new_str == display_channel and orig_str in df_channel_set:
                            valid_channels.append(orig_str)
                            break
                        # Case 2: display_channel is the original name (not selected in listbox)
                        elif orig_str == display_channel and new_str in df_channel_set:
                            valid_channels.append(new_str)
                            break
        
        if not valid_channels:
            # If no valid channels found, return all data
            return df
        
        # Filter the dataframe
        return df[df['Channel'].astype(str).isin(valid_channels)]
    
    def create_channel_params_ui(self):
        if self.original_calculated_data is None:
            return
            
        for widget in self.channel_params_frame.winfo_children():
            widget.destroy()
        
        # Always use ORIGINAL channel names for the parameter table
        channels = self.original_calculated_data['Channel'].unique()
        headers = ["Channel", "Rename", "Leaf Area", "Air Flow Rate", 
                  "Absorbed Radiation", "Boundary Layer Cond", "Cutic Conductance"]
        
        for i, header in enumerate(headers):
            ttk.Label(self.channel_params_frame, text=header, font=("Arial", 9, "bold")).grid(
                row=0, column=i, padx=5, pady=5, sticky=tk.W)
        
        for idx, channel in enumerate(channels):
            row = idx + 1
            
            # Use original channel name for the label
            channel_str = str(channel)
            
            # Show current display name (either original or renamed)
            if channel_str in self.channel_names:
                display_name = self.channel_names[channel_str]
            else:
                display_name = channel_str
            
            ttk.Label(self.channel_params_frame, text=display_name).grid(
                row=row, column=0, padx=5, pady=2, sticky=tk.W)
            
            # Get the appropriate data for this channel
            channel_data = None
            
            # First try to find data in modified dataframe
            if self.modified_calculated_data is not None and not self.modified_calculated_data.empty:
                # Try to find by original channel name first
                if channel_str in self.channel_names:
                    # Channel was renamed - look for it by new name
                    new_name = self.channel_names[channel_str]
                    modified_channel_data = self.modified_calculated_data[
                        self.modified_calculated_data['Channel'].astype(str) == new_name
                    ]
                else:
                    # Channel wasn't renamed - use original name
                    modified_channel_data = self.modified_calculated_data[
                        self.modified_calculated_data['Channel'].astype(str) == channel_str
                    ]
                
                if not modified_channel_data.empty:
                    channel_data = modified_channel_data
            
            # Fall back to original data
            if channel_data is None:
                channel_data = self.original_calculated_data[self.original_calculated_data['Channel'] == channel]
            
            # Pre-fill rename field with current name (either original or renamed)
            rename_var = tk.StringVar()
            if channel_str in self.channel_names:
                rename_var.set(self.channel_names[channel_str])
            else:
                rename_var.set(channel_str)
            
            rename_entry = ttk.Entry(self.channel_params_frame, textvariable=rename_var, width=15)
            rename_entry.grid(row=row, column=1, padx=5, pady=2)
            setattr(self, f"channel_name_{channel}", rename_entry)
            
            # Rest of the method remains the same...
            leaf_area_var = tk.StringVar(value=str(channel_data['leaf_area'].iloc[0]))
            leaf_area_entry = ttk.Entry(self.channel_params_frame, textvariable=leaf_area_var, width=10)
            leaf_area_entry.grid(row=row, column=2, padx=5, pady=2)
            setattr(self, f"leaf_area_{channel}", leaf_area_entry)
            
            air_flow_var = tk.StringVar(value=str(channel_data['air_flow_rate'].iloc[0]))
            air_flow_entry = ttk.Entry(self.channel_params_frame, textvariable=air_flow_var, width=10)
            air_flow_entry.grid(row=row, column=3, padx=5, pady=2)
            setattr(self, f"air_flow_rate_{channel}", air_flow_entry)
            
            radiation_var = tk.StringVar(value=str(channel_data['absorbed_radiation'].iloc[0]))
            radiation_entry = ttk.Entry(self.channel_params_frame, textvariable=radiation_var, width=10)
            radiation_entry.grid(row=row, column=4, padx=5, pady=2)
            setattr(self, f"absorbed_radiation_{channel}", radiation_entry)
            
            boundary_var = tk.StringVar(value=str(channel_data['boundary_layer_cond'].iloc[0]))
            boundary_entry = ttk.Entry(self.channel_params_frame, textvariable=boundary_var, width=10)
            boundary_entry.grid(row=row, column=5, padx=5, pady=2)
            setattr(self, f"boundary_layer_cond_{channel}", boundary_entry)
            
            cutic_var = tk.StringVar(value=str(channel_data['cutic_conductance'].iloc[0]))
            cutic_entry = ttk.Entry(self.channel_params_frame, textvariable=cutic_var, width=10)
            cutic_entry.grid(row=row, column=6, padx=5, pady=2)
            setattr(self, f"cutic_conductance_{channel}", cutic_entry)
    
    def create_darkness_inputs(self):
        """Create input boxes for darkness periods for each channel - EMPTY by default, preserve manual entries"""
        # Use CURRENT data to get channel names
        df = self.get_data_for_plots()
        if df is None:
            return
    
        # Clear existing widgets
        for widget in self.darkness_inputs_frame.winfo_children():
            widget.destroy()

        # Get current channel names from the data
        channels = df['Channel'].unique()

        # Store widget references in a dictionary
        self.darkness_widgets = {}

        # Create input boxes for each channel
        row = 0
        col = 0
        max_per_col = 6

        for idx, channel in enumerate(channels):
            channel_str = str(channel)
    
            if idx % max_per_col == 0 and idx > 0:
                col += 2
                row = 0
    
            # Channel label - use current display name
            label = ttk.Label(self.darkness_inputs_frame, text=f"Channel {channel_str}:")
            label.grid(row=row, column=col, padx=5, pady=2, sticky=tk.W)
    
            # Create entry widget - START EMPTY
            darkness_entry = ttk.Entry(self.darkness_inputs_frame, width=20)
            darkness_entry.grid(row=row, column=col+1, padx=5, pady=2, sticky=tk.W)
    
            # --- ONLY SHOW MANUAL ENTRIES IN THE INPUT BOXES ---
            # Auto-detected values are for plotting only and should NOT appear in input boxes
            if hasattr(self, 'darkness_settings') and self.darkness_settings:
            
                # Look for settings with this exact channel name
                if channel_str in self.darkness_settings:
                    settings = self.darkness_settings[channel_str]
                    # Check if this is a manual entry with manual_times
                    if settings.get('type') == 'manual' and 'manual_times' in settings:
                        times = settings['manual_times']
                        if times:
                            # Format times as integers (no decimal places) for input boxes
                            formatted_times = ', '.join(str(int(t)) for t in times)
                            darkness_entry.insert(0, formatted_times)
        
            # Store reference in dictionary
            self.darkness_widgets[channel_str] = darkness_entry
        
            row += 1

        # Configure columns to expand properly
        for i in range(col+2):
            self.darkness_inputs_frame.columnconfigure(i, weight=1)
    
    def clear_darkness_inputs(self):
        """Clear all darkness input boxes"""
        if hasattr(self, 'darkness_widgets'):
            for channel_str, entry_widget in self.darkness_widgets.items():
                entry_widget.delete(0, tk.END)
    
    def create_radiation_inputs(self):
        """Create input boxes for radiation intervals for each channel"""
        # Use CURRENT data to get channel names
        df = self.get_data_for_plots()
        if df is None:
            return
            
        # Clear existing widgets
        for widget in self.radiation_inputs_frame.winfo_children():
            widget.destroy()
        
        # Get current channel names from the data
        channels = df['Channel'].unique()
        
        # Store widget references in a dictionary
        self.radiation_widgets = {}  # Add this line
        
        # Create input boxes for each channel
        row = 0
        col = 0
        max_per_col = 6
        
        for idx, channel in enumerate(channels):
            channel_str = str(channel)
            
            if idx % max_per_col == 0 and idx > 0:
                col += 2
                row = 0
            
            # Channel label
            label = ttk.Label(self.radiation_inputs_frame, text=f"Channel {channel_str}:")
            label.grid(row=row, column=col, padx=5, pady=2, sticky=tk.W)
            
            # Create entry widget
            radiation_entry = ttk.Entry(self.radiation_inputs_frame, width=30)
            radiation_entry.grid(row=row, column=col+1, padx=5, pady=2, sticky=tk.W)
            
            # Store reference in dictionary instead of attribute
            self.radiation_widgets[channel_str] = radiation_entry
            
            row += 1
        
        # Configure columns to expand properly
        for i in range(col+2):
            self.radiation_inputs_frame.columnconfigure(i, weight=1)
    
    def update_all_plots(self):
        """Update all plots with current data and settings"""
        print("\n=== UPDATE ALL PLOTS CALLED ===")
        print(f"Current selected_channels: {self.selected_channels}")
    
        # Clear all old figures first
        self.clear_old_figures()
        plt.close('all')  # Additional cleanup
    
        df = self.get_data_for_plots()
        print(f"Data for plots - shape: {df.shape if df is not None else 'None'}, empty: {df.empty if df is not None else 'N/A'}")
    
        # Get current palette from the variable
        palette = self.palette_var.get()
    
        # Update main plots with current palette
        print("Calling update_main_plots")
        self.update_main_plots(df, palette)
        
        # Create darkness inputs (will create empty boxes)
        self.create_darkness_inputs()  # This should recreate with current channel names
        self.create_radiation_inputs()  # This recreates with current channel names
        
        # Update darkness plots if data exists
        if self.darkness_data is not None:
            self.update_darkness_plots(palette)
        
        # Update radiation plots if data exists
        if self.radiation_data is not None:
            self.update_radiation_plots(palette)
        
        # Update time range info
        self.update_time_range_info()
    
    def update_main_plots(self, df, palette):
        plot_categories = {
            'environment': [
                ("air_temp", "Air Temperature", "Temperature"),
                ("saturating_air_humidity", "Saturation Humidity at Air Temperature", "Saturation Humidity"),
                ("absorbed_radiation", "Absorbed Radiation", "Radiation"),
                ("relative_air_humidity", "Relative Air Humidity", "Humidity"),
                ("VPD", "Vapour Pressure Deficit", "VPD"),
                ("rad_ind_leaf_temp_increase", "Radiation-induced Leaf Temperature Increase", "Temperature Increase")
            ],
            'base_fluxes': [
                ("CO2_exchange_rate", "CO₂ Exchange Rate", "CO₂ Exchange"),
                ("Transpiration_H2O_evol_rate", "Transpiration H₂O Evolution Rate", "Transpiration"),
                ("transp_ind_leaf_temp_depr", "Transpiration-induced Leaf Temperature Depression", "Temperature Depression"),
                ("leaf_temp_c", "Leaf Temperature", "Leaf Temperature")
            ],
            'humidity_vapor': [
                ("leaf_air_hum_grad", "Leaf Air Humidity Gradient", "Gradient"),
                ("satur_hum_at_leaf_temp_c", "Saturation Humidity at Leaf Temperature", "Saturation Humidity"),
                ("mean_vapor_pressure", "Mean Vapor Pressure", "Pressure")
            ],
            'conductances': [
                ("overall_conductance", "Overall Conductance", "Conductance"),
                ("corr_d_overall_conductance", "Corr d Overall Conductance", "Conductance"),
                ("massflow_correction_for_overall_resistance", "Massflow Correction for Overall Resistance", "Correction Factor"),
                ("b_layer_conductance", "Boundary Layer Conductance", "Conductance"),
                ("corr_leaf_conductance", "Corrected Leaf Conductance", "Conductance"),
                ("corr_stomatal_conductance", "Corr Stomatal Conductance", "Conductance"),
                ("Stomatal_conductance_corrected", "Stomatal Conductance Corrected", "Conductance"),
                ("normalised_corr_stomatal_conductance", "Normalised Stomatal Conductance", "Normalised Conductance"),
                ("corr_uncorr_stom_cond", "Corrected/Uncorrected Stomatal Conductance Ratio", "Ratio"),
                ("corr_stomatal_change_rate", "Corrected Stomatal Change Rate", "Change Rate"),
                ("relat_stom_cond_change_rate", "Relative Stomatal Conductance Change Rate", "Change Rate")
            ],
            'ozone': [
                ("overall_O3_uptake_rate_c", "Overall O₃ Uptake Rate", "Uptake Rate"),
                ("corr_conductance_for_O3_c", "Corrected Conductance for O₃", "Conductance"),
                ("corr_stomatal_O3_uptake_rate", "Corrected Stomatal O₃ Uptake Rate", "Uptake Rate"),
                ("corr_cumul_O3_dose", "Corrected Cumulative O₃ Dose", "Dose")
            ],
            'co2': [
                ("massflow_corr_d_ca", "Massflow Corr d ca", "CO₂"),
                ("massflow_corr_for_CO2", "Massflow Correction for CO₂", "Correction Factor"),
                ("mflowcorr_Ca_gradient", "Massflow Correction for CO₂ Gradient", "Correction Factor"),
                ("corr_uncorr_ca", "Corrected/Uncorrected Ambient CO₂ Ratio", "Ratio"),
                ("stom_resist_to_CO2", "Stomatal Resistance to CO₂", "Resistance"),
                ("b_layer_resist_to_CO2", "Boundary Layer Resistance to CO₂", "Resistance"),
                ("corr_ted_CO2_grad", "Corr ted CO₂ grad", "Gradient"),
                ("intercell_CO2_conc_in_gas_phase", "Intercellular CO₂ Concentration in Gas Phase", "CO₂"),
                ("uncorr_ci", "Uncorr ci", "ci"),
                ("corr_uncorr_ci", "Corr uncorr ci", "Ratio")
            ],
            'mesophyll': [
                ("mesophyll_conductance_for_CO2", "Mesophyll Conductance for CO₂", "Conductance"),
                ("uncorr_gm_prima", "Uncorrected gm'", "gm'"),
                ("corr_uncorr_gm_prima", "Corrected/Uncorrected gm'", "Ratio"),
                ("CO2_comp_point", "CO₂ Compensation Point", "CO₂"),
                ("uncorr_gamma", "Uncorrected Gamma", "Gamma"),
                ("corr_uncorr_gamma", "Corrected/Uncorrected Gamma", "Ratio")
            ],
            'efficiency': [
                ("WUEi", "Intrinsic Water Use Efficiency", "WUEi"),
                ("WUE", "Water Use Efficiency", "WUE")
            ]
        }
        
        for category, plots in plot_categories.items():
            if category in self.plot_tabs:
                self.update_plot_tab(category, plots, df, palette)
        
        # Update special plots with palette
        self.update_special_plots(df, palette)
        
        # Add WUEi comparison plot to efficiency tab (this will be the 3rd plot)
        self.update_wuei_comparison_plot(df, palette)
        
    def update_wuei_comparison_plot(self, df, palette):
        """Update the WUEi comparison plot in the efficiency tab - as the 3rd plot"""
        if df is None or 'WUEi' not in df.columns:
            return
        
        # Get the efficiency tab frame
        tab_frame = self.plot_tabs['efficiency']
        
        # Count how many plots are already in the tab (should be 2: WUEi and WUE)
        existing_plot_containers = len([w for w in tab_frame.winfo_children() 
                                       if isinstance(w, ttk.Frame) and hasattr(w, 'grid_info')])
        
        # Create the combined plot
        fig = create_WUEi_comparison_plot(df, palette)
        fig.set_size_inches(12, 6)  # Set consistent figure size
        
        # Create a container frame for the plot - place it after existing plots
        plot_container = ttk.Frame(tab_frame)
        
        # Calculate the correct row position
        row_position = existing_plot_containers  # This should be 2 (0-indexed after 2 plots)
        
        plot_container.grid(row=row_position, column=0, 
                           sticky=(tk.W, tk.E), padx=10, pady=(20, 0), columnspan=2)
        
        # Add title label
        ttk.Label(plot_container, text="WUEi Distribution by Channel", 
                 font=("Arial", 11, "bold")).grid(
            row=0, column=0, sticky=tk.W)
        
        # Create canvas for the plot
        canvas = FigureCanvasTkAgg(fig, plot_container)
        canvas.draw()
        canvas.get_tk_widget().grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # Add download button
        btn = ttk.Button(plot_container, text="Download WUEi Distribution Analysis",
                    command=lambda f=fig, t="WUEi_Distribution_Analysis": self.download_plot(f, t))
        btn.grid(row=2, column=0, sticky=tk.W, pady=(5, 20))
        
        # Configure plot container column to expand
        plot_container.columnconfigure(0, weight=1)
        
        # Make sure the tab frame expands properly
        tab_frame.columnconfigure(0, weight=1)
    
    def update_darkness_plots(self, palette=None):
        """Update darkness plots with yellow and black rectangles - ALL in one tab"""
        # Use the provided palette or get the current one from the UI
        if palette is None:
            palette = self.palette_var.get()
    
        # IMPORTANT: Apply channel filtering to darkness_data
        if self.darkness_data is not None:
            df = self.get_filtered_data(self.darkness_data)
            print(f"Darkness plots - filtered df shape: {df.shape if df is not None else 'None'}")
        else:
            df = None
    
        if df is None or len(df) == 0:
            print("No darkness data available or all channels hidden - showing empty plots")
        # Clear the darkness plots tab and show empty plots
            if 'darkness_plots' in self.plot_tabs:
                self.clear_plot_tab('darkness_plots')
            return
            
        # Define ALL darkness plots configuration in the exact order requested
        darkness_plots = [
            # 1. Absorbed radiation
            ("absorbed_radiation", "Absorbed Radiation with Darkness Periods", "Radiation"),
            
            # 2. CO₂ exchange rate
            ("CO2_exchange_rate", "CO₂ Exchange Rate - Light/Dark", "CO₂ Exchange Rate"),
            
            # 3. Transpiration H₂O evolution rate
            ("Transpiration_H2O_evol_rate", "Transpiration H₂O Evolution Rate - Light/Dark", "Transpiration Rate"),
            
            # 4. Leaf temperature
            ("leaf_temp_c", "Leaf Temperature - Light/Dark", "Leaf Temperature"),
            
            # 5. VPD (Vapor Pressure Deficit)
            ("VPD", "Vapor Pressure Deficit (VPD) - Light/Dark", "VPD"),
            
            # 6. Stomatal conductance (corrected)
            ("Stomatal_conductance_corrected", "Stomatal Conductance (Corrected) - Light/Dark", "Conductance"),
            
            # 7. Overall conductance
            ("overall_conductance", "Overall Conductance - Light/Dark", "Conductance"),
            
            # 8. Boundary layer conductance
            ("b_layer_conductance", "Boundary Layer Conductance - Light/Dark", "Conductance"),
            
            # 9. Corrected leaf conductance
            ("corr_leaf_conductance", "Corrected Leaf Conductance - Light/Dark", "Conductance"),
            
            # 10. Corrected stomatal conductance
            ("corr_stomatal_conductance", "Corrected Stomatal Conductance - Light/Dark", "Conductance"),
            
            # 11. Stomatal resistance to CO₂
            ("stom_resist_to_CO2", "Stomatal Resistance to CO₂ - Light/Dark", "Resistance"),
            
            # 12. Boundary layer resistance to CO₂
            ("b_layer_resist_to_CO2", "Boundary Layer Resistance to CO₂ - Light/Dark", "Resistance"),
            
            # 13. Vapor pressure around leaves
            ("vapor_pressure_around_leaves", "Vapor Pressure Around Leaves - Light/Dark", "Vapor Pressure"),
            
            # 14. Saturation vapor pressure inside leaves
            ("saturation_vapor_pressure_inside_leaves", "Saturation Vapor Pressure Inside Leaves - Light/Dark", "Vapor Pressure"),
            
            # 15. Normalised corrected stomatal conductance
            ("normalised_corr_stomatal_conductance", "Normalised Corrected Stomatal Conductance - Light/Dark", "Normalised Conductance"),
            
            # 16. Intercellular CO₂ concentration
            ("intercell_CO2_conc_in_gas_phase", "Intercellular CO₂ Concentration - Light/Dark", "CO₂ Concentration"),
            
            # 17. Mesophyll conductance
            ("mesophyll_conductance_for_CO2", "Mesophyll Conductance for CO₂ - Light/Dark", "Conductance"),
            
            # 18. Corrected CO₂ gradient
            ("corr_ted_CO2_grad", "Corrected CO₂ Gradient - Light/Dark", "CO₂ Gradient"),
            
            # 19. Corrected conductance for O₃
            ("corr_conductance_for_O3_c", "Corrected Conductance for O₃ - Light/Dark", "Conductance"),
            
            # 20. Corrected stomatal O₃ uptake rate
            ("corr_stomatal_O3_uptake_rate", "Corrected Stomatal O₃ Uptake Rate - Light/Dark", "O₃ Uptake Rate"),
            
            # 21. Corrected cumulative O₃ dose
            ("corr_cumul_O3_dose", "Corrected Cumulative O₃ Dose - Light/Dark", "O₃ Dose"),
            
            # 22. Corrected stomatal change rate
            ("corr_stomatal_change_rate", "Corrected Stomatal Change Rate - Light/Dark", "Change Rate"),
            
            # 23. Relative stomatal conductance change rate
            ("relat_stom_cond_change_rate", "Relative Stomatal Conductance Change Rate - Light/Dark", "Change Rate")
        ]
        
        # Filter out plots that don't exist in the data
        available_plots = []
        for yvar, title, ylab in darkness_plots:
            if yvar in df.columns:
                available_plots.append((yvar, title, ylab))
            else:
                print(f"Warning: {yvar} not found in data, skipping plot")
    
        # Get darkness times for plotting (also filtered)
        darkness_times = self.get_darkness_times_for_plotting()
    
        # Update the single darkness plots tab
        if 'darkness_plots' in self.plot_tabs:
            self.update_darkness_plot_tab('darkness_plots', available_plots, df, palette, darkness_times)
                
    def get_darkness_times_for_plotting(self):
        """Extract just the times list for plotting - includes both auto and manual times"""
        if not self.darkness_settings or self.darkness_data is None:
            return {}
    
        df = self.darkness_data.copy()
        plot_times = {}
    
        # IMPORTANT: Only include channels that are currently selected
        selected_channels_display = self.selected_channels
    
        for channel_str, dark_info in self.darkness_settings.items():
            # Check if this channel should be shown based on current selection
            channel_visible = False
        
            # Map the stored channel name to its current display name
            current_display_name = channel_str
            if hasattr(self, 'channel_names') and self.channel_names:
                if channel_str in self.channel_names:
                    current_display_name = self.channel_names[channel_str]
                elif channel_str in self.channel_names.values():
                    current_display_name = channel_str
        
            # Check if this channel (by any name) is in selected channels
            if current_display_name in selected_channels_display:
                channel_visible = True
            elif channel_str in selected_channels_display:
                channel_visible = True
        
            if not channel_visible:
                print(f"Channel {channel_str} not selected, skipping darkness times")
                continue
        
            # Try to find channel data
            channel_data = df[df['Channel'].astype(str) == channel_str]
            if channel_data.empty:
                # Try with the display name
                channel_data = df[df['Channel'].astype(str) == current_display_name]
        
            if channel_data.empty:
                continue
        
            # Get valid times that exist in current data
            valid_times = []
        
            # Use ALL times (auto + manual combined)
            for original_time in dark_info['times']:
                # Find the closest time point in current data
                time_diffs = abs(channel_data['time_repeated'] - original_time)
                if not time_diffs.empty and time_diffs.min() < 0.01:
                    closest_idx = time_diffs.idxmin()
                    valid_times.append(channel_data.loc[closest_idx, 'time_repeated'])
        
            if valid_times:
                plot_times[current_display_name] = valid_times
    
        return plot_times
                
    def update_darkness_plot_tab(self, category, plot_configs, df, palette, darkness_times):
        """Update the darkness plot tab with yellow and black rectangles"""
        tab_frame = self.plot_tabs[category]

        # Destroy old canvas widgets and close figures
        for widget in tab_frame.winfo_children():
            if hasattr(widget, 'destroy'):
                widget.destroy()
            elif hasattr(widget, 'get_tk_widget'):
                widget.get_tk_widget().destroy()
    
        # If no data or empty dataframe, show message
        if df is None or len(df) == 0:
            fig, ax = plt.subplots(figsize=(12, 6))
            ax.text(0.5, 0.5, 'No data available - all channels hidden', 
                    ha='center', va='center', fontsize=14, fontweight='bold')
            ax.set_axis_off()
        
            plot_container = ttk.Frame(tab_frame)
            plot_container.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=20)
        
            canvas = FigureCanvasTkAgg(fig, plot_container)
            canvas.draw()
            canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.W, tk.E))
        
            tab_frame.columnconfigure(0, weight=1)
            return
    
        # Rest of your existing code for plotting...
        row = 0
        for yvar, title, ylab in plot_configs:
            if yvar in df.columns:
                fig = make_light_dark_plot(df, yvar, title, ylab, palette, darkness_times)
                fig.set_size_inches(12, 6)
            
                plot_container = ttk.Frame(tab_frame)
                plot_container.grid(row=row*3, column=0, sticky=(tk.W, tk.E), 
                                  padx=10, pady=(20, 0), columnspan=2)
            
                ttk.Label(plot_container, text=title, font=("Arial", 11, "bold")).grid(
                    row=0, column=0, sticky=tk.W)
            
                canvas = FigureCanvasTkAgg(fig, plot_container)
                canvas.draw()
                canvas.get_tk_widget().grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
            
                btn = ttk.Button(plot_container, text=f"Download {title}",
                          command=lambda f=fig, t=title: self.download_plot(f, t))
                btn.grid(row=2, column=0, sticky=tk.W, pady=(5, 20))
            
                plot_container.columnconfigure(0, weight=1)
                row += 1
    
        tab_frame.columnconfigure(0, weight=1)
    
    def update_radiation_plots(self, palette=None):
        """Update radiation plots with ALL variables - MATCHING DARKNESS TAB STRUCTURE"""
        if palette is None:
            palette = self.palette_var.get()
    
        # IMPORTANT: Apply channel filtering to radiation_data
        if self.radiation_data is not None:
            df = self.get_filtered_data(self.radiation_data)
            print(f"Radiation plots - filtered df shape: {df.shape if df is not None else 'None'}")
            print(f"Radiation plots - filtered df empty: {df.empty if df is not None else 'N/A'}")
        else:
            df = None
            print("Radiation plots - no radiation_data available")
    
        if df is None or len(df) == 0:
            print("No radiation data available or all channels hidden - showing empty plots")
            # Clear the radiation plots tab and show empty plots
            if 'radiation_plots' in self.plot_tabs:
                # Create empty plot configs list - will trigger empty handling in update_radiation_plot_tab
                self.update_radiation_plot_tab('radiation_plots', [], df, palette)
            return
        
        # Define ALL radiation plots configuration - SAME 23 VARIABLES AS DARKNESS TAB
        radiation_plots = [
            # 1. Absorbed radiation
            ("absorbed_radiation", "Absorbed Radiation (R B light)", "Radiation"),
        
            # 2. CO₂ exchange rate
            ("CO2_exchange_rate", "CO₂ Exchange Rate - (R B light)", "CO₂ Exchange Rate"),
        
            # 3. Transpiration H₂O evolution rate
            ("Transpiration_H2O_evol_rate", "Transpiration H₂O Evolution Rate - (R B light)", "Transpiration Rate"),
        
            # 4. Leaf temperature
            ("leaf_temp_c", "Leaf Temperature - (R B light)", "Leaf Temperature"),
        
            # 5. VPD (Vapor Pressure Deficit)
            ("VPD", "Vapor Pressure Deficit (VPD) - (R B light)", "VPD"),
        
            # 6. Stomatal conductance (corrected)
            ("Stomatal_conductance_corrected", "Stomatal Conductance (Corrected) - (R B light)", "Conductance"),
        
            # 7. Overall conductance
            ("overall_conductance", "7. Overall Conductance - (R B light)", "Conductance"),
        
            # 8. Boundary layer conductance
            ("b_layer_conductance", "Boundary Layer Conductance - (R B light)", "Conductance"),
        
            # 9. Corrected leaf conductance
            ("corr_leaf_conductance", "Corrected Leaf Conductance - (R B light)", "Conductance"),
        
            # 10. Corrected stomatal conductance
            ("corr_stomatal_conductance", "Corrected Stomatal Conductance - (R B light)", "Conductance"),
        
            # 11. Stomatal resistance to CO₂
            ("stom_resist_to_CO2", "Stomatal Resistance to CO₂ - (R B light)", "Resistance"),
        
            # 12. Boundary layer resistance to CO₂
            ("b_layer_resist_to_CO2", "Boundary Layer Resistance to CO₂ - (R B light)", "Resistance"),
        
            # 13. Vapor pressure around leaves
            ("vapor_pressure_around_leaves", "Vapor Pressure Around Leaves - (R B light)", "Vapor Pressure"),
        
            # 14. Saturation vapor pressure inside leaves
            ("saturation_vapor_pressure_inside_leaves", "Saturation Vapor Pressure Inside Leaves - (R B light)", "Vapor Pressure"),
        
            # 15. Normalised corrected stomatal conductance
            ("normalised_corr_stomatal_conductance", "Normalised Corrected Stomatal Conductance - (R B light)", "Normalised Conductance"),
        
            # 16. Intercellular CO₂ concentration
            ("intercell_CO2_conc_in_gas_phase", "Intercellular CO₂ Concentration - (R B light)", "CO₂ Concentration"),
        
            # 17. Mesophyll conductance
            ("mesophyll_conductance_for_CO2", "Mesophyll Conductance for CO₂ - (R B light)", "Conductance"),
        
            # 18. Corrected CO₂ gradient
            ("corr_ted_CO2_grad", "Corrected CO₂ Gradient - (R B light)", "CO₂ Gradient"),
        
            # 19. Corrected conductance for O₃
            ("corr_conductance_for_O3_c", "Corrected Conductance for O₃ - (R B light)", "Conductance"),
        
            # 20. Corrected stomatal O₃ uptake rate
            ("corr_stomatal_O3_uptake_rate", "Corrected Stomatal O₃ Uptake Rate - (R B light)", "O₃ Uptake Rate"),
        
            # 21. Corrected cumulative O₃ dose
            ("corr_cumul_O3_dose", "Corrected Cumulative O₃ Dose - (R B light)", "O₃ Dose"),
        
            # 22. Corrected stomatal change rate
            ("corr_stomatal_change_rate", "Corrected Stomatal Change Rate - (R B light)", "Change Rate"),
        
            # 23. Relative stomatal conductance change rate
            ("relat_stom_cond_change_rate", "Relative Stomatal Conductance Change Rate - (R B light)", "Change Rate")
        ]
    
        # Filter out plots that don't exist in the data
        available_plots = []
        for yvar, title, ylab in radiation_plots:
            if yvar in df.columns:
                available_plots.append((yvar, title, ylab))
            else:
                print(f"Warning: {yvar} not found in radiation data, skipping plot")
    
        # Update the single radiation plots tab
        if 'radiation_plots' in self.plot_tabs:
            self.update_radiation_plot_tab('radiation_plots', available_plots, df, palette)
    
    def clear_plot_tab(self, tab_name):
        """Clear a plot tab and show empty plots with 'No data available - all channels hidden' messages"""
        if tab_name not in self.plot_tabs:
            return
    
        tab_frame = self.plot_tabs[tab_name]
    
        # Destroy old canvas widgets
        for widget in tab_frame.winfo_children():
            if hasattr(widget, 'destroy'):
                widget.destroy()
            elif hasattr(widget, 'get_tk_widget'):
                widget.get_tk_widget().destroy()
    
        # Show a single message that no data is available
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.text(0.5, 0.5, 'No data available', 
                ha='center', va='center', fontsize=14, fontweight='bold')
        ax.set_axis_off()
    
        # Create container for the message
        plot_container = ttk.Frame(tab_frame)
        plot_container.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=20)
    
        canvas = FigureCanvasTkAgg(fig, plot_container)
        canvas.draw()
        canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.W, tk.E))
    
        tab_frame.columnconfigure(0, weight=1)
    
    def update_plot_tab(self, category, plot_configs, df, palette, highlight_darkness=False):
        tab_frame = self.plot_tabs[category]

        # Destroy old canvas widgets and close figures
        for widget in tab_frame.winfo_children():
            if hasattr(widget, 'destroy'):
                widget.destroy()
            elif hasattr(widget, 'get_tk_widget'):
                # This is a FigureCanvasTkAgg widget
                widget.get_tk_widget().destroy()
        
        row = 0
        for yvar, title, ylab in plot_configs:
            if yvar in df.columns:
                # Use the passed palette parameter
                fig = make_channel_plot(df, yvar, title, ylab, palette, 
                                       highlight_darkness, self.darkness_settings)
                
                # Set consistent figure size for ALL tabs (including Light-Dark)
                fig.set_size_inches(12, 6)  # Increased from 10,6 to match main plots
                
                # Create a container frame for each plot to ensure proper layout
                plot_container = ttk.Frame(tab_frame)
                plot_container.grid(row=row*3, column=0, sticky=(tk.W, tk.E), 
                                  padx=10, pady=(20, 0), columnspan=2)
                
                # Add title label
                ttk.Label(plot_container, text=title, font=("Arial", 11, "bold")).grid(
                    row=0, column=0, sticky=tk.W)
                
                # Create canvas for the plot
                canvas = FigureCanvasTkAgg(fig, plot_container)
                canvas.draw()
                canvas.get_tk_widget().grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
                
                # Add download button
                btn = ttk.Button(plot_container, text=f"Download {title}",
                          command=lambda f=fig, t=title: self.download_plot(f, t))
                btn.grid(row=2, column=0, sticky=tk.W, pady=(5, 20))
                
                # Configure plot container column to expand
                plot_container.columnconfigure(0, weight=1)
                
                row += 1
        
        # Make sure the tab frame expands properly
        tab_frame.columnconfigure(0, weight=1)
      
    def update_special_plots(self, df, palette):
        special_plots = [
            ('humidity_vapor', 'vapor_pressure_comparison',
             ["saturation_vapor_pressure_inside_leaves", "vapor_pressure_around_leaves"],
             "Vapor Pressure: Inside vs Around Leaves", "Vapor Pressure",
             ["Inside Leaves", "Around Leaves"])
        ]
        
        for category, plot_name, cols, title, ylab, series_names in special_plots:
            if category in self.plot_tabs and all(col in df.columns for col in cols):
                tab_frame = self.plot_tabs[category]
                
                # Close any existing figure
                plt.close('all')
                
                # Use the passed palette parameter
                fig = make_dual_channel_plot(df, cols[0], cols[1], title, ylab, series_names, palette)
                fig.set_size_inches(12, 6)  # Increased to match main plots
                
                # Create a container frame for each plot to ensure proper layout
                plot_container = ttk.Frame(tab_frame)
                plot_container.grid(row=len(tab_frame.winfo_children()), column=0, 
                                  sticky=(tk.W, tk.E), padx=10, pady=(20, 0), columnspan=2)
                
                # Add title label
                ttk.Label(plot_container, text=title, font=("Arial", 11, "bold")).grid(
                    row=0, column=0, sticky=tk.W)
                
                # Create canvas for the plot
                canvas = FigureCanvasTkAgg(fig, plot_container)
                canvas.draw()
                canvas.get_tk_widget().grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
                
                # Add download button
                btn = ttk.Button(plot_container, text=f"Download {title}",
                          command=lambda f=fig, t=title: self.download_plot(f, t))
                btn.grid(row=2, column=0, sticky=tk.W, pady=(5, 20))
                
                # Configure plot container column to expand
                plot_container.columnconfigure(0, weight=1)
      
    def update_wuei_plot(self, df):
        """Update the WUEi comparison plot in the efficiency tab"""
        if df is None or 'WUEi' not in df.columns:
            return
        
        # Get the efficiency tab frame
        tab_frame = self.plot_tabs['efficiency']
        
        # Count existing plots to determine position
        existing_plots = len(tab_frame.winfo_children()) // 3
        row = existing_plots
        
        # Create the combined plot with palette parameter
        fig = create_WUEi_comparison_plot(df, self.palette_var.get())
        
        canvas = FigureCanvasTkAgg(fig, tab_frame)
        canvas.draw()
        
        # Add title
        ttk.Label(tab_frame, text="WUEi Distribution Analysis", 
                 font=("Arial", 11, "bold")).grid(
            row=row*2, column=0, sticky=tk.W, padx=10, pady=(20, 0))
        
        # Add canvas
        canvas.get_tk_widget().grid(row=row*2+1, column=0, 
                                   sticky=(tk.W, tk.E), padx=10, pady=(0, 20))
        
        # Add download button
        btn = ttk.Button(tab_frame, text="Download WUEi Analysis",
                        command=lambda f=fig, t="WUEi_Analysis": self.download_plot(f, t))
        btn.grid(row=row*2+1, column=1, padx=10, pady=10)
    
    def _set_dialog_icon(self, dialog):
        """Helper method to set icon on dialog windows"""
        try:
            # Try to set icon using iconbitmap (Windows)
            dialog.iconbitmap("favicon32.ico")
        except:
            try:
                # Try using PhotoImage (cross-platform)
                icon = tk.PhotoImage(file="favicon32.ico")
                dialog.iconphoto(True, icon)
                dialog.icon_image = icon  # Keep reference
            except:
                pass  # Icon is optional, continue without it
    
    def download_data(self):
        """Export data with option: re-import (28 columns) or all columns (complete data)"""

        # Determine which data to export
        if self.modified_calculated_data is not None:
            df = self.modified_calculated_data.copy()
        elif self.original_calculated_data is not None:
            df = self.original_calculated_data.copy()
        else:
            messagebox.showwarning("Warning", "No data available to download")
            return

        # Create custom dialog for export options
        dialog = tk.Toplevel(self.root)
        dialog.title("Export Options")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Set icon for the dialog
        self._set_dialog_icon(dialog)
    
        # Center the dialog
        dialog_width = 400
        dialog_height = 250
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        x = (screen_width - dialog_width) // 2
        y = (screen_height - dialog_height) // 2
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
    
        # Add instruction label
        ttk.Label(dialog, text="Select export format:", font=("Arial", 12, "bold")).pack(pady=20)
    
        # Add description
        desc_frame = ttk.Frame(dialog)
        desc_frame.pack(pady=10)
        ttk.Label(desc_frame, text="Choose the format for your exported data:\n\n- RE-IMPORT (28 columns compatible)\n- ALL columns (all calculated columns):", font=("Arial", 10)).pack()
    
        # Variable to store the result
        result = [None]
    
        def on_reimport():
            result[0] = "reimport"
            dialog.destroy()
    
        def on_all_columns():
            result[0] = "all_columns"
            dialog.destroy()
    
        # Button frame
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
    
        ttk.Button(button_frame, text="RE-IMPORT", command=on_reimport, 
                   width=20).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="ALL COLUMNS", command=on_all_columns, 
                   width=20).pack(side=tk.LEFT, padx=10)
    
        # Wait for dialog to close
        self.root.wait_window(dialog)
    
        # Check if user closed the dialog without choosing
        if result[0] is None:
            return
    
        export_mode = result[0]

        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[
                ("TSV/TXT files", "*.txt"),
                ("All files", "*.*")
            ]
        )

        if not filename:
            return

        # Ensure .txt extension
        if not filename.endswith('.txt'):
            filename = filename.rsplit('.', 1)[0] + '.txt' if '.' in filename else filename + '.txt'

        try:
            # Get current date and time
            now = datetime.now()
            date_str = now.strftime("%Y-%m-%d")
            time_str = now.strftime("%H:%M:%S")
        
            # Get the original filename if available
            original_filename = self.file_path_var.get() if hasattr(self, 'file_path_var') else "Unknown"
            modified_text = f"Modified file: {os.path.basename(original_filename)}"
        
            # Define the original 28 columns (exactly as from raw data)
            original_28_columns = [
                "Channel", "CO2_M", "CO2_R", "H2O_M", "H2O_R", "O3_M", "O3_R",
                "CO2_exch_rate", "H2O_evol_rate", "overall_O3_uptake_rate", "air_temp", 
                "transp_ind_leaf_temp_depr", "leaf_temp", "satur_hum_at_leaf_temp", 
                "leaf_air_hum_grad", "stomatal_cond", "stomatal_O3_uptake_rate", "O3_cumul_dose",
                "stom_change_rate", "normalised_stom_cond", "relstom_condchange_rate", "leaf_area", 
                "air_flow_rate", "absorbed_radiation", "boundary_layer_cond", "cutic_conductance", 
                "temp_calibr", "time"
            ]
        
            # Create a copy for export
            df_export = df.copy()
        
            # Apply column selection based on user choice
            if export_mode == "reimport":
                # Only keep columns that actually exist in the dataframe
                export_columns = [col for col in original_28_columns if col in df_export.columns]
                df_export = df_export[export_columns].copy()
                mode_text = "Re-import compatible (original 28 columns only)"
            else:  # all_columns
                mode_text = "Complete data (all calculated columns)"
        
            # Calculate continuous seconds for sorting (handles day rollover)
            time_seconds = []
            for time_val in df_export['time']:
                parts = str(time_val).split(':')
                if len(parts) >= 3:
                    secs = float(parts[0]) * 3600 + float(parts[1]) * 60 + float(parts[2])
                else:
                    secs = 0.0
                time_seconds.append(secs)
        
            # Detect day rollover and calculate continuous seconds
            day_increment = 0
            full_seconds = []
            prev_sec = time_seconds[0] if time_seconds else 0
            max_day_increment = 0  # Track the maximum day increment
        
            for sec in time_seconds:
                if sec < prev_sec:
                    day_increment += 1
                    if day_increment > max_day_increment:
                        max_day_increment = day_increment
                prev_sec = sec
                full_seconds.append(sec + day_increment * 24 * 3600)
        
            # Add temporary continuous time column for sorting
            df_export['_sort_time'] = full_seconds
        
            # Sort by continuous time first, then by channel
            df_export = df_export.sort_values(['_sort_time', 'Channel'])
        
            # Remove the temporary sorting column
            df_export = df_export.drop(columns=['_sort_time'])
        
            # Convert numeric values to string with comma as decimal separator
            for col in df_export.columns:
                if df_export[col].dtype in ['float64', 'float32', 'int64', 'int32']:
                    df_export[col] = df_export[col].apply(
                        lambda x: str(x).replace('.', ',') if pd.notna(x) else ''
                    )
        
            # Create units row
            units_row = []
            for col in df_export.columns:
                unit = UNITS_MAP.get(col, "")
                unit = unit.replace('₂', '2').replace('₃', '3').replace('ₒ', 'o')
                unit = unit.replace('µ', 'u').replace('μ', 'u')
                units_row.append(unit)
        
            # Add empty column at the beginning
            headers = df_export.columns.tolist()
            headers.insert(0, "")
            units_row.insert(0, "")
        
            # Add empty column to dataframe
            df_export.insert(0, "", "")
        
            # Write as TSV (tab-separated) with .txt extension
            with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f, delimiter='\t')
            
                # Row 1: Empty first cell, then Date, Time, Modified text, Export mode
                writer.writerow(["", f"Date: {date_str}", f"Time: {time_str}", modified_text, f"Mode: {mode_text}"])
            
                # Row 2: Empty row
                writer.writerow([])
            
                # Row 3: Headers
                writer.writerow(headers)
            
                # Row 4: Units row
                writer.writerow(units_row)
            
                # Write data starting from row 5
                for _, row in df_export.iterrows():
                    writer.writerow(row)
        
            # Count columns exported
            num_cols_exported = len(df_export.columns) - 1  # Subtract the empty column we added
        
            # Check if multi-day experiment was detected
            multi_day_note = ""
            if max_day_increment > 0:
                multi_day_note = f"\nMulti-day experiment detected ({max_day_increment + 1} days) - Sorted correctly by continuous time"
        
            messagebox.showinfo("Success", 
                f"Data saved to {filename}\n\n"
                f"Export mode: {mode_text}\n"
                f"Format: TSV (tab-separated) with .txt extension\n"
                f"Decimal separator: Comma (,)\n"
                f"Sorting: By continuous time (handles day rollover){multi_day_note}\n"
                f"Total rows: {len(df)}\n"
                f"Channels: {len(df['Channel'].unique())}\n"
                f"Columns exported: {num_cols_exported}\n\n"
                f"File structure:\n"
                f"  Row 1: Empty in A1, Date in B1, Time in C1, Modified file info in D1, Mode in E1\n"
                f"  Row 2: Empty\n"
                f"  Row 3: Column headers (starting with empty column A)\n"
                f"  Row 4: Units\n"
                f"  Row 5+: Data with empty first column")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {str(e)}")
            traceback.print_exc()

    def _perform_export(self, df, filename, export_format):
        """Perform the actual export with selected format"""
        try:
            now = datetime.now()
            date_str = now.strftime("%Y-%m-%d")
            time_str = now.strftime("%H:%M:%S")
        
            original_filename = self.file_path_var.get() if hasattr(self, 'file_path_var') else "Unknown"
            modified_text = f"Modified file: {os.path.basename(original_filename)}"
        
            # Create a copy for export
            df_export = df.copy()
        
            if export_format == "reimport":
                # For re-import: sort by time, then by channel
                df_export = df_export.sort_values(['time_repeated', 'Channel'])
        
            # Replace dot with comma in numeric values (if comma decimal separator is preferred)
            # Convert numeric columns to string with comma as decimal separator
            for col in df_export.columns:
                if df_export[col].dtype in ['float64', 'float32', 'int64', 'int32']:
                    # Convert to string with comma as decimal separator
                    df_export[col] = df_export[col].apply(
                        lambda x: str(x).replace('.', ',') if pd.notna(x) else ''
                    )
        
            # Create units row
            units_row = []
            for col in df.columns:
                unit = UNITS_MAP.get(col, "")
                unit = unit.replace('₂', '2').replace('₃', '3').replace('ₒ', 'o')
                unit = unit.replace('µ', 'u').replace('μ', 'u')
                units_row.append(unit)
        
            # Add empty column at the beginning
            headers = df_export.columns.tolist()
            headers.insert(0, "")
            units_row.insert(0, "")
        
            # Add empty column to dataframe
            df_export.insert(0, "", "")
        
            # Determine file format based on extension
            if filename.endswith('.xlsx'):
                # Excel format - always use dot as decimal
                # Revert back to dot for Excel
                for col in df_export.columns:
                    if col != "":
                        df_export[col] = df_export[col].apply(
                            lambda x: str(x).replace(',', '.') if isinstance(x, str) and ',' in x else x
                        )
            
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    worksheet = writer.book.active
                    worksheet.title = 'Data'
                
                    worksheet.cell(row=1, column=1, value="")
                    worksheet.cell(row=1, column=2, value=f"Date: {date_str}")
                    worksheet.cell(row=1, column=3, value=f"Time: {time_str}")
                    worksheet.cell(row=1, column=4, value=modified_text)
                    worksheet.cell(row=1, column=5, value=f"Format: {export_format}")
                
                    # Row 2: Empty row
                
                    # Row 3: Headers
                    for col_idx, header in enumerate(headers, 1):
                        worksheet.cell(row=3, column=col_idx, value=header)
                
                    # Row 4: Units
                    for col_idx, unit in enumerate(units_row, 1):
                        cell = worksheet.cell(row=4, column=col_idx, value=unit)
                        cell.font = Font(italic=True, color='666666')
                
                    # Write data starting from row 5
                    for r_idx, row in enumerate(df_export.values, 5):
                        for c_idx, value in enumerate(row, 1):
                            worksheet.cell(row=r_idx, column=c_idx, value=value)
                
                    # Auto-adjust column widths
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if cell.value and len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                
                    worksheet.freeze_panes = 'B5'
        
            elif filename.endswith('.tsv'):
                # TSV format - tab-separated with comma as decimal separator
                with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
                    writer = csv.writer(f, delimiter='\t')
                
                    writer.writerow(["", f"Date: {date_str}", f"Time: {time_str}", modified_text, f"Format: {export_format}"])
                    writer.writerow([])
                    writer.writerow(headers)
                    writer.writerow(units_row)
                
                    for _, row in df_export.iterrows():
                        writer.writerow(row)
        
            else:
                # CSV format - comma-separated with comma as decimal separator
                with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
                    writer = csv.writer(f, delimiter=',')
                
                    writer.writerow(["", f"Date: {date_str}", f"Time: {time_str}", modified_text, f"Format: {export_format}"])
                    writer.writerow([])
                    writer.writerow(headers)
                    writer.writerow(units_row)
                
                    for _, row in df_export.iterrows():
                        writer.writerow(row)
        
            # Update the download_data method to include TSV option
            messagebox.showinfo("Success", 
                f"Data saved to {filename}\n\n"
                f"Format: {'Re-import ready (time-sorted)' if export_format == 'reimport' else 'Viewing (channel-grouped)'}\n"
                f"Decimal separator: Comma (,)\n"
                f"File type: {'TSV (tab-separated)' if filename.endswith('.tsv') else 'CSV (comma-separated)' if filename.endswith('.csv') else 'Excel'}\n"
                f"Total rows: {len(df)}\n"
                f"Channels: {len(df['Channel'].unique())}\n"
                f"Columns: {len(df.columns)}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {str(e)}")
            traceback.print_exc()
    
    def download_plot(self, fig, title):
        filename = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG files", "*.png"), ("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                fig.savefig(filename, dpi=300, bbox_inches='tight')
                messagebox.showinfo("Success", f"Plot saved to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save plot: {str(e)}")
    
    def download_wuei_plot(self):
        if hasattr(self, 'wuei_fig'):
            self.download_plot(self.wuei_fig, "WUEi Boxplot")
    
    def clear_old_figures(self):
        # Close all matplotlib figures
        plt.close('all')
        
        # Force garbage collection to free memory
        gc.collect()

# ============================================
# MAIN ENTRY POINT
# ============================================

def main():
    print("\n" + "=" * 60)
    print("GAS EXCHANGE ANALYZER")
    print("=" * 60)
    
    if not run_comprehensive_dependency_check():
        print("\nDependency check failed. Exiting...")
        return
    
    # Show splash screen
    try:
        splash = SplashScreen()
    except Exception as e:
        print(f"Warning: Could not show splash screen: {e}")
        # Continue anyway
    
    try:
        root = tk.Tk()
        app = GasExchangeApp(root)
        
        # Try to set icon
        try:
            root.iconbitmap("favicon32.ico")
        except:
            pass  # Icon is optional
            
        root.mainloop()
    except Exception as e:
        print(f"\nFATAL ERROR: {e}")
        traceback.print_exc()
        
        try:
            input("\nPress Enter to exit...")
        except:
            pass

if __name__ == "__main__":
    main()
