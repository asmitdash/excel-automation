import streamlit as st
import pandas as pd
import io
import zipfile

# --- CONFIGURATION ---
st.set_page_config(page_title="Town Summary Generator", layout="wide")

def process_town_data(town_name, town_df):
    """
    Generates the specific 'Azamgarh-style' summary rows for a given town.
    """
    summary_data = []

    # --- ROW 1: Town Name Header ---
    summary_data.append([
        "", town_name, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
    ])

    # --- ROW 2: Column Headers ---
    headers = [
        "", "Stratification", "# Store Count", "BAL", "TVS", "", "Store Gap", 
        "Unique Location Gap", "IND S1", "S1 BAL Vol", "BAL MS", "S1 TVS Vol", 
        "Vol Gap (TVS-BAL)", "CR", "Addition", "", "Reduction", "", 
        "BAL Network Count @ UP 2.0", "Unique Location Gap post appointment"
    ]
    summary_data.append(headers)

    # --- ROW 3: Sub-Headers (Primary/Secondary) ---
    sub_headers = [
        "", "", "", "Primary", "Secondary", "", "", "", "", "", "", "", "", "", 
        "Primary", "Secondary", "Primary", "Secondary", "", ""
    ]
    summary_data.append(sub_headers)

    # Filter out NaNs from Stratification
    stratifications = [x for x in town_df['Updated Stratification'].dropna().unique()]

    for strat in stratifications:
        # Filter data for this specific stratification
        strat_df = town_df[town_df['Updated Stratification'] == strat]

        # --- CALCULATIONS ---
        
        # Store Counts (Logic: Check 'Store Type' columns)
        # Using exact string matching logic from your file structure
        bal_pri = len(strat_df[strat_df['BAL Store Type'].astype(str).str.contains('MD|Dealer', case=False, na=False)])
        bal_sec = len(strat_df[strat_df['BAL Store Type'].astype(str).str.contains('ASD|AD|Sub|Rep', case=False, na=False)])
        
        tvs_pri = len(strat_df[strat_df['TVS Store Type'].astype(str).str.contains('MD|Dealer', case=False, na=False)])
        tvs_sec = len(strat_df[strat_df['TVS Store Type'].astype(str).str.contains('ASD|AD|Sub', case=False, na=False)])
        
        # Volumes (Summing the columns from Accessibility File)
        # Note: We cleaned column names to remove newlines, so we use the clean versions here
        ind_s1 = pd.to_numeric(strat_df['S1 Ind Vistaar'], errors='coerce').sum()
        bal_s1_vol = pd.to_numeric(strat_df['BAL S1 Vol - Vistaar'], errors='coerce').sum()
        tvs_s1_vol = pd.to_numeric(strat_df['TVS S1 Vol'], errors='coerce').sum()
        
        # Derived Metrics
        bal_ms = (bal_s1_vol / ind_s1) if ind_s1 > 0 else 0
        vol_gap = tvs_s1_vol - bal_s1_vol
        cr_val = pd.to_numeric(strat_df['CR'], errors='coerce').mean()

        # --- BUILDING THE ROWS ---

        # Row A: Primary Store Row
        row_pri = [
            "", strat, "Pri Store", bal_pri, tvs_pri, "", 
            (bal_pri - tvs_pri), "", 
            ind_s1, bal_s1_vol, f"{bal_ms:.2%}", tvs_s1_vol, vol_gap, f"{cr_val:.2f}",
            "", "", "", "", "", "" 
        ]
        summary_data.append(row_pri)

        # Row B: Secondary (ASD) Row
        row_sec = [
            "", "", "ASD", bal_sec, tvs_sec, "",
            (bal_sec - tvs_sec), "",
            "", "", "", "", "", "-", 
            "", "", "", "", "", ""
        ]
        summary_data.append(row_sec)

        # Row C: Vacant 
        row_vac = [
            "", "", "Vacant", 0, 0, "", "", "", 
            "", "", "", "", "", "-", 
            "", "", "", "", "", ""
        ]
        summary_data.append(row_vac)

        # Row D: Total
        row_tot = [
            "", "", "Total", (bal_pri + bal_sec), (tvs_pri + tvs_sec), "", "", "",
            "", "", "", "", "", "", 
            "", "", "", "", "", ""
        ]
        summary_data.append(row_tot)
        
        # Spacer row
        summary_data.append([""] * len(headers))

    return pd.DataFrame(summary_data)


# --- MAIN APP UI ---
st.title("📊 Automator: Town Summary Generator (Fixed)")

uploaded_file = st.file_uploader("Upload 'Accessibility Excel' (CSV or Excel)", type=['xlsx', 'csv'])

if uploaded_file:
    try:
        # Load Data
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=2)
        else:
            df = pd.read_excel(uploaded_file, header=2)
            
        if len(df) > 0:
            df = df.drop(0).reset_index(drop=True)

        # --- CRITICAL FIX: Clean Column Names ---
        # Replaces newlines '\n' with spaces so 'S1 Ind\nVistaar' becomes 'S1 Ind Vistaar'
        df.columns = df.columns.str.replace('\n', ' ').str.strip()

        # Clean 'Total' rows
        clean_df = df[~df['Town'].astype(str).str.contains('Total', case=False, na=False)].copy()
        
        unique_towns = clean_df['Town'].dropna().unique()
        st.success(f"✅ Data Loaded. Found {len(unique_towns)} unique towns.")

        if st.button("Generate & Download Summaries"):
            
            # Create a ZIP file in memory
            zip_buffer = io.BytesIO()
            
            # --- CRITICAL FIX: Removed invalid 'false_compress' argument ---
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                progress_bar = st.progress(0)
                
                for i, town in enumerate(unique_towns):
                    town_summary_df = process_town_data(town, clean_df)
                    
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        town_summary_df.to_excel(writer, index=False, header=False)
                    
                    file_name = f"Summary_{town}.xlsx"
                    zip_file.writestr(file_name, excel_buffer.getvalue())
                    
                    progress_bar.progress((i + 1) / len(unique_towns))
            
            st.download_button(
                label="📥 Download All Files (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="All_Town_Summaries.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"Error processing file: {e}")