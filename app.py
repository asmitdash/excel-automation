import streamlit as st
import pandas as pd
import io
import zipfile

# --- CONFIGURATION ---
st.set_page_config(page_title="Town Scorecard Generator", layout="wide")

def generate_town_excel(town_name, town_df):
    """
    Generates an Excel file with:
    1. 'Scorecard': Exact replica of Azamgarh template (Formatted).
    2. 'Network_Plan': The bottom 2 tables (Summary & Intervention).
    """
    
    # ==========================================
    # DATA PREPARATION & CLEANING
    # ==========================================
    
    # 1. Filter out 'Closed' stores for the Scorecard counts
    # We look for the specific column indicating closed status
    scorecard_df = town_df.copy()
    
    # Robustly find the 'Sub-Location' column even if headers vary slightly
    closed_col = next((c for c in scorecard_df.columns if 'Highlight' in c and 'closed' in c), None)
    
    # If the column exists, exclude rows containing "closed" (case insensitive)
    if closed_col:
        scorecard_df = scorecard_df[~scorecard_df[closed_col].astype(str).str.contains('closed', case=False, na=False)]

    # 2. Prepare Data Rows
    scorecard_rows = []
    stratifications = [x for x in scorecard_df['Updated Stratification'].dropna().unique()]

    for strat in stratifications:
        strat_df = scorecard_df[scorecard_df['Updated Stratification'] == strat]

        # --- COUNTS (Only Active Stores) ---
        bal_pri = len(strat_df[strat_df['BAL Store Type'].astype(str).str.contains('MD|Dealer', case=False, na=False)])
        bal_sec = len(strat_df[strat_df['BAL Store Type'].astype(str).str.contains('ASD|AD|Sub|Rep', case=False, na=False)])
        tvs_pri = len(strat_df[strat_df['TVS Store Type'].astype(str).str.contains('MD|Dealer', case=False, na=False)])
        tvs_sec = len(strat_df[strat_df['TVS Store Type'].astype(str).str.contains('ASD|AD|Sub', case=False, na=False)])
        
        # --- VOLUMES (Sum of S1 Ind, BAL, TVS) ---
        ind_s1 = pd.to_numeric(strat_df['S1 Ind Vistaar'], errors='coerce').sum()
        bal_s1_vol = pd.to_numeric(strat_df['BAL S1 Vol - Vistaar'], errors='coerce').sum()
        tvs_s1_vol = pd.to_numeric(strat_df['TVS S1 Vol'], errors='coerce').sum()
        
        # --- METRICS ---
        bal_ms = (bal_s1_vol / ind_s1) if ind_s1 > 0 else 0
        vol_gap = tvs_s1_vol - bal_s1_vol # Matches Header: Vol Gap (TVS-BAL)
        cr_val = pd.to_numeric(strat_df['CR'], errors='coerce').mean()

        # --- GAPS ---
        # Store Gap: If BAL < TVS, it's a gap. 
        # Usually represented as (BAL - TVS) -> Negative means gap.
        # Or (TVS - BAL) -> Positive means gap. 
        # Azamgarh template has "Store Gap". I will use BAL - TVS.
        store_gap_pri = bal_pri - tvs_pri
        store_gap_sec = bal_sec - tvs_sec
        
        # --- BUILD ROWS FOR THIS STRATIFICATION ---
        # Row 1: Primary
        scorecard_rows.append([
            strat, "Pri Store", bal_pri, tvs_pri, "", 
            store_gap_pri, 0, ind_s1, bal_s1_vol, bal_ms, tvs_s1_vol, vol_gap, cr_val, 
            "", "", "", "", bal_pri, 0
        ])
        # Row 2: Secondary
        scorecard_rows.append([
            "", "ASD", bal_sec, "", tvs_sec, 
            store_gap_sec, 0, "", "", "", "", "", "", 
            "", "", "", "", bal_sec, 0
        ])
        # Row 3: Vacant
        scorecard_rows.append([
            "", "Vacant", 0, "", "", 
            0, 0, "", "", "", "", "", "", 
            "", "", "", "", 0, 0
        ])
        # Row 4: Total
        scorecard_rows.append([
            "", "Total", (bal_pri + bal_sec), (tvs_pri + tvs_sec), "", 
            (store_gap_pri + store_gap_sec), 0, "", "", "", "", "", "", 
            "", "", "", "", (bal_pri + bal_sec), 0
        ])
        # Spacer Row
        scorecard_rows.append([""] * 19)

    df_scorecard = pd.DataFrame(scorecard_rows)

    # ==========================================
    # SHEET 2: NETWORK PLAN LOGIC
    # ==========================================
    # Logic for Bottom 2 Tables (Summary & Intervention)
    
    # TABLE 1: PRE vs POST SUMMARY
    summary_rows = []
    categories = ['Primary', 'Secondary', 'Vacant']
    
    total_pre_count = 0
    total_post_count = 0
    total_vol_gain = 0
    
    for cat in categories:
        # Pre
        pre_mask = town_df['Pre - Network - BAL'].astype(str).str.contains(cat, case=False, na=False)
        pre_count = len(town_df[pre_mask])
        pre_ind = pd.to_numeric(town_df.loc[pre_mask, 'S1 Ind Vistaar'], errors='coerce').sum()
        
        # Post
        post_mask = town_df['Post Net Bal'].astype(str).str.contains(cat, case=False, na=False)
        post_count = len(town_df[post_mask])
        post_ind = pd.to_numeric(town_df.loc[post_mask, 'S1 Ind Vistaar'], errors='coerce').sum()
        
        # Gains
        ind_cov_gain = post_ind - pre_ind
        vol_gain = pd.to_numeric(town_df.loc[post_mask, 'Vol Gain'], errors='coerce').sum()
        ms_gain = (vol_gain / ind_cov_gain) if ind_cov_gain > 0 else 0
        
        total_pre_count += pre_count
        total_post_count += post_count
        total_vol_gain += vol_gain
        
        summary_rows.append([cat, pre_count, pre_ind, post_count, post_ind, ind_cov_gain, vol_gain, ms_gain])

    summary_rows.append(["Total", total_pre_count, "", total_post_count, "", "", total_vol_gain, ""])
    
    df_network_summary = pd.DataFrame(summary_rows, columns=[
        "Channel Type", "Pre Count", "Pre Ind", "Post Count", "Post Ind", "Ind Coverage Gain", "Vol Gain", "MS Gain"
    ])

    # TABLE 2: INTERVENTION LIST
    intervention_df = town_df[town_df['Network Intervention'].notna()].copy()
    cols_to_keep = ['Location / T2T', 'Updated Stratification', 'Channel Type - TVS', 'BAL Store Type', 'S1 Ind Vistaar', 'Nature of Intervention', 'Network Intervention', 'Remarks']
    # Use only columns that actually exist
    existing_cols = [c for c in cols_to_keep if c in intervention_df.columns]
    df_intervention_list = intervention_df[existing_cols]

    # ==========================================
    # EXCEL WRITING (XlsxWriter Engine)
    # ==========================================
    output = io.BytesIO()
    workbook = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook_obj = workbook.book

    # --- STYLES ---
    header_format = workbook_obj.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        'fg_color': '#D9E1F2', 'border': 1
    })
    subheader_format = workbook_obj.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        'fg_color': '#D9E1F2', 'border': 1, 'font_size': 9
    })
    simple_header = workbook_obj.add_format({
        'bold': True, 'bg_color': '#FFFF00', 'border': 1
    })
    
    # --- SHEET 1: SCORECARD ---
    # We write data starting at Row 3 (Index 3), leaving room for headers
    df_scorecard.to_excel(workbook, sheet_name='Scorecard', startrow=3, header=False, index=False)
    ws1 = workbook.sheets['Scorecard']
    
    # 1. Town Name (Top Left)
    ws1.write(0, 1, town_name, workbook_obj.add_format({'bold': True, 'font_size': 14}))
    
    # 2. Main Headers (Row 2 / Index 1)
    ws1.write(1, 0, "Stratification", header_format)
    ws1.write(1, 1, "# Store Count", header_format)
    ws1.write(1, 2, "BAL", header_format)
    ws1.merge_range(1, 3, 1, 4, "TVS", header_format) # Merged TVS
    ws1.write(1, 5, "Store Gap", header_format)
    ws1.write(1, 6, "Unique Location Gap", header_format)
    ws1.write(1, 7, "IND S1", header_format)
    ws1.write(1, 8, "S1 BAL Vol", header_format)
    ws1.write(1, 9, "BAL MS", header_format)
    ws1.write(1, 10, "S1 TVS Vol", header_format)
    ws1.write(1, 11, "Vol Gap\n(TVS-BAL)", header_format)
    ws1.write(1, 12, "CR", header_format)
    ws1.merge_range(1, 13, 1, 14, "Addition", header_format) # Merged Addition
    ws1.merge_range(1, 15, 1, 16, "Reduction", header_format) # Merged Reduction
    ws1.write(1, 17, "BAL Network Count\n@ UP 2.0", header_format)
    ws1.write(1, 18, "Unique Location Gap\npost appointment", header_format)

    # 3. Sub-Headers (Row 3 / Index 2)
    for col in [0, 1, 2, 5, 6, 7, 8, 9, 10, 11, 12, 17, 18]:
        ws1.write(2, col, "", subheader_format) # Blank subheaders for non-merged cols
        
    ws1.write(2, 3, "Primary", subheader_format)
    ws1.write(2, 4, "Secondary", subheader_format)
    ws1.write(2, 13, "Primary", subheader_format)
    ws1.write(2, 14, "Secondary", subheader_format)
    ws1.write(2, 15, "Primary", subheader_format)
    ws1.write(2, 16, "Secondary", subheader_format)

    # Column Widths
    ws1.set_column(0, 0, 15)
    ws1.set_column(7, 12, 10)

    # --- SHEET 2: NETWORK PLAN ---
    ws2_name = 'Network_Plan'
    df_network_summary.to_excel(workbook, sheet_name=ws2_name, startrow=1, index=False)
    ws2 = workbook.sheets[ws2_name]
    
    # Headers for Table 1
    for col_num, value in enumerate(df_network_summary.columns.values):
        ws2.write(0, col_num, value, simple_header)

    # Intervention List (Table 2)
    start_row_t2 = len(df_network_summary) + 4
    ws2.write(start_row_t2 - 1, 0, "Intervention List", workbook_obj.add_format({'bold': True, 'font_size': 12}))
    df_intervention_list.to_excel(workbook, sheet_name=ws2_name, startrow=start_row_t2, index=False)
    
    # Headers for Table 2
    for col_num, value in enumerate(df_intervention_list.columns.values):
        ws2.write(start_row_t2, col_num, value, simple_header)
        
    ws2.set_column(0, 8, 15)

    workbook.close()
    return output.getvalue()


# --- MAIN APP UI ---
st.title("📊 Master Scorecard Generator (Gold Standard)")
st.markdown("Generates exact Azamgarh-style formatting with corrected data logic.")

uploaded_file = st.file_uploader("Upload 'Accessibility Excel' (CSV or Excel)", type=['xlsx', 'csv'])

if uploaded_file:
    try:
        # Load File
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=2)
        else:
            df = pd.read_excel(uploaded_file, header=2)
            
        if len(df) > 0:
            df = df.drop(0).reset_index(drop=True)

        # 1. CLEAN HEADERS (Handle Newlines)
        df.columns = df.columns.str.replace('\n', ' ').str.strip()
        
        # 2. FILTER TOTALS
        clean_df = df[~df['Town'].astype(str).str.contains('Total', case=False, na=False)].copy()
        
        unique_towns = clean_df['Town'].dropna().unique()
        st.success(f"✅ Data Loaded. Found {len(unique_towns)} unique towns.")

        if st.button("Generate Scorecards"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                progress_bar = st.progress(0)
                for i, town in enumerate(unique_towns):
                    town_df = clean_df[clean_df['Town'] == town]
                    excel_data = generate_town_excel(town, town_df)
                    zip_file.writestr(f"Scorecard_{town}.xlsx", excel_data)
                    progress_bar.progress((i + 1) / len(unique_towns))
            
            st.download_button(
                label="📥 Download All Scorecards (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Town_Scorecards_Gold.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"Error: {e}")
