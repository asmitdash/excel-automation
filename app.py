import pandas as pd
import io
import xlsxwriter

def generate_summary_logic(source_file_path, output_file_path):
    """
    Generates the Town Summary Scorecard based on the Azamgarh logic.
    """
    # 1. Load Data
    # Handle potential header row issues (usually row 1 or 2)
    try:
        df = pd.read_csv(source_file_path, header=1)
        if 'Region' not in df.columns: # Fallback if header is row 2
             df = pd.read_csv(source_file_path, header=2)
    except:
        df = pd.read_csv(source_file_path, header=2) # Default assumption

    # 2. Clean Headers (Remove Newlines)
    df.columns = df.columns.str.replace('\n', ' ').str.strip()
    
    # 3. Filter Data
    # Exclude 'Total' rows if any
    df = df[~df['Town'].astype(str).str.contains('Total', case=False, na=False)]
    
    # 4. Helper Functions for Logic
    def get_bal_class(row):
        """Classify BAL Store Type into Primary, Secondary, or Vacant"""
        stype = str(row['BAL Store Type']).upper()
        # User Logic: Main dealership (MD) and branch (BR) -> Primary
        if 'MD' in stype or 'BRANCH' in stype or 'BR' == stype or 'DEALER' in stype:
            return 'Pri Store'
        # User Logic: ASD count -> Secondary
        elif 'ASD' in stype or 'AD' in stype or 'SUB' in stype or 'REP' in stype:
            return 'ASD'
        else:
            return 'Vacant'

    def get_tvs_counts(row):
        """Return (Primary Count, Secondary Count) for TVS"""
        stype = str(row['TVS Store Type']).upper()
        pri = 1 if ('MD' in stype or 'DEALER' in stype or 'BRANCH' in stype) else 0
        sec = 1 if ('ASD' in stype or 'AD' in stype or 'SUB' in stype) else 0
        return pri, sec

    # Apply Classification
    df['Row_Class'] = df.apply(get_bal_class, axis=1)
    
    # 5. Process Each Town
    summary_data = []
    
    # Get Unique Stratifications (Large Town, Small Town, etc.)
    strats = ['Large Town', 'Small Town', 'Rural', 'Deep Rural'] # Ordered list
    existing_strats = df['Updated Stratification'].dropna().unique()
    
    for strat in strats:
        if strat not in existing_strats:
            continue
            
        strat_df = df[df['Updated Stratification'] == strat]
        
        # We need 3 rows: Pri Store, ASD, Vacant
        row_types = ['Pri Store', 'ASD', 'Vacant']
        
        for rtype in row_types:
            # Filter for this row type
            row_df = strat_df[strat_df['Row_Class'] == rtype]
            
            # --- CALCULATE METRICS ---
            
            # 1. Counts
            bal_count = len(row_df) # Count of BAL stores in this category
            
            # TVS Counts (Sum of Primary and Secondary markers)
            tvs_counts = row_df.apply(get_tvs_counts, axis=1)
            tvs_pri = sum([x[0] for x in tvs_counts])
            tvs_sec = sum([x[1] for x in tvs_counts])
            tvs_total = tvs_pri + tvs_sec
            
            # 2. Volumes
            # Use 'S1 Ind - F Vistaar' (or similar column name from file)
            # Snippet shows 'S1 Ind - F Vistaar'
            ind_s1 = pd.to_numeric(row_df['S1 Ind - F Vistaar'], errors='coerce').sum()
            bal_s1_vol = pd.to_numeric(row_df['BAL S1 Vol - Vistaar'], errors='coerce').sum()
            # Snippet shows 'TVS S1 Vol Basis MS'
            tvs_s1_vol = pd.to_numeric(row_df['TVS S1 Vol Basis MS'], errors='coerce').sum()
            
            # 3. Calculated Columns
            bal_ms = (bal_s1_vol / ind_s1) if ind_s1 > 0 else 0
            vol_gap = tvs_s1_vol - bal_s1_vol
            cr = pd.to_numeric(row_df['CR'], errors='coerce').mean() # Assuming 'CR' column exists, or derive
            if 'CR' not in row_df.columns:
                 cr = (tvs_s1_vol / bal_s1_vol) if bal_s1_vol > 0 else 0
            
            # 4. Gaps
            # Store Gap: TVS Total - BAL (Positive means we are behind)
            store_gap = tvs_total - bal_count
            
            # Unique Location Gap: If Vacant, count rows. Else 0.
            unique_loc_gap = bal_count if rtype == 'Vacant' else 0 # Wait, if Vacant, bal_count is rows.
            # Actually, logic: If 'Vacant', it's the count of rows (opportunities).
            if rtype == 'Vacant':
                 unique_loc_gap = len(row_df)
            else:
                 unique_loc_gap = 0
            
            # 5. Interventions (Addition/Reduction)
            # Check 'Nature of Intervention' column
            # 'Branch' -> Add Primary
            # 'ASD' -> Add Secondary
            add_pri = len(row_df[row_df['Nature of Intervention'].astype(str).str.contains('Branch', case=False, na=False)])
            add_sec = len(row_df[row_df['Nature of Intervention'].astype(str).str.contains('ASD', case=False, na=False)])
            
            # Reduction? (Not explicitly in snippet, assume 0 for now)
            red_pri = 0
            red_sec = 0
            
            # 6. Post Network Counts
            # Current BAL + Addition - Reduction
            bal_net_count_up = bal_count + (add_pri + add_sec) - (red_pri + red_sec)
            
            # Unique Location Gap Post
            # Current Gap - Total Addition
            unique_loc_gap_post = max(0, unique_loc_gap - (add_pri + add_sec))
            
            # Append Row Data
            summary_data.append([
                strat if rtype == 'Pri Store' else "", # Stratification only on first row
                rtype, # Row Label (Pri Store, ASD, Vacant)
                bal_count, # BAL Count
                tvs_pri, # TVS Primary
                tvs_sec, # TVS Secondary
                store_gap,
                unique_loc_gap,
                ind_s1,
                bal_s1_vol,
                bal_ms,
                tvs_s1_vol,
                vol_gap,
                cr,
                add_pri,
                add_sec,
                red_pri,
                red_sec,
                bal_net_count_up,
                unique_loc_gap_post
            ])
            
        # Add Total Row for Stratification
        # (Summing up previous 3 rows)
        # Logic: Sum numeric columns, skip others
        # For simplicity in this logic generator, I'll skip calculating Total row here, 
        # but in Excel generation, it's just sum formula.
        
    # Create DataFrame
    cols = [
        'Stratification', '# Store Count', 'BAL', 'TVS Primary', 'TVS Secondary', 
        'Store Gap', 'Unique Location Gap', 'IND S1', 'S1 BAL Vol', 'BAL MS', 
        'S1 TVS Vol', 'Vol Gap (TVS-BAL)', 'CR', 'Addition Pri', 'Addition Sec', 
        'Reduction Pri', 'Reduction Sec', 'BAL Network Count @ UP 2.0', 'Unique Location Gap post appointment'
    ]
    df_summary = pd.DataFrame(summary_data, columns=cols)
    
    return df_summary

# Usage Example (commented out):
# df_result = generate_summary_logic('azamghar.xlsx - Sheet1.csv', 'Output.xlsx')
