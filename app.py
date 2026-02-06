import streamlit as st
import pandas as pd
import io
import zipfile

# --- CONFIGURATION ---
st.set_page_config(page_title="Town Scorecard Generator", layout="wide")

def generate_town_excel(town_name, town_df):

    scorecard_rows = []

    # ----------------------------
    # CLEAN CLOSED STORES
    # ----------------------------
    clean_df = town_df.copy()
    closed_col = next((c for c in clean_df.columns if 'Highlight' in c and 'closed' in c.lower()), None)
    if closed_col:
        clean_df = clean_df[
            ~clean_df[closed_col].astype(str).str.contains('closed', case=False, na=False)
        ]

    stratifications = clean_df['Updated Stratification'].dropna().unique()

    # ----------------------------
    # METRICS FUNCTION (FIXED)
    # ----------------------------
    def get_metrics(df):
        bal_count = len(df)

        tvs_p = len(df[df['TVS Store Type'].astype(str).str.contains('MD|Dealer', case=False, na=False)])
        tvs_s = len(df[df['TVS Store Type'].astype(str).str.contains('ASD|AD|Sub', case=False, na=False)])

        ind = pd.to_numeric(df['S1 Ind Vistaar'], errors='coerce').sum()
        bal_vol = pd.to_numeric(df['BAL S1 Vol - Vistaar'], errors='coerce').sum()
        tvs_vol = pd.to_numeric(df['TVS S1 Vol'], errors='coerce').sum()

        ms = (bal_vol / ind) if ind > 0 else ""
        vol_gap = bal_vol - tvs_vol
        cr = (tvs_vol / bal_vol) if bal_vol > 0 else ""

        return bal_count, tvs_p, tvs_s, ind, bal_vol, tvs_vol, ms, vol_gap, cr

    # ----------------------------
    # STRAT LOOP
    # ----------------------------
    for strat in stratifications:

        strat_df = clean_df[clean_df['Updated Stratification'] == strat]

        pri_mask = strat_df['BAL Store Type'].astype(str).str.contains('MD|Dealer', case=False, na=False)
        sec_mask = strat_df['BAL Store Type'].astype(str).str.contains('ASD|AD|Sub|Rep', case=False, na=False)

        pri_df = strat_df[pri_mask]
        sec_df = strat_df[sec_mask]
        vac_df = strat_df[~(pri_mask | sec_mask)]

        # --- PRIMARY ---
        p_bal, p_tvs_p, p_tvs_s, p_ind, p_bal_vol, p_tvs_vol, p_ms, p_vg, p_cr = get_metrics(pri_df)
        p_store_gap = p_bal - (p_tvs_p + p_tvs_s)

        # --- SECONDARY ---
        s_bal, s_tvs_p, s_tvs_s, s_ind, s_bal_vol, s_tvs_vol, s_ms, s_vg, s_cr = get_metrics(sec_df)
        s_store_gap = s_bal - (s_tvs_p + s_tvs_s)

        # --- VACANT (FIXED LOGIC) ---
        v_unique_locations = vac_df['Location / T2T'].nunique()
        v_ind = pd.to_numeric(vac_df['S1 Ind Vistaar'], errors='coerce').sum()

        # --- TOTALS ---
        t_bal = p_bal + s_bal
        t_tvs = (p_tvs_p + p_tvs_s) + (s_tvs_p + s_tvs_s)
        t_gap = p_store_gap + s_store_gap

        # --- ROW BUILD ---
        scorecard_rows.append([
            strat, "Pri Store", p_bal, p_tvs_p, "",
            p_store_gap, 0, p_ind, p_bal_vol, p_ms,
            p_tvs_vol, p_vg, p_cr,
            "", "", "", "", t_bal, 0
        ])

        scorecard_rows.append([
            "", "ASD", s_bal, "", s_tvs_s,
            s_store_gap, 0, s_ind, s_bal_vol, s_ms,
            s_tvs_vol, s_vg, s_cr,
            "", "", "", "", t_bal, 0
        ])

        scorecard_rows.append([
            "", "Vacant", 0, "", "",
            0, v_unique_locations, v_ind, 0, "",
            0, 0, "",
            "", "", "", "", 0, v_unique_locations
        ])

        scorecard_rows.append([
            "", "Total", t_bal, t_tvs, "",
            t_gap, v_unique_locations,
            (p_ind + s_ind + v_ind),
            (p_bal_vol + s_bal_vol),
            "",
            (p_tvs_vol + s_tvs_vol),
            (p_vg + s_vg),
            "",
            "", "", "", "", t_bal, v_unique_locations
        ])

        scorecard_rows.append([""] * 19)

    df_scorecard = pd.DataFrame(scorecard_rows)

    # ==========================
    # NETWORK PLAN (UNCHANGED)
    # ==========================
    summary_rows = []
    categories = ['Primary', 'Secondary', 'Vacant']

    total_pre_count = total_post_count = total_vol_gain = 0

    for cat in categories:
        pre_mask = town_df['Pre - Network - BAL'].astype(str).str.contains(cat, case=False, na=False)
        post_mask = town_df['Post Net Bal'].astype(str).str.contains(cat, case=False, na=False)

        pre_count = len(town_df[pre_mask])
        post_count = len(town_df[post_mask])

        pre_ind = pd.to_numeric(town_df.loc[pre_mask, 'S1 Ind Vistaar'], errors='coerce').sum()
        post_ind = pd.to_numeric(town_df.loc[post_mask, 'S1 Ind Vistaar'], errors='coerce').sum()

        ind_gain = post_ind - pre_ind
        vol_gain = pd.to_numeric(town_df.loc[post_mask, 'Vol Gain'], errors='coerce').sum()
        ms_gain = (vol_gain / ind_gain) if ind_gain > 0 else ""

        total_pre_count += pre_count
        total_post_count += post_count
        total_vol_gain += vol_gain

        summary_rows.append([cat, pre_count, pre_ind, post_count, post_ind, ind_gain, vol_gain, ms_gain])

    summary_rows.append(["Total", total_pre_count, "", total_post_count, "", "", total_vol_gain, ""])

    df_network_summary = pd.DataFrame(summary_rows, columns=[
        "Channel Type", "Pre Count", "Pre Ind", "Post Count", "Post Ind",
        "Ind Coverage Gain", "Vol Gain", "MS Gain"
    ])

    intervention_df = town_df[town_df['Network Intervention'].notna()].copy()
    cols = [
        'Location / T2T', 'Updated Stratification', 'Channel Type - TVS',
        'BAL Store Type', 'S1 Ind Vistaar',
        'Nature of Intervention', 'Network Intervention', 'Remarks'
    ]
    df_intervention_list = intervention_df[[c for c in cols if c in intervention_df.columns]]

    # ==========================
    # EXCEL WRITE
    # ==========================
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    wb = writer.book

    header = wb.add_format({'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#D9E1F2'})
    subheader = wb.add_format({'align': 'center', 'border': 1})
    yellow = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#FFFF00'})

    df_scorecard.to_excel(writer, sheet_name='Scorecard', startrow=3, index=False, header=False)
    ws1 = writer.sheets['Scorecard']
    ws1.write(0, 1, town_name, wb.add_format({'bold': True, 'font_size': 14}))

    headers = [
        "Stratification", "# Store Count", "BAL", "TVS", "TVS",
        "Store Gap", "Unique Location Gap", "IND S1", "S1 BAL Vol",
        "BAL MS", "S1 TVS Vol", "Vol Gap\n(TVS-BAL)", "CR",
        "Addition", "Addition", "Reduction", "Reduction",
        "BAL Network Count\n@ UP 2.0", "Unique Location Gap\npost appointment"
    ]

    for i, h in enumerate(headers):
        ws1.write(1, i, h, header)

    ws1.write(2, 3, "Primary", subheader)
    ws1.write(2, 4, "Secondary", subheader)
    ws1.write(2, 13, "Primary", subheader)
    ws1.write(2, 14, "Secondary", subheader)
    ws1.write(2, 15, "Primary", subheader)
    ws1.write(2, 16, "Secondary", subheader)

    df_network_summary.to_excel(writer, sheet_name='Network_Plan', startrow=1, index=False)
    ws2 = writer.sheets['Network_Plan']
    for c, col in enumerate(df_network_summary.columns):
        ws2.write(0, c, col, yellow)

    start = len(df_network_summary) + 4
    ws2.write(start - 1, 0, "Intervention List", wb.add_format({'bold': True}))
    df_intervention_list.to_excel(writer, sheet_name='Network_Plan', startrow=start, index=False)
    for c, col in enumerate(df_intervention_list.columns):
        ws2.write(start, c, col, yellow)

    writer.close()
    return output.getvalue()


# ==========================
# STREAMLIT UI
# ==========================
st.title("📊 Master Scorecard Generator (Audit Compliant)")

uploaded_file = st.file_uploader("Upload Accessibility Excel", type=['xlsx', 'csv'])

if uploaded_file:
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file, header=2)
    else:
        df = pd.read_excel(uploaded_file, header=2)

    df = df.drop(0).reset_index(drop=True)
    df.columns = df.columns.str.replace('\n', ' ').str.strip()

    clean_df = df[~df['Town'].astype(str).str.contains('Total', case=False, na=False)]
    towns = clean_df['Town'].dropna().unique()

    st.success(f"✅ Loaded {len(towns)} towns")

    if st.button("Generate Scorecards"):
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
            for town in towns:
                excel = generate_town_excel(town, clean_df[clean_df['Town'] == town])
                z.writestr(f"Scorecard_{town}.xlsx", excel)

        st.download_button(
            "📥 Download All Scorecards",
            zip_buffer.getvalue(),
            "Town_Scorecards_Audit_Compliant.zip",
            "application/zip"
        )
