"""
Retailer Analysis App
----------------------
A simple, professional Streamlit app that:
  1. Lets the user upload the three input files (or use the bundled defaults).
  2. Builds the full multi-sheet Excel report on pressing "Execute".
  3. Offers the generated report for download.
  4. Provides a "Username Lookup" screen that mirrors the workbook's own
     Username_Lookup dashboard, computed directly in the app.

Run with:
    pip install -r requirements.txt
    streamlit run app.py
"""

import os
from datetime import datetime

import streamlit as st

import report_builder as rb

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_POS = os.path.join(APP_DIR, "defaults", "pos_details_1784298149125.csv")
DEFAULT_BCD = os.path.join(APP_DIR, "defaults", "bcd_2026-07_51.csv")
DEFAULT_DBN = os.path.join(APP_DIR, "defaults", "Retailer_DBN_Division_Wise.xlsx")

st.set_page_config(page_title="Retailer Analysis", page_icon="📊", layout="wide")

# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
        .main > div { padding-top: 1rem; }
        .app-header {
            background: linear-gradient(90deg, #1F3864 0%, #2E75B6 100%);
            padding: 1.6rem 2rem;
            border-radius: 10px;
            color: white;
            margin-bottom: 1.5rem;
        }
        .app-header h1 { margin: 0; font-size: 1.6rem; }
        .app-header p { margin: 0.3rem 0 0 0; opacity: 0.85; font-size: 0.95rem; }
        .metric-card {
            background: #F2F2F2;
            border-radius: 8px;
            padding: 0.9rem 1rem;
            border-left: 4px solid #2E75B6;
        }
        div[data-testid="stMetricValue"] { color: #1F3864; }
        .section-title {
            background: #2E75B6;
            color: white;
            padding: 0.4rem 0.8rem;
            border-radius: 6px;
            font-weight: 600;
            margin: 0.8rem 0 0.6rem 0;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="app-header">
        <h1>📊 Retailer Analysis</h1>
        <p>Generate the Franchisee_CSC report from POS, activation, and division-wise data — and look up any retailer's monthly performance.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

tab_generate, tab_lookup = st.tabs(["⚙️  Generate Report", "🔍  Username Lookup"])

# ---------------------------------------------------------------------------
# Shared: load the bundled default files, cached so repeat visits are instant
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def _load_default_pos():
    return rb.load_pos_file(DEFAULT_POS)


@st.cache_data(show_spinner=False)
def _load_default_bcd():
    return rb.load_bcd_file(DEFAULT_BCD)


@st.cache_data(show_spinner=False)
def _load_default_dbn():
    return rb.load_dbn_file(DEFAULT_DBN)


# ---------------------------------------------------------------------------
# TAB 1: Generate Report
# ---------------------------------------------------------------------------
with tab_generate:
    st.markdown('<div class="section-title">1. Input files</div>', unsafe_allow_html=True)
    st.caption("Each file defaults to the bundled sample data. Upload your own to override any of them.")

    c1, c2, c3 = st.columns(3)
    with c1:
        pos_upload = st.file_uploader("POS Details CSV", type=["csv"], key="pos_upload")
        st.caption("Default: `pos_details_1784298149125.csv`" if pos_upload is None else f"Using uploaded file: **{pos_upload.name}**")
    with c2:
        bcd_upload = st.file_uploader("Monthly Activations CSV (BCD)", type=["csv"], key="bcd_upload")
        st.caption("Default: `bcd_2026-07_51.csv`" if bcd_upload is None else f"Using uploaded file: **{bcd_upload.name}**")
    with c3:
        dbn_upload = st.file_uploader("Retailer DBN Division Wise (XLSX)", type=["xlsx"], key="dbn_upload")
        st.caption("Default: `Retailer_DBN_Division_Wise.xlsx`" if dbn_upload is None else f"Using uploaded file: **{dbn_upload.name}**")

    st.markdown('<div class="section-title">2. Execute</div>', unsafe_allow_html=True)
    execute = st.button("▶  Execute", type="primary", use_container_width=False)

    if execute:
        try:
            with st.spinner("Reading input files..."):
                pos_data = rb.load_pos_file(pos_upload if pos_upload is not None else DEFAULT_POS)
                bcd_data = rb.load_bcd_file(bcd_upload if bcd_upload is not None else DEFAULT_BCD)
                dbn_data = rb.load_dbn_file(dbn_upload if dbn_upload is not None else DEFAULT_DBN)

            status_box = st.empty()

            def progress(msg):
                status_box.info(msg)

            with st.spinner("Building workbook..."):
                buf, stats = rb.build_workbook(pos_data, bcd_data, dbn_data, progress_callback=progress)
            status_box.empty()

            st.session_state["output_buf"] = buf.getvalue()
            st.session_state["output_stats"] = stats
            st.session_state["pos_data"] = pos_data
            st.session_state["bcd_data"] = bcd_data
            st.session_state["dbn_data"] = dbn_data

            st.success("Report generated successfully.")
        except Exception as e:
            st.error(f"Something went wrong while building the report: {e}")

    if "output_buf" in st.session_state:
        stats = st.session_state["output_stats"]
        st.markdown('<div class="section-title">3. Summary</div>', unsafe_allow_html=True)
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Active POS", f"{stats['active_pos_rows']:,}", f"of {stats['total_pos_rows']:,} total")
        m2.metric("Franchisee / CSC codes", stats["unique_franchisees"])
        m3.metric("BCD activation records", f"{stats['bcd_rows']:,}")
        m4.metric("DBN usernames matched", stats["dbn_matched_usernames"])

        st.markdown('<div class="section-title">4. Download</div>', unsafe_allow_html=True)
        fname = f"pos_details_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            "⬇  Download Output File (.xlsx)",
            data=st.session_state["output_buf"],
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption(f"{stats['sheets']} sheets — POS Details, Active_POS, Franchisee_CSC_Wise, Monthly_Activations, "
                   "Daywise_Pivot, Retailer_DBN_Division_Wise, Username_Lookup, and one sheet per Franchisee_CSC.")

# ---------------------------------------------------------------------------
# TAB 2: Username Lookup
# ---------------------------------------------------------------------------
with tab_lookup:
    st.markdown('<div class="section-title">Look up a retailer</div>', unsafe_allow_html=True)
    st.caption("Uses the files generated in the Generate Report tab if available, otherwise the bundled defaults.")

    if "pos_data" in st.session_state:
        pos_data = st.session_state["pos_data"]
        bcd_data = st.session_state["bcd_data"]
        dbn_data = st.session_state["dbn_data"]
    else:
        with st.spinner("Loading default data..."):
            pos_data = _load_default_pos()
            bcd_data = _load_default_bcd()
            dbn_data = _load_default_dbn()

    pivot, day_values = rb.build_daywise_pivot(bcd_data["rows"])
    active_usernames = sorted({(r.get("username") or "").strip() for r in pos_data["active_rows"]})

    username = st.selectbox(
        "Select or search a username",
        options=[""] + active_usernames,
        format_func=lambda u: "Type or select a username..." if u == "" else u,
    )

    if username:
        profile = rb.get_username_profile(
            username, pos_data["active_rows"], dbn_data["by_username"], pivot, day_values
        )
        if profile is None:
            st.error("Username not found among ACTIVE POS records.")
        else:
            colA, colB = st.columns(2)
            with colA:
                st.markdown('<div class="section-title">POS Profile</div>', unsafe_allow_html=True)
                for label, value in profile["pos_profile"]:
                    st.write(f"**{label}:** {value if value not in (None, '') else '—'}")
            with colB:
                st.markdown('<div class="section-title">Franchisee / Division</div>', unsafe_allow_html=True)
                if not profile["has_dbn_record"]:
                    st.info("No matching record in Retailer_DBN_Division_Wise for this username.")
                for label, value in profile["dbn_profile"]:
                    st.write(f"**{label}:** {value if value not in (None, '') else '—'}")

            st.markdown('<div class="section-title">Monthly Activation Analysis</div>', unsafe_allow_html=True)
            st.metric("Total Sims Sold (This Month)", profile["sims_sold"])

            chart_data = {"Date": profile["day_values"], "Activations": profile["day_counts"]}
            try:
                import pandas as pd
                df = pd.DataFrame(chart_data).set_index("Date")
                st.bar_chart(df)
            except ImportError:
                st.table(chart_data)
    else:
        st.info("Choose a username above to see their profile and monthly activation breakdown.")
