import streamlit as st
import pandas as pd
import datetime
import io
import json
import os
import base64
import re
import warnings
from fpdf import FPDF
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT
import pyotp
import qrcode
import time

# --- Setup Logo & Page Config ---
LOGO_URL = "https://cdn.prod.website-files.com/679a0d90717922eb8830bfb4/6ad7e88d6bb176ea1ab26c34_esentry%20white%20logo-p-2000.png"
st.set_page_config(page_title="Attendance Tracker", page_icon=LOGO_URL, layout="wide")

is_auth = st.session_state.get('authenticated', False)
page_id = "dashboard" if is_auth else "gateway"

# ── PREMIUM UI/UX CSS OVERHAUL ─────────────────────────────────────────────
st.markdown("""
<style>
/* --- Core Theme Variables --- */
:root {
    --bg-black: #000000;
    --navy-primary: #000047;
    --mint-accent: #CAEEC8;
    --mint-dim: rgba(202, 238, 200, 0.15);
    --text-main: #ffffff;
    --text-muted: #a0a0b0;
}

/* Base App Styling */
.stApp {
    background-color: var(--bg-black) !important;
    color: var(--text-main);
    font-family: 'Inter', -apple-system, sans-serif;
}
header[data-testid="stHeader"] {
    background-color: transparent !important;
}

/* Hide exact Streamlit branding */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* --- Animations --- */
@keyframes fadeInUp {
    0% { opacity: 0; transform: translateY(40px); }
    100% { opacity: 1; transform: translateY(0); }
}
@keyframes pulseGlow {
    0% { box-shadow: 0 0 5px var(--mint-dim); }
    50% { box-shadow: 0 0 20px var(--mint-dim); }
    100% { box-shadow: 0 0 5px var(--mint-dim); }
}
@keyframes slideInRight {
    from { opacity: 0; transform: translateX(30px); }
    to { opacity: 1; transform: translateX(0); }
}

/* Smoothen the entrance of main blocks */
.block-container {
    animation: fadeInUp 0.8s cubic-bezier(0.2, 0.8, 0.2, 1) forwards;
}
.stTabs {
    animation: slideInRight 0.6s ease-out forwards;
}

/* --- Interactive Elements Styling --- */

/* File Uploader */
[data-testid="stFileUploader"] {
    border: 2px dashed var(--navy-primary);
    border-radius: 16px;
    background: linear-gradient(145deg, rgba(0,0,71,0.2), rgba(0,0,0,0.8));
    padding: 20px;
    transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
}
[data-testid="stFileUploader"]:hover {
    border-color: var(--mint-accent);
    background: rgba(0,0,71,0.5);
    transform: translateY(-5px);
    box-shadow: 0 10px 30px rgba(202, 238, 200, 0.1);
}
[data-testid="stFileUploader"] small {
    color: var(--text-muted) !important;
}

/* Buttons */
.stButton>button, .stDownloadButton>button {
    background: var(--navy-primary);
    color: var(--mint-accent) !important;
    border-radius: 8px;
    border: 1px solid var(--navy-primary);
    font-weight: 600;
    letter-spacing: 0.5px;
    padding: 0.5rem 1.5rem;
    transition: all 0.3s ease;
    text-transform: uppercase;
    font-size: 0.85rem;
}
.stButton>button:hover, .stDownloadButton>button:hover {
    transform: translateY(-3px) scale(1.02);
    background: var(--mint-accent);
    color: var(--bg-black) !important;
    box-shadow: 0 8px 25px rgba(202, 238, 200, 0.25);
    border-color: var(--mint-accent);
}

/* Inputs & Dropdowns */
.stTextInput>div>div>input, .stSelectbox>div>div>div {
    background-color: rgba(0, 0, 71, 0.4) !important;
    color: var(--text-main) !important;
    border: 1px solid var(--navy-primary) !important;
    border-radius: 8px;
    transition: all 0.3s ease;
}
.stTextInput>div>div>input:focus, .stSelectbox>div>div>div:focus-within {
    border-color: var(--mint-accent) !important;
    box-shadow: 0 0 10px rgba(202, 238, 200, 0.2) !important;
}

/* Radio Buttons */
div[role="radiogroup"] > label {
    background: rgba(0,0,71,0.3);
    padding: 10px 20px;
    border-radius: 8px;
    border: 1px solid var(--navy-primary);
    transition: all 0.3s;
}
div[role="radiogroup"] > label:hover {
    border-color: var(--mint-accent);
}

/* --- Data Presentation --- */

/* DataFrames */
div[data-testid="stDataFrame"] {
    border: 1px solid var(--navy-primary);
    border-radius: 12px;
    background-color: rgba(0, 0, 71, 0.1);
    box-shadow: 0 8px 30px rgba(0,0,0,0.6) !important;
    transition: all 0.4s ease;
}
div[data-testid="stDataFrame"]:hover {
    border-color: rgba(202, 238, 200, 0.3);
    box-shadow: 0 12px 40px rgba(202, 238, 200, 0.08) !important;
}

/* Custom Tabs */
button[data-baseweb="tab"] {
    background-color: transparent !important;
    color: var(--text-muted) !important;
    font-size: 1.1rem !important;
    font-weight: 500 !important;
    padding: 1rem 2rem !important;
    border-radius: 8px 8px 0 0;
    transition: color 0.3s ease;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color: var(--mint-accent) !important;
    border-bottom: 3px solid var(--mint-accent) !important;
    background: linear-gradient(0deg, rgba(202,238,200,0.05) 0%, transparent 100%) !important;
}

/* Alerts and Information Boxes */
[data-testid="stAlert"] {
    background-color: rgba(0, 0, 71, 0.3);
    border: 1px solid var(--navy-primary);
    color: var(--text-main);
    border-radius: 12px;
}
.stProgress > div > div > div {
    background-color: var(--mint-accent) !important;
}

/* Pop-up Dashboard Cards */
.metric-card {
    background: linear-gradient(135deg, var(--navy-primary) 0%, rgba(0,0,0,1) 100%);
    border-left: 4px solid var(--mint-accent);
    padding: 20px 25px;
    border-radius: 12px;
    margin-bottom: 2rem;
    box-shadow: 0 8px 25px rgba(0,0,0,0.5);
    animation: fadeInUp 0.6s ease-out forwards;
    transition: transform 0.3s ease;
}
.metric-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 12px 30px rgba(202,238,200,0.1);
}
.metric-title { margin: 0 0 8px 0; color: var(--mint-accent); font-size: 1.3rem; font-weight: 700; letter-spacing: 0.5px;}
.metric-sub { margin: 0; color: var(--text-main); font-size: 1rem; line-height: 1.6; font-weight: 300; }

/* Custom Headings */
h1, h2, h3 { color: var(--text-main); }
h2 { border-bottom: 1px solid rgba(202,238,200,0.2); padding-bottom: 10px; margin-top: 2rem; font-weight: 300;}
</style>
""", unsafe_allow_html=True)


def get_header_html(title_text, subtitle_text=""):
    """Generates a premium, animated header component"""
    return f"""
    <div style="display: flex; align-items: center; margin-bottom: 2.5rem; background: linear-gradient(90deg, rgba(0,0,71,0.6) 0%, transparent 100%); padding: 25px 30px; border-radius: 16px; border-left: 4px solid #CAEEC8; animation: slideInRight 0.8s ease-out;">
        <img src="{LOGO_URL}" width="160" style="margin-right: 35px; filter: drop-shadow(0 4px 10px rgba(202,238,200,0.2)); flex-shrink: 0;">
        <div style="display: flex; flex-direction: column; justify-content: center;">
            <h1 style="margin: 0; padding: 0; color: #ffffff; font-weight: 800; font-size: 2.4rem; letter-spacing: -1px; text-transform: uppercase;">{title_text}</h1>
            <div style="margin-top: 10px; font-size: 1.1rem; color: #CAEEC8; font-weight: 400; letter-spacing: 1px;">{subtitle_text}</div>
        </div>
    </div>
    """

# --- Security Gateway (Google Authenticator) ---
if 'totp_secret' not in st.session_state:
    st.session_state['totp_secret'] = 'JBSWY3DPEBLW64TN'

if not st.session_state.get('authenticated', False):
    st.markdown(get_header_html("Attendance Tracker", "ATTENDACE TRACKING PORTAL"), unsafe_allow_html=True)
    
    st.markdown("""
    <div style="text-align: center; margin-top: 40px;">
        <h3 style="color: #CAEEC8; font-weight: 300;">Authentication Required</h3>
        <p style="color: #a0a0b0; margin-bottom: 30px;">Welcome, Identify with your code.</p>
    </div>
    """, unsafe_allow_html=True)
    
    totp = pyotp.TOTP(st.session_state['totp_secret'])
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("totp_form", clear_on_submit=True):
            user_code = st.text_input("Enter Passcode", type="password", placeholder="• • • • • •")
            submit_button = st.form_submit_button("Authenticate Session")
            
            if submit_button:
                if totp.verify(user_code):
                    st.session_state['authenticated'] = True
                    st.rerun()
                else:
                    st.error("Authentication failed. Invalid code provided.👀")

    st.stop()


# ── Intelligent Date Parser Engine ─────────────────────────────────────────
def robust_parse_date(series, default_dayfirst=False):
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    
    is_dayfirst = default_dayfirst
    s_clean = series.dropna().astype(str).str.strip()
    s_clean = s_clean[s_clean != '']
    
    for val in s_clean.head(1000):
        m = re.search(r'^(\d{1,2})[/-](\d{1,2})[/-]\d{2,4}', val)
        if m:
            p1, p2 = int(m.group(1)), int(m.group(2))
            if p1 > 12 and p2 <= 12:
                is_dayfirst = True
                break
            elif p1 <= 12 and p2 > 12:
                is_dayfirst = False
                break

    try:
        return pd.to_datetime(series, format='mixed', dayfirst=is_dayfirst, errors='coerce')
    except ValueError:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            return pd.to_datetime(series, infer_datetime_format=True, dayfirst=is_dayfirst, errors='coerce')


# ── MAIN DASHBOARD APPLICATION ─────────────────────────────────────────────
st.markdown(get_header_html("Attendance Tracker", "WORKSPACE INITIALIZATION & ANALYTICS"), unsafe_allow_html=True)

st.markdown(
    """<div style="background-color: rgba(0, 0, 71, 0.4); border-left: 3px solid #CAEEC8; padding: 20px 25px; border-radius: 12px; margin-bottom: 2.5rem; font-size: 1.05rem; font-weight: 300; color: #e2e8f0; line-height: 1.6; box-shadow: 0 4px 15px rgba(0,0,0,0.3);">
    Synchronize raw access control logs with organizational matrices (Staff Rosters / SOC Schedules) to instantly compute highly accurate compliance metrics, late arrival detections, and absenteeism reporting.
    </div>""", unsafe_allow_html=True
)

# ── Upload UI ──
up_col1, up_col2, up_col3 = st.columns(3)

with up_col1:
    uploaded_files = st.file_uploader(
        "Access Control Logs",
        type=['csv', 'xlsx', 'xls'],
        accept_multiple_files=True,
        help="System-generated badge-in event logs."
    )

with up_col2:
    roster_file = st.file_uploader(
        "Corporate Staff Roster",
        type=['csv', 'xlsx', 'xls'],
        accept_multiple_files=False,
        help="Standard Mon-Fri organizational alignment matrix."
    )

with up_col3:
    soc_file = st.file_uploader(
        "SOC Shift Schedule",
        type=['csv', 'xlsx', 'xls'],
        accept_multiple_files=False,
        help="24/7 rotational schedule matrix."
    )

st.markdown("<br>", unsafe_allow_html=True)

# ── Configurations ──
SOC_SHIFT_TIMES = {
    'morning':   datetime.time(7, 0, 0),
    'afternoon': datetime.time(14, 0, 0),
    'night':     datetime.time(19, 0, 0),
}
ENTRY_DOORS = {
    'main entrance ground flr',
    'main entrance upfloor'
}

# ── Reusable Component: Compact Export Action ──
def render_export_dropdown(dataframe, period_label, report_title, report_subtitle, col_widths_pdf, key_prefix):
    with st.popover("Export Output"):
        buf_xlsx = io.BytesIO()
        with pd.ExcelWriter(buf_xlsx, engine='openpyxl') as writer:
            dataframe.to_excel(writer, index=False, sheet_name=report_title[:31])
        st.download_button(
            label="Export Spreadsheet (.xlsx)",
            data=buf_xlsx.getvalue(),
            file_name=f"{key_prefix}_{period_label.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"xlsx_{key_prefix}_{period_label}"
        )

        doc = Document()
        doc.add_heading(f'{report_title} — {period_label}', level=1)
        doc.add_paragraph(report_subtitle)
        tbl = doc.add_table(rows=1, cols=len(dataframe.columns), style='Light Grid Accent 1')
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr = tbl.rows[0].cells
        for i, cn in enumerate(dataframe.columns):
            hdr[i].text = str(cn)
        for _, row in dataframe.iterrows():
            rc = tbl.add_row().cells
            for i, cn in enumerate(dataframe.columns):
                rc[i].text = str(row[cn])
        buf_docx = io.BytesIO()
        doc.save(buf_docx)
        st.download_button(
            label="Export Document (.docx)",
            data=buf_docx.getvalue(),
            file_name=f"{key_prefix}_{period_label.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"docx_{key_prefix}_{period_label}"
        )

        pdf = FPDF()
        pdf.add_page(orientation='L')
        pdf.set_font('Helvetica', 'B', 14)
        pdf.cell(0, 10, f'{report_title} - {period_label}', ln=1, align='C')
        pdf.set_font('Helvetica', '', 9)
        pdf.cell(0, 8, report_subtitle, ln=1, align='C')
        pdf.ln(5)
        pdf.set_font('Helvetica', 'B', 8)
        for i, h in enumerate(dataframe.columns):
            pdf.cell(col_widths_pdf[i] if i < len(col_widths_pdf) else 40, 8, str(h), border=1, align='C')
        pdf.ln()
        pdf.set_font('Helvetica', '', 7)
        for _, row in dataframe.iterrows():
            for i, cn in enumerate(dataframe.columns):
                w = col_widths_pdf[i] if i < len(col_widths_pdf) else 40
                val = str(row[cn])
                val = val.encode('latin-1', 'replace').decode('latin-1') 
                pdf.cell(w, 7, val, border=1, align='C' if i in [0, 2] else 'L')
            pdf.ln()
        pdf_out = pdf.output(dest='S')
        pdf_bytes = pdf_out.encode('latin-1', errors='ignore') if isinstance(pdf_out, str) else bytes(pdf_out)
        st.download_button(
            label="Export Document (.pdf)",
            data=pdf_bytes,
            file_name=f"{key_prefix}_{period_label.replace(' ', '_')}.pdf",
            mime="application/pdf",
            key=f"pdf_{key_prefix}_{period_label}"
        )

# ── Reusable Component: Time Filter ──
def period_filter(df, date_col, parsed_time_col, radio_key):
    period = st.radio("Temporal Filter:", ["All-Time", "Daily", "Weekly", "Monthly"], horizontal=True, key=radio_key)
    if period == "Daily":
        df['_date_label'] = df[parsed_time_col].dt.date
        opts = sorted(df['_date_label'].unique(), reverse=True)
        sel = st.selectbox("Specify Date", opts, key=f"{radio_key}_sel")
        return df[df['_date_label'] == sel].copy(), f"Day {sel}"
    elif period == "Weekly":
        df['_week'] = df[parsed_time_col].dt.isocalendar().year.astype(str) + '-W' + \
                      df[parsed_time_col].dt.isocalendar().week.astype(str).str.zfill(2)
        opts = sorted(df['_week'].unique(), reverse=True)
        sel = st.selectbox("Specify Week", opts, key=f"{radio_key}_sel")
        return df[df['_week'] == sel].copy(), f"Week {sel}"
    elif period == "Monthly":
        df['_month'] = df[parsed_time_col].dt.to_period('M').astype(str)
        opts = sorted(df['_month'].unique(), reverse=True)
        sel = st.selectbox("Specify Month", opts, key=f"{radio_key}_sel")
        return df[df['_month'] == sel].copy(), f"Month {sel}"
    return df.copy(), "All Time Data"


if uploaded_files:
    st.markdown("<hr style='border-color: #000047; margin: 10px 0 30px 0;'>", unsafe_allow_html=True)
    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        # ── Data Processing Core ──
        
        def get_norm_col(c):
            cl = str(c).strip().lower()
            if cl in ['time', 'date/time', 'datetime', 'date time', 'event time']: return 'Time'
            if cl in ['date']: return 'Date'
            if cl in ['door name', 'location', 'door', 'point', 'reader name']: return 'Door Name'
            if cl in ['event description', 'event', 'state', 'status', 'event type', 'description']: return 'Event Description'
            if cl in ['personnel id', 'id', 'emp no', 'employee no', 'staff id', 'card no', 'pin', 'emp pin', 'ac no.']: return 'Personnel ID'
            if cl in ['first name', 'name', 'employee name', 'staff name']: return 'First Name'
            if cl in ['last name', 'surname']: return 'Last Name'
            if cl in ['door number']: return 'Door Number'
            return str(c).strip()

        def normalize_columns(df_in):
            return df_in.rename(columns={c: get_norm_col(c) for c in df_in.columns})

        def has_required_cols(cols):
            lc = [get_norm_col(c).lower() for c in cols]
            return 'time' in lc and 'door name' in lc

        all_dfs = []
        total_files = len(uploaded_files)

        for idx, uf in enumerate(uploaded_files):
            status_text.markdown(f"<span style='color:#CAEEC8; font-weight:600;'>Processing Batch {idx + 1}/{total_files}:</span> `{uf.name}`...", unsafe_allow_html=True)
            
            if uf.name.endswith('.csv'):
                tmp = pd.read_csv(uf)
                tmp.columns = tmp.columns.astype(str).str.strip()
                tmp = normalize_columns(tmp)
                
                if 'Date' in tmp.columns and 'Time' in tmp.columns:
                    mask = tmp['Date'].notna() & tmp['Time'].notna()
                    tmp.loc[mask, 'Time'] = tmp.loc[mask, 'Date'].astype(str).str.split(' ').str[0] + ' ' + tmp.loc[mask, 'Time'].astype(str).str.split(' ').str[-1]
                
                all_dfs.append(tmp)
            else:
                if uf.name.endswith('.xls'):
                    sheets = pd.read_excel(uf, sheet_name=None, engine='xlrd')
                else:
                    sheets = pd.read_excel(uf, sheet_name=None)
                    
                for sname, d in sheets.items():
                    if d.empty:
                        continue
                    d.columns = d.columns.astype(str).str.strip()
                    
                    if not has_required_cols(d.columns):
                        found = False
                        for i in range(min(30, len(d))):
                            rv = [str(v).strip() for v in d.iloc[i].values]
                            if has_required_cols(rv):
                                d.columns = rv
                                d = d.iloc[i+1:].reset_index(drop=True)
                                found = True
                                break
                        if not found:
                            continue
                    
                    d = d.dropna(how='all').reset_index(drop=True)
                    d = normalize_columns(d)
                    
                    if 'Date' in d.columns and 'Time' in d.columns:
                        mask = d['Date'].notna() & d['Time'].notna()
                        d.loc[mask, 'Time'] = d.loc[mask, 'Date'].astype(str).str.split(' ').str[0] + ' ' + d.loc[mask, 'Time'].astype(str).str.split(' ').str[-1]

                    if 'Time' in d.columns and 'Door Name' in d.columns:
                        all_dfs.append(d)
                        
            progress_bar.progress((idx + 1) / total_files)

        status_text.markdown("✨ <strong style='color:#CAEEC8;'>Compilation Complete. Matrix Initialized.</strong>", unsafe_allow_html=True)
        time.sleep(0.6)
        status_text.empty()
        progress_bar.empty()

        if not all_dfs:
            st.error("Engine failed to locate standard chronometric structures in files.")
            st.stop()

        df = pd.concat(all_dfs, ignore_index=True)

        def clean_id(val):
            if pd.isna(val) or val == '':
                return ""
            try:
                if isinstance(val, float) and val.is_integer():
                    return str(int(val))
            except:
                pass
            val_str = str(val).strip()
            if val_str.endswith('.0') and val_str.replace('.', '').isdigit():
                val_str = val_str[:-2]
            if val_str == '0':
                return ""
            return val_str

        for col in ['Time', 'Door Name']:
            if col not in df.columns:
                st.error(f"Critical schema missing: **{col}**")
                st.stop()
                
        if 'Event Description' not in df.columns:
            df['Event Description'] = 'Normal Open'

        if 'Personnel ID' not in df.columns:
            df['Personnel ID'] = df.get('First Name', df.get('Last Name', pd.Series([""] * len(df))))

        df['Personnel ID'] = df['Personnel ID'].apply(clean_id)

        df['Parsed_Time'] = robust_parse_date(df['Time'], default_dayfirst=False)
        df = df.dropna(subset=['Parsed_Time']).copy()

        df['Event Description'] = df['Event Description'].astype(str).str.strip()
        df['Door Name'] = df['Door Name'].astype(str).str.strip()
        df['Door_Name_Norm'] = df['Door Name'].str.lower().str.strip()

        fail_keywords = 'denied|invalid|alarm|error|fail|reject|unregistered|not register|unauthorized|access denied|validate interval'
        
        df_entry = df[
            df['Door_Name_Norm'].isin(ENTRY_DOORS) &
            ~df['Event Description'].astype(str).str.lower().str.contains(fail_keywords, na=False)
        ].copy()

        df_entry = df_entry[df_entry['Personnel ID'] != ""]

        df_entry['Date']      = df_entry['Parsed_Time'].dt.date
        df_entry['Time_Only'] = df_entry['Parsed_Time'].dt.time
        df_entry['Date_dt']   = pd.to_datetime(df_entry['Date'])

        df_entry = df_entry.sort_values('Parsed_Time')
        df_first = df_entry.drop_duplicates(subset=['Date', 'Personnel ID'], keep='first').copy()

        def load_simple_excel(file_obj):
            if file_obj.name.endswith('.csv'):
                return pd.read_csv(file_obj)
            if file_obj.name.endswith('.xls'):
                return pd.read_excel(file_obj, engine='xlrd')
            return pd.read_excel(file_obj)

        df_roster = None
        soc_ids_from_roster = set()
        
        if roster_file:
            df_roster = load_simple_excel(roster_file)
            df_roster.columns = df_roster.columns.astype(str).str.strip()
            
            col_map = {}
            for c in df_roster.columns:
                cl = c.lower()
                if 'personnel' in cl or cl == 'id':
                    col_map[c] = 'Personnel ID'
                elif 'expected' in cl or 'days' in cl:
                    col_map[c] = 'Expected Days'
            df_roster = df_roster.rename(columns=col_map)

            if 'Personnel ID' not in df_roster.columns or 'Expected Days' not in df_roster.columns:
                st.warning("Master List must contain 'Personnel ID' and 'Expected Days' columns. Absentee tracking disabled.")
                df_roster = None
            else:
                df_roster['Personnel ID'] = df_roster['Personnel ID'].apply(clean_id)
                df_roster['Expected Days'] = df_roster['Expected Days'].astype(str).str.strip()
                soc_mask = df_roster['Expected Days'].str.upper() == 'SOC'
                soc_ids_from_roster = set(df_roster.loc[soc_mask, 'Personnel ID'].unique())
                
                df_std_roster = df_roster[~soc_mask].copy()
                df_std_roster['Days_Per_Week'] = pd.to_numeric(df_std_roster['Expected Days'], errors='coerce')
                df_std_roster = df_std_roster.dropna(subset=['Days_Per_Week']).copy()
                df_std_roster['Days_Per_Week'] = df_std_roster['Days_Per_Week'].astype(int)

        df_soc = None
        if soc_file:
            df_soc = load_simple_excel(soc_file)
            df_soc.columns = df_soc.columns.astype(str).str.strip()
            df_soc = normalize_columns(df_soc)
            
            for needed in ['Personnel ID', 'Date', 'Shift']:
                if needed not in df_soc.columns:
                    st.warning(f"SOC Schedule must contain a '{needed}' column. SOC tab disabled.")
                    df_soc = None
                    break
                    
            if df_soc is not None:
                df_soc['Date_parsed'] = robust_parse_date(df_soc['Date'], default_dayfirst=True)
                df_soc['Date'] = df_soc['Date_parsed'].dt.date
                df_soc = df_soc.dropna(subset=['Date']).copy()

                df_soc['Shift_Lower'] = df_soc['Shift'].astype(str).str.strip().str.lower()
                df_soc['Shift_Start'] = df_soc['Shift_Lower'].map(SOC_SHIFT_TIMES)
                df_soc['Personnel ID'] = df_soc['Personnel ID'].apply(clean_id)
                
                fn_cols = [c for c in ['First Name', 'Last Name'] if c in df_soc.columns]
                df_soc['Full Name'] = df_soc[fn_cols].fillna('').astype(str).apply(
                    lambda r: ' '.join(r).strip(), axis=1) if fn_cols else ''

        # ── OUTPUT TABS ────────────────────────────────────────────────────
        tab_std, tab_soc = st.tabs(["Corporate Staff Analytics", "SOC Operations Matrix"])

        # ════════════════════════════════════════════════════════════════════
        # TAB 1 — STANDARD STAFF
        # ════════════════════════════════════════════════════════════════════
        with tab_std:
            st.markdown("<br>", unsafe_allow_html=True)
            CUTOFF = datetime.time(8, 31, 0)
            WORKDAYS = {0, 1, 2, 3, 4}

            soc_ids = soc_ids_from_roster.copy()
            if df_soc is not None:
                soc_ids.update(df_soc['Personnel ID'].astype(str).str.strip().unique())

            df_std_first = df_first[~df_first['Personnel ID'].astype(str).str.strip().isin(soc_ids)].copy()

            # ── Late Arrivals ────────────────────────────────────────────────
            st.markdown("<h2>Late Arrival Detections</h2>", unsafe_allow_html=True)

            df_std_workday = df_std_first[df_std_first['Date_dt'].dt.dayofweek.isin(WORKDAYS)].copy()
            df_late = df_std_workday[df_std_workday['Time_Only'] >= CUTOFF].copy()

            name_cols = [c for c in ['First Name', 'Last Name'] if c in df_late.columns]
            df_late['Full Name'] = df_late[name_cols].fillna('').astype(str).apply(
                lambda r: ' '.join(r).strip(), axis=1) if name_cols else ''
            df_late['Date_Str'] = df_late['Date'].apply(lambda d: d.strftime('%Y-%m-%d'))
            df_late['Time_Str'] = df_late['Time_Only'].apply(lambda t: t.strftime('%H:%M:%S'))

            if df_late.empty:
                st.markdown(
                    '<div class="metric-card" style="border-left-color: #30d158;">'
                    '<h3 class="metric-title" style="color:#30d158;">Zero Anomalies Detected</h3>'
                    '<p class="metric-sub">100% of corporate staff authenticated prior to the 08:31 AM threshold.</p></div>',
                    unsafe_allow_html=True
                )
            else:
                st.markdown(
                    f'<div class="metric-card" style="border-left-color: #ff453a;">'
                    f'<h3 class="metric-title" style="color:#ff6b6b;">{len(df_late)} Temporal Anomalies Detected</h3>'
                    f'<p class="metric-sub">Access logs flagged at or following the <strong>08:31 AM</strong> cutoff parameter on designated workdays.</p></div>',
                    unsafe_allow_html=True
                )

                pf_col, ex_col = st.columns([4, 1])
                with pf_col:
                    df_late_f, period_label_late = period_filter(df_late.copy(), 'Date', 'Parsed_Time', 'std_late_period')

                if df_late_f.empty:
                    st.info("No anomalies recorded for the isolated timeframe.")
                else:
                    late_agg = (
                        df_late_f.sort_values(['Date', 'Time_Only'])
                        .groupby(['Personnel ID', 'Full Name'])
                        .agg(
                            Times_Late=('Date', 'nunique'),
                            Dates_Late=('Date_Str', lambda x: ', '.join(x)),
                            Times_In=('Time_Str', lambda x: ', '.join(x)),
                        )
                        .reset_index()
                        .sort_values('Times_Late', ascending=False)
                    )

                    if df_roster is not None:
                        late_agg = late_agg.merge(
                            df_std_roster[['Personnel ID', 'Days_Per_Week']],
                            on='Personnel ID', how='left'
                        )
                        late_agg['Days_Per_Week'] = late_agg['Days_Per_Week'].fillna('–')
                        late_agg = late_agg[['Personnel ID', 'Full Name', 'Days_Per_Week',
                                            'Times_Late', 'Dates_Late', 'Times_In']]
                        late_agg.columns = ['Personnel ID', 'Full Name', 'Expected/Week',
                                            'Frequency', 'Dates Logged', 'Timestamps']
                    else:
                        late_agg = late_agg[['Personnel ID', 'Full Name', 'Times_Late', 'Dates_Late', 'Times_In']]
                        late_agg.columns = ['Personnel ID', 'Full Name', 'Frequency', 'Dates Logged', 'Timestamps']

                    st.dataframe(late_agg, hide_index=True, use_container_width=True)

                    with ex_col:
                        st.write("") # Spacer
                        render_export_dropdown(
                            late_agg, period_label_late,
                            "Corporate Staff Anomalies",
                            "Entities violating the 08:31 AM threshold on standard operational days.",
                            [20, 40, 15, 15, 80, 80], "std_late"
                        )

            # ── Weekly Attendance Compliance ──────────────────────────────────
            st.markdown("<h2>Weekly Compliance Matrix</h2>", unsafe_allow_html=True)

            if df_roster is None:
                st.info("Dependency Required: Upload Corporate Staff Roster to activate structural compliance computations.")
            else:
                _name_lkp = {}
                if 'First Name' in df_first.columns:
                    _nl = df_first[['Personnel ID', 'First Name', 'Last Name']].copy()
                    _nl['Personnel ID'] = _nl['Personnel ID'].astype(str).str.strip()
                    _nc_cols = [c for c in ['First Name', 'Last Name'] if c in _nl.columns]
                    _nl['Full Name'] = _nl[_nc_cols].fillna('').astype(str).apply(
                        lambda r: ' '.join(r).strip(), axis=1)
                    _nl = _nl.drop_duplicates(subset='Personnel ID')
                    _name_lkp = dict(zip(_nl['Personnel ID'], _nl['Full Name']))

                _df_att = df_std_first[df_std_first['Date_dt'].dt.dayofweek.isin(WORKDAYS)].copy()
                _df_att['Personnel ID'] = _df_att['Personnel ID'].astype(str).str.strip()
                _df_att['ISO_Year_n'] = _df_att['Parsed_Time'].dt.isocalendar().year.astype(int)
                _df_att['ISO_Week_n'] = _df_att['Parsed_Time'].dt.isocalendar().week.astype(int)
                _df_att['Week_Label'] = (_df_att['ISO_Year_n'].astype(str) + '-W' +
                                         _df_att['ISO_Week_n'].astype(str).str.zfill(2))
                _df_att['Date_Str'] = _df_att['Date'].apply(lambda d: d.strftime('%Y-%m-%d'))
                _df_att['Time_Str'] = _df_att['Time_Only'].apply(lambda t: t.strftime('%H:%M'))

                if _df_att.empty:
                    st.info("Insufficient chronometric data mapping to standard workdays.")
                else:
                    _min_dt = _df_att['Parsed_Time'].min()
                    _max_dt = _df_att['Parsed_Time'].max()

                    _wk_starts = pd.date_range(
                        _min_dt.to_period('W').to_timestamp(),
                        _max_dt.to_period('W').to_timestamp(),
                        freq='W-MON'
                    )
                    _iso_weeks = []
                    for _ws in _wk_starts:
                        _ic = _ws.isocalendar()
                        _iso_weeks.append({
                            'year': int(_ic[0]), 'week': int(_ic[1]),
                            'label': f"{int(_ic[0])}-W{str(int(_ic[1])).zfill(2)}"
                        })
                    if not _iso_weeks:
                        _ic = _min_dt.isocalendar()
                        _iso_weeks = [{'year': int(_ic[0]), 'week': int(_ic[1]),
                                       'label': f"{int(_ic[0])}-W{str(int(_ic[1])).zfill(2)}"}]

                    _comp_rows = []
                    for _, _pr in df_std_roster.iterrows():
                        _pid  = str(_pr['Personnel ID'])
                        _exp  = int(_pr['Days_Per_Week'])
                        _fname = _name_lkp.get(_pid, '')
                        _p_att = _df_att[_df_att['Personnel ID'] == _pid]

                        for _wk in _iso_weeks:
                            try:
                                _mon = datetime.date.fromisocalendar(_wk['year'], _wk['week'], 1)
                            except Exception:
                                continue
                            
                            _wd_strs = [
                                (_mon + datetime.timedelta(days=_d)).strftime('%Y-%m-%d')
                                for _d in range(5)
                            ]
                            _wk_att = _p_att[_p_att['Week_Label'] == _wk['label']].sort_values('Date')
                            _att_strs = set(_wk_att['Date_Str'].tolist())
                            
                            _att_entries = [
                                f"{_r['Date_Str']} ({_r['Time_Str']})"
                                for _, _r in _wk_att.iterrows()
                            ]
                            _absent = [_d for _d in _wd_strs if _d not in _att_strs]
                            _actual  = len(_att_strs)
                            _deficit = max(0, _exp - _actual)
                            _comp_rows.append({
                                'Personnel ID':         _pid,
                                'Full Name':            _fname,
                                'Week':                 _wk['label'],
                                'Expected Days':        _exp,
                                'Days Present':         _actual,
                                'Deficit':              _deficit,
                                'Status':               '✅ Verified' if _deficit == 0 else f'⚠️ Deficit: {_deficit}',
                                'Verified Timestamps':  ', '.join(_att_entries) if _att_entries else '—',
                                'Absent Dates':         ', '.join(_absent) if _absent else '—',
                            })

                    _df_comp = pd.DataFrame(_comp_rows)

                    _nc_count = _df_comp[_df_comp['Deficit'] > 0]['Personnel ID'].nunique()
                    if _nc_count == 0:
                        st.markdown(
                            '<div class="metric-card" style="border-left-color: #30d158;">'
                            '<h3 class="metric-title" style="color:#30d158;">100% Operational Compliance</h3>'
                            '<p class="metric-sub">Structural alignment with mandatory operational days successfully verified.</p></div>',
                            unsafe_allow_html=True
                        )
                    else:
                        st.markdown(
                            f'<div class="metric-card" style="border-left-color: #ff9500;">'
                            f'<h3 class="metric-title" style="color:#ff9500;">⚠️ Deficit Detected: {_nc_count} Entities</h3>'
                            f'<p class="metric-sub">Calculated divergence between mandatory days requirement and authenticated system badges.</p></div>',
                            unsafe_allow_html=True
                        )

                    _pf2_col, _ex2_col = st.columns([4, 1])
                    _all_wk_opts = sorted(_df_comp['Week'].unique(), reverse=True)
                    with _pf2_col:
                        _comp_view = st.radio(
                            "Matrix Scope:", ["Global View", "Isolate Week"],
                            horizontal=True, key="comp_period_radio"
                        )
                    if _comp_view == "Isolate Week" and _all_wk_opts:
                        _sel_wk = st.selectbox("Specify Iso-Week", _all_wk_opts, key="comp_week_sel")
                        _detail_view = _df_comp[_df_comp['Week'] == _sel_wk].copy()
                    else:
                        _detail_view = _df_comp.copy()

                    st.markdown("<p style='color:#a0a0b0; font-size:0.9rem;'><strong>Granular Logs</strong> — Complete timeline of authenticated presence and computed absences.</p>", unsafe_allow_html=True)
                    st.dataframe(_detail_view, hide_index=True, use_container_width=True)

                    st.markdown("<p style='color:#a0a0b0; font-size:0.9rem; margin-top:20px;'><strong>Aggregated Analytics</strong> — Cross-referenced macro view of individual deficits.</p>", unsafe_allow_html=True)
                    
                    _nc_wk_list = (
                        _df_comp[_df_comp['Deficit'] > 0]
                        .groupby('Personnel ID')['Week']
                        .apply(lambda x: ', '.join(sorted(x)))
                        .reset_index()
                        .rename(columns={'Week': 'Non-Compliant Weeks'})
                    )
                    _summary = (
                        _df_comp.groupby(['Personnel ID', 'Full Name'])
                        .agg(
                            Expected_Per_Week=('Expected Days', 'first'),
                            Weeks_NC=('Deficit', lambda x: int((x > 0).sum())),
                            Total_Deficit=('Deficit', 'sum'),
                        )
                        .reset_index()
                    )
                    _summary = _summary.merge(_nc_wk_list, on='Personnel ID', how='left')
                    _summary['Non-Compliant Weeks'] = _summary['Non-Compliant Weeks'].fillna('—')
                    _summary = _summary.sort_values('Total_Deficit', ascending=False)
                    _summary.columns = [
                        'Personnel ID', 'Full Name', 'Target/Week',
                        'Weeks Flagged', 'Accumulated Deficit', 'Flagged Cycles'
                    ]
                    st.dataframe(_summary, hide_index=True, use_container_width=True)

                    with _ex2_col:
                        st.write("") # Spacer
                        render_export_dropdown(
                            _detail_view, "All Time",
                            "Structural Compliance Report",
                            "Corporate framework verification against system access logs.",
                            [15, 35, 12, 12, 12, 8, 15, 60, 60], "std_compliance"
                        )

        # ════════════════════════════════════════════════════════════════════
        # TAB 2 — SOC TEAM
        # ════════════════════════════════════════════════════════════════════
        with tab_soc:
            st.markdown("<br>", unsafe_allow_html=True)
            if df_soc is None:
                st.info("Dependency Required: Upload SOC Shift Schedule to activate continuous operations tracking.")
                st.markdown("""
                <div style="background: rgba(0,0,71,0.2); padding: 15px; border-radius: 8px; border: 1px solid #000047;">
                <p style="color:#CAEEC8; margin-top:0;"><strong>Expected Data Architecture:</strong></p>
                
                | Personnel ID | First Name | Last Name | Date | Shift |
                |---|---|---|---|---|
                | 201 | Jane | Smith | 14/04/2026 | Morning |
                | 202 | James | Okoro | 14/04/2026 | Afternoon |

                <small style="color:#a0a0b0;"><b>Shift Parameters:</b> <code>Morning</code> (07:00) · <code>Afternoon</code> (14:00) · <code>Night</code> (19:00)<br>
                <em>Note: Omit rows entirely for rest cycles (days off).</em></small>
                </div>
                """, unsafe_allow_html=True)
            else:
                _soc_pids = df_soc['Personnel ID'].unique()
                _soc_logs = df_first[df_first['Personnel ID'].astype(str).str.strip().isin(_soc_pids)].copy()
                _soc_logs['Personnel ID'] = _soc_logs['Personnel ID'].astype(str).str.strip()

                _merged = df_soc.merge(
                    _soc_logs[['Personnel ID', 'Date', 'Time_Only', 'Parsed_Time']],
                    on=['Personnel ID', 'Date'],
                    how='left'
                )

                _sc = _merged.copy()
                _sc['Date_Str']        = _sc['Date'].apply(
                    lambda d: d.strftime('%Y-%m-%d') if hasattr(d, 'strftime') else str(d))
                _sc['Sched_Start_Str'] = _sc['Shift_Start'].apply(
                    lambda t: t.strftime('%H:%M') if pd.notna(t) else '—')
                _sc['Actual_In_Str']   = _sc['Time_Only'].apply(
                    lambda t: t.strftime('%H:%M') if pd.notna(t) else '—')
                _sc['Shift_Title']     = _sc['Shift'].astype(str).str.strip().str.title()
                _sc['Date_dt_col']     = pd.to_datetime(_sc['Date_Str'])

                def _soc_status(row):
                    if pd.isna(row['Time_Only']):
                        return '🚫 Absent'
                    if pd.notna(row['Shift_Start']):
                        shift_start_dt = datetime.datetime.combine(row['Date'], row['Shift_Start'])
                        cutoff_time = (shift_start_dt + datetime.timedelta(minutes=1)).time()
                        if row['Time_Only'] >= cutoff_time:
                            return '⚠️ Late'
                    return '✅ Logged'

                def _soc_mins(row):
                    if pd.isna(row['Time_Only']) or pd.isna(row['Shift_Start']):
                        return 0
                    shift_start_dt = datetime.datetime.combine(row['Date'], row['Shift_Start'])
                    cutoff_time = (shift_start_dt + datetime.timedelta(minutes=1)).time()
                    if row['Time_Only'] >= cutoff_time:
                        return int(
                            (datetime.datetime.combine(row['Date'], row['Time_Only']) -
                             shift_start_dt).total_seconds() / 60
                        )
                    return 0

                _sc['Status']    = _sc.apply(_soc_status, axis=1)
                _sc['Mins Late'] = _sc.apply(_soc_mins, axis=1)

                _soc_comp_all = (
                    _sc[['Personnel ID', 'Full Name', 'Date_Str', 'Shift_Title',
                          'Sched_Start_Str', 'Actual_In_Str', 'Status', 'Mins Late']]
                    .rename(columns={
                        'Date_Str': 'Target Date', 'Shift_Title': 'Cycle',
                        'Sched_Start_Str': 'T-Zero (Expected)',
                        'Actual_In_Str':   'Recorded Auth'
                    })
                    .sort_values(['Personnel ID', 'Target Date'])
                )

                _late_pids = _sc[_sc['Status'] == '⚠️ Late']['Personnel ID'].nunique()
                _abs_pids  = _sc[_sc['Status'] == '🚫 Absent']['Personnel ID'].nunique()
                _banner    = []
                if _late_pids:
                    _banner.append(f"<strong style='color:#ff9500;'>{_late_pids}</strong> operators breached T-Zero")
                if _abs_pids:
                    _banner.append(f"<strong style='color:#ff453a;'>{_abs_pids}</strong> operators untraceable for shift")

                if not _banner:
                    st.markdown(
                        '<div class="metric-card" style="border-left-color: #CAEEC8;">'
                        '<h3 class="metric-title" style="color:#CAEEC8;">Flawless Rotational Integrity</h3>'
                        '<p class="metric-sub">Zero deviations detected across all continuous schedule matrices.</p></div>',
                        unsafe_allow_html=True
                    )
                else:
                    st.markdown(
                        f'<div class="metric-card" style="border-left-color: #ff453a;">'
                        f'<h3 class="metric-title" style="color:#ff6b6b;">Schedule Deviations Detected</h3>'
                        f'<p class="metric-sub">{" &nbsp;|&nbsp; ".join(_banner)}</p></div>',
                        unsafe_allow_html=True
                    )

                st.markdown("<h2>Continuous Operations Matrix</h2>", unsafe_allow_html=True)
                st.markdown(
                    "<p style='color:#a0a0b0; font-size:0.95rem;'>"
                    "<strong>Legend:</strong> ✅ Logged &nbsp;·&nbsp; ⚠️ Late (breach of T-Zero + 1 min threshold) "
                    "&nbsp;·&nbsp; 🚫 Absent (authentication void)."
                    "</p>", unsafe_allow_html=True
                )

                _sc_pf_col, _sc_ex_col = st.columns([4, 1])
                with _sc_pf_col:
                    _sc_f, _sc_period = period_filter(
                        _sc.copy(), 'Date', 'Date_dt_col', 'soc_comp_period'
                    )

                if _sc_f.empty:
                    st.info("No schedule matrices loaded for the specified temporal bracket.")
                else:
                    _sc_display = (
                        _sc_f[['Personnel ID', 'Full Name', 'Date_Str', 'Shift_Title',
                                'Sched_Start_Str', 'Actual_In_Str', 'Status', 'Mins Late']]
                        .rename(columns={
                            'Date_Str': 'Target Date', 'Shift_Title': 'Cycle',
                            'Sched_Start_Str': 'T-Zero (Expected)',
                            'Actual_In_Str':   'Recorded Auth'
                        })
                        .sort_values(['Personnel ID', 'Target Date'])
                    )
                    st.dataframe(_sc_display, hide_index=True, use_container_width=True)

                    with _sc_ex_col:
                        st.write("")
                        render_export_dropdown(
                            _sc_display, _sc_period,
                            "SOC Operations Audit",
                            "High-fidelity track of continuous shift authentication events.",
                            [18, 40, 14, 14, 16, 16, 14, 12], "soc_compliance"
                        )

                st.markdown("<h2>Operator Aggregation View</h2>", unsafe_allow_html=True)
                st.markdown("<p style='color:#a0a0b0; font-size:0.9rem;'>Consolidated performance metrics isolating tardiness duration and total absence volumes.</p>", unsafe_allow_html=True)

                _soc_summary = (
                    _soc_comp_all
                    .groupby(['Personnel ID', 'Full Name'])
                    .agg(
                        Scheduled_Shifts=('Target Date', 'count'),
                        On_Time=('Status',    lambda x: (x == '✅ Logged').sum()),
                        Late=   ('Status',    lambda x: x.str.startswith('⚠️').sum()),
                        Absent= ('Status',    lambda x: (x == '🚫 Absent').sum()),
                        Total_Mins_Late=('Mins Late', 'sum'),
                    )
                    .reset_index()
                    .sort_values(['Absent', 'Late'], ascending=False)
                )
                _soc_summary.columns = [
                    'Operator ID', 'Designation (Name)', 'Cycles Assigned',
                    'Flawless Auth', 'Breaches', 'Void Auths', 'Total Duration Late (m)'
                ]
                st.dataframe(_soc_summary, hide_index=True, use_container_width=True)

                render_export_dropdown(
                    _soc_summary, "All Time",
                    "SOC Operator Macro Analytics",
                    "Individual aggregate tracking of SOC team integrity.",
                    [18, 40, 20, 14, 14, 14, 20], "soc_summary"
                )

    except Exception as e:
        st.error(f"Critical Engine Error: {e}")
        import traceback
        st.code(traceback.format_exc())
