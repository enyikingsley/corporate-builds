import os
import sys
import io
import re
import time
import shutil
import platform
import subprocess
import urllib.request
import datetime
import pandas as pd
import streamlit as st
from fpdf import FPDF
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT

# ==========================================
# 0. RUNTIME CONTEXT CHECK & LAUNCHER (CLOUDFLARE)
# ==========================================
try:
    from streamlit.runtime.scriptrunner import get_script_run_ctx
    RUNNING_IN_STREAMLIT = get_script_run_ctx() is not None
except Exception:
    RUNNING_IN_STREAMLIT = False

TUNNEL_CACHE_FILE = ".cf_tunnel_url"

def get_cloudflared_binary():
    if shutil.which("cloudflared"):
        return "cloudflared"
    binary_name = "cloudflared.exe" if sys.platform.startswith("win") else "cloudflared"
    local_binary = os.path.abspath(binary_name)
    if os.path.exists(local_binary):
        return local_binary

    print("⚡ cloudflared not found. Downloading standalone Cloudflare Tunnel binary...")
    machine = platform.machine().lower()
    is_arm = "arm" in machine or "aarch64" in machine
    if sys.platform.startswith("linux"):
        arch = "arm64" if is_arm else "amd64"
        url = f"https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-{arch}"
    elif sys.platform == "darwin":
        arch = "arm64" if is_arm else "amd64"
        url = f"https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-darwin-{arch}"
    elif sys.platform.startswith("win"):
        url = "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe"
    else:
        raise RuntimeError(f"Unsupported OS: {sys.platform}")

    urllib.request.urlretrieve(url, local_binary)
    if not sys.platform.startswith("win"):
        os.chmod(local_binary, 0o755)
    return local_binary

def run_with_cloudflare():
    port = 8501
    script_path = os.path.abspath(__file__)

    st_cmd = [sys.executable, "-m", "streamlit", "run", script_path, "--server.port", str(port), "--server.headless", "true"]
    print(f"🚀 Starting Streamlit backend on port {port}...")
    st_proc = subprocess.Popen(st_cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    cf_bin = get_cloudflared_binary()
    print("🌐 Initiating Cloudflare Tunnel...")
    cf_cmd = [cf_bin, "tunnel", "--url", f"http://127.0.0.1:{port}"]
    cf_proc = subprocess.Popen(cf_cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1)

    url_regex = re.compile(r"https://[a-zA-Z0-9-]+\.trycloudflare\.com")
    public_url = None

    try:
        for line in iter(cf_proc.stdout.readline, ''):
            match = url_regex.search(line)
            if match:
                public_url = match.group(0)
                break

        if public_url:
            with open(TUNNEL_CACHE_FILE, "w") as f:
                f.write(public_url)
            print("\n" + "=" * 65)
            print("🚀 CLOUDFLARE PUBLIC TUNNEL IS LIVE!")
            print(f"🔗 Public URL: {public_url}")
            print(f"🏠 Local URL:  http://localhost:{port}")
            print("=" * 65)
            print("App is active. Press Ctrl+C anytime to stop.\n")
        
        while True:
            time.sleep(1)

    except KeyboardInterrupt:
        print("\n🛑 Shutting down Cloudflare Tunnel and Streamlit...")
    finally:
        if os.path.exists(TUNNEL_CACHE_FILE):
            try: os.remove(TUNNEL_CACHE_FILE)
            except: pass
        st_proc.terminate()
        cf_proc.terminate()
        sys.exit(0)

if not RUNNING_IN_STREAMLIT and __name__ == "__main__":
    run_with_cloudflare()

# ==========================================
# 1. UI SETUP & ANIMATED CSS
# ==========================================
st.set_page_config(page_title="esentry Attendance", page_icon="⏱️", layout="wide")
page_id = "dashboard"

st.markdown("""
<style>
.stApp { background-color: #000000; color: #CAEEC8; font-family: 'Inter', sans-serif; }
p, div, span, label, h1, h2, h3, h4, h5, h6 { color: #CAEEC8 !important; }
#MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
@keyframes curtainLeft { 0% { right: 50%; opacity: 1; z-index: 9999999; } 99% { right: 100%; opacity: 0; z-index: 9999999; } 100% { right: 100%; opacity: 0; z-index: -1; visibility: hidden; } }
@keyframes curtainRight { 0% { left: 50%; opacity: 1; z-index: 9999999; } 99% { left: 100%; opacity: 0; z-index: 9999999; } 100% { left: 100%; opacity: 0; z-index: -1; visibility: hidden; } }
.block-container { animation: slideUpFade 1.2s ease-out forwards; animation-delay: 0.1s; opacity: 0; padding-top: 2rem; }
@keyframes slideUpFade { 0% { opacity: 0; transform: translateY(40px); } 100% { opacity: 1; transform: translateY(0); } }
@keyframes floatLogo { 0% { transform: translateY(0px); box-shadow: 0 10px 20px rgba(0, 0, 71, 0.6); } 50% { transform: translateY(-10px); box-shadow: 0 20px 30px rgba(202, 238, 200, 0.2); } 100% { transform: translateY(0px); box-shadow: 0 10px 20px rgba(0, 0, 71, 0.6); } }
.logo-container { background-color: #CAEEC8; padding: 15px 30px; border-radius: 16px; display: inline-block; animation: floatLogo 4s ease-in-out infinite; border: 2px solid #000047; }
[data-testid="stFileUploader"] { border: 2px dashed #000047; border-radius: 12px; background: rgba(0, 0, 71, 0.2); padding: 10px; transition: all 0.4s ease; cursor: pointer; }
[data-testid="stFileUploader"]:hover { border-color: #CAEEC8; background: rgba(0, 0, 71, 0.6); box-shadow: 0 0 20px rgba(202, 238, 200, 0.15); transform: translateY(-2px); }
.stButton>button { background: linear-gradient(135deg, #000047 0%, #000000 100%); color: #CAEEC8 !important; border-radius: 8px; border: 1px solid #CAEEC8; font-weight: bold; transition: all 0.3s ease; }
.stButton>button:hover { background: #CAEEC8; color: #000000 !important; transform: translateY(-3px) scale(1.02); box-shadow: 0 6px 20px rgba(202, 238, 200, 0.3); border-color: #000000; }
[data-testid="stDataFrame"] > div { border: 1px solid #000047; border-radius: 10px; overflow: hidden; box-shadow: 0 8px 30px rgba(0,0,0,0.5) !important; background-color: #000047; }
th { background-color: #000000 !important; color: #CAEEC8 !important; }
td { background-color: rgba(0, 0, 71, 0.4) !important; color: white !important; }
@keyframes outputPopUp { 0%, 30% { transform: scale(0.95) translateY(30px); opacity: 0; } 100% { transform: scale(1) translateY(0); opacity: 1; } }
div[data-testid="stDataFrame"], .output-alert-box { animation: outputPopUp 1.0s cubic-bezier(0.175, 0.885, 0.32, 1.275) forwards; }
.custom-alert { background-color: #000047; border-left: 4px solid #CAEEC8; padding: 15px 20px; border-radius: 0 8px 8px 0; margin-bottom: 2rem; box-shadow: 0 4px 15px rgba(0,0,0,0.3); transition: transform 0.3s ease; }
button[data-baseweb="tab"] { color: #CAEEC8 !important; background-color: transparent !important; transition: all 0.3s ease; }
button[data-baseweb="tab"]:hover { background-color: rgba(202, 238, 200, 0.1) !important; }
button[data-baseweb="tab"][aria-selected="true"] { border-bottom: 3px solid #CAEEC8 !important; background-color: #000047 !important; font-weight: bold; }
</style>
<div class="curtain-panel-left-dashboard" style="position: fixed; top: 0; left: 0; right: 50%; bottom: 0; background-color: #000000; border-right: 2px solid #CAEEC8; box-shadow: 5px 0 30px rgba(0,0,71,0.8); animation: curtainLeft 1.2s cubic-bezier(0.86, 0, 0.07, 1) forwards; animation-delay: 0.1s; z-index: 9999999; pointer-events: none;"></div>
<div class="curtain-panel-right-dashboard" style="position: fixed; top: 0; left: 50%; right: 0; bottom: 0; background-color: #000000; border-left: 2px solid #CAEEC8; box-shadow: -5px 0 30px rgba(0,0,71,0.8); animation: curtainRight 1.2s cubic-bezier(0.86, 0, 0.07, 1) forwards; animation-delay: 0.1s; z-index: 9999999; pointer-events: none;"></div>
""", unsafe_allow_html=True)

def get_header_html(title_text, motto=""):
    logo_url = "https://cdn.prod.website-files.com/679a0d90717922eb8830bfb4/6abe670bbb0b6bb7d828edef_esentry%20black-01-p-2000.png"
    motto_html = f'<div style="margin-top: 12px; font-size: 1.15rem; color: #CAEEC8; border-left: 3px solid #CAEEC8; padding-left: 15px; font-weight: 400; opacity: 0.9;">✦ {motto}</div>' if motto else ""
    return f"""
    <div style="display: flex; align-items: center; margin-bottom: 2.5rem; gap: 30px; flex-wrap: wrap;">
        <div class="logo-container">
            <img src="{logo_url}" width="160" style="display: block;">
        </div>
        <div style="display: flex; flex-direction: column; justify-content: center;">
            <h1 style="margin: 0; padding: 0; color: #CAEEC8; font-weight: 800; letter-spacing: -1px; text-transform: uppercase;">{title_text}</h1>
            {motto_html}
        </div>
    </div>
    """

if os.path.exists(TUNNEL_CACHE_FILE):
    try:
        with open(TUNNEL_CACHE_FILE, "r") as f:
            cf_url = f.read().strip()
            if cf_url:
                st.markdown(f"**🌐 Live App Broadcast Tunnel:** [{cf_url}]({cf_url})")
    except: pass

st.markdown(get_header_html("Attendance Analytics", motto="Automated Latecomer Tracking System"), unsafe_allow_html=True)
st.markdown("""<div class="custom-alert">Upload access logs alongside the <strong>Staff Master List</strong> and <strong>SOC Schedule</strong> to generate reports. The system strictly isolates the <strong>absolute single first valid entrance scan</strong> per day on <strong>Main entrance Upfloor / Ground Flr</strong>. Anyone registering at <strong>08:31:00 AM</strong> or later is logged as Late.</div>""", unsafe_allow_html=True)

# ==========================================
# 3. UPLOADS & EXPORT HELPERS
# ==========================================
up_col1, up_col2, up_col3 = st.columns(3)
with up_col1:
    uploaded_files = st.file_uploader("Access Control Logs (Unlimited)", type=['csv', 'xlsx', 'xls'], accept_multiple_files=True)
with up_col2:
    roster_file = st.file_uploader("Standard Staff Master List", type=['csv', 'xlsx', 'xls'])
with up_col3:
    soc_file = st.file_uploader("SOC Team Schedule", type=['csv', 'xlsx', 'xls'])

SOC_SHIFT_TIMES = {'morning': datetime.time(7, 0, 0), 'afternoon': datetime.time(14, 0, 0), 'night': datetime.time(19, 0, 0)}

def render_export_dropdown(dataframe, period_label, report_title, report_subtitle, col_widths_pdf, key_prefix):
    with st.popover("Export Report Options"):
        # Excel
        buf_xlsx = io.BytesIO()
        with pd.ExcelWriter(buf_xlsx, engine='openpyxl') as writer:
            dataframe.to_excel(writer, index=False, sheet_name=report_title[:31])
        st.download_button("Download as Excel (.xlsx)", data=buf_xlsx.getvalue(), file_name=f"{key_prefix}_{period_label.replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"xlsx_{key_prefix}_{period_label}")
        
        # Word
        doc = Document()
        doc.add_heading(f'{report_title} — {period_label}', level=1)
        doc.add_paragraph(report_subtitle)
        tbl = doc.add_table(rows=1, cols=len(dataframe.columns), style='Light Grid Accent 1')
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr = tbl.rows[0].cells
        for i, cn in enumerate(dataframe.columns): hdr[i].text = str(cn)
        for _, row in dataframe.iterrows():
            rc = tbl.add_row().cells
            for i, cn in enumerate(dataframe.columns): rc[i].text = str(row[cn])
        buf_docx = io.BytesIO()
        doc.save(buf_docx)
        st.download_button("Download as Word (.docx)", data=buf_docx.getvalue(), file_name=f"{key_prefix}_{period_label.replace(' ', '_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"docx_{key_prefix}_{period_label}")
        
        # PDF
        def clean_pdf_text(txt):
            return str(txt).replace('✅', '').replace('⚠️', '').replace('🚫', '').strip().encode('latin-1', 'ignore').decode('latin-1')

        pdf = FPDF(orientation='L')
        pdf.add_page()
        pdf.set_font('Helvetica', 'B', 14)
        pdf.cell(w=0, h=10, txt=clean_pdf_text(f'{report_title} - {period_label}'), border=0, ln=1, align='C')
        pdf.set_font('Helvetica', '', 9)
        pdf.cell(w=0, h=8, txt=clean_pdf_text(report_subtitle), border=0, ln=1, align='C')
        pdf.ln(5)
        pdf.set_font('Helvetica', 'B', 8)
        for i, h in enumerate(dataframe.columns):
            w = col_widths_pdf[i] if i < len(col_widths_pdf) else 40
            pdf.cell(w=w, h=8, txt=clean_pdf_text(h), border=1, ln=0, align='C')
        pdf.ln()
        pdf.set_font('Helvetica', '', 7)
        for _, row in dataframe.iterrows():
            for i, cn in enumerate(dataframe.columns):
                w = col_widths_pdf[i] if i < len(col_widths_pdf) else 40
                align = 'C' if i in [0, 2] else 'L'
                pdf.cell(w=w, h=7, txt=clean_pdf_text(row[cn]), border=1, ln=0, align=align)
            pdf.ln()

        try:
            buf_pdf = io.BytesIO()
            pdf.output(buf_pdf)
            pdf_bytes = buf_pdf.getvalue()
        except Exception:
            pdf_bytes = pdf.output(dest='S').encode('latin-1')

        st.download_button("Download as PDF (.pdf)", data=pdf_bytes, file_name=f"{key_prefix}_{period_label.replace(' ', '_')}.pdf", mime="application/pdf", key=f"pdf_{key_prefix}_{period_label}")

def period_filter(df, date_col, parsed_time_col, radio_key):
    period = st.radio("View period:", ["All", "Daily", "Weekly", "Monthly"], horizontal=True, key=radio_key)
    if period == "Daily":
        df['_date_label'] = df[parsed_time_col].dt.date
        opts = sorted(df['_date_label'].unique(), reverse=True)
        sel = st.selectbox("Select Date", opts, key=f"{radio_key}_sel")
        return df[df['_date_label'] == sel].copy(), f"Day {sel}"
    elif period == "Weekly":
        df['_week'] = df[parsed_time_col].dt.isocalendar().year.astype(str) + '-W' + df[parsed_time_col].dt.isocalendar().week.astype(str).str.zfill(2)
        opts = sorted(df['_week'].unique(), reverse=True)
        sel = st.selectbox("Select Week", opts, key=f"{radio_key}_sel")
        return df[df['_week'] == sel].copy(), f"Week {sel}"
    elif period == "Monthly":
        df['_month'] = df[parsed_time_col].dt.to_period('M').astype(str)
        opts = sorted(df['_month'].unique(), reverse=True)
        sel = st.selectbox("Select Month", opts, key=f"{radio_key}_sel")
        return df[df['_month'] == sel].copy(), f"Month {sel}"
    return df.copy(), "All Time"

# ==========================================
# 4. CORE PROCESSING LOGIC
# ==========================================
if uploaded_files:
    with st.spinner("Decrypting Access Logs & Compiling High-Volume Report..."):
        time.sleep(0.5)

    try:
        STANDARD_COLS = {'time': 'Time', 'date': 'Date', 'door name': 'Door Name', 'event description': 'Event Description', 'personnel id': 'Personnel ID', 'first name': 'First Name', 'last name': 'Last Name'}
        
        def normalize_columns(df_in):
            col_map = {}
            for c in df_in.columns:
                cl = str(c).lower().strip()
                if cl in STANDARD_COLS:
                    col_map[c] = STANDARD_COLS[cl]
                elif 'door' in cl or 'reader' in cl:
                    col_map[c] = 'Door Name'
                elif 'time' in cl and 'date' in cl:
                    col_map[c] = 'Time'
                elif 'event time' in cl or cl == 'datetime' or cl == 'date time':
                    col_map[c] = 'Time'
                elif 'event' in cl and 'desc' in cl:
                    col_map[c] = 'Event Description'
                elif 'pin' in cl or 'personnel' in cl or 'emp no' in cl or 'emp id' in cl:
                    col_map[c] = 'Personnel ID'
                elif 'first name' in cl or 'fname' in cl:
                    col_map[c] = 'First Name'
                elif 'last name' in cl or 'lname' in cl:
                    col_map[c] = 'Last Name'
                else:
                    col_map[c] = str(c).strip()
            return df_in.rename(columns=col_map)

        def clean_id(val):
            if pd.isna(val): return ""
            val_str = str(val).strip().upper() # Explicitly uppercase to ensure matches between logs and masterlist
            return val_str[:-2] if val_str.endswith('.0') else val_str

        def has_req_cols(cols):
            lc = [str(c).strip().lower() for c in cols]
            # Match variations of time and door mappings flexibly 
            has_time = any(kw in c for c in lc for kw in ['time', 'date'])
            has_door = any(kw in c for c in lc for kw in ['door', 'reader'])
            return has_time and has_door

        # -- Load access sheets (handles CSV, multi-sheet XLSX seamlessly)
        all_dfs = []
        for uf in uploaded_files:
            if uf.name.endswith('.csv'):
                tmp = pd.read_csv(uf)
                tmp.columns = tmp.columns.astype(str).str.strip()
                all_dfs.append(normalize_columns(tmp))
            else:
                xls = pd.ExcelFile(uf)
                for sheet in xls.sheet_names:
                    d = pd.read_excel(xls, sheet_name=sheet)
                    d = d.dropna(how='all').reset_index(drop=True)
                    if d.empty: continue
                    
                    if not has_req_cols(d.columns):
                        found = False
                        for i in range(min(20, len(d))):
                            rv = [str(v).strip() for v in d.iloc[i].values]
                            if has_req_cols(rv):
                                d.columns = rv
                                d = d.iloc[i+1:].reset_index(drop=True)
                                found = True
                                break
                        if not found: continue
                    
                    d.columns = d.columns.astype(str).str.strip()
                    d = normalize_columns(d)
                    if 'Time' in d.columns and 'Door Name' in d.columns:
                        all_dfs.append(d)

        if not all_dfs:
            st.error("No valid data found in the uploaded log files. Ensure they contain 'Time' and 'Door/Reader Name'.")
            st.stop()

        df = pd.concat(all_dfs, ignore_index=True)
        if 'Personnel ID' in df.columns:
            df['Personnel ID'] = df['Personnel ID'].apply(clean_id)
            df = df[df['Personnel ID'] != '0']
            df = df[df['Personnel ID'] != '']

        # -- RULE 1: EXCLUDE EXITS (Prevent tailgaters' exit times from becoming their arrival time)
        if 'Event Description' in df.columns:
            evt_lower = df['Event Description'].astype(str).str.lower()
            df = df[~evt_lower.str.contains('exit|out', na=False)].copy()

        # -- RULE 2: Filter explicitly to TARGET DOORS ONLY safely without crushing string variations
        dc = df['Door Name'].astype(str).str.lower()
        # Look for "main" paired with either "ent", "door", "up", "ground", "gnd", or "flr"
        door_mask = dc.str.contains('main', na=False) & (
            dc.str.contains('ent|door', na=False) | 
            dc.str.contains('up|ground|gnd|flr', na=False)
        )
        df = df[door_mask].copy()

        # -- RULE 3: Robust Date & Time Parsing preventing timezone/format shift bugs
        def safe_parse_date_series(col):
            if pd.api.types.is_datetime64_any_dtype(col):
                if hasattr(col.dt, 'tz') and col.dt.tz is not None:
                    return col.dt.tz_localize(None)
                return col
            
            # Critical Fix: Force conversion via string to avoid `datetime.time` object corruption by pd.to_datetime 
            def to_str(x):
                if pd.isna(x): return ""
                if isinstance(x, (datetime.time, datetime.date, datetime.datetime)):
                    return str(x)
                return str(x).strip()
            
            # Enforce dayfirst=True to accommodate Nigeria (DD/MM/YYYY) format so months don't swap
            parsed = pd.to_datetime(col.apply(to_str), errors='coerce', dayfirst=True)
            if hasattr(parsed.dt, 'tz') and parsed.dt.tz is not None:
                parsed = parsed.dt.tz_localize(None)
            return parsed

        df['Parsed_Time'] = safe_parse_date_series(df['Time'])
        
        # If separated Date and Time columns exist, merge them appropriately extracting Time correctly
        if 'Date' in df.columns:
            parsed_date_col = safe_parse_date_series(df['Date'])
            valid = parsed_date_col.notna() & df['Parsed_Time'].notna()
            if valid.any():
                combined_strs = parsed_date_col[valid].dt.strftime('%Y-%m-%d') + ' ' + df.loc[valid, 'Parsed_Time'].dt.strftime('%H:%M:%S')
                df.loc[valid, 'Parsed_Time'] = pd.to_datetime(combined_strs, errors='coerce')

        df = df.dropna(subset=['Parsed_Time']).copy()
        df['Date'] = df['Parsed_Time'].dt.date
        df['Time_Only'] = df['Parsed_Time'].dt.time
        df['Date_dt'] = pd.to_datetime(df['Date'])

        # -- RULE 4: STRICT SINGLE EARLIEST ENTRY LOGIC --
        df_first = df.sort_values('Parsed_Time').drop_duplicates(subset=['Date', 'Personnel ID'], keep='first').copy()
        
        # NOTE: Using strictly '<' so 08:31:00 exactly evaluates to False (meaning they are Late)
        CUTOFF_TIME = datetime.time(8, 31, 0)
        df_first['Is_Early'] = df_first['Time_Only'] < CUTOFF_TIME

        # -- Load Rosters --
        def load_simple_excel(f_obj):
            return pd.read_csv(f_obj) if f_obj.name.endswith('.csv') else pd.read_excel(f_obj)

        df_roster, df_std_roster, soc_ids_from_roster = None, None, set()
        
        if roster_file:
            df_roster_temp = load_simple_excel(roster_file)
            
            # Fuzzy match columns to avoid silent crashes
            col_map = {}
            for c in df_roster_temp.columns:
                cl = str(c).lower().strip()
                if 'personnel' in cl or 'id' in cl or 'pin' in cl or 'employee' in cl:
                    if 'Personnel ID' not in col_map.values():
                        col_map[c] = 'Personnel ID'
                elif 'expected' in cl or 'days' in cl or 'work' in cl:
                    if 'Expected Days' not in col_map.values():
                        col_map[c] = 'Expected Days'
                        
            df_roster = df_roster_temp.rename(columns=col_map)
            
            if 'Personnel ID' in df_roster.columns:
                df_roster['Personnel ID'] = df_roster['Personnel ID'].apply(clean_id)
                
                # Check for SOC and construct std_roster safely even if Expected Days is completely missing
                if 'Expected Days' in df_roster.columns:
                    df_roster['Expected Days'] = df_roster['Expected Days'].astype(str).str.strip()
                    soc_mask = df_roster['Expected Days'].str.upper() == 'SOC'
                    soc_ids_from_roster = set(df_roster.loc[soc_mask, 'Personnel ID'].unique())
                    df_std_roster = df_roster[~soc_mask].copy()
                    df_std_roster['Days_Per_Week'] = pd.to_numeric(df_std_roster['Expected Days'], errors='coerce').fillna(5).astype(int)
                else:
                    df_std_roster = df_roster.copy()
                    df_std_roster['Days_Per_Week'] = 5 
            else:
                st.warning("Master List is missing a recognizable ID column (e.g., 'Personnel ID'). Weekly compliance will be disabled.")

        df_soc = None
        if soc_file:
            df_soc = load_simple_excel(soc_file)
            df_soc = normalize_columns(df_soc)
            if all(c in df_soc.columns for c in ['Personnel ID', 'Date', 'Shift']):
                df_soc['Date'] = safe_parse_date_series(df_soc['Date']).dt.date
                df_soc = df_soc.dropna(subset=['Date']).copy()
                df_soc['Shift_Start'] = df_soc['Shift'].astype(str).str.strip().str.lower().map(SOC_SHIFT_TIMES)
                df_soc['Personnel ID'] = df_soc['Personnel ID'].apply(clean_id)
                fn_cols = [c for c in ['First Name', 'Last Name'] if c in df_soc.columns]
                df_soc['Full Name'] = df_soc[fn_cols].fillna('').astype(str).apply(lambda r: ' '.join(r).strip(), axis=1) if fn_cols else ''

        # ==========================================
        # 5. TABS & REPORT GENERATION
        # ==========================================
        st.divider()
        tab_std, tab_soc = st.tabs(["📊 Standard Staff", "🚨 SOC Team"])

        with tab_std:
            soc_ids = soc_ids_from_roster.copy()
            if df_soc is not None: soc_ids.update(df_soc['Personnel ID'].unique())
            
            df_std_first = df_first[~df_first['Personnel ID'].astype(str).str.strip().isin(soc_ids)].copy()
            
            name_cols = [c for c in ['First Name', 'Last Name'] if c in df_std_first.columns]
            df_std_first['Full Name'] = df_std_first[name_cols].fillna('').astype(str).apply(lambda r: ' '.join(r).strip(), axis=1) if name_cols else ''
            
            WORKDAYS = {0, 1, 2, 3, 4} # Mon-Fri Only
            df_std_workday = df_std_first[df_std_first['Date_dt'].dt.dayofweek.isin(WORKDAYS)].copy()
            df_std_workday['Date_Str'] = df_std_workday['Date'].apply(lambda d: d.strftime('%Y-%m-%d'))
            df_std_workday['Time_Str'] = df_std_workday['Time_Only'].apply(lambda t: t.strftime('%H:%M:%S'))

            # ── Section 1: 📋 Complete Daily Log ────────
            st.markdown("### 📋 Daily Check-in Log (Earliest Scans)")
            st.markdown("Shows the absolute earliest scan per person on the target doors, and whether they were **On Time** (< 08:31:00) or **Late**.")
            
            df_std_workday['Status'] = df_std_workday['Is_Early'].apply(lambda x: '✅ On Time ' if x else '⚠️ Late ')
            
            log_pf_col, log_ex_col = st.columns([4, 1])
            with log_pf_col:
                daily_log_f, daily_log_period = period_filter(df_std_workday.copy(), 'Date', 'Parsed_Time', 'std_daily_log_period')
            
            if not daily_log_f.empty:
                daily_log_disp = daily_log_f[['Personnel ID', 'Full Name', 'Date_Str', 'Time_Str', 'Door Name', 'Status']].sort_values(['Date_Str', 'Time_Str']).rename(columns={'Date_Str': 'Date', 'Time_Str': 'Earliest Time', 'Door Name': 'Logged Door'})
                st.dataframe(daily_log_disp, hide_index=True, use_container_width=True)
                with log_ex_col:
                    render_export_dropdown(daily_log_disp, daily_log_period, "Standard Staff Daily Log", "Earliest scans at target doors.", [15, 35, 15, 15, 45, 15], "std_daily_log")
            else:
                st.info("No logs found for the selected period.")

            st.divider()

            # ── Section 2: ⚠️ Strictly Late Arrivals Counter View ────────────
            st.markdown("### ⚠️ Late Arrivals Dashboard")
            st.markdown("This view exclusively tracks standard staff who arrived at **08:31:00 AM or later**. It counts the number of infractions per person within the selected timeframe.")

            df_late = df_std_workday[~df_std_workday['Is_Early']].copy()

            if df_late.empty:
                st.markdown('<div class="output-alert-box" style="background:#000047;border-left:5px solid #CAEEC8;padding:15px 20px;border-radius:6px;margin-bottom:1.5rem;"><h3 style="margin:0;color:#CAEEC8;">🏆 Perfect Attendance</h3><p style="margin:5px 0 0;color:white;">No standard staff checked in at or past 8:31 AM during this period.</p></div>', unsafe_allow_html=True)
            else:
                pf_col, ex_col = st.columns([4, 1])
                with pf_col: 
                    df_late_f, period_label_late = period_filter(df_late.copy(), 'Date', 'Parsed_Time', 'std_late_period')
                
                if df_late_f.empty:
                    st.info("No late arrivals recorded for the specifically selected date/timeframe.")
                else:
                    late_agg = (df_late_f.sort_values(['Date', 'Time_Only'])
                                .groupby(['Personnel ID', 'Full Name'])
                                .agg(
                                    Times_Late=('Date', 'nunique'),
                                    Dates_Late=('Date_Str', lambda x: ', '.join(x)),
                                    Times_In=('Time_Str', lambda x: ', '.join(x))
                                )
                                .reset_index()
                                .sort_values('Times_Late', ascending=False))
                    
                    late_agg.rename(columns={
                        'Times_Late': 'Total Times Late',
                        'Dates_Late': 'Dates Late',
                        'Times_In': 'Exact Arrival Times'
                    }, inplace=True)

                    if df_std_roster is not None and 'Days_Per_Week' in df_std_roster.columns:
                        late_agg = late_agg.merge(df_std_roster[['Personnel ID', 'Days_Per_Week']], on='Personnel ID', how='left')
                        late_agg['Days_Per_Week'] = late_agg['Days_Per_Week'].fillna('–')
                        late_agg = late_agg[['Personnel ID', 'Full Name', 'Days_Per_Week', 'Total Times Late', 'Dates Late', 'Exact Arrival Times']]
                        late_agg.columns = ['Personnel ID', 'Full Name', 'Expected/Week', 'Total Times Late', 'Dates Late', 'Exact Arrival Times']
                        pdf_widths = [15, 30, 15, 15, 60, 60]
                    else:
                        late_agg = late_agg[['Personnel ID', 'Full Name', 'Total Times Late', 'Dates Late', 'Exact Arrival Times']]
                        pdf_widths = [18, 40, 20, 60, 60]

                    total_incidents = late_agg['Total Times Late'].sum()
                    st.markdown(f'<div class="output-alert-box" style="background:#000000;border:1px solid #CAEEC8;border-left:5px solid #CAEEC8;padding:15px 20px;border-radius:6px;margin-bottom:1.5rem;"><h3 style="margin:0;color:#CAEEC8;">⚠️ {len(late_agg)} Personnel Late ({total_incidents} Total Incidents)</h3><p style="margin:5px 0 0;color:white;">Showing late count metrics for {period_label_late}.</p></div>', unsafe_allow_html=True)

                    st.dataframe(late_agg, hide_index=True, use_container_width=True)
                    
                    with ex_col: 
                        render_export_dropdown(late_agg, period_label_late, "Standard Staff Late Arrivals", "Personnel mapped to first scan >= 08:31 at Main Entrances.", pdf_widths, "std_late")

            st.divider()

            # ── Section 3: 📅 Weekly Attendance Compliance ────────────────────
            st.markdown("### 📅 Weekly Attendance Compliance (Days in Office)")

            if df_std_roster is None:
                st.info("Upload the Standard Staff Master List to enable weekly compliance tracking.")
            else:
                _name_lkp = {}
                if 'First Name' in df_first.columns:
                    _nl = df_first[['Personnel ID', 'First Name', 'Last Name']].copy()
                    _nc_cols = [c for c in ['First Name', 'Last Name'] if c in _nl.columns]
                    _nl['Full Name'] = _nl[_nc_cols].fillna('').astype(str).apply(lambda r: ' '.join(r).strip(), axis=1)
                    _name_lkp = dict(zip(_nl['Personnel ID'], _nl['Full Name']))

                _df_att = df_first[(~df_first['Personnel ID'].isin(soc_ids)) & (df_first['Date_dt'].dt.dayofweek.isin(WORKDAYS))].copy()
                _df_att['ISO_Year_n'] = _df_att['Parsed_Time'].dt.isocalendar().year.astype(int)
                _df_att['ISO_Week_n'] = _df_att['Parsed_Time'].dt.isocalendar().week.astype(int)
                _df_att['Week_Label'] = _df_att['ISO_Year_n'].astype(str) + '-W' + _df_att['ISO_Week_n'].astype(str).str.zfill(2)
                _df_att['Date_Str'] = _df_att['Date'].apply(lambda d: d.strftime('%Y-%m-%d'))
                _df_att['Time_Str'] = _df_att['Time_Only'].apply(lambda t: t.strftime('%H:%M'))

                if _df_att.empty: st.info("No workday attendance data found.")
                else:
                    _min_dt, _max_dt = _df_att['Parsed_Time'].min(), _df_att['Parsed_Time'].max()
                    _wk_starts = pd.date_range(_min_dt.to_period('W').to_timestamp(), _max_dt.to_period('W').to_timestamp(), freq='W-MON')
                    _iso_weeks = [{'year': int(_ws.isocalendar()[0]), 'week': int(_ws.isocalendar()[1]), 'label': f"{int(_ws.isocalendar()[0])}-W{str(int(_ws.isocalendar()[1])).zfill(2)}"} for _ws in _wk_starts]
                    if not _iso_weeks: _iso_weeks = [{'year': int(_min_dt.isocalendar()[0]), 'week': int(_min_dt.isocalendar()[1]), 'label': f"{int(_min_dt.isocalendar()[0])}-W{str(int(_min_dt.isocalendar()[1])).zfill(2)}"}]

                    _comp_rows = []
                    for _, _pr in df_std_roster.iterrows():
                        _pid, _exp = str(_pr['Personnel ID']), int(_pr['Days_Per_Week'])
                        _fname = _name_lkp.get(_pid, '')
                        _p_att = _df_att[_df_att['Personnel ID'] == _pid]

                        for _wk in _iso_weeks:
                            try: _mon = datetime.date.fromisocalendar(_wk['year'], _wk['week'], 1)
                            except: continue
                            _wd_strs = [(_mon + datetime.timedelta(days=_d)).strftime('%Y-%m-%d') for _d in range(5)]
                            _wk_att = _p_att[_p_att['Week_Label'] == _wk['label']].sort_values('Date')
                            _att_strs = set(_wk_att['Date_Str'].tolist())
                            _att_entries = [f"{_r['Date_Str']} ({_r['Time_Str']})" for _, _r in _wk_att.iterrows()]
                            _absent = [_d for _d in _wd_strs if _d not in _att_strs]
                            _actual = len(_att_strs)
                            _deficit = max(0, _exp - _actual)
                            _comp_rows.append({'Personnel ID': _pid, 'Full Name': _fname, 'Week': _wk['label'], 'Expected Days': _exp, 'Days Present': _actual, 'Deficit': _deficit, 'Status': '✅ Met' if _deficit == 0 else f'⚠️ Short by {_deficit}', 'Days Attended & Times': ', '.join(_att_entries) if _att_entries else '—', 'Absent Dates': ', '.join(_absent) if _absent else '—'})

                    _df_comp = pd.DataFrame(_comp_rows)
                    _nc_count = _df_comp[_df_comp['Deficit'] > 0]['Personnel ID'].nunique()
                    if _nc_count == 0:
                        st.markdown('<div style="background:#000047;border-left:5px solid #CAEEC8;padding:15px 20px;border-radius:6px;margin-bottom:1.5rem;"><h3 style="margin:0;color:#CAEEC8;">✅ Full Compliance</h3><p style="margin:5px 0 0;color:white;">All standard staff met required weekly attendance.</p></div>', unsafe_allow_html=True)
                    else:
                        st.markdown(f'<div style="background:#000000;border:1px solid #CAEEC8;border-left:5px solid #CAEEC8;padding:15px 20px;border-radius:6px;margin-bottom:1.5rem;"><h3 style="margin:0;color:#CAEEC8;">⚠️ {_nc_count} staff member(s) non-compliant</h3><p style="margin:5px 0 0;color:white;">Based on required days vs. actual badge records.</p></div>', unsafe_allow_html=True)

                    _pf2_col, _ex2_col = st.columns([4, 1])
                    _all_wk_opts = sorted(_df_comp['Week'].unique(), reverse=True)
                    with _pf2_col: _comp_view = st.radio("View period:", ["All Weeks", "Select Week"], horizontal=True, key="comp_period_radio")
                    _detail_view = _df_comp[_df_comp['Week'] == st.selectbox("Select Week", _all_wk_opts, key="comp_week_sel")].copy() if _comp_view == "Select Week" and _all_wk_opts else _df_comp.copy()

                    st.markdown("**Weekly Detail**")
                    st.dataframe(_detail_view, hide_index=True, use_container_width=True)

                    st.markdown("**Summary**")
                    _nc_wk_list = _df_comp[_df_comp['Deficit'] > 0].groupby('Personnel ID')['Week'].apply(lambda x: ', '.join(sorted(x))).reset_index().rename(columns={'Week': 'Non-Compliant Weeks'})
                    _summary = _df_comp.groupby(['Personnel ID', 'Full Name']).agg(Expected_Per_Week=('Expected Days', 'first'), Weeks_NC=('Deficit', lambda x: int((x > 0).sum())), Total_Deficit=('Deficit', 'sum')).reset_index().merge(_nc_wk_list, on='Personnel ID', how='left')
                    _summary['Non-Compliant Weeks'] = _summary['Non-Compliant Weeks'].fillna('—')
                    _summary = _summary.sort_values('Total_Deficit', ascending=False)
                    _summary.columns = ['Personnel ID', 'Full Name', 'Expected/Week', 'Weeks Non-Compliant', 'Total Deficit Days', 'Non-Compliant Weeks']
                    st.dataframe(_summary, hide_index=True, use_container_width=True)
                    
                    with _ex2_col: 
                        render_export_dropdown(_detail_view, "All Time", "Weekly Attendance Compliance", "Standard staff weekly attendance.", [15, 35, 12, 12, 12, 8, 15, 60, 60], "std_compliance")

        with tab_soc:
            if df_soc is None:
                st.info("Upload the SOC Team Schedule file to enable the SOC report.")
            else:
                _soc_pids = df_soc['Personnel ID'].unique()
                _soc_logs = df_first[df_first['Personnel ID'].isin(_soc_pids)].copy()

                _merged = df_soc.merge(_soc_logs[['Personnel ID', 'Date', 'Time_Only', 'Parsed_Time']], on=['Personnel ID', 'Date'], how='left')
                _sc = _merged.copy()
                _sc['Date_Str'] = _sc['Date'].apply(lambda d: d.strftime('%Y-%m-%d') if hasattr(d, 'strftime') else str(d))
                _sc['Sched_Start_Str'] = _sc['Shift_Start'].apply(lambda t: t.strftime('%H:%M') if pd.notna(t) else '—')
                _sc['Actual_In_Str'] = _sc['Time_Only'].apply(lambda t: t.strftime('%H:%M') if pd.notna(t) else '—')
                _sc['Shift_Title'] = _sc['Shift'].astype(str).str.strip().str.title()
                _sc['Date_dt_col'] = pd.to_datetime(_sc['Date_Str'])

                def _soc_status(row):
                    if pd.isna(row['Time_Only']): return '🚫 Absent '
                    if pd.notna(row['Shift_Start']) and row['Time_Only'] > row['Shift_Start']: return '⚠️ Late '
                    return '✅ On Time '

                def _soc_mins(row):
                    if pd.isna(row['Time_Only']) or pd.isna(row['Shift_Start']): return 0
                    if row['Time_Only'] > row['Shift_Start']:
                        return int((datetime.datetime.combine(row['Date'], row['Time_Only']) - datetime.datetime.combine(row['Date'], row['Shift_Start'])).total_seconds() / 60)
                    return 0

                _sc['Status'] = _sc.apply(_soc_status, axis=1)
                _sc['Mins Late'] = _sc.apply(_soc_mins, axis=1)

                _soc_comp_all = _sc[['Personnel ID', 'Full Name', 'Date_Str', 'Shift_Title', 'Sched_Start_Str', 'Actual_In_Str', 'Status', 'Mins Late']].rename(columns={'Date_Str': 'Shift Date', 'Shift_Title': 'Shift', 'Sched_Start_Str': 'Scheduled Start', 'Actual_In_Str': 'Actual Check-in'}).sort_values(['Personnel ID', 'Shift Date'])

                _late_pids, _abs_pids = _sc[_sc['Status'] == '⚠️ Late ']['Personnel ID'].nunique(), _sc[_sc['Status'] == '🚫 Absent ']['Personnel ID'].nunique()
                _banner = []
                if _late_pids: _banner.append(f"<strong style='color:#CAEEC8;'>{_late_pids}</strong> member(s) arrived late")
                if _abs_pids: _banner.append(f"<strong style='color:#CAEEC8;'>{_abs_pids}</strong> member(s) had absences")

                if not _banner: st.markdown('<div class="output-alert-box" style="background:#000047;border-left:5px solid #CAEEC8;padding:15px 20px;border-radius:6px;margin-bottom:1.5rem;"><h3 style="margin:0;color:#CAEEC8;">✅ Perfect SOC Compliance!</h3></div>', unsafe_allow_html=True)
                else: st.markdown(f'<div class="output-alert-box" style="background:#000000;border:1px solid #CAEEC8;border-left:5px solid #CAEEC8;padding:15px 20px;border-radius:6px;margin-bottom:1.5rem;"><h3 style="margin:0;color:#CAEEC8;">⚠️ SOC Shift Compliance Issues</h3><p style="margin:5px 0 0;color:white;">{" · ".join(_banner)}</p></div>', unsafe_allow_html=True)

                st.markdown("### Shift Compliance — SOC Team")
                _sc_pf_col, _sc_ex_col = st.columns([4, 1])
                with _sc_pf_col: _sc_f, _sc_period = period_filter(_sc.copy(), 'Date', 'Date_dt_col', 'soc_comp_period')

                if not _sc_f.empty:
                    _sc_display = _sc_f[['Personnel ID', 'Full Name', 'Date_Str', 'Shift_Title', 'Sched_Start_Str', 'Actual_In_Str', 'Status', 'Mins Late']].rename(columns={'Date_Str': 'Shift Date', 'Shift_Title': 'Shift', 'Sched_Start_Str': 'Scheduled Start', 'Actual_In_Str': 'Actual Check-in'}).sort_values(['Personnel ID', 'Shift Date'])
                    st.dataframe(_sc_display, hide_index=True, use_container_width=True)
                    with _sc_ex_col: render_export_dropdown(_sc_display, _sc_period, "SOC Team Shift Compliance", "SOC personnel attendance.", [18, 40, 14, 14, 16, 16, 14, 12], "soc_compliance")

                st.divider()
                st.markdown("### Compliance Summary — SOC Team")
                _soc_summary = _soc_comp_all.groupby(['Personnel ID', 'Full Name']).agg(Scheduled_Shifts=('Shift Date', 'count'), On_Time=('Status', lambda x: (x == '✅ On Time ').sum()), Late=('Status', lambda x: x.str.startswith('⚠️').sum()), Absent=('Status', lambda x: (x == '🚫 Absent ').sum()), Total_Mins_Late=('Mins Late', 'sum')).reset_index().sort_values(['Absent', 'Late'], ascending=False)
                _soc_summary.columns = ['Personnel ID', 'Full Name', 'Scheduled Shifts', 'On Time', 'Late', 'Absent', 'Total Mins Late']
                st.dataframe(_soc_summary, hide_index=True, use_container_width=True)
                render_export_dropdown(_soc_summary, "All Time", "SOC Team Summary", "Per-person SOC overview.", [18, 40, 20, 14, 14, 14, 20], "soc_summary")

    except Exception as e:
        st.error(f"An error occurred while processing: {e}")
