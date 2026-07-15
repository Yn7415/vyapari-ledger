import os
import io
import datetime
import pandas as pd
import streamlit as strlt
from supabase import create_client, Client

# Safe PDF Library Import
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    pdf_available = True
except ImportError:
    pdf_available = False

# Initialize speech recognition safely
try:
    import speech_recognition as sr
    speech_available = True
except ImportError:
    speech_available = False

strlt.set_page_config(page_title="Smart Vyapari Ledger Pro+", page_icon="💎", layout="wide")

# ==========================================================
# 🔌 SUPABASE CONFIGURATION (Secrets check & app stop on missing)
# ==========================================================
if "SUPABASE_URL" not in strlt.secrets or "SUPABASE_KEY" not in strlt.secrets:
    strlt.error("❌ Critical Error: SUPABASE_URL ya SUPABASE_KEY missing hai Streamlit Secrets me! App ko rok diya gaya hai.")
    strlt.stop()

SUPABASE_URL = strlt.secrets["SUPABASE_URL"]
SUPABASE_KEY = strlt.secrets["SUPABASE_KEY"]

@strlt.cache_resource
def init_supabase() -> Client:
    try:
        return create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception as e:
        strlt.error(f"❌ Supabase Client initialization fail ho gaya: {str(e)}")
        strlt.stop()

supabase = init_supabase()

# Helper function to safely convert values to integer
def safe_int(val):
    try:
        if pd.isna(val) or val == "" or val is None:
            return 0
        return int(float(val))
    except (ValueError, TypeError):
        return 0

# ==========================================================
# 📊 DATA MANAGEMENT CORE (Optimized Storage & Queries)
# ==========================================================
def load_data():
    try:
        dash_res = supabase.table("dashboard").select("*").execute()
        df_d = pd.DataFrame(dash_res.data)
        if df_d.empty:
            df_d = pd.DataFrame(columns=["Naam", "Mobile", "Total Maal", "Total Jama", "Baki Balance"])
        else:
            df_d["Mobile"] = df_d["Mobile"].astype(str).str.replace(r'\.0$', '', regex=True)
            
        trans_res = supabase.table("transactions").select("*").execute()
        df_t = pd.DataFrame(trans_res.data)
        if df_t.empty:
            df_t = pd.DataFrame(columns=["ID", "Tarikh", "Naam", "Type", "Amount", "Note"])
        else:
            df_t["ID"] = df_t["ID"].astype(int)
            if "Note" not in df_t.columns:
                df_t["Note"] = ""
                
        df_t["Tarikh"] = pd.to_datetime(
            df_t["Tarikh"],
            format="%Y-%m-%d",
            errors="coerce"
        )
        return df_d, df_t
    except Exception as e:
        strlt.error(f"❌ Database se data load karne me samasya aayi: {str(e)}")
        return pd.DataFrame(columns=["Naam", "Mobile", "Total Maal", "Total Jama", "Baki Balance"]), pd.DataFrame(columns=["ID", "Tarikh", "Naam", "Type", "Amount", "Note"])

def recalculate_all(vyapari_name=None):
    try:
        if vyapari_name:
            t_maal = df_trans[(df_trans["Naam"] == vyapari_name) & (df_trans["Type"] == "Maal")]["Amount"].sum()
            t_jama = df_trans[(df_trans["Naam"] == vyapari_name) & (df_trans["Type"] == "Jama")]["Amount"].sum()
            
            idx_list = df_dash[df_dash["Naam"] == vyapari_name].index
            if not idx_list.empty:
                idx = idx_list[0]
                df_dash.at[idx, "Total Maal"] = safe_int(t_maal)
                df_dash.at[idx, "Total Jama"] = safe_int(t_jama)
                df_dash.at[idx, "Baki Balance"] = safe_int(t_maal - t_jama)
                
                dash_row = df_dash.iloc[idx].to_dict()
                dash_row["Mobile"] = str(dash_row["Mobile"])
                supabase.table("dashboard").upsert(dash_row, on_conflict="Naam").execute()
        else:
            for idx, row in df_dash.iterrows():
                v_name = row["Naam"]
                t_maal = df_trans[(df_trans["Naam"] == v_name) & (df_trans["Type"] == "Maal")]["Amount"].sum()
                t_jama = df_trans[(df_trans["Naam"] == v_name) & (df_trans["Type"] == "Jama")]["Amount"].sum()
                df_dash.at[idx, "Total Maal"] = safe_int(t_maal)
                df_dash.at[idx, "Total Jama"] = safe_int(t_jama)
                df_dash.at[idx, "Baki Balance"] = safe_int(t_maal - t_jama)
            
            dash_records = df_dash.copy().fillna("").to_dict(orient="records")
            for r in dash_records:
                r["Mobile"] = str(r["Mobile"])
            if dash_records:
                supabase.table("dashboard").upsert(dash_records, on_conflict="Naam").execute()
    except Exception as e:
        strlt.error(f"❌ Calculation auto sync fail ho gaya: {str(e)}")

# Primary Data State Matrix Init
df_dash, df_trans = load_data()

# --- State Management ---
if "selected_vyapari" not in strlt.session_state:
    strlt.session_state.selected_vyapari = None
if "editing_id" not in strlt.session_state:
    strlt.session_state.editing_id = None
if "editing_vyapari_name" not in strlt.session_state:
    strlt.session_state.editing_vyapari_name = None

if "delete_confirm_target" not in strlt.session_state:
    strlt.session_state.delete_confirm_target = None
if "last_deleted_vyapari" not in strlt.session_state:
    strlt.session_state.last_deleted_vyapari = None
if "last_deleted_trans" not in strlt.session_state:
    strlt.session_state.last_deleted_trans = None

# --- PDF Generator Function ---
def generate_ledger_pdf(v_name, v_row, df_m, df_j):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=22, textColor=colors.HexColor('#1a73e8'), spaceAfter=10)
    sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontSize=12, textColor=colors.gray, spaceAfter=20)
    section_style = ParagraphStyle('SectionStyle', parent=styles['Heading2'], fontSize=14, textColor=colors.HexColor('#333333'), spaceBefore=15, spaceAfter=10)
    bold_summary = ParagraphStyle('BoldSummary', parent=styles['Normal'], fontSize=12, fontName='Helvetica-Bold')

    story.append(Paragraph(f"Vyapari Khata Statement: {v_name}", title_style))
    story.append(Paragraph(f"Mobile: {v_row['Mobile']} | Date Generated: {datetime.date.today().strftime('%d-%m-%Y')}", sub_style))
    story.append(Spacer(1, 10))
    
    story.append(Paragraph("<b>📊 TRANSACTION DETAILS</b>", section_style))
    
    list_m = df_m.reset_index(drop=True)
    list_j = df_j.reset_index(drop=True)
    max_len = max(len(list_m), len(list_j))
    
    table_data = [["📦 MAAL LIYA HISTORY", "", "💰 JAMA PAISA HISTORY", ""]]
    table_data.append(["Tarikh", "Amount (₹)", "Tarikh", "Amount (₹)"])
    
    for i in range(max_len):
        row_cells = ["", "", "", ""]
        if i < len(list_m):
            row_m = list_m.iloc[i]
            dt_val = row_m['Tarikh'].strftime('%d-%m-%Y') if pd.notnull(row_m['Tarikh']) else "N/A"
            row_cells[0] = dt_val
            row_cells[1] = f"{safe_int(row_m['Amount']):,}"
        if i < len(list_j):
            row_j = list_j.iloc[i]
            dt_val = row_j['Tarikh'].strftime('%d-%m-%Y') if pd.notnull(row_j['Tarikh']) else "N/A"
            note_txt = f" ({row_j['Note']})" if pd.notnull(row_j['Note']) and row_j['Note'] != "" else ""
            row_cells[2] = f"{dt_val}{note_txt}"
            row_cells[3] = f"{safe_int(row_j['Amount']):,}"
        table_data.append(row_cells)
        
    t = Table(table_data, colWidths=[110, 110, 150, 110])
    t.setStyle(TableStyle([
        ('SPAN', (0, 0), (1, 0)),
        ('SPAN', (2, 0), (3, 0)),
        ('BACKGROUND', (0, 0), (1, 0), colors.HexColor('#e8f0fe')),
        ('BACKGROUND', (2, 0), (3, 0), colors.HexColor('#e6f4ea')),
        ('TEXTCOLOR', (0, 0), (1, 0), colors.HexColor('#1a73e8')),
        ('TEXTCOLOR', (2, 0), (3, 0), colors.HexColor('#137333')),
        ('FONTNAME', (0, 0), (3, 1), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
    ]))
    story.append(t)
    story.append(Spacer(1, 25))
    
    story.append(Paragraph("<b>📈 KHATA SUMMARY</b>", section_style))
    summary_data = [
        [Paragraph("<b>Total Maal Business:</b>", styles['Normal']), f"₹ {safe_int(v_row['Total Maal']):,}"],
        [Paragraph("<b>Total Jama Received:</b>", styles['Normal']), f"₹ {safe_int(v_row['Total Jama']):,}"],
        [Paragraph("<b>Total Baki Balance (Outstanding):</b>", bold_summary), f"₹ {safe_int(v_row['Baki Balance']):,}"]
    ]
    st_table = Table(summary_data, colWidths=[200, 150])
    st_table.setStyle(TableStyle([
        ('LINEBELOW', (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ('FONTNAME', (1, 0), (1, -1), 'Helvetica-Bold'),
        ('TEXTCOLOR', (1, 2), (1, 2), colors.HexColor('#c5221f')),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(st_table)
    
    doc.build(story)
    buffer.seek(0)
    return buffer

# --- Voice Command Logic ---
def listen_voice_command():
    if not speech_available:
        strlt.error("Speech Recognition library available nahi hai.")
        return ""
    r = sr.Recognizer()
    with sr.Microphone() as source:
        try:
            r.adjust_for_ambient_noise(source, duration=1)
            audio = r.listen(source, timeout=4, phrase_time_limit=5)
            text = r.recognize_google(audio, language="hi-IN")
            return text
        except Exception as e:
            strlt.error(f"🎙️ Mic capture optimization check fail: {str(e)}")
            return ""

# ===== UI PART 1 START =====
strlt.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800&display=swap');
    
    html, body, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        font-family: 'Poppins', sans-serif !important;
        background-color: #0F172A !important;
        color: #FFFFFF !important;
    }
    
    .hero-box {
        background: linear-gradient(135deg, #1E3A8A 0%, #0F766E 50%, #111827 100%);
        padding: 40px 30px;
        border-radius: 18px;
        text-align: center;
        color: #FFFFFF;
        margin-bottom: 25px;
        border: 1px solid rgba(255, 255, 255, 0.08);
        box-shadow: 0 15px 35px rgba(0, 0, 0, 0.4);
    }
    
    .hero-box h1 {
        font-size: 38px;
        font-weight: 700;
        margin: 0 0 10px 0 !important;
        color: #FFFFFF !important;
        letter-spacing: -0.5px;
    }
    
    .hero-box p {
        font-size: 16px;
        font-weight: 400;
        color: #94A3B8;
        margin: 0 !important;
    }
    
    .metric-card {
        background: rgba(30, 41, 59, 0.65);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        padding: 24px;
        border-radius: 18px;
        border: 1px solid rgba(255, 255, 255, 0.08);
        box-shadow: 0 8px 24 rgba(0, 0, 0, 0.2);
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    
    .metric-card:hover {
        transform: translateY(-4px);
        background: rgba(30, 41, 59, 0.85);
        border: 1px solid rgba(255, 255, 255, 0.15);
        box-shadow: 0 12px 30px rgba(0, 0, 0, 0.35);
    }
    
    .metric-title {
        font-size: 13px;
        font-weight: 600;
        color: #94A3B8;
        letter-spacing: 0.5px;
        margin-bottom: 6px;
    }
    
    .metric-value {
        font-size: 30px;
        font-weight: 700;
        color: #FFFFFF;
        letter-spacing: -0.5px;
    }
    
    .premium-vyapari-card {
        background: rgba(30, 41, 59, 0.5) !important;
        border-radius: 18px !important;
        border: 1px solid rgba(255, 255, 255, 0.06) !important;
        padding: 24px !important;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.2) !important;
        transition: all 0.3s ease-in-out !important;
        margin-bottom: 16px !important;
    }
    
    .premium-vyapari-card:hover {
        transform: translateY(-5px) scale(1.005) !important;
        background: rgba(30, 41, 59, 0.8) !important;
        border: 1px solid rgba(59, 130, 246, 0.3) !important;
        box-shadow: 0 12px 40px rgba(0, 0, 0, 0.4) !important;
    }

    .v-title-large { font-size: 24px; font-weight: 700; color: #FFFFFF; margin-bottom: 2px; }
    .v-sub-mobile { font-size: 14px; color: #94A3B8; font-weight: 500; }
    
    .v-grid-stats { font-size: 15px; font-weight: 600; color: #CBD5E1; }
    .v-baki-highlight { font-size: 20px; font-weight: 700; color: #EF4444; }

    .lbl-highlight { 
        font-size: 15px; 
        font-weight: 600; 
        padding: 5px 12px; 
        border-radius: 6px; 
        display: inline-block;
    }
    
    .lbl-settled { background-color: rgba(16, 185, 129, 0.15) !important; color: #10B981 !important; border: 1px solid #10B981 !important; }
    .lbl-maal { background-color: rgba(59, 130, 246, 0.15) !important; color: #3B82F6 !important; border: 1px solid #3B82F6 !important; }
    
    .lbl-j-slab-1 { background-color: rgba(59, 130, 246, 0.12) !important; color: #3B82F6 !important; border: 1px solid #3B82F6 !important; }
    .lbl-j-slab-2 { background-color: rgba(245, 158, 11, 0.12) !important; color: #F59E0B !important; border: 1px solid #F59E0B !important; }
    .lbl-j-slab-3 { background-color: rgba(168, 85, 247, 0.12) !important; color: #A855F7 !important; border: 1px solid #A855F7 !important; }
    .lbl-j-slab-4 { background-color: rgba(6, 182, 212, 0.12) !important; color: #06B6D4 !important; border: 1px solid #06B6D4 !important; }
    .lbl-j-slab-5 { background-color: rgba(239, 68, 68, 0.12) !important; color: #EF4444 !important; border: 1px solid #EF4444 !important; }
    
    .lbl-date { font-size: 14px; color: #E2E8F0; font-weight: 500; }
    .undo-banner { background-color: #1E293B; color: #F59E0B; padding: 14px; border-radius: 14px; font-weight: 600; text-align: center; border: 1px solid rgba(245, 158, 11, 0.3); }
    .alert-inline-banner { background-color: rgba(220, 38, 38, 0.12); color: #EF4444; padding: 14px; border-radius: 14px; font-weight: 600; border: 1px solid rgba(239, 68, 68, 0.3); }
    .alert-tag { background-color: #EF4444; color: white; padding: 4px 8px; border-radius: 6px; font-size: 11px; font-weight: bold; }
    .editing-badge { background-color: rgba(245, 158, 11, 0.15); color: #F59E0B; border: 1px solid rgba(245, 158, 11, 0.4); font-weight: 700; padding: 6px 14px; border-radius: 10px; margin-bottom: 12px; display: inline-block; }

    /* REQUESTED INPUT BOX COMPACT WRAPPER CHANGES */
    div[data-testid="stTextInput"]{
        margin-top:0px !important;
    }
    div[data-testid="stTextInput"] input{
        height:34px !important;
        width:95px !important;
        min-width:95px !important;
        max-width:95px !important;
        padding:4px 8px !important;
        text-align:left !important;
    }

    [data-testid="stForm"], div[data-testid="stExpander"], .stAlert {
        background-color: #1E293B !important;
        border-radius: 18px !important;
        border: 1px solid rgba(255, 255, 255, 0.08) !important;
        box-shadow: 0 4px 25px rgba(0,0,0,0.2) !important;
    }
    
    button[kind="primary"], button[kind="secondary"] {
        border-radius: 10px !important;
        padding: 8px 14px !important;
        font-weight: 600 !important;
        font-size: 13px !important;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }
    button[kind="primary"] { background-color: #3B82F6 !important; border: none !important; color: white !important; }
    button[kind="primary"]:hover { background-color: #2563EB !important; transform: translateY(-1px); }
    button[kind="secondary"] { background-color: rgba(255, 255, 255, 0.06) !important; color: #FFFFFF !important; border: 1px solid rgba(255, 255, 255, 0.1) !important; }
    button[kind="secondary"]:hover { background-color: rgba(255, 255, 255, 0.12) !important; }
    .no-data-text { text-align: center; padding: 40px; color: #64748B; font-size: 16px; font-weight: 500; }
    </style>
""", unsafe_allow_html=True)
# ===== UI PART 1 END =====

has_deleted_vyapari = strlt.session_state.last_deleted_vyapari is not None
has_deleted_trans = (strlt.session_state.last_deleted_trans is not None and not strlt.session_state.last_deleted_trans.empty)

# ==========================================
# SCREEN 2: SMART LEDGER DETAIL PAGE
# ==========================================
if strlt.session_state.selected_vyapari is not None:
    v_name = strlt.session_state.selected_vyapari
    
    match = df_dash[df_dash["Naam"] == v_name]
    if match.empty:
        strlt.error("Vyapari nahi mila.")
        strlt.stop()
    v_row = match.iloc[0]
    
    col_top1, col_top2 = strlt.columns([7, 3])
    with col_top1:
        if strlt.button("⬅️ Back to Main Dashboard Home"):
            strlt.session_state.selected_vyapari = None
            strlt.session_state.editing_id = None
            strlt.rerun()
            
    df_trans["Tarikh"] = pd.to_datetime(df_trans["Tarikh"], errors="coerce")
    df_m = df_trans[(df_trans["Naam"] == v_name) & (df_trans["Type"] == "Maal")].sort_values(by="Tarikh", ascending=True)
    df_j = df_trans[(df_trans["Naam"] == v_name) & (df_trans["Type"] == "Jama")].sort_values(by="Tarikh", ascending=False)

    # FIFO Bill Chukta Advanced Logic
    settled_trans_ids = set()
    maal_list = df_m[["ID", "Amount"]].to_dict('records')
    jama_list = df_j[["ID", "Amount"]].to_dict('records')
    jama_list_fifo = sorted(jama_list, key=lambda x: x['ID']) 
    
    j_idx = 0
    remaining_jama = jama_list_fifo[j_idx]['Amount'] if jama_list_fifo else 0

    for m_entry in maal_list:
        m_amt = m_entry['Amount']
        temp_used_pieces = []
        while m_amt > 0 and j_idx < len(jama_list_fifo):
            if remaining_jama <= m_amt:
                m_amt -= remaining_jama
                temp_used_pieces.append(jama_list_fifo[j_idx]['ID'])
                j_idx += 1
                if j_idx < len(jama_list_fifo): remaining_jama = jama_list_fifo[j_idx]['Amount']
            else:
                remaining_jama -= m_amt
                temp_used_pieces.append(jama_list_fifo[j_idx]['ID'])
                m_amt = 0
        if m_amt == 0:
            settled_trans_ids.add(m_entry['ID'])
            for j_id in temp_used_pieces: settled_trans_ids.add(j_id)

    with col_top2:
        if pdf_available:
            pdf_data = generate_ledger_pdf(v_name, v_row, df_m, df_j)
            strlt.download_button(label="📥 Download PDF Layout Report", data=pdf_data, file_name=f"{v_name}_khata_report.pdf", mime="application/pdf")

    strlt.markdown(f"<div class='hero-box'><h1 style='font-size: 42px;'>👤 {v_name}</h1><p style='font-size: 18px; color: #CBD5E1;'>📞 Mobile: {v_row['Mobile']}</p></div>", unsafe_allow_html=True)
    
    m1, m2, m3 = strlt.columns(3)
    m1.markdown(f"<div class='metric-card'><div class='metric-title'>📦 KUL MAAL (TOTAL BUSINESS)</div><div class='metric-value' style='color:#3B82F6;'>₹ {safe_int(v_row['Total Maal']):,}</div></div>", unsafe_allow_html=True)
    m2.markdown(f"<div class='metric-card'><div class='metric-title'>💰 KUL JAMA (TOTAL RECEIVED)</div><div class='metric-value' style='color:#10B981;'>₹ {safe_int(v_row['Total Jama']):,}</div></div>", unsafe_allow_html=True)
    m3.markdown(f"<div class='metric-card' style='background: rgba(239, 68, 68, 0.08);'><div class='metric-title' style='color: #FCA5A5;'>🔴 CURRENT BALANCE (TOTAL BAKI)</div><div class='metric-value' style='color:#EF4444;'>₹ {safe_int(v_row['Baki Balance']):,}</div></div>", unsafe_allow_html=True)
    
    # --- Single Vyapari Pie Chart Panel ---
    strlt.markdown("<br>", unsafe_allow_html=True)
    if "show_vyapari_pie" not in strlt.session_state:
        strlt.session_state.show_vyapari_pie = False
        
    if strlt.button(f"📊 Toggle {v_name} Monthly Len-Den Pie Chart Panel", type="primary", use_container_width=True):
        strlt.session_state.show_vyapari_pie = not strlt.session_state.show_vyapari_pie
        strlt.rerun()
        
    if strlt.session_state.show_vyapari_pie:
        df_v_all = df_trans[df_trans["Naam"] == v_name].copy()
        if not df_v_all.empty:
            df_v_all["Month_Key"] = df_v_all["Tarikh"].dt.strftime("%b %Y")
            df_pie_data = df_v_all.groupby("Month_Key")["Amount"].sum()
            
            strlt.markdown(f"<div class='metric-card'><div class='metric-title'>📈 Monthly Len-Den Share for {v_name}</div>", unsafe_allow_html=True)
            strlt.bar_chart(df_pie_data, color="#3B82F6")
            strlt.markdown("</div>", unsafe_allow_html=True)
        else:
            strlt.markdown("<div class='metric-card'><p class='no-data-text'>📊 No Transaction Data Available to Plot Pie Chart</p></div>", unsafe_allow_html=True)

    strlt.markdown("<br>", unsafe_allow_html=True)

    # --- FORM INPUT ---
    strlt.subheader("✏️ Entry Form (Amount daalkar Enter dabayein)")
    if strlt.session_state.editing_id:
        strlt.markdown("<div class='editing-badge'>🟡 Editing Transaction Pipeline Mode Active</div>", unsafe_allow_html=True)

    voice_trans_text = ""
    if strlt.button("🎙️ Speak Transaction Entry"):
        with strlt.spinner("Aap boliye..."):
            voice_trans_text = listen_voice_command()
            if voice_trans_text: strlt.info(f"Suna: '{voice_trans_text}'")
            
    v_amt_val = None
    v_type_val = "Maal"
    if voice_trans_text:
        words = voice_trans_text.lower().split()
        for w in words:
            if w.isdigit(): v_amt_val = safe_int(w)
        if "jama" in voice_trans_text or "paisa" in voice_trans_text: v_type_val = "Jama"

    with strlt.form(key="ledger_entry_form", clear_on_submit=False):
        col_date_block, col_amt, col_tp = strlt.columns([2.5, 4, 2.5])
        today = datetime.date.today()
        if strlt.session_state.editing_id:
            edit_match = df_trans[df_trans["ID"] == strlt.session_state.editing_id]
            if edit_match.empty:
                strlt.error("Edit karne ke liye transaction nahi mila.")
                strlt.stop()
            edit_row = edit_match.iloc[0]
            ref_date = edit_row["Tarikh"] if pd.notnull(edit_row["Tarikh"]) else today
        else:
            ref_date = today
            
        day_options = [str(i).zfill(2) for i in range(1, 32)]
        month_options = [str(i).zfill(2) for i in range(1, 13)]
        year_options = [str(i) for i in range(2020, today.year + 1)]
        
        try: def_day_idx = day_options.index(str(ref_date.day).zfill(2))
        except ValueError: def_day_idx = today.day - 1
        try: def_mon_idx = month_options.index(str(ref_date.month).zfill(2))
        except ValueError: def_mon_idx = today.month - 1
        try: def_year_idx = year_options.index(str(ref_date.year))
        except ValueError: def_year_idx = len(year_options) - 1

        with col_date_block:
            strlt.markdown("<span style='font-size:14px; font-weight:500; color:gray;'>Tarikh (DD-MM-YYYY)</span>", unsafe_allow_html=True)
            d_sub1, d_sub2, d_sub3 = strlt.columns(3)
            selected_dd = d_sub1.selectbox("DD", day_options, index=def_day_idx, label_visibility="collapsed")
            selected_mm = d_sub2.selectbox("MM", month_options, index=def_mon_idx, label_visibility="collapsed")
            selected_yyyy = d_sub3.selectbox("YYYY", year_options, index=def_year_idx, label_visibility="collapsed")
        
        final_constructed_date_str = f"{selected_dd}-{selected_mm}-{selected_yyyy}"
        
        if strlt.session_state.editing_id:
            e_amt = col_amt.number_input("Amount (₹)", min_value=0, value=safe_int(edit_row["Amount"]))
            e_type = col_tp.selectbox("Entry Type", ["Maal", "Jama"], index=0 if edit_row["Type"] == "Maal" else 1)
        else:
            e_amt = col_amt.number_input("Amount (₹)", min_value=0, value=v_amt_val, placeholder="Type amount...")
            e_type = col_tp.selectbox("Entry Type", ["Maal", "Jama"], index=0 if v_type_val == "Maal" else 1)

        btn_label = "💾 Update Transaction" if strlt.session_state.editing_id else "⚡ Save Record"
        submit_ledger = strlt.form_submit_button(label=btn_label, type="primary")

    if submit_ledger:
        if e_amt is None or safe_int(e_amt) == 0: strlt.error("Kripya valid Amount enter karein!")
        else:
            try: parsed_date = datetime.datetime.strptime(final_constructed_date_str, "%d-%m-%Y")
            except Exception: parsed_date = pd.NaT
            if pd.isna(parsed_date): strlt.error("Chuni hui tarikh calendar ke hisab se valid nahi h!")
            elif parsed_date.date() > datetime.date.today(): strlt.error("Aage ki future tarikh block h!")
            else:
                try:
                    if strlt.session_state.editing_id:
                        tr_id = int(strlt.session_state.editing_id)
                        payload = {"Tarikh": parsed_date.strftime("%Y-%m-%d"), "Type": e_type, "Amount": safe_int(e_amt)}
                        supabase.table("transactions").update(payload).eq("ID", tr_id).execute()
                        idx = df_trans[df_trans["ID"] == tr_id].index[0]
                        df_trans.at[idx, "Tarikh"] = parsed_date
                        df_trans.at[idx, "Type"] = e_type
                        df_trans.at[idx, "Amount"] = safe_int(e_amt)
                        strlt.session_state.editing_id = None
                    else:
                        payload = {"Tarikh": parsed_date.strftime("%Y-%m-%d"), "Naam": v_name, "Type": e_type, "Amount": safe_int(e_amt), "Note": ""}
                        insert_res = supabase.table("transactions").insert(payload).execute()
                        if insert_res.data:
                            new_row_db = insert_res.data[0]
                            new_row_db["Tarikh"] = pd.to_datetime(new_row_db["Tarikh"])
                            df_trans = pd.concat([df_trans, pd.DataFrame([new_row_db])], ignore_index=True)
                    recalculate_all(vyapari_name=v_name)
                    strlt.rerun()
                except Exception as ex: strlt.error(f"❌ Database error: {str(ex)}")

    strlt.markdown("---")

    # ===== UI PART 2 START =====
    col_l, col_r = strlt.columns(2, gap="large")
    
    with col_l:
        strlt.markdown("### 📦 MAAL LIYA HISTORY")
        for _, row in df_m.iterrows():
            formatted_date = row["Tarikh"].strftime("%d-%m-%Y") if pd.notnull(row["Tarikh"]) else ""
            amt_val = safe_int(row['Amount'])
            
            style_class = "lbl-settled" if row['ID'] in settled_trans_ids else "lbl-maal"
            is_this_target = (strlt.session_state.delete_confirm_target is not None and strlt.session_state.delete_confirm_target["scope"] == "transaction" and strlt.session_state.delete_confirm_target["key"] == row["ID"])
            
            with strlt.container():
                r_amt, r_date, r_ed, r_del = strlt.columns([1.3, 1.8, 0.55, 0.55], vertical_alignment="center")
                r_amt.markdown(f"<div><span class='lbl-highlight {style_class}'>₹ {amt_val:,}</span></div>", unsafe_allow_html=True)
                r_date.markdown(f"<div style='padding-top:2px;'><span class='lbl-date'>📅 {formatted_date}</span></div>", unsafe_allow_html=True)
                
                if r_ed.button("✏️", key=f"ed_{row['ID']}", use_container_width=True, type="secondary"):
                    strlt.session_state.editing_id = row["ID"]
                    strlt.rerun()
                if r_del.button("❌", key=f"del_{row['ID']}", use_container_width=True, type="secondary"):
                    strlt.session_state.delete_confirm_target = {"scope": "transaction", "key": row["ID"]}
                    strlt.rerun()
                    
                if is_this_target:
                    strlt.markdown("<div class='alert-inline-banner'>⚠️ Mitaana chahte hain?</div>", unsafe_allow_html=True)
                    b1, b2 = strlt.columns(2)
                    if b1.button("✅ HAAN", key=f"yes_t_{row['ID']}", type="primary", use_container_width=True):
                        supabase.table("transactions").delete().eq("ID", int(row["ID"])).execute()
                        df_trans = df_trans[df_trans["ID"] != int(row["ID"])]
                        recalculate_all(vyapari_name=v_name)
                        strlt.session_state.delete_confirm_target = None
                        strlt.rerun()
                    if b2.button("❌ ROKO", key=f"no_t_{row['ID']}", use_container_width=True):
                        strlt.session_state.delete_confirm_target = None
                        strlt.rerun()

    with col_r:
        strlt.markdown("### 💰 HAR HAFTE JAMA PAISA")
        for _, row in df_j.iterrows():
            formatted_date = row["Tarikh"].strftime("%d-%m-%Y") if pd.notnull(row["Tarikh"]) else ""
            amt_val = safe_int(row['Amount'])
            note_val = str(row['Note']) if pd.notnull(row['Note']) and str(row['Note']) != "nan" else ""
            is_this_target = (strlt.session_state.delete_confirm_target is not None and strlt.session_state.delete_confirm_target["scope"] == "transaction" and strlt.session_state.delete_confirm_target["key"] == row["ID"])
            
            if row['ID'] in settled_trans_ids: 
                jama_class = "lbl-settled"
            else:
                if amt_val <= 5000: jama_class = "lbl-j-slab-1"
                elif amt_val <= 10000: jama_class = "lbl-j-slab-2"
                elif amt_val <= 15000: jama_class = "lbl-j-slab-3"
                elif amt_val <= 20000: jama_class = "lbl-j-slab-4"
                else: jama_class = "lbl-j-slab-5"
                
            with strlt.container():
                # FIXED REQUESTED PROPORTIONS: [1.3, 1.8, 0.55, 0.55, 1.1] with vertical center alignment
                r_amt, r_date, r_ed, r_del, r_nt = strlt.columns([1.3, 1.8, 0.55, 0.55, 1.1], vertical_alignment="center")
                
                r_amt.markdown(f"<div><span class='lbl-highlight {jama_class}'>₹ {amt_val:,}</span></div>", unsafe_allow_html=True)
                r_date.markdown(f"<div style='padding-top:2px;'><span class='lbl-date'>📅 {formatted_date}</span></div>", unsafe_allow_html=True)
                
                if r_ed.button("✏️", key=f"ed_j_{row['ID']}", use_container_width=True, type="secondary"):
                    strlt.session_state.editing_id = row["ID"]
                    strlt.rerun()
                if r_del.button("❌", key=f"del_j_{row['ID']}", use_container_width=True, type="secondary"):
                    strlt.session_state.delete_confirm_target = {"scope": "transaction", "key": row["ID"]}
                    strlt.rerun()
                    
                with r_nt:
                    updated_note = strlt.text_input(
                        "",
                        value=note_val,
                        key=f"nt_{row['ID']}",
                        label_visibility="collapsed",
                        placeholder="Note",
                    )
                    
                if updated_note != note_val:
                    supabase.table("transactions").update({"Note": str(updated_note).strip()}).eq("ID", int(row["ID"])).execute()
                    idx_list = df_trans[df_trans["ID"] == int(row["ID"])].index
                    if not idx_list.empty: df_trans.at[idx_list[0], "Note"] = str(updated_note).strip()
                    strlt.rerun()
                    
                if is_this_target:
                    strlt.markdown("<div class='alert-inline-banner'>⚠️ Mitaana chahte hain?</div>", unsafe_allow_html=True)
                    b1, b2 = strlt.columns(2)
                    if b1.button("✅ HAAN", key=f"yes_j_{row['ID']}", type="primary", use_container_width=True):
                        supabase.table("transactions").delete().eq("ID", int(row["ID"])).execute()
                        df_trans = df_trans[df_trans["ID"] != int(row["ID"])]
                        recalculate_all(vyapari_name=v_name)
                        strlt.session_state.delete_confirm_target = None
                        strlt.rerun()
                    if b2.button("❌ ROKO", key=f"no_j_{row['ID']}", use_container_width=True):
                        strlt.session_state.delete_confirm_target = None
                        strlt.rerun()
    # ===== UI PART 2 END =====

# ==========================================
# SCREEN 1: MAIN DASHBOARD HOME SCREEN
# ==========================================
else:
    strlt.markdown("<div class='hero-box'><h1>💎 Smart Vyapari Ledger Pro+</h1><p>Digital Business Ledger Management Matrix System</p></div>", unsafe_allow_html=True)
    
    t_business = df_dash["Total Maal"].sum() if not df_dash.empty else 0
    t_received = df_dash["Total Jama"].sum() if not df_dash.empty else 0
    t_outstanding = df_dash["Baki Balance"].sum() if not df_dash.empty else 0
    
    card1, card2, card3, card4 = strlt.columns(4)
    card1.markdown(f"<div class='metric-card'><div class='metric-title'>📦 TOTAL MAAL</div><div class='metric-value' style='color:#3B82F6;'>₹ {safe_int(t_business):,}</div></div>", unsafe_allow_html=True)
    card2.markdown(f"<div class='metric-card'><div class='metric-title'>💰 TOTAL JAMA</div><div class='metric-value' style='color:#10B981;'>₹ {safe_int(t_received):,}</div></div>", unsafe_allow_html=True)
    card3.markdown(f"<div class='metric-card'><div class='metric-title'>🔴 TOTAL BAKI</div><div class='metric-value' style='color:#EF4444;'>₹ {safe_int(t_outstanding):,}</div></div>", unsafe_allow_html=True)
    card4.markdown(f"<div class='metric-card'><div class='metric-title'>👥 TOTAL VYAPARI</div><div class='metric-value' style='color:#06B6D4;'>{len(df_dash)}</div></div>", unsafe_allow_html=True)
    
    # --- Analytics Panel Component Section ---
    strlt.markdown("<br>", unsafe_allow_html=True)
    if "show_analytics" not in strlt.session_state:
        strlt.session_state.show_analytics = False
        
    if strlt.button("📊 Toggle 12-Month Financial Business Analytics Panel", type="primary", use_container_width=True):
        strlt.session_state.show_analytics = not strlt.session_state.show_analytics
        strlt.rerun()
        
    if strlt.session_state.show_analytics:
        strlt.markdown("<h3 style='margin-top:10px; font-weight:600;'>📊 12-Month Financial Business Analytics</h3>", unsafe_allow_html=True)
        if not df_trans.empty:
            one_year_ago = datetime.datetime.now() - datetime.timedelta(days=365)
            df_analytics_slice = df_trans[df_trans["Tarikh"] >= one_year_ago].copy()
            
            if not df_analytics_slice.empty:
                df_analytics_slice["Month_Key"] = df_analytics_slice["Tarikh"].dt.strftime("%b %Y")
                df_pivot = df_analytics_slice.groupby(["Month_Key", "Type"])["Amount"].sum().unstack().fillna(0)
                
                if "Maal" not in df_pivot.columns: df_pivot["Maal"] = 0
                if "Jama" not in df_pivot.columns: df_pivot["Jama"] = 0
                df_pivot["Baki"] = df_pivot["Maal"] - df_pivot["Jama"]
                
                ch1, ch2, ch3 = strlt.columns(3)
                with ch1:
                    strlt.markdown("<div class='metric-card'><div class='metric-title'>📈 Monthly Maal Matrix</div>", unsafe_allow_html=True)
                    strlt.bar_chart(df_pivot["Maal"], color="#3B82F6")
                    strlt.markdown("</div>", unsafe_allow_html=True)
                with ch2:
                    strlt.markdown("<div class='metric-card'><div class='metric-title'>📈 Monthly Jama Tracking</div>", unsafe_allow_html=True)
                    strlt.bar_chart(df_pivot["Jama"], color="#10B981")
                    strlt.markdown("</div>", unsafe_allow_html=True)
                with ch3:
                    strlt.markdown("<div class='metric-card'><div class='metric-title'>📈 Monthly Net Baki Breakdown</div>", unsafe_allow_html=True)
                    strlt.area_chart(df_pivot["Baki"], color="#EF4444")
                    strlt.markdown("</div>", unsafe_allow_html=True)
            else:
                strlt.markdown("<div class='metric-card'><p class='no-data-text'>📊 No Data Available For Last 12 Months</p></div>", unsafe_allow_html=True)
        else:
            strlt.markdown("<div class='metric-card'><p class='no-data-text'>📊 No Data Available For Last 12 Months</p></div>", unsafe_allow_html=True)

    strlt.markdown("<br>", unsafe_allow_html=True)
    search_term = strlt.text_input("🔍 Search Vyapari Log Matrix", placeholder="Type Vyapari ka Naam ya Mobile Number to instantly filter metrics locally...", label_visibility="collapsed")
    
    strlt.markdown("<br>", unsafe_allow_html=True)
    box_title = "👤 ➕ Naya Vyapari Jodein"
    init_name, init_mob = "", ""
    
    with strlt.expander(box_title, expanded=False):
        voice_v_text = ""
        if strlt.button("🎙️ Speak Vyapari Info"):
            with strlt.spinner("Aap boliye..."):
                voice_v_text = listen_voice_command()
                if voice_v_text: strlt.info(f"Suna: '{voice_v_text}'")
                
        if voice_v_text:
            v_words = voice_v_text.split()
            init_name = v_words[0] if len(v_words) > 0 else ""
            init_mob = v_words[1] if len(v_words) > 1 else ""

        with strlt.form(key="merchant_add_form", clear_on_submit=True):
            n_name = strlt.text_input("Vyapari Ka Naam", value=init_name)
            n_mob = strlt.text_input("Mobile Number", value=init_mob)
            submit_merchant = strlt.form_submit_button(label="💾 Account Detail Save Karein", type="primary")
        
        if submit_merchant:
            if n_name and n_mob:
                if not df_dash.empty and n_name in df_dash["Naam"].values: strlt.warning("Vyapari pehle se add hai.")
                else:
                    try:
                        new_profile = {"Naam": n_name, "Mobile": str(n_mob), "Total Maal": 0, "Total Jama": 0, "Baki Balance": 0}
                        supabase.table("dashboard").upsert(new_profile, on_conflict="Naam").execute()
                        df_dash = pd.concat([df_dash, pd.DataFrame([new_profile])], ignore_index=True)
                        strlt.rerun()
                    except Exception as ex: strlt.error(f"❌ Error: {str(ex)}")

    strlt.markdown("---")
    strlt.subheader("👥 Active Vyapari Ledger Logs")

    filtered_df_dash = df_dash.copy()
    if search_term:
        term = str(search_term).strip().lower()
        filtered_df_dash = filtered_df_dash[
            (filtered_df_dash["Naam"].astype(str).str.lower().str.contains(term, na=False)) |
            (filtered_df_dash["Mobile"].astype(str).str.lower().str.contains(term, na=False)) |
            (filtered_df_dash["Baki Balance"].astype(str).str.lower().str.contains(term, na=False)) |
            (filtered_df_dash["Total Maal"].astype(str).str.lower().str.contains(term, na=False)) |
            (filtered_df_dash["Total Jama"].astype(str).str.lower().str.contains(term, na=False))
        ]
        # YAHAN PASTE KAREIN
    if strlt.session_state.get("editing_vyapari_name"):
        with strlt.expander("🟡 Vyapari Details Edit Karein", expanded=True):
            edit_v_name = strlt.session_state.editing_vyapari_name
            match = df_dash[df_dash["Naam"] == edit_v_name]
            if not match.empty:
                current_data = match.iloc[0]
                with strlt.form("edit_v_form"):
                    n_name = strlt.text_input("Naya Naam", value=current_data["Naam"])
                    n_mob = strlt.text_input("Naya Mobile", value=current_data["Mobile"])
                    if strlt.form_submit_button("✅ Update Karein"):
                        supabase.table("dashboard").update({"Naam": n_name, "Mobile": str(n_mob)}).eq("Naam", edit_v_name).execute()
                        strlt.session_state.editing_vyapari_name = None
                        strlt.rerun()

    # AB ISKE NEECHE TUMHARA WOH "if not filtered_df_dash.empty:" WALA LOOP CHALEGA
    if not filtered_df_dash.empty:
     
        for idx, row in filtered_df_dash.iterrows():
            v_name = row["Naam"]
            serial_no = idx + 1
            
            df_v_trans = df_trans[df_trans["Naam"] == v_name] if not df_trans.empty else pd.DataFrame()
            df_v_maal = df_v_trans[df_v_trans["Type"] == "Maal"] if not df_v_trans.empty else pd.DataFrame()
            df_v_jama = df_v_trans[df_v_trans["Type"] == "Jama"] if not df_v_trans.empty else pd.DataFrame()
            
            alert_active = False
            if not df_v_maal.empty and safe_int(row["Baki Balance"]) > 0:
                last_maal_date = df_v_maal["Tarikh"].max()
                if not df_v_jama.empty:
                    last_jama_date = df_v_jama["Tarikh"].max()
                    if (datetime.datetime.now() - last_jama_date).days > 90 and last_maal_date > last_jama_date: alert_active = True
                else:
                    if (datetime.datetime.now() - last_maal_date).days > 90: alert_active = True
                        
            strlt.markdown(f"<div class='premium-vyapari-card'>", unsafe_allow_html=True)
            alert_html = "<span class='alert-tag'>🔴 OVERDUE</span>" if alert_active else ""
            c_info, c_stats, c_actions = strlt.columns([4.5, 3.5, 2])
            
            with c_info:
                strlt.markdown(f"<div class='v-title-large'>👤 {serial_no}. {v_name} {alert_html}</div><div class='v-sub-mobile'>📞 Mobile: {row['Mobile']}</div>", unsafe_allow_html=True)
            with c_stats:
                strlt.markdown(f"<div class='v-grid-stats'>📦 Maal: ₹{safe_int(row['Total Maal']):,} &nbsp;|&nbsp; 💵 Jama: ₹{safe_int(row['Total Jama']):,}</div><div class='v-baki-highlight'>🔴 Baki Outstanding: ₹{safe_int(row['Baki Balance']):,}</div>", unsafe_allow_html=True)
            with c_actions:
                btn_col1, btn_col2, btn_col3 = strlt.columns(3)
                if btn_col1.button("📖 Open", key=f"open_{v_name}", use_container_width=True, type="primary"):
                    strlt.session_state.selected_vyapari = v_name
                    strlt.rerun()
                if btn_col2.button("✏️", key=f"edit_v_{v_name}", use_container_width=True, type="secondary"):
                    strlt.session_state.editing_vyapari_name = v_name
                    strlt.rerun()
                if btn_col3.button("❌", key=f"del_v_{v_name}", use_container_width=True, type="secondary"):
                    strlt.session_state.delete_confirm_target = {"scope": "vyapari", "key": v_name}
                    strlt.rerun()
                    
            if (strlt.session_state.delete_confirm_target is not None and strlt.session_state.delete_confirm_target["scope"] == "vyapari" and strlt.session_state.delete_confirm_target["key"] == v_name):
                strlt.markdown(f"<div class='alert-inline-banner'>⚠️ Deletion Alert Profile Account '{v_name}' Purge Sequence Warning?</div>", unsafe_allow_html=True)
                v_btn1, v_btn2 = strlt.columns(2)
                if v_btn1.button("🔥 YES, PURGE RECS", key=f"y_v_{v_name}", type="primary", use_container_width=True):
                    supabase.table("dashboard").delete().eq("Naam", v_name).execute()
                    supabase.table("transactions").delete().eq("Naam", v_name).execute()
                    df_dash = df_dash[df_dash["Naam"] != v_name]
                    df_trans = df_trans[df_trans["Naam"] != v_name]
                    strlt.session_state.delete_confirm_target = None
                    strlt.rerun()
                if v_btn2.button("❌ CANCEL ROKO", key=f"n_v_{v_name}", use_container_width=True):
                    strlt.session_state.delete_confirm_target = None
                    strlt.rerun()
            strlt.markdown("</div>", unsafe_allow_html=True)
    else:
        strlt.info("Koi Vyapari data mila nahi.")

strlt.markdown("<div style='border-bottom:1px solid rgba(255,255,255,0.08); margin-bottom:12px; margin-top:4px;'></div>", unsafe_allow_html=True)