import os
import io
import datetime
import pandas as pd
import streamlit as strlt

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

EXCEL_FILE = "vyapari_hisab.xlsx"

# Helper function to safely convert values to integer
def safe_int(val):
    try:
        if pd.isna(val) or val == "" or val is None:
            return 0
        return int(float(val))
    except (ValueError, TypeError):
        return 0

# --- Data Management Core ---
def load_data():
    if not os.path.exists(EXCEL_FILE):
        df_dash = pd.DataFrame(columns=["Naam", "Mobile", "Total Maal", "Total Jama", "Baki Balance"])
        df_trans = pd.DataFrame(columns=["ID", "Tarikh", "Naam", "Type", "Amount", "Note"])
        with pd.ExcelWriter(EXCEL_FILE) as writer:
            df_dash.to_excel(writer, sheet_name="Dashboard", index=False)
            df_trans.to_excel(writer, sheet_name="Transactions", index=False)
    
    df_d = pd.read_excel(EXCEL_FILE, sheet_name="Dashboard")
    df_t = pd.read_excel(EXCEL_FILE, sheet_name="Transactions")
    
    df_d["Mobile"] = df_d["Mobile"].astype(str).str.replace(r'\.0$', '', regex=True)
    if "ID" not in df_t.columns:
        df_t.insert(0, "ID", range(1, len(df_t) + 1))
    if "Note" not in df_t.columns:
        df_t["Note"] = ""
        
    # ✅ Fix 1 Kept: Indian Date parsing configuration during Excel load[cite: 1]
    df_t["Tarikh"] = pd.to_datetime(
        df_t["Tarikh"],
        format="%d-%m-%Y",
        dayfirst=True,
        errors="coerce"
    )
    return df_d, df_t

def save_data(df_dash, df_trans):
    df_dash_save = df_dash.copy()
    df_trans_save = df_trans.copy()
    
    df_dash_save["Mobile"] = df_dash_save["Mobile"].astype(str)
    
    # ✅ Fix 4 Kept: Force parse as datetime object right before string format conversion[cite: 1]
    if pd.api.types.is_datetime64_any_dtype(df_trans_save["Tarikh"]) or True:
        df_trans_save["Tarikh"] = (
            pd.to_datetime(df_trans_save["Tarikh"])
            .dt.strftime("%d-%m-%Y")
        )
    
    with pd.ExcelWriter(EXCEL_FILE) as writer:
        df_dash_save.to_excel(writer, sheet_name="Dashboard", index=False)
        df_trans_save.to_excel(writer, sheet_name="Transactions", index=False)

def recalculate_all():
    df_dash, df_trans = load_data()
    for idx, row in df_dash.iterrows():
        v_name = row["Naam"]
        t_maal = df_trans[(df_trans["Naam"] == v_name) & (df_trans["Type"] == "Maal")]["Amount"].sum()
        t_jama = df_trans[(df_trans["Naam"] == v_name) & (df_trans["Type"] == "Jama")]["Amount"].sum()
        df_dash.at[idx, "Total Maal"] = safe_int(t_maal)
        df_dash.at[idx, "Total Jama"] = safe_int(t_jama)
        df_dash.at[idx, "Baki Balance"] = safe_int(t_maal - t_jama)
    save_data(df_dash, df_trans)

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

# --- PDF Generator ---
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
        except Exception:
            return ""

# --- UI Premium Styling ---
strlt.markdown("""
    <style>
    .main-title { font-size: 32px; font-weight: bold; color: #1a73e8; text-align: center; margin-bottom: 25px; }
    .kpi-box { padding: 22px; border-radius: 15px; text-align: center; font-weight: bold; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
    .kpi-business { background-color: #e8f0fe; color: #1a73e8; border: 1px solid #d2e3fc; }
    .kpi-received { background-color: #e6f4ea; color: #137333; border: 1px solid #ceead6; }
    .kpi-outstanding { background-color: #fce8e6; color: #c5221f; border: 1px solid #fad2cf; }
    .alert-tag { background-color: #ff4d4d; color: white; padding: 4px 8px; border-radius: 6px; font-size: 11px; font-weight: bold; animation: blinker 1.5s linear infinite; }
    @keyframes blinker { 50% { opacity: 0.4; } }
    
    .lbl-highlight { 
        font-size: 16px; 
        font-weight: bold; 
        padding: 5px 12px; 
        border-radius: 6px; 
        display: inline-block;
        margin-bottom: 3px;
    }
    .lbl-maal { background-color: #d6e4ff; color: #1a73e8 !important; border-left: 5px solid #1a73e8; }
    
    .lbl-j-1 { background-color: #e8f0fe; color: #1a73e8 !important; border-left: 5px solid #4285f4; }
    .lbl-j-2 { background-color: #fff3cd; color: #856404 !important; border-left: 5px solid #ffc107; }
    .lbl-j-3 { background-color: #f3e5f5; color: #4a148c !important; border-left: 5px solid #9c27b0; }
    .lbl-j-4 { background-color: #e0f7fa; color: #006064 !important; border-left: 5px solid #00bcd4; }
    .lbl-j-5 { background-color: #d4edda; color: #155724 !important; border: 2px solid #28a745; font-size: 17px; }
    
    .lbl-settled { background-color: #d4edda !important; color: #155724 !important; border: 2px solid #28a745 !important; }
    .lbl-date { font-size: 13px; color: #888888; font-weight: 500; margin-left: 8px; }
    </style>
""", unsafe_allow_html=True)

has_deleted_vyapari = strlt.session_state.last_deleted_vyapari is not None
has_deleted_trans = (strlt.session_state.last_deleted_trans is not None and not strlt.session_state.last_deleted_trans.empty)

# ==========================================
# SCREEN 2: SMART LEDGER DETAIL PAGE
# ==========================================
if strlt.session_state.selected_vyapari is not None:
    v_name = strlt.session_state.selected_vyapari
    v_row = df_dash[df_dash["Naam"] == v_name].iloc[0]
    
    col_top1, col_top2 = strlt.columns([7, 3])
    with col_top1:
        if strlt.button("⬅️ Back to Main Dashboard Home"):
            strlt.session_state.selected_vyapari = None
            strlt.session_state.editing_id = None
            strlt.rerun()
            
    # ✅ Fix 3 Kept: Safety alignment parser right before index matrix layout splitting[cite: 1]
    df_trans["Tarikh"] = pd.to_datetime(
        df_trans["Tarikh"],
        format="%d-%m-%Y",
        errors="coerce",
        dayfirst=True
    )
            
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
                if j_idx < len(jama_list_fifo):
                    remaining_jama = jama_list_fifo[j_idx]['Amount']
            else:
                remaining_jama -= m_amt
                temp_used_pieces.append(jama_list_fifo[j_idx]['ID'])
                m_amt = 0
        if m_amt == 0:
            settled_trans_ids.add(m_entry['ID'])
            for j_id in temp_used_pieces:
                settled_trans_ids.add(j_id)

    with col_top2:
        if pdf_available:
            pdf_data = generate_ledger_pdf(v_name, v_row, df_m, df_j)
            strlt.download_button(label="📥 Download PDF Layout Report", data=pdf_data, file_name=f"{v_name}_khata_report.pdf", mime="application/pdf")

    strlt.markdown(f"<div class='main-title'>📊 Khata Statement: {v_name}</div>", unsafe_allow_html=True)
    strlt.caption(f"📱 Mobile: {v_row['Mobile']}")
    
    m1, m2, m3 = strlt.columns(3)
    m1.markdown(f"<div class='kpi-box kpi-business'><span style='font-size:13px;'>📦 KUL MAAL (TOTAL BUSINESS)</span><br><span style='font-size:22px;'>₹ {safe_int(v_row['Total Maal']):,}</span></div>", unsafe_allow_html=True)
    m2.markdown(f"<div class='kpi-box kpi-received'><span style='font-size:13px;'>💰 KUL JAMA (TOTAL RECEIVED)</span><br><span style='font-size:22px;'>₹ {safe_int(v_row['Total Jama']):,}</span></div>", unsafe_allow_html=True)
    m3.markdown(f"<div class='kpi-box kpi-outstanding'><span style='font-size:13px;'>🔴 TOTAL BAKI (OUTSTANDING)</span><br><span style='font-size:22px;'>₹ {safe_int(v_row['Baki Balance']):,}</span></div>", unsafe_allow_html=True)
    
    strlt.markdown("<br>", unsafe_allow_html=True)

    # --- FORM INPUT ---
    strlt.subheader("✏️ Entry Form (Amount daalkar Enter dabayein)")
    
    voice_trans_text = ""
    if strlt.button("🎙️ Speak Transaction Entry (Bolein: 'Bees hazar maal' ya 'Das hazar jama')"):
        with strlt.spinner("Aap boliye, system sun raha hai..."):
            voice_trans_text = listen_voice_command()
            if voice_trans_text: strlt.info(f"System ne suna: '{voice_trans_text}'")
            
    v_amt_val = None
    v_type_val = "Maal"
    if voice_trans_text:
        words = voice_trans_text.lower().split()
        for w in words:
            if w.isdigit(): v_amt_val = safe_int(w)
        if "jama" in voice_trans_text or "paisa" in voice_trans_text: v_type_val = "Jama"

    with strlt.form(key="ledger_entry_form", clear_on_submit=False):
        # 💎 SHRUNK DATE BLOCK COLUMN: Reduced date column size ratio from 4 to 2.5 to make DD-MM-YYYY tight & small
        col_date_block, col_amt, col_tp = strlt.columns([2.5, 4, 2.5])
        
        today = datetime.date.today()
        if strlt.session_state.editing_id:
            edit_row = df_trans[df_trans["ID"] == strlt.session_state.editing_id].iloc[0]
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

        btn_label = "💾 Update Badlav Save Karein" if strlt.session_state.editing_id else "⚡ Save Record (Ya Keyboard se Enter dabayein)"
        submit_ledger = strlt.form_submit_button(label=btn_label, type="primary")

    if submit_ledger:
        if e_amt is None or safe_int(e_amt) == 0:
            strlt.error("Kripya valid Amount enter karein!")
        else:
            # ✅ Fix 2 Kept: Try-Except string parsing strategy fallback execution[cite: 1]
            try:
                parsed_date = datetime.datetime.strptime(
                    final_constructed_date_str,
                    "%d-%m-%Y"
                )
            except:
                parsed_date = pd.NaT
                
            if pd.isna(parsed_date):
                strlt.error("Chuni hui tarikh calendar ke hisab se valid nahi h!")
            elif parsed_date.date() > datetime.date.today():
                strlt.error("Aage ki future tarikh block h!")
            else:
                if strlt.session_state.editing_id:
                    idx = df_trans[df_trans["ID"] == strlt.session_state.editing_id].index[0]
                    df_trans.at[idx, "Tarikh"] = parsed_date
                    df_trans.at[idx, "Type"] = e_type
                    df_trans.at[idx, "Amount"] = safe_int(e_amt)
                    strlt.session_state.editing_id = None
                else:
                    new_id = df_trans["ID"].max() + 1 if not df_trans.empty else 1
                    new_tr = pd.DataFrame([{"ID": new_id, "Tarikh": parsed_date, "Naam": v_name, "Type": e_type, "Amount": safe_int(e_amt), "Note": ""}])
                    df_trans = pd.concat([df_trans, new_tr], ignore_index=True)
                    
                save_data(df_dash, df_trans)
                recalculate_all()
                strlt.rerun()

    strlt.markdown("---")

    # --- ↩️ LOCAL LEDGER UNDO BAR ---
    if has_deleted_trans and not has_deleted_vyapari:
        strlt.markdown("<div class='undo-banner'>⚠️ Entry Mita Diya Gaya Hai!</div>", unsafe_allow_html=True)
        ub_c1, ub_c2 = strlt.columns([5, 5])
        if ub_c1.button("↩️ UNDO (Wapas Layein)", type="primary", use_container_width=True):
            df_trans = pd.concat([df_trans, strlt.session_state.last_deleted_trans], ignore_index=True)
            save_data(df_dash, df_trans)
            recalculate_all()
            strlt.session_state.last_deleted_trans = None
            strlt.success("Entry safaltapurvak wapas aa gayi!")
            strlt.rerun()
        if ub_c2.button("❌ Close (Hatayein)", use_container_width=True):
            strlt.session_state.last_deleted_trans = None
            strlt.rerun()

    col_l, col_r = strlt.columns(2)
    
    # --- LEFT COLUMN: MAAL HISTORY ---
    with col_l:
        strlt.markdown("### 📦 MAAL LIYA HISTORY")
        for _, row in df_m.iterrows():
            # ✅ Fix 5 Kept: Safe check format condition rule[cite: 1]
            formatted_date = (
                row["Tarikh"].strftime("%d-%m-%Y")
                if pd.notnull(row["Tarikh"])
                else ""
            )
            amt_val = safe_int(row['Amount'])
            
            is_this_target = (strlt.session_state.delete_confirm_target is not None and 
                              strlt.session_state.delete_confirm_target["scope"] == "transaction" and 
                              strlt.session_state.delete_confirm_target["key"] == row["ID"])
            
            style_class = "lbl-maal lbl-settled" if row['ID'] in settled_trans_ids else "lbl-maal"
            
            c_box1, c_box2, c_box3 = strlt.columns([3, 1, 1])
            c_box1.markdown(f"<div style='margin-bottom:8px;'><span class='lbl-highlight {style_class}'>₹ {amt_val:,}</span><span class='lbl-date'>📅 {formatted_date}</span></div>", unsafe_allow_html=True)
            
            if c_box2.button("✏️", key=f"ed_{row['ID']}"):
                strlt.session_state.editing_id = row["ID"]
                strlt.rerun()
                
            if c_box3.button("❌", key=f"del_{row['ID']}"):
                strlt.session_state.delete_confirm_target = {"scope": "transaction", "key": row["ID"], "display_info": f"Maal Entry worth ₹{amt_val:,}"}
                strlt.rerun()
                
            if is_this_target:
                strlt.markdown(f"<div class='alert-inline-banner'>⚠️ Mitaana chahte hain?</div>", unsafe_allow_html=True)
                btn_grid1, btn_grid2 = strlt.columns(2)
                if btn_grid1.button("✅ HAAN", key=f"yes_t_{row['ID']}", type="primary"):
                    strlt.session_state.last_deleted_trans = df_trans[df_trans["ID"] == row["ID"]]
                    df_trans = df_trans[df_trans["ID"] != row["ID"]]
                    save_data(df_dash, df_trans)
                    recalculate_all()
                    strlt.session_state.delete_confirm_target = None
                    strlt.rerun()
                if btn_grid2.button("❌ ROKO", key=f"no_t_{row['ID']}"):
                    strlt.session_state.delete_confirm_target = None
                    strlt.rerun()

    # --- RIGHT COLUMN: JAMA HISTORY ---
    with col_r:
        strlt.markdown("### 💰 HAR HAFTE JAMA PAISA")
        for _, row in df_j.iterrows():
            # ✅ Fix 5 Kept: Same conditional logic validation[cite: 1]
            formatted_date = (
                row["Tarikh"].strftime("%d-%m-%Y")
                if pd.notnull(row["Tarikh"])
                else ""
            )
            amt_val = safe_int(row['Amount'])
            note_val = str(row['Note']) if pd.notnull(row['Note']) and str(row['Note']) != "nan" else ""
            
            is_this_target = (strlt.session_state.delete_confirm_target is not None and 
                              strlt.session_state.delete_confirm_target["scope"] == "transaction" and 
                              strlt.session_state.delete_confirm_target["key"] == row["ID"])
            
            if row['ID'] in settled_trans_ids:
                jama_class = "lbl-settled"
            else:
                if amt_val <= 5000: jama_class = "lbl-j-1"
                elif amt_val <= 10000: jama_class = "lbl-j-2"
                elif amt_val <= 15000: jama_class = "lbl-j-3"
                elif amt_val <= 20000: jama_class = "lbl-j-4"
                else: jama_class = "lbl-j-5"
                
            c_box1, c_box2, c_box3, c_box4 = strlt.columns([3, 1, 1, 2])
            c_box1.markdown(f"<div style='margin-bottom:8px;'><span class='lbl-highlight {jama_class}'>₹ {amt_val:,}</span><span class='lbl-date'>📅 {formatted_date}</span></div>", unsafe_allow_html=True)
            
            if c_box2.button("✏️", key=f"ed_j_{row['ID']}"):
                strlt.session_state.editing_id = row["ID"]
                strlt.rerun()
                
            if c_box3.button("❌", key=f"del_j_{row['ID']}"):
                strlt.session_state.delete_confirm_target = {"scope": "transaction", "key": row["ID"], "display_info": f"Jama Entry worth ₹{amt_val:,}"}
                strlt.rerun()
                
            updated_note = c_box4.text_input(
                "📝 Note", 
                value=note_val, 
                key=f"nt_input_{row['ID']}", 
                label_visibility="collapsed",
                placeholder="......"
            )
            if updated_note != note_val:
                idx = df_trans[df_trans["ID"] == row["ID"]].index[0]
                df_trans.at[idx, "Note"] = str(updated_note).strip()
                save_data(df_dash, df_trans)
                
            if is_this_target:
                strlt.markdown(f"<div class='alert-inline-banner'>⚠️ Mitaana chahte hain?</div>", unsafe_allow_html=True)
                btn_grid1, btn_grid2 = strlt.columns(2)
                if btn_grid1.button("✅ HAAN", key=f"yes_j_{row['ID']}", type="primary"):
                    strlt.session_state.last_deleted_trans = df_trans[df_trans["ID"] == row["ID"]]
                    df_trans = df_trans[df_trans["ID"] != row["ID"]]
                    save_data(df_dash, df_trans)
                    recalculate_all()
                    strlt.session_state.delete_confirm_target = None
                    strlt.rerun()
                if btn_grid2.button("❌ ROKO", key=f"no_j_{row['ID']}"):
                    strlt.session_state.delete_confirm_target = None
                    strlt.rerun()

# ==========================================
# SCREEN 1: MAIN DASHBOARD HOME SCREEN
# ==========================================
else:
    strlt.markdown("<div class='main-title'>✨ Smart Digital Vyapari Ledger Pro+</div>", unsafe_allow_html=True)
    
    t_business = df_dash["Total Maal"].sum()
    t_received = df_dash["Total Jama"].sum()
    t_outstanding = df_dash["Baki Balance"].sum()
    
    card1, card2, card3 = strlt.columns(3)
    card1.markdown(f"<div class='kpi-box kpi-business'><span style='font-size:14px;'>📦 TOTAL BUSINESS (KUL MAAL)</span><br><span style='font-size:24px;'>₹ {safe_int(t_business):,}</span></div>", unsafe_allow_html=True)
    card2.markdown(f"<div class='kpi-box kpi-received'><span style='font-size:14px;'>💰 TOTAL RECEIVED (KUL JAMA)</span><br><span style='font-size:24px;'>₹ {safe_int(t_received):,}</span></div>", unsafe_allow_html=True)
    card3.markdown(f"<div class='kpi-box kpi-outstanding'><span style='font-size:14px;'>🔴 OUTSTANDING (KUL BAKI)</span><br><span style='font-size:24px;'>₹ {safe_int(t_outstanding):,}</span></div>", unsafe_allow_html=True)
    
    strlt.markdown("<br>", unsafe_allow_html=True)
    
    box_title = "👤 ➕ Naya Vyapari Jodein"
    init_name = ""
    init_mob = ""
    
    with strlt.expander(box_title, expanded=False):
        voice_v_text = ""
        if strlt.button("🎙️ Speak Vyapari Info (Bolein: 'Ramesh 91xxxxxxxxxx')"):
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
                if n_name in df_dash["Naam"].values:
                    strlt.warning("Vyapari pehle se add hai.")
                else:
                    new_v = pd.DataFrame([{"Naam": n_name, "Mobile": str(n_mob), "Total Maal": 0, "Total Jama": 0, "Baki Balance": 0}])
                    df_dash = pd.concat([df_dash, new_v], ignore_index=True)
                    save_data(df_dash, df_trans)
                    recalculate_all()
                    strlt.rerun()

    strlt.markdown("---")
    strlt.subheader("👥 Active Vyapari Ledger Logs")

    if has_deleted_vyapari:
        strlt.markdown("<div class='undo-banner'>⚠️ Vyapari Profile Mita Diya Gaya Hai!</div>", unsafe_allow_html=True)
        h_ub_c1, h_ub_c2 = strlt.columns([5, 5])
        if h_ub_c1.button("↩️ UNDO (Wapas Layein)", type="primary", use_container_width=True):
            df_dash = pd.concat([df_dash, pd.DataFrame([strlt.session_state.last_deleted_vyapari])], ignore_index=True)
            if has_deleted_trans:
                df_trans = pd.concat([df_trans, strlt.session_state.last_deleted_trans], ignore_index=True)
            save_data(df_dash, df_trans)
            recalculate_all()
            strlt.session_state.last_deleted_vyapari = None
            strlt.session_state.last_deleted_trans = None
            strlt.success("Profile safaltapurvak wapas aa gaya!")
            strlt.rerun()
        if h_ub_c2.button("❌ Close (Hatayein)", use_container_width=True):
            strlt.session_state.last_deleted_vyapari = None
            strlt.session_state.last_deleted_trans = None
            strlt.rerun()
    
    for idx, row in df_dash.iterrows():
        v_name = row["Naam"]
        serial_no = idx + 1
        
        df_v_trans = df_trans[df_trans["Naam"] == v_name]
        df_v_maal = df_v_trans[df_v_trans["Type"] == "Maal"]
        df_v_jama = df_v_trans[df_v_trans["Type"] == "Jama"]
        
        alert_active = False
        if not df_v_maal.empty and safe_int(row["Baki Balance"]) > 0:
            last_maal_date = df_v_maal["Tarikh"].max()
            if not df_v_jama.empty:
                last_jama_date = df_v_jama["Tarikh"].max()
                if (datetime.datetime.now() - last_jama_date).days > 90 and last_maal_date > last_jama_date:
                    alert_active = True
            else:
                if (datetime.datetime.now() - last_maal_date).days > 90:
                    alert_active = True
                    
        g1, g2, g3, g4 = strlt.columns([5, 2, 1, 1])
        alert_html = "<span class='alert-tag'>🔴 3-MONTH OVERDUE ALERT</span>" if alert_active else ""
        
        g1.markdown(f"👤 **{serial_no}. {v_name}** &nbsp;&nbsp; {alert_html}<br><span style='color:gray; font-size:13px;'>📱 Mobile: {row['Mobile']}</span>", unsafe_allow_html=True)
        g2.markdown(f"<div style='text-align: right;'><b>Baki: ₹{safe_int(row['Baki Balance']):,}</b><br><span style='font-size:12px; color:gray;'>M: ₹{safe_int(row['Total Maal']):,} / J: ₹{safe_int(row['Total Jama']):,}</span></div>", unsafe_allow_html=True)
        
        if g3.button("📖 Open", key=f"open_{v_name}"):
            strlt.session_state.selected_vyapari = v_name
            strlt.rerun()
            
        col_icon1, col_icon2 = g4.columns(2)
        if col_icon1.button("✏️", key=f"edit_v_{v_name}"):
            strlt.session_state.editing_vyapari_name = v_name
            strlt.rerun()
            
        if col_icon2.button("❌", key=f"del_v_{v_name}"):
            strlt.session_state.delete_confirm_target = {"scope": "vyapari", "key": v_name, "display_info": f"Vyapari Profile ({v_name})"}
            strlt.rerun()
            
        if (strlt.session_state.delete_confirm_target is not None and 
            strlt.session_state.delete_confirm_target["scope"] == "vyapari" and 
            strlt.session_state.delete_confirm_target["key"] == v_name):
            strlt.markdown(f"<div class='alert-inline-banner'>⚠️ Vyapari Profile {v_name} ko mitaana chahte hain?</div>", unsafe_allow_html=True)
            v_btn1, v_btn2 = strlt.columns(2)
            if v_btn1.button("✅ CONFIG HAAN", key=f"y_v_{v_name}", type="primary"):
                strlt.session_state.last_deleted_vyapari = df_dash[df_dash["Naam"] == v_name].iloc[0].to_dict()
                strlt.session_state.last_deleted_trans = df_trans[df_trans["Naam"] == v_name]
                df_dash = df_dash[df_dash["Naam"] != v_name]
                df_trans = df_trans[df_trans["Naam"] != v_name]
                strlt.session_state.delete_confirm_target = None
                save_data(df_dash, df_trans)
                recalculate_all()
                strlt.rerun()
            if v_btn2.button("❌ ROKO", key=f"n_v_{v_name}"):
                strlt.session_state.delete_confirm_target = None
                strlt.rerun()
            
        strlt.markdown("<div style='border-bottom:1px solid #e1e5eb; margin-bottom:12px; margin-top:4px;'></div>", unsafe_allow_html=True)