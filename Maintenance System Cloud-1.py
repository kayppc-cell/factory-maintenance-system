import streamlit as st
import requests
import datetime
import qrcode
import io
import os
import gspread
from google.oauth2.service_account import Credentials

# =========================================================================
# 🔑 1. CONFIGURATION SYSTEM & LINE NOTIFY (ดึงจากไฟล์ของคุณเป๊ะๆ)
# =========================================================================
LINE_ACCESS_TOKEN = "RRtpOuJT8oWgvglsSFUqc7LC1zZqL2jD8qTdJx5iIpAkG4GiJjAkaetvEKLGLuNOJ7j9dpyNMSTviG06LCe//YM1+r5TqRQx09p8nLNh5lZzKy78CvGLfGAWjFSOtyj89Bu3nm8iVlTh0pNQtc737gdB04t89/1O/w1cDnyilFU=" 
LINE_TARGET_ID = "Cbf3d27d5280ae8b258727047a26b399a"  
BOSS_PASSWORD = "boss1234"  

# ⚠️ นำรหัส Spreadsheet ID ที่ได้จากก้าวที่ 2 มาวางตรงนี้แทนนะครับเพื่อนรัก
SPREADSHEET_ID = "1hXBpjrZMJDGmBC0ib9tSP-FeCISCgg9QOYG8NwHt6cA" 

def get_google_sheet_client():
    """เปิดประตูเชื่อมสายเน็ตไปยังคลาวด์ Google Sheets"""
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    return gspread.authorize(Credentials.from_service_account_file("google_creds.json", scopes=scopes))

def send_line_notify(message):
    """ฟังก์ชันแจ้งเตือนเข้ากลุ่ม LINE ดั้งเดิมของโรงงานเรา"""
    url = "https://notify-api.line.me/api/notify"
    headers = {"Authorization": f"Bearer {LINE_ACCESS_TOKEN}"}
    data = {"message": message}
    try:
        requests.post(url, headers=headers, data=data, timeout=5)
    except Exception:
        pass

# =========================================================================
# 📑 2. MASTER DATA (รายการเครื่องจักร 54 รายการ ตรงตามบรีฟล่าสุด 100%)
# =========================================================================
MACHINES = {
    # 1. แผนก CNC (3 แกน & 5 แกน)
    "CNC3X-01": "CNC 3 แกน #01", "CNC3X-02": "CNC 3 แกน #02",
    "CNC3X-03": "CNC 3 แกน #03", "CNC3X-04": "CNC 3 แกน #04",
    "CNC3X-05": "CNC 3 แกน #05", "CNC3X-06": "CNC 3 แกน #06",
    "CNC3X-07": "CNC 3 แกน #07", "CNC3X-08": "CNC 3 แกน #08",
    "CNC5X-01": "CNC 5 แกน #พิเศษ",
    # 2. แผนก เครน (Crane)
    "Crane no.1": "เครน CNC NO.1", "Crane no.2": "เครน QC NO.2",
    # 3. แผนก เครื่องมือวัด QC
    "QC-01": "เครื่องวัดความแข็ง QC-01", "QC-02": "เวอร์เนีย 1000 QC-02",
    "QC-03": "เวอร์เนีย 300 QC-03", "QC-04": "เวอร์เนีย 300 QC-04",
    "QC-05": "เวอร์เนีย 200 QC-05", "QC-06": "เวอร์เนีย 200 QC-06",
    "QC-07": "เวอร์เนีย 200 QC-07", "QC-08": "เวอร์เนีย 200 QC-08",
    "QC-09": "เวอร์เนีย 200 QC-09", "QC-10": "ไฮเกจ 300 QC-10",
    "QC-11": "ไฮเกจ 600 QC-11", "QC-12": "ไฮเกจ 600 QC-12",
    "QC-13": "ไมโคร 0-25 QC-13", "QC-14": "ไมโคร 5-30 QC-14",
    "QC-15": "CMM QC-15", "QC-16": "Laser QC-16",
    "QC-17": "เลื่อยสายพาน QC-17", "QC-18": "Faro Arm QC-18",
    "QC-19": "Arm 2.8 QC-19", "QC-20": "Arm 2.4 QC-20",
    "QC-21": "Arm 3.5 QC-21",
    # 4. แผนก ปั๊มลม
    "COMP-01": "ปั๊มลม 1 COMP-01", "COMP-02": "ปั๊มลม 2 COMP-02",
    # 5. แผนก เครื่องเจียร & ลับคม (Grinding / Cutter)
    "GRINDING-01": "เครื่องเจียร GRINDING #01", "GRINDING-02": "เครื่องเจียร GRINDING #02",
    "CUTTER GRINDING-01": "เครื่องลับคม CUTTER GRINDING #01",
    # 6. แผนก มิลลิ่ง (Milling)
    "MILLING-01": "เครื่องมิลลิ่ง #01", "MILLING-02": "เครื่องมิลลิ่ง #02", "MILLING-03": "เครื่องมิลลิ่ง #03",
    # 7. แผนก เครื่องตัด (Cutting)
    "CUTTING-01": "เครื่องตัด CUTTING #01", "CUTTING-02": "เครื่องตัด CUTTING #02",
    # 8. แผนก เครื่องเชื่อม (MIG / ARGON)
    "MIG CO2-01": "เครื่องเชื่อม MIG CO2 #01", "MIG CO2-02": "เครื่องเชื่อม MIG CO2 #02", "MIG CO2-03": "เครื่องเชื่อม MIG CO2 #03",
    "ARGON-01": "เครื่องเชื่อม ARGON #01",
    # 9. แผนก เครื่องเลื่อยสายพาน (Band Saw)
    "BAND SAW-01": "เครื่องเลื่อยสายพาน #01", "BAND SAW-02": "เครื่องเลื่อยสายพาน #02", "BAND SAW-03": "เครื่องเลื่อยสายพาน #03"
}

# รายการเช็คลิสต์ดั้งเดิมตามหน้างานจริง
CHECKLISTS = {
    "CNC_STD": ["เช็คระดับน้ำมันหล่อลื่นทางไหล (Way Lubricant)", "เช็คระดับน้ำยาหล่อเย็น (Coolant)", "เช็คความดันลม (Air Pressure) แหล่งจ่าย ต้องได้ 5-6 bar", "เช็คเสียงผิดปกติของชุดสปินเดิลและแกนเคลื่อนที่", "ทำความสะอาดกรองอากาศตู้คอนโทรล"],
    "CRANE_STD": ["เช็คระบบสวิตช์ควบคุม/ปุ่มฉุกเฉิน (EMG)", "เช็คการเคลื่อนที่แนวราบและแนวดิ่งไม่มีสะดุด", "เช็คสภาพลวดสลิง/โซ่ยกและตะขอเกี่ยวยกของ", "เช็คเสียงผิดปกติของมอเตอร์เกียร์ขับเคลื่อน"],
    "QC-HARDNESS": ["ตรวจดูสภาพตัวเครื่องทั่วไปพร้อมใช้งานหรือไม่", "ตรวจสอบเข็มหรือหน้าจอแสดงผลความแข็ง", "ทดสอบกดแผ่นมาตรฐานเช็คค่าความเที่ยงตรง", "หลังเลิกงานทำความสะอาดชิ้นงานและปิดเครื่องทุกครั้ง"],
    "QC-VERNIER": ["ตรวจดูสภาพของเวอร์เนียพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด", "ตรวจสอบที่คีบตรงปลายที่ใช้วัดชิ้นงาน เช็คว่ามีรอยบิ่น หรือสึกหล่อหรือไม่", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-HIGAUGE": ["ตรวจดูสภาพของไฮเกจพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-MICRO": ["ตรวจดูสภาพของไมโครพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการหมุนเข้าหมุนออกต้องไม่ติดขัด", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-15": ["ตรวจดูสภาพของสายไฟ", "ตรวจดูสภาพของลมพร้อมใช้งานหรือไม่", "ตรวจดูสภาพของ SURFACE BASE", "ตรวจดูสภาพของ STICKER", "ตรวจสอบ COMPUTER", "ตรวจดูสภาพของหัว PROBE คตงอหรือไม่"],
    "QC-16": ["ตรวจดูสภาพของสายไฟ", "ตรวจดูสภาพ ของเครื่อง", "ตรวจดูสภาพของ Lend Laser", "ตรวจสอบ STICKER", "ตรวจสอบ COMPUTER"],
    "QC-17": ["ตรวจดูสภาพสายไฟ", "ตรวจดูสภาพของใบเลื่อย", "ตรวจดูสภาพของมอเตอร์", "ตรวจดูสภาพของสายพาน", "ตรวจสอบเครื่องเลื่อยสายพาน"],
    "QC-ARM": ["ตรวจดูสภาพของสายไฟ", "ตรวจดูสภาพ ARM ของเครื่อง", "ตรวจดูสภาพของหัว PROBE คตงอหรือไม่", "ตรวจสอบ STICKER", "ตรวจสอบ NOTEBOOK COMPUTER"],
    "COMP_STD": ["เช็คแรงดัน (Pressure) ต้องไม่ต่ำกว่า 7 bar", "ตรวจสอบระดับน้ำมันไฮดรอลิก ต้องไม่ต่ำกว่าระดับต่ำสุด", "เช็คอุณหภูมิความร้อนต้องไม่เกิน 80 องศา", "เช็คการรั่วซีมของระบบน้ำมัน", "เช็คระบบเดรนน้ำ (Water Draen)"],
    "GRIND_STD": ["เช็คระดับน้ำมันไฮดรอลิกชุดขับเคลื่อน", "เช็คสภาพหินเจียรและเพลาจับหิน", "เช็คระบบดูดฝุ่น/ละอองน้ำยาหล่อเย็น", "ทำความสะอาดแม่เหล็กจับชิ้นงาน (Magnetic Chuck)"],
    "MILL_STD": ["เช็คระดับน้ำมันหล่อลื่นตามจุดหยอดน้ำมัน", "เช็คความลื่นไหลของแท่นประคอง (Saddle/Slide)", "เช็คสภาพระบบเบรกและลิมิตสวิตช์", "เช็คระบบน้ำยาหล่อเย็นและการรั่วซึม"],
    "WELD_STD": ["เช็คสภาพสายไฟและสายเชื่อมไม่มีรอยฉีกขาด", "เช็คระบบแก๊สปกคลุมและการรั่วซึม (MIG/ARGON)", "เช็คพัดลมระบายความร้อนหลังตู้เชื่อม", "ตรวจสอบหัวทิช/ปลอกตลับทิชทำความสะอาดเศษสะเก็ด"],
    "SAW_STD": ["เช็คความตึงของใบเลื่อยและสภาพฟันใบเลื่อย", "เช็คระดับน้ำยาหล่อเย็นตัดชิ้นงาน", "เช็คระบบป้อนกินชิ้นงานอัตโนมัติ (Feed Rate)", "ทำความสะอาดเศษเหล็กในตู้สายพาน"]
}

def get_checklist(m_id):
    if "CNC" in m_id: return CHECKLISTS["CNC_STD"]
    if "Crane" in m_id: return CHECKLISTS["CRANE_STD"]
    if m_id == "QC-01": return CHECKLISTS["QC-HARDNESS"]
    if m_id in ["QC-02", "QC-03", "QC-04", "QC-05", "QC-06", "QC-07", "QC-08", "QC-09"]: return CHECKLISTS["QC-VERNIER"]
    if m_id in ["QC-10", "QC-11", "QC-12"]: return CHECKLISTS["QC-HIGAUGE"]
    if m_id in ["QC-13", "QC-14"]: return CHECKLISTS["QC-MICRO"]
    if m_id == "QC-15": return CHECKLISTS["QC-15"]
    if m_id == "QC-16": return CHECKLISTS["QC-16"]
    if m_id == "QC-17": return CHECKLISTS["QC-17"]
    if m_id in ["QC-18", "QC-19", "QC-20", "QC-21"]: return CHECKLISTS["QC-ARM"]
    if "COMP-" in m_id: return CHECKLISTS["COMP_STD"]
    if "GRIND" in m_id or "CUTTER" in m_id: return CHECKLISTS["GRIND_STD"]
    if "MILL" in m_id: return CHECKLISTS["MILL_STD"]
    if "CUTTING" in m_id or "SAW" in m_id: return CHECKLISTS["SAW_STD"]
    if "MIG" in m_id or "ARGON" in m_id: return CHECKLISTS["WELD_STD"]
    return []

# =========================================================================
# 📐 3. COORDINATES MAPPER (ล็อกพิกัดตามที่คุณกำหนดไว้เป๊ะๆ)
# =========================================================================
def get_coordinates(m_id):
    # 1. กลุ่มเครื่องมือวัด QC และปั๊มลม (ช่าง 11,12 / หัวหน้า 13,14 / อาการเสียสะสม B16)
    if m_id in ["QC-01", "QC-02", "QC-03", "QC-04", "QC-05", "QC-06", "QC-07", "QC-08", "QC-09", "QC-10", "QC-11", "QC-12", "QC-13", "QC-14", "QC-17", "QC-18", "QC-19", "QC-20", "QC-21"] or "COMP-" in m_id:
        return 11, 13, 16
    # 2. ข้อยกเว้นเครื่อง CMM QC-15 (ช่าง 12,13 / หัวหน้า 14,15 / อาการเสียสะสม B17)
    if m_id == "QC-15":
        return 12, 14, 17
    # 3. กลุ่มแผนกโรงงานอื่นๆ ทั้งหมด (ช่าง 11,12 / หัวหน้า 13,14 / อาการเสียสะสม B15)
    return 11, 13, 15

def get_photo_rules(m_id):
    if "CNC" in m_id: return [1, 3]
    if m_id in ["QC-02", "QC-03", "QC-04", "QC-05", "QC-06", "QC-07", "QC-08", "QC-09"]: return [2, 4]
    if m_id in ["QC-10", "QC-11", "QC-12"]: return [2]
    if m_id in ["QC-13", "QC-14"]: return [2]
    if m_id == "QC-15": return [6]
    if m_id == "QC-16": return [3]
    if m_id == "QC-17": return [2]
    if m_id in ["QC-18", "QC-19", "QC-20", "QC-21"]: return [3]
    if "COMP" in m_id: return [1, 3]
    return [1]

# =========================================================================
# ☁️ 4. GOOGLE SHEETS CLOUD ENGINE (ระบบจัดการบันทึกข้ามอินเทอร์เน็ต)
# =========================================================================
def get_or_setup_worksheet(sh, m_id):
    try:
        return sh.worksheet(m_id)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=m_id, rows="50", cols="40")
        ws.update_cell(1, 1, f"ใบตรวจสอบประจำเครื่องจักร: {MACHINES.get(m_id, m_id)}")
        ws.update_cell(4, 2, "รายการตรวจสอบ / วันที่ประจำเดือน")
        for d in range(1, 32): ws.update_cell(4, d + 2, d)
        
        items = get_checklist(m_id)
        for idx, text in enumerate(items):
            ws.update_cell(idx + 5, 1, idx + 1)
            ws.update_cell(idx + 5, 2, text)
            
        t_row, b_row, n_row = get_coordinates(m_id)
        ws.update_cell(t_row, 2, "ช่างเทคนิคผู้เข้าตรวจ (ลงชื่อ)")
        ws.update_cell(t_row + 1, 2, "ผลลัพธ์ภาพรวมประจำวัน (OK/NG)")
        ws.update_cell(b_row, 2, "หัวหน้างานผู้อนุมัติ (ลงชื่อ)")
        ws.update_cell(b_row + 1, 2, "เวลาที่อนุมัติออนไลน์")
        ws.update_cell(n_row, 1, "บันทึกอาการเสียสะสม (ช่อง B)")
        return ws

def save_tech_data_to_cloud(m_id, tech_name, check_results, note):
    try:
        client = get_google_sheet_client()
        sh = client.open_by_key(SPREADSHEET_ID)
        ws = get_or_setup_worksheet(sh, m_id)
        day = datetime.datetime.now().day
        target_col = day + 2
        t_row, _, n_row = get_coordinates(m_id)
        
        for idx, res in enumerate(check_results):
            ws.update_cell(idx + 5, target_col, res)
        
        ws.update_cell(t_row, target_col, tech_name)
        ws.update_cell(t_row + 1, target_col, "OK")
        
        if note:
            ws.update_cell(n_row, 2, note)
            
        send_line_notify(f"\n🛠️ ช่าง [{tech_name}] ได้ส่งผลตรวจเครื่อง [{MACHINES[m_id]}] เรียบร้อยแล้ววันนี้ครับ!")
        return True, "บันทึกผลตรวจสอบเข้าสู่ระบบ Cloud Google Sheets เรียบร้อย!"
    except Exception as e:
        return False, f"ระบบคลาวด์ขัดข้อง: {str(e)}"

def save_boss_approval_to_cloud(m_id, boss_name):
    try:
        client = get_google_sheet_client()
        sh = client.open_by_key(SPREADSHEET_ID)
        ws = get_or_setup_worksheet(sh, m_id)
        day = datetime.datetime.now().day
        target_col = day + 2
        _, b_row, _ = get_coordinates(m_id)
        
        ws.update_cell(b_row, target_col, boss_name)
        ws.update_cell(b_row + 1, target_col, datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))
        
        send_line_notify(f"\n✅ หัวหน้า [{boss_name}] ได้ทำการกดอนุมัติใบงานเครื่อง [{MACHINES[m_id]}] ผ่านคลาวด์แล้วครับ!")
        return True, "อนุมัติใบงานสำเร็จ ข้อมูลเซ็นชื่อลง Google Sheets เรียบร้อยครับหัวหน้า!"
    except Exception as e:
        return False, f"อนุมัติล้มเหลวเนื่องจาก: {str(e)}"

def read_status_from_cloud(m_id):
    try:
        client = get_google_sheet_client()
        sh = client.open_by_key(SPREADSHEET_ID)
        ws = get_or_setup_worksheet(sh, m_id)
        day = datetime.datetime.now().day
        target_col = day + 2
        t_row, b_row, n_row = get_coordinates(m_id)
        
        t_val = ws.cell(t_row, target_col).value
        b_val = ws.cell(b_row, target_col).value
        n_val = ws.cell(n_row, 2).value
        return t_val, b_val, n_val
    except Exception:
        return None, None, None

# =========================================================================
# 🎨 5. STREAMLIT WEB APP UI (ระบบแยก 9 แผนก ชัดเจนตามจริงของโรงงานเรา)
# =========================================================================
st.set_page_config(page_title="ระบบซ่อมบำรุงโรงงานดิจิทัล", layout="wide")
st.title("🏭 ระบบบันทึกผลตรวจสอบเครื่องจักรและเครื่องมือวัด (Cloud ISO-9001)")

query_params = st.query_params
url_id = query_params.get("machine_id", None)
selected_machine = url_id.upper().strip() if url_id and url_id.upper().strip() in MACHINES else None

tab_form, tab_boss, tab_qr_gen = st.tabs(["📝 หน้าฟอร์มตรวจเช็ค (ช่าง)", "📊 บอร์ดควบคุมรวมแผนก (หัวหน้า)", "🖨️ พิมพ์ป้าย QR Code ติดเครื่อง"])

# -------------------------------------------------------------------------
# แท็บที่ 1: หน้าฟอร์มตรวจเช็ค (ช่าง)
# -------------------------------------------------------------------------
with tab_form:
    st.subheader("📋 บันทึกรายงานตรวจเช็คสภาพเครื่องจักรรายวัน")
    if selected_machine:
        st.success(f"🟢 สแกนสิทธิ์เข้าถึงสำเร็จ: **{MACHINES[selected_machine]} ({selected_machine})**")
        active_m = selected_machine
    else:
        active_m = st.selectbox("🎯 เลือกเครื่องจักรที่เข้าทำงาน:", list(MACHINES.keys()), format_func=lambda x: f"[{x}] {MACHINES[x]}")
        
    st.write("---")
    tech_name = st.text_input("👤 ชื่อ-นามสกุล ช่างผู้ตรวจสอบหน้างาน:")
    
    t_user, b_user, c_note = read_status_from_cloud(active_m)
    st.info(f"🔍 ประวัติวันนี้: ช่างตรวจแล้ว [{t_user or 'ยังไม่มี'}] | หัวหน้าอนุมัติแล้ว [{b_user or 'รอดำเนินการ'}]")
    if c_note: st.error(f"⚠️ อาการเสียสะสมปัจจุบันในช่อง B: {c_note}")
        
    items = get_checklist(active_m)
    p_rules = get_photo_rules(active_m)
    
    st.write("### ⚙️ รายการตรวจสอบ:")
    results = []
    photo_ok = True
    
    for i, text in enumerate(items):
        step = i + 1
        c_text, c_radio, c_file = st.columns([5, 2, 3])
        with c_text: st.write(f"**{step}. {text}**")
        with c_radio:
            res = st.radio(f"ข้อ {step}", ["✓ OK", "✗ NG"], key=f"r_{active_m}_{step}", label_visibility="collapsed")
            results.append("OK" if "✓" in res else "NG")
        with c_file:
            if step in p_rules:
                up_f = st.file_uploader(f"📷 รูปหลักฐานข้อ {step}", type=["jpg", "png", "jpeg"], key=f"f_{active_m}_{step}", label_visibility="collapsed")
                if not up_f: photo_ok = False
            else:
                st.caption("ไม่ต้องแนบภาพ")
                
    note_text = st.text_area("✍️ บันทึกอาการเสียหรือข้อมูลเพิ่มเติม (ลงในคอลัมน์ B):")
    
    if st.button("🚀 ยิงผลตรวจบันทึกขึ้นระบบคลาวด์ตารางถาวร", use_container_width=True):
        if not tech_name: st.error("❌ โปรดป้อนชื่อช่างเทคนิคผู้ตรวจงานก่อนครับ")
        elif not photo_ok: st.error("❌ ปฏิเสธการส่ง! กรุณาถ่ายรูปภาพหลักฐานประจำข้อบังคับให้ครบถ้วนก่อน")
        else:
            with st.spinner("กำลังเชื่อมต่อไปยัง Google Cloud..."):
                ok, msg = save_tech_data_to_cloud(active_m, tech_name, results, note_text)
                if ok: st.success(msg); st.balloons()
                else: st.error(msg)

# -------------------------------------------------------------------------
# แท็บที่ 2: หน้าบอร์ดอนุมัติงาน (หัวหน้า)
# -------------------------------------------------------------------------
with tab_boss:
    st.subheader("📊 แดชบอร์ดคุมสถานะการอนุมัติใบงานซ่อมบำรุง (สิทธิ์หัวหน้า)")
    boss_name = st.text_input("🔑 ชื่อ-นามสกุล หัวหน้างาน/ผู้จัดการ:")
    boss_pass = st.text_input("🔒 รหัสผ่านอนุมัติ (Password):", type="password")
    st.write("---")
    
    def draw_card(m_id, m_name):
        t_name, b_name, note = read_status_from_cloud(m_id)
        with st.container(border=True):
            st.write(f"##### **{m_id}**")
            st.caption(m_name)
            if t_name:
                st.markdown("🟢 **ช่างลงนามแล้ว**")
                st.text(f"โดย: {t_name}")
                if b_name:
                    st.markdown(f"✅ **อนุมัติแล้ว**\n({b_name})")
                else:
                    st.markdown("🟡 **รออนุมัติ**")
                    if st.button("📥 กดอนุมัติใบงาน", key=f"b_{m_id}"):
                        if boss_pass != BOSS_PASSWORD: st.error("รหัสผ่านไม่ถูกต้อง")
                        elif not boss_name: st.error("โปรดใส่ชื่อผู้จัดการ")
                        else:
                            s, m = save_boss_approval_to_cloud(m_id, boss_name)
                            if s: st.success(m); st.rerun()
                            else: st.error(m)
            else:
                st.markdown("🔴 **วันนี้ยังไม่มีคนตรวจ**")
            if note: st.caption(f"⚠️ เสียสะสม: {note}")

    # ดึงการจัดกรุ๊ปแผนกตามไฟล์จริงล่าสุดของคุณ
    categories = {
        "⚙️ 1. แผนก CNC (3 แกน / 5 แกน)": ["CNC3X-01", "CNC3X-02", "CNC3X-03", "CNC3X-04", "CNC3X-05", "CNC3X-06", "CNC3X-07", "CNC3X-08", "CNC5X-01"],
        "🏗️ 2. แผนก เครน (Crane)": ["Crane no.1", "Crane no.2"],
        "🔎 3. แผนก เครื่องมือวัด QC": ["QC-01", "QC-02", "QC-03", "QC-04", "QC-05", "QC-06", "QC-07", "QC-08", "QC-09", "QC-10", "QC-11", "QC-12", "QC-13", "QC-14", "QC-15", "QC-16", "QC-17", "QC-18", "QC-19", "QC-20", "QC-21"],
        "💨 4. แผนก ปั๊มลม (Compressor)": ["COMP-01", "COMP-02"],
        "💎 5. แผนก เครื่องเจียร & ลับคม": ["GRINDING-01", "GRINDING-02", "CUTTER GRINDING-01"],
        "📐 6. แผนก เครื่องมิลลิ่ง (Milling)": ["MILLING-01", "MILLING-02", "MILLING-03"],
        "✂️ 7. แผนก เครื่องตัด (Cutting)": ["CUTTING-01", "CUTTING-02"],
        "🔥 8. แผนก เครื่องเชื่อม (MIG / ARGON)": ["MIG CO2-01", "MIG CO2-02", "MIG CO2-03", "ARGON-01"],
        "🪚 9. แผนก เครื่องเลื่อยสายพาน (Band Saw)": ["BAND SAW-01", "BAND SAW-02", "BAND SAW-03"]
    }
    
    for cat_title, machine_list in categories.items():
        st.write(f"#### {cat_title}")
        cols = st.columns(3)
        for idx, m_id in enumerate(machine_list):
            with cols[idx % 3]:
                draw_card(m_id, MACHINES[m_id])

# -------------------------------------------------------------------------
# แท็บที่ 3: ระบบพิมพ์ QR Code
# -------------------------------------------------------------------------
with tab_qr_gen:
    st.subheader("🖨️ ออกใบฉลาก QR Code ประจำเครื่องเพื่อพิมพ์แปะตัวเครื่องจักร")
    qr_target = st.selectbox("เลือกเครื่องจักรที่ต้องการออกบาร์โค้ด QR Code:", list(MACHINES.keys()), format_func=lambda x: f"[{x}] {MACHINES[x]}")
    
    base_url = "https://factory-maintenance.streamlit.app/" 
    full_url = f"{base_url}?machine_id={qr_target}"
    
    qr = qrcode.QRCode(version=1, box_size=10, border=4)
    qr.add_data(full_url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white")
    
    buf = io.BytesIO()
    qr_img.save(buf, format="PNG")
    byte_im = buf.getvalue()
    
    st.image(byte_im, width=200)
    st.info(f"🔗 ลิงก์ภายในป้ายนี้: `{full_url}`")
    st.download_button("💾 ดาวน์โหลดไฟล์รูปภาพ QR Code (.png)", data=byte_im, file_name=f"QR_{qr_target}.png", mime="image/png")