import streamlit as st
import requests
import datetime
import qrcode
import io
from io import BytesIO  
import json             
import os
from openpyxl import load_workbook

# =========================================================================
# 🔑 1. CONFIGURATION SYSTEM & LINE NOTIFY (ดึงจากไฟล์เดิมของคุณ)
# =========================================================================
LINE_ACCESS_TOKEN = "RRtpOuJT8oWgvglsSFUqc7LC1zZqL2jD8qTdJx5iIpAkG4GiJjAkaetvEKLGLuNOJ7j9dpyNMSTviG06LCe//YM1+r5TqRQx09p8nLNh5lZzKy78CvGLfGAWjFSOtyj89Bu3nm8iVlTh0pNQtc737gdB04t89/1O/w1cDnyilFU=" 
LINE_TARGET_ID = "Cbf3d27d5280ae8b258727047a26b399a"  
BOSS_PASSWORD = "boss1234"  

def send_line_alert(msg_text):
    """ฟังก์ชันส่งสัญญาณแจ้งเตือนเข้า LINE กลุ่มแบบ Push Message ดั้งเดิม"""
    url = 'https://api.line.me/v2/bot/message/push'
    headers = {'Content-Type': 'application/json', 'Authorization': f'Bearer {LINE_ACCESS_TOKEN}'}
    payload = {"to": LINE_TARGET_ID, "messages": [{"type": "text", "text": msg_text}]}
    try: 
        requests.post(url, headers=headers, data=json.dumps(payload), timeout=5)
    except Exception: 
        pass

# =========================================================================
# 📑 2. MASTER DATA (รายการเครื่องจักร 54 รายการ + เช็คลิสต์คำต่อคำ เพื่อ ISO)
# =========================================================================
MACHINES = {
    "CNC3X-01": "CNC 3 แกน #01", "CNC3X-02": "CNC 3 แกน #02",
    "CNC3X-03": "CNC 3 แกน #03", "CNC3X-04": "CNC 3 แกน #04",
    "CNC3X-05": "CNC 3 แกน #05", "CNC3X-06": "CNC 3 แกน #06",
    "CNC3X-07": "CNC 3 แกน #07", "CNC3X-08": "CNC 3 แกน #08",
    "CNC5X-01": "CNC 5 แกน #พิเศษ",
    "Crane no.1": "เครน CNC NO.1", "Crane no.2": "เครน QC NO.2",
    "QC-01": "เครื่องวัดความแข็ง QC-01",  
    "QC-02": "เวอร์เนีย 1000 QC-02",       
    "QC-03": "เวอร์เนีย 300 QC-03",  
    "QC-04": "เวอร์เนีย 300 QC-04",   
    "QC-05": "เวอร์เนีย 200 QC-05",
    "QC-06": "เวอร์เนีย 200 QC-06",  
    "QC-07": "เวอร์เนีย 200 QC-07",
    "QC-08": "เวอร์เนีย 200 QC-08", 
    "QC-09": "เวอร์เนีย 200 QC-09",
    "QC-10": "ไฮเกจ 300 QC-10",
    "QC-11": "ไฮเกจ 600 QC-11",
    "QC-12": "ไฮเกจ 600 QC-12",
    "QC-13": "ไมโคร 0-25 QC-13",
    "QC-14": "ไมโคร 5-30 QC-14",
    "QC-15": "CMM QC-15",
    "QC-16": "Laser QC-16",
    "QC-17": "เลื่อยสายพาน QC-17",
    "QC-18": "Faro Arm QC-18",
    "QC-19": "Arm 2.8 QC-19",
    "QC-20": "Arm 2.4 QC-20",               
    "QC-21": "Arm 3.5 QC-21",
    "COMP-01": "ปั๊มลม 1 COMP-01",          
    "COMP-02": "ปั๊มลม 2 COMP-02",
    "GRINDING-01": "เครื่องเจียร GRINDING #01", "GRINDING-02": "เครื่องเจียร GRINDING #02",
    "CUTTER GRINDING-01": "เครื่องลับคม CUTTER GRINDING #01",
    "MILLING-01": "เครื่องมิลลิ่ง #01", "MILLING-02": "เครื่องมิลลิ่ง #02", "MILLING-03": "เครื่องมิลลิ่ง #03",
    "CUTTING-01": "เครื่องตัด CUTTING #01", "CUTTING-02": "เครื่องตัด CUTTING #02",
    "MIG CO2-01": "เครื่องเชื่อม MIG CO2 #01", "MIG CO2-02": "เครื่องเชื่อม MIG CO2 #02", "MIG CO2-03": "เครื่องเชื่อม MIG CO2 #03",
    "ARGON-01": "เครื่องเชื่อม ARGON #01",
    "BAND SAW-01": "เครื่องเลื่อยสายพาน #01", "BAND SAW-02": "เครื่องเลื่อยสายพาน #02", "BAND SAW-03": "เครื่องเลื่อยสายพาน #03"
}

CHECKLISTS = {
    "CNC": [
        "Worm up เครื่องจักร 15 นาที ทุกครั้งที่ใช้งาน", "เช็คระดับนำมัน Oil Matic Mesh ทุกวัน เติมเมื่อพร่อง",
        "ทำความสะอาด Air filter Mesh ทุกวัน", "ตรวจเช็คแรงดัน Air Control Unit ปกติเฉลี่ยที่ 0.5 Mpa",
        "เช็คน้ำมันชุด Gear ของ ATC ทุกวัน(เปลี่ยนถ่ายทุกปี)", "อัดจารบี Ballscrew และ Linear Guideทุก 1000 ชม.",
        "ตรวจสอบการเคลื่อนที่ของแกนทุกแกน (X,Y,Z)", "ตรวจสอบสภาพของน้ำ Coolant ถ่ายรูปค่าที่วัดได้ส่งเข้าระบบ",
        "การทำงานของ Coolant pump", "การทำงานของ Unclamp และการเปลี่ยน Tool", "การทำงานของ Spindle",
        "การทำงานของ Arm เปลี่ยน Tool", "ตรวจสอบระดับของน้ำ Coolant เติมเมื่ออยู่ระดับที่ต่ำ",
        "ความสะอาดทั่วไปของเครื่องจักรโดยรวม", "ตรวจสอบความพร้อมสภาพโดยรวม(ฟังด้วยหู ดูด้วยตา)", "ตรวจสอบสายไฮโดรลิกส์ และสายลม"
    ],
    "Crane no.1": [
        "ตรวจเช็คปุ่มกดต้องอยู่ในสภาพพร้อมใช้งาน ไม่แตก ไม่ชำรุด เสียหาย", "ตรวจเช็คการหยุดเครน เดินหน้า และถอยหลัง เมื่อปล่อยปุ่มกดต้องหยุดทันที",
        "ตรวจเช็คสลิงต้องไม่แตกฝอยเป็นหนาม\nและบิดงอ", "ตรวจเช็คตะขอต้องไม่มีรอยแตกร้าวสูญหายกิ๊ปปากตะขอไม่ชำรุดหรือหลุดหาย",
        "ตรวจเช็คสายบังคับเครนต้องไม่ชำรุดสายไฟไม่ขาดรุ่งริ่ง\nไม่เรียบร้อย", "ตรวจเช็คสัญญานเสียงเมื่อเริ่มเดินเครนต้องมีเสียงเตือนการทำงาน"
    ],
    "Crane no.2": [
        "ตรวจเช็คปุ่มกดต้องอยู่ในสภาพพร้อมใช้งาน ไม่แตก ไม่ชำรุด เสียหาย", "ตรวจเช็คการหยุดเครน เดินหน้า และถอยหลัง เมื่อปล่อยปุ่มกดต้องหยุดทันที",
        "ตรวจเช็คสลิงต้องไม่แตกฝอยเป็นหนาม\nและบิดงอ", "ตรวจเช็คตะขอต้องไม่มีรอยแตกร้าวสูญหายกิ๊ปปากตะขอไม่ชำรุดหรือหลุดหาย",
        "ตรวจเช็คสายบังคับเครนต้องไม่ชำรุดสายไฟไม่ขาดรุ่งริ่ง\nไม่เรียบร้อย", "ตรวจเช็คสัญญานเสียงเมื่อเริ่มเดินเครนต้องมีเสียงเตือนการทำงาน"
    ],
    "QC-01": ["ตรวจสอบความสะอาด", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบปุ่มกดต่างๆๆ", "ตรวจสอบอุปกรณ์การชาร์จ"],
    "QC-VERNIER_STD": [
        "ตรวจดูสภาพของเวอร์เนียพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด",
        "ตรวจสอบที่คีบตรงปลายที่ใช้วัดชิ้นงาน เช็คว่ามีรอยบิ่น หรือสึกหล่อหรือไม่", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"
    ],
    "QC-HIGAUGE_STD": ["ตรวจดูสภาพของไมโครพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-MICRO_STD": ["ตรวจดูสภาพของไมโครพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการหมุนเข้าหมุนออกต้องไม่ติดขัด", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-15": ["ตรวจดูสภาพของสายไฟ", "ตรวจดูสภาพของลมพร้อมใช้งานหรือไม่", "ตรวจดูสภาพของ SURFACE BASE", "ตรวจดูสภาพของ STICKER", "ตรวจสอบ COMPUTER", "ตรวจดูสภาพของหัว PROBE คตงอหรือไม่"],
    "QC-16": ["ตรวจดูสภาพของสายไฟ", "ตรวจดูสภาพ ของเครื่อง", "ตรวจดูสภาพของ Lend Laser", "ตรวจสอบ STICKER", "ตรวจสอบ COMPUTER"],
    "QC-17": ["ตรวจดูสภาพสายไฟ", "ตรวจดูสภาพของใบเลื่อย", "ตรวจดูสภาพของมอเตอร์", "ตรวจดูสภาพของสายพาน", "ตรวจสอบเครื่องเลื่อยสายพาน"],
    "QC-ARM_STD": ["ตรวจดูสภาพของสายไฟ", "ตรวจดูสภาพ ARM ของเครื่อง", "ตรวจดูสภาพของหัว PROBE คตงอหรือไม่", "ตรวจสอบ STICKER", "ตรวจสอบ NOTEBOOK COMPUTER"],
    "COMP_STD": [
        "เช็คแรงดัน (Pressure) ต้องไม่ต่ำกว่า 7 bar", "ตรวจสอบระดับน้ำมันไฮดรอลิก ต้องไม่ต่ำกว่าระดับต่ำสุด",
        "เช็คอุณหภูมิความร้อนต้องไม่เกิน 80 องศา", "เช็คการรั่วซีมของระบบน้ำมัน", "เช็คระบบเดรนน้ำ (Water Draen)"
    ],
    "GRINDING": [
        "การ Worm spindle และ TABLE SLIDE", "เช็คระดับนำมันไฮดรอลิก และ การทำงานของ PUMP", "เช็คระดับของน้ำยา COOLANNT PUMP",
        "ตรวจสอบการทำงานของแม่เหล็ก", "ตรวจสอบการทำงานของ SLIDE X,Y", "ตรวจสอบสภาพความพร้อมโดยรวมของเครื่องจักร",
        "ตรวจสอบระดับน้ำมันของ PUMPน้ำมันหล่อลื่น", "ตรวจสอบการทำงานของไฟฟ้าและแสงสว่าง", "ตรวจสอบการทำงานของตัวดูดอากศ"
    ],
    "CUTTER GRINDING": ["การ WORM UP แกน Y พร้อมใช้งาน", "การ WORM UP แกน Z พร้อมใช้งาน", "ตรวจสอบการทำงานของไฟฟ้าและแสงสว่าง", "ตรวจสอบการทำงานของมอเตอร์ มีการหมุนปกติ", "ตรวจสอบการจับหัวคอเรต"],
    "MILLING": [
        "Worm Spindle ก่อนเริมงาน ตรวจสอบความ ผิดปกติของชุด  Back gauge  และ Motor", "เช็ค Auto  Up-Down back gauge  และ Manual ( ความคร่องตัวในการเคลื่อนที่ของ Spindle )",
        "ตรวจสอบการ SLIDE  ของแกน X", "ตรวจสอบการ SLIDE  ของแกน Y", "ตรวจสอบการ SLIDE  ของแกน Z",
        "ระดับน้ำมันไฮดรอลิค ตรวจสอบน้ำมันในปั้มน้ำมันหล่อ ลื่นแกน  X,Y,Z", "ตรวจน้ำมันหล่อลื่นเย็น ตรวจสอบการทำงานของปั้ม COOLANT และสภาพของน้ำ  COOLANT",
        "ตรวจสอบหน้าจอ  DIGITAL READ OUT และการทำ งานของ LINEAR SCALE", "หยอดน้ำมันหล่อลื่นทุกวันจันทร์",
        "ตรวจสอบการทำงานของไฟฟ้าแสงสว่างของเครื่อง", "ตรวจสอบสภาพความพร้อมโดยรวมของเครื่องจักร  และอุกรณ์เสริมต่าง ๆ"
    ],
    "CUTTING": [
        "การ Worm spindle ก่อนเริ่มงาน เพื่อตรวจ ความผิดปกติของชุด Back gauge และ Motor", "เช็ค Auto Up-Down back gauge และ Manual ( ความคล่องตัวในการเคลื่อนที่ )",
        "ระดับน้ำมันไฮดรอลิค ตรวจสอบระดับในปั้มน้ำมัน หล่อลืนแกน  Back gauge", "ตรวจเช็ค  Switch  เปิด-ปิด",
        "ตรวจสอบ Digital  read out และการทำงานของ Linear  scale", "อัดจาระบีตามจุดที่อัดจาระบีทุกๆจุด",
        "ตรวจสอบใบมีด  บนและล่าง", "ตรวจสอบความพร้อมสภาพโดยรวมของเครื่อง จักรและอุปกรณ์เสริมต่าง ๆ"
    ],
    "MIG CO2": [
        "ตรวจสภาพความพร้อมโดยรวมของเครื่อง", "เช็ค BREAKER เพื่อเช็คระบบไฟฟ้า ตามตำแหน่งไฟ โชว์ และสวิชท์ต่าง ๆ",
        "ตรวจสภาพความพร้อมของมาตราวัดแรงดัน ของก๊าซ CO2 และปรับตั้งอย่างถูกต้อง", "ตรวจจุดต่อของก๊าซ CO2 รั่วหรือไม่",
        "ตรวจสภาพความพร้อมของสายไฟ สายก๊าซ  CO2 ว่ารั่วหรือไม่", "ตรวจสภาพความพร้อมของสายกราวด์", "ทำความสะอาดหัวเชื่อมก่อนใช้งาน"
    ],
    "ARGON": [
        "ตรวจสภาพความพรัอมโดยรวมของเครื่อง", "เช็ค  BREAKER  เพื่อเช็คระบบไฟฟ้า ตามตำแหน่งไฟ โชว์  และ SWITCH  ต่าง ๆ",
        "ตรวจสภาพความพร้อมของมาตราวัดแรงดันของมาตรา วัดแรงดันของก๊าช  ARGON  และปรับตั้งอย่างถูกวิธี", "ตรวจุดต่อของสายก๊าช  ARGON  ก่อนว่ารั่วหรือไม่",
        "ตรวจสภาพความพร้อมของสายกราว์", "ตรวจสภาความพร้อมของสายไฟฟ้าสายก๊าช  ARGON และชุดหัวเชื่อม", "ตรวจสภาพความพร้อมของ  SWITCH  หัวเชื่อม", "ทำความสะดาดชุดหัวเชื่อมก่อนใช้งาน"
    ],
    "BAND SAW": [
        "เช็ค Auto Up-Down Back Gauge และ Manual (ความคล่องตัวในการเคลื่อนที่ของ Spindle)", "เช็คระดับน้ำมันไฮดรอลิค",
        "ตรวจน้ำมันหล่อลื่นเย็น ตรวจสอบการทำงานของปั๊ม COOLANT และสภาพของน้ำ COOLANT", "ตรวจสอบ Switch (สวิตซ์) หน้า BOX CONTROL", "ตรวจสอบระดับน้ำมันหล่อลื่นในห้องเกียร์"
    ]
}

def get_coordinates(m_type):
    if m_type == "CNC": return 22, 24, "B28"
    if "CRANE" in m_type.upper(): return 14, 16, "B19"
    if m_type == "QC-01": return 10, 12, "B15"
    if m_type in ["QC-VERNIER_STD", "QC-MICRO_STD", "QC-ARM_STD", "QC-16", "QC-17", "BAND SAW", "CUTTING", "ARGON"]: return 11, 13, "B16"
    if m_type == "QC-HIGAUGE_STD": return 11, 13, "B15"
    if m_type == "QC-15": return 12, 14, "B17"
    if m_type == "GRINDING": return 16, 18, "B21"
    if m_type == "CUTTER GRINDING": return 13, 15, "B18"
    if m_type == "MILLING": return 20, 22, "B25"
    if m_type == "MIG CO2": return 13, 15, "B18"
    return 11, 13, "B16"

# =========================================================================
# ⚙️ 3. MULTI-FILE EXCEL STORAGE ENGINE (ระบบเขียนแยก 1 เครื่องต่อ 1 ไฟล์เด็ดขาด)
# =========================================================================
def save_tech_data_to_separate_excel(machine_id, tech_name, results_dict, m_type):
    """ฟังก์ชันเปิดไฟล์แม่แบบรายเครื่องจักร หยอดเครื่องหมาย และบันทึกแยกไฟล์ไม่ปนกัน"""
    try:
        template_file = f"{machine_id}.xlsx"       # เช่น CNC3X-01.xlsx
        live_output_file = f"LIVE_{machine_id}.xlsx" # ไฟล์ที่อัปเดตข้อมูลแล้ว
        
        # ถ้ายังไม่มีไฟล์สะสมข้อมูล ให้จำลองคัดลอกมาจากต้นฉบับรายเครื่องจักร
        if not os.path.exists(live_output_file):
            if os.path.exists(template_file):
                import shutil
                shutil.copy(template_file, live_output_file)
            else:
                return False, f"ไม่พบไฟล์ต้นฉบับ {template_file} บน GitHub กรุณาอัปโหลดไฟล์เทมเพลตของเครื่องนี้ก่อนนะครับเพื่อนรัก"
        
        # เปิดไฟล์รายเครื่องขึ้นมาเขียนทับ
        wb = load_workbook(live_output_file)
        ws = wb.active # เขียนลงหน้าหลักของไฟล์เครื่องนั้นๆ
        
        day_num = datetime.datetime.now().day
        target_col = day_num + 2 # วันที่ 1 ตรงกับคอลัมน์ C (3)
        
        items = CHECKLISTS[m_type]
        for idx, item in enumerate(items, 5):
            status_val = results_dict[item]["status"]
            short_status = "✓" if "ปกติ" in status_val else "❌" if "ต้องแก้ไข" in status_val else "-"
            ws.cell(row=idx, column=target_col, value=short_status)
            
        # ลงชื่อช่างประจำวันในช่องวันที่
        t_row, _, n_cell = get_coordinates(m_type)
        ws.cell(row=t_row, column=target_col, value=tech_name)
        
        # บันทึกอาการเสียสะสมลงช่อง B ด้านล่าง
        notes_collected = [results_dict[item]["note"] for item in items if results_dict[item]["note"]]
        if notes_collected:
            note_row = int(n_cell[1:])
            old_val = ws.cell(row=note_row, column=2).value or ""
            new_val = old_val + ("\n" if old_val else "") + f"[วันที่ {day_num}]: " + ", ".join(notes_collected)
            ws.cell(row=note_row, column=2, value=new_val)
            
        wb.save(live_output_file)
        return True, "บันทึกเข้าฟอร์มเครื่องเฉพาะสำเร็จ"
    except Exception as e:
        return False, str(e)

# =========================================================================
# 🎨 4. STREAMLIT WEB APP UI 
# =========================================================================
st.sidebar.title("🏢 เมนูควบคุมโรงงานรวม")
user_role = st.sidebar.radio("เลือกสิทธิ์การเข้าใช้งานด้านล่าง:", ["🔧 ช่างเทคนิค (ส่งฟอร์ม)", "🔐 หัวหน้างาน/ผู้ตรวจสอบ"])

now = datetime.datetime.now()
current_time_str = now.strftime("%Y-%m-%d %H:%M:%S")

query_params = st.query_params
raw_machine_id = query_params.get("id") or query_params.get("machine_id") or query_params.get("machine")
machine_id = "CNC3X-01" if not raw_machine_id else str(raw_machine_id if not isinstance(raw_machine_id, list) else raw_machine_id[0]).strip().replace("%20", " ")

for actual_id in MACHINES.keys():
    if machine_id.upper() == actual_id.upper():
        machine_id = actual_id
        break

m_type_selected = get_machine_type(machine_id)

if user_role == "🔧 ช่างเทคนิค (ส่งฟอร์ม)":
    st.caption("PHOLLAWAT ENGINEERING SUPPLY CO., LTD.")
    st.title(f"📋 ใบตรวจสอบเครื่อง {machine_id} ประจำวัน")
    st.info(f"📄 มาตรฐาน ISO: **FM-MN-07 Rev.00** บันทึกแยกไฟล์ตรงระบบเครื่อง `[ {machine_id} ]`")

    if not query_params.get("id") and not query_params.get("machine_id"):
        machine_id = st.selectbox("🎯 เลือกเครื่องจักรที่เข้าทำงาน:", list(MACHINES.keys()), format_func=lambda x: f"[{x}] {MACHINES[x]}")
        m_type_selected = get_machine_type(machine_id)
    else:
        st.success(f"⚙️ คุณกำลังตรวจเครื่อง: **{machine_id} ({MACHINES.get(machine_id, '')})**")
    st.divider()

    with st.form("pm_form"):
        tech_name = st.text_input("👤 ชื่อช่างผู้ตรวจเช็ค (ผู้รับผิดชอบ)", placeholder="ระบุชื่อ-นามสกุลของคุณ")
        results = {}
        current_checklist = CHECKLISTS[m_type_selected]
        
        for i, item in enumerate(current_checklist, 1):
            st.write(f"**{i}. {item}**")
            status = st.radio(f"ผลการตรวจข้อ {i}", ["ใช้งานได้ปกติ", "ทำการแก้ไขใช้งานได้ปกติ", "ใช้งานไม่ได้ต้องแก้ไข", "ไม่ได้ทำงาน"], horizontal=True, key=f"check_{i}", label_visibility="collapsed", index=None)
            note = st.text_input(f"หมายเหตุ/อาการเสีย (ข้อ {i})", key=f"note_{i}", placeholder="ระบุรายละเอียดหากพบจุดพังหรือบันทึกงานซ่อมแก้ไข")
            results[item] = {"status": status, "note": note}
            st.divider()

        submitted = st.form_submit_button("💾 ส่งรายงานการตรวจเช็คประจำวัน (SUBMIT)")

    if submitted:
        if not tech_name: st.error("❌ กรุณากรอกชื่อช่างผู้ตรวจสอบก่อนส่งรายงานครับ!")
        elif any(results[item]["status"] is None for item in current_checklist): st.error("❌ ปฏิเสธการบันทึก! ช่างยังเลือกผลการตรวจสอบไม่ครบทุกหัวข้อ")
        else:
            fails = [f"- ข้อ {i}. {item}" for i, item in enumerate(current_checklist, 1) if results[item]["status"] == "ใช้งานไม่ได้ต้องแก้ไข"]
            
            with st.spinner(f"🚀 กำลังเขียนข้อมูลลงในฟอร์ม Excel ไฟล์ {machine_id}.xlsx..."):
                success, err_msg = save_tech_data_to_separate_excel(machine_id, tech_name, results, m_type_selected)
                if success:
                    audit_tag = f"\n\n🔒 [ISO Status]: แยกไฟล์บันทึกเรียบร้อยเป็นไฟล์ {machine_id}.xlsx ชัดเจน"
                    if fails:
                        summary_msg = f"\n🚨 [แจ้งซ่อมด่วนจากใบตรวจเช็ค ISO]\n🔧 เครื่อง: {machine_id}\n📅 วันที่: {current_time_str}\n👤 ผู้ตรวจ: {tech_name}\n\n❌ ไม่ผ่านมาตรฐาน:\n" + "\n".join(fails)
                        send_line_alert(summary_msg + audit_tag)
                    else:
                        ok_msg = f"\n🎉 [รายงานเครื่องจักรปกติ - ISO]\n🔧 เครื่อง: {machine_id}\n✅ ปกติทุกหัวข้อ\n👤 ผู้ตรวจสอบ: {tech_name}"
                        send_line_alert(ok_msg + audit_tag)
                
                    st.success(f"🎉 บันทึกข้อมูลลงใบฟอร์ม Excel ประจำเครื่อง {machine_id} แยกไฟล์เฉพาะเรียบร้อย!")
                    st.balloons()
                else:
                    st.error(f"เกิดข้อผิดพลาด: {err_msg}")

else:
    st.title("🔐 หน้าต่างดาวน์โหลดฟอร์มแยกรายเครื่องจักร (สำหรับหัวหน้างาน)")
    password_input = st.text_input("🔑 กรุณากรอกรหัสผ่านผู้เข้าตรวจสอบเพื่อเข้าถึงระบบอนุมัติ:", type="password")
    if password_input == BOSS_PASSWORD:
        st.success("🔓 รหัสผ่านถูกต้อง ยินดีต้อนรับคุณพลวัฒน์")
        st.write("### 🖨️ คลังเลือกดาวน์โหลดไฟล์เอกสาร ISO แยกเครื่อง")
        
        # ให้หัวหน้างานเลือกชื่อเครื่องจักรที่ต้องการจะดึงไฟล์ Excel สรุปผลออกไปพิมพ์งาน
        selected_download_machine = st.selectbox("🎯 เลือกเครื่องจักรที่ต้องการดาวน์โหลดไฟล์ Excel:", list(MACHINES.keys()), format_func=lambda x: f"[{x}] {MACHINES[x]}")
        
        live_file_target = f"LIVE_{selected_download_machine}.xlsx"
        
        if os.path.exists(live_file_target):
            st.info(f"💡 ไฟล์ของเครื่อง {selected_download_machine} พร้อมให้ดาวน์โหลดแล้วครับ")
            with open(live_file_target, "rb") as f:
                st.download_button(
                    label=f"📥 คลิกดาวน์โหลดฟอร์มแท้ของเครื่อง {selected_download_machine}",
                    data=f,
                    file_name=f"FM-MN-07-Rev00_{selected_download_machine}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.warning(f"💡 ปัจจุบันเครื่อง {selected_download_machine} ยังไม่มีช่างหน้างานกดส่งรายงานเข้ามารอบใหม่ครับ (หรือคุณยังไม่ได้อัปโหลดไฟล์แม่แบบ {selected_download_machine}.xlsx ขึ้น GitHub งับ)")
            
    elif password_input != "": 
        st.error("❌ รหัสผ่านไม่ถูกต้อง")
