import streamlit as st
import requests
import datetime
import qrcode
import io
from io import BytesIO  
import json             
import os
import pandas as pd

# =========================================================================
# 🔑 1. CONFIGURATION SYSTEM & LINE NOTIFY (ระบบดั้งเดิมของคุณ)
# =========================================================================
LINE_ACCESS_TOKEN = "RRtpOuJT8oWgvglsSFUqc7LC1zZqL2jD8qTdJx5iIpAkG4GiJjAkaetvEKLGLuNOJ7j9dpyNMSTviG06LCe//YM1+r5TqRQx09p8nLNh5lZzKy78CvGLfGAWjFSOtyj89Bu3nm8iVlTh0pNQtc737gdB04t89/1O/w1cDnyilFU=" 
LINE_TARGET_ID = "Cbf3d27d5280ae8b258727047a26b399a"  
BOSS_PASSWORD = "boss1234"  
DATA_FILE = "local_maintenance_data.csv" # ไฟล์เก็บข้อมูลในเซิร์ฟเวอร์ตัวเอง

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
# 📑 2. MASTER DATA (รายการเครื่องจักร 54 รายการ + เช็คลิสต์คำต่อคำ ห้ามเปลี่ยนเพื่อ ISO)
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

def get_machine_type(m_id):
    if "CNC" in m_id and "CRANE" not in m_id.upper() and "QC-" not in m_id.upper(): return "CNC"
    if "CRANE NO.1" in m_id.upper(): return "Crane no.1"
    if "CRANE NO.2" in m_id.upper(): return "Crane no.2"
    if m_id == "QC-01": return "QC-01"
    if m_id in ["QC-02", "QC-03", "QC-04", "QC-05", "QC-06", "QC-07", "QC-08", "QC-09"]: return "QC-VERNIER_STD"
    if m_id in ["QC-10", "QC-11", "QC-12"]: return "QC-HIGAUGE_STD"
    if m_id in ["QC-13", "QC-14"]: return "QC-MICRO_STD"
    if m_id == "QC-15": return "QC-15"
    if m_id == "QC-16": return "QC-16"
    if m_id == "QC-17": return "QC-17"
    if m_id in ["QC-18", "QC-19", "QC-20", "QC-21"]: return "QC-ARM_STD"
    if "COMP-" in m_id: return "COMP_STD"  
    if "GRINDING" in m_id: return "GRINDING"
    if "CUTTER" in m_id: return "CUTTER GRINDING"
    if "MILLING" in m_id: return "MILLING"
    if "CUTTING" in m_id: return "CUTTING"
    if "MIG" in m_id: return "MIG CO2"
    if "ARGON" in m_id: return "ARGON"
    if "BAND" in m_id: return "BAND SAW"
    return "CNC"

PHOTO_RULES = {
    "CNC": [2, 3, 4, 5, 8, 13], "Crane no.1": [3, 4], "Crane no.2": [3, 4], "QC-01": [4],
    "QC-VERNIER_STD": [2, 4], "QC-HIGAUGE_STD": [2], "QC-MICRO_STD": [2], "QC-15": [6], "QC-16": [3], "QC-17": [2],
    "QC-ARM_STD": [3], "COMP_STD": [1, 2, 3], "GRINDING": [2, 4, 7], "CUTTER GRINDING": [],
    "MILLING": [6, 7], "CUTTING": [3, 5, 7], "MIG CO2": [3, 4, 5], "ARGON": [3, 4, 6], "BAND SAW": [2, 3, 5]
}

# =========================================================================
# 💾 3. LOCAL FILE STORAGE ENGINE (ระบบบันทึกในตัว เสถียรที่สุด 100%)
# =========================================================================
def save_tech_data_locally(machine_id, tech_name, results_dict):
    """บันทึกข้อมูลแบบด่วนลงเป็นไฟล์ตารางในระบบภายในแอปเพื่อความชัวร์"""
    try:
        summary_list = []
        for item, details in results_dict.items():
            status_text = details['status'] if details['status'] else "ไม่ได้เลือกสถานะ"
            summary_list.append(f"{item}: {status_text}" + (f" ({details['note']})" if details['note'] else ""))
        all_form_data = " | ".join(summary_list)
        
        new_data = {
            "Timestamp": [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
            "Machine ID": [machine_id],
            "Machine Name": [MACHINES.get(machine_id, "")],
            "Tech Name": [tech_name],
            "Form Data": [all_form_data]
        }
        df_new = pd.DataFrame(new_data)
        
        if os.path.exists(DATA_FILE):
            df_old = pd.read_csv(DATA_FILE)
            df_total = pd.concat([df_old, df_new], ignore_index=True)
        else:
            df_total = df_new
            
        df_total.to_csv(DATA_FILE, index=False)
        return True, "บันทึกข้อมูลเรียบร้อย"
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
    st.title(f"📋 ใบตรวจสอบเครื่อง {m_type_selected} ประจำวัน")
    st.info("📄 มาตรฐานระบบคุณภาพโรงงาน: **FM-MN-07 Rev.00**")

    if not query_params.get("id") and not query_params.get("machine_id"):
        machine_id = st.selectbox("🎯 เลือกเครื่องจักรที่เข้าทำงาน:", list(MACHINES.keys()), format_func=lambda x: f"[{x}] {MACHINES[x]}")
        m_type_selected = get_machine_type(machine_id)
    else:
        st.success(f"⚙️ คุณกำลังตรวจเครื่อง: **{machine_id} ({MACHINES.get(machine_id, '')})**")
    st.divider()

    with st.form("pm_form"):
        tech_name = st.text_input("👤 ชื่อช่างผู้ตรวจเช็ค (ผู้รับผิดชอบ)", placeholder="ระบุชื่อ-นามสกุลของคุณ")
        results, uploaded_photos = {}, {}
        current_checklist = CHECKLISTS[m_type_selected]
        required_photo_indexes = PHOTO_RULES.get(m_type_selected, [])
        
        for i, item in enumerate(current_checklist, 1):
            st.write(f"**{i}. {item}**")
            status = st.radio(f"ผลการตรวจข้อ {i}", ["ใช้งานได้ปกติ", "ทำการแก้ไขใช้งานได้ปกติ", "ใช้งานไม่ได้ต้องแก้ไข", "ไม่ได้ทำงาน"], horizontal=True, key=f"check_{i}", label_visibility="collapsed", index=None)
            if i in required_photo_indexes:
                uploaded_file = st.file_uploader(f"แนบรูปข้อ {i}", type=["jpg", "jpeg", "png"], key=f"photo_{i}")
                uploaded_photos[i] = {"file": uploaded_file, "index": i}
            note = st.text_input(f"หมายเหตุ/อาการเสีย (ข้อ {i})", key=f"note_{i}", placeholder="ระบุรายละเอียดหากพบจุดพังหรือบันทึกงานซ่อมแก้ไข")
            results[item] = {"status": status, "note": note}
            st.divider()

        submitted = st.form_submit_button("💾 ส่งรายงานการตรวจเช็คประจำวัน (SUBMIT)")

    if submitted:
        if not tech_name: st.error("❌ กรุณากรอกชื่อช่างผู้ตรวจสอบก่อนส่งรายงานครับ!")
        elif any(results[item]["status"] is None for item in current_checklist): st.error("❌ ปฏิเสธการบันทึก! ช่างยังเลือกผลการตรวจสอบไม่ครบทุกหัวข้อ")
        elif any(uploaded_photos[idx]["file"] is None for idx in required_photo_indexes): st.error(f"❌ ปฏิเสธการบันทึกฟอร์ม! กรุณาแนบภาพให้ครบตามเกณฑ์")
        else:
            fails, fixed_items = [], []
            for i, item in enumerate(current_checklist, 1):
                status_val = results[item]["status"]
                note_val = results[item]["note"]
                if status_val == "ใช้งานไม่ได้ต้องแก้ไข": fails.append(f"- ข้อ {i}. {item}" + (f" ({note_val})" if note_val else ""))
                elif status_val == "ทำการแก้ไขใช้งานได้ปกติ": fixed_items.append(f"- ข้อ {i}. {item}" + (f" ({note_val})" if note_val else ""))
            
            with st.spinner("🚀 กำลังบันทึกข้อมูลรายงานเข้าระบบส่วนกลาง..."):
                success, err_msg = save_tech_data_locally(machine_id, tech_name, results)
                if success:
                    audit_tag = f"\n\n🔒 [ISO Status]: บันทึกข้อมูลเข้าระบบตาราง Sheet ในแอปสำเร็จ (รอเจ้าหน้าที่ดาวน์โหลดออกเอกสาร)"
                    if fails:
                        summary_msg = f"\n🚨 [แจ้งซ่อมด่วนจากใบตรวจเช็ค ISO]\n🔧 เครื่อง: {machine_id}\n📅 วันที่: {current_time_str}\n👤 ผู้ตรวจ: {tech_name}\n\n❌ ไม่ผ่านมาตรฐาน:\n" + "\n".join(fails)
                        send_line_alert(summary_msg + audit_tag)
                    else:
                        ok_msg = f"\n🎉 [รายงานเครื่องจักรปกติ - ISO]\n🔧 เครื่อง: {machine_id}\n✅ ปกติทุกหัวข้อ\n👤 ผู้ตรวจสอบ: {tech_name}"
                        send_line_alert(ok_msg + audit_tag)
                
                    st.success(f"🎉 บันทึกรายงานเครื่อง {machine_id} สำเร็จเสร็จสิ้น 100%!")
                    st.balloons()
                else:
                    st.error(f"เกิดข้อผิดพลาดคลาวด์: {err_msg}")

else:
    st.title("🔐 หน้าต่างสำหรับเจ้าหน้าที่ / หัวหน้างาน (ดาวน์โหลดตาราง ISO)")
    password_input = st.text_input("🔑 กรุณากรอกรหัสผ่านผู้เข้าตรวจสอบเพื่อเข้าถึงระบบอนุมัติ:", type="password")
    if password_input == BOSS_PASSWORD:
        st.success("🔓 รหัสผ่านถูกต้อง ยินดีต้อนรับคุณพลวัฒน์")
        st.write("### 📊 บอร์ดจัดเก็บรายงานเอกสารอิเล็กทรอนิกส์")
        
        if os.path.exists(DATA_FILE):
            df_report = pd.read_csv(DATA_FILE)
            st.dataframe(df_report, use_container_width=True)
            
            # แปลงข้อมูลเป็น Excel เพื่อให้ดาวน์โหลดไปพิมพ์ฟอร์ม ISO
            towrite = io.BytesIO()
            df_report.to_excel(towrite, index=False, header=True)
            towrite.seek(0)
            
            st.download_button(
                label="📥 คลิกที่นี่เพื่อดาวน์โหลดไฟล์ตารางรายงานออกมาเป็น Excel",
                data=towrite,
                file_name=f"ISO_Maintenance_Report_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.ms-excel",
                use_container_width=True
            )
        else:
            st.info("💡 ปัจจุบันยังไม่มีข้อมูลช่างเทคนิคส่งฟอร์มเข้ามาในระบบรอบใหม่นี้ครับ")
    elif password_input != "": 
        st.error("❌ รหัสผ่านไม่ถูกต้อง")
