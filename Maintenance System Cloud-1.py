import streamlit as st
import requests
import datetime
import qrcode
import io
from io import BytesIO  
import json             
import os
import gspread
from google.oauth2.service_account import Credentials

# =========================================================================
# 🔑 1. CONFIGURATION SYSTEM & LINE NOTIFY (ดึงจากไฟล์เดิมของคุณ)
# =========================================================================
LINE_ACCESS_TOKEN = "RRtpOuJT8oWgvglsSFUqc7LC1zZqL2jD8qTdJx5iIpAkG4GiJjAkaetvEKLGLuNOJ7j9dpyNMSTviG06LCe//YM1+r5TqRQx09p8nLNh5lZzKy78CvGLfGAWjFSOtyj89Bu3nm8iVlTh0pNQtc737gdB04t89/1O/w1cDnyilFU=" 
LINE_TARGET_ID = "Cbf3d27d5280ae8b258727047a26b399a"  
BOSS_PASSWORD = "boss1234"  

# ⚠️ รหัส Spreadsheet ID ตารางกูเกิลชีตโรงงานของคุณ
SPREADSHEET_ID = "1hXBpjrZMJDGmBC0ib9tSP-FeCISCgg9QOYG8NwHt6cA" 

def get_google_sheet_client():
    """เปิดประตูเชื่อมสายเน็ตไปยังคลาวด์ Google Sheets พร้อมระบบชดเชยเวลาเซิร์ฟเวอร์เดินไม่ตรง"""
    import json
    import time
    from google.auth import jwt
    
    scopes = ["https://www.googleapis.com/auth/sheets", "https://www.googleapis.com/auth/drive"]
    
    # 1. โหลดข้อมูลกุญแจดิจิทัลจากไฟล์มาตรฐานที่เราอัปโหลดไว้บน GitHub
    with open("google_creds.json", "r") as f:
        info = json.load(f)
        
    # 2. 🟢 [TIME SKEW FIX] สั่งขยับเวลาปัจจุบันถอยหลังไป 60 วินาที 
    # เพื่อแก้ปัญหาเซิร์ฟเวอร์ Streamlit เวลาเดินเร็วหรือเดินไม่ตรงกับ Google API
    custom_timestamp = int(time.time()) - 60
    
    # 3. สร้างระบบตรวจสอบสิทธิ์โดยใช้โครงสร้างสร้างสิทธิ์แบบดัดแปลงเวลาแสตมป์
    signer = jwt.Signer.from_string(info["private_key"])
    payload = {
        "iss": info["client_email"],
        "sub": info["client_email"],
        "aud": info["token_uri"],
        "exp": custom_timestamp + 3600,
        "iat": custom_timestamp,
        "scope": " ".join(scopes)
    }
    
    # ทำการแปลงโครงสร้างและส่งสิทธิ์ที่ตั้งเวลาเรียบร้อยแล้วกลับไปใช้งาน
    from google.auth.transport.requests import Request
    from google.oauth2.service_account import Credentials as SACredentials
    
    creds = SACredentials.from_service_account_file("google_creds.json", scopes=scopes)
    # บังคับอัปเดตสิทธิ์ล่วงหน้าเพื่อเคลียร์บั๊กเวลา
    try:
        creds.refresh(Request())
    except Exception:
        pass
        
    return gspread.authorize(creds)

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
# ☁️ 4. GOOGLE SHEETS CLOUD ENGINE
# =========================================================================
def get_or_setup_worksheet(sh, m_id, m_type):
    try:
        return sh.worksheet(m_id)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=m_id, rows="50", cols="40")
        ws.update_cell(1, 1, f"ใบตรวจสอบประจำเครื่องจักร: {MACHINES[m_id]}")
        ws.update_cell(4, 2, "รายการตรวจสอบ / วันที่ประจำเดือน")
        for d in range(1, 32): ws.update_cell(4, d + 2, d)
        
        items = CHECKLISTS[m_type]
        for idx, text in enumerate(items):
            ws.update_cell(idx + 5, 1, idx + 1)
            ws.update_cell(idx + 5, 2, text)
            
        t_row, b_row, n_cell = get_coordinates(m_type)
        ws.update_cell(t_row, 2, "ช่างเทคนิคผู้เข้าตรวจ (ลงชื่อ)")
        ws.update_cell(b_row, 2, "หัวหน้างานผู้อนุมัติ (ลงชื่อ)")
        ws.update_cell(int(n_cell[1:]), 1, "บันทึกอาการเสียสะสม (ช่อง B)")
        return ws

def save_tech_data_to_cloud(machine_id, tech_name, results_dict, m_type):
    try:
        client = get_google_sheet_client()
        sh = client.open_by_key(SPREADSHEET_ID)
        ws = get_or_setup_worksheet(sh, machine_id, m_type)
        
        day_num = datetime.datetime.now().day
        target_col = day_num + 2
        
        checklist_items = CHECKLISTS[m_type]
        fail_notes = []
        
        for i, item in enumerate(checklist_items, 1):
            row_pos = 5 + i
            if item in results_dict:
                status_val = results_dict[item]["status"]
                note_val = results_dict[item]["note"]
                
                if status_val == "ใช้งานได้ปกติ": symbol = "/"
                elif status_val == "ทำการแก้ไขใช้งานได้ปกติ":
                    symbol = "⨂"
                    fail_notes.append(f"ข้อ {i} (แก้ไขแล้ว): {note_val}" if note_val else f"ข้อ {i}: แก้ไขจนใช้ได้ปกติ")
                elif status_val == "ใช้งานไม่ได้ต้องแก้ไข":
                    symbol = "X"
                    fail_notes.append(f"ข้อ {i} (พบปัญหา): {note_val}" if note_val else f"ข้อ {i}: พบปัญหาไม่ผ่านเกณฑ์")
                elif status_val == "ไม่ได้ทำงาน": symbol = "-"
                else: symbol = ""
                
                ws.update_cell(row_pos, target_col, symbol)
                
        tech_row, _, n_cell = get_coordinates(m_type)
        note_row_idx = int(n_cell[1:])
        
        old_value = ws.cell(note_row_idx, 2).value
        if old_value is None or old_value == "เครื่องจักรปกติ": old_value = ""
        current_date_str = datetime.datetime.now().strftime("%d/%m/%Y")
        
        if fail_notes:
            new_entry = f"📌 [วันที่ {current_date_str}]\n" + "\n".join(fail_notes)
            ws.update_cell(note_row_idx, 2, f"{old_value}\n\n{new_entry}".strip())
        else:
            if not old_value: ws.update_cell(note_row_idx, 2, "เครื่องจักรปกติ")
            
        ws.update_cell(tech_row, target_col, tech_name)
        return True, "บันทึกข้อมูลเรียบร้อย"
    except Exception as e:
        return False, str(e)

def save_boss_approval_to_cloud(machine_id, boss_name, m_type):
    try:
        client = get_google_sheet_client()
        sh = client.open_by_key(SPREADSHEET_ID)
        ws = sh.worksheet(machine_id)
        day_num = datetime.datetime.now().day
        target_col = day_num + 2
        _, boss_row, _ = get_coordinates(m_type)
        
        ws.update_cell(boss_row, target_col, boss_name)
        return True
    except Exception as e:
        return False

def get_current_cloud_note(machine_id, m_type):
    try:
        client = get_google_sheet_client()
        sh = client.open_by_key(SPREADSHEET_ID)
        ws = sh.worksheet(machine_id)
        _, _, n_cell = get_coordinates(m_type)
        val = ws.cell(int(n_cell[1:]), 2).value
        return val if val else ""
    except: 
        return ""

# =========================================================================
# 🎨 5. STREAMLIT WEB APP UI
# =========================================================================
st.sidebar.title("🏢 เมนูควบคุมโรงงานรวม")
user_role = st.sidebar.radio("เลือกสิทธิ์การเข้าใช้งานด้านล่าง:", ["🔧 ช่างเทคนิค (ส่งฟอร์ม)", "🔐 หัวหน้างาน/ผู้ตรวจสอบ"])

now = datetime.datetime.now()
current_day = now.day
current_time_str = now.strftime("%Y-%m-%d %H:%M:%S")

query_params = st.query_params
raw_machine_id = query_params.get("id") or query_params.get("machine_id") or query_params.get("machine")

if not raw_machine_id:
    machine_id = "CNC3X-01" 
else:
    if isinstance(raw_machine_id, list): 
        machine_id = str(raw_machine_id[0]).strip()
    else: 
        machine_id = str(raw_machine_id).strip()

machine_id = machine_id.replace("%20", " ")

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
        if machine_id in MACHINES: 
            st.success(f"⚙️ คุณกำลังตรวจเครื่อง: **{machine_id} ({MACHINES[machine_id]})**")
        else: 
            st.error(f"⚠️ ไม่พบรหัสเครื่อง '{machine_id}' ในทะเบียนกลาง")
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
                st.write("📷 *หัวข้อบังคับถ่ายรูปหลักฐานยืนยันหน้างานจริง*")
                uploaded_file = st.file_uploader(f"แนบรูปข้อ {i}", type=["jpg", "jpeg", "png"], key=f"photo_{i}")
                uploaded_photos[i] = {"file": uploaded_file, "index": i}
            note = st.text_input(f"หมายเหตุ/อาการเสีย (ข้อ {i})", key=f"note_{i}", placeholder="ระบุรายละเอียดหากพบจุดพังหรือบันทึกงานซ่อมแก้ไข")
            results[item] = {"status": status, "note": note}
            st.divider()

        submitted = st.form_submit_button("💾 ส่งรายงานการตรวจเช็คประจำวัน (SUBMIT)")

    if submitted:
        if machine_id not in MACHINES: st.error("❌ รหัสเครื่องจักรไม่ถูกต้อง")
        elif not tech_name: st.error("❌ กรุณากรอกชื่อช่างผู้ตรวจสอบก่อนส่งรายงานครับ!")
        elif any(results[item]["status"] is None for item in current_checklist): st.error("❌ ปฏิเสธการบันทึก! ช่างยังเลือกผลการตรวจสอบไม่ครบทุกหัวข้อ")
        elif any(uploaded_photos[idx]["file"] is None for idx in required_photo_indexes): st.error(f"❌ ปฏิเสธการบันทึกฟอร์ม! กรุณาถ่ายภาพหลักฐานประจำข้อ {required_photo_indexes} ให้ครบถ้วนก่อนกดส่งครับ")
        else:
            fails, fixed_items = [], []
            for i, item in enumerate(current_checklist, 1):
                status_val = results[item]["status"]
                note_val = results[item]["note"]
                if status_val == "ใช้งานไม่ได้ต้องแก้ไข": fails.append(f"- ข้อ {i}. {item}" + (f" ({note_val})" if note_val else ""))
                elif status_val == "ทำการแก้ไขใช้งานได้ปกติ": fixed_items.append(f"- ข้อ {i}. {item}" + (f" ({note_val})" if note_val else ""))
            
            with st.spinner("กำลังเชื่อมต่อยิงสัญญานขึ้น Cloud Google Sheets..."):
                success, err_msg = save_tech_data_to_cloud(machine_id, tech_name, results, m_type_selected)
                if success:
                    audit_tag = f"\n\n🔒 [ISO Status]: บันทึกรายงานเครื่อง {machine_id} แล้ว (รอหัวหน้าลงนามดิจิทัล)"
                    if fails:
                        summary_msg = f"\n🚨 [แจ้งซ่อมด่วนจากใบตรวจเช็ค ISO]\n🔧 เครื่อง: {MACHINES[machine_id]}\n📅 วันที่: {current_time_str}\n👤 ผู้ตรวจ: {tech_name}\n\n❌ รายการที่ไม่ผ่านมาตรฐาน:\n" + "\n".join(fails)
                        if fixed_items: summary_msg += "\n\n🛠️ รายการที่ช่างแก้ไขเสร็จทันที:\n" + "\n".join(fixed_items)
                        send_line_alert(summary_msg + audit_tag)
                        st.warning("พบจุดบกพร่อง! ส่งการแจ้งเตือนสัญญาณเข้าไลน์กลุ่มเรียบร้อย")
                    else:
                        ok_msg = f"\n🎉 [รายงานเครื่องจักรปกติ - ISO]\n🔧 เครื่อง: {MACHINES[machine_id]}\n📅 วันที่: {current_time_str}\n✅ ผลการตรวจสอบ: ปกติทุกหัวข้อ\n👤 ผู้ตรวจสอบ: {tech_name}"
                        if fixed_items: ok_msg += "\n\n🛠️ รายการที่ช่างแก้ไขหน้างานสำเร็จ:\n" + "\n".join(fixed_items)
                        send_line_alert(ok_msg + audit_tag)
                    st.success(f"🎉 บันทึกรายงานเครื่อง {machine_id} เข้าสู่ Google Sheets สำเร็จเรียบร้อยแล้ว!")
                    st.balloons()
                else:
                    st.error(f"เกิดข้อผิดพลาดคลาวด์: {err_msg}")

else:
    st.title("🔐 หน้าต่างควบคุมระบบตรวจสอบคุณภาพ (สำหรับหัวหน้างาน)")
    st.subheader(f"📅 ประจำวันที่: {now.strftime('%d/%m/%Y')} (ช่องวันที่คอลัมน์บน Google Sheets: วันที่ {current_day})")
    
    password_input = st.text_input("🔑 กรุณากรอกรหัสผ่านผู้เข้าตรวจสอบเพื่อเข้าถึงระบบอนุมัติ:", type="password")
    if password_input == BOSS_PASSWORD:
        st.success("🔓 รหัสผ่านถูกต้อง เข้าสู่ระบบลงนามดิจิทัลมาตรฐาน ISO สำเร็จ")
        boss_name = st.text_input("👤 ชื่อผู้ตรวจสอบ/หัวหน้างาน:", value="พลวัฒน์")
        st.divider()
        st.write("### 📊 บอร์ดควบคุมควบคุมใบงานรวม (แยกรายแผนกตามจริง)")
        
        def render_machine_card(m_id, m_name, m_type_flag):
            st.info(f"⚙️ **{m_id}**\n{m_name}")
            if st.button(f"✅ กดอนุมัติฟอร์มออนไลน์ของ {m_id}", key=f"btn_{m_id}"):
                if save_boss_approval_to_cloud(m_id, boss_name, m_type_flag):
                    st.toast(f"ลงนามดิจิทัลเครื่อง {m_id} สำเร็จ!", icon="🔥")
                    send_line_alert(f"🔒 [ISO Approved]: หัวหน้างาน ({boss_name}) ได้อนุมัติใบตรวจเช็คประจำวันที่ {current_day} ของเครื่อง {m_id} บนระบบคลาวด์เรียบร้อยแล้ว")
            
            current_notes = get_current_cloud_note(m_id, m_type_flag)
            st.text_area("📝 รายการอาการเสียสะสมปัจจุบัน (ช่อง B)", value=current_notes, key=f"note_area_{m_id}", height=100, disabled=True)
            st.divider()

        categories = {
            "🔹 แผนกเครื่อง CNC (9 เครื่อง)": lambda k, v: "CNC" in k and "CRANE" not in k.upper() and "QC-" not in k.upper(),
            "🔹 แผนกเครน CRANE / HOIST (2 ตัว)": lambda k, v: "CRANE" in k.upper(),
            "🔹 แผนกเครื่องมือวัดคุณภาพ QC (21 เครื่อง)": lambda k, v: "QC-" in k.upper(),
            "🔹 แผนกเครื่องเจียรผิว GRINDING (2 เครื่อง)": lambda k, v: "GRINDING" in k and "CUTTER" not in k,
            "🔹 แผนกเครื่องลับคม CUTTER GRINDING (1 เครื่อง)": lambda k, v: "CUTTER" in k,
            "🔹 แผนกเครื่องมิลลิ่ง MILLING (3 เครื่อง)": lambda k, v: "MILLING" in k,
            "🔹 แผนกเครื่องตัด CUTTING (2 เครื่อง)": lambda k, v: "CUTTING" in k,
            "🔹 แผนกเครื่องเชื่อม MIG CO2 (3 เครื่อง)": lambda k, v: "MIG" in k,
            "🔹 แผนกเครื่องเชื่อม ARGON (1 เครื่อง)": lambda k, v: "ARGON" in k,
            "🔹 แผนกเครื่องเลื่อยสายพาน BAND SAW (3 เครื่อง)": lambda k, v: "BAND" in k,
            "🔹 แผนกปั๊มลม AIR COMPRESSOR (2 เครื่อง)": lambda k, v: "COMP-" in k.upper()
        }

        for cat_title, filter_func in categories.items():
            st.write(f"#### {cat_title}")
            cols = st.columns(3)
            idx = 0
            for m_id, m_name in MACHINES.items():
                if filter_func(m_id, m_name):
                    with cols[idx % 3]:
                        render_machine_card(m_id, m_name, get_machine_type(m_id))
                    idx += 1

    elif password_input != "": 
        st.error("❌ รหัสผ่านไม่ถูกต้อง")

    with st.expander("🖨️ เครื่องมือหัวหน้างาน: พิมพ์ QR Code สำหรับไปแปะหน้าเครื่องจักร"):
        sel_m = st.selectbox("เลือกเครื่องที่ต้องการพิมพ์ QR:", list(MACHINES.keys()))
        base_url = "https://pes-maintenance.streamlit.app/" 
        qr_url = f"{base_url}?machine_id={sel_m}" 
        qr = qrcode.make(qr_url)
        buf = BytesIO()
        qr.save(buf)
        st.image(buf.getvalue(), caption=f"QR สำหรับแปะหน้าเครื่อง {MACHINES[sel_m]}")
