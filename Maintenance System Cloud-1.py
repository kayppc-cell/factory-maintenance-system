import base64
import datetime
from io import BytesIO
import json
import mimetypes
import os
import shutil
import time
import zipfile
import gc
import html
from urllib.parse import quote
import streamlit as st
import streamlit.components.v1 as components
from supabase import create_client, Client

# --- 1. CONFIGURATION & SUPABASE SETUP ---
# กำหนด Base URL ของระบบ Streamlit App
DEFAULT_APP_URL = "https://factory-maintenance-system.streamlit.app"

BASE_FOLDER = (
    os.path.dirname(os.path.abspath(__file__))
    if "__file__" in locals()
    else os.getcwd()
)

# รองรับทั้ง st.secrets บน Streamlit Cloud และ Environment Variables
def get_secret(key_name, default=""):
    if key_name in st.secrets:
        return st.secrets[key_name]
    return os.getenv(key_name, default)

SUPABASE_URL = get_secret("SUPABASE_URL")
SUPABASE_KEY = get_secret("SUPABASE_KEY")
LINE_ACCESS_TOKEN = get_secret("LINE_ACCESS_TOKEN")
LINE_TARGET_ID = get_secret("LINE_TARGET_ID")
BOSS_PASSWORD = get_secret("BOSS_PASSWORD")
BIGBOSS_PASSWORD = get_secret("BIGBOSS_PASSWORD")

@st.cache_resource
def init_supabase():
    try:
        if not SUPABASE_URL or not SUPABASE_KEY:
            return None
        return create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception as e:
        print(f"Supabase init error: {e}")
        return None

supabase = init_supabase()

now = datetime.datetime.now()
current_time_str = now.strftime("%Y-%m-%d %H:%M:%S")

MACHINES = {
    "CNC3X-01": "CNC 3 แกน #01", "CNC3X-02": "CNC 3 แกน #02",
    "CNC3X-03": "CNC 3 แกน #03", "CNC3X-04": "CNC 3 แกน #04",
    "CNC3X-05": "CNC 3 แกน #05", "CNC3X-06": "CNC 3 แกน #06",
    "CNC3X-07": "CNC 3 แกน #07", "CNC3X-08": "CNC 3 แกน #08",
    "CNC5X-01": "CNC 5 แกน #พิเศษ",
    "Crane no.1": "เครน CNC NO.1", "Crane no.2": "เครน QC NO.2",
    "QC-01": "เครื่องวัดความแข็ง",
    "QC-02": "เวอร์เนีย 1000 QC-VN-009", "QC-03": "เวอร์เนีย 300 QC-VN-011",
    "QC-04": "เวอร์เนีย 300 QC-VN-029", "QC-05": "เวอร์เนีย 200 QC-VN-025",
    "QC-06": "เวอร์เนีย 200 QC-VN-026", "QC-07": "เวอร์เนีย 200 QC-VN-027",
    "QC-08": "เวอร์เนีย 200 QC-VN-028", "QC-09": "เวอร์เนีย 600 QC-VN-010",
    "QC-10": "ไฮเกจ 300 QC-HG-006", "QC-11": "ไฮเกจ 600 QC-HG-007",
    "QC-12": "ไฮเกจ 1000 QC-HG-008", "QC-13": "ไมโคร 0-25 QC-MC-013",
    "QC-14": "ไมโคร 5-30 QC-MC-024", "QC-15": "CMM QC-CMM-001",
    "QC-16": "Laser QC-Laser-001", "QC-17": "เลื่อยสายพาน QC-SAW-001",
    "QC-18": "Faro Arm QC-AC-001", "QC-19": "Cimcore Arm 2.8 QC-AC-003",
    "QC-20": "Cimcore Arm 2.4 QC-AC-002", "QC-21": "Cimcore Arm 3.5 QC-AC-004",
    "COMP-01": "ปั๊มลม 1 COMP-01", "COMP-02": "ปั๊มลม 2 COMP-02",
    "GRINDING-01": "เครื่องเจียร GRINDING #01", "GRINDING-02": "เครื่องเจียร GRINDING #02",
    "CUTTER GRINDING-01": "เครื่องลับคม CUTTER GRINDING #01",
    "MILLING-01": "เครื่องมิลลิ่ง #01", "MILLING-02": "เครื่องมิลลิ่ง #02",
    "MILLING-03": "เครื่องมิลลิ่ง #03", "MILLING-04": "เครื่องเฟสเก็บขนาด",
    "LATHE-01": "เครื่องกลึง LATHE #01",
    "CUTTING-01": "เครื่องตัด CUTTING #01",
    "MIG CO2-01": "เครื่องเชื่อม MIG CO2 #01", "MIG CO2-02": "เครื่องเชื่อม MIG CO2 #02",
    "MIG CO2-03": "เครื่องเชื่อม MIG CO2 #03",
    "ARGON-01": "เครื่องเชื่อม ARGON #01", "ARGON-02": "เครื่องเชื่อม ARGON #02",
    "WELDING_ALUMINUM-01": "เครื่องเชื่อมอลูมิเนียม WELDING ALUMINUM #01",
    "BAND SAW-01": "เครื่องเลื่อยสายพาน #01", "BAND SAW-02": "เครื่องเลื่อยสายพาน #02",
    "BAND SAW-03": "เครื่องเลื่อยสายพาน #03",
    "FORKLIFT-01": "รถโฟคลิฟ FORKLIFT #01",
    "CAR-2ฒข-5050": "รถยนต์ทะเบียน 2ฒข-5050",
    "CAR-2ฒข-5353": "รถยนต์ทะเบียน 2ฒข-5353",
    "CAR-2ฒฆ-5151": "รถยนต์ทะเบียน 2ฒฆ-5151",
    "CAR-1ฒถ-5252": "รถยนต์ทะเบียน 1ฒถ-5252",
    "Truck-83-2329": "รถบรรทุกทะเบียน 83-2329",
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
    "QC-02": ["ตรวจดูสภาพของเวอร์เนียพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด", "ตรวจสอบที่คีบตรงปลายที่ใช้วัดชิ้นงาน เช็คว่ามีรอยบิ่น หรือสึกหล่อหรือไม่", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-03": ["ตรวจดูสภาพของเวอร์เนียพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด", "ตรวจสอบที่คีบตรงปลายที่ใช้วัดชิ้นงาน เช็คว่ามีรอยบิ่น หรือสึกหล่อหรือไม่", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-04": ["ตรวจดูสภาพของเวอร์เนียพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด", "ตรวจสอบที่คีบตรงปลายที่ใช้วัดชิ้นงาน เช็คว่ามีรอยบิ่น หรือสึกหล่อหรือไม่", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-05": ["ตรวจดูสภาพของเวอร์เนียพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด", "ตรวจสอบที่คีบตรงปลายที่ใช้วัดชิ้นงาน เช็คว่ามีรอยบิ่น หรือสึกหล่อหรือไม่", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-06": ["ตรวจดูสภาพของเวอร์เนียพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด", "ตรวจสอบที่คีบตรงปลายที่ใช้วัดชิ้นงาน เช็คว่ามีรอยบิ่น หรือสึกหล่อหรือไม่", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-07": ["ตรวจดูสภาพของเวอร์เนียพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด", "ตรวจสอบที่คีบตรงปลายที่ใช้วัดชิ้นงาน เช็คว่ามีรอยบิ่น หรือสึกหล่อหรือไม่", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-08": ["ตรวจดูสภาพของเวอร์เนียพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด", "ตรวจสอบที่คีบตรงปลายที่ใช้วัดชิ้นงาน เช็คว่ามีรอยบิ่น หรือสึกหล่อหรือไม่", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-09": ["ตรวจดูสภาพของเวอร์เนียพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด", "ตรวจสอบที่คีบตรงปลายที่ใช้วัดชิ้นงาน เช็คว่ามีรอยบิ่น หรือสึกหล่อหรือไม่", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-10": ["ตรวจดูสภาพของไมโครพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-11": ["ตรวจดูสภาพของ Hight Gauge พร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-12": ["ตรวจดูสภาพของ Hight Gauge พร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-13": ["ตรวจดูสภาพของไมโครพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการหมุนเข้าหมุนออกต้องไม่ติดขัด", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-14": ["ตรวจดูสภาพของไมโครพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการหมุนเข้าหมุนออกต้องไม่ติดขัด", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-15": ["ตรวจดูสภาพของสายไฟ", "ตรวจดูสภาพของลมพร้อมใช้งานหรือไม่", "ตรวจดูสภาพของ SURFACE BASE", "ตรวจดูสภาพของ STICKER", "ตรวจสอบ COMPUTER", "ตรวจดูสภาพของหัว PROBE คตงอหรือไม่"],
    "QC-16": ["ตรวจดูสภาพของสายไฟ", "ตรวจดูสภาพ ของเครื่อง", "ตรวจดูสภาพของ Lend Laser", "ตรวจสอบ STICKER", "ตรวจสอบ COMPUTER"],
    "QC-17": ["ตรวจดูสภาพสายไฟ", "ตรวจดูสภาพของใบเลื่อย", "ตรวจดูสภาพของมอเตอร์", "ตรวจดูสภาพของสายพาน", "ตรวจสอบเครื่องเลื่อยสายพาน"],
    "QC-18": ["ตรวจดูสภาพของสายไฟ", "ตรวจดูสภาพ ARM ของเครื่อง", "ตรวจดูสภาพของหัว PROBE คตงอหรือไม่", "ตรวจสอบ STICKER", "ตรวจสอบ NOTEBOOK COMPUTER"],
    "QC-19": ["ตรวจดูสภาพของสายไฟ", "ตรวจดูสภาพ ARM ของเครื่อง", "ตรวจดูสภาพของหัว PROBE คตงอหรือไม่", "ตรวจสอบ STICKER", "ตรวจสอบ NOTEBOOK COMPUTER"],
    "QC-20": ["ตรวจดูสภาพของสายไฟ", "ตรวจดูสภาพ ARM ของเครื่อง", "ตรวจดูสภาพของหัว PROBE คตงอหรือไม่", "ตรวจสอบ STICKER", "ตรวจสอบ NOTEBOOK COMPUTER"],
    "QC-21": ["ตรวจดูสภาพของสายไฟ", "ตรวจดูสภาพ ARM ของเครื่อง", "ตรวจดูสภาพของหัว PROBE คตงอหรือไม่", "ตรวจสอบ STICKER", "ตรวจสอบ NOTEBOOK COMPUTER"],
    "COMP-01": ["เช็คแรงดัน (Pressure) ต้องไม่ต่ำกว่า 7 bar", "ตรวจสอบระดับน้ำมันไฮดรอลิก ต้องไม่ต่ำกว่าระดับต่ำสุด", "เช็คอุณหภูมิความร้อนต้องไม่เกิน 80 องศา", "เช็คการรั่วซีมของระบบน้ำมัน", "เช็คระบบเดรนน้ำ (Water Draen)"],
    "COMP-02": ["เช็คแรงดัน (Pressure) ต้องไม่ต่ำกว่า 7 bar", "ตรวจสอบระดับน้ำมันไฮดรอลิก ต้องไม่ต่ำกว่าระดับต่ำสุด", "เช็คอุณหภูมิความร้อนต้องไม่เกิน 80 องศา", "เช็คการรั่วซีมของระบบน้ำมัน", "เช็คระบบเดรนน้ำ (Water Draen)"],
    "GRINDING-01": ["การ Worm spindle และ TABLE SLIDE", "เช็คระดับนำมันไฮดรอลิก และ การทำงานของ PUMP", "เช็คระดับของน้ำยา COOLANNT PUMP", "ตรวจสอบการทำงานของแม่เหล็ก", "ตรวจสอบการทำงานของ SLIDE X,Y", "ตรวจสอบสภาพความพร้อมโดยรวมของเครื่องจักร", "ตรวจสอบระดับน้ำมันของ PUMPน้ำมันหล่อลื่น", "ตรวจสอบการทำงานของไฟฟ้าและแสงสว่าง", "ตรวจสอบการทำงานของตัวดูดอากศ"],
    "GRINDING-02": ["การ Worm spindle และ TABLE SLIDE", "เช็คระดับนำมันไฮดรอลิก และ การทำงานของ PUMP", "เช็คระดับของน้ำยา COOLANNT PUMP", "ตรวจสอบการทำงานของแม่เหล็ก", "ตรวจสอบการทำงานของ SLIDE X,Y", "ตรวจสอบสภาพความพร้อมโดยรวมของเครื่องจักร", "ตรวจสอบระดับน้ำมันของ PUMPน้ำมันหล่อลื่น", "ตรวจสอบการทำงานของไฟฟ้าและแสงสว่าง", "ตรวจสอบการทำงานของตัวดูดอากศ"],
    "CUTTER GRINDING-01": ["การ WORM UP แกน Y พร้อมใช้งาน", "การ WORM UP แกน Z พร้อมใช้งาน", "ตรวจสอบการทำงานของไฟฟ้าและแสงสว่าง", "ตรวจสอบการทำงานของมอเตอร์ มีการหมุนปกติ", "ตรวจสอบการจับหัวคอเรต"],
    "MILLING": [
        "Worm Spindle ก่อนเริมงาน ตรวจสอบความ ผิดปกติของชุด  Back gauge  และ Motor", 
        "เช็ค Auto  Up-Down back gauge  และ Manual ( ความคร่องตัวในการเคลื่อนที่ของ Spindle )", 
        "ตรวจสอบการ SLIDE  ของแกน X", "ตรวจสอบการ SLIDE  ของแกน Y", "ตรวจสอบการ SLIDE  ของแกน Z", 
        "ระดับน้ำมันไฮดรอลิค ตรวจสอบน้ำมันในปั้มน้ำมันหล่อ ลื่นแกน  X,Y,Z", 
        "ตรวจน้ำมันหล่อลื่นเย็น ตรวจสอบการทำงานของปั้ม COOLANT และสภาพของน้ำ  COOLANT", 
        "ตรวจสอบหน้าจอ  DIGITAL READ OUT และการทำ งานของ LINEAR SCALE", "หยอดน้ำมันหล่อลื่นทุกวันจันทร์", 
        "ตรวจสอบการทำงานของไฟฟ้าแสงสว่างของเครื่อง", "ตรวจสอบสภาพความพร้อมโดยรวมของเครื่องจักร  และอุกรณ์เสริมต่าง ๆ"
    ],
    "LATHE": [
        "Spindle ก่อนเริ่มงาน 15 นาที", "เช็คระดับน้ำมันเครื่องในห้องเกียร์", "เช็คระบบเฟื่องทดลองเปลี่ยนรอบตวามเร็วต่าง ๆ",
        "เช็ค AUTO แกน X,Y", "เช็คระดับน้ำมันหล่อลื่นใน PUMP", "เช็คน้ำมันหล่อเย็นและการทำงานของปั้ม",
        "ตรวจสอบหน้าจอ  DIGITAL  READ OUT และการ ทำงานของ  LINEAR  SCALE", "ตรวจเช็คสภาพและความตึงของสายพาน",
        "ตรวจสอบการทำงานของไฟฟ้าแสงสว่าง", "อัดจาระบีตามหัวอัดจาระบีทุก ๆ จุด", "ตรวจสอบความพร้อมสภาพโดยรวมของเครื่อง"
    ],
    "CUTTING": ["การ Worm spindle ก่อนเริ่มงาน เพื่อตรวจ ความผิดปกติของชุด Back gauge และ Motor", "เช็ค Auto Up-Down back gauge และ Manual ( ความคล่องตัวในการเคลื่อนที่ )", "ระดับน้ำมันไฮดรอลิค ตรวจสอบระดับในปั้มน้ำมัน หล่อลืนแกน  Back gauge", "ตรวจเช็ค  Switch  เปิด-ปิด", "ตรวจสอบ Digital  read out และการทำงาน ของ Linear  scale", "อัดจาระบีตามจุดที่อัดจาระบีทุกๆจุด", "ตรวจสอบใบมีด  บนและล่าง", "ตรวจสอบความพร้อมสภาพโดยรวมของเครื่อง จักรและอุปกรณ์เสริมต่าง ๆ"],
    "MIG CO2": ["ตรวจสภาพความพร้อมโดยรวมของเครื่อง", "เช็ค BREAKER เพื่อเช็คระบบไฟฟ้า ตามตำแหน่งไฟ โชว์ และสวิชท์ต่าง ๆ", "ตรวจสภาพความพร้อมของมาตราวัดแรงดัน ของก๊าซ CO2 และปรับตั้งอย่างถูกต้อง", "ตรวจจุดต่อของก๊าซ CO2 รั่วหรือไม่", "ตรวจสภาพความพร้อมของสายไฟ สายก๊าซ  CO2 ว่ารั่วหรือไม่", "ตรวจสภาพความพร้อมของสายกราวด์", "ทำความสะอาดหัวเชื่อมก่อนใช้งาน"],
    "ARGON": [
        "ตรวจสภาพความพรัอมโดยรวมของเครื่อง", "เช็ค  BREAKER  เพื่อเช็คระบบไฟฟ้า ตามตำแหน่งไฟ โชว์  และ SWITCH  ต่าง ๆ", 
        "ตรวจสภาพความพร้อมของมาตราวัดแรงดันของมาตรา วัดแรงดันของก๊าช  ARGON  และปรับตั้งอย่างถูกวิธี", "ตรวจุดต่อของสายก๊าช  ARGON  ก่อนว่ารั่วหรือไม่", 
        "ตรวจสภาพความพร้อมของสายกราว์", "ตรวจสภาพความพร้อมของสายไฟฟ้าสายก๊าช  ARGON และชุดหัวเชื่อม", 
        "ตรวจสภาพความพร้อมของ  SWITCH  หัวเชื่อม", "ทำความสะดาดชุดหัวเชื่อมก่อนใช้งาน"
    ],
    "WELDING_ALUMINUM": [
        "ตรวจสภาพความพรัอมโดยรวมของเครื่อง", "เช็ค  BREAKER  เพื่อเช็คระบบไฟฟ้า ตามตำแหน่งไฟ โชว์  และ SWITCH  ต่าง ๆ",
        "ตรวจสภาพความพร้อมของสายกราวด์ให้อยู๋ในสภาพ ความพร้อมอยู่เสมอ", "ตรวจจุดต่อของสายท่อแก๊สว่ารั่วหรือไม่",
        "เช็คระดับน้ำหล่อเย็นให้อยู่ในระดับพร้อมใช้งาน", "เช็คระดับแรงดันในถังแก๊สให้พร้อมใช้งาน",
        "ทำความสะอาดชุดหัวเชื่อมก่อนใช้งานอย่างสม่ำเสมอ"
    ],
    "BAND SAW": ["เช็ค Auto Up-Down Back Gauge และ Manual (ความคล่องตัวในการเคลื่อนที่ of Spindle)", "เช็คระดับน้ำมันไฮดรอลิค", "ตรวจน้ำมันหล่อลื่นเย็น ตรวจสอบการทำงานของปั๊ม COOLANT และสภาพของน้ำ COOLANT", "ตรวจสอบ Switch (สวิตซ์) หน้า BOX CONTROL", "ตรวจสอบระดับน้ำมันหล่อลื่นในห้องเกียร์"],
    "FORKLIFT": [
        "ตรวจเช็คระบบน้ำหม้อน้ำให้อยู่ในระดับ Hight", "ตรวจเช็คน้ำมันเครื่องยนต์ต้องอยู่ไม่เกินขีดที่3ของตัวเช็ค", "ตรวจเช็คไส้กรองและเป่าลมทำความสะอาด",
        "ตรวจเช็คการรั่วซึมofน้ำมันไฮดรอริก", "ตรวจเช็คระบบเบรคและน้ำมันเบรค", "ตรวจเช็คไฟส่องสว่างและไฟเลี้ยว",
        "ตรวจเช็คสัญญานแตร"
    ],
    "VEHICLE": [
        "เช็คยางรถยนต์ ไม่บวม หรือฉีกขาด",
        "เช็คหน้ากระจกรถ ต้องไม่มีรอยแตกร้าว",
        "เช็คระบบไฟหน้า ไฟเบรก ไฟเลี้ยว ไฟหรี่ และไฟฉุกเฉิน ต้องใช้ได้ทุกดวง",
        "เช็คที่ปัดน้ำฝน ต้องไม่แข็ง แห้ง หรือกรอบ",
        "เช็คเลขหน้าปัดไมล์ และถ่ายรูปเลขหน้าปัดไมล์ยืนยันก่อนเริ่มใช้งาน",
        "เช็คระดับน้ำมันเบรก ต้องอยู่ในระดับปกติ",
        "เช็คระดับน้ำมันเครื่อง ต้องอยู่ในระดับปกติ",
        "เช็คระดับหม้อพักน้ำ ต้องอยู่ในระดับปกติ",
        "เช็คระดับหม้อพักน้ำฉีดกระจก ต้องอยู่ในระดับปกติ",
        "เช็คระบบสายพานต่าง ๆ ต้องไม่เปื่อยหรือมีรอยฉีกขาด",
        "ตรวจสอบเอกสาร ประกัน / พ.ร.บ. / ภาษี"
    ]
}

PHOTO_RULES = {
    "CNC": [2, 3, 4, 5, 8, 13], "Crane no.1": [3, 4], "Crane no.2": [3, 4], "QC-01": [4],
    "QC-02": [2, 4], "QC-03": [2, 4], "QC-04": [2, 4], "QC-05": [2, 4], "QC-06": [2, 4],
    "QC-07": [2, 4], "QC-08": [2, 4], "QC-09": [2, 4], "QC-10": [2, 3], "QC-11": [2, 3], "QC-12": [2, 3],
    "QC-13": [2, 3], "QC-14": [2, 3], "QC-15": [6], "QC-16": [3], "QC-17": [2], "QC-18": [3], "QC-19": [3],
    "QC-20": [3], "QC-21": [3], "COMP-01": [1, 2, 3], "COMP-02": [1, 2, 3], "GRINDING-01": [2, 4, 7], "GRINDING-02": [4, 7],
    "CUTTER GRINDING-01": [], "MILLING": [6, 7], "LATHE": [2, 5], "CUTTING": [3, 5, 7], "MIG CO2": [3, 4, 5],
    "ARGON": [3, 4, 6], "WELDING_ALUMINUM": [5, 6], "BAND SAW": [3, 5], "FORKLIFT": [1, 2, 5],
    "VEHICLE": list(range(1, 12))
}

def get_machine_type_by_id(machine_id):
    u_id = str(machine_id).upper().strip()
    if u_id.startswith("CAR-") or u_id.startswith("TRUCK-"): return "VEHICLE"
    elif "CUTTER" in u_id: return "CUTTER GRINDING-01"
    elif "MILLING" in u_id or "MILL" in u_id: return "MILLING"
    elif "CRANE NO.1" in u_id or "CRANE NO. 1" in u_id: return "Crane no.1"
    elif "CRANE NO.2" in u_id or "CRANE NO. 2" in u_id: return "Crane no.2"
    elif any(f"QC-{i:02d}" in u_id for i in range(1, 22)):
        for i in range(1, 22):
            if f"QC-{i:02d}" in u_id: return f"QC-{i:02d}"
    elif "COMP-01" in u_id: return "COMP-01"
    elif "COMP-02" in u_id: return "COMP-02"
    elif "GRINDING-01" in u_id: return "GRINDING-01"
    elif "GRINDING-02" in u_id: return "GRINDING-02"
    elif "LATHE" in u_id: return "LATHE"
    elif "CUTTING" in u_id: return "CUTTING"
    elif "MIG" in u_id: return "MIG CO2"
    elif "ARGON" in u_id: return "ARGON"
    elif "WELDING_ALUMINUM" in u_id: return "WELDING_ALUMINUM"
    elif "BAND" in u_id: return "BAND SAW"
    elif "FORKLIFT" in u_id: return "FORKLIFT"
    return "CNC"

def get_coordinates_by_machine(m_id, m_type):
    u_id = str(m_id).upper().strip()

    # แบบฟอร์มรถ: รายการตรวจแถว 6-16, ผู้ตรวจแถว 17, ผู้อนุมัติแถว 19, หมายเหตุ B22
    if m_type == "VEHICLE" or u_id.startswith("CAR-") or u_id.startswith("TRUCK-"):
        return 17, 19, "B22"
    
    # ⚡ 1. ล็อกเฉพาะ MILLING-04 (ช่างแถว 17, หัวหน้าแถว 19, บันทึกเพิ่มเติม B22)
    if "MILLING-04" in u_id or "MILLING NO. 4" in u_id or "MILLING NO.4" in u_id or "MILLING_04" in u_id:
        return 17, 19, "B22"

    # ⚡ 2. MILLING-01, 02, 03 (ช่างแถว 20, หัวหน้าแถว 22, บันทึกเพิ่มเติม B25)
    if "MILLING" in u_id or "MILL" in u_id or m_type == "MILLING": 
        return 20, 22, "B25"
        
    if "CUTTER" in u_id or m_type == "CUTTER GRINDING-01": return 13, 15, "B18"
    
    if any(k in u_id for k in ["QC-01", "QC-10", "QC-11", "QC-12", "QC-13", "QC-14"]): 
        return 10, 12, "B15"

    if any(k in u_id for k in ["QC-02", "QC-03", "QC-04", "QC-05", "QC-06", "QC-07", "QC-08", "QC-09", "QC-16", "QC-17", "QC-18", "QC-19", "QC-20", "QC-21"]): 
        return 11, 13, "B16"

    if "QC-15" in u_id: return 12, 14, "B17"
    if "ARGON-02" in u_id or "ARGON-01" in u_id: return 14, 16, "B19"
    if m_type == "FORKLIFT" or "FORKLIFT" in u_id: return 13, 15, "B18"
    if m_type == "CNC" or "CNC" in u_id: return 22, 24, "B28"
    if "CRANE" in u_id: return 14, 16, "B19"
    if "GRINDING" in m_type or "GRINDING" in u_id: return 16, 18, "B21"
    if m_type == "LATHE" or "LATHE" in u_id: return 17, 19, "B22"
    
    # ⚡ 3. CUTTING-01 (ช่างแถว 14, ผู้ตรวจสอบแถว 16, บันทึกเพิ่มเติม B19)
    if m_type == "CUTTING" or "CUTTING" in u_id: 
        return 14, 16, "B19"
        
    if m_type == "WELDING_ALUMINUM" or "WELDING_ALUMINUM" in u_id: return 13, 15, "B18"
    if m_type == "MIG CO2" or "MIG" in u_id: return 13, 15, "B18"
    if m_type == "BAND SAW" or "BAND" in u_id: return 11, 13, "B16"
    return 11, 13, "B16"

def set_cell_value_safe(ws, coordinate_str, value, alignment=None):
    try:
        cell = ws[coordinate_str]
        cell.value = value
        if alignment: cell.alignment = alignment
    except Exception as e:
        print(f"Cell write error at {coordinate_str}: {e}")

# --- 2. REALTIME SUPABASE ENGINE WITH PAGINATION ---
def save_log_to_supabase_bulk(list_of_logs):
    if not supabase or not list_of_logs:
        return False
    try:
        # อ่าน ID ชุดเดิมก่อน แล้วบันทึกชุดใหม่ให้สำเร็จก่อนค่อยลบชุดเก่า
        # จึงไม่ทำข้อมูลเดิมหายหากการ insert ใหม่ล้มเหลว
        contexts = {
            (
                str(log["machine_id"]), str(log["year_month"]),
                int(log["day_num"]), str(log["role"])
            )
            for log in list_of_logs
        }
        old_ids = []
        for machine_id, year_month, day_num, role in contexts:
            old = supabase.table("maintenance_logs").select("id")\
                .eq("machine_id", machine_id)\
                .eq("year_month", year_month)\
                .eq("day_num", day_num)\
                .eq("role", role)\
                .execute()
            old_ids.extend(row["id"] for row in (old.data or []) if row.get("id") is not None)

        supabase.table("maintenance_logs").insert(list_of_logs).execute()
        if old_ids:
            for start in range(0, len(old_ids), 200):
                supabase.table("maintenance_logs").delete().in_("id", old_ids[start:start + 200]).execute()
        st.cache_data.clear()
        return True
    except Exception as e:
        print(f"Supabase Safe Replace Error: {e}")
        return False

LOG_COLUMNS = "id,timestamp,machine_id,day_num,year_month,tech_name,item_no,checklist_item,status,note,role"

def logs_to_dataframe(rows):
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows).rename(columns={
        "timestamp": "Timestamp", "machine_id": "Machine_ID",
        "day_num": "Day_Num", "year_month": "Year_Month",
        "tech_name": "Tech_Name", "item_no": "Item_No",
        "checklist_item": "Checklist_Item", "status": "Status",
        "note": "Note", "role": "Role"
    })
    if "Day_Num" in df.columns:
        df["Day_Num"] = pd.to_numeric(df["Day_Num"], errors="coerce").fillna(0).astype(int)
    return df

@st.cache_data(ttl=30, show_spinner=False)
def fetch_day_logs(year_month, day_num):
    if not supabase: return pd.DataFrame()
    try:
        all_data = []
        page_size = 1000
        start = 0
        while True:
            res = supabase.table("maintenance_logs").select(LOG_COLUMNS)\
                .eq("year_month", year_month)\
                .eq("day_num", int(day_num))\
                .order("id")\
                .range(start, start + page_size - 1)\
                .execute()
            if not res.data:
                break
            all_data.extend(res.data)
            if len(res.data) < page_size:
                break
            start += page_size
            
        return logs_to_dataframe(all_data)
    except Exception as e:
        print(f"Supabase Fetch Day Error: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=60, show_spinner=False)
def fetch_month_logs(year_month):
    if not supabase: return pd.DataFrame()
    try:
        all_data = []
        page_size = 1000
        start = 0
        while True:
            res = supabase.table("maintenance_logs").select(LOG_COLUMNS)\
                .eq("year_month", year_month)\
                .order("id")\
                .range(start, start + page_size - 1)\
                .execute()
            if not res.data:
                break
            all_data.extend(res.data)
            if len(res.data) < page_size:
                break
            start += page_size
            
        return logs_to_dataframe(all_data)
    except Exception as e:
        print(f"Fetch Month Log Error: {e}")
        return pd.DataFrame()

def fetch_machine_all_month_logs(machine_id, year_month, month_logs=None):
    df = fetch_month_logs(year_month) if month_logs is None else month_logs
    if df.empty or "Machine_ID" not in df.columns:
        return pd.DataFrame()
    target = str(machine_id).strip().upper()
    normalized_ids = df["Machine_ID"].astype(str).str.strip().str.upper()
    return df[normalized_ids.eq(target)].copy()

def generate_excel_bytes(machine_id, year_month, m_type, target_day=None):
    import openpyxl
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter
    excel_file_name = f"FM-MN-07_{machine_id}.xlsx"
    target_excel_path = os.path.join(BASE_FOLDER, excel_file_name)
    if not os.path.isfile(target_excel_path): return None
    
    df_logs = fetch_machine_all_month_logs(machine_id, year_month)
    if target_day is not None and not df_logs.empty:
        df_logs = df_logs[df_logs["Day_Num"].eq(int(target_day))].copy()
    try:
        wb = openpyxl.load_workbook(target_excel_path, data_only=False)
        ws = wb.active
        if m_type == "VEHICLE":
            # ปรับหัวแบบฟอร์มและหน้ากระดาษให้ตรงกับรถที่เลือกทุกครั้งที่ดาวน์โหลด
            vehicle_plate = machine_id.split("-", 1)[1] if "-" in machine_id else machine_id
            ws["A1"] = f"ใบตรวจสอบสภาพรถยนต์ {vehicle_plate}"
            ws["AB1"] = machine_id
            ws.print_area = "A1:AG28"
            ws.page_setup.orientation = "landscape"
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_margins.left = 0
            ws.page_margins.right = 0
            ws.page_margins.top = 0
            ws.page_margins.bottom = 0
            ws.page_margins.header = 0
            ws.page_margins.footer = 0
        t_row, boss_row, n_cell = get_coordinates_by_machine(machine_id, m_type)
        center_align = Alignment(horizontal='center', vertical='center')
        
        # ⚡ 1. ล้างข้อมูลรอยติ๊กเก่า วันที่ 1-31 ทุกข้อตรวจ และช่องลงชื่อ ให้สะอาดก่อนเสมอ
        for d in range(1, 32):
            col_letter = get_column_letter(2 + d)
            # ล้างรอยติ๊กข้อตรวจ (แถว 6 ถึง 21)
            for r in range(6, 22):
                set_cell_value_safe(ws, f"{col_letter}{r}", None)
            # ล้างช่องชื่อช่างและชื่อหัวหน้า
            set_cell_value_safe(ws, f"{col_letter}{t_row}", None)
            set_cell_value_safe(ws, f"{col_letter}{boss_row}", None)
        # ล้างช่องบันทึกเพิ่มเติม
        set_cell_value_safe(ws, n_cell, None)
        
        # ⚡ 2. นำข้อมูล Log เฉพาะของเดือนที่เลือก (year_month) มาเขียนลงตาราง
        if not df_logs.empty:
            df_logs = df_logs.sort_values(by="Timestamp")
            notes_accumulator = []
            
            for _, row in df_logs.iterrows():
                day_val = int(row["Day_Num"])
                if not (1 <= day_val <= 31):
                    continue
                    
                col_letter = get_column_letter(2 + day_val)
                role_val = str(row["Role"]).strip().lower()
                tech_boss_name = str(row["Tech_Name"]).strip()
                
                if role_val == "tech":
                    status_val = str(row["Status"]).strip()
                    item_idx = int(row["Item_No"])
                    note_val = str(row["Note"]).strip()
                    
                    if item_idx > 0:
                        cell_coord = f"{col_letter}{5 + item_idx}"
                        
                        if "ไม่ได้ทำงาน" in status_val or status_val == "-":
                            mark = "-"
                        elif "แก้ไข" in status_val and "ปกติ" in status_val:
                            mark = "⨂"
                        elif "ปกติ" in status_val and "แก้ไข" not in status_val:
                            mark = "/"
                        elif "ต้องแก้ไข" in status_val or "ไม่ได้" in status_val:
                            mark = "X"
                        else:
                            mark = "-"
                            
                        set_cell_value_safe(ws, cell_coord, mark, center_align)
                        
                        if note_val and note_val.lower() != "nan":
                            notes_accumulator.append(f"[วันที่ {day_val}]: ข้อ {item_idx} {note_val}")
                        
                    if tech_boss_name and tech_boss_name.lower() != "nan":
                        set_cell_value_safe(ws, f"{col_letter}{t_row}", tech_boss_name, Alignment(text_rotation=90, horizontal='center', vertical='center'))
                elif role_val == "boss":
                    if tech_boss_name and tech_boss_name.lower() != "nan":
                        set_cell_value_safe(ws, f"{col_letter}{boss_row}", tech_boss_name, Alignment(text_rotation=90, horizontal="center", vertical="center"))

            # บันทึกหมายเหตุลงช่องบันทึกเพิ่มเติม
            if notes_accumulator:
                combined_notes = ", ".join(notes_accumulator)
                set_cell_value_safe(ws, n_cell, combined_notes)

        output_stream = BytesIO()
        wb.save(output_stream)
        wb.close()
        data = output_stream.getvalue()
        output_stream.close()
        gc.collect()
        return data
    except Exception as e:
        print(f"Generate Excel Error: {e}")
        return None

def zip_all_factory_excel(year_month_key, target_day=None):
    zip_buffer = BytesIO()
    has_file = False
    try:
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for m_id in MACHINES.keys():
                m_type = get_machine_type_by_id(m_id)
                excel_bytes = generate_excel_bytes(m_id, year_month_key, m_type, target_day=target_day)
                if excel_bytes:
                    period_tag = f"Day_{target_day}" if target_day is not None else year_month_key
                    zip_file.writestr(f"FM-MN-07_{m_id}_{period_tag}.xlsx", excel_bytes)
                    has_file = True
        if not has_file: return None
        zip_buffer.seek(0)
        gc.collect()
        return zip_buffer
    except Exception as e:
        print(f"Zip All Excel Error: {e}")
        return None

# --- PHOTO & DUAL STORAGE (LOCAL + SUPABASE CLOUD) ---
def send_line_alert(msg_text):
    import requests
    if not LINE_ACCESS_TOKEN or not LINE_TARGET_ID:
        print("LINE alert skipped: missing LINE_ACCESS_TOKEN or LINE_TARGET_ID")
        return
    url = 'https://api.line.me/v2/bot/message/push'
    headers = {'Content-Type': 'application/json', 'Authorization': f'Bearer {LINE_ACCESS_TOKEN}'}
    payload = {"to": LINE_TARGET_ID, "messages": [{"type": "text", "text": msg_text}]}
    try: requests.post(url, headers=headers, data=json.dumps(payload), timeout=5)
    except Exception as e: print(f"ส่งไลน์ไม่สำเร็จ: {e}")

def save_uploaded_photos_dict(machine_id, day_num, uploaded_photos_dict, current_date_obj=None):
    if current_date_obj is None: current_date_obj = datetime.date.today()
    current_year_month = current_date_obj.strftime("%Y_%B")
    
    local_day_dir = os.path.join(BASE_FOLDER, "maintenance_photos", str(machine_id), current_year_month, f"Day_{day_num}")
    os.makedirs(local_day_dir, exist_ok=True)
    
    for item_idx, photo_data in uploaded_photos_dict.items():
        files_list = photo_data.get("files", [])
        if files_list:
            for f_order, uploaded_file in enumerate(files_list, 1):
                file_ext = os.path.splitext(uploaded_file.name)[1]
                file_name = f"photo_item_{item_idx}_{f_order}{file_ext}"
                file_bytes = uploaded_file.getvalue()
                content_type = mimetypes.guess_type(file_name)[0] or "application/octet-stream"
                
                full_local_path = os.path.join(local_day_dir, file_name)
                with open(full_local_path, "wb") as f_out:
                    f_out.write(file_bytes)
                
                if supabase:
                    try:
                        storage_path = f"{machine_id}/{current_year_month}/Day_{day_num}/{file_name}"
                        supabase.storage.from_("maintenance-photos").upload(
                            storage_path, file_bytes,
                            {"content-type": content_type, "upsert": "true"}
                        )
                    except Exception as e_up:
                        print(f"Photo Supabase Upload Error: {e_up}")
    gc.collect()

def get_machine_photos(machine_id, year_month, day_num):
    photos = []
    # Cloud เป็นแหล่งข้อมูลหลัก เพราะ local disk ของ Streamlit Cloud ไม่ถาวร
    if supabase:
        try:
            folder_path = f"{machine_id}/{year_month}/Day_{day_num}"
            files = supabase.storage.from_("maintenance-photos").list(folder_path)
            if files:
                for f in sorted(files, key=lambda x: x.get("name", "")):
                    f_name = f.get("name")
                    if f_name and f_name.lower().endswith(('.png', '.jpg', '.jpeg')):
                        file_data = supabase.storage.from_("maintenance-photos").download(f"{folder_path}/{f_name}")
                        photos.append((f_name, file_data))
        except Exception as e_sb:
            print(f"Supabase fetch photo error: {e_sb}")

    # ใช้ local เป็น fallback เมื่อ Cloud ไม่มีข้อมูลหรือเชื่อมต่อไม่ได้
    if not photos:
        local_dir = os.path.join(BASE_FOLDER, "maintenance_photos", str(machine_id), year_month, f"Day_{day_num}")
        if os.path.exists(local_dir):
            for f in sorted(os.listdir(local_dir)):
                if f.lower().endswith(('.png', '.jpg', '.jpeg')):
                    try:
                        with open(os.path.join(local_dir, f), "rb") as f_img:
                            photos.append((f, f_img.read()))
                    except Exception:
                        pass
    return photos

def list_storage_files_recursive(prefix=""):
    """คืน path ของไฟล์ทั้งหมดใน Supabase Storage รวมทุกโฟลเดอร์ย่อย"""
    if not supabase:
        return []
    paths = []
    entries = supabase.storage.from_("maintenance-photos").list(prefix)
    for entry in entries or []:
        name = entry.get("name")
        if not name:
            continue
        path = f"{prefix}/{name}" if prefix else name
        if entry.get("id") is not None or entry.get("metadata"):
            paths.append(path)
        else:
            paths.extend(list_storage_files_recursive(path))
    return paths

def delete_all_storage_photos():
    paths = list_storage_files_recursive()
    for start in range(0, len(paths), 100):
        supabase.storage.from_("maintenance-photos").remove(paths[start:start + 100])
    return len(paths)

PHOTO_DEPARTMENTS = [
    "ทั้งโรงงาน", "CNC", "GRINDING", "CRANE", "COMPRESSOR", "QC",
    "MILLING", "MIG CO2", "ARGON", "รถยนต์และรถบรรทุก",
    "เครื่องจักรอื่น ๆ (ตัด/กลึง/โฟคลิฟ)"
]

def machine_codes_by_department(filter_type):
    selected = []
    for machine_code in MACHINES:
        code = machine_code.upper()
        match = (
            filter_type == "ทั้งโรงงาน"
            or (filter_type == "CNC" and "CNC" in code)
            or (filter_type == "GRINDING" and "GRINDING" in code and "CUTTER" not in code)
            or (filter_type == "CRANE" and "CRANE" in code)
            or (filter_type == "COMPRESSOR" and "COMP-" in code)
            or (filter_type == "QC" and "QC-" in code)
            or (filter_type == "MILLING" and "MILLING" in code)
            or (filter_type == "MIG CO2" and "MIG" in code)
            or (filter_type == "ARGON" and "ARGON" in code)
            or (filter_type == "รถยนต์และรถบรรทุก" and (code.startswith("CAR-") or code.startswith("TRUCK-")))
            or (
                filter_type == "เครื่องจักรอื่น ๆ (ตัด/กลึง/โฟคลิฟ)"
                and any(k in code for k in ["CUTTING", "LATHE", "FORKLIFT", "WELDING_ALUMINUM", "BAND SAW", "CUTTER GRINDING"])
            )
        )
        if match:
            selected.append(machine_code)
    return selected

def zip_single_machine_photos(machine_id, target_date_obj, target_day=None):
    current_year_month = target_date_obj.strftime("%Y_%B")
    zip_buffer = BytesIO()
    has_file = False
    
    try:
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            days_to_check = [target_day] if target_day else range(1, 32)
                
            for d in days_to_check:
                photos = get_machine_photos(machine_id, current_year_month, d)
                for f_name, f_bytes in photos:
                    zip_file.writestr(f"Day_{d}/{f_name}", f_bytes)
                    has_file = True
                    
        if not has_file: return None
        zip_buffer.seek(0)
        gc.collect()
        return zip_buffer
    except Exception as e:
        print(f"Zip Machine Photos Error: {e}")
        return None

def zip_all_factory_photos_by_filter(filter_type="ทั้งโรงงาน", target_date_obj=None, target_day=None):
    if target_date_obj is None: target_date_obj = datetime.date.today()
    current_year_month = target_date_obj.strftime("%Y_%B")
    zip_buffer = BytesIO()
    has_file = False
    
    try:
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            days_to_check = [int(target_day)] if target_day is not None else range(1, 32)
            for machine_code in machine_codes_by_department(filter_type):
                for d in days_to_check:
                    photos = get_machine_photos(machine_code, current_year_month, d)
                    for f_name, f_bytes in photos:
                        zip_file.writestr(f"{machine_code}/Day_{d}/{f_name}", f_bytes)
                        has_file = True
                            
        if not has_file: return None
        zip_buffer.seek(0)
        gc.collect()
        return zip_buffer
    except Exception as e:
        print(f"Zip Filter Photos Error: {e}")
        return None

def delete_photos_by_department_and_period(filter_type, target_date_obj, period_scope):
    """ลบรูปตามแผนกและช่วงเวลา ทั้ง Cloud และ local fallback"""
    machine_codes = machine_codes_by_department(filter_type)
    year_month = target_date_obj.strftime("%Y_%B")
    day_num = target_date_obj.day
    cloud_paths = []

    for machine_code in machine_codes:
        if period_scope == "เฉพาะวันที่เลือก":
            prefix = f"{machine_code}/{year_month}/Day_{day_num}"
            local_target = os.path.join(BASE_FOLDER, "maintenance_photos", machine_code, year_month, f"Day_{day_num}")
        elif period_scope == "ทั้งเดือนที่เลือก":
            prefix = f"{machine_code}/{year_month}"
            local_target = os.path.join(BASE_FOLDER, "maintenance_photos", machine_code, year_month)
        else:  # ทั้งหมดของแผนก
            prefix = machine_code
            local_target = os.path.join(BASE_FOLDER, "maintenance_photos", machine_code)

        if supabase:
            cloud_paths.extend(list_storage_files_recursive(prefix))
        if os.path.exists(local_target):
            shutil.rmtree(local_target)

    if supabase:
        for start in range(0, len(cloud_paths), 100):
            supabase.storage.from_("maintenance-photos").remove(cloud_paths[start:start + 100])
    return len(cloud_paths), len(machine_codes)

def build_factory_issue_print_html(month_logs, year_month_key):
    """สร้างรายงาน HTML สำหรับ Print / Save as PDF แยกปัญหาตามเครื่องจักร"""
    required = {"Note", "Role", "Machine_ID"}
    if month_logs.empty or not required.issubset(month_logs.columns):
        return None

    notes = month_logs[
        month_logs["Note"].notna()
        & month_logs["Note"].astype(str).str.strip().ne("")
        & month_logs["Note"].astype(str).str.strip().str.lower().ne("nan")
        & month_logs["Role"].astype(str).str.strip().str.lower().eq("tech")
    ].copy()
    if notes.empty:
        return None

    notes = notes.sort_values(["Machine_ID", "Day_Num", "Item_No", "Timestamp"])
    sections, total_issues = [], 0
    for machine_code, rows in notes.groupby("Machine_ID", sort=True):
        machine_code = str(machine_code).strip()
        issue_rows = []
        for _, row in rows.iterrows():
            total_issues += 1
            safe = lambda value: html.escape(str(value if pd.notna(value) else ""))
            issue_rows.append(
                f"<tr><td>{safe(row.get('Day_Num', ''))}</td><td>{safe(row.get('Item_No', ''))}</td>"
                f"<td>{safe(row.get('Checklist_Item', ''))}</td><td>{safe(row.get('Status', ''))}</td>"
                f"<td>{safe(row.get('Note', ''))}</td><td>{safe(row.get('Tech_Name', ''))}</td></tr>"
            )
        sections.append(
            "<section class='machine'>"
            f"<h2>{html.escape(machine_code)} - {html.escape(MACHINES.get(machine_code, machine_code))}</h2>"
            f"<div class='count'>จำนวนรายการที่บันทึก: {len(issue_rows)} รายการ</div>"
            "<table><thead><tr><th>วันที่</th><th>ข้อ</th><th>หัวข้อตรวจ</th><th>ผลตรวจ</th>"
            "<th>รายละเอียดปัญหา/การแก้ไข</th><th>ผู้ตรวจ</th></tr></thead>"
            f"<tbody>{''.join(issue_rows)}</tbody></table></section>"
        )

    created_at = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    report_document = f"""<!doctype html><html lang='th'><head><meta charset='utf-8'>
<title>Factory Issues - {html.escape(year_month_key)}</title><style>
@page {{ size:A4 landscape; margin:12mm; }} * {{ box-sizing:border-box; }}
body {{ font-family:Tahoma,'Noto Sans Thai',Arial,sans-serif; color:#172033; margin:0; font-size:10pt; }}
h1 {{ margin:0 0 4px; color:#8b1e2d; font-size:20pt; }} .meta {{ color:#526070; margin-bottom:14px; }}
.summary {{ background:#f7ecee; border-left:5px solid #8b1e2d; padding:8px 12px; margin-bottom:14px; }}
.machine {{ break-inside:avoid; page-break-inside:avoid; margin:0 0 15px; }}
h2 {{ color:white; background:#8b1e2d; padding:7px 10px; margin:0; font-size:13pt; }}
.count {{ border:1px solid #ccd2da; border-bottom:0; padding:5px 8px; font-weight:bold; }}
table {{ width:100%; border-collapse:collapse; table-layout:fixed; }}
th,td {{ border:1px solid #aeb7c2; padding:5px; vertical-align:top; overflow-wrap:anywhere; }}
th {{ background:#e9edf2; text-align:center; }}
th:nth-child(1),td:nth-child(1) {{ width:6%; text-align:center; }} th:nth-child(2),td:nth-child(2) {{ width:5%; text-align:center; }}
th:nth-child(3),td:nth-child(3) {{ width:27%; }} th:nth-child(4),td:nth-child(4) {{ width:14%; }}
th:nth-child(5),td:nth-child(5) {{ width:34%; }} th:nth-child(6),td:nth-child(6) {{ width:14%; }}
.footer {{ margin-top:10px; color:#687386; text-align:right; font-size:8pt; }}
</style></head><body><h1>รายงานรายการปัญหาสะสมของเครื่องจักรทั้งโรงงาน</h1>
<div class='meta'>PHOLLAWAT ENGINEERING SUPPLY CO., LTD. | เดือนข้อมูล: {html.escape(year_month_key)} | จัดทำเมื่อ: {created_at}</div>
<div class='summary'>เครื่องที่มีการบันทึกปัญหา {len(sections)} เครื่อง | รวม {total_issues} รายการ</div>
{''.join(sections)}<div class='footer'>Smart Factory PM SYSTEM - FM-MN-07</div></body></html>"""

    document_json = json.dumps(report_document, ensure_ascii=False).replace("</", "<\\/")
    return f"""<button onclick='printIssueReport()' style='width:100%;padding:12px;border:0;border-radius:8px;background:#8b1e2d;color:white;font-size:16px;font-weight:bold;cursor:pointer;'>🖨️ พิมพ์ / บันทึก PDF รายการปัญหาสะสมทั้งโรงงาน</button>
<script>function printIssueReport() {{ const w=window.open('', '_blank'); if(!w){{alert('กรุณาอนุญาต Pop-up ของเว็บไซต์ก่อนสั่งพิมพ์ PDF');return;}} w.document.open();w.document.write({document_json});w.document.close();w.onload=()=>{{w.focus();w.print();}}; }}</script>"""

# --- 3. UI NAVIGATION SIDEBAR & QUERY PARAMETERS ---
st.set_page_config(page_title="Smart Factory PM SYSTEM", page_icon="🔧", layout="wide")

query_params = st.query_params
raw_role = query_params.get("role", "tech")
is_boss_link = str(raw_role).strip().lower() == "boss"

if is_boss_link:
    user_role = "🔐 Engineer/ผู้ตรวจสอบ"
else:
    st.sidebar.title("🏢 เมนูควบคุมโรงงานรวม")
    user_role = st.sidebar.radio("เลือกสิทธิ์การเข้าใช้งานด้านล่าง:", [
        "🔧 ช่างเทคนิค (ส่งฟอร์ม)",
        "🔐 Engineer/ผู้ตรวจสอบ",
        "👑 ผู้บริหารสูงสุด (Big Boss Zone)"
    ])

raw_machine_id = query_params.get("id", "CNC3X-01")
if isinstance(raw_machine_id, list): machine_id = str(raw_machine_id[0]).strip()
else: machine_id = str(raw_machine_id).strip()
machine_id = machine_id.replace("%20", " ")

# รองรับ QR Code เดิมหลังเปลี่ยนทะเบียนรถ โดยแปลงเป็นรหัสปัจจุบันก่อนแสดงผลและบันทึก
MACHINE_ID_ALIASES = {
    "CAR-2ฒถ-5252": "CAR-1ฒถ-5252",
}
machine_id = MACHINE_ID_ALIASES.get(machine_id, machine_id)

m_type_selected = get_machine_type_by_id(machine_id)

# หน้า Engineer/ผู้บริหารต้องใช้ pandas แต่หน้าช่างไม่ต้องโหลดไลบรารีขนาดใหญ่นี้
if user_role != "🔧 ช่างเทคนิค (ส่งฟอร์ม)":
    import pandas as pd

def enable_required_photo_camera_uploads():
    """ขอให้ทุกช่องบังคับถ่ายรูปบนมือถือเปิดกล้องหลังเป็นลำดับแรก"""
    components.html(
        """
        <script>
        (() => {
          const parentDoc = window.parent.document;
          const patchCameraInputs = () => {
            parentDoc.querySelectorAll('input[type="file"]').forEach((input) => {
              input.setAttribute('accept', 'image/*');
              input.setAttribute('capture', 'environment');
            });
          };
          patchCameraInputs();
          const observer = new MutationObserver(patchCameraInputs);
          observer.observe(parentDoc.body, {childList: true, subtree: true});
        })();
        </script>
        """,
        height=0,
        width=0,
    )

# ==========================================
# 🔧 [โหมดที่ 1: ฝั่งช่างเทคนิคส่งฟอร์มประจำวัน]
# ==========================================
if user_role == "🔧 ช่างเทคนิค (ส่งฟอร์ม)":
    if os.path.exists("Logo_Pes.png"): st.image("Logo_Pes.png", width=240)
    st.caption("PHOLLAWAT ENGINEERING SUPPLY CO., LTD.")

    if PHOTO_RULES.get(m_type_selected, []):
        enable_required_photo_camera_uploads()

    st.title(f"📋 ใบตรวจสอบเครื่อง {machine_id} ประจำวัน")
    st.info("📄 มาตรฐานระบบคุณภาพโรงงาน: **FM-MN-07 Rev.00**")

    if machine_id in MACHINES: st.success(f"⚙️ คุณกำลังตรวจเครื่อง: **{machine_id} ({MACHINES[machine_id]})**")
    else: st.error(f"⚠️ ไม่พบรหัสเครื่อง '{machine_id}' ในทะเบียนกลาง")
    st.divider()

    report_date = st.date_input("📆 เลือกวันที่ตรวจสอบงานฟอร์ม:", value=datetime.date.today())
    current_day = report_date.day
    year_month_key = report_date.strftime("%Y_%B")

    with st.form("pm_form"):
        tech_name = st.text_input("👤 ชื่อช่างผู้ตรวจเช็ค (ผู้รับผิดชอบ)", placeholder="ระบุชื่อ-นามสกุลของคุณ")
        results, uploaded_photos = {}, {}
        current_checklist = CHECKLISTS.get(m_type_selected, CHECKLISTS["CNC"])
        required_photo_indexes = PHOTO_RULES.get(m_type_selected, [])

        if required_photo_indexes:
            st.info("📷 หัวข้อที่มีสัญลักษณ์กล้องต้องถ่ายรูปปัจจุบันอย่างน้อย 1 รูปก่อนส่ง สามารถถ่ายมากกว่า 1 รูปต่อหัวข้อได้")
        
        for i, item in enumerate(current_checklist, 1):
            st.write(f"**{i}. {item}**")
            status = st.radio(f"ผลการตรวจข้อ {i}", ["ใช้งานได้ปกติ", "ทำการแก้ไขใช้งานได้ปกติ", "ใช้งานไม่ได้ต้องแก้ไข", "ไม่ได้ทำงาน"], horizontal=True, key=f"check_{i}", label_visibility="collapsed", index=None)
            if i in required_photo_indexes:
                st.write("📷 *บังคับถ่ายรูปปัจจุบันหัวข้อนี้ก่อนส่งรายงาน*")
                uploaded_files = st.file_uploader(
                    f"📷 ถ่ายรูปข้อ {i} (เลือกได้มากกว่า 1 รูป)",
                    type=["jpg", "jpeg", "png"],
                    key=f"required_camera_{m_type_selected}_{i}",
                    accept_multiple_files=True,
                )
                uploaded_photos[i] = {"files": uploaded_files, "index": i}
            note = st.text_input(f"หมายเหตุ/อาการเสีย (ข้อ {i})", key=f"note_{i}", placeholder="ระบุรายละเอียดหากพบจุดพังหรือบันทึกงานซ่อมแก้ไข")
            results[item] = {"status": status, "note": note}
            st.divider()

        submitted = st.form_submit_button("💾 ส่งรายงานการตรวจเช็คประจำวัน (SUBMIT)")

    if submitted:
        if machine_id not in MACHINES: st.error("❌ รหัสเครื่องจักรไม่ถูกต้อง")
        elif not tech_name: st.error("❌ กรุณาระบุชื่อผู้ตรวจสอบก่อนส่งรายงานครับ!")
        elif any(results[item]["status"] is None for item in current_checklist): st.error("❌ ปฏิเสธการบันทึก! ช่างยังเลือกผลการตรวจสอบไม่ครบทุกหัวข้อ")
        elif any((uploaded_photos[idx]["files"] is None or len(uploaded_photos[idx]["files"]) == 0) for idx in required_photo_indexes): st.error(f"❌ ปฏิเสธการบันทึกฟอร์ม! กรุณาถ่ายภาพหลักฐานประจำข้อ {required_photo_indexes} ให้ครบถ้วนก่อนกดส่งครับ")
        else:
            save_uploaded_photos_dict(machine_id, current_day, uploaded_photos, current_date_obj=report_date)

            logs_to_save = []
            for i, item in enumerate(current_checklist, 1):
                logs_to_save.append({
                    "machine_id": machine_id,
                    "day_num": int(current_day),
                    "year_month": year_month_key,
                    "tech_name": tech_name,
                    "item_no": int(i),
                    "checklist_item": item,
                    "status": str(results[item]["status"]).strip(),
                    "note": str(results[item]["note"]).strip(),
                    "role": "tech"
                })
            if not save_log_to_supabase_bulk(logs_to_save):
                st.error("❌ บันทึกฐานข้อมูลไม่สำเร็จ กรุณาตรวจสอบการเชื่อมต่อแล้วลองใหม่")
                st.stop()

            fails, fixed_items = [], []
            for i, item in enumerate(current_checklist, 1):
                status_val = str(results[item]["status"]).strip()
                note_val = str(results[item]["note"]).strip()
                if "ไม่ได้" in status_val or "ต้องแก้ไข" in status_val: fails.append(f"- ข้อ {i}. {item}" + (f" ({note_val})" if note_val else ""))
                elif "ทำการแก้ไข" in status_val: fixed_items.append(f"- ข้อ {i}. {item}" + (f" ({note_val})" if note_val else ""))
            
            boss_review_url = f"{DEFAULT_APP_URL}/?role=boss&id={quote(machine_id, safe='')}"
            audit_tag = f"\n\n📂 [คลิกเปิดตรวจรายงานและดูภาพหลักฐานคลาวด์]:\n👉 {boss_review_url}"
            
            if fails:
                summary_msg = f"\n🚨 [แจ้งซ่อมด่วนจากใบตรวจเช็ค ISO]\n🔧 เครื่อง: {MACHINES[machine_id]}\n📅 วันที่: {current_time_str}\n👤 ผู้ตรวจ: {tech_name}\n\n❌ รายการที่ไม่ผ่านมาตรฐาน:\n" + "\n".join(fails)
                if fixed_items: summary_msg += "\n\n🛠️ รายการที่ช่างแก้ไขเสร็จทันที:\n" + "\n".join(fixed_items)
                send_line_alert(summary_msg + audit_tag)
                st.warning("พบจุดบกพร่อง! ส่งการแจ้งเตือนเตือนเข้าไลน์กลุ่มช่างแล้ว")
            else:
                ok_msg = f"\n🎉 [รายงานเครื่องจักรปกติ - ISO]\n🔧 เครื่อง: {MACHINES[machine_id]}\n📅 วันที่: {current_time_str}\n✅ ผลการตรวจสอบ: ปกติทุกหัวข้อ\n👤 ผู้ตรวจสอบ: {tech_name}"
                if fixed_items: ok_msg += "\n\n🛠️ รายการที่ช่างแก้ไขหน้างานสำเร็จ (ลงตาราง ⨂):\n" + "\n".join(fixed_items)
                send_line_alert(ok_msg + audit_tag)
            st.success(f"🎉 บันทึกรายงานเครื่อง {machine_id} สำเร็จ! ข้อมูลรอยติ๊กและรูปภาพอัปเดตเรียบร้อยแล้ว")

# ==========================================
# 🔐 [โหมดที่ 2: ฝั่งหัวหน้างาน ดูบอร์ดตรวจเช็ค/กดอนุมัติ]
# ==========================================
elif user_role == "🔐 Engineer/ผู้ตรวจสอบ":
    if os.path.exists("Logo_Pes.png"): st.image("Logo_Pes.png", width=240)
    st.caption("PHOLLAWAT ENGINEERING SUPPLY CO., LTD.")
    st.title("🔐 หน้าต่างควบคุมระบบตรวจสอบคุณภาพ (สำหรับ Engineer)")
    
    col_date, col_refresh = st.columns([3, 1])
    with col_date:
        selected_date = st.date_input("📆 เลือกวันที่ต้องการตรวจสอบเอกสารและดูรูปภาพย้อนหลัง:", value=datetime.date.today())
    with col_refresh:
        st.write("")
        st.write("")
        if st.button("🔄 รีเฟรชข้อมูลล่าสุด", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    target_day_check = selected_date.day
    year_month_key = selected_date.strftime("%Y_%B")
    
    st.subheader(f"📅 ประจำวันที่เลือก: {selected_date.strftime('%d/%m/%Y')} (คอลัมน์ Excel ช่องวันที่ {target_day_check})")
    
    password_input = st.text_input("🔑 กรุณากรอกรหัสผ่านเพื่อเข้าสู่ระบบบอร์ดควบคุม Engineer:", type="password")
    if not BOSS_PASSWORD and not BIGBOSS_PASSWORD:
        st.error("❌ ยังไม่ได้ตั้งค่า BOSS_PASSWORD และ BIGBOSS_PASSWORD ใน Streamlit Secrets")
    
    if password_input != "":
        if password_input == BOSS_PASSWORD or password_input == BIGBOSS_PASSWORD:
            st.success("🔓 ยืนยันสิทธิ์: เข้าสู่ระบบตรวจสอบและบันทึกประจำวันได้")
            boss_name = st.text_input("👤 ชื่อผู้ตรวจสอบ/Engineer (บังคับระบุชื่อเพื่อกดอนุมัติ):", value="", placeholder="ระบุชื่อ-นามสกุลของคุณเพื่อลงนามอนุมัติ")
            
            st.divider()
            st.write("### 📊 บอร์ดควบคุมการรายงานตรวจเช็ค ทั้งโรงงาน")
       
            month_logs_all = fetch_month_logs(year_month_key)
            if not month_logs_all.empty and "Day_Num" in month_logs_all.columns:
                day_logs_all = month_logs_all[month_logs_all["Day_Num"].eq(int(target_day_check))]
            else:
                day_logs_all = pd.DataFrame()

            st.write("#### 🖨️ รายงานรายการปัญหาสะสมทั้งโรงงาน")
            issue_print_html = build_factory_issue_print_html(month_logs_all, year_month_key)
            if issue_print_html:
                st.caption("รายงานจะแยกหัวข้อของแต่ละเครื่อง และใช้ข้อมูลปัญหาที่มีการบันทึกหมายเหตุในเดือนที่เลือก")
                components.html(issue_print_html, height=64)
            else:
                st.info("เดือนที่เลือกยังไม่มีรายการปัญหาหรือหมายเหตุสำหรับจัดทำรายงาน PDF")
            st.divider()

            # สร้างดัชนีครั้งเดียว แล้วให้การ์ดทุกเครื่องใช้ข้อมูลในหน่วยความจำร่วมกัน
            # แทนการกรอง DataFrame ทั้งก้อนและยิง Supabase ซ้ำประมาณ 50 รอบ
            day_logs_by_machine = {}
            if not day_logs_all.empty and "Machine_ID" in day_logs_all.columns:
                day_index = day_logs_all["Machine_ID"].astype(str).str.strip().str.upper()
                for normalized_id, indexes in day_logs_all.groupby(day_index).groups.items():
                    day_logs_by_machine[normalized_id] = day_logs_all.loc[indexes]

            month_logs_by_machine = {}
            if not month_logs_all.empty and "Machine_ID" in month_logs_all.columns:
                month_index = month_logs_all["Machine_ID"].astype(str).str.strip().str.upper()
                for normalized_id, indexes in month_logs_all.groupby(month_index).groups.items():
                    month_logs_by_machine[normalized_id] = month_logs_all.loc[indexes]

            @st.fragment
            def render_machine_card(m_id, m_name, m_type_flag):
                approve_state_key = f"approved_{m_id}_{year_month_key}_{target_day_check}"
                
                is_reported = False
                is_approved = st.session_state.get(approve_state_key, False)
                tech_who_checked = ""
                boss_who_approved = st.session_state.get(f"boss_name_{approve_state_key}", "")
                
                target_mid_clean = str(m_id).strip().upper()
                
                df_day = day_logs_by_machine.get(target_mid_clean, pd.DataFrame())
                if not df_day.empty:
                    boss_rows = df_day[df_day["Role"].astype(str).str.strip().str.lower() == "boss"]
                    tech_rows = df_day[df_day["Role"].astype(str).str.strip().str.lower() == "tech"]
                    if not boss_rows.empty:
                        is_approved = True
                        boss_who_approved = str(boss_rows.iloc[0]["Tech_Name"])
                    if not tech_rows.empty:
                        is_reported = True
                        tech_who_checked = str(tech_rows.iloc[0]["Tech_Name"])

                st.info(f"⚙️ **{m_id}**\n{m_name}")
                
                if is_approved:
                    st.success(f"✅ อนุมัติแล้ว โดย: {boss_who_approved}")
                elif is_reported:
                    st.warning(f"📋 ช่างตรวจแล้ว ({tech_who_checked}) - รอหัวหน้าอนุมัติ")
                else:
                    st.caption("⚪ ยังไม่ได้ดำเนินการ")

                if is_approved:
                    st.button(f"🔒 อนุมัติแล้วโดย {boss_who_approved}", key=f"btn_approved_disabled_{m_id}", disabled=True)
                elif not boss_name.strip():
                    st.button(f"⚠️ กรุณาระบุชื่อผู้ตรวจสอบก่อนกดอนุมัติ ({m_id})", key=f"btn_disabled_{m_id}", disabled=True)
                else:
                    if st.button(f"✅ อนุมัติฟอร์มของ {m_id}", key=f"btn_{m_id}"):
                        approval_saved = save_log_to_supabase_bulk([{
                            "machine_id": m_id,
                            "day_num": int(target_day_check),
                            "year_month": year_month_key,
                            "tech_name": boss_name,
                            "item_no": 0,
                            "checklist_item": "BOSS APPROVAL",
                            "status": "APPROVED",
                            "note": "",
                            "role": "boss"
                        }])
                        if approval_saved:
                            st.session_state[approve_state_key] = True
                            st.session_state[f"boss_name_{approve_state_key}"] = boss_name
                            st.toast(f"ลงนามดิจิทัลเครื่อง {m_id} สำเร็จ!", icon="🔥")
                            send_line_alert(f"🔒 [ISO Approved]: หัวหน้างาน/Engineer ({boss_name}) ได้อนุมัติใบตรวจประจำวันที่ {target_day_check} ของเครื่อง {m_id} แล้ว")
                            st.rerun(scope="fragment")
                        else:
                            st.error("❌ บันทึกการอนุมัติไม่สำเร็จ กรุณาลองใหม่")
                
                # ⚡ แสดงรูปภาพหลักฐาน พร้อมชื่อหัวข้อข้อตรวจจริงใต้ภาพ (Caption)
                with st.expander(f"📸 ตรวจรูปภาพหลักฐานวันที่ {target_day_check}"):
                    load_photos = st.toggle("โหลดและแสดงรูปภาพ", key=f"load_photos_{m_id}_{year_month_key}_{target_day_check}")
                    if load_photos:
                        from PIL import Image
                        photos_list = get_machine_photos(m_id, year_month_key, target_day_check)
                        if photos_list:
                            machine_checklist = CHECKLISTS.get(m_type_flag, CHECKLISTS["CNC"])
                            for f_name, f_bytes in photos_list:
                                try:
                                    img_obj = Image.open(BytesIO(f_bytes))
                                    caption_title = f_name
                                    if "item_" in f_name:
                                        try:
                                            item_num = int(f_name.split("item_")[1].split("_")[0])
                                            if 1 <= item_num <= len(machine_checklist):
                                                item_desc = machine_checklist[item_num - 1]
                                                caption_title = f"📌 [ข้อ {item_num}] {item_desc}"
                                        except Exception:
                                            pass
                                    st.image(img_obj, caption=caption_title, use_container_width=True)
                                except Exception:
                                    st.warning(f"ไฟล์ภาพ {f_name} ไม่สามารถแสดงได้")
                        else:
                            st.caption(f"ℹ️ วันที่ {target_day_check} ไม่มีรูปภาพหลักฐาน")
                    else:
                        st.caption("เปิดสวิตช์เมื่อต้องการดูรูป ระบบจะไม่ดาวน์โหลดรูปโดยอัตโนมัติ")

                # ดึง Note อาการเสียสะสมของเครื่องนี้
                df_machine_logs = month_logs_by_machine.get(target_mid_clean, pd.DataFrame())
                current_notes_list = []
                if not df_machine_logs.empty:
                    notes_rows = df_machine_logs[(df_machine_logs["Note"].notnull()) & (df_machine_logs["Note"] != "") & (df_machine_logs["Note"] != "nan")]
                    for _, n_row in notes_rows.iterrows():
                        current_notes_list.append(f"[วันที่ {n_row['Day_Num']}]: ข้อ {n_row['Item_No']} {n_row['Note']}")
                
                current_notes = ", ".join(current_notes_list) if current_notes_list else ""
                st.text_area(f"📝 รายการอาการเสียสะสม ({m_id})", value=current_notes, key=f"note_area_{m_id}", height=100, disabled=True)

                st.write("---")
                excel_col, zip_day_col, zip_month_col = st.columns(3)
                
                with excel_col:
                    with st.popover("📥 ดึง Excel", use_container_width=True):
                        st.write(f"**ดาวน์โหลดแบบฟอร์ม {m_id}**")
                        if st.button(f"สร้างไฟล์ Excel ({m_id})", key=f"btn_gen_excel_{m_id}"):
                            excel_bytes = generate_excel_bytes(m_id, year_month_key, m_type_flag)
                            if excel_bytes:
                                st.download_button(
                                    label=f"💾 กดบันทึกไฟล์ {m_id}.xlsx", 
                                    data=excel_bytes, 
                                    file_name=f"FM-MN-07_{m_id}.xlsx", 
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                                    key=f"dl_confirm_{m_id}",
                                    use_container_width=True
                                )
                            else:
                                st.error("ไม่พบไฟล์แบบฟอร์ม")
                        
                with zip_day_col:
                    with st.popover(f"📸 รูปวันที่ {target_day_check}", use_container_width=True):
                        st.write(f"**รูปภาพประจำวันที่ {target_day_check}**")
                        if st.button(f"สร้างไฟล์ Zip วันนี้ ({m_id})", key=f"btn_gen_zip_day_{m_id}"):
                            zip_day_data = zip_single_machine_photos(m_id, target_date_obj=selected_date, target_day=target_day_check)
                            if zip_day_data:
                                st.download_button(
                                    label=f"💾 กดบันทึกรูป Day {target_day_check}", 
                                    data=zip_day_data, 
                                    file_name=f"Photos_{m_id}_Day_{target_day_check}.zip", 
                                    mime="application/zip", 
                                    key=f"zip_day_confirm_{m_id}",
                                    use_container_width=True
                                )
                            else:
                                st.caption("วันนี้ไม่มีรูปภาพหลักฐาน")
                        
                with zip_month_col:
                    with st.popover("📦 รูปทั้งเดือน", use_container_width=True):
                        st.write(f"**รูปภาพรวมทั้งเดือน {selected_date.strftime('%Y_%B')}**")
                        if st.button(f"สร้างไฟล์ Zip ทั้งเดือน ({m_id})", key=f"btn_gen_zip_month_{m_id}"):
                            zip_month_data = zip_single_machine_photos(m_id, target_date_obj=selected_date, target_day=None)
                            if zip_month_data:
                                st.download_button(
                                    label="💾 กดบันทึกรูปทั้งเดือน", 
                                    data=zip_month_data, 
                                    file_name=f"Photos_{m_id}_Full_{selected_date.strftime('%Y_%B')}.zip", 
                                    mime="application/zip", 
                                    key=f"zip_month_confirm_{m_id}",
                                    use_container_width=True
                                )
                            else:
                                st.caption("เดือนนี้ไม่มีรูปภาพหลักฐาน")
                st.divider()

            # ---- 1. แผนก CNC ----
            st.write("#### 🔹 เครื่อง CNC (9 เครื่อง)")
            cnc_col1, cnc_col2, cnc_col3 = st.columns(3)
            cnc_idx = 0
            for m_id, m_name in MACHINES.items():
                if "CNC" in m_id and "CRANE" not in m_id.upper() and "QC-" not in m_id.upper():
                    with (cnc_col1 if cnc_idx % 3 == 0 else (cnc_col2 if cnc_idx % 3 == 1 else cnc_col3)):
                        render_machine_card(m_id, m_name, "CNC")
                    cnc_idx += 1

            # ---- 2. แผนก GRINDING ----
            st.write("#### 🔹 เครื่องเจียรผิว GRINDING (2 เครื่อง)")
            grind_col1, grind_col2 = st.columns(2)
            grind_idx = 0
            for m_id, m_name in MACHINES.items():
                if "GRINDING" in m_id and "CUTTER" not in m_id:
                    with (grind_col1 if grind_idx % 2 == 0 else grind_col2):
                        render_machine_card(m_id, m_name, "GRINDING")
                    grind_idx += 1

            # ---- 3. แผนก CUTTER GRINDING ----
            st.write("#### 🔹 เครื่องลับคม CUTTER GRINDING (1 เครื่อง)")
            cutter_grind_col1, = st.columns(1)
            with cutter_grind_col1: render_machine_card("CUTTER GRINDING-01", MACHINES["CUTTER GRINDING-01"], "CUTTER GRINDING-01")

            # ---- 4. แผนกปั๊มลม COMPRESSOR ----
            st.write("#### 🔹 เครื่องปั๊มลม AIR COMPRESSOR (2 เครื่อง)")
            comp_col1, comp_col2, comp_col3 = st.columns(3)
            comp_idx = 0
            for m_id, m_name in MACHINES.items():
                if "COMP-" in m_id.upper():
                    with (comp_col1 if comp_idx % 3 == 0 else (comp_col2 if comp_idx % 3 == 1 else comp_col3)):
                        render_machine_card(m_id, m_name, m_id) 
                    comp_idx += 1

            # ---- 5. แผนก CRANE ----
            st.write("#### 🔹 เครน CRANE (2 แผนก)")
            crane_col1, crane_col2 = st.columns(2)
            crane_idx = 0
            for m_id, m_name in MACHINES.items():
                if "CRANE" in m_id.upper() or "Crane" in m_id:
                    with (crane_col1 if crane_idx % 2 == 0 else crane_col2):
                        render_machine_card(m_id, m_name, "Crane no.1" if "no.1" in m_id else "Crane no.2") 
                    crane_idx += 1

            # ---- 6. แผนก QC ----
            st.write("#### 🔹 เครื่องมือวัดคุณภาพ QC (19 เครื่องมือวัด, 2 เครื่องจักรทำงาน)")
            qc_col1, qc_col2, qc_col3 = st.columns(3)
            qc_idx = 0
            for m_id, m_name in MACHINES.items():
                if "QC-" in m_id.upper():
                    with (qc_col1 if qc_idx % 3 == 0 else (qc_col2 if qc_idx % 3 == 1 else qc_col3)):
                        render_machine_card(m_id, m_name, m_id) 
                    qc_idx += 1

            # ---- 7. แผนก MILLING ----
            st.write("#### 🔹 เครื่องมิลลิ่ง MILLING (4 เครื่อง)")
            mill_col1, mill_col2, mill_col3 = st.columns(3)
            mill_idx = 0
            for m_id, m_name in MACHINES.items():
                if "MILLING" in m_id:
                    with (mill_col1 if mill_idx % 3 == 0 else (mill_idx % 3 == 1 and mill_col2 or mill_col3)):
                        render_machine_card(m_id, m_name, "MILLING")
                    mill_idx += 1

            # ---- 8. แผนก LATHE ----
            st.write("#### 🔹 เครื่องกลึง LATHE (1 เครื่อง)")
            lathe_col1, = st.columns(1)
            with lathe_col1: render_machine_card("LATHE-01", MACHINES["LATHE-01"], "LATHE")

            # ---- 9. แผนก CUTTING ----
            st.write("#### 🔹 เครื่องตัด CUTTING (1 เครื่อง)")
            cut_col1, = st.columns(1)
            with cut_col1: render_machine_card("CUTTING-01", MACHINES["CUTTING-01"], "CUTTING")

            # ---- 11. แผนก MIG CO2 ----
            st.write("#### 🔹 เครื่องเชื่อม MIG CO2 (3 เครื่อง)")
            mig_col1, mig_col2, mig_col3 = st.columns(3)
            mig_idx = 0
            for m_id, m_name in MACHINES.items():
                if "MIG" in m_id:
                    with (mig_col1 if mig_idx % 3 == 0 else (mig_col2 if mig_idx % 3 == 1 else mig_col3)):
                        render_machine_card(m_id, m_name, "MIG CO2")
                    mig_idx += 1

            # ---- 12. แผนก ARGON ----
            st.write("#### 🔹 เครื่องเชื่อม ARGON (2 เครื่อง)")
            argon_col1, argon_col2 = st.columns(2)
            argon_idx = 0
            for m_id, m_name in MACHINES.items():
                if "ARGON" in m_id:
                    with (argon_col1 if argon_idx % 2 == 0 else argon_col2):
                        render_machine_card(m_id, m_name, "ARGON")
                    argon_idx += 1

            # ---- 13. แผนก WELDING ALUMINUM ----
            st.write("#### 🔹 เครื่องเชื่อมอลูมิเนียม WELDING ALUMINUM (1 เครื่อง)")
            wel_al_col1, = st.columns(1)
            with wel_al_col1: render_machine_card("WELDING_ALUMINUM-01", MACHINES["WELDING_ALUMINUM-01"], "WELDING_ALUMINUM")

            # ---- 14. แผนก BAND SAW ----
            st.write("#### 🔹 เครื่องเลื่อยสายพาน BAND SAW (3 เครื่อง)")
            saw_col1, saw_col2, saw_col3 = st.columns(3)
            saw_idx = 0
            for m_id, m_name in MACHINES.items():
                if "BAND" in m_id.upper():
                    with (saw_col1 if saw_idx % 3 == 0 else (saw_idx % 3 == 1 and saw_col2 or saw_col3)):
                        render_machine_card(m_id, m_name, "BAND SAW")
                    saw_idx += 1

            # ---- 15. แผนก FORKLIFT ----
            st.write("#### 🔹 รถโฟคลิฟ FORKLIFT (1 เครื่อง)")
            fork_col1, = st.columns(1)
            with fork_col1: render_machine_card("FORKLIFT-01", MACHINES["FORKLIFT-01"], "FORKLIFT")

            # ---- 16. รถยนต์และรถบรรทุก ----
            st.write("#### 🚗 รถยนต์และรถบรรทุก (5 คัน)")
            vehicle_cols = st.columns(3)
            vehicle_order = [
                "CAR-2ฒข-5050",
                "CAR-2ฒฆ-5151",
                "CAR-1ฒถ-5252",
                "CAR-2ฒข-5353",
                "Truck-83-2329",
            ]
            for vehicle_idx, m_id in enumerate(vehicle_order):
                with vehicle_cols[vehicle_idx % 3]:
                    render_machine_card(m_id, MACHINES[m_id], "VEHICLE")
        else:
            st.error("❌ รหัสผ่านไม่ถูกต้อง ไม่พบสิทธิ์เข้าใช้งานระบบตามรหัสนี้ครับ")

# ==========================================
# 👑 [โหมดที่ 3: 👑 พื้นที่ควบคุมผู้บริหารสูงสุด (Big Boss Zone)]
# ==========================================
else:
    if os.path.exists("Logo_Pes.png"): st.image("Logo_Pes.png", width=240)
    st.caption("PHOLLAWAT ENGINEERING SUPPLY CO., LTD.")
    st.title("👑 ศูนย์ควบคุมระบบผู้บริหารสูงสุด (Big Boss Zone)")
    st.info("🔐 พื้นที่ความปลอดภัยระดับสูง สำหรับดาวน์โหลดไฟล์สำรอง พิมพ์คิวอาร์โค้ด และจัดการฐานข้อมูลหลัก")
    
    bigboss_code_input = st.text_input("🔑 กรุณากรอกรหัสผ่านผู้บริหารสูงสุด เพื่อปลดล็อกศูนย์ควบคุม:", type="password", key="bigboss_outside_secret_key")
    if not BIGBOSS_PASSWORD:
        st.error("❌ ยังไม่ได้ตั้งค่า BIGBOSS_PASSWORD ใน Streamlit Secrets")
    
    if bigboss_code_input != "":
        if bigboss_code_input == BIGBOSS_PASSWORD:
            st.success("🎯 ยืนยันสิทธิ์ผู้บริหารสูงสุด สำเร็จ ปลดล็อกเรียบร้อยแล้วครับ!")
            st.divider()
            
            selected_date = st.date_input("📆 เลือกวันที่สำหรับอ้างอิงการดาวน์โหลดข้อมูลย้อนหลัง:", value=datetime.date.today())
            current_boss_month = selected_date.strftime("%Y_%B")
            
            with st.expander("📊 [เฉพาะผู้บริหารสูงสุด] ดาวน์โหลดไฟล์ Excel รวมทุกเครื่องจักรทั้งโรงงาน (.zip)"):
                excel_period = st.radio(
                    "เลือกช่วงข้อมูล Excel:",
                    ["เฉพาะวันที่เลือก", "ทั้งเดือนที่เลือก"],
                    horizontal=True,
                    key="bigboss_excel_period"
                )
                excel_target_day = selected_date.day if excel_period == "เฉพาะวันที่เลือก" else None
                excel_period_label = selected_date.strftime("Day_%d_%Y_%m") if excel_target_day else current_boss_month
                st.info(f"📂 ระบบจะรวมแบบฟอร์ม Excel (FM-MN-07) ของทุกเครื่อง สำหรับ **{excel_period_label}** เป็นไฟล์ .zip")
                if st.button(f"📦 สร้าง Excel ทุกเครื่อง - {excel_period}", type="primary"):
                    with st.spinner("กำลังประกอบไฟล์ Excel ทุกเครื่องจักร กรุณารอสักครู่..."):
                        excel_all_zip = zip_all_factory_excel(current_boss_month, target_day=excel_target_day)
                        if excel_all_zip:
                            st.download_button(
                                label=f"💾 ดาวน์โหลด Excel ทั้งโรงงาน - {excel_period_label}",
                                data=excel_all_zip,
                                file_name=f"Excel_All_Machines_{excel_period_label}.zip",
                                mime="application/zip"
                            )
                        else:
                            st.error("ไม่สามารถสร้างไฟล์ Zip ได้ หรือไม่พบไฟล์เทมเพลต Excel")

            with st.expander("📦 [เฉพาะผู้บริหารสูงสุด] ดาวน์โหลดไฟล์ดิบฐานข้อมูลหลัก (SUPABASE BACKUP)"):
                st.info("📂 ปุ่มนี้ทำหน้าที่ดึงประวัติข้อมูลใน Supabase ออกมาเป็นไฟล์ .csv (จำกัด 3,000 แถวล่าสุด เพื่อความเสถียร)")
                if supabase:
                    if st.button("📥 เตรียมไฟล์สำรองข้อมูล (CSV)", type="primary"):
                        try:
                            res = supabase.table("maintenance_logs").select("*").order("timestamp", desc=True).limit(3000).execute()
                            if res.data:
                                df_db = pd.DataFrame(res.data)
                                csv_data = df_db.to_csv(index=False, encoding="utf-8-sig")
                                st.download_button(
                                    label="💾 ยืนยันดาวน์โหลดไฟล์ CSV",
                                    data=csv_data,
                                    file_name=f"Backup_Master_Database_{datetime.datetime.now().strftime('%Y_%m_%d')}.csv",
                                    mime="text/csv"
                                )
                                gc.collect()
                            else: st.caption("ℹ️ ยังไม่มีข้อมูลบันทึกในฐานข้อมูล Supabase")
                        except Exception as e_db: st.error(f"Error: {e_db}")

            with st.expander("📸 [เฉพาะผู้บริหารสูงสุด] ดาวน์โหลดรูปภาพ PM รวมหมดทั้งโรงงาน (.zip)"):
                dept_target = st.selectbox("เลือกแผนกที่ต้องการดาวน์โหลดรูปภาพ:", PHOTO_DEPARTMENTS, key="photo_download_department")
                photo_period = st.radio(
                    "เลือกช่วงรูปภาพ:", ["เฉพาะวันที่เลือก", "ทั้งเดือนที่เลือก"],
                    horizontal=True, key="photo_download_period"
                )
                photo_target_day = selected_date.day if photo_period == "เฉพาะวันที่เลือก" else None
                photo_period_label = selected_date.strftime("Day_%d_%Y_%m") if photo_target_day else current_boss_month
                st.info(f"📦 รวมรูปแผนก **{dept_target}** ช่วง **{photo_period_label}** เป็นไฟล์ .zip")
                if st.button(f"📦 สร้าง Zip รูป [{dept_target}] - {photo_period}", type="primary"):
                    filtered_zip_data = zip_all_factory_photos_by_filter(
                        filter_type=dept_target, target_date_obj=selected_date, target_day=photo_target_day
                    )
                    if filtered_zip_data:
                        st.download_button(
                            label=f"💾 ดาวน์โหลดรูป [{dept_target}] - {photo_period_label}",
                            data=filtered_zip_data, 
                            file_name=f"Photos_Filter_{dept_target}_{photo_period_label}.zip",
                            mime="application/zip"
                        )
                    else:
                        st.warning(f"⚠️ ไม่พบรูปภาพของแผนก [{dept_target}] ในช่วงที่เลือก")

            with st.expander("🖨️ [เฉพาะผู้บริหารสูงสุด] เครื่องมือพิมพ์ QR Code สำหรับไปแปะหน้าเครื่องจักร"):
                sel_m = st.selectbox("เลือกเครื่องที่ต้องการพิมพ์ QR:", list(MACHINES.keys()), key="bigboss_qr_select_box_outside")
                
                base_web_url = st.text_input(
                    "🔗 ตรวจสอบ Base URL ของระบบ (ต้องขึ้นต้นด้วย https://):", 
                    value=DEFAULT_APP_URL
                )
                
                import qrcode
                qr_url = f"{base_web_url.rstrip('/')}/?id={quote(sel_m, safe='')}" 
                qr = qrcode.make(qr_url)
                buf = BytesIO()
                qr.save(buf)
                st.image(buf, caption=f"QR สำหรับแปะหน้าเครื่อง {MACHINES[sel_m]} (ลิงก์: {qr_url})")
                buf.close()
                gc.collect()

            # ----------------------------------------------------
            # 🧹 กล่องเครื่องมือล้างระบบ แยก 2 ฟังก์ชั่นชัดเจน
            # ----------------------------------------------------
            with st.expander("🧹 [เฉพาะผู้บริหารสูงสุด] กล่องเครื่องมือล้างระบบ (SYSTEM RESET ZONE)"):
                st.warning("⚠️ โซนควบคุมความปลอดภัย: กรุณาตรวจสอบก่อนกดใช้งาน การลบข้อมูลจะไม่สามารถย้อนคืนได้")
                
                col_reset_photos, col_reset_db = st.columns(2)
                
                with col_reset_photos:
                    st.write("#### 📸 1. ลบเฉพาะระบบรูปภาพ")
                    st.caption("เลือกแผนกและช่วงเวลาที่ต้องการลบ โดยไม่กระทบประวัติตารางตรวจ")
                    reset_photo_dept = st.selectbox(
                        "เลือกแผนกที่จะลบรูป:", PHOTO_DEPARTMENTS,
                        key="reset_photo_department"
                    )
                    reset_photo_period = st.radio(
                        "เลือกช่วงรูปที่จะลบ:",
                        ["เฉพาะวันที่เลือก", "ทั้งเดือนที่เลือก", "ทั้งหมดของแผนก"],
                        key="reset_photo_period"
                    )
                    if reset_photo_period == "เฉพาะวันที่เลือก":
                        reset_scope_label = selected_date.strftime("วันที่ %d/%m/%Y")
                    elif reset_photo_period == "ทั้งเดือนที่เลือก":
                        reset_scope_label = f"เดือน {current_boss_month}"
                    else:
                        reset_scope_label = "ทุกวันและทุกเดือน"

                    st.warning(f"กำลังเลือก: แผนก [{reset_photo_dept}] | {reset_scope_label}")
                    confirm_delete_photos = st.checkbox(
                        f"ยืนยันลบรูปแผนก [{reset_photo_dept}] - {reset_scope_label}",
                        key="confirm_delete_photos_filtered"
                    )
                    if st.button(
                        "🗑️ ลบรูปตามตัวเลือก",
                        type="primary",
                        key="btn_reset_filtered_photos",
                        disabled=(not confirm_delete_photos or not supabase)
                    ):
                        try:
                            deleted_count, affected_machines = delete_photos_by_department_and_period(
                                reset_photo_dept, selected_date, reset_photo_period
                            )
                            gc.collect()
                            st.success(
                                f"✅ ลบรูปสำเร็จ {deleted_count} ไฟล์ "
                                f"จากขอบเขต {affected_machines} เครื่อง"
                            )
                            st.toast("ลบรูปตามตัวเลือกสำเร็จ", icon="📸")
                        except Exception as e_st_del:
                            st.error(f"ลบรูปไม่สำเร็จ: {e_st_del}")

                with col_reset_db:
                    st.write("#### 🗄️ 2. ลบเฉพาะประวัติตารางข้อมูล")
                    st.caption("ทำหน้าที่ล้าง Log การตรวจเช็คใน Supabase Database (ไม่ลบไฟล์รูปภาพหลักฐาน)")
                    confirm_delete_db = st.checkbox("ยืนยันว่าต้องการล้างประวัติทั้งหมด", key="confirm_delete_db")
                    if st.button("🧹 สั่งล้างฐานข้อมูลประวัติ", type="primary", key="btn_reset_only_db", disabled=not confirm_delete_db):
                        if supabase:
                            try: 
                                supabase.table("maintenance_logs").delete().neq("id", 0).execute()
                                st.cache_data.clear()
                                gc.collect()
                                st.success("✅ ล้างประวัติตารางข้อมูลใน Supabase เรียบร้อยแล้ว!")
                                st.balloons()
                            except Exception as e_del: 
                                st.error(f"เกิดข้อผิดพลาดในการลบ Database: {e_del}")
                        else:
                            st.error("❌ ไม่สามารถเชื่อมต่อ Supabase ได้")
        else:
            st.error("❌ รหัสผ่านไม่ถูกต้อง ไม่พบสิทธิ์เข้าใช้งานระบบตามรหัสนี้ครับ")
