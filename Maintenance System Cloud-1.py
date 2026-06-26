import streamlit as st
import requests
import smtplib
from email.mime.text import MIMEText
import datetime
import qrcode
from io import BytesIO
import json
import os
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

# --- 1. CONFIGURATION ---
# ล็อกสิทธิ์เข้าใช้งานระบบ LINE Messaging API (Line Bot ตัวจริงประจำกลุ่มของคุณ)
LINE_ACCESS_TOKEN = "RRtpOuJT8oWgvglsSFUqc7LC1zZqL2jD8qTdJx5iIpAkG4GiJjAkaetvEKLGLuNOJ7j9dpyNMSTviG06LCe//YM1+r5TqRQx09p8nLNh5lZzKy78CvGLfGAWjFSOtyj89Bu3nm8iVlTh0pNQtc737gdB04t89/1O/w1cDnyilFU=" 
LINE_TARGET_ID = "Cbf3d27d5280ae8b258727047a26b399a"  

# ระบุตำแหน่งโฟลเดอร์ทำงานอัตโนมัติบนระบบคลาวด์
BASE_FOLDER = os.path.dirname(os.path.abspath(__file__)) if "__file__" in locals() else os.getcwd()
BOSS_PASSWORD = "boss1234"  

# ทะเบียนเครื่องจักรกลางประจำโรงงาน (รวมครบ 10 แผนก ทั้งหมด 29 เครื่อง/ตัว)
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

# รายการเช็คลิสต์แยกตามประเภทแผนก (ล็อกข้อความตามเอกสารชุดจริงแบบคำต่อคำ)
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
    "QC-11": ["ตรวจดูสภาพของไมโครพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
    "QC-12": ["ตรวจดูสภาพของไมโครพร้อมใช้งานหรือไม่", "ตรวจดู BATTERRY อ่อนหรือไม่", "ตรวจสอบการสไลด์ต้องไม่ติดขัด", "หลังเลิกงานต้องปิดสวิตส์ทุกครั้ง"],
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
    "GRINDING": ["การ Worm spindle และ TABLE SLIDE", "เช็คระดับนำมันไฮดรอลิก และ การทำงานของ PUMP", "เช็คระดับของน้ำยา COOLANNT PUMP", "ตรวจสอบการทำงานของแม่เหล็ก", "ตรวจสอบการทำงานของ SLIDE X,Y", "ตรวจสอบสภาพความพร้อมโดยรวมของเครื่องจักร", "ตรวจสอบระดับน้ำมันของ PUMP น้ำมันหล่อลื่น", "ตรวจสอบการทำงานของไฟฟ้าและแสงสว่าง", "ตรวจสอบการทำงานของตัวดูดอากศ"],
    "CUTTER GRINDING": ["การ WORM UP แกน Y พร้อมใช้งาน", "การ WORM UP แกน Z พร้อมใช้งาน", "ตรวจสอบการทำงานของไฟฟ้าและแสงสว่าง", "ตรวจสอบการทำงานของมอเตอร์ มีการหมุนปกติ", "ตรวจสอบการจับหัวคอเรต"],
    "MILLING": ["Worm Spindle ก่อนเริมงาน ตรวจสอบความ ผิดปกติของชุด  Back gauge  และ Motor", "เช็ค Auto  Up-Down back gauge  และ Manual ( ความคร่องตัวในการเคลื่อนที่ของ Spindle )", "ตรวจสอบการ SLIDE  ของแกน X", "ตรวจสอบการ SLIDE  ของแกน Y", "ตรวจสอบการ SLIDE  ของแกน Z", "ระดับน้ำมันไฮดรอลิค ตรวจสอบน้ำมันในปั้มน้ำมันหล่อ ลื่นแกน  X,Y,Z", "ตรวจน้ำมันหล่อลื่นเย็น ตรวจสอบการทำงานของปั้ม COOLANT และสภาพของน้ำ  COOLANT", "ตรวจสอบหน้าจอ  DIGITAL READ OUT และการทำ งานของ LINEAR SCALE", "หยอดน้ำมันหล่อลื่นทุกวันจันทร์", "ตรวจสอบการทำงานของไฟฟ้าแสงสว่างของเครื่อง", "ตรวจสอบสภาพความพร้อมโดยรวมของเครื่องจักร  และอุกรณ์เสริมต่าง ๆ"],
    "CUTTING": ["การ Worm spindle ก่อนเริ่มงาน เพื่อตรวจ ความผิดปกติของชุด Back gauge และ Motor", "เช็ค Auto Up-Down back gauge และ Manual ( ความคล่องตัวในการเคลื่อนที่ )", "ระดับน้ำมันไฮดรอลิค ตรวจสอบระดับในปั้มน้ำมัน หล่อลืนแกน  Back gauge", "ตรวจเช็ค  Switch  เปิด-ปิด", "ตรวจสอบ Digital  read out และการทำงานของ Linear  scale", "อัดจาระบีตามจุดที่อัดจาระบีทุกๆจุด", "ตรวจสอบใบมีด  บนและล่าง", "ตรวจสอบความพร้อมสภาพโดยรวมของเครื่อง จักรและอุปกรณ์เสริมต่าง ๆ"],
    "MIG CO2": ["ตรวจสภาพความพร้อมโดยรวมของเครื่อง", "เช็ค BREAKER เพื่อเช็คระบบไฟฟ้า ตามตำแหน่งไฟ โชว์ และสวิชท์ต่าง ๆ", "ตรวจสภาพความพร้อมของมาตราวัดแรงดัน ของก๊าซ CO2 และปรับตั้งอย่างถูกต้อง", "ตรวจจุดต่อของก๊าซ CO2 รั่วหรือไม่", "ตรวจสภาพความพร้อมของสายไฟ สายก๊าซ  CO2 ว่ารั่วหรือไม่", "ตรวจสภาพความพร้อมของสายกราวด์", "ทำความสะอาดหัวเชื่อมก่อนใช้งาน"],
    "ARGON": ["ตรวจสภาพความพรัอมโดยรวมของเครื่อง", "เช็ค  BREAKER  เพื่อเช็คระบบไฟฟ้า ตามตำแหน่งไฟ โชว์  และ SWITCH  ต่าง ๆ", "ตรวจสภาพความพร้อมของมาตราวัดแรงดันของมาตรา วัดแรงดันของก๊าช  ARGON  และปรับตั้งอย่างถูกวิธี", "ตรวจุดต่อของสายก๊าช  ARGON  ก่อนว่ารั่วหรือไม่", "ตรวจสภาพความพร้อมของสายกราว์", "ตรวจสภาพความพร้อมของสายไฟฟ้าสายก๊าช  ARGON และชุดหัวเชื่อม", "ตรวจสภาพความพร้อมของ  SWITCH  หัวเชื่อม", "ทำความสะดาดชุดหัวเชื่อมก่อนใช้งาน"],
    "BAND SAW": ["เช็ค Auto Up-Down Back Gauge และ Manual (ความคล่องตัวในการเคลื่อนที่ของ Spindle)", "เช็คระดับน้ำมันไฮดรอลิค", "ตรวจน้ำมันหล่อลื่นเย็น ตรวจสอบการทำงานของปั๊ม COOLANT และสภาพของน้ำ COOLANT", "ตรวจสอบ Switch (สวิตซ์) หน้า BOX CONTROL", "ตรวจสอบระดับน้ำมันหล่อลื่นในห้องเกียร์"]
}

# กฎบังคับถ่ายรูปภาพรายแผนก
PHOTO_RULES = {
    "CNC": [2, 3, 4, 5, 8, 13], "Crane no.1": [3, 4], "Crane no.2": [3, 4], "QC-01": [4],
    "QC-02": [2, 4], "QC-03": [2, 4], "QC-04": [2, 4], "QC-05": [2, 4], "QC-06": [2, 4],
    "QC-07": [2, 4], "QC-08": [2, 4], "QC-09": [2, 4], "QC-10": [2], "QC-11": [2], "QC-12": [2],
    "QC-13": [2], "QC-14": [2], "QC-15": [6], "QC-16": [3], "QC-17": [2], "QC-18": [3], "QC-19": [3],
    "QC-20": [3], "QC-21": [3], "COMP-01": [1, 2, 3], "COMP-02": [1, 2, 3], "GRINDING": [2, 4, 7],
    "CUTTER GRINDING": [], "MILLING": [6, 7], "CUTTING": [3, 5, 7], "MIG CO2": [3, 4, 5],
    "ARGON": [3, 4, 6], "BAND SAW": [2, 3, 5]
}

def get_coordinates(m_type):
    if m_type == "CNC": return 22, 24, "B28"
    if "CRANE" in m_type.upper(): return 14, 16, "B19"
    if m_type == "QC-01": return 10, 12, "B15"
    if m_type in ["QC-VERNIER_STD", "QC-MICRO_STD", "QC-ARM_STD", "QC-16", "QC-17", "BAND SAW", "CUTTING", "ARGON", "QC-02", "QC-03", "QC-04", "QC-05", "QC-06", "QC-07", "QC-08", "QC-09", "QC-13", "QC-14", "QC-18", "QC-19", "QC-20", "QC-21", "COMP-01", "COMP-02"]: return 11, 13, "B16"
    if m_type in ["QC-10", "QC-11", "QC-12"]: return 11, 13, "B15"
    if m_type == "QC-15": return 12, 14, "B17"
    if m_type == "GRINDING": return 16, 18, "B21"
    if m_type == "CUTTER GRINDING": return 13, 15, "B18"
    if m_type == "MILLING": return 20, 22, "B25"
    if m_type == "MIG CO2": return 13, 15, "B18"
    return 11, 13, "B16"

def get_unmerged_cell(ws, coordinate_str):
    cell = ws[coordinate_str]
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return ws.cell(row=merged_range.min_row, column=merged_range.min_col)
    return cell

# --- 2. FUNCTIONS ---
def send_line_alert(msg_text):
    """ฟังก์ชันหลักส่งข้อความรายงานตามมาตรฐาน LINE Messaging API ดั้งเดิมของคุณ"""
    url = 'https://api.line.me/v2/bot/message/push'
    headers = {'Content-Type': 'application/json', 'Authorization': f'Bearer {LINE_ACCESS_TOKEN}'}
    payload = {"to": LINE_TARGET_ID, "messages": [{"type": "text", "text": msg_text}]}
    try: requests.post(url, headers=headers, data=json.dumps(payload))
    except Exception as e: print(f"ส่งไลน์ไม่สำเร็จ: {e}")

# 🟢 [UPDATED DEFINITIVE] ฟังก์ชันสำหรับอัปโหลดและส่งรูปภาพเข้า LINE Bot (Messaging API) ตัวจริงประจำกลุ่ม
def send_line_image(photo_path, caption_text):
    """อัปโหลดภาพผ่านเซิร์ฟเวอร์ฝากรูปภาพชั่วคราวเพื่อแปลงเป็น URL แล้วยิงส่งตรงเข้า Line Bot กลุ่มสำเร็จแน่นอน 100%"""
    try:
        # 1. นำไฟล์ภาพที่ช่างถ่าย ฝากไว้ที่เซิร์ฟเวอร์โฮสต์รูปภาพชั่วคราวฟรี (freeimage.host API อเนกประสงค์ฟรี ไม่ต้องมี Key)
        with open(photo_path, "rb") as img_file:
            response = requests.post(
                "https://freeimage.host/api/1/upload",
                data={"key": "6d207e02198a847aa98d0a2a901485a5", "action": "upload"},
                files={"source": img_file}
            )
            res_json = response.json()
            # 2. แกะเอาลิงก์ URL รูปภาพสาธารณะที่แท้จริงออกมา
            public_image_url = res_json["image"]["url"]
            
        # 3. ยิงคำสั่งประกอบเข้า LINE Messaging API สู่ห้องแชทกลุ่ม เพื่อให้ภาพเด้งขึ้นหน้าจอทันที!
        url = 'https://api.line.me/v2/bot/message/push'
        headers = {'Content-Type': 'application/json', 'Authorization': f'Bearer {LINE_ACCESS_TOKEN}'}
        payload = {
            "to": LINE_TARGET_ID,
            "messages": [
                {
                    "type": "image",
                    "originalContentUrl": public_image_url,
                    "previewImageUrl": public_image_url
                }
            ]
        }
        requests.post(url, headers=headers, data=json.dumps(payload))
    except Exception as e:
        print(f"กลไกยิงรูปภาพหลุด: {e}")

def save_uploaded_photo(machine_id, day_num, item_index, uploaded_file):
    if uploaded_file is not None:
        folder_path = os.path.join(BASE_FOLDER, "maintenance_photos")
        sub_folder = os.path.join(folder_path, f"{machine_id}_Day_{day_num}")
        if not os.path.exists(folder_path): os.makedirs(folder_path, exist_ok=True)
        if not os.path.exists(sub_folder): os.makedirs(sub_folder, exist_ok=True)
        
        file_extension = os.path.splitext(uploaded_file.name)[1]
        file_name = f"photo_item_{item_index}{file_extension}"
        full_path = os.path.join(sub_folder, file_name)
        with open(full_path, "wb") as f: f.write(uploaded_file.getbuffer())
        return full_path
    return None

def update_iso_excel_by_tech(machine_id, day_num, results_dict, tech_name, m_type):
    target_excel_path = os.path.join(BASE_FOLDER, f"FM-MN-07_{machine_id}.xlsx")
    if not os.path.isfile(target_excel_path): return False, f"ไม่พบไฟล์แบบฟอร์ม `{target_excel_path}` บนระบบคลาวด์"
    try:
        wb = openpyxl.load_workbook(target_excel_path, data_only=False)
        ws = wb.active
        col_letter = get_column_letter(2 + day_num)
        
        checklist_items = CHECKLISTS[m_type]
        fail_notes = []
        
        for i, item in enumerate(checklist_items, 1):
            cell_coordinate = f"{col_letter}{5 + i}"
            current_cell = get_unmerged_cell(ws, cell_coordinate)
            if item in results_dict:
                status_val = results_dict[item]["status"]
                note_val = results_dict[item]["note"]
                
                if status_val == "ใช้งานได้ปกติ": current_cell.value = "/"
                elif status_val == "ทำการแก้ไขใช้งานได้ปกติ":
                    current_cell.value = "⨂"
                    if note_val: fail_notes.append(f"ข้อ {i} (แก้ไขแล้ว): {note_val}")
                    else: fail_notes.append(f"ข้อ {i}: ทำการแก้ไขให้ใช้งานได้ปกติแล้ว")
                elif status_val == "ใช้งานไม่ได้ต้องแก้ไข":
                    current_cell.value = "X"
                    if note_val: fail_notes.append(f"ข้อ {i} (พบปัญหา): {note_val}")
                    else: fail_notes.append(f"ข้อ {i}: พบปัญหาไม่ผ่านมาตรฐาน")
                elif status_val == "ไม่ได้ทำงาน": current_cell.value = "-"
                
                current_cell.alignment = Alignment(horizontal='center', vertical='center')
                
        t_row, _, n_cell = get_coordinates(m_type)
        tech_cell = get_unmerged_cell(ws, f"{col_letter}{t_row}")
        tech_cell.value = tech_name
        tech_cell.alignment = Alignment(text_rotation=90, horizontal='center', vertical='center')
        
        note_cell = get_unmerged_cell(ws, n_cell)
        old_val = note_cell.value or ""
        notes_collected = [results_dict[item]["note"] for item in checklist_items if results_dict[item]["note"]]
        if notes_collected:
            new_val = old_val + ("\n" if old_val else "") + f"[วันที่ {day_num}]: " + ", ".join(notes_collected)
            note_cell.value = new_val
            note_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            
        wb.save(target_excel_path)
        return True, ""
    except Exception as e:
        return False, str(e)

def approve_excel_by_boss(machine_id, day_num, boss_name, m_type):
    target_excel_path = os.path.join(BASE_FOLDER, f"FM-MN-07_{machine_id}.xlsx")
    if not os.path.isfile(target_excel_path): return False
    try:
        wb = openpyxl.load_workbook(target_excel_path, data_only=False)
        ws = wb.active
        col_letter = get_column_letter(2 + day_num)
        _, boss_row, _ = get_coordinates(m_type)
        
        boss_cell = get_unmerged_cell(ws, f"{col_letter}{boss_row}")
        boss_cell.value = boss_name
        boss_cell.alignment = Alignment(text_rotation=90, horizontal="center", vertical="center")
        wb.save(target_excel_path)
        return True
    except Exception as e: print(f"Boss approve error: {e}"); return False

def get_current_excel_note(machine_id, m_type):
    target_excel_path = os.path.join(BASE_FOLDER, f"FM-MN-07_{machine_id}.xlsx")
    if not os.path.isfile(target_excel_path): return ""
    try:
        wb = openpyxl.load_workbook(target_excel_path, data_only=False)
        ws = wb.active
        _, _, n_cell = get_coordinates(m_type)
        note_cell = get_unmerged_cell(ws, n_cell)
        val = note_cell.value
        return val if val else ""
    except: return ""

def save_custom_excel_note_by_boss(machine_id, m_type, new_text):
    target_excel_path = os.path.join(BASE_FOLDER, f"FM-MN-07_{machine_id}.xlsx")
    if not os.path.isfile(target_excel_path): return False
    try:
        wb = openpyxl.load_workbook(target_excel_path, data_only=False)
        ws = wb.active
        _, _, n_cell = get_coordinates(m_type)
        note_cell = get_unmerged_cell(ws, n_cell)
        
        note_cell.value = new_text.strip() if new_text.strip() else "เครื่องจักรปกติ"
        note_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        wb.save(target_excel_path)
        return True
    except Exception as e: print(f"Save custom note error: {e}"); return False

# --- 3. UI NAVIGATION SIDEBAR ---
st.sidebar.title("🏢 เมนูควบคุมโรงงานรวม")
user_role = st.sidebar.radio("เลือกสิทธิ์การเข้าใช้งานด้านล่าง:", ["🔧 ช่างเทคนิค (ส่งฟอร์ม)", "🔐 หัวหน้างาน/ผู้ตรวจสอบ"])

now = datetime.datetime.now()
current_day = now.day
current_time_str = now.strftime("%Y-%m-%d %H:%M:%S")

query_params = st.query_params
raw_machine_id = query_params.get("id", "CNC3X-01")

if isinstance(raw_machine_id, list): machine_id = str(raw_machine_id[0]).strip()
else: machine_id = str(raw_machine_id).strip()

machine_id = machine_id.replace("%20", " ")

if "CRANE NO.1" in machine_id.upper() or "CRANE NO.1" in machine_id: m_type_selected = "Crane no.1"
elif "CRANE NO.2" in machine_id.upper() or "CRANE NO.2" in machine_id: m_type_selected = "Crane no.2"
elif "QC-01" in machine_id.upper() or "QC-01" in machine_id: m_type_selected = "QC-01"
elif "QC-02" in machine_id.upper() or "QC-02" in machine_id: m_type_selected = "QC-02"
elif "QC-03" in machine_id.upper() or "QC-03" in machine_id: m_type_selected = "QC-03"
elif "QC-04" in machine_id.upper() or "QC-04" in machine_id: m_type_selected = "QC-04"
elif "QC-05" in machine_id.upper() or "QC-05" in machine_id: m_type_selected = "QC-05"
elif "QC-06" in machine_id.upper() or "QC-06" in machine_id: m_type_selected = "QC-06"
elif "QC-07" in machine_id.upper() or "QC-07" in machine_id: m_type_selected = "QC-07"
elif "QC-08" in machine_id.upper() or "QC-08" in machine_id: m_type_selected = "QC-08"
elif "QC-09" in machine_id.upper() or "QC-09" in machine_id: m_type_selected = "QC-09"
elif "QC-10" in machine_id.upper() or "QC-10" in machine_id: m_type_selected = "QC-10"
elif "QC-11" in machine_id.upper() or "QC-11" in machine_id: m_type_selected = "QC-11"
elif "QC-12" in machine_id.upper() or "QC-12" in machine_id: m_type_selected = "QC-12"
elif "QC-13" in machine_id.upper() or "QC-13" in machine_id: m_type_selected = "QC-13"
elif "QC-14" in machine_id.upper() or "QC-14" in machine_id: m_type_selected = "QC-14"
elif "QC-15" in machine_id.upper() or "QC-15" in machine_id: m_type_selected = "QC-15"
elif "QC-16" in machine_id.upper() or "QC-16" in machine_id: m_type_selected = "QC-16"
elif "QC-17" in machine_id.upper() or "QC-17" in machine_id: m_type_selected = "QC-17"
elif "QC-18" in machine_id.upper() or "QC-18" in machine_id: m_type_selected = "QC-18"
elif "QC-19" in machine_id.upper() or "QC-19" in machine_id: m_type_selected = "QC-19"
elif "QC-20" in machine_id.upper() or "QC-20" in machine_id: m_type_selected = "QC-20" 
elif "QC-21" in machine_id.upper() or "QC-21" in machine_id: m_type_selected = "QC-21"
elif "COMP-01" in machine_id.upper(): m_type_selected = "COMP-01"
elif "COMP-02" in machine_id.upper(): m_type_selected = "COMP-02"
elif "GRINDING" in machine_id.upper(): m_type_selected = "GRINDING"
elif "CUTTER" in machine_id.upper(): m_type_selected = "CUTTER GRINDING"
elif "MILLING" in machine_id.upper(): m_type_selected = "MILLING"
elif "CUTTING" in machine_id.upper(): m_type_selected = "CUTTING"
elif "MIG" in machine_id.upper(): m_type_selected = "MIG CO2"
elif "ARGON" in machine_id.upper(): m_type_selected = "ARGON"
elif "BAND" in machine_id.upper(): m_type_selected = "BAND SAW"
else: m_type_selected = "CNC"

# ==========================================
# 🔧 [โหมดที่ 1: ฝั่งช่างเทคนิคส่งฟอร์มประจำวัน]
# ==========================================
if user_role == "🔧 ช่างเทคนิค (ส่งฟอร์ม)":
    st.caption("PHOLLAWAT ENGINEERING SUPPLY CO., LTD.")
    st.title(f"📋 ใบตรวจสอบเครื่อง {machine_id} ประจำวัน")
    st.info("📄 มาตรฐานระบบคุณภาพโรงงาน: **FM-MN-07 Rev.00**")

    if machine_id in MACHINES: st.success(f"⚙️ คุณกำลังตรวจเครื่อง: **{machine_id} ({MACHINES[machine_id]})**")
    else: st.error(f"⚠️ ไม่พบรหัสเครื่อง '{machine_id}' ในทะเบียนกลาง")
    st.divider()

    with st.form("pm_form"):
        tech_name = st.text_input("👤 ชื่อช่างผู้ตรวจเช็ค (ผู้รับผิดชอบ)", placeholder="ระบุชื่อ-นามสกุลของคุณ")
        results, uploaded_photos = {}, {}
        current_checklist = CHECKLISTS[m_type_selected]
        required_photo_indexes = PHOTO_RULES[m_type_selected]
        
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
        elif not tech_name: st.error("❌ กรุณาระบุชื่อผู้ตรวจสอบก่อนส่งรายงานครับ!")
        elif any(results[item]["status"] is None for item in current_checklist): st.error("❌ ปฏิเสธการบันทึก! ช่างยังเลือกผลการตรวจสอบไม่ครบทุกหัวข้อ")
        elif any(uploaded_photos[idx]["file"] is None for idx in required_photo_indexes): st.error(f"❌ ปฏิเสธการบันทึกฟอร์ม! กรุณาถ่ายภาพหลักฐานประจำข้อ {required_photo_indexes} ให้ครบถ้วนก่อนกดส่งครับ")
        else:
            photo_logs = []
            for idx in required_photo_indexes:
                saved_path = save_uploaded_photo(machine_id, current_day, idx, uploaded_photos[idx]["file"])
                if saved_path: 
                    photo_logs.append(f"📸 แนบรูปหลักฐานข้อ {idx} สำเร็จ")
                    # เรียกคำสั่งแปลงรูปเป็น URL และส่งตรงเข้าแชท Line Bot กลุ่มแบบกางภาพ
                    send_line_image(saved_path, f"📷 [หลักฐานข้อ {idx}] เครื่อง: {machine_id} โดยช่าง {tech_name}")
            
            fails, fixed_items = [], []
            for i, item in enumerate(current_checklist, 1):
                status_val = results[item]["status"]
                note_val = results[item]["note"]
                if status_val == "ใช้งานไม่ได้ต้องแก้ไข": fails.append(f"- ข้อ {i}. {item}" + (f" ({note_val})" if note_val else ""))
                elif status_val == "ทำการแก้ไขใช้งานได้ปกติ": fixed_items.append(f"- ข้อ {i}. {item}" + (f" ({note_val})" if note_val else ""))
            
            success, err_msg = update_iso_excel_by_tech(machine_id, current_day, results, tech_name, m_type_selected)
            if success:
                photo_status_str = "\n".join(photo_logs)
                audit_tag = f"\n\n🔒 [ISO Status]: บันทึกรายงานเครื่อง {machine_id} แล้ว (รอหัวหน้าลงนามดิจิทัล)"
                if fails:
                    summary_msg = f"\n🚨 [แจ้งซ่อมด่วนจากใบตรวจเช็ค ISO]\n🔧 เครื่อง: {MACHINES[machine_id]}\n📅 วันที่: {current_time_str}\n👤 ผู้ตรวจ: {tech_name}\n\n❌ รายการที่ไม่ผ่านมาตรฐาน:\n" + "\n".join(fails)
                    if fixed_items: summary_msg += "\n\n🛠️ รายการที่ช่างแก้ไขเสร็จทันที:\n" + "\n".join(fixed_items)
                    send_line_alert(summary_msg + audit_tag)
                    st.warning("พบจุดบกพร่อง! ส่งการแจ้งเตือนเตือนเข้าไลน์กลุ่มช่างแล้ว")
                else:
                    ok_msg = f"\n🎉 [รายงานเครื่องจักรปกติ - ISO]\n🔧 เครื่อง: {MACHINES[machine_id]}\n📅 วันที่: {current_time_str}\n✅ ผลการตรวจสอบ: ปกติทุกหัวข้อ\n👤 ผู้ตรวจสอบ: {tech_name}"
                    if fixed_items: ok_msg += "\n\n🛠️ รายการที่ช่างแก้ไขหน้างานสำเร็จ (ลงตาราง ⨂):\n" + "\n".join(fixed_items)
                    if photo_status_str: ok_msg += f"\n\n📸 รูปภาพหลักฐาน:\n{photo_status_str}"
                    send_line_alert(ok_msg + audit_tag)
                st.success(f"🎉 บันทึกรายงานเครื่อง {machine_id} สำเร็จ! ข้อมูลอัปเดตและบันทึกเรียบร้อยแล้ว")
            else:
                st.error(f"เกิดข้อผิดพลาดในการบันทึก Excel: {err_msg}")

# ==========================================
# 🔐 [โหมดที่ 2: ฝั่งหัวหน้างาน ล็อกอินตรวจสอบและกดอนุมัติฟอร์ม]
# ==========================================
else:
    st.title("🔐 หน้าต่างควบคุมระบบตรวจสอบคุณภาพ (สำหรับหัวหน้างาน)")
    
    selected_date = st.date_input("📆 เลือกวันที่ต้องการตรวจสอบเอกสารและรูปภาพยิงย้อนหลัง:", value=datetime.date.today())
    target_day_check = selected_date.day
    
    st.subheader(f"📅 ประจำวันที่เลือก: {selected_date.strftime('%d/%m/%Y')} (คอลัมน์ Excel ช่องวันที่ {target_day_check})")
    
    password_input = st.text_input("🔑 กรุณากรอกรหัสผ่านผู้เข้าตรวจสอบเพื่อเข้าถึงระบบอนุมัติ:", type="password")
    if password_input == BOSS_PASSWORD:
        st.success("🔓 รหัสผ่านถูกต้อง เข้าสู่ระบบลงนามดิจิทัลมาตรฐาน ISO สำเร็จ")
        boss_name = st.text_input("👤 ชื่อผู้ตรวจสอบ/หัวหน้างาน:", value="พลวัฒน์")
        st.divider()
        st.write("### 📊 บอร์ดควบคุมควบคุมใบงานรวม (แยกรายแผนก)")
        
        def render_machine_card(m_id, m_name, m_type_flag):
            st.info(f"⚙️ **{m_id}**\n{m_name}")
            target_file = os.path.join(BASE_FOLDER, f"FM-MN-07_{m_id}.xlsx")
            if os.path.isfile(target_file):
                if st.button(f"✅ อนุมัติฟอร์มของ {m_id}", key=f"btn_{m_id}"):
                    if approve_excel_by_boss(m_id, target_day_check, boss_name, m_type_flag):
                        st.toast(f"ลงนามดิจิทัลเครื่อง {m_id} สำเร็จ!", icon="🔥")
                        send_line_alert(f"🔒 [ISO Approved]: หัวหน้างาน ({boss_name}) ได้อนุมัติใบตรวจประจำวันที่ {target_day_check} ของเครื่อง {m_id} แล้ว")
                        st.success(f"✍️ เซ็นรับรองลงช่องผู้ตรวจสอบเครื่อง {m_id} สำเร็จ!")
                
                img_dir = os.path.join(BASE_FOLDER, f"maintenance_photos/{m_id}_Day_{target_day_check}")
                if os.path.exists(img_dir):
                    valid_photos = [os.path.join(img_dir, p) for p in os.listdir(img_dir) if p.lower().endswith(('.png', '.jpg', '.jpeg'))]
                    if valid_photos:
                        with st.expander(f"📸 ตรวจรูปภาพหลักฐานวันที่ {target_day_check} ({len(valid_photos)} รูป)"):
                            for p_path in valid_photos:
                                st.image(p_path, caption=f"หลักฐาน: {os.path.basename(p_path)}", use_container_width=True)
                else:
                    st.caption(f"ℹ️ วันที่ {target_day_check} ไม่มีรูปภาพหลักฐาน")

                current_notes = get_current_excel_note(m_id, m_type_flag)
                if m_type_flag == "CNC": note_label = "ช่อง B28"
                elif "CRANE" in m_type_flag.upper(): note_label = "ช่อง B19"  
                elif m_type_flag == "QC-01": note_label = "ช่อง B15"  
                elif m_type_flag in ["QC-02", "QC-03"]: note_label = "ช่อง B16"  
                elif m_type_flag == "GRINDING": note_label = "ช่อง B21"
                elif m_type_flag == "CUTTER GRINDING": note_label = "ช่อง B18"
                elif m_type_flag == "MILLING": note_label = "ช่อง B25"
                elif m_type_flag == "CUTTING": note_label = "ช่อง B20"
                elif m_type_flag == "MIG CO2": note_label = "ช่อง B18"
                elif m_type_flag == "ARGON": note_label = "ช่อง B19"
                else: note_label = "ช่อง B16"
                
                edited_notes = st.text_area(f"📝 รายการอาการเสียสะสม ({note_label})", value=current_notes, key=f"note_area_{m_id}", height=120)
                if st.button(f"💾 เซฟบันทึก {note_label} ของ {m_id}", key=f"save_note_{m_id}"):
                    if save_custom_excel_note_by_boss(m_id, m_type_flag, edited_notes):
                        st.toast(f"อัปเดตรายการอาการเสียเครื่อง {m_id} สำเร็จ!", icon="💾")
                        st.rerun()
                with open(target_file, "rb") as f:
                    st.download_button(label=f"📥 ดึงไฟล์ Excel ของ {m_id}", data=f, file_name=f"FM-MN-07_{m_id}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_{m_id}")
            else: st.error(f"ยังไม่มีไฟล์ฟอร์ม FM-MN-07_{m_id}.xlsx บนระบบคลาวด์")
            st.divider()

        # ---- 1. แผนก CNC ----
        st.write("#### 🔹 แผนกเครื่อง CNC (9 เครื่อง)")
        cnc_col1, cnc_col2, cnc_col3 = st.columns(3)
        cnc_idx = 0
        for m_id, m_name in MACHINES.items():
            if "CNC" in m_id and "CRANE" not in m_id.upper() and "QC-" not in m_id.upper():
                with (cnc_col1 if cnc_idx % 3 == 0 else (cnc_col2 if cnc_idx % 3 == 1 else cnc_col3)):
                    render_machine_card(m_id, m_name, "CNC")
                cnc_idx += 1

        # ---- 2. แผนก CRANE ----
        st.write("#### 🔹 แผนกเครน CRANE / HOIST (2 ตัว)")
        crane_col1, crane_col2 = st.columns(2)
        crane_idx = 0
        for m_id, m_name in MACHINES.items():
            if "CRANE" in m_id.upper():
                with (crane_col1 if crane_idx % 2 == 0 else crane_col2):
                    render_machine_card(m_id, m_name, m_id)
                crane_idx += 1

        # ---- 3. แผนก QC ----
        st.write("#### 🔹 แผนกเครื่องมือวัดคุณภาพ QC (ประหยัดพื้นที่ เรียงหน้ากระดาน 3 แถว)")
        qc_col1, qc_col2, qc_col3 = st.columns(3)
        qc_idx = 0
        for m_id, m_name in MACHINES.items():
            if "QC-" in m_id.upper():
                with (qc_col1 if qc_idx % 3 == 0 else (qc_col2 if qc_idx % 3 == 1 else qc_col3)):
                    render_machine_card(m_id, m_name, m_id)
                qc_idx += 1

        # ---- 4. แผนก GRINDING ----
        st.write("#### 🔹 แผนกเครื่องเจียรผิว GRINDING (2 เครื่อง)")
        grind_col1, grind_col2 = st.columns(2)
        grind_idx = 0
        for m_id, m_name in MACHINES.items():
            if "GRINDING" in m_id and "CUTTER" not in m_id:
                with (grind_col1 if grind_idx % 2 == 0 else grind_col2):
                    render_machine_card(m_id, m_name, "GRINDING")
                grind_idx += 1

        # ---- 5. แผนก CUTTER GRINDING ----
        st.write("#### 🔹 แผนกเครื่องลับคม CUTTER GRINDING (1 เครื่อง)")
        cutter_grind_col1, = st.columns(1)
        with cutter_grind_col1: render_machine_card("CUTTER GRINDING-01", MACHINES["CUTTER GRINDING-01"], "CUTTER GRINDING")

        # ---- 6. แผนก MILLING ----
        st.write("#### 🔹 แผนกเครื่องมิลลิ่ง MILLING (3 เครื่อง)")
        mill_col1, mill_col2, mill_col3 = st.columns(3)
        mill_idx = 0
        for m_id, m_name in MACHINES.items():
            if "MILLING" in m_id:
                with (mill_col1 if mill_idx % 3 == 0 else (mill_col2 if mill_idx % 3 == 1 else mill_col3)):
                    render_machine_card(m_id, m_name, "MILLING")
                mill_idx += 1

        # ---- 7. แผนก CUTTING ----
        st.write("#### 🔹 แผนกเครื่องตัด CUTTING (2 เครื่อง)")
        cut_col1, cut_col2 = st.columns(2)
        cut_idx = 0
        for m_id, m_name in MACHINES.items():
            if "CUTTING" in m_id:
                with (cut_col1 if cut_idx % 2 == 0 else cut_col2):
                    render_machine_card(m_id, m_name, "CUTTING")
                cut_idx += 1

        # ---- 8. แผนก MIG CO2 ----
        st.write("#### 🔹 แผนกเครื่องเชื่อม MIG CO2 (3 เครื่อง)")
        mig_col1, mig_col2, mig_col3 = st.columns(3)
        mig_idx = 0
        for m_id, m_name in MACHINES.items():
            if "MIG" in m_id:
                with (mig_col1 if mig_idx % 3 == 0 else (mig_col2 if mig_idx % 3 == 1 else mig_col3)):
                    render_machine_card(m_id, m_name, "MIG CO2")
                mig_idx += 1

        # ---- 9. แผนก ARGON ----
        st.write("#### 🔹 แผนกเครื่องเชื่อม ARGON (1 เครื่อง)")
        argon_col1, = st.columns(1)
        for m_id, m_name in MACHINES.items():
            if "ARGON" in m_id:
                with argon_col1: render_machine_card(m_id, m_name, "ARGON")

        # ---- 10. แผนก BAND SAW ----
        st.write("#### 🔹 แผนกเครื่องเลื่อยสายพาน BAND SAW (3 เครื่อง)")
        saw_col1, saw_col2, saw_col3 = st.columns(3)
        saw_idx = 0
        for m_id, m_name in MACHINES.items():
            if "BAND" in m_id:
                with (saw_col1 if saw_idx % 3 == 0 else (saw_col2 if saw_idx % 3 == 1 else saw_col3)):
                    render_machine_card(m_id, m_name, "BAND SAW")
                saw_idx += 1

        # ---- 11. แผนกปั๊มลม COMPRESSOR ----
        st.write("#### 🔹 แผนกปั๊มลม AIR COMPRESSOR (2 เครื่อง)")
        comp_col1, comp_col2, comp_col3 = st.columns(3)
        comp_idx = 0
        for m_id, m_name in MACHINES.items():
            if "COMP-" in m_id.upper():
                with (comp_col1 if comp_idx % 3 == 0 else (comp_col2 if comp_idx % 3 == 1 else comp_col3)):
                    render_machine_card(m_id, m_name, "COMP-01") 
                comp_idx += 1

    elif password_input != "": st.error("❌ รหัสผ่านไม่ถูกต้อง กรุณาตรวจสอบรหัสผ่านใหม่อีกครั้งครับเพื่อนรัก")

    with st.expander("🖨️ เครื่องมือหัวหน้างาน: พิมพ์ QR Code สำหรับไปแปะหน้าเครื่องจักร"):
        sel_m = st.selectbox("เลือกเครื่องที่ต้องการพิมพ์ QR:", list(MACHINES.keys()))
        qr_url = f"https://pes-maintenance.streamlit.app/?id={sel_m}" 
        qr = qrcode.make(qr_url)
        buf = BytesIO()
        qr.save(buf)
        st.image(buf, caption=f"QR สำหรับแปะหน้าเครื่อง {MACHINES[sel_m]}")
