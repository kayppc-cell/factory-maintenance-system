import streamlit as st
import requests
import smtplib
from email.mime.text import MIMEText
import datetime
import qrcode
from io import BytesIO
@@ -13,17 +11,14 @@
from openpyxl.styles import Alignment

# --- 1. CONFIGURATION ---
# 🔑 [🔒 UPDATE: สัญญาณ Token ใหม่เอี่ยมอัปเดตเข้าท่อหลักเรียบร้อย]
LINE_ACCESS_TOKEN = "SOs7DeGwVsFpuK/JN8zm58Wn3EOiB75Ww0q57z1/yht4H1imzYonre4QuPfQ3cxbJ7j9dpyNMSTviG06LCe//YM1+r5TqRQx09p8nLNh5lYwCp4biq7N20ffJqzGm+ZYNgtEzt2rYZ/GYVRV725EiAdB04t89/1O/w1cDnyilFU=" 
LINE_TARGET_ID = "Cbf3d27d5280ae8b258727047a26b399a"  

BASE_FOLDER = os.path.dirname(os.path.abspath(__file__)) if "__file__" in locals() else os.getcwd()

# รหัสความปลอดภัยประจำโรงงาน
BOSS_PASSWORD = "boss1234"       
BIGBOSS_PASSWORD = "bigboss9999" 

# ทะเบียนเครื่องจักรกลางประจำโรงงาน
MACHINES = {
    "CNC3X-01": "CNC 3 แกน #01", "CNC3X-02": "CNC 3 แกน #02",
    "CNC3X-03": "CNC 3 แกน #03", "CNC3X-04": "CNC 3 แกน #04",
@@ -67,7 +62,6 @@
    "FORKLIFT-01": "รถโฟคลิฟ FORKLIFT #01"
}

# รายการเช็คลิสต์แยกตามประเภทแผนก
CHECKLISTS = {
    "CNC": [
        "Worm up เครื่องจักร 15 นาที ทุกครั้งที่ใช้งาน", "เช็คระดับนำมัน Oil Matic Mesh ทุกวัน เติมเมื่อพร่อง",
@@ -158,7 +152,6 @@
    ]
}

# กฎบังคับถ่ายรูปภาพรายแผนก
PHOTO_RULES = {
    "CNC": [2, 3, 4, 5, 8, 13], "Crane no.1": [3, 4], "Crane no.2": [3, 4], "QC-01": [4],
    "QC-02": [2, 4], "QC-03": [2, 4], "QC-04": [2, 4], "QC-05": [2, 4], "QC-06": [2, 4],
@@ -200,41 +193,13 @@ def get_unmerged_cell(ws, coordinate_str):
            return ws.cell(row=merged_range.min_row, column=merged_range.min_col)
    return cell

# --- 2. FUNCTIONS ---
def send_line_alert(msg_text):
    url = 'https://api.line.me/v2/bot/message/push'
    headers = {'Content-Type': 'application/json', 'Authorization': f'Bearer {LINE_ACCESS_TOKEN}'}
    payload = {"to": LINE_TARGET_ID, "messages": [{"type": "text", "text": msg_text}]}
    try: requests.post(url, headers=headers, data=json.dumps(payload))
    except Exception as e: print(f"ส่งไลน์ไม่สำเร็จ: {e}")

def send_line_image(photo_path, caption_text):
    try:
        with open(photo_path, "rb") as img_file:
            response = requests.post(
                "https://freeimage.host/api/1/upload",
                data={"key": "6d207e02198a847aa98d0a2a901485a5", "action": "upload"},
                files={"source": img_file}
            )
            res_json = response.json()
            public_image_url = res_json["image"]["url"]
            
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
@@ -257,37 +222,6 @@ def update_iso_excel_by_tech(machine_id, day_num, results_dict, tech_name, m_typ
        ws = wb.active

        t_row, boss_row, n_cell = get_coordinates_by_machine(machine_id, m_type)
        
        check_col_1 = get_column_letter(3) 
        first_cell_of_month = get_unmerged_cell(ws, f"{check_col_1}{t_row}")
        
        if day_num == 1 and (first_cell_of_month.value is None or first_cell_of_month.value == ""):
            backup_folder = os.path.join(BASE_FOLDER, "maintenance_backups")
            if not os.path.exists(backup_folder): os.makedirs(backup_folder, exist_ok=True)
            today = datetime.date.today()
            first_day_of_current_month = today.replace(day=1)
            last_day_of_last_month = first_day_of_current_month - datetime.timedelta(days=1)
            last_month_str = last_day_of_last_month.strftime("%B_%Y")
            
            backup_file_name = f"Backup_{last_month_str}_FM-MN-07_{machine_id}.xlsx"
            backup_excel_path = os.path.join(backup_folder, backup_file_name)
            shutil.copy2(target_excel_path, backup_excel_path)
            
            try: send_line_alert(f"📦 [Auto-Backup Completed]: ระบบสำรองไฟล์ของเครื่อง {machine_id} ประจำเดือน {last_month_str} สำเร็จเรียบร้อยแล้ว!")
            except: pass

            checklist_items = CHECKLISTS[m_type]
            for d in range(1, 32):
                c_letter = get_column_letter(2 + d)
                for row_idx in range(6, 6 + len(checklist_items)):
                    ws.cell(row=row_idx, column=2 + d, value="")
                get_unmerged_cell(ws, f"{c_letter}{t_row}").value = ""
                get_unmerged_cell(ws, f"{c_letter}{boss_row}").value = ""
            
            note_cell = get_unmerged_cell(ws, n_cell)
            note_cell.value = "เครื่องจักรปกติ"
            note_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        col_letter = get_column_letter(2 + day_num)
        checklist_items = CHECKLISTS[m_type]

@@ -296,12 +230,10 @@ def update_iso_excel_by_tech(machine_id, day_num, results_dict, tech_name, m_typ
            current_cell = get_unmerged_cell(ws, cell_coordinate)
            if item in results_dict:
                status_val = results_dict[item]["status"]
                
                if status_val == "ใช้งานได้ปกติ": current_cell.value = "/"
                elif status_val == "ทำการแก้ไขใช้งานได้ปกติ": current_cell.value = "⨂"
                elif status_val == "ใช้งานไม่ได้ต้องแก้ไข": current_cell.value = "X"
                elif status_val == "ไม่ได้ทำงาน": current_cell.value = "-"
                
                current_cell.alignment = Alignment(horizontal='center', vertical='center')

        tech_cell = get_unmerged_cell(ws, f"{col_letter}{t_row}")
@@ -318,8 +250,7 @@ def update_iso_excel_by_tech(machine_id, day_num, results_dict, tech_name, m_typ

        wb.save(target_excel_path)
        return True, ""
    except Exception as e:
        return False, str(e)
    except Exception as e: return False, str(e)

def approve_excel_by_boss(machine_id, day_num, boss_name, m_type):
    target_excel_path = os.path.join(BASE_FOLDER, f"FM-MN-07_{machine_id}.xlsx")
@@ -329,13 +260,12 @@ def approve_excel_by_boss(machine_id, day_num, boss_name, m_type):
        ws = wb.active
        col_letter = get_column_letter(2 + day_num)
        _, boss_row, _ = get_coordinates_by_machine(machine_id, m_type)
        
        boss_cell = get_unmerged_cell(ws, f"{col_letter}{boss_row}")
        boss_cell.value = boss_name
        boss_cell.alignment = Alignment(text_rotation=90, horizontal="center", vertical="center")
        wb.save(target_excel_path)
        return True
    except Exception as e: print(f"Boss approve error: {e}"); return False
    except: return False

def get_current_excel_note(machine_id, m_type):
    target_excel_path = os.path.join(BASE_FOLDER, f"FM-MN-07_{machine_id}.xlsx")
@@ -345,8 +275,7 @@ def get_current_excel_note(machine_id, m_type):
        ws = wb.active
        _, _, n_cell = get_coordinates_by_machine(machine_id, m_type)
        note_cell = get_unmerged_cell(ws, n_cell)
        val = note_cell.value
        return val if val else ""
        return note_cell.value or ""
    except: return ""

def save_custom_excel_note_by_boss(machine_id, m_type, new_text):
@@ -357,14 +286,13 @@ def save_custom_excel_note_by_boss(machine_id, m_type, new_text):
        ws = wb.active
        _, _, n_cell = get_coordinates_by_machine(machine_id, m_type)
        note_cell = get_unmerged_cell(ws, n_cell)
        
        note_cell.value = new_text.strip() if new_text.strip() else "เครื่องจักรปกติ"
        note_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        wb.save(target_excel_path)
        return True
    except Exception as e: print(f"Save custom note error: {e}"); return False
    except: return False

# --- 3. UI NAVIGATION SIDEBAR & QUERY PARAMETERS ---
# --- UI NAVIGATION ---
st.set_page_config(page_title="Smart Factory PM SYSTEM", page_icon="🔧", layout="wide")

query_params = st.query_params
@@ -388,27 +316,8 @@ def save_custom_excel_note_by_boss(machine_id, m_type, new_text):

if "CRANE NO.1" in machine_id.upper() or "CRANE no.1" in machine_id: m_type_selected = "Crane no.1"
elif "CRANE NO.2" in machine_id.upper() or "CRANE no.2" in machine_id: m_type_selected = "Crane no.2"
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
elif "QC-01" in machine_id.upper(): m_type_selected = "QC-01"
elif "QC-" in machine_id.upper(): m_type_selected = "QC-02"
elif "COMP-01" in machine_id.upper(): m_type_selected = "COMP-01"
elif "COMP-02" in machine_id.upper(): m_type_selected = "COMP-02"
elif "GRINDING" in machine_id.upper(): m_type_selected = "GRINDING"
@@ -438,7 +347,7 @@ def save_custom_excel_note_by_boss(machine_id, m_type, new_text):
    st.divider()

    with st.form("pm_form"):
        tech_name = st.text_input("👤 ชื่อช่างผู้ตรวจเช็ค (ผู้รับผิดชอบ)", placeholder="ระบุชื่อ-นามสกุลของคุณ")
        tech_name = st.text_input("👤 ชื่อช่างผู้ตรวจเช็ค", placeholder="ระบุชื่อ-นามสกุลของคุณ")
        results, uploaded_photos = {}, {}
        current_checklist = CHECKLISTS[m_type_selected]
        required_photo_indexes = PHOTO_RULES[m_type_selected]
@@ -447,10 +356,10 @@ def save_custom_excel_note_by_boss(machine_id, m_type, new_text):
            st.write(f"**{i}. {item}**")
            status = st.radio(f"ผลการตรวจข้อ {i}", ["ใช้งานได้ปกติ", "ทำการแก้ไขใช้งานได้ปกติ", "ใช้งานไม่ได้ต้องแก้ไข", "ไม่ได้ทำงาน"], horizontal=True, key=f"check_{i}", label_visibility="collapsed", index=None)
            if i in required_photo_indexes:
                st.write("📷 *หัวข้อบังคับถ่ายรูปหลักฐานยืนยันหน้างานจริง*")
                st.write("📷 *หัวข้อบังคับถ่ายรูปหลักฐานยืนยันหน้างาน*")
                uploaded_file = st.file_uploader(f"แนบรูปข้อ {i}", type=["jpg", "jpeg", "png"], key=f"photo_{i}")
                uploaded_photos[i] = {"file": uploaded_file, "index": i}
            note = st.text_input(f"หมายเหตุ/อาการเสีย (ข้อ {i})", key=f"note_{i}", placeholder="ระบุรายละเอียดหากพบจุดพังหรือบันทึกงานซ่อมแก้ไข")
            note = st.text_input(f"หมายเหตุ (ข้อ {i})", key=f"note_{i}")
            results[item] = {"status": status, "note": note}
            st.divider()

@@ -460,39 +369,36 @@ def save_custom_excel_note_by_boss(machine_id, m_type, new_text):
        if machine_id not in MACHINES: st.error("❌ รหัสเครื่องจักรไม่ถูกต้อง")
        elif not tech_name: st.error("❌ กรุณาระบุชื่อผู้ตรวจสอบก่อนส่งรายงานครับ!")
        elif any(results[item]["status"] is None for item in current_checklist): st.error("❌ ปฏิเสธการบันทึก! ช่างยังเลือกผลการตรวจสอบไม่ครบทุกหัวข้อ")
        elif any(uploaded_photos[idx]["file"] is None for idx in required_photo_indexes): st.error(f"❌ ปฏิเสธการบันทึกฟอร์ม! กรุณาถ่ายภาพหลักฐานประจำข้อ {required_photo_indexes} ให้ครบถ้วนก่อนกดส่งครับ")
        elif any(uploaded_photos[idx]["file"] is None for idx in required_photo_indexes): st.error(f"❌ ปฏิเสธการบันทึกฟอร์ม! กรุณาถ่ายภาพหลักฐานประจำข้อ {required_photo_indexes} ให้ครบถ้วน")
        else:
            photo_logs = []
            # บันทึกรูปภาพเก็บไว้ในคลาวด์ภายในตัวแอปอย่างเดียว (ไม่ส่งไฟล์สดเข้า LINE)
            for idx in required_photo_indexes:
                saved_path = save_uploaded_photo(machine_id, current_day, idx, uploaded_photos[idx]["file"])
                if saved_path: 
                    photo_logs.append(f"📸 แนบรูปหลักฐานข้อ {idx} สำเร็จ")
                    try: send_line_image(saved_path, f"📷 [หลักฐานข้อ {idx}] เครื่อง: {machine_id} โดยช่าง {tech_name}")
                    except: pass
            
            fails, fixed_items = [], []
            for i, item in enumerate(current_checklist, 1):
                status_val = results[item]["status"]
                note_val = results[item]["note"]
                if status_val == "ใช้งานไม่ได้ต้องแก้ไข": fails.append(f"- ข้อ {i}. {item}" + (f" ({note_val})" if note_val else ""))
                elif status_val == "ทำการแก้ไขใช้งานได้ปกติ": fixed_items.append(f"- ข้อ {i}. {item}" + (f" ({note_val})" if note_val else ""))
            
                save_uploaded_photo(machine_id, current_day, idx, uploaded_photos[idx]["file"])
                
            success, err_msg = update_iso_excel_by_tech(machine_id, current_day, results, tech_name, m_type_selected)
            if success:
                photo_status_str = "\n".join(photo_logs)
                audit_tag = f"\n\n🔒 [ISO Status]: บันทึกรายงานเครื่อง {machine_id} แล้ว (รอหัวหน้าลงนามดิจิทัล)"
                fails, fixed_items = [], []
                for i, item in enumerate(current_checklist, 1):
                    status_val = results[item]["status"]
                    note_val = results[item]["note"]
                    if status_val == "ใช้งานไม่ได้ต้องแก้ไข": fails.append(f"- ข้อ {i}. {item}" + (f" ({note_val})" if note_val else ""))
                    elif status_val == "ทำการแก้ไขใช้งานได้ปกติ": fixed_items.append(f"- ข้อ {i}. {item}" + (f" ({note_val})" if note_val else ""))
                
                # 🔗 ไฮไลท์เด็ด: สร้างลิงก์เปิดหน้าตรวจสอบของหัวหน้างานตรงเครื่องนั้น ๆ อัตโนมัติ
                boss_review_url = f"https://pes-maintenance.streamlit.app/?role=boss&id={machine_id}"
                audit_tag = f"\n\n📂 [คลิกเพื่อตรวจเช็คและดูรูปหลักฐานทั้งหมด]:\n👉 {boss_review_url}"

                if fails:
                    summary_msg = f"\n🚨 [แจ้งซ่อมด่วนจากใบตรวจเช็ค ISO]\n🔧 เครื่อง: {MACHINES[machine_id]}\n📅 วันที่: {current_time_str}\n👤 ผู้ตรวจ: {tech_name}\n\n❌ รายการที่ไม่ผ่านมาตรฐาน:\n" + "\n".join(fails)
                    if fixed_items: summary_msg += "\n\n🛠️ รายการที่ช่างแก้ไขเสร็จทันที:\n" + "\n".join(fixed_items)
                    summary_msg = f"\n🚨 [แจ้งซ่อมด่วน - ISO]\n🔧 เครื่อง: {MACHINES[machine_id]}\n📅 วันที่: {current_time_str}\n👤 ผู้ตรวจ: {tech_name}\n\n❌ รายการที่ไม่ผ่าน:\n" + "\n".join(fails)
                    if fixed_items: summary_msg += "\n\n🛠️ รายการที่แก้ไขเสร็จทันที:\n" + "\n".join(fixed_items)
                    send_line_alert(summary_msg + audit_tag)
                    st.warning("พบจุดบกพร่อง! ส่งการแจ้งเตือนเตือนเข้าไลน์กลุ่มช่างแล้ว")
                    st.warning("พบจุดบกพร่อง! แจ้งเตือนเข้าไลน์กลุ่มแล้ว")
                else:
                    ok_msg = f"\n🎉 [รายงานเครื่องจักรปกติ - ISO]\n🔧 เครื่อง: {MACHINES[machine_id]}\n📅 วันที่: {current_time_str}\n✅ ผลการตรวจสอบ: ปกติทุกหัวข้อ\n👤 ผู้ตรวจสอบ: {tech_name}"
                    if fixed_items: ok_msg += "\n\n🛠️ รายการที่ช่างแก้ไขหน้างานสำเร็จ (ลงตาราง ⨂):\n" + "\n".join(fixed_items)
                    if photo_status_str: ok_msg += f"\n\n📸 รูปภาพหลักฐาน:\n{photo_status_str}"
                    ok_msg = f"\n🎉 [รายงานปกติ - ISO]\n🔧 เครื่อง: {MACHINES[machine_id]}\n📅 วันที่: {current_time_str}\n✅ ผลการตรวจสอบ: ปกติทุกหัวข้อ\n👤 ผู้ตรวจสอบ: {tech_name}"
                    if fixed_items: ok_msg += "\n\n🛠️ รายการที่แก้ไขหน้างานสำเร็จ (⨂):\n" + "\n".join(fixed_items)
                    send_line_alert(ok_msg + audit_tag)
                st.success(f"🎉 บันทึกรายงานเครื่อง {machine_id} สำเร็จ! ข้อมูลอัปเดตและบันทึกเรียบร้อยแล้ว")
                
                st.success(f"🎉 บันทึกรายงานเครื่อง {machine_id} สำเร็จ! ข้อมูลอัปเดตและส่งลิงก์แจ้งเตือนเข้าไลน์เรียบร้อย")
            else:
                st.error(f"เกิดข้อผิดพลาดในการบันทึก Excel: {err_msg}")

@@ -507,240 +413,41 @@ def save_custom_excel_note_by_boss(machine_id, m_type, new_text):
    selected_date = st.date_input("📆 เลือกวันที่ต้องการตรวจสอบเอกสารและรูปภาพยิงย้อนหลัง:", value=datetime.date.today())
    target_day_check = selected_date.day

    st.subheader(f"📅 ประจำวันที่เลือก: {selected_date.strftime('%d/%m/%Y')} (คอลัมน์ Excel ช่องวันที่ {target_day_check})")
    
    is_supervisor = False
    is_bigboss = False
    
    password_input = st.text_input("🔑 กรุณากรอกรหัสผ่านหลักเพื่อเข้าสู่ระบบบอร์ดควบคุมหลัก:", type="password")
    
    if password_input != "":
        is_supervisor = (password_input == BOSS_PASSWORD)
        is_bigboss = (password_input == BIGBOSS_PASSWORD)
    is_supervisor = (password_input == BOSS_PASSWORD)
    is_bigboss = (password_input == BIGBOSS_PASSWORD)

    if is_supervisor or is_bigboss:
        if is_bigboss:
            st.success("👑 [สิทธิ์ผู้บริหารสูงสุด]: ล็อกอินผ่านรหัสแอดมินหลักเรียบร้อย")
            boss_name = st.text_input("👤 ชื่อผู้ตรวจสอบ/บิ๊กบอส:", value="พลวัฒน์ (Big Boss)")
        else:
            st.success("🔓 ยืนยันสิทธิ์: เข้าสู่ระบบตรวจสอบและบันทึกประจำวันได้")
            boss_name = st.text_input("👤 ชื่อผู้ตรวจสอบ/หัวหน้างาน:", value="พลวัฒน์")
            
        boss_name = "พลวัฒน์ (Big Boss)" if is_bigboss else "พลวัฒน์"
        st.success(f"🔓 ยืนยันสิทธิ์: ยินดีต้อนรับคุณ {boss_name}")
        st.divider()
        st.write("### 📊 บอร์ดควบคุมการรายงานตรวจเช็ค ทั้งโรงงาน")
   
        
        def render_machine_card(m_id, m_name, m_type_flag):
            st.info(f"⚙️ **{m_id}**\n{m_name}")
            target_file = os.path.join(BASE_FOLDER, f"FM-MN-07_{m_id}.xlsx")
            if os.path.isfile(target_file):
                if st.button(f"✅ อนุมัติฟอร์มของ {m_id}", key=f"btn_{m_id}"):
                    if approve_excel_by_boss(m_id, target_day_check, boss_name, m_type_flag):
                        st.toast(f"ลงนามดิจิทัลเครื่อง {m_id} สำเร็จ!", icon="🔥")
                        send_line_alert(f"🔒 [ISO Approved]: หัวหน้างาน ({boss_name}) ได้อนุมัติใบตรวจประจำวันที่ {target_day_check} ของเครื่อง {m_id} แล้ว")
                        st.success(f"✍️ เซ็นรับรองลงช่องผู้ตรวจสอบเครื่อง {m_id} สำเร็จ!")
                        st.toast(f"ลงนามดิจิทัลเครื่อง {m_id} สำเร็จ!")
                        send_line_alert(f"🔒 [ISO Approved]: หัวหน้างาน ({boss_name}) ได้อนุมัติใบตรวจวันที่ {target_day_check} ของเครื่อง {m_id} เรียบร้อยแล้ว")

                # ดึงรูปภาพที่บันทึกไว้ในตัวแอปคลาวด์มาเปิดให้หัวหน้างานส่องเช็คตรงนี้ได้ทันที!
                img_dir = os.path.join(BASE_FOLDER, f"maintenance_photos/{m_id}_Day_{target_day_check}")
                if os.path.exists(img_dir):
                    valid_photos = [os.path.join(img_dir, p) for p in os.listdir(img_dir) if p.lower().endswith(('.png', '.jpg', '.jpeg'))]
                    if valid_photos:
                        with st.expander(f"📸 ตรวจรูปภาพหลักฐานวันที่ {target_day_check} ({len(valid_photos)} รูป)"):
                        with st.expander(f"📸 ตรวจรูปภาพหลักฐานข้อบังคับ ({len(valid_photos)} รูป)"):
                            for p_path in valid_photos:
                                st.image(p_path, caption=f"หลักฐาน: {os.path.basename(p_path)}", use_container_width=True)
                else:
                    st.caption(f"ℹ️ วันที่ {target_day_check} ไม่มีรูปภาพหลักฐาน")

                current_notes = get_current_excel_note(m_id, m_type_flag)
                
                u_id = str(m_id).upper()
                if "ARGON-02" in u_id or "ARGON-01" in u_id or "CRANE" in u_id: note_label = "ช่อง B19"
                elif "WELDING_ALUMINUM" in u_id or "FORKLIFT" in u_id or "CUTTER" in u_id or "CUTTING" in u_id: note_label = "ช่อง B18"
                elif "CNC" in u_id: note_label = "ช่อง B28"
                elif "QC-01" in u_id or "QC-10" in u_id or "QC-11" in u_id or "QC-12" in u_id: note_label = "ช่อง B15"
                elif "QC-15" in u_id: note_label = "ช่อง B17"
                elif "GRINDING" in u_id: note_label = "ช่อง B21"
                elif "MILLING" in u_id or "LATHE" in u_id: note_label = "ช่อง B22"
                elif "BENDING" in u_id: note_label = "ช่อง B20"
                else: note_label = "ช่อง B16"
                else: st.caption("ℹ️ วันนี้ไม่มีรูปภาพหลักฐาน")

                edited_notes = st.text_area(f"📝 รายการอาการเสียสะสม ({note_label})", value=current_notes, key=f"note_area_{m_id}", height=120)
                if st.button(f"💾 เซฟบันทึก {note_label} ของ {m_id}", key=f"save_note_{m_id}"):
                    if save_custom_excel_note_by_boss(m_id, m_type_flag, edited_notes):
                        st.toast(f"อัปเดตรายการอาการเสียเครื่อง {m_id} สำเร็จ!", icon="💾")
                        st.rerun()
                with open(target_file, "rb") as f:
                    st.download_button(label=f"📥 ดึงไฟล์ Excel ของ {m_id}", data=f, file_name=f"FM-MN-07_{m_id}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_{m_id}")
            else: st.error(f"ยังไม่มีไฟล์ฟอร์ม FM-MN-07_{m_id}.xlsx บนระบบคลาวด์")
            else: st.error("ยังไม่มีไฟล์ฟอร์มบนระบบ")
            st.divider()

        # ---- 1. แผนก CNC ----
        st.write("#### 🔹 เครื่อง CNC (9 เครื่อง)")
        cnc_col1, cnc_col2, cnc_col3 = st.columns(3)
        cnc_idx = 0
        # ---- บอร์ดแสดงรายแผนก ----
        st.write("#### 🔹 เครื่องจักรและเครื่องมือวัดทั้งหมด")
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
        with cutter_grind_col1: render_machine_card("CUTTER GRINDING-01", MACHINES["CUTTER GRINDING-01"], "CUTTER GRINDING")

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
                    render_machine_card(m_id, m_name, m_id) 
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
                with (mill_col1 if mill_idx % 3 == 0 else (mill_col2 if mill_idx % 3 == 1 else mill_col3)):
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

        # ---- 10. แผนก BENDING ----
        st.write("#### 🔹 เครื่องพับ BENDING (1 เครื่อง)")
        bend_col1, = st.columns(1)
        with bend_col1: render_machine_card("BENDING-01", MACHINES["BENDING-01"], "BENDING")

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
            if "BAND" in m_id:
                with (saw_col1 if saw_idx % 3 == 0 else (saw_col2 if saw_idx % 3 == 1 else saw_col3)):
                    render_machine_card(m_id, m_name, "BAND SAW")
                saw_idx += 1

        # ---- 15. แผนก FORKLIFT ----
        st.write("#### 🔹 รถโฟคลิฟ FORKLIFT (1 เครื่อง)")
        fork_col1, = st.columns(1)
        with fork_col1: render_machine_card("FORKLIFT-01", MACHINES["FORKLIFT-01"], "FORKLIFT")

        # 👑 พื้นที่ควบคุมระดับความปลอดภัยสูงสุด (สำหรับผู้บริหารสูงสุด)
        st.markdown("---")
        st.write("### 👑 พื้นที่ควบคุมระดับความปลอดภัยสูงสุด (สำหรับผู้บริหาร)")
        bigboss_code_input = st.text_input("🔐 กรุณากรอกรหัสผ่านผู้บริหาร เพื่อเปิดตู้นิรภัย พิมพ์คิวอาร์โค้ด และระบบล้างประวัติหลังบ้าน:", type="password", key="bigboss_secret_key_field")
        
        if bigboss_code_input == BIGBOSS_PASSWORD:
            st.success("🎯 ยืนยันสิทธิ์ สำเร็จ ปลดล็อกเรียบร้อยแล้วครับ!")
            
            with st.expander("🖨️ [เฉพาะผู้บริหารสูงสุด] เครื่องมือพิมพ์ QR Code สำหรับไปแปะหน้าเครื่องจักร"):
                sel_m = st.selectbox("เลือกเครื่องที่ต้องการพิมพ์ QR:", list(MACHINES.keys()), key="bigboss_qr_select_box")
                qr_url = f"https://pes-maintenance.streamlit.app/?id={sel_m}" 
                qr = qrcode.make(qr_url)
                buf = BytesIO()
                qr.save(buf)
                st.image(buf, caption=f"QR สำหรับแปะหน้าเครื่อง {MACHINES[sel_m]}")

            with st.expander("📦 [เฉพาะผู้บริหารสูงสุด] ตู้เซฟเก็บประวัติเอกสารย้อนหลังอัตโนมัติ (BACKUP HISTORY ARCHIVES)"):
                st.info("📂 ส่วนนี้เป็นที่รวบรวมไฟล์ Excel ประจำเดือนเก่าที่ระบบทำการคัดลอกสำรอง (Auto-Backup) เก็บไว้ให้โดยอัตโนมัติทุก ๆ สิ้นเดือน")
                backup_folder_path = os.path.join(BASE_FOLDER, "maintenance_backups")
                if os.path.exists(backup_folder_path):
                    all_backups = [f for f in os.listdir(backup_folder_path) if f.lower().endswith('.xlsx')]
                    if all_backups:
                        for b_file in sorted(all_backups):
                            b_file_path = os.path.join(backup_folder_path, b_file)
                            with open(b_file_path, "rb") as f_data:
                                st.download_button(
                                    label=f"📥 ดาวน์โหลดไฟล์สำรอง: {b_file}",
                                    data=f_data,
                                    file_name=b_file,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=f"dl_backup_{b_file}"
                                )
                    else:
                        st.caption("ℹ️ ยังไม่มีไฟล์สำรองประวัติเดือนเก่าจัดเก็บในตู้นี้")
                else:
                    st.caption("ℹ️ ระบบกำลังเตรียมตู้เซฟ (จะปรากฏไฟล์แรกเมื่อช่างส่งฟอร์มประเดิมคนแรกในวันที่ 1 ของเดือนถัดไปครับ)")

            with st.expander("🧹 [เฉพาะผู้บริหารสูงสุด] กล่องเครื่องมือล้างระบบภาพถ่ายทดสอบ (RESET SYSTEM)"):
                st.warning("⚠️ คำเตือน: ปุ่มนี้จะทำการลบโฟลเดอร์รูปภาพหลักฐานที่ส่งทดสอบก่อนหน้านี้ทั้งหมดออกไปอย่างถาวร เพื่อให้ระบบสะอาดพร้อมเปิดใช้งานจริง")
                if st.button("🚨 สั่งลบรูปภาพทดสอบทั้งหมดกริบ 100%", type="primary"):
                    target_photo_folder = os.path.join(BASE_FOLDER, "maintenance_photos")
                    if os.path.exists(target_photo_folder):
                        shutil.rmtree(target_photo_folder) 
                        st.success("🧹 ลบโฟลเดอร์รูปภาพทดสอบทั้งหมดออกไปจากระบบคลาวด์สะอาดบริสุทธิ์เรียบร้อยแล้วครับ!")
                        st.balloons()
                    else:
                        st.info("✨ ระบบสะอาดอยู่แล้ว ไม่มีโฟลเดอร์ภาพเก่าค้างให้ลบครับ")
        elif bigboss_code_input != "":
            st.error("❌ รหัสผ่าน Big Boss ไม่ถูกต้อง! ปฏิเสธสิทธิ์การพิมพ์คิวอาร์ เข้าถึงไฟล์ประวัติย้อนหลัง และปุ่มล้างระบบ")

    elif password_input != "": 
        st.error("❌ รหัสผ่านไม่ถูกต้อง ไม่พบสิทธิ์เข้าใช้งานระบบตามรหัสนี้ครับ")
            if machine_id == m_id: # แสดงไฮไลท์เครื่องที่กดมาจากลิงก์ไลน์ก่อนเป็นอันดับแรก
                render_machine_card(m_id, m_name, m_type_selected)
