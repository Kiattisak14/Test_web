import psutil
import pyautogui
import time
import numpy      as np
import os
import easyocr
import pandas     as pd
import pyperclip
import subprocess
import requests
import win32gui
import win32process
from datetime import datetime

from Get_Download import process_all_excel_files
from Get_Download import delete_all_files_from_base_folder

# เพิ่ม fail-safe ของ pyautogui
pyautogui.FAILSAFE = True  # เอาเมาส์ไปที่มุมบนซ้ายของหน้าจอเพื่อหยุด
pyautogui.PAUSE = 0.1  # เพิ่มการหน่วงเวลาระหว่างคำสั่ง

# ---- การตั้งค่า ----
scroll_duration  = 3  # ระยะเวลาที่ต้องการให้สกอลล์ (วินาที)
scroll_speed     = 70  # ความเร็วในการสกอลล์ (ยิ่งมากยิ่งเร็ว)
scroll_direction = 1  # 1 คือเลื่อนขึ้น, -1 คือเลื่อนลง
NTFY_TOPIC    = "test-server14"
NTFY_SERVER   = "https://ntfy.sh"
COM           = 'COM1'
# กำหนดพื้นที่ค้นหา (ปรับตามหน้าจอของคุณ)
SEARCH_REGION = (23, 309, 320, 695)  # พื้นที่ค้นหาเดียว: x, y, width, height
export_counter = 0  # คลิกระฆังครั้งเดียวเมื่อส่งออกครั้งแรก
folder_download_status = r"C:\Users\Kiatt\OneDrive - Energy Absolute Public Co Ltd\Data_storage\Python_Project\TSB_Download\Com_Download_Status\COM1_Download_Status.xlsx"

# ---- การตั้งค่าเวลา (เป็นเวลาในแต่ละวัน) ว่าหลังจากเวลานี้ค่อยเริ่ม Auto_Download ----
# รูปแบบ 'HH:MM:SS'  เช่น '14:00:00' คือ รอจนกว่าเวลาในวันนี้จะเกินหรือเท่ากับ 14:00 แล้วค่อยทำ Auto_Download
AFTER_TIME_STR = '09:50:00'

def get_active_process_name():
    try:
        # 1. หา Handle ของหน้าต่างที่ Active
        hwnd = win32gui.GetForegroundWindow()

        # 2. จาก Handle หา Process ID (PID)
        # tid, pid = win32process.GetWindowThreadProcessId(hwnd)
        # ใช้ [1] เพื่อเอาเฉพาะ pid
        pid = win32process.GetWindowThreadProcessId(hwnd)[1]

        # 3. จาก PID หาชื่อ Process
        # ใช้ psutil จะง่ายกว่าการใช้ win32api โดยตรง
        process = psutil.Process(pid)
        return process.name().lower()

    except Exception as e:
        # อาจเกิด Error ถ้าสลับหน้าต่างเร็วไป หรือเป็น Process ของระบบ
        # print(f"Error: {e}")
        return None
def open_chrome(url=None):
    candidate_paths = [
        r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        r"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
    ]

    for chrome_path in candidate_paths:
        if os.path.exists(chrome_path):
            if url:
                subprocess.Popen([chrome_path, url])
            else:
                subprocess.Popen([chrome_path])
            return True

    try:
        if url:
            subprocess.Popen(f'start "" chrome "{url}"', shell=True)
        else:
            subprocess.Popen('start "" chrome', shell=True)
        return True
    except Exception:
        return False
def Alert(mess):
    """
    ส่งแจ้งเตือนไปยัง ntfy ว่าเครื่องกำลังจะปิด
    """
    # -----------------------------------------------------------------
    # [FIX] แก้ไขตรงนี้ครับ
    # -----------------------------------------------------------------
    # Title (Header) ต้องเป็น ASCII (ภาษาอังกฤษ) เท่านั้น
    title = f"TSB-Download Alert! ({COM})"

    # Message (Data) เป็น UTF-8 (ไทย + Emoji) ได้
    message = f"🔴 {mess} '{COM}' มาเเก้ไขที"
    # -----------------------------------------------------------------

    requests.post(
        f"{NTFY_SERVER}/{NTFY_TOPIC}",
        data=message.encode('utf-8'),  # Message เข้ารหัส UTF-8
        headers={
            "Title": title,  # Title เป็น ASCII แล้ว
            "Priority": "high",
            "Tags": "warning,computer"
        },
        timeout=5
    )
def close_chrome():
    """
    ปิด Google Chrome แบบสุภาพก่อน แล้วค่อยบังคับปิด (กันแท็บค้าง)
    """
    try:
        # พยายามปิดแบบปกติ
        subprocess.run(
            ["taskkill", "/IM", "chrome.exe"],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        time.sleep(1)
        # กันเหนียว ปิดรวม subprocess ทั้งหมด
        subprocess.run(
            ["taskkill", "/IM", "chrome.exe", "/F", "/T"],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        print("ปิด Google Chrome แล้ว")
    except Exception as e:
        print(f"ปิด Google Chrome ไม่สำเร็จ: {e}")
def Restart_Chrome():
    global export_counter
    export_counter = 0  # reset counter so the bell can be clicked again after restart

    # วนลูปจนกว่าจะ login สำเร็จ
    login_success = False
    login_fail_count = 0  # นับจำนวนครั้งที่ login ไม่สำเร็จ
    restart_fail_count = 0  # นับจำนวนครั้งที่ restart ไม่สำเร็จ (อยู่นอกลูป login เพื่อเก็บค่าต่อเนื่อง)
    while not login_success:
        pyautogui.click(x=1615, y=774)  # กดล้าง download
        time.sleep(1)
        pyautogui.click(x=1893, y=19)  # กดปิด Chrome
        time.sleep(3)

        # เปิด Google Chrome
        open_chrome(
            "http://fleet.thaismilebus.com/_tsb/login?redirect=%2Findex")  # ถ้าต้องการเปิด URL ให้ใส่เช่น open_chrome("https://google.com")
        time.sleep(20)
        pyautogui.click(x=1520, y=454)  # กด User name
        time.sleep(1)
        pyautogui.write('SaranakomCh', interval=0.1)  # พิมพ์คำว่า SaranakomCh
        time.sleep(1)
        pyautogui.click(x=1497, y=532)  # กด Password
        pyautogui.write('On$542809', interval=0.1)  # พิมพ์คำว่า On$542809
        time.sleep(1)
        pyautogui.click(x=1463, y=602)  # กด Login
        time.sleep(1)
        pyautogui.click(x=1463, y=602)  # กด Login
        time.sleep(20)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(1)
        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.5)  # รอสักครู่ให้ copy เสร็จ
        clipboard_data = pyperclip.paste()
        cleaned_data = clipboard_data.replace('\r\n', '\n').replace('\r', '\n')
        lines = cleaned_data.strip().split('\n')
        lines = [line.strip() for line in lines if line.strip()]

        # ตรวจสอบว่ามีคำที่ต้องการทั้งหมดหรือไม่
        required_words = [
            'Dashboard',
            'วิดีโอ',
            'ความปลอดภัย',
            'จัดตารางเดินรถ',
            'จัดการการเดินรถ',
            'รายงานสถิติผู้โดยสาร',
            'รายงานแผนเดินรถ',
            'Operation Management',
            'ข้อมูลพื้นฐาน',
            'การจัดการระบบ',
            'ERP'
        ]

        # รวม lines ทั้งหมดเป็น string เดียวเพื่อตรวจสอบ substring
        all_text = ' '.join(lines)

        # ตรวจสอบแต่ละคำและเก็บคำที่ขาด
        missing_words = []
        for word in required_words:
            if word not in all_text:
                missing_words.append(word)

        # Print ผลลัพธ์
        if not missing_words:
            print("login สำเร็จ")
            login_success = True  # ตั้งค่าให้ออกจากลูป

            # วนลูปตรวจสอบ restart จนกว่าจะสำเร็จหรือครบ 3 ครั้ง
            restart_success = False
            while not restart_success:
                # ขั้นตอนที่ต้องทำก่อนตรวจสอบ restart
                pyautogui.click(x=1298, y=397)  # กดรายงานสถิติผู้โดยสาร
                time.sleep(10)
                pyautogui.click(x=445, y=151)  # กดทำได้
                time.sleep(5)
                pyautogui.click(x=770, y=289)  # กดประวัติข้อมูล
                time.sleep(15)

                # ตรวจสอบ restart
                pyautogui.click(x=1353, y=979)
                time.sleep(1)
                pyautogui.click(x=1353, y=979)
                time.sleep(1)
                pyautogui.hotkey("ctrl", "a")
                time.sleep(1)
                pyautogui.hotkey("ctrl", "c")
                time.sleep(0.5)  # รอสักครู่ให้ copy เสร็จ
                clipboard_data = pyperclip.paste()
                if clipboard_data == '1':
                    print('restart สำเร็จ')
                    restart_success = True  # ตั้งค่าให้ออกจากลูป
                    restart_fail_count = 0  # reset counter เมื่อ restart สำเร็จ
                else:
                    restart_fail_count += 1  # เพิ่มจำนวนครั้งที่ restart ไม่สำเร็จ
                    print('restart ไม่สำเร็จ')
                    print(f"จำนวนครั้งที่ restart ไม่สำเร็จ: {restart_fail_count}/3")

                    if restart_fail_count >= 3:
                        print("restart ไม่สำเร็จครบ 3 ครั้งแล้ว ส่ง Alert และหยุดการทำงาน")
                        Alert('แจ้งเตือน Restart Chrome ไม่สำเร็จครบ 3 ครั้ง')
                        # วนลูปรอแก้ไขเรื่อยๆ
                        while True:
                            print("รอแก้ไข...")
                            time.sleep(5)  # รอ 5 วินาทีแล้ว print อีกครั้ง
                    else:
                        print("กำลังลอง restart ใหม่...")
                        time.sleep(2)  # รอสักครู่ก่อนลองใหม่
                        # ย้อนกลับไปเริ่มใหม่ตั้งแต่ global export_counter
                        export_counter = 0  # reset counter so the bell can be clicked again after restart
                        login_success = False  # reset เพื่อให้ลูป login เริ่มใหม่
                        login_fail_count = 0  # reset counter
                        break  # ออกจากลูป restart เพื่อเริ่มลูป login ใหม่
        else:
            login_fail_count += 1  # เพิ่มจำนวนครั้งที่ login ไม่สำเร็จ
            print("login ไม่สำเร็จ")
            print(f"คำที่ขาด: {', '.join(missing_words)}")
            print(f"จำนวนครั้งที่ login ไม่สำเร็จ: {login_fail_count}/5")

            if login_fail_count >= 5:
                print("login ไม่สำเร็จครบ 5 ครั้งแล้ว ส่ง Alert และหยุดการทำงาน")
                Alert('แจ้งเตือน Login ไม่สำเร็จครบ 5 ครั้ง')
                # วนลูปรอแก้ไขเรื่อยๆ
                while True:
                    print("รอแก้ไข...")
                    time.sleep(5)  # รอ 5 วินาทีแล้ว print อีกครั้ง
            else:
                print("กำลังรีสตาร์ท Chrome และลองใหม่...")
                time.sleep(2)  # รอสักครู่ก่อนลองใหม่
def Open_Chrome():
    global export_counter
    export_counter = 0  # reset counter so the bell can be clicked again after restart

    # วนลูปจนกว่าจะ login สำเร็จ
    login_success = False
    login_fail_count = 0  # นับจำนวนครั้งที่ login ไม่สำเร็จ
    restart_fail_count = 0  # นับจำนวนครั้งที่ restart ไม่สำเร็จ (อยู่นอกลูป login เพื่อเก็บค่าต่อเนื่อง)
    while not login_success:
        # เปิด Google Chrome
        open_chrome(
            "http://fleet.thaismilebus.com/_tsb/login?redirect=%2Findex")  # ถ้าต้องการเปิด URL ให้ใส่เช่น open_chrome("https://google.com")
        time.sleep(20)
        pyautogui.click(x=1882, y=137)  # กดปิด
        time.sleep(1)
        pyautogui.click(x=1882, y=137)  # กดปิด
        time.sleep(1)
        pyautogui.click(x=1520, y=454)  # กด User name
        time.sleep(1)
        pyautogui.write('SaranakomCh', interval=0.1)  # พิมพ์คำว่า SaranakomCh
        time.sleep(1)
        pyautogui.click(x=1497, y=532)  # กด Password
        pyautogui.write('On$542809', interval=0.1)  # พิมพ์คำว่า On$542809
        time.sleep(1)
        pyautogui.click(x=1463, y=602)  # กด Login
        time.sleep(1)
        pyautogui.click(x=1463, y=602)  # กด Login
        time.sleep(20)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(1)
        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.5)  # รอสักครู่ให้ copy เสร็จ
        clipboard_data = pyperclip.paste()
        cleaned_data = clipboard_data.replace('\r\n', '\n').replace('\r', '\n')
        lines = cleaned_data.strip().split('\n')
        lines = [line.strip() for line in lines if line.strip()]

        # ตรวจสอบว่ามีคำที่ต้องการทั้งหมดหรือไม่
        required_words = [
            'Dashboard',
            'วิดีโอ',
            'ความปลอดภัย',
            'จัดตารางเดินรถ',
            'จัดการการเดินรถ',
            'รายงานสถิติผู้โดยสาร',
            'รายงานแผนเดินรถ',
            'Operation Management',
            'ข้อมูลพื้นฐาน',
            'การจัดการระบบ',
            'ERP'
        ]

        # รวม lines ทั้งหมดเป็น string เดียวเพื่อตรวจสอบ substring
        all_text = ' '.join(lines)

        # ตรวจสอบแต่ละคำและเก็บคำที่ขาด
        missing_words = []
        for word in required_words:
            if word not in all_text:
                missing_words.append(word)

        # Print ผลลัพธ์
        if not missing_words:
            print("login สำเร็จ")
            login_success = True  # ตั้งค่าให้ออกจากลูป

            # วนลูปตรวจสอบ restart จนกว่าจะสำเร็จหรือครบ 3 ครั้ง
            restart_success = False
            while not restart_success:
                # ขั้นตอนที่ต้องทำก่อนตรวจสอบ restart
                pyautogui.click(x=1298, y=397)  # กดรายงานสถิติผู้โดยสาร
                time.sleep(10)
                pyautogui.click(x=445, y=151)  # กดทำได้
                time.sleep(5)
                pyautogui.click(x=770, y=289)  # กดประวัติข้อมูล
                time.sleep(15)

                # ตรวจสอบ restart
                pyautogui.click(x=1353, y=979)
                time.sleep(1)
                pyautogui.click(x=1353, y=979)
                time.sleep(1)
                pyautogui.hotkey("ctrl", "a")
                time.sleep(1)
                pyautogui.hotkey("ctrl", "c")
                time.sleep(0.5)  # รอสักครู่ให้ copy เสร็จ
                clipboard_data = pyperclip.paste()
                if clipboard_data == '1':
                    print('restart สำเร็จ')
                    restart_success = True  # ตั้งค่าให้ออกจากลูป
                    restart_fail_count = 0  # reset counter เมื่อ restart สำเร็จ
                else:
                    restart_fail_count += 1  # เพิ่มจำนวนครั้งที่ restart ไม่สำเร็จ
                    print('restart ไม่สำเร็จ')
                    print(f"จำนวนครั้งที่ restart ไม่สำเร็จ: {restart_fail_count}/3")

                    if restart_fail_count >= 3:
                        print("restart ไม่สำเร็จครบ 3 ครั้งแล้ว ส่ง Alert และหยุดการทำงาน")
                        Alert('แจ้งเตือน Restart Chrome ไม่สำเร็จครบ 3 ครั้ง')
                        # วนลูปรอแก้ไขเรื่อยๆ
                        while True:
                            print("รอแก้ไข...")
                            time.sleep(5)  # รอ 5 วินาทีแล้ว print อีกครั้ง
                    else:
                        print("กำลังลอง restart ใหม่...")
                        time.sleep(2)  # รอสักครู่ก่อนลองใหม่
                        # ย้อนกลับไปเริ่มใหม่ตั้งแต่ global export_counter
                        export_counter = 0  # reset counter so the bell can be clicked again after restart
                        login_success = False  # reset เพื่อให้ลูป login เริ่มใหม่
                        login_fail_count = 0  # reset counter
                        break  # ออกจากลูป restart เพื่อเริ่มลูป login ใหม่
        else:
            login_fail_count += 1  # เพิ่มจำนวนครั้งที่ login ไม่สำเร็จ
            print("login ไม่สำเร็จ")
            print(f"คำที่ขาด: {', '.join(missing_words)}")
            print(f"จำนวนครั้งที่ login ไม่สำเร็จ: {login_fail_count}/5")

            if login_fail_count >= 5:
                print("login ไม่สำเร็จครบ 5 ครั้งแล้ว ส่ง Alert และหยุดการทำงาน")
                Alert('แจ้งเตือน Login ไม่สำเร็จครบ 5 ครั้ง')
                # วนลูปรอแก้ไขเรื่อยๆ
                while True:
                    print("รอแก้ไข...")
                    time.sleep(5)  # รอ 5 วินาทีแล้ว print อีกครั้ง
            else:
                print("กำลังรีสตาร์ท Chrome และลองใหม่...")
                time.sleep(2)  # รอสักครู่ก่อนลองใหม่
def Re_login():
    time.sleep(5)
    pyautogui.click(x=1520, y=454)  # กด User name
    time.sleep(1)
    pyautogui.click(x=1549, y=533)  # กด User saranacom
    time.sleep(1)
    pyautogui.click(x=1546, y=605)  # กด login
    time.sleep(4)
    pyautogui.click(x=1288, y=403)  # กดรายงานสถิติผู้โดยสาร
    time.sleep(3)
    pyautogui.click(x=445, y=151)  # กดทำได้
    time.sleep(2)
    pyautogui.click(x=770, y=289)  # กดประวัติข้อมูล
    time.sleep(15)
def find_text_in_region(text_to_find, region=None):
    """
    ค้นหาข้อความในพื้นที่ที่กำหนดด้วย EasyOCR
    """
    if region is None:
        region = SEARCH_REGION

    x, y, w, h = region
    print(f"กำลังค้นหาข้อความ: '{text_to_find}' ในพื้นที่: ({x}, {y}, {w}, {h})")

    # ถ่ายภาพเฉพาะพื้นที่
    screenshot    = pyautogui.screenshot(region=(x, y, w, h))
    screenshot_np = np.array(screenshot)

    # ใช้ EasyOCR อ่านข้อความ
    try:
        reader  = easyocr.Reader(['en'], gpu=False)
        results = reader.readtext(screenshot_np)

        print(f"ข้อความที่อ่านได้: {[text for _, text, _ in results]}")

        for (bbox, text, prob) in results:
            if text_to_find.lower() in text.lower():
                print(f"เจอแล้ว! ในพื้นที่ที่กำหนด")

                # หาตำแหน่งที่แม่นยำ
                top_left = tuple(bbox[0])
                bottom_right = tuple(bbox[2])

                word_x = top_left[0] + x  # ตำแหน่งสัมพัทธ์กับหน้าจอ
                word_y = top_left[1] + y
                word_w = bottom_right[0] - top_left[0]
                word_h = bottom_right[1] - top_left[1]

                print(f"ตำแหน่ง: ({word_x}, {word_y}, {word_w}, {word_h})")
                return (word_x, word_y, word_w, word_h)

        print("ไม่พบข้อความที่ต้องการในพื้นที่ที่กำหนด")
        return None

    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการอ่านพื้นที่: {e}")
        return None
def find_text_with_easyocr_fast(text_to_find):
    """
    ฟังก์ชันสำหรับค้นหาข้อความบนหน้าจอด้วย EasyOCR (เร็วกว่าเดิม)
    """
    print(f"กำลังค้นหาข้อความ: '{text_to_find}' ด้วย EasyOCR...")

    # 1. ถ่ายภาพหน้าจอ
    screenshot = pyautogui.screenshot()
    screenshot_np = np.array(screenshot)

    # 2. ใช้ EasyOCR อ่านข้อความทั้งหมด
    try:
        reader = easyocr.Reader(['en'], gpu=False)
        results = reader.readtext(screenshot_np)

        print(f"ข้อความที่อ่านได้: {[text for _, text, _ in results][:5]}...")  # แสดงแค่ 5 ข้อความแรก

        for (bbox, text, prob) in results:
            if text_to_find.lower() in text.lower():
                print(f"เจอแล้ว! '{text_to_find}' ในข้อความ")

                # 3. หาตำแหน่งที่แม่นยำ
                top_left = tuple(bbox[0])
                bottom_right = tuple(bbox[2])

                x = top_left[0]
                y = top_left[1]
                w = bottom_right[0] - top_left[0]
                h = bottom_right[1] - top_left[1]

                print(f"ตำแหน่ง: ({x}, {y}, {w}, {h})")
                return (x, y, w, h)

        print("ไม่พบข้อความที่ต้องการในหน้าจอ")
        return None

    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการอ่านข้อความ: {e}")
        return None
def find_text_with_easyocr(text_to_find, language_codes=['en']):
    """
    ฟังก์ชันสำหรับค้นหาข้อความบนหน้าจอด้วย EasyOCR และคืนค่าตำแหน่ง
    (เก็บไว้เป็นตัวสำรอง)
    """
    print(f"กำลังเตรียม EasyOCR Reader... (ครั้งแรกอาจจะช้าเพราะต้องโหลดโมเดล)")
    # 1. สร้าง Reader ของ EasyOCR (ระบุภาษาที่ต้องการใช้)
    # หากต้องการใช้ภาษาไทยด้วย ให้ใส่ ['th', 'en']
    reader = easyocr.Reader(language_codes, gpu=False)  # gpu=False เพื่อให้แน่ใจว่าทำงานได้บนเครื่องที่ไม่มีการ์ดจอแรงๆ

    print(f"กำลังค้นหาข้อความ: '{text_to_find}'...")

    # 2. ถ่ายภาพหน้าจอและแปลงเป็นรูปแบบที่ EasyOCR ใช้
    screenshot = pyautogui.screenshot()
    screenshot_np = np.array(screenshot)

    # 3. ใช้ EasyOCR อ่านข้อความจากภาพ
    results = reader.readtext(screenshot_np)

    # 4. วนลูปเพื่อค้นหาข้อความที่ตรงกัน
    for (bbox, text, prob) in results:
        # bbox คือ list ของตำแหน่งมุม 4 มุม [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
        # text คือข้อความที่อ่านได้
        # prob คือระดับความน่าเชื่อถือ (0 ถึง 1)
        if text_to_find.lower() in text.lower():
            print(f"เจอแล้ว! '{text}' มีข้อความที่เราค้นหา")

            # ดึงตำแหน่งมุมบนซ้าย (top_left) และมุมล่างขวา (bottom_right)
            top_left = tuple(bbox[0])
            bottom_right = tuple(bbox[2])

            # คืนค่าเป็นตำแหน่งและขนาดเพื่อง่ายต่อการใช้งาน
            x = top_left[0]
            y = top_left[1]
            w = bottom_right[0] - top_left[0]
            h = bottom_right[1] - top_left[1]

            return (x, y, w, h)

    # 5. ถ้าวนจนครบแล้วยังหาไม่เจอ
    print("ไม่พบข้อความที่ต้องการบนหน้าจอ")
    return None
def find_download_state(text_to_find, language_codes=['en']):
    """
    ฟังก์ชันสำหรับค้นหาข้อความบนหน้าจอด้วย EasyOCR และคืนค่าตำแหน่ง
    """
    print(f"กำลังเตรียม EasyOCR Reader... (ครั้งแรกอาจจะช้าเพราะต้องโหลดโมเดล)")
    reader = easyocr.Reader(language_codes, gpu=False)  # gpu=False เพื่อให้แน่ใจว่าทำงานได้บนเครื่องที่ไม่มีการ์ดจอแรงๆ
    print(f"กำลังค้นหาข้อความ: '{text_to_find}'...")
    screenshot = pyautogui.screenshot()
    screenshot_np = np.array(screenshot)
    results = reader.readtext(screenshot_np)
    for (bbox, text, prob) in results:
        if text_to_find.lower() in text.lower():
            print(f"เจอแล้ว! '{text}' มีข้อความที่เราค้นหา")
            return 1
    print("ไม่พบข้อความที่ต้องการบนหน้าจอ")
    return 0
def find_plate_num(Plate_num):
    pyautogui.click(x=191, y=283)
    time.sleep(1)
    pyautogui.click(x=191, y=283)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(1)
    pyautogui.press("backspace")
    time.sleep(1)
    pyautogui.write(str(Plate_num))
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.hotkey("backspace")
    time.sleep(1)

    # ใช้ Region-based Search + EasyOCR แทน Tesseract
    location = find_text_in_region(Plate_num)
    if location:
        x, y, w, h = location
        center_x = x + w / 2
        center_y = y + h / 2
        pyautogui.click(x=178, y=center_y)
    else:
        # ถ้าไม่เจอ ให้ใช้ EasyOCR ทั้งหน้าจอ
        location = find_text_with_easyocr_fast(Plate_num)
        if location:
            x, y, w, h = location
            center_x = x + w / 2
            center_y = y + h / 2
            pyautogui.click(x=178, y=center_y)
        else:
            print(f"ไม่พบ {Plate_num} บนหน้าจอ")

    time.sleep(10)
def Select_Time_and_Download(Time_Start, Time_Stop, Date):
    global export_counter
    time.sleep(2)

    # ตั้งค่า Date/Time แค่ครั้งแรกครั้งเดียว
    if not hasattr(Select_Time_and_Download, "_initialized"):
        pyautogui.click(x=660, y=424)
        time.sleep(0.5)
        pyautogui.click(x=660, y=424)
        time.sleep(2)
        pyautogui.click(x=746, y=490)
        time.sleep(1)
        pyautogui.click(x=746, y=490)
        time.sleep(1)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.2)
        pyautogui.press("backspace")
        time.sleep(0.2)
        pyautogui.write(str(Date))
        time.sleep(0.2)
        pyautogui.click(x=910, y=490)
        time.sleep(0.2)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.2)
        pyautogui.press("backspace")
        time.sleep(0.2)
        pyautogui.write(str(Time_Start))
        time.sleep(0.2)
        pyautogui.click(x=1145, y=490)
        time.sleep(0.2)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.2)
        pyautogui.press("backspace")
        time.sleep(0.2)
        pyautogui.write(str(Date))
        time.sleep(0.2)
        pyautogui.click(x=1315, y=490)
        time.sleep(0.2)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.2)
        pyautogui.press("backspace")
        time.sleep(0.2)
        pyautogui.write(str(Time_Stop))
        time.sleep(0.2)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.2)
        pyautogui.press("backspace")
        time.sleep(0.2)
        pyautogui.write(str(Time_Stop))
        time.sleep(0.2)
        pyautogui.click(x=1356, y=951)  # กดตกลง
        time.sleep(0.2)
        pyautogui.click(x=1043, y=424)  # กดค้นหา
        time.sleep(10)
        # mark ว่าตั้งค่า Date/Time แล้ว
        Select_Time_and_Download._initialized = True


    pyautogui.hotkey("ctrl", "a")
    time.sleep(1)
    pyautogui.hotkey("ctrl", "c")
    time.sleep(0.5)  # รอสักครู่ให้ copy เสร็จ
    clipboard_data = pyperclip.paste()
    cleaned_data   = clipboard_data.replace('\r\n', '\n').replace('\r', '\n')
    lines = cleaned_data.strip().split('\n')
    lines = [line.strip() for line in lines if line.strip()]
    # ตรวจสอบว่ามีสมาชิกของ lines ที่มีค่าเท่ากับ 'ไม่มีข้อมูล' หรือไม่
    has_no_data     = any(line == 'ไม่มีข้อมูล' for line in lines)
    download_detect = any(line == 'กำลังส่งออก…' for line in lines)
    login_state     = any(line == 'โหมดเข้าสู่ระบบในเครื่อง' for line in lines)
    if login_state:
        Re_login()
    else:
        if download_detect:
            pyautogui.click(x=1120, y=621)  # กดดาวน์โหลดด้านหลัง
            time.sleep(1)
        else:
            if has_no_data:
                print("ไม่มีข้อมูล")
                pyautogui.click(x=1285, y=658)  # กดกลางจอ
                time.sleep(0.2)
                return 0
            else:
                print("มีข้อมูล")
                pyautogui.click(x=1285, y=658)  # กดกลางจอ
                time.sleep(1)
                pyautogui.click(x=453, y=333)  # กดส่งออก
                time.sleep(5)
                pyautogui.click(x=1120, y=621)  # กดดาวน์โหลดด้านหลัง
                time.sleep(1)
                # mark ว่ามีการสั่งส่งออกแล้ว (แต่ยังไม่กดระฆังในขั้นนี้)
                if export_counter == 0:
                    export_counter = 1
                return 1
def count_files_in_folder(folder_path):
    total_files = 0
    for root, dirs, files in os.walk(folder_path):
        total_files += len(files)
    return total_files
def Column_name(df):
    if len(df.columns) == 32:
        df.columns = ['Vehicle_ID', 'Plate_NO', 'Route', 'Group',
                      'Data_Refresh_Time', 'Charging_Status', 'SOC', 'Speed',
                      'Total_Distance', 'CAN_Bus_Instantaneous', 'Service_Status', 'Indicator_Light_Status',
                      'Gear_Status', 'Brake_Pedal_Switch', 'Front_Braking_Pressure',
                      'Rear_Braking_Pressure', 'Total_Battery_Voltage', 'Total_Current',
                      'Min_Battery_Cell_Voltage',
                      'DCDC_Current', 'Max_Battery_Cell_Voltage', 'Min_Battery_Cell_Temperature',
                      'Max_Battery_Cell_Temperature',
                      'Positive_Insulation_Impedance', 'Negative_Insulation_Impedance',
                      'The_drive_motor_state', 'Motor_Controller_Voltage',
                      'Motor_Controller_Current', 'Drive_Motor_Temperature',
                      'Drive_Motor_Controller_Temperature', 'Drive_Motor_Rotate_Speed',
                      'Alarm_Data']
        return df
    elif len(df.columns) == 61:
        df.columns.values[0] = 'Data_Refresh_Time'
        df.columns.values[1] = 'Vehicle_ID'
        df.columns.values[2] = 'Plate_NO'
        df.columns.values[3] = 'Route'
        df.columns.values[4] = 'Group'
        df.columns.values[5] = 'Service_Status'
        df.columns.values[7] = 'Charging_Status'
        df.columns.values[8] = 'Speed'
        df.columns.values[9] = 'Total_Distance'
        df.columns.values[10] = 'Total_Battery_Voltage'
        df.columns.values[11] = 'Total_Current'
        df.columns.values[12] = 'SOC'
        df.columns.values[15] = 'Gear_Status'
        df.columns.values[19] = 'Positive_Insulation_Impedance'
        df.columns.values[20] = 'Negative_Insulation_Impedance'
        df.columns.values[24] = 'The_drive_motor_state'
        df.columns.values[25] = 'Drive_Motor_Temperature'
        df.columns.values[26] = 'Drive_Motor_Controller_Temperature'
        df.columns.values[27] = 'Drive_Motor_Rotate_Speed'
        df.columns.values[29] = 'Motor_Controller_Voltage'
        df.columns.values[30] = 'Motor_Controller_Current'
        df.columns.values[36] = 'Max_Battery_Cell_Voltage'
        df.columns.values[39] = 'Min_Battery_Cell_Voltage'
        df.columns.values[42] = 'Max_Battery_Cell_Temperature'
        df.columns.values[45] = 'Min_Battery_Cell_Temperature'
        return df
    else:
        print('Data Type Error')
        return df
def combine_excel_files(folder_path, output_file):
    combined_df = pd.DataFrame()
    column_names = None
    download_state = 0
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            file_path = os.path.join(folder_path, filename)
            try:
                # อ่านไฟล์โดยข้าม 3 แถวแรก
                df = pd.read_excel(file_path, skiprows=2)
                df = Column_name(df)
                combined_df = pd.concat([combined_df, df], ignore_index=True)
            except Exception as e:
                print(f"ไม่สามารถอ่านไฟล์ {filename} ได้: {e}")

    # บันทึกไฟล์รวม
    combined_df = combined_df[
        ['Vehicle_ID', 'Plate_NO', 'Route', 'Group', 'Data_Refresh_Time', 'Charging_Status', 'SOC',
         'Speed', 'Total_Distance', 'Service_Status', 'Gear_Status',
         'Total_Battery_Voltage', 'Total_Current', 'Min_Battery_Cell_Voltage', 'Max_Battery_Cell_Voltage',
         'Min_Battery_Cell_Temperature', 'Max_Battery_Cell_Temperature', 'Positive_Insulation_Impedance',
         'Negative_Insulation_Impedance',
         'The_drive_motor_state', 'Motor_Controller_Voltage', 'Motor_Controller_Current',
         'Drive_Motor_Temperature',
         'Drive_Motor_Controller_Temperature', 'Drive_Motor_Rotate_Speed']]
    combined_df = combined_df.sort_values(by='Data_Refresh_Time', ascending=True)
    if not combined_df.empty:
        combined_df.to_excel(output_file, index=False)
        print(f"บันทึกข้อมูลลงใน {output_file} เรียบร้อยแล้ว")
        download_state = 1
    else:
        print("DataFrame ว่างเปล่า ไม่มีการบันทึกไฟล์ Excel")
        download_state = 0
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            os.remove(file_path)
            print(f"ลบไฟล์: {filename}")
        except Exception as e:
            print(f"ไม่สามารถลบไฟล์ {filename} ได้: {e}")
    return download_state
def check_duplicate_file_sizes(folder_path):
    """
    ฟังก์ชันสำหรับเช็คไฟล์ที่มีขนาดเท่ากันในโฟลเดอร์
    """
    file_sizes = {}
    duplicate_files = {}

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path):
            file_size = os.path.getsize(file_path)
            if file_size in file_sizes:
                if file_size not in duplicate_files:
                    duplicate_files[file_size] = [file_sizes[file_size]]
                duplicate_files[file_size].append(filename)
            else:
                file_sizes[file_size] = filename

    # แสดงผลไฟล์ที่มีขนาดเท่ากัน
    if duplicate_files:
        print("พบไฟล์ที่มีขนาดเท่ากัน:")
        for size, files in duplicate_files.items():
            print(f"ขนาด: {size} bytes")
            for file in files:
                print(f"  - {file}")
            print()
        return duplicate_files
    else:
        print("ไม่พบไฟล์ที่มีขนาดเท่ากัน")
        return {}
def refresh():
    time.sleep(3)
    pyautogui.click(x=1590, y=120)  # กดจดหมาย
    time.sleep(3)
    pyautogui.click(x=1682,y=618)   # กดล้างสถานะ
    time.sleep(1)
    pyautogui.click(x=1682, y=618)  # กดล้างสถานะ
    time.sleep(1)
    pyautogui.click(x=90, y=61)   # กด refresh
    time.sleep(10)
    pyautogui.click(x=609, y=226)  # กดประวัติข้อมูล
    time.sleep(10)
def check_export_status():
    """
    ตรวจข้อความบนหน้าเว็บว่า:
      - export_success: พบคำว่า 'ส่งออกสำเร็จ'
      - exporting:      พบคำขึ้นต้น '17กําลังส่งออก' หรือ '0กําลังส่งออก'
    """
    pyautogui.hotkey("ctrl", "a")
    time.sleep(1)
    pyautogui.hotkey("ctrl", "c")
    time.sleep(0.5)
    clipboard_data = pyperclip.paste()
    cleaned_data = clipboard_data.replace('\r\n', '\n').replace('\r', '\n')
    lines = [line.strip() for line in cleaned_data.strip().split('\n') if line.strip()]

    export_success = any('ส่งออกสำเร็จ' in line for line in lines)
    exporting      = any(
        line.startswith('17กําลังส่งออก') or line.startswith('0กําลังส่งออก')
        for line in lines
    )
    return export_success, exporting
def Auto_Download(folder_path,refresh_number):
    initial_count      = count_files_in_folder(folder_path)
    download_detected  = False
    check_attempt_auto = 0  # ตัวนับเคสไม่พบทั้ง "ส่งออกสำเร็จ" และ "กำลังส่งออก"
    for i in range(0, refresh_number):
        # เช็กสถานะส่งออกทุกครั้งก่อนสั่งโหลด
        export_success, exporting = check_export_status()
        if export_success and not exporting:
            # ผ่านเงื่อนไข ไปขั้นโหลดไฟล์ได้
            pass
        elif exporting:
            print("ยังพบ 'กำลังส่งออก' -> รอ 5 นาทีแล้วตรวจใหม่")
            time.sleep(300)  # รอ 5 นาที
            continue
        else:
            # ไม่พบทั้ง 'ส่งออกสำเร็จ' และ 'กำลังส่งออก'
            check_attempt_auto += 1
            print(f"ยังไม่พบข้อความ 'ส่งออกสำเร็จ' (ครั้งที่ {check_attempt_auto}/10) ใน Auto_Download")

            if check_attempt_auto >= 10:
                print("ตรวจครบ 10 ครั้งแล้วยังไม่พบ 'ส่งออกสำเร็จ' -> Alert และรอให้ผู้ใช้มาแก้ไข (Auto_Download)")
                Alert("หน้าต่าง Download ไม่สามารถใช้งานได้ (Auto_Download ตรวจไม่พบ 'ส่งออกสำเร็จ' ครบ 10 ครั้ง)")
                while True:
                    print("รอให้ผู้ใช้มาแก้ไขหน้าต่าง Download ... (Auto_Download)")
                    time.sleep(5)

            print("ยังไม่พบ 'ส่งออกสำเร็จ' -> กดปุ่ม (1785, 243) แล้วกดระฆังใหม่ จากนั้นจะตรวจอีกครั้ง (Auto_Download)")
            pyautogui.click(x=1785, y=243)  # ปุ่มที่คุณระบุ
            time.sleep(5)
            pyautogui.click(x=1507, y=150)  # กดระฆังอีกครั้ง
            time.sleep(8)
            continue

        file_count_pre = count_files_in_folder(folder_path)
        pyautogui.click(x=1736, y=393)  # กด โหลดเอง
        time.sleep(5)
        file_count = count_files_in_folder(folder_path)
        if file_count > file_count_pre:
            download_detected = True
            pyautogui.click(x=1769, y=393)  # กดลบ
            time.sleep(10)
        else:
            time.sleep(5)
            pyautogui.click(x=1736, y=393)
            time.sleep(5)
            file_count = count_files_in_folder(folder_path)
            if file_count > file_count_pre:
                download_detected = True
                pyautogui.click(x=1769, y=393)  # กดลบ
                time.sleep(10)
            else:
                pyautogui.click(x=1769, y=393)  # กดลบ
                time.sleep(10)
    final_count = count_files_in_folder(folder_path)
    files_downloaded = final_count - initial_count  # จำนวนไฟล์ที่ดาวน์โหลดได้
    if final_count > initial_count:
        download_detected = True
    pyautogui.click(x=1612, y=776)  # กดล้าง
    time.sleep(1)
    pyautogui.click(x=1612, y=776)  # กดล้าง
    return download_detected, files_downloaded
def write_excel_download_status(Date_List,status):
    print('write_excel_download_status')
    try:
        # เตรียมแถวใหม่สำหรับบันทึกสถานะ (เขียนทับไฟล์เดิมเสมอ)
        df_out = pd.DataFrame([{
            'Date': Date_List[0],
            'Status': status
        }])[['Date', 'Status']]

        # ให้แน่ใจว่าโฟลเดอร์ปลายทางมีอยู่
        dst_dir = os.path.dirname(folder_download_status)
        if dst_dir and not os.path.exists(dst_dir):
            os.makedirs(dst_dir, exist_ok=True)

        # เขียนทับไฟล์เดิมทันที
        df_out.to_excel(folder_download_status, index=False)
    except Exception as e:
        print(f'write_excel_download_status error: {e}')

Time_Start_List        = ['00:00:00']
Time_Stop_List         = ['23:59:59']
folder_History         = 'D:\Python_Project\Process_Auto_BUS\History_Download_Com2.xlsx'
plate_dowload          = []
No_Data_Download       = []
time.sleep(10)

def wait_until_time_of_day(time_str: str):
    """
    รอจนกว่าเวลาปัจจุบัน (ของวันนี้) จะมากกว่าหรือเท่ากับ time_str (รูปแบบ 'HH:MM:SS')
    ถ้าตอนนี้เวลาเลย time_str ไปแล้ว จะไม่รอ
    """
    if not time_str:
        return

    try:
        target_time = datetime.strptime(time_str, "%H:%M:%S").time()
    except Exception as e:
        print(f"รูปแบบ AFTER_TIME_STR ไม่ถูกต้อง ({time_str}) ข้ามการรอเวลา: {e}")
        return

    print(f"เริ่มรอเวลาให้ถึง {time_str} ก่อนค่อยทำ Auto_Download ...")

    while True:
        now = datetime.now()
        if now.time() >= target_time:
            print(f"เวลาปัจจุบัน {now.time().strftime('%H:%M:%S')} ถึง/เกิน {time_str} แล้ว เริ่มทำ Auto_Download")
            break

        # เวลา (วินาที) ที่ยังเหลือจนถึง time_str ของวันนี้
        today_target = datetime.combine(now.date(), target_time)
        remaining = (today_target - now).total_seconds()

        if remaining <= 0:
            print(f"เวลาปัจจุบัน {now.time().strftime('%H:%M:%S')} ถึง/เกิน {time_str} แล้ว เริ่มทำ Auto_Download")
            break

        # เลือกช่วงเวลานอนให้เหมาะสม
        if remaining > 600:
            sleep_sec = 60
        elif remaining > 300:
            sleep_sec = 30
        elif remaining > 60:
            sleep_sec = 10
        else:
            sleep_sec = 5

        print(f"รอเวลาอยู่ เหลือประมาณ {int(remaining)} วินาที ก่อนถึง {time_str} (sleep {sleep_sec}s)")
        time.sleep(sleep_sec)
def convert_index_to_details(total_index, len_plate, len_date, len_time):
    # Adjust for 0-based indexing
    total_index_0_based = total_index - 1

    # Calculate Plate
    plate_index = total_index_0_based // (len_date * len_time)
    plate = plate_index + 1

    # Calculate Date
    date_index = (total_index_0_based % (len_date * len_time)) // len_time
    date = date_index + 1

    # Calculate Time
    time_index = total_index_0_based % len_time
    time = time_index + 1

    return (plate, date, time)
def TSB_Download(Plate_num_List, Date_List):
    global export_counter
    # Reset export_counter เมื่อเริ่ม TSB_Download ใหม่ทุกครั้ง
    export_counter = 0
    
    # เปิด Chrome และล็อกอินทุกครั้งก่อนเริ่มดาวน์โหลด
    Open_Chrome()

    # รีเซ็ตให้ Select_Time_and_Download ตั้งค่า Date/Time ใหม่ทุกครั้งที่เริ่ม TSB_Download
    if hasattr(Select_Time_and_Download, "_initialized"):
        delattr(Select_Time_and_Download, "_initialized")

    len_plate     = len(Plate_num_List)
    len_date      = len(Date_List)
    len_time      = len(Time_Start_List)
    Plate_Counter = 0           # นับจำนวน plate ที่มีข้อมูลจริง ๆ
    folder_path   = "D:\Download"
    plate_index   = 1           # เก็บ index ของ plate ปัจจุบัน
    Plate_num_1   = '16-xxxx'
    alert_shown   = False       # ตัวแปรติดตามว่าได้ Alert ไปแล้วหรือยัง
    bell_clicked  = False       # ตัวแปรติดตามว่ากดระฆังครั้งแรกแล้วหรือยัง

    # เดินให้ครบทุก combination ของ plate/date/time ในรอบเดียว
    while plate_index < (len_plate * len_date * len_time) + 1:
        # รอจนกว่าจะกลับมาที่ Google Chrome
        while True:
            process_name = get_active_process_name()
            if process_name == "chrome.exe":
                print(f"Active: Google Chrome (Process: {process_name})")
                alert_shown = False  # Reset เมื่อกลับมา Chrome
                break  # ออกจากลูปเมื่อเจอ Chrome
            else:
                print(f"Not Chrome. (Active: {process_name}) - รอจนกว่าจะกลับมาที่ Chrome...")
                # Alert แค่ครั้งเดียว
                if not alert_shown:
                    Alert("ตรวจพบว่าคุณอยู่นอก Google Chrome โปรดกลับมาที่ Google Chrome เพื่อดำเนินการต่อ")
                    alert_shown = True
                time.sleep(1)  # รอ 1 วินาทีแล้วตรวจสอบใหม่
        
        current_plate_index = plate_index
        print(Plate_Counter + 1)
        plate, date, time_download = convert_index_to_details(plate_index, len_plate, len_date, len_time)
        Plate_num  = Plate_num_List[plate-1]
        Date       = Date_List[date-1]
        Time_Start = Time_Start_List[time_download-1]
        Time_Stop  = Time_Stop_List[time_download-1]
        if Plate_num == Plate_num_1:
            pass
        else:
            find_plate_num(Plate_num)
        Data_State = Select_Time_and_Download(Time_Start, Time_Stop, Date)
        if Data_State == 1:
            # มีข้อมูล → เก็บ plate นี้ไว้ในประวัติ และนับเพิ่ม
            plate_dowload.append(Plate_num)
            Plate_Counter += 1
            
            # กดระฆังครั้งแรกทันทีหลังจาก plate แรกที่มีข้อมูลเสร็จ
            if not bell_clicked and export_counter == 1:
                print("เริ่มกดระฆังเพื่อโหลดไฟล์จาก TSB ... (หลัง plate แรกที่มีข้อมูล)")
                for _ in range(5):
                    pyautogui.click(x=1507, y=150)  # กดระฆัง
                    time.sleep(2)
                # กันไม่ให้มากดซ้ำอีกในรอบนี้
                bell_clicked = True
                export_counter = 2

        # ไป plate ถัดไป ไม่ต้องมี batch ตาม refresh_number แล้ว
        Plate_num_1 = Plate_num
        plate_index = current_plate_index + 1

    # เมื่อเดินครบทุก plate/date/time แล้ว:
    # ตรวจสอบว่ามีการกดระฆังแล้วหรือยัง ถ้ายังไม่กดและมีการส่งออก ให้กดระฆัง
    if export_counter == 1 and not bell_clicked:
        print("เริ่มกดระฆังเพื่อโหลดไฟล์จาก TSB ...")
        for _ in range(5):
            pyautogui.click(x=1507, y=150)  # กดระฆัง
            time.sleep(2)
        export_counter = 2
        bell_clicked = True

    # ตรวจสอบสถานะส่งออก (ถ้ามีการกดระฆังแล้ว)
    if bell_clicked:
        # หลังจากกดระฆังแล้ว ตรวจสอบซ้ำ ๆ ว่ามีข้อความ "ส่งออกสำเร็จ" บนหน้าเว็บหรือไม่
        # กรณีต่าง ๆ:
        #   - ถ้ามีคำว่า "ส่งออกสำเร็จ" -> ออกจากลูปแล้วไป Auto_Download
        #   - ถ้าไม่มี "ส่งออกสำเร็จ" แต่มีคำว่า "กำลังส่งออก" -> รอ 5 นาทีแล้วตรวจใหม่ (ไม่เพิ่มตัวนับครั้ง)
        #   - ถ้าไม่มีทั้งสองคำ -> กดปุ่ม (1785, 243) รอ 5 วินาที แล้วกดระฆังใหม่ และตรวจอีกครั้ง
        #   - ถ้าทำแบบนี้ครบ 10 ครั้งแล้วยังไม่เจอ "ส่งออกสำเร็จ" เลย -> Alert และรอให้ผู้ใช้มาแก้ไข
        check_attempt = 0
        while True:
            print("ตรวจสอบข้อความ 'ส่งออกสำเร็จ' บนหน้าเว็บก่อนเริ่ม Auto_Download ...")
            pyautogui.hotkey("ctrl", "a")
            time.sleep(1)
            pyautogui.hotkey("ctrl", "c")
            time.sleep(0.5)
            clipboard_data = pyperclip.paste()
            cleaned_data = clipboard_data.replace('\r\n', '\n').replace('\r', '\n')
            lines = [line.strip() for line in cleaned_data.strip().split('\n') if line.strip()]

            export_success = any('ส่งออกสำเร็จ' in line for line in lines)
            # สถานะ "กำลังส่งออก" ตอนนี้แสดงในรูปแบบ เช่น
            #   "17กําลังส่งออก %s..." หรือ "0กําลังส่งออก %s..."
            # จึงตรวจด้วยการเช็คขึ้นต้นด้วย "17กําลังส่งออก" หรือ "0กําลังส่งออก"
            exporting      = any(
                line.startswith('17กําลังส่งออก') or line.startswith('0กําลังส่งออก')
                for line in lines
            )

            # ต้องมี "ส่งออกสำเร็จ" และ "ไม่เหลือกำลังส่งออก" จึงถือว่าสำเร็จ
            if export_success and not exporting:
                print("พบข้อความ 'ส่งออกสำเร็จ' และไม่เหลือ 'กำลังส่งออก' แล้ว เริ่ม Auto_Download และประมวลผลไฟล์ต่อไป")
                time.sleep(5)  # รอ 5 วินาทีก่อนเริ่ม Auto_Download
                break

            # ถ้ามี "กำลังส่งออก" ไม่ว่าจะมี "ส่งออกสำเร็จ" หรือไม่ ให้รอ 5 นาที
            if exporting:
                print("ยังพบข้อความ 'กำลังส่งออก' -> รออีก 5 นาทีแล้วตรวจสอบใหม่")
                time.sleep(300)  # รอ 5 นาที
                continue

            check_attempt += 1
            print(f"ยังไม่พบข้อความ 'ส่งออกสำเร็จ' (ครั้งที่ {check_attempt}/10)")

            if check_attempt >= 10:
                print("ตรวจสอบครบ 10 ครั้งแล้วยังไม่พบ 'ส่งออกสำเร็จ' -> Alert และรอให้ผู้ใช้มาแก้ไข")
                Alert("หน้าต่าง Download ไม่สามารถใช้งานได้ (ตรวจไม่พบ 'ส่งออกสำเร็จ' ครบ 10 ครั้ง)")
                # วนรอให้ผู้ใช้มาแก้ไข
                while True:
                    print("รอให้ผู้ใช้มาแก้ไขหน้าต่าง Download ...")
                    time.sleep(5)

            print("ยังไม่พบข้อความ 'ส่งออกสำเร็จ' บนหน้าเว็บ -> กดปุ่ม (1785, 243) แล้วกดระฆังใหม่ จากนั้นจะตรวจสอบอีกครั้ง")
            pyautogui.click(x=1785, y=243)  # ปุ่มที่คุณระบุ
            time.sleep(5)
            pyautogui.click(x=1507, y=150)  # กดระฆังอีกครั้ง
            time.sleep(8)

    # มาถึงตรงนี้แปลว่า:
    #   - ไม่มีการสั่งส่งออกเลย (export_counter != 1) และข้ามการกดระฆัง
    #   หรือ
    #   - กดระฆังแล้ว และตรวจสอบพบ 'ส่งออกสำเร็จ'
    download_success, files_downloaded = Auto_Download(folder_path, Plate_Counter)
    process_all_excel_files()

    # บันทึกประวัติ plate ที่ดาวน์โหลดสำเร็จ (ทั้งรอบ)
    if len(plate_dowload) > 0:
        plate_dowload_df = pd.DataFrame(plate_dowload, columns=['Plate_Number'])
        plate_dowload_df.to_csv('plate_download.txt', index=False, sep='\t')
        plate_dowload_df.to_excel('plate_download.xlsx', index=0)
        plate_dowload_df.to_excel(folder_History, index=0)

    # เขียนสถานะสรุปว่ารอบนี้จบสมบูรณ์
    write_excel_download_status(Date_List, 1)
    close_chrome()
#------------------------------------------------------
#Plate_Num_Data        = pd.read_excel('Amita_plate.xlsx')
#Plate_Num_List        = Plate_Num_Data['Plate'].tolist()
#Date_Data             = pd.read_excel('Date_List_manual.xlsx')
#Date_List             = Date_Data['Date'].tolist()
#Date_List             = ['2025-11-30']
#Plate_Num_List        = ['16-6971','16-6593']
#TSB_Download(Plate_Num_List, Date_List)
#folder_path = "D:\Download"
#Auto_Download(folder_path,400)
#process_all_excel_files()
