from flask import Flask, request, jsonify
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import traceback
import re
import os

app = Flask(__name__)

CHANNEL_ACCESS_TOKEN = 'SpWk6UUGZwG8nCf4ch6DL4fnSVU9NTHHeCsPobNuuprT2t+/FEkTY0Z7IEhR2sNDfSTlUuCHIVPu+eZL3NdPmJfceJqc6WK7zzpY0SrCXk0+AUOqXkYx8zgwoFEQ7FtFABvECsXpFQnYkkydnjf66AdB04t89/1O/w1cDnyilFU='

LINE_HEADERS = {
    'Content-Type': 'application/json',
    'Authorization': f'Bearer {CHANNEL_ACCESS_TOKEN}'
}

# 🔁 เก็บ state ชั่วคราวของผู้ใช้แต่ละคน
user_states = {}

@app.route("/", methods=["GET"])
def home():
    return "✅ LINE Helpdesk is running.", 200

@app.route("/webhook", methods=["POST"])
def webhook():
    try:
        payload = request.json
        print("📥 Received payload:", payload)
        events = payload.get('events', [])

        for event in events:
            if event.get('type') == 'message' and event['message'].get('type') == 'text':
                user_message = event['message']['text'].strip()
                reply_token = event['replyToken']
                user_id = event['source']['userId']

                print(f"🗣️ User {user_id} sent: {user_message}")
                
                reset_keywords = ["แจ้งปัญหา", "เช็กสถานะ", "ติดต่อเจ้าหน้าที่", "ยกเลิก"]
                if user_id in user_states and any(user_message.startswith(k) for k in reset_keywords):
                    del user_states[user_id]

                # 🧠 ตรวจสอบว่าผู้ใช้กำลังอยู่ใน flow หรือไม่
                if user_id in user_states:
                    state = user_states[user_id]
                    step = state.get("step")

                    if step == "ask_issue":
                        state["issue"] = user_message
                        state["step"] = "ask_category"
                        reply(reply_token, "📂 กรุณาบอกชื่อผู้ใช้")

                    elif step == "ask_category":
                        state["category"] = user_message
                        state["step"] = "ask_phone"
                        reply(reply_token, "📞 กรุณาบอกเบอร์ติดต่อกลับ")

                    elif step == "ask_phone":
                        phone = user_message
                        if not re.fullmatch(r"0\d{9}", phone):
                            reply(reply_token, "⚠️ กรุณาระบุเบอร์ติดต่อ 10 หลักให้ถูกต้อง เช่น 0812345678")
                            continue

                        state["phone"] = phone
                        ticket_id = generate_ticket_id()
                        success = save_ticket_to_sheet(user_id, state, ticket_id)
                        if success:
                            reply(reply_token, f"✅ แจ้งปัญหาเรียบร้อยแล้วค่ะ\nหมายเลข Ticket ของคุณคือ: {ticket_id}")
                        else:
                            reply(reply_token, "❌ เกิดปัญหาในการบันทึกข้อมูลลง Google Sheet")
                        del user_states[user_id]

                    continue  # ข้ามการประมวลผลด้านล่าง

                # 🔘 คำสั่งหลัก
                if user_message == "แจ้งปัญหา":
                    user_states[user_id] = {"step": "ask_issue"}
                    reply(reply_token, "📝 กรุณาบอกปัญหาที่พบ")
                    
                elif user_message == "ยกเลิก":
                    if user_id in user_states:
                        del user_states[user_id]
                        reply(reply_token, "❎ ยกเลิกการแจ้งปัญหาเรียบร้อยแล้ว")
                    else:
                        reply(reply_token, "❎ ยกเลิกการแจ้งปัญหาเรียบร้อยแล้ว")

                elif user_message == "เช็กสถานะ":
                    reply(reply_token, "🔍 กรุณาพิมพ์หมายเลข Ticket ของคุณ เช่น TICKET-20250521123000")

                elif user_message == "ติดต่อเจ้าหน้าที่":
                    reply(reply_token, "📞 ติดต่อเจ้าหน้าที่ได้ที่เบอร์ 02-xxx-xxxx หรืออีเมล support@example.com")

                elif user_message.startswith("แจ้งปัญหา"):
                    parsed = parse_issue_message(user_message)
                    if parsed:
                        ticket_id = generate_ticket_id()
                        success = save_ticket_to_sheet(user_id, parsed, ticket_id)
                        if success:
                            reply(reply_token, f"✅ แจ้งปัญหาเรียบร้อยแล้วค่ะ\nหมายเลข Ticket ของคุณคือ: {ticket_id}")
                        else:
                            reply(reply_token, "❌ เกิดปัญหาในการบันทึกข้อมูลลง Google Sheet")
                    else:
                        reply(reply_token, "⚠️ กรุณาระบุข้อมูลให้ครบถ้วน:\n\nแจ้งปัญหา: ...\nประเภท: ...\nเบอร์ติดต่อ: ...")

                elif re.search(r"TICKET-\d{14}", user_message):
                    match = re.search(r"(TICKET-\d{14})", user_message)
                    ticket_id = match.group(1)
                    status = check_ticket_status(ticket_id)
                    if status:
                        reply(reply_token, f"📄 สถานะของ {ticket_id}:{status}")
                    else:
                        reply(reply_token, "❌ ไม่พบ Ticket นี้ในระบบ")

                else:
                    reply(reply_token, "📌 ไปที่เมนูและเลือกบริการ")

        return jsonify({"status": "ok"}), 200

    except Exception as e:
        print("❌ ERROR in webhook():", e)
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500

def reply(reply_token, text):
    body = {
        "replyToken": reply_token,
        "messages": [{"type": "text", "text": text}]
    }
    try:
        res = requests.post('https://api.line.me/v2/bot/message/reply', headers=LINE_HEADERS, json=body)
        print("📤 Reply response:", res.status_code, res.text)
    except Exception as e:
        print("❌ Failed to reply:", e)
        traceback.print_exc()

# 🔍 แยกข้อความแจ้งปัญหา
def parse_issue_message(message):
    try:
        issue = re.search(r"แจ้งปัญหา[:：]\s*(.*)", message)
        category = re.search(r"ประเภท[:：]\s*(.*)", message)
        phone = re.search(r"เบอร์ติดต่อ[:：]\s*(.*)", message)
        if issue and category and phone:
            return {
                "issue": issue.group(1).strip(),
                "category": category.group(1).strip(),
                "phone": phone.group(1).strip()
            }
        else:
            return None
    except:
        return None

# 🧾 สร้างเลข Ticket
def generate_ticket_id():
    now = datetime.now()
    return f"TICKET-{now.strftime('%Y%m%d%H%M%S')}"

# 📤 บันทึกข้อมูลลง Google Sheet
def save_ticket_to_sheet(user_id, data, ticket_id):
    try:
        scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(creds)
        sheet = client.open("Tickets").sheet1

        sheet.append_row([
        ticket_id,
        user_id,
        data['issue'],
        data['category'],
        data['phone'],  # ✅ ใส่เบอร์ตรง ๆ ไม่มี quote
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "รอดำเนินการ"
        ], value_input_option='RAW')  

        print(f"✅ Ticket {ticket_id} saved")
        return True
    except Exception as e:
        print("❌ Error saving ticket:", e)
        traceback.print_exc()
        return False

def check_ticket_status(ticket_id):
    try:
        scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(creds)
        sheet = client.open("Tickets").sheet1
        data = sheet.get_all_records()

        for row in data:
            if row.get('หมายเลข Ticket') == ticket_id or row.get('Ticket ID') == ticket_id:
                return (
                    f"\n📋 สถานะ \n"
                    f"ปัญหา: {row.get('รายละเอียด') or row.get('issue') or '-'}\n"
                    f"ชื่อ: {row.get('ชื่อ') or row.get('category') or '-'}\n"
                    f"เบอร์ติดต่อ: {display_phone_number(row.get('เบอร์ติดต่อ') or row.get('phone'))}\n"
                    f"สถานะ: {row.get('สถานะ') or row.get('status') or '-'}"
                )

        return "❌ ไม่พบหมายเลข Ticket นี้ในระบบ"
    except Exception as e:
        print("❌ Error checking status:", e)
        traceback.print_exc()
        return "❌ เกิดข้อผิดพลาดในการค้นหาข้อมูล"
    
def display_phone_number(phone):
    try:
        phone_str = str(phone).strip().replace("'", "").replace('"', "")

        if phone_str.startswith("66") and len(phone_str) == 11:
            return "0" + phone_str[2:]

        if len(phone_str) == 9 and not phone_str.startswith("0"):
            return "0" + phone_str

        return phone_str
    except:
        return "-"


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    LINE_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
    GOOGLE_CREDS_PATH = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    app.run(host="0.0.0.0", port=port, debug=True)


