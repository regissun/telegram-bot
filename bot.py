import os
import io
import re
import requests
import pandas as pd
from datetime import datetime
from flask import Flask, request
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters
from openpyxl import Workbook, load_workbook

# ====== Cấu hình ======
LOG_FILE = "bot_user_log.xlsx"
OUTLOOK_LINK = "https://1drv.ms/x/c/63897167e619733d/IQAAsw4pLS6ZQ46oKJfSgbmRASMpiNzmZcrm1cKRWGwB1Tc?e=cTvuRI"
TACT_LINK = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSxJMSJZcwlD4ZUiY0a_N1KfeAyKp2HDUGzhXWA1wDxRkU1fFCU3BjfQZnquOEtwA/pubhtml?gid=248455740&single=true"

TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
ONEDRIVE_SHORT_URL = os.getenv("ONEDRIVE_LINK")

user_data = {}

# ====== Hàm xử lý OneDrive ======
def get_direct_onedrive_link(short_url: str) -> str:
    resp = requests.get(short_url, allow_redirects=True)
    final_url = resp.url
    match = re.search(r"resid=([^&]+)", final_url)
    if not match:
        raise ValueError("Không tìm thấy resid trong link OneDrive")
    resid = match.group(1)
    return f"https://onedrive.live.com/download?resid={resid}"

DIRECT_URL = get_direct_onedrive_link(ONEDRIVE_SHORT_URL)

def load_excel_from_onedrive(sheet_name=" "):  # sheet name là dấu cách
    response = requests.get(DIRECT_URL)
    response.raise_for_status()
    return pd.read_excel(io.BytesIO(response.content), sheet_name=sheet_name, header=None)

# ====== Ghi log ======
def save_log(user_id, name, company, question, timestamp):
    try:
        try:
            wb = load_workbook(LOG_FILE)
            ws = wb.active
        except FileNotFoundError:
            wb = Workbook()
            ws = wb.active
            ws.append(["User ID", "Tên", "Công ty", "Câu hỏi", "Thời gian"])
        ws.append([user_id, name, company, question, timestamp])
        wb.save(LOG_FILE)
    except Exception as e:
        print(f"❌ Lỗi ghi log: {e}")

# ====== Các lệnh bot ======
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    user_data[user_id] = {"step": "name"}
    await update.message.reply_text("Xin chào! Vui lòng nhập Tên của bạn:")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    answer = (
        "📖 Hướng dẫn sử dụng bot:\n\n"
        "- /start: Bắt đầu trò chuyện với bot (yêu cầu nhập Tên và Công ty).\n"
        "- /list_dest: Liệt kê toàn bộ mã Dest trong cột B.\n"
        "- /help: Hiển thị hướng dẫn chi tiết.\n\n"
        "Sau khi nhập đủ thông tin, bạn có thể gõ trực tiếp mã Dest (ví dụ: SIN, CGK, BKK, KUL) "
        "để nhận thông tin chi tiết."
    )
    await update.message.reply_text(answer)

async def list_dest(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        df = load_excel_from_onedrive(sheet_name=" ")
        dest_values = df[1].dropna().unique()
        dest_list = ", ".join(sorted(dest_values.astype(str)))
        answer = f"📋 Danh sách tất cả Dest trong cột B:\n{dest_list}"
    except Exception as e:
        answer = f"⚠️ Có lỗi xảy ra khi đọc file: {e}"
    await update.message.reply_text(answer)

async def reply(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    text = update.message.text.strip()

    if user_id in user_data and user_data[user_id].get("step") == "name":
        user_data[user_id]["name"] = text
        user_data[user_id]["step"] = "company"
        await update.message.reply_text("Cảm ơn! Bây giờ hãy nhập Tên công ty:")
        return

    if user_id in user_data and user_data[user_id].get("step") == "company":
        user_data[user_id]["company"] = text
        user_data[user_id]["step"] = "done"
        await update.message.reply_text("✅ Đã lưu thông tin. Giờ bạn có thể nhập mã Dest để tra cứu.")
        return

    if user_id not in user_data or user_data[user_id].get("step") != "done":
        await update.message.reply_text("⚠️ Vui lòng nhập Tên và Công ty trước bằng lệnh /start.")
        return

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_log(user_id, user_data[user_id]["name"], user_data[user_id]["company"], text, timestamp)

    try:
        df = load_excel_from_onedrive(sheet_name=" ")
        row = df[df[1].astype(str).str.strip().str.lower() == text.lower()]

        if not row.empty:
            r = row.iloc[0]
            c_val = f"{r[2]:.2f}" if pd.notnull(r[2]) else ""
            i_val = str(r[8]) if pd.notnull(r[8]) else ""
            j_val = str(r[9]) if pd.notnull(r[9]) else ""
            k_val = pd.to_datetime(r[10], errors="coerce")
            k_val = k_val.strftime("%d/%m/%Y") if pd.notnull(k_val) else ""
            l_val = pd.to_datetime(r[11], errors="coerce")
            l_val = l_val.strftime("%d/%m/%Y") if pd.notnull(l_val) else ""

            extra_text = "\n".join(
                str(df.iloc[i, 13]) for i in range(1, 6) if pd.notnull(df.iloc[i, 13])
            )

            answer = (
                f"📊 Kết quả cho Dest: {text.upper()}\n"
                f"- Giá All-in USD/kg: {c_val}\n"
                f"- Điều kiện 1: {i_val}\n"
                f"- Điều kiện 2: {j_val}\n"
                f"- Valid from: {k_val}\n"
                f"- Valid till: {l_val}\n\n"
                f"📌 Thông tin bổ sung:\n{extra_text}\n\n"
                f"🔗 <a href='{OUTLOOK_LINK}'>Space Outlook</a>\n"
                f"🔗 <a href='{TACT_LINK}'>TACT Rate</a>"
            )
        else:
            answer = (
                "Xin lỗi, chưa có dữ liệu cho giá trị này.\n"
                "👉 Bạn có thể dùng lệnh /list_dest để xem danh sách Dest có sẵn.\n\n"
                f"🔗 <a href='{OUTLOOK_LINK}'>Space Outlook</a>\n"
                f"🔗 <a href='{TACT_LINK}'>TACT Rate</a>"
            )

    except Exception as e:
        answer = f"⚠️ Có lỗi xảy ra: {e}"

    await update.message.reply_text(answer, parse_mode="HTML")

# ====== Flask + Webhook ======
app = Flask(__name__)
application = ApplicationBuilder().token(TOKEN).build()
application.add_handler(CommandHandler("start", start))
application.add_handler(CommandHandler("help", help_command))
application.add_handler(CommandHandler("list_dest", list_dest))
application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, reply))

@app.route("/")
def home():
    return "Bot is alive!"

@app.route("/webhook", methods=["POST"])
def webhook():
    update = Update.de_json(request.get_json(force=True), application.bot)
    application.update_queue.put_nowait(update)  # dùng put_nowait để tránh lỗi async
    print(f"📩 Nhận update: {update.to_dict()}")
    return "ok"

if __name__ == "__main__":
    application.initialize()
    application.start()
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
    application.stop()
