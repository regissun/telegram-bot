# bot.py
import os
import io
import base64
import requests
import pandas as pd
import asyncio
import threading
from datetime import datetime
from flask import Flask, request, abort
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)
from openpyxl import Workbook, load_workbook

# ====== Cấu hình ======
LOG_FILE = "bot_user_log.xlsx"
OUTLOOK_LINK = "https://1drv.ms/x/c/63897167e619733d/IQAAsw4pLS6ZQ46oKJfSgbmRASMpiNzmZcrm1cKRWGwB1Tc?e=cTvuRI"
TACT_LINK = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSxJMSJZcwlD4ZUiY0a_N1KfeAyKp2HDUGzhXWA1wDxRkU1fFCU3BjfQZnquOEtwA/pubhtml?gid=248455740&single=true"

# Env vars
TOKEN = os.getenv("TOKEN") or os.getenv("TELEGRAM_BOT_TOKEN")
ONEDRIVE_URL = os.getenv("ONEDRIVE_URL") or os.getenv("ONEDRIVE_LINK")
RENDER_URL = os.getenv("RENDER_URL")  # e.g. "https://telegram-bot-o4h8.onrender.com"
PORT = int(os.environ.get("PORT", 10000))
WEBHOOK_PATH = "webhook"  # final path: /webhook

if not TOKEN:
    raise RuntimeError("Missing TELEGRAM token. Set environment variable TOKEN.")
if not ONEDRIVE_URL:
    raise RuntimeError("Missing ONEDRIVE_URL environment variable.")
if not RENDER_URL:
    raise RuntimeError("Missing RENDER_URL environment variable.")

# ====== Global state ======
application = None
user_data = {}

# ====== Flask app for webhook + health-check ======
flask_app = Flask(__name__)

@flask_app.route("/", methods=["GET", "HEAD"])
def health():
    return "Bot is alive!"

@flask_app.route(f"/{WEBHOOK_PATH}", methods=["POST"])
def webhook_receiver():
    """
    Receive Telegram webhook POST from Telegram (synchronous Flask route).
    Convert JSON to Update and push into application's update_queue synchronously.
    """
    global application
    if application is None:
        # Not ready yet
        abort(503, "Bot not ready")
    try:
        data = request.get_json(force=True)
        update = Update.de_json(data, application.bot)
        # put_nowait is synchronous and safe here
        application.update_queue.put_nowait(update)
        return "ok"
    except Exception as e:
        # Return 200 to Telegram if you prefer, but log error
        print("Error handling webhook POST:", e)
        abort(400, str(e))

def run_flask():
    # Run Flask dev server only for health and webhook receiving.
    # Render will map external port to container port.
    flask_app.run(host="0.0.0.0", port=PORT)

# ====== OneDrive helper (sheet name is single space " ") ======
def get_direct_link(share_url: str) -> str:
    """
    Convert OneDrive share URL to direct download via onedrive API share token.
    Uses base64 'u!' encoding approach.
    """
    encoded = base64.b64encode(share_url.encode()).decode()
    encoded = encoded.rstrip("=").replace("/", "_").replace("+", "-")
    return f"https://api.onedrive.com/v1.0/shares/u!{encoded}/root/content"

DIRECT_URL = get_direct_link(ONEDRIVE_URL)

def load_excel_from_onedrive(sheet_name=" "):
    """
    Download workbook and read the sheet named single-space " ".
    Returns pandas DataFrame with header=None.
    """
    resp = requests.get(DIRECT_URL, timeout=30)
    resp.raise_for_status()
    return pd.read_excel(io.BytesIO(resp.content), sheet_name=sheet_name, header=None)

# ====== Ghi log vào Excel cục bộ ======
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

# ====== Handlers ======
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data[user_id] = {"step": "name"}
    await update.message.reply_text("Xin chào! Vui lòng nhập Tên của bạn:")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    answer = (
        "📖 Hướng dẫn sử dụng bot:\n\n"
        "- /start: Bắt đầu (yêu cầu nhập Tên và Công ty).\n"
        "- /list_dest: Liệt kê toàn bộ mã Dest trong cột B.\n"
        "- /help: Hiển thị hướng dẫn.\n\n"
        "Sau khi nhập Tên và Công ty, gõ mã Dest (ví dụ: SIN, CGK) để tra cứu."
    )
    await update.message.reply_text(answer)

async def list_dest(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        df = load_excel_from_onedrive(sheet_name=" ")
        dest_values = df[1].dropna().unique()
        dest_list = ", ".join(sorted(dest_values.astype(str)))
        answer = f"📋 Danh sách tất cả Dest trong cột B:\n{dest_list}"
    except Exception as e:
        answer = f"⚠️ Lỗi khi đọc file: {e}"
    await update.message.reply_text(answer)

async def reply(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = (update.message.text or "").strip()
    if not text:
        return

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
        answer = f"⚠️ Có lỗi xảy ra khi tra cứu: {e}"

    await update.message.reply_text(answer, parse_mode="HTML")

# ====== Build application ======
def build_application():
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("list_dest", list_dest))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, reply))
    return app

# ====== Helper: set webhook via Telegram API ======
def set_telegram_webhook(webhook_url: str):
    """
    Call Telegram setWebhook synchronously. Raise on failure.
    """
    url = f"https://api.telegram.org/bot{TOKEN}/setWebhook"
    resp = requests.post(url, data={"url": webhook_url}, timeout=15)
    try:
        resp.raise_for_status()
    except Exception as e:
        raise RuntimeError(f"Failed to setWebhook: {e} - {resp.text}")
    j = resp.json()
    if not j.get("ok"):
        raise RuntimeError(f"setWebhook returned not ok: {j}")

# ====== Entrypoint: run application in asyncio main thread and Flask in background thread ======
async def main_async():
    global application
    application = build_application()

    # compute webhook URL
    webhook_url = RENDER_URL.rstrip("/") + f"/{WEBHOOK_PATH}"
    print("Webhook URL will be:", webhook_url)

    # Try to set webhook (best-effort). If fails, print error but continue.
    try:
        set_telegram_webhook(webhook_url)
        print("✅ Telegram webhook set to:", webhook_url)
    except Exception as e:
        print("⚠️ setWebhook failed:", e)

    # Initialize and start application (registers handlers and starts dispatcher)
    await application.initialize()
    await application.start()
    print("✅ Application initialized and started")

    # Start Flask server in background thread to receive webhook POSTs and health checks
    flask_thread = threading.Thread(target=run_flask, daemon=True)
    flask_thread.start()
    print(f"✅ Flask started on port {PORT} (health + webhook receiver)")

    # Keep running until cancelled
    try:
        await asyncio.Event().wait()
    finally:
        # graceful shutdown
        await application.stop()
        await application.shutdown()
        print("Application stopped")

if __name__ == "__main__":
    try:
        asyncio.run(main_async())
    except KeyboardInterrupt:
        print("Interrupted, exiting")
    except Exception as e:
        print("Fatal error in main:", e)
