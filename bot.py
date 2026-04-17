# bot.py
import os
import io
import base64
import requests
import pandas as pd
from datetime import datetime
from flask import Flask, request, abort, jsonify
from openpyxl import Workbook, load_workbook

# ====== Cấu hình ======
LOG_FILE = "bot_user_log.xlsx"
OUTLOOK_LINK = "https://1drv.ms/x/c/63897167e619733d/IQAAsw4pLS6ZQ46oKJfSgbmRASMpiNzmZcrm1cKRWGwB1Tc?e=cTvuRI"
TACT_LINK = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSxJMSJZcwlD4ZUiY0a_N1KfeAyKp2HDUGzhXWA1wDxRkU1fFCU3BjfQZnquOEtwA/pubhtml?gid=248455740&single=true"

# Env vars (bạn đã cấu hình RENDER_URL trên Render)
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
user_data = {}  # {user_id: {"step": "name"/"company"/"done", "name":..., "company":...}}

# ====== Flask app (health + webhook receiver) ======
app = Flask(__name__)

@app.route("/", methods=["GET", "HEAD"])
def health():
    return "Bot is alive!"

@app.route(f"/{WEBHOOK_PATH}", methods=["POST"])
def webhook_receiver():
    """
    Synchronous webhook receiver.
    Parse incoming update JSON and handle message synchronously,
    then reply via Telegram sendMessage API.
    """
    try:
        data = request.get_json(force=True)
    except Exception as e:
        print("Invalid JSON in webhook POST:", e)
        abort(400)

    # Only handle message updates (ignore other update types for now)
    if "message" not in data:
        # respond 200 so Telegram doesn't retry too aggressively
        return jsonify({"ok": True, "note": "no message"}), 200

    try:
        handle_update_sync(data)
    except Exception as e:
        # Log error but return 200 to Telegram to avoid retries
        print("Error handling update:", e)
    return jsonify({"ok": True}), 200

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

# ====== Telegram send helper (synchronous) ======
TELEGRAM_API = f"https://api.telegram.org/bot{TOKEN}"

def send_message(chat_id: int, text: str, parse_mode: str = None):
    payload = {"chat_id": chat_id, "text": text}
    if parse_mode:
        payload["parse_mode"] = parse_mode
        payload["disable_web_page_preview"] = True
    try:
        resp = requests.post(f"{TELEGRAM_API}/sendMessage", data=payload, timeout=15)
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        print("Failed to send message:", e, resp.text if 'resp' in locals() else "")
        return None

# ====== Core synchronous update handling (keeps same UX as before) ======
def handle_update_sync(update_json: dict):
    msg = update_json.get("message", {})
    if not msg:
        return

    chat = msg.get("chat", {})
    from_user = msg.get("from", {})
    chat_id = chat.get("id")
    user_id = from_user.get("id")
    text = msg.get("text", "").strip() if msg.get("text") else ""

    if not chat_id or not user_id:
        return

    # Commands
    if text.startswith("/start"):
        user_data[user_id] = {"step": "name"}
        send_message(chat_id, "Xin chào! Vui lòng nhập Tên của bạn:")
        return

    if text.startswith("/help"):
        help_text = (
            "📖 Hướng dẫn sử dụng bot:\n\n"
            "- /start: Bắt đầu trò chuyện với bot (yêu cầu nhập Tên và Công ty).\n"
            "- /list_dest: Liệt kê toàn bộ mã Dest trong cột B.\n"
            "- /help: Hiển thị hướng dẫn chi tiết.\n\n"
            "Sau khi nhập đủ thông tin, bạn có thể gõ trực tiếp mã Dest (ví dụ: SIN, CGK, BKK, KUL) "
            "để nhận thông tin chi tiết."
        )
        send_message(chat_id, help_text)
        return

    if text.startswith("/list_dest"):
        try:
            df = load_excel_from_onedrive(sheet_name=" ")
            dest_values = df[1].dropna().unique()
            dest_list = ", ".join(sorted(dest_values.astype(str)))
            answer = f"📋 Danh sách tất cả Dest trong cột B:\n{dest_list}"
        except Exception as e:
            answer = f"⚠️ Có lỗi xảy ra khi đọc file: {e}"
        send_message(chat_id, answer)
        return

    # Conversation flow: name -> company -> done
    state = user_data.get(user_id, {})
    step = state.get("step")

    if step == "name":
        # Save name, ask company
        user_data[user_id]["name"] = text
        user_data[user_id]["step"] = "company"
        send_message(chat_id, "Cảm ơn! Bây giờ hãy nhập Tên công ty:")
        return

    if step == "company":
        user_data[user_id]["company"] = text
        user_data[user_id]["step"] = "done"
        send_message(chat_id, "✅ Đã lưu thông tin. Giờ bạn có thể nhập mã Dest để tra cứu.")
        return

    if step != "done":
        send_message(chat_id, "⚠️ Vui lòng nhập Tên và Công ty trước bằng lệnh /start.")
        return

    # Now user is in 'done' state: treat text as Dest query
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_log(user_id, user_data[user_id].get("name", ""), user_data[user_id].get("company", ""), text, timestamp)

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

    send_message(chat_id, answer, parse_mode="HTML")

# ====== Helper: set webhook via Telegram API (synchronous) ======
def set_telegram_webhook(webhook_url: str):
    url = f"https://api.telegram.org/bot{TOKEN}/setWebhook"
    resp = requests.post(url, data={"url": webhook_url}, timeout=15)
    resp.raise_for_status()
    j = resp.json()
    if not j.get("ok"):
        raise RuntimeError(f"setWebhook returned not ok: {j}")
    return j

# ====== Entrypoint ======
if __name__ == "__main__":
    webhook_url = RENDER_URL.rstrip("/") + f"/{WEBHOOK_PATH}"
    print("Webhook URL will be:", webhook_url)

    # Try to set webhook (best-effort)
    try:
        res = set_telegram_webhook(webhook_url)
        print("✅ Telegram webhook set to:", webhook_url)
    except Exception as e:
        print("⚠️ setWebhook failed (will still run):", e)

    # Start Flask (this will receive webhook POSTs and health checks)
    # Use Flask dev server here; Render maps external port to container port.
    print("Starting Flask on port", PORT)
    app.run(host="0.0.0.0", port=PORT)
