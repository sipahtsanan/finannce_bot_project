"""Finance Tracker Bot"""

import os
import re
import json
import logging
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters
)
import google.generativeai as genai
from excel_manager import ExcelManager
import base64

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "YOUR_BOT_TOKEN_HERE")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY_HERE")
EXCEL_FILE = "finance_tracker.xlsx"

CATEGORIES = [
    "อาหาร", "เดินทาง", "ที่พัก", "สุขภาพ",
    "ช้อปปิ้ง", "บันเทิง", "โทรศัพท์", "ค่าน้ำไฟ",
    "การศึกษา", "ออมลงทุน", "รายได้", "อื่นๆ",
]

genai.configure(api_key=GEMINI_API_KEY)
excel = ExcelManager(EXCEL_FILE)
TODAY = lambda: datetime.now().strftime("%Y-%m-%d")


def parse_json_response(text: str):
    text = re.sub(r"```json\s*", "", text.strip())
    text = re.sub(r"```\s*", "", text).strip()
    return json.loads(text)


def analyze_slip(image_bytes: bytes) -> list:
    model = genai.GenerativeModel("gemini-2.5-flash")
    today = TODAY()
    prompt = (
        "Analyze this document and return ONLY a JSON array. "
        "Include ALL transactions found. No extra text. "
        "IMPORTANT: If no date is shown in the document, use " + today + " as the date. "
        'Format: [{"date":"YYYY-MM-DD","amount":number,"description":"detail",'
        '"merchant":"name","transaction_type":"expense or income",'
        '"suggested_category":"อาหาร/เดินทาง/ที่พัก/สุขภาพ/ช้อปปิ้ง/บันเทิง/โทรศัพท์/ค่าน้ำไฟ/การศึกษา/ออมลงทุน/อื่นๆ"}]'
    )
    image_part = {"mime_type": "image/jpeg", "data": base64.b64encode(image_bytes).decode()}
    response = model.generate_content([prompt, image_part])
    result = parse_json_response(response.text)
    items = result if isinstance(result, list) else [result]
    # ถ้าวันที่เป็น null หรือ default ที่ผิดพลาด ให้ใช้วันนี้แทน
    for item in items:
        d = item.get("date", "")
        if not d or d < "2020-01-01":
            item["date"] = today
    return items


def build_category_keyboard():
    keyboard = []
    row = []
    for cat in CATEGORIES:
        row.append(InlineKeyboardButton(cat, callback_data="cat:" + cat))
        if len(row) == 2:
            keyboard.append(row)
            row = []
    if row:
        keyboard.append(row)
    keyboard.append([InlineKeyboardButton("X ยกเลิก", callback_data="cat:cancel")])
    return InlineKeyboardMarkup(keyboard)


def make_pending(data: dict) -> dict:
    date_str = data.get("date") or TODAY()
    if date_str < "2020-01-01":
        date_str = TODAY()
    amount = abs(float(data.get("amount", 0)))
    desc = data.get("description", "")
    merchant = data.get("merchant", "")
    return {
        "date": date_str,
        "amount": amount,
        "description": (merchant + " - " + desc).strip(" -"),
        "transaction_type": data.get("transaction_type", "expense"),
        "suggested": data.get("suggested_category", "")
    }


async def show_category_prompt(reply_func, context, data: dict, prefix=""):
    pending = make_pending(data)
    context.user_data["pending"] = pending
    msg = (prefix +
           "ข้อมูล:\n" +
           "วันที่: " + pending["date"] + "\n" +
           "จำนวน: " + f"{pending['amount']:,.2f}" + " บาท\n" +
           "รายละเอียด: " + pending["description"] + "\n" +
           "AI แนะนำหมวด: " + pending["suggested"] + "\n\n" +
           "เลือกหมวดหมู่:")
    await reply_func(msg, reply_markup=build_category_keyboard())


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = ("สวัสดี! Finance Tracker Bot\n\n"
           "วิธีใช้:\n"
           "- ส่งรูป slip หรือ PDF\n"
           "- พิมพ์ เช่น กาแฟ 80 หรือ เงินเดือน 50000\n\n"
           "คำสั่ง:\n"
           "/list - รายการล่าสุด + ลบ\n"
           "/edit - แก้ไขราคาและวันที่\n"
           "/summary - สรุปเดือนนี้\n"
           "/yearly - สรุปรายปี\n"
           "/export - ดาวน์โหลด Excel")
    await update.message.reply_text(msg)


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("กำลังอ่าน slip...")
    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)
    image_bytes = await file.download_as_bytearray()
    try:
        items = analyze_slip(bytes(image_bytes))
        context.user_data["queue"] = items[1:]
        await show_category_prompt(update.message.reply_text, context, items[0])
    except Exception as e:
        logger.error("Photo error: " + str(e))
        await update.message.reply_text("อ่าน slip ไม่ได้ ลองพิมพ์เองได้เลย เช่น กาแฟ 80")


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if doc.mime_type != "application/pdf":
        await update.message.reply_text("รองรับเฉพาะ PDF หรือรูปภาพ")
        return
    await update.message.reply_text("กำลังอ่าน PDF...")
    file = await context.bot.get_file(doc.file_id)
    pdf_bytes = await file.download_as_bytearray()
    try:
        import fitz
        pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        pix = pdf_doc[0].get_pixmap(dpi=200)
        items = analyze_slip(pix.tobytes("jpeg"))
        context.user_data["queue"] = items[1:]
        total = len(items)
        prefix = ("พบ " + str(total) + " รายการ กรุณาเลือกหมวดหมู่ทีละรายการ\n\n") if total > 1 else ""
        await show_category_prompt(update.message.reply_text, context, items[0], prefix)
    except ImportError:
        await update.message.reply_text("ต้องติดตั้ง PyMuPDF: pip3 install pymupdf")
    except Exception as e:
        logger.error("PDF error: " + str(e))
        await update.message.reply_text("อ่าน PDF ไม่ได้ ลองส่งเป็นรูปภาพแทน")


async def handle_category_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    try:
        await query.answer()
    except Exception:
        pass

    if query.data == "cat:cancel":
        context.user_data.pop("queue", None)
        context.user_data.pop("pending", None)
        await query.edit_message_text("ยกเลิกแล้ว")
        return

    category = query.data.replace("cat:", "")
    pending = context.user_data.get("pending", {})
    if not pending:
        await query.edit_message_text("ไม่พบข้อมูล ลองส่งใหม่")
        return

    excel.add_transaction(
        date=pending["date"],
        amount=pending["amount"],
        category=category,
        description=pending["description"],
        transaction_type=pending["transaction_type"]
    )
    context.user_data.pop("pending", None)

    saved_msg = ("บันทึกแล้ว: " + pending["description"] +
                 " | " + f"{pending['amount']:,.2f}" + " บาท | " + category)
    await query.edit_message_text(saved_msg)

    queue = context.user_data.get("queue", [])
    if queue:
        next_item = queue.pop(0)
        context.user_data["queue"] = queue
        remaining = len(queue)
        prefix = ("(เหลืออีก " + str(remaining) + " รายการ)\n\n") if remaining > 0 else ""
        await show_category_prompt(query.message.reply_text, context, next_item, prefix)
    else:
        context.user_data.pop("queue", None)


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # edit mode: รับราคาหรือวันที่ใหม่
    if context.user_data.get("editing_row"):
        await handle_edit_input(update, context)
        return

    text = update.message.text.strip()
    model = genai.GenerativeModel("gemini-2.5-flash")
    today = TODAY()
    prompt = ('Parse this text: "' + text + '"\n'
              "Return ONLY JSON (no markdown):\n"
              '{"amount":number,"description":"detail","merchant":"",'
              '"date":"' + today + '","transaction_type":"expense or income","suggested_category":"category"}')
    try:
        response = model.generate_content(prompt)
        result = parse_json_response(response.text)
        data = result[0] if isinstance(result, list) else result
        if not data.get("date") or data["date"] < "2020-01-01":
            data["date"] = today
        context.user_data["queue"] = []
        await show_category_prompt(update.message.reply_text, context, data)
    except Exception as e:
        logger.error("handle_text error: " + str(e))
        await update.message.reply_text("ไม่เข้าใจ ลองใหม่ เช่น กาแฟ 80 หรือ เงินเดือน 30000")


async def list_recent(update: Update, context: ContextTypes.DEFAULT_TYPE):
    rows = excel.get_recent_transactions(10)
    if not rows:
        await update.message.reply_text("ยังไม่มีข้อมูล")
        return
    keyboard = []
    lines = ["10 รายการล่าสุด:\n"]
    for i, row in enumerate(rows):
        row_num, date, cat, desc, ttype, amount = row
        sign = "+" if ttype == "รายได้" else "-"
        lines.append(str(i+1) + ". " + str(date) + " | " + str(desc) + " | " + sign + f"{abs(amount):,.0f}" + " | " + str(cat))
        keyboard.append([InlineKeyboardButton("ลบ " + str(i+1) + ": " + str(desc)[:15], callback_data="del:" + str(row_num))])
    keyboard.append([InlineKeyboardButton("ปิด", callback_data="del:close")])
    await update.message.reply_text("\n".join(lines), reply_markup=InlineKeyboardMarkup(keyboard))


async def handle_delete_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    try:
        await query.answer()
    except Exception:
        pass
    if query.data == "del:close":
        await query.edit_message_reply_markup(reply_markup=None)
        return
    row_num = int(query.data.replace("del:", ""))
    excel.delete_transaction(row_num)
    await query.edit_message_text("ลบรายการแล้ว\nพิมพ์ /list เพื่อดูรายการที่เหลือ")


async def list_for_edit(update: Update, context: ContextTypes.DEFAULT_TYPE):
    rows = excel.get_recent_transactions(10)
    if not rows:
        await update.message.reply_text("ยังไม่มีข้อมูล")
        return
    keyboard = []
    lines = ["เลือกรายการที่ต้องการแก้ไข:\n"]
    for i, row in enumerate(rows):
        row_num, date, cat, desc, ttype, amount = row
        sign = "+" if ttype == "รายได้" else "-"
        lines.append(str(i+1) + ". " + str(date) + " | " + str(desc) + " | " + sign + f"{abs(amount):,.0f}" + " | " + str(cat))
        keyboard.append([InlineKeyboardButton("แก้ " + str(i+1) + ": " + str(desc)[:15],
            callback_data="editrow:" + str(row_num) + ":" + str(date) + ":" + f"{abs(amount):.0f}")])
    keyboard.append([InlineKeyboardButton("ปิด", callback_data="editrow:close")])
    await update.message.reply_text("\n".join(lines), reply_markup=InlineKeyboardMarkup(keyboard))


async def handle_edit_select_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    try:
        await query.answer()
    except Exception:
        pass
    if query.data == "editrow:close":
        await query.edit_message_reply_markup(reply_markup=None)
        return
    # format: editrow:ROW_NUM:DATE:AMOUNT
    parts = query.data.replace("editrow:", "").split(":")
    row_num = int(parts[0])
    old_date = parts[1]
    old_amount = parts[2]
    context.user_data["editing_row"] = row_num
    context.user_data["editing_step"] = "amount"
    context.user_data["editing_date"] = old_date
    await query.edit_message_text(
        "แก้ไขรายการ\n" +
        "วันที่เดิม: " + old_date + "\n" +
        "ราคาเดิม: " + old_amount + " บาท\n\n" +
        "พิมพ์ราคาใหม่ได้เลย (หรือพิมพ์ skip เพื่อข้าม)"
    )


async def handle_edit_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    row_num = context.user_data.get("editing_row")
    step = context.user_data.get("editing_step", "amount")

    if step == "amount":
        if text.lower() == "skip":
            context.user_data["editing_step"] = "date"
            await update.message.reply_text(
                "ข้ามการแก้ราคา\n\nพิมพ์วันที่ใหม่ รูปแบบ YYYY-MM-DD เช่น " + TODAY() +
                "\n(หรือพิมพ์ skip เพื่อข้าม)"
            )
        else:
            try:
                new_amount = float(text.replace(",", ""))
                excel.update_amount(row_num, new_amount)
                context.user_data["editing_step"] = "date"
                await update.message.reply_text(
                    "แก้ราคาเป็น " + f"{new_amount:,.2f}" + " บาทแล้ว\n\n" +
                    "พิมพ์วันที่ใหม่ รูปแบบ YYYY-MM-DD เช่น " + TODAY() +
                    "\n(หรือพิมพ์ skip เพื่อข้าม)"
                )
            except ValueError:
                await update.message.reply_text("กรุณาพิมพ์ตัวเลขอย่างเดียว เช่น 150 หรือพิมพ์ skip")

    elif step == "date":
        if text.lower() == "skip":
            context.user_data.pop("editing_row", None)
            context.user_data.pop("editing_step", None)
            context.user_data.pop("editing_date", None)
            await update.message.reply_text("แก้ไขเสร็จแล้ว!")
        else:
            try:
                datetime.strptime(text, "%Y-%m-%d")
                excel.update_date(row_num, text)
                context.user_data.pop("editing_row", None)
                context.user_data.pop("editing_step", None)
                context.user_data.pop("editing_date", None)
                await update.message.reply_text("แก้วันที่เป็น " + text + " แล้ว!")
            except ValueError:
                await update.message.reply_text("รูปแบบวันที่ไม่ถูกต้อง ต้องเป็น YYYY-MM-DD เช่น " + TODAY())


async def summary(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(excel.get_monthly_summary())


async def yearly(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(excel.get_yearly_summary())


async def export_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename="finance_" + datetime.now().strftime("%Y%m") + ".xlsx",
                caption="ไฟล์ Excel รายรับรายจ่าย"
            )
    else:
        await update.message.reply_text("ยังไม่มีข้อมูล ส่ง slip มาก่อนนะครับ")


def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", start))
    app.add_handler(CommandHandler("list", list_recent))
    app.add_handler(CommandHandler("edit", list_for_edit))
    app.add_handler(CommandHandler("summary", summary))
    app.add_handler(CommandHandler("yearly", yearly))
    app.add_handler(CommandHandler("export", export_excel))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_handler(CallbackQueryHandler(handle_category_callback, pattern="^cat:"))
    app.add_handler(CallbackQueryHandler(handle_delete_callback, pattern="^del:"))
    app.add_handler(CallbackQueryHandler(handle_edit_select_callback, pattern="^editrow:"))
    print("Bot running... Ctrl+C to stop")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
