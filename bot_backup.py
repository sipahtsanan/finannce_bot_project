"""
Finance Tracker Bot
ส่ง slip ธนาคาร (รูปหรือ PDF) ผ่าน Telegram → AI อ่านข้อมูล → บันทึกลง Excel
"""

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
from PIL import Image
import io
import base64

# ── Config ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "YOUR_BOT_TOKEN_HERE")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY_HERE")
EXCEL_FILE = "finance_tracker.xlsx"

# หมวดหมู่ค่าใช้จ่าย - แก้ไขได้ตามต้องการ
CATEGORIES = [
    "🍔 อาหาร",
    "🚗 เดินทาง",
    "🏠 ที่พัก/ค่าเช่า",
    "💊 สุขภาพ/ยา",
    "🛍️ ช้อปปิ้ง",
    "🎮 บันเทิง",
    "📱 โทรศัพท์/อินเทอร์เน็ต",
    "💡 ค่าน้ำ/ไฟ",
    "📚 การศึกษา",
    "💰 ออม/ลงทุน",
    "💳 รายได้",
    "❓ อื่นๆ",
]

genai.configure(api_key=GEMINI_API_KEY)
excel = ExcelManager(EXCEL_FILE)

# ── Gemini: อ่าน slip ───────────────────────────────────────────────────────
def analyze_slip(image_bytes: bytes) -> dict:
    """ส่งรูปให้ Gemini วิเคราะห์ slip แล้วคืนค่าเป็น dict"""
    model = genai.GenerativeModel("gemini-2.5-flash")
    
    prompt = """วิเคราะห์ slip ธนาคารหรือใบเสร็จนี้ แล้วตอบเป็น JSON เท่านั้น ไม่ต้องมีข้อความอื่น
    
    รูปแบบ JSON:
    {
        "date": "วันที่ในรูปแบบ YYYY-MM-DD หรือ null ถ้าไม่มี",
        "amount": ตัวเลข (จำนวนเงิน บวก=รายได้ ลบ=รายจ่าย),
        "description": "รายละเอียดการทำรายการ",
        "merchant": "ชื่อร้านค้าหรือผู้รับเงิน",
        "transaction_type": "expense หรือ income",
        "suggested_category": "หมวดหมู่ที่แนะนำจากรายการ: อาหาร/เดินทาง/ที่พัก/สุขภาพ/ช้อปปิ้ง/บันเทิง/โทรศัพท์/ค่าน้ำไฟ/การศึกษา/ออมลงทุน/อื่นๆ"
    }"""
    
    image_part = {
        "mime_type": "image/jpeg",
        "data": base64.b64encode(image_bytes).decode()
    }
    
    response = model.generate_content([prompt, image_part])
    
    # Parse JSON จาก response
    text = response.text.strip()
    if "```json" in text:
        text = text.split("```json")[1].split("```")[0].strip()
    elif "```" in text:
        text = text.split("```")[1].split("```")[0].strip()
    
    return json.loads(text)

# ── Handlers ─────────────────────────────────────────────────────────────────
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = """👋 สวัสดี! ฉันคือ Finance Tracker Bot

📤 **วิธีใช้:**
• ส่งรูป slip หรือ PDF → ฉันจะอ่านและบันทึกให้อัตโนมัติ
• ส่งข้อความเช่น: `กาแฟ 80` หรือ `ค่าแท็กซี่ 150` ก็ได้

📊 **คำสั่ง:**
/summary - สรุปรายจ่ายเดือนนี้
/yearly - สรุปรายปี แยกตามหมวด
/export - ดาวน์โหลดไฟล์ Excel
/help - ดูวิธีใช้"""
    await update.message.reply_text(msg, parse_mode="Markdown")


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """รับรูปภาพ slip"""
    await update.message.reply_text("⏳ กำลังอ่าน slip...")
    
    photo = update.message.photo[-1]  # ขนาดใหญ่สุด
    file = await context.bot.get_file(photo.file_id)
    image_bytes = await file.download_as_bytearray()
    
    try:
        data = analyze_slip(bytes(image_bytes))
        await ask_confirm_category(update, context, data)
    except Exception as e:
        logger.error(f"Error analyzing slip: {e}")
        await update.message.reply_text(
            "❌ อ่าน slip ไม่ได้ ลองส่งใหม่ หรือพิมพ์ข้อมูลเองได้เลย\n"
            "รูปแบบ: `ชื่อรายการ จำนวนเงิน` เช่น `กาแฟ 80`",
            parse_mode="Markdown"
        )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """รับ PDF slip"""
    doc = update.message.document
    if doc.mime_type != "application/pdf":
        await update.message.reply_text("⚠️ รองรับเฉพาะ PDF หรือรูปภาพเท่านั้น")
        return
    
    await update.message.reply_text("⏳ กำลังอ่าน PDF slip...")
    
    file = await context.bot.get_file(doc.file_id)
    pdf_bytes = await file.download_as_bytearray()
    
    # แปลง PDF หน้าแรกเป็นรูป
    try:
        import fitz  # PyMuPDF
        pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page = pdf_doc[0]
        pix = page.get_pixmap(dpi=200)
        image_bytes = pix.tobytes("jpeg")
        
        data = analyze_slip(image_bytes)
        await ask_confirm_category(update, context, data)
    except ImportError:
        await update.message.reply_text(
            "⚠️ ต้องติดตั้ง PyMuPDF ก่อน:\n`pip install pymupdf`\n\n"
            "หรือส่งเป็นรูปภาพแทนได้เลยครับ",
            parse_mode="Markdown"
        )
    except Exception as e:
        logger.error(f"PDF error: {e}")
        await update.message.reply_text("❌ อ่าน PDF ไม่ได้ ลองส่งเป็นรูปภาพแทนครับ")


async def ask_confirm_category(update: Update, context: ContextTypes.DEFAULT_TYPE, data: dict):
    """แสดงข้อมูลที่อ่านได้และให้เลือกหมวด"""
    date_str = data.get("date") or datetime.now().strftime("%Y-%m-%d")
    amount = data.get("amount", 0)
    desc = data.get("description", "")
    merchant = data.get("merchant", "")
    suggested = data.get("suggested_category", "")
    
    # บันทึก pending transaction ไว้ใน context
    context.user_data["pending"] = {
        "date": date_str,
        "amount": abs(float(amount)),
        "description": f"{merchant} - {desc}".strip(" -"),
        "transaction_type": data.get("transaction_type", "expense")
    }
    
    msg = f"""📋 **ข้อมูลที่อ่านได้:**
📅 วันที่: {date_str}
💰 จำนวน: {abs(float(amount)):,.2f} บาท
📝 รายละเอียด: {merchant} {desc}
🏷️ AI แนะนำหมวด: {suggested}

**เลือกหมวดหมู่:**"""
    
    # สร้างปุ่ม category
    keyboard = []
    row = []
    for i, cat in enumerate(CATEGORIES):
        row.append(InlineKeyboardButton(cat, callback_data=f"cat:{cat}"))
        if len(row) == 2:
            keyboard.append(row)
            row = []
    if row:
        keyboard.append(row)
    
    keyboard.append([InlineKeyboardButton("❌ ยกเลิก", callback_data="cat:cancel")])
    
    await update.message.reply_text(
        msg,
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )


async def handle_category_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """รับหมวดหมู่ที่เลือกแล้วบันทึกลง Excel"""
    query = update.callback_query
    await query.answer()
    
    if query.data == "cat:cancel":
        await query.edit_message_text("❌ ยกเลิกแล้ว")
        return
    
    category = query.data.replace("cat:", "")
    pending = context.user_data.get("pending", {})
    
    if not pending:
        await query.edit_message_text("❌ ไม่พบข้อมูล ลองส่งใหม่อีกครั้ง")
        return
    
    # บันทึกลง Excel
    excel.add_transaction(
        date=pending["date"],
        amount=pending["amount"],
        category=category,
        description=pending["description"],
        transaction_type=pending["transaction_type"]
    )
    
    emoji = "💰" if pending["transaction_type"] == "income" else "💸"
    await query.edit_message_text(
        f"✅ บันทึกแล้ว!\n"
        f"{emoji} {pending['description']}\n"
        f"💵 {pending['amount']:,.2f} บาท\n"
        f"🏷️ {category}"
    )
    context.user_data.pop("pending", None)


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """รับข้อความ เช่น 'กาแฟ 80' หรือ 'เงินเดือน 30000'"""
    text = update.message.text.strip()
    
    # ใช้ Gemini parse ข้อความ
    model = genai.GenerativeModel("gemini-2.5-flash")
    prompt = f"""แปลงข้อความนี้เป็น JSON: "{text}"
    
    JSON รูปแบบ:
    {{
        "amount": ตัวเลข,
        "description": "รายละเอียด",
        "merchant": "",
        "date": "{datetime.now().strftime('%Y-%m-%d')}",
        "transaction_type": "expense หรือ income",
        "suggested_category": "หมวดหมู่"
    }}
    ตอบเฉพาะ JSON เท่านั้น"""
    
    try:
        response = model.generate_content(prompt)
        text_resp = response.text.strip()
        # Strip markdown code blocks
        text_resp = re.sub(r"```json\s*", "", text_resp)
        text_resp = re.sub(r"```\s*", "", text_resp)
        text_resp = text_resp.strip()
        data = json.loads(text_resp)
        await ask_confirm_category(update, context, data)
    except Exception as e:
        await update.message.reply_text(
            "❌ ไม่เข้าใจ ลองใหม่นะครับ เช่น:\n"
            "• `กาแฟ 80`\n• `ค่าแท็กซี่ 150`\n• `เงินเดือน 30000`",
            parse_mode="Markdown"
        )


async def summary(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """สรุปรายจ่ายเดือนนี้"""
    result = excel.get_monthly_summary()
    await update.message.reply_text(result, parse_mode="Markdown")


async def yearly(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """สรุปรายปี"""
    result = excel.get_yearly_summary()
    await update.message.reply_text(result, parse_mode="Markdown")


async def export_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """ส่งไฟล์ Excel"""
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=f"finance_{datetime.now().strftime('%Y%m')}.xlsx",
                caption="📊 ไฟล์ Excel รายรับรายจ่ายของคุณ"
            )
    else:
        await update.message.reply_text("❌ ยังไม่มีข้อมูล ส่ง slip มาก่อนนะครับ")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", start))
    app.add_handler(CommandHandler("summary", summary))
    app.add_handler(CommandHandler("yearly", yearly))
    app.add_handler(CommandHandler("export", export_excel))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_handler(CallbackQueryHandler(handle_category_callback, pattern="^cat:"))
    
    print("🤖 Bot กำลังทำงาน... กด Ctrl+C เพื่อหยุด")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
