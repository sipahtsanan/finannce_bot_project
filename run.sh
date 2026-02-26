#!/bin/bash
# run.sh - สคริปต์เริ่ม bot

echo "🤖 กำลังเริ่ม Finance Tracker Bot..."

# โหลด environment variables
if [ -f .env ]; then
    export $(cat .env | grep -v '#' | xargs)
    echo "✅ โหลด .env แล้ว"
else
    echo "❌ ไม่พบไฟล์ .env - กรุณา copy .env.example เป็น .env แล้วใส่ API Keys"
    exit 1
fi

# ตรวจสอบ dependencies
python3 -c "import telegram" 2>/dev/null || {
    echo "📦 ติดตั้ง dependencies..."
    pip3 install -r requirements.txt
}

# รัน bot
python3 bot.py
