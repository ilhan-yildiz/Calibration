import openpyxl
import requests
from io import BytesIO
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import os
import logging
from flask import Flask
import threading
import asyncio

# Logging ayarları
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
EXCEL_URL = os.getenv("EXCEL_URL")

def search_excel(search_value):
    """Excel'de C sütununda ara (pandas kullanmadan)"""
    try:
        logger.info(f"Aranan değer: {search_value}")
        
        response = requests.get(EXCEL_URL, timeout=30)
        response.raise_for_status()
        
        workbook = openpyxl.load_workbook(BytesIO(response.content), data_only=True)
        sheet = workbook["TX Detail List"]
        
        results = []
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, values_only=True):
            if len(row) > 2 and row[2] is not None:
                if str(row[2]).lower() == str(search_value).lower():
                    columns_needed = [3, 4, 7, 8, 9, 10, 11, 12]
                    column_letters = ['D', 'E', 'H', 'I', 'J', 'K', 'L', 'M']
                    
                    values = []
                    for i, col_idx in enumerate(columns_needed):
                        if col_idx < len(row):
                            val = row[col_idx] if row[col_idx] is not None else "Boş"
                            values.append(f"📌 {column_letters[i]}: {val}")
                        else:
                            values.append(f"📌 {column_letters[i]}: Yok")
                    
                    results.append("\n".join(values))
        
        if not results:
            return None
        
        return "\n\n" + "\n" + "-"*30 + "\n".join(results)
    
    except Exception as e:
        logger.error(f"Hata: {str(e)}")
        return f"❌ Hata oluştu: {str(e)}"

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome_text = """🤖 Merhaba! Ben Excel botuyum.

📊 Nasıl çalışırım:
• Excel dosyasının "TX Detail List" sheet'indeki C sütununda arama yaparım
• Eşleşme varsa D, E, H, I, J, K, L, M sütunlarındaki verileri gönderirim

🔍 Kullanım:
Sadece aramak istediğiniz değeri yazın.

Örnek: INV12345
"""
    await update.message.reply_text(welcome_text)

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    search_text = update.message.text.strip()
    
    logger.info(f"Mesaj alındı - Aranan: {search_text}")
    
    await update.message.reply_text(f"🔍 '{search_text}' aranıyor... Lütfen bekleyin.")
    
    result = search_excel(search_text)
    
    if result:
        if len(result) > 4000:
            for i in range(0, len(result), 4000):
                await update.message.reply_text(result[i:i+4000])
        else:
            await update.message.reply_text(f"✅ **BULUNDU:**\n\n{result}", parse_mode="Markdown")
    else:
        await update.message.reply_text(f"❌ '{search_text}' için eşleşme bulunamadı.")

async def error_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    logger.error(f"Update {update} caused error {context.error}")

# Flask uygulaması (Render'ın sağlık kontrolü için)
flask_app = Flask(__name__)

@flask_app.route('/')
def health_check():
    return "Bot is running!", 200

def run_http_server():
    port = int(os.environ.get('PORT', 10000))
    flask_app.run(host='0.0.0.0', port=port)

async def run_bot():
    """Bot'u başlatan asenkron fonksiyon"""
    app = Application.builder().token(TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_error_handler(error_handler)
    
    logger.info("Bot çalışıyor...")
    await app.run_polling(allowed_updates=Update.ALL_TYPES)

def main():
    if not TOKEN:
        logger.error("TELEGRAM_BOT_TOKEN environment variable bulunamadı!")
        return
    
    if not EXCEL_URL:
        logger.error("EXCEL_URL environment variable bulunamadı!")
        return

    # HTTP sunucusunu ayrı bir thread'de başlat
    threading.Thread(target=run_http_server, daemon=True).start()
    
    logger.info("Bot başlatılıyor...")
    
    # Yeni event loop oluştur ve bot'u çalıştır
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    
    try:
        loop.run_until_complete(run_bot())
    except KeyboardInterrupt:
        logger.info("Bot durduruluyor...")
    finally:
        loop.close()

if __name__ == "__main__":
    main()
