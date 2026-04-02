import pandas as pd
import requests
from io import BytesIO
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import os
import logging

# Logging ayarları
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Telegram Bot Token (Render'dan environment variable olarak gelecek)
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")

# GitHub'daki Excel dosyasının RAW linki (Render'dan environment variable)
EXCEL_URL = os.getenv("EXCEL_URL")

def search_excel(search_value):
    """Excel dosyasında C sütununda ara ve istenen sütunları döndür"""
    try:
        logger.info(f"Aranan değer: {search_value}")
        
        # Excel'i indir
        response = requests.get(EXCEL_URL, timeout=30)
        response.raise_for_status()
        
        # Excel'i oku (header yok varsayımıyla)
        df = pd.read_excel(BytesIO(response.content), sheet_name="TX Detail List", header=None, engine="openpyxl")
        
        # C sütunu (index 2) tam eşleşme ara
        mask = df[2].astype(str).str.lower() == str(search_value).lower()
        results = df[mask]
        
        if results.empty:
            logger.info("Eşleşme bulunamadı")
            return None
        
        # İstenen sütunlar: D(3), E(4), H(7), I(8), J(9), K(10), L(11), M(12)
        columns_needed = [3, 4, 7, 8, 9, 10, 11, 12]
        column_letters = ['D', 'E', 'H', 'I', 'J', 'K', 'L', 'M']
        
        output_lines = []
        for idx, row in results.iterrows():
            values = []
            for i, col in enumerate(columns_needed):
                val = row[col] if pd.notna(row[col]) else "Boş"
                values.append(f"📌 {column_letters[i]}: {val}")
            output_lines.append("\n".join(values))
            output_lines.append("-" * 30)
        
        return "\n".join(output_lines)
    
    except Exception as e:
        logger.error(f"Hata: {str(e)}")
        return f"❌ Hata oluştu: {str(e)}"

# /start komutu
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

# Mesajları işle
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    search_text = update.message.text.strip()
    chat_id = update.effective_chat.id
    
    logger.info(f"Mesaj alındı - Chat ID: {chat_id}, Aranan: {search_text}")
    
    # Bekleme mesajı
    await update.message.reply_text(f"🔍 '{search_text}' aranıyor... Lütfen bekleyin.")
    
    # Arama yap
    result = search_excel(search_text)
    
    if result:
        # Cevap çok uzunsa parçala
        if len(result) > 4000:
            for i in range(0, len(result), 4000):
                await update.message.reply_text(result[i:i+4000])
        else:
            await update.message.reply_text(f"✅ **BULUNDU:**\n\n{result}", parse_mode="Markdown")
    else:
        await update.message.reply_text(f"❌ '{search_text}' için eşleşme bulunamadı.")

# Hata yakalama
async def error_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    logger.error(f"Update {update} caused error {context.error}")

# Ana fonksiyon
def main():
    if not TOKEN:
        logger.error("TELEGRAM_BOT_TOKEN environment variable bulunamadı!")
        return
    
    if not EXCEL_URL:
        logger.error("EXCEL_URL environment variable bulunamadı!")
        return
    
    logger.info("Bot başlatılıyor...")
    
    app = Application.builder().token(TOKEN).build()
    
    # Handler'ları ekle
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_error_handler(error_handler)
    
    # Botu başlat
    logger.info("Bot çalışıyor...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
