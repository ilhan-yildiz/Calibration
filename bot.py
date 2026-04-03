import os
import logging
from flask import Flask, jsonify
import threading
import openpyxl
import requests
from io import BytesIO
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from functools import lru_cache
import time
import asyncio
from datetime import datetime

# Logging ayarları
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Environment variables
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
EXCEL_URL = os.getenv("EXCEL_URL", "https://raw.githubusercontent.com/ilhan-yildiz/Calibration/main/Yearly_Calibration_Schedule.xlsx")
PORT = int(os.environ.get('PORT', 10000))

# Excel cache sınıfı
class ExcelSearcher:
    def __init__(self, excel_url, cache_duration=300):
        self.excel_url = excel_url
        self.cache_duration = cache_duration
        self.cached_workbook = None
        self.last_update = 0
    
    def get_workbook(self):
        current_time = time.time()
        
        if self.cached_workbook and (current_time - self.last_update) < self.cache_duration:
            logger.info("Excel cache'den alındı")
            return self.cached_workbook
        
        try:
            logger.info(f"Excel dosyası indiriliyor: {self.excel_url}")
            response = requests.get(self.excel_url, timeout=30)
            response.raise_for_status()
            
            if len(response.content) > 50 * 1024 * 1024:
                raise ValueError("Excel dosyası çok büyük (max 50MB)")
            
            self.cached_workbook = openpyxl.load_workbook(
                BytesIO(response.content), 
                data_only=True, 
                read_only=True
            )
            self.last_update = current_time
            logger.info("Excel dosyası başarıyla yüklendi")
            return self.cached_workbook
            
        except Exception as e:
            logger.error(f"Excel yükleme hatası: {str(e)}")
            raise
    
    def search(self, search_value):
        try:
            workbook = self.get_workbook()
            
            # İlk sayfayı kullan (veya belirli bir sayfa)
            sheet = workbook.active
            
            # Sayfa adını logla
            logger.info(f"Arama yapılan sayfa: {sheet.title}")
            
            results = []
            row_count = 0
            
            # Tüm satırları tara
            for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, values_only=True):
                row_count += 1
                if row and len(row) > 0:
                    # İlk sütunda (A sütunu) arama yap
                    first_col = row[0] if row[0] is not None else ""
                    if str(first_col).lower() == str(search_value).lower():
                        # Tüm sütunları topla
                        values = []
                        for idx, val in enumerate(row[:10]):  # İlk 10 sütun
                            col_letter = openpyxl.utils.get_column_letter(idx + 1)
                            val_str = str(val) if val is not None else "Boş"
                            values.append(f"📌 {col_letter}: {val_str}")
                        
                        results.append("\n".join(values))
                        
                        if len(results) >= 5:
                            results.append("⚠️ *Sadece ilk 5 sonuç gösteriliyor*")
                            break
            
            logger.info(f"Toplam {row_count} satır tarandı, {len(results)} sonuç bulundu")
            
            if not results:
                return None
            
            return "\n\n" + "\n" + "─" * 40 + "\n".join(results)
        
        except Exception as e:
            logger.error(f"Arama hatası: {str(e)}")
            return f"❌ Hata: {str(e)}"

# Flask uygulaması
flask_app = Flask(__name__)

@flask_app.route('/')
def health_check():
    return jsonify({
        "status": "ok",
        "bot": "Calibration Bot",
        "timestamp": datetime.now().isoformat()
    }), 200

@flask_app.route('/health')
def health():
    return jsonify({
        "status": "healthy",
        "excel_url": EXCEL_URL,
        "time": time.time()
    }), 200

def run_http_server():
    try:
        flask_app.run(host='0.0.0.0', port=PORT, use_reloader=False)
    except Exception as e:
        logger.error(f"HTTP sunucu hatası: {str(e)}")

# HTTP sunucusunu başlat
server_thread = threading.Thread(target=run_http_server, daemon=True)
server_thread.start()
logger.info(f"HTTP sunucusu başlatıldı - Port: {PORT}")

# Global searcher
searcher = ExcelSearcher(EXCEL_URL) if EXCEL_URL else None

# Telegram komutları
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome_text = """🔧 *Kalibrasyon Botu'na Hoş Geldiniz!*

📊 *Özellikler:*
• Yıllık Kalibrasyon Takvimi'nde arama yapar
• Excel dosyasındaki verileri hızlıca bulur
• Cache sistemi ile hızlı yanıt verir

🔍 *Kullanım:*
Aramak istediğiniz cihaz adını veya ID'yi yazın

📝 *Örnek:*
`Multimeter`
`FLUKE-1234`

⚙️ *Komutlar:*
/start - Botu başlat
/help - Yardım
/info - Bot bilgisi
"""
    await update.message.reply_text(welcome_text, parse_mode="Markdown")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = """📖 *Yardım Menüsü*

*Nasıl kullanılır?*
1. Aradığınız cihaz adını yazın
2. Bot Excel'de arasın
3. Size sonuçları göstersin

*Önemli Notlar:*
• Arama büyük/küçük harf duyarsızdır
• En fazla 5 sonuç gösterilir
• Excel her 5 dakikada bir güncellenir

*Örnek aramalar:*
• `Kalibrasyon`
• `Test Cihazı`
• `2024-001`
"""
    await update.message.reply_text(help_text, parse_mode="Markdown")

async def info_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    info_text = f"""ℹ️ *Bot Bilgisi*

• Versiyon: 1.0.0
• Excel Kaynağı: GitHub
• Cache Süresi: 5 dakika
• Durum: ✅ Aktif

📊 *İstatistikler*
• Token: {'✅ Var' if TOKEN else '❌ Yok'}
• Excel: {'✅ Yüklü' if EXCEL_URL else '❌ Yok'}
"""
    await update.message.reply_text(info_text, parse_mode="Markdown")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    search_text = update.message.text.strip()
    
    if not search_text:
        await update.message.reply_text("⚠️ Lütfen bir arama terimi girin.")
        return
    
    if len(search_text) < 2:
        await update.message.reply_text("⚠️ En az 2 karakter girmelisiniz.")
        return
    
    logger.info(f"🔍 Arama: '{search_text}' - Kullanıcı: {update.effective_user.id}")
    
    waiting_msg = await update.message.reply_text(f"🔍 *'{search_text}'* aranıyor...", parse_mode="Markdown")
    
    try:
        result = await asyncio.to_thread(searcher.search, search_text)
        
        await waiting_msg.delete()
        
        if result and "❌" not in result:
            if len(result) > 4000:
                await update.message.reply_text(f"✅ *'{search_text}' için sonuçlar:*")
                for i in range(0, len(result), 3500):
                    await update.message.reply_text(result[i:i+3500], parse_mode="Markdown")
            else:
                await update.message.reply_text(
                    f"✅ *'{search_text}' için sonuçlar:*\n{result}", 
                    parse_mode="Markdown"
                )
        else:
            await update.message.reply_text(f"❌ *'{search_text}'* için eşleşme bulunamadı.", parse_mode="Markdown")
            
    except Exception as e:
        logger.error(f"Hata: {str(e)}")
        await waiting_msg.delete()
        await update.message.reply_text("⚠️ Bir hata oluştu. Lütfen tekrar deneyin.")

async def error_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    logger.error(f"Telegram hatası: {context.error}")

def main():
    if not TOKEN:
        logger.error("❌ TELEGRAM_BOT_TOKEN bulunamadı!")
        return
    
    logger.info("🤖 Kalibrasyon Botu başlatılıyor...")
    
    application = Application.builder().token(TOKEN).build()
    
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("info", info_command))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    application.add_error_handler(error_handler)
    
    logger.info("✅ Bot başarıyla başlatıldı!")
    
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
