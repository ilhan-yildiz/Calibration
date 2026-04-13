import os
import logging
from flask import Flask, jsonify
import threading
import openpyxl
import requests
from io import BytesIO
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import time
import asyncio

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
EXCEL_URL = os.getenv("EXCEL_URL")
SHEET_NAME = os.getenv("SHEET_NAME", "TX Detail List")
PORT = int(os.environ.get('PORT', 10000))

# Flask uygulaması
flask_app = Flask(__name__)

@flask_app.route('/')
def health_check():
    return jsonify({"status": "ok"}), 200

def run_http_server():
    flask_app.run(host='0.0.0.0', port=PORT, use_reloader=False)

threading.Thread(target=run_http_server, daemon=True).start()
logger.info(f"HTTP sunucusu başlatıldı - Port: {PORT}")

def test_excel_url():
    """Excel URL'sini test et"""
    try:
        logger.info(f"Test edilen URL: {EXCEL_URL}")
        response = requests.get(EXCEL_URL, timeout=30, verify=True)
        logger.info(f"HTTP Status: {response.status_code}")
        logger.info(f"Content-Type: {response.headers.get('Content-Type')}")
        logger.info(f"Dosya boyutu: {len(response.content)} bytes")
        
        if response.status_code == 200:
            # Excel dosyası mı kontrol et
            if b'PK' in response.content[:2]:  # Excel dosyaları PK ile başlar
                logger.info("✅ Dosya geçerli bir Excel dosyası")
                return True, "Excel dosyası geçerli"
            else:
                return False, "Dosya Excel formatında değil"
        else:
            return False, f"HTTP {response.status_code} hatası"
            
    except requests.exceptions.Timeout:
        return False, "Bağlantı zaman aşımı"
    except requests.exceptions.ConnectionError:
        return False, "Bağlantı hatası - URL'ye ulaşılamıyor"
    except Exception as e:
        return False, f"Hata: {str(e)}"

def search_in_column_c(search_value, workbook, sheet_name):
    """C sütununda arama yap"""
    try:
        if sheet_name not in workbook.sheetnames:
            return None, f"❌ '{sheet_name}' sayfası bulunamadı!\nMevcut sayfalar: {', '.join(workbook.sheetnames)}"
        
        sheet = workbook[sheet_name]
        results = []
        
        for row_idx, row in enumerate(sheet.iter_rows(min_row=1, values_only=True), 1):
            if len(row) > 2 and row[2] is not None:
                if str(row[2]).lower() == str(search_value).lower():
                    # D, E, H, I, J, K, L, M sütunları
                    cols = {
                        'D': row[3] if len(row) > 3 else None,
                        'E': row[4] if len(row) > 4 else None,
                        'H': row[7] if len(row) > 7 else None,
                        'I': row[8] if len(row) > 8 else None,
                        'J': row[9] if len(row) > 9 else None,
                        'K': row[10] if len(row) > 10 else None,
                        'L': row[11] if len(row) > 11 else None,
                        'M': row[12] if len(row) > 12 else None,
                    }
                    
                    result_text = f"📌 Satır {row_idx} (C: {row[2]})\n"
                    for col, val in cols.items():
                        result_text += f"   {col}: {val if val else 'Boş'}\n"
                    results.append(result_text)
                    
                    if len(results) >= 5:
                        break
        
        if not results:
            return None, None
        return results, None
        
    except Exception as e:
        return None, f"Arama hatası: {str(e)}"

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("🔍 Excel bot başlatılıyor...\n\nExcel dosyası kontrol ediliyor...")
    
    # URL'yi test et
    is_valid, message = test_excel_url()
    
    if not is_valid:
        await update.message.reply_text(f"❌ *Excel Hatası*\n\n{message}\n\n📋 Excel URL: `{EXCEL_URL}`\n\nLütfen URL'yi kontrol edin.", parse_mode="Markdown")
        return
    
    await update.message.reply_text(f"✅ Excel dosyasına erişim başarılı!\n\n{message}\n\nŞimdi arama yapabilirsiniz.", parse_mode="Markdown")
    
    # Sayfa bilgilerini al
    try:
        response = requests.get(EXCEL_URL, timeout=30)
        workbook = openpyxl.load_workbook(BytesIO(response.content), data_only=True)
        
        sheets = workbook.sheetnames
        sheet_found = SHEET_NAME in sheets
        
        info_text = f"📊 *Excel Bilgileri*\n\n"
        info_text += f"• Sayfalar: {', '.join(sheets)}\n"
        info_text += f"• Aranan sayfa '{SHEET_NAME}': {'✅ Bulundu' if sheet_found else '❌ Bulunamadı'}\n"
        
        if sheet_found:
            sheet = workbook[SHEET_NAME]
            info_text += f"• Toplam satır: {sheet.max_row}\n"
            info_text += f"• Toplam sütun: {sheet.max_column}\n\n"
            
            # İlk 2 satırı göster
            info_text += "📝 *Test Verileri (İlk 2 satır)*\n\n"
            for row_num in range(1, min(3, sheet.max_row + 1)):
                row = list(sheet[row_num])
                if row:
                    info_text += f"Satır {row_num}:\n"
                    info_text += f"  A: {row[0].value if len(row) > 0 else 'Boş'}\n"
                    info_text += f"  B: {row[1].value if len(row) > 1 else 'Boş'}\n"
                    info_text += f"  C: {row[2].value if len(row) > 2 else 'Boş'}\n"
                    info_text += f"  D: {row[3].value if len(row) > 3 else 'Boş'}\n\n"
        
        await update.message.reply_text(info_text, parse_mode="Markdown")
        
    except Exception as e:
        await update.message.reply_text(f"⚠️ Excel okuma hatası: {str(e)}")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    search_text = update.message.text.strip()
    
    if not search_text:
        await update.message.reply_text("⚠️ Lütfen bir arama değeri girin.")
        return
    
    await update.message.reply_text(f"🔍 '{search_text}' aranıyor... (Sayfa: {SHEET_NAME}, Sütun: C)")
    
    try:
        # Excel'i indir
        response = requests.get(EXCEL_URL, timeout=30)
        response.raise_for_status()
        
        workbook = openpyxl.load_workbook(BytesIO(response.content), data_only=True)
        
        # Ara
        results, error = search_in_column_c(search_text, workbook, SHEET_NAME)
        
        if error:
            await update.message.reply_text(error)
        elif results:
            await update.message.reply_text(f"✅ {len(results)} sonuç bulundu:\n\n" + "\n".join(results))
        else:
            await update.message.reply_text(f"❌ '{search_text}' için C sütununda eşleşme bulunamadı.")
            
    except Exception as e:
        await update.message.reply_text(f"❌ Hata: {str(e)}\n\nURL: {EXCEL_URL}")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = """📖 *Yardım*

• Arama yapmak için bir değer yazın
• Bot C sütununda tam eşleşme arar
• D, E, H, I, J, K, L, M sütunlarını gösterir

*Komutlar:*
/start - Botu başlat ve Excel'i test et
/help - Bu yardım

*Not:* Excel URL'nin doğru olduğundan emin olun."""
    await update.message.reply_text(help_text, parse_mode="Markdown")

def main():
    if not TOKEN:
        logger.error("❌ TELEGRAM_BOT_TOKEN yok!")
        return
    
    if not EXCEL_URL:
        logger.error("❌ EXCEL_URL yok!")
        return
    
    logger.info(f"Excel URL: {EXCEL_URL}")
    
    application = Application.builder().token(TOKEN).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    application.run_polling()

if __name__ == "__main__":
    main()
