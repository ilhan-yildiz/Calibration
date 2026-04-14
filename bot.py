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
SHEET_NAME = os.getenv("SHEET_NAME", "Tx Detail List")
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
            if b'PK' in response.content[:2]:
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

def search_in_all_columns(search_value, workbook, sheet_name):
    """Tüm sütunlarda arama yap (kısmi eşleşme)"""
    try:
        if sheet_name not in workbook.sheetnames:
            return None, f"❌ '{sheet_name}' sayfası bulunamadı!\nMevcut sayfalar: {', '.join(workbook.sheetnames)}"
        
        sheet = workbook[sheet_name]
        results = []
        search_lower = str(search_value).lower()
        
        for row_idx, row in enumerate(sheet.iter_rows(min_row=1, values_only=True), 1):
            for col_idx, cell_value in enumerate(row):
                if cell_value is not None and search_lower in str(cell_value).lower():
                    # Sütun harfini bul (A, B, C, ... AA, AB, ...)
                    col_letter = ""
                    temp = col_idx
                    while temp >= 0:
                        temp -= 1
                        col_letter = chr(65 + (temp % 26)) + col_letter
                        temp = temp // 26 - 1
                        if temp < 0:
                            break
                    if not col_letter:
                        col_letter = chr(65 + col_idx)
                    
                    # Tüm satırı al (ilk 15 sütun)
                    row_data = []
                    for i in range(min(15, len(row))):
                        if row[i] is not None:
                            col_let = ""
                            t = i
                            while t >= 0:
                                t -= 1
                                col_let = chr(65 + (t % 26)) + col_let
                                t = t // 26 - 1
                                if t < 0:
                                    break
                            if not col_let:
                                col_let = chr(65 + i)
                            row_data.append(f"{col_let}: {row[i]}")
                    
                    result_text = f"📌 *Satır {row_idx}* (Bulunan: {col_letter} sütununda '{cell_value}')\n"
                    result_text += "   " + "\n   ".join(row_data[:10])  # İlk 10 sütunu göster
                    results.append(result_text)
                    
                    if len(results) >= 10:
                        break
            if len(results) >= 10:
                break
        
        if not results:
            return None, None
        return results, None
        
    except Exception as e:
        return None, f"Arama hatası: {str(e)}"

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("🔍 Excel bot başlatılıyor...\n\nExcel dosyası kontrol ediliyor...")
    
    is_valid, message = test_excel_url()
    
    if not is_valid:
        await update.message.reply_text(f"❌ *Excel Hatası*\n\n{message}\n\n📋 Excel URL: `{EXCEL_URL}`\n\nLütfen URL'yi kontrol edin.", parse_mode="Markdown")
        return
    
    await update.message.reply_text(f"✅ Excel dosyasına erişim başarılı!\n\n{message}\n\nŞimdi arama yapabilirsiniz.", parse_mode="Markdown")
    
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
            
            info_text += "📝 *Örnek Veriler (İlk 3 satır, ilk 5 sütun)*\n\n"
            for row_num in range(1, min(4, sheet.max_row + 1)):
                info_text += f"Satır {row_num}:\n"
                for col_num in range(1, min(6, sheet.max_column + 1)):
                    cell_value = sheet.cell(row_num, col_num).value
                    col_letter = openpyxl.utils.get_column_letter(col_num)
                    info_text += f"  {col_letter}: {cell_value if cell_value else 'Boş'}\n"
                info_text += "\n"
        
        await update.message.reply_text(info_text, parse_mode="Markdown")
        
    except Exception as e:
        await update.message.reply_text(f"⚠️ Excel okuma hatası: {str(e)}")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    search_text = update.message.text.strip()
    
    if not search_text:
        await update.message.reply_text("⚠️ Lütfen bir arama değeri girin.")
        return
    
    logger.info(f"🔍 Arama yapılıyor: '{search_text}'")
    await update.message.reply_text(f"🔍 '{search_text}' aranıyor... (Tüm sütunlarda, kısmi eşleşme)")
    
    try:
        response = requests.get(EXCEL_URL, timeout=30)
        response.raise_for_status()
        
        workbook = openpyxl.load_workbook(BytesIO(response.content), data_only=True)
        
        results, error = search_in_all_columns(search_text, workbook, SHEET_NAME)
        
        if error:
            logger.error(f"Hata: {error}")
            await update.message.reply_text(error)
        elif results:
            logger.info(f"✅ {len(results)} sonuç bulundu")
            # Sonuçları parçalara böl (Telegram mesaj limiti 4096 karakter)
            for i, result in enumerate(results):
                await update.message.reply_text(result, parse_mode="Markdown")
                if i < len(results) - 1:
                    await asyncio.sleep(0.5)
        else:
            logger.info(f"❌ Sonuç bulunamadı: '{search_text}'")
            await update.message.reply_text(f"❌ '{search_text}' için hiçbir sütunda eşleşme bulunamadı.\n\n💡 İpucu: Büyük/küçük harf fark etmez, kısmi eşleşme yaparım.")
            
    except Exception as e:
        logger.error(f"Exception: {str(e)}")
        await update.message.reply_text(f"❌ Hata: {str(e)}")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = """📖 *Yardım*

• Arama yapmak için bir değer yazın
• **Tüm sütunlarda** ve **kısmi eşleşme** ile ararım
• Büyük/küçük harf fark etmez
• Örnek: `KKS` yazarsanız "KKS-001", "AKKS01" gibi değerleri de bulurum

*Komutlar:*
/start - Botu başlat ve Excel'i test et
/help - Bu yardım

*Not:* Excel dosyası her aramada yeniden indirilir, biraz bekleyebilir."""
    await update.message.reply_text(help_text, parse_mode="Markdown")

def main():
    if not TOKEN:
        logger.error("❌ TELEGRAM_BOT_TOKEN yok!")
        return
    
    if not EXCEL_URL:
        logger.error("❌ EXCEL_URL yok!")
        return
    
    logger.info(f"Excel URL: {EXCEL_URL}")
    logger.info(f"Aranan sayfa: {SHEET_NAME}")
    
    application = Application.builder().token(TOKEN).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    application.run_polling()

if __name__ == "__main__":
    main()
