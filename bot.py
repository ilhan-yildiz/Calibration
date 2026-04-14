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
        
        if response.status_code == 200:
            if b'PK' in response.content[:2]:
                logger.info("✅ Dosya geçerli bir Excel dosyası")
                return True, "Excel dosyası geçerli"
            else:
                return False, "Dosya Excel formatında değil"
        else:
            return False, f"HTTP {response.status_code} hatası"
            
    except Exception as e:
        return False, f"Hata: {str(e)}"

def search_in_column_c(search_value, workbook, sheet_name):
    """C sütununda tam eşleşme arama (3. satırdan itibaren)"""
    try:
        if sheet_name not in workbook.sheetnames:
            return None, f"❌ '{sheet_name}' sayfası bulunamadı!\nMevcut sayfalar: {', '.join(workbook.sheetnames)}"
        
        sheet = workbook[sheet_name]
        results = []
        search_lower = str(search_value).lower().strip()
        
        for row_idx, row in enumerate(sheet.iter_rows(min_row=3, values_only=True), 3):  # 3. satırdan başla (başlıkları atla)
            if len(row) > 2 and row[2] is not None:
                cell_value = str(row[2]).strip()
                if cell_value.lower() == search_lower:
                    # Tüm satırı göster (ilk 15 sütun)
                    result_text = f"📌 *Satır {row_idx}*\n"
                    result_text += f"   C sütunu: {row[2]}\n\n"
                    
                    for col_idx in range(min(15, len(row))):
                        if row[col_idx] is not None and col_idx != 2:  # C hariç diğerlerini göster
                            col_letter = openpyxl.utils.get_column_letter(col_idx + 1)
                            result_text += f"   {col_letter}: {row[col_idx]}\n"
                    
                    results.append(result_text)
                    
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
        await update.message.reply_text(f"❌ *Excel Hatası*\n\n{message}", parse_mode="Markdown")
        return
    
    await update.message.reply_text(f"✅ Excel dosyasına erişim başarılı!\n\n{message}\n\nŞimdi `/ara` komutunu kullanabilirsiniz.", parse_mode="Markdown")
    
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
            
            info_text += "📝 *Örnek Veriler (C sütunu, ilk 5 veri satırı)*\n\n"
            count = 0
            for row_num in range(3, min(8, sheet.max_row + 1)):
                c_value = sheet.cell(row_num, 3).value
                if c_value:
                    info_text += f"Satır {row_num}: {c_value}\n"
                    count += 1
                    if count >= 5:
                        break
        
        await update.message.reply_text(info_text, parse_mode="Markdown")
        
    except Exception as e:
        await update.message.reply_text(f"⚠️ Excel okuma hatası: {str(e)}")

async def search_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/ara komutu - C sütununda arama yapar"""
    if not context.args:
        await update.message.reply_text("❌ *Kullanım:* `/ara ARANACAK_DEGER`\n\nÖrnek: `/ara 12LAB20CF101`", parse_mode="Markdown")
        return
    
    search_text = " ".join(context.args).strip()
    
    await update.message.reply_text(f"🔍 '{search_text}' aranıyor... (C sütununda, tam eşleşme)")
    logger.info(f"Arama yapılıyor: '{search_text}'")
    
    try:
        response = requests.get(EXCEL_URL, timeout=30)
        response.raise_for_status()
        
        workbook = openpyxl.load_workbook(BytesIO(response.content), data_only=True)
        
        results, error = search_in_column_c(search_text, workbook, SHEET_NAME)
        
        if error:
            logger.error(f"Hata: {error}")
            await update.message.reply_text(error)
        elif results:
            logger.info(f"✅ {len(results)} sonuç bulundu")
            for result in results:
                await update.message.reply_text(result, parse_mode="Markdown")
                await asyncio.sleep(0.3)
        else:
            logger.info(f"❌ Sonuç bulunamadı: '{search_text}'")
            await update.message.reply_text(f"❌ '{search_text}' için C sütununda eşleşme bulunamadı.\n\n💡 İpucu: Tam eşleşme arıyorum. Büyük/küçük harf fark etmez.")
            
    except Exception as e:
        logger.error(f"Exception: {str(e)}")
        await update.message.reply_text(f"❌ Hata: {str(e)}")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = """📖 *Yardım*

*Komutlar:*
/start - Botu başlat ve Excel'i test et
/ara <deger> - C sütununda tam eşleşme arama yapar
/help - Bu yardım

*Örnekler:*
`/ara 12LAB20CF101`
`/ara TX-001`

*Not:* 
• C sütununda tam eşleşme arar
• Büyük/küçük harf fark etmez
• 3. satırdan itibaren arama yapar"""
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
    application.add_handler(CommandHandler("ara", search_command))
    application.add_handler(CommandHandler("help", help_command))
    # Normal mesajları işleme (opsiyonel - isterseniz kaldırabilirsiniz)
    # application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    application.run_polling()

if __name__ == "__main__":
    main()
