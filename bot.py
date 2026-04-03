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

# Detaylı logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
EXCEL_URL = os.getenv("EXCEL_URL", "https://raw.githubusercontent.com/ilhan-yildiz/Calibration/main/Yearly_Calibration_Schedule.xlsx")
SHEET_NAME = os.getenv("SHEET_NAME", "TX Detail List")  # Varsayılan sheet adı
PORT = int(os.environ.get('PORT', 10000))

# Flask uygulaması
flask_app = Flask(__name__)

@flask_app.route('/')
def health_check():
    return jsonify({"status": "ok", "message": "Bot is running"}), 200

@flask_app.route('/health')
def health():
    return jsonify({
        "status": "healthy",
        "excel_url": EXCEL_URL,
        "sheet_name": SHEET_NAME,
        "timestamp": time.time()
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

def get_excel_info():
    """Excel dosyası bilgilerini al"""
    try:
        response = requests.get(EXCEL_URL, timeout=30)
        response.raise_for_status()
        
        workbook = openpyxl.load_workbook(BytesIO(response.content), data_only=True)
        
        info = {
            "sheets": workbook.sheetnames,
            "active_sheet": workbook.active.title if workbook.active else None,
            "total_sheets": len(workbook.sheetnames)
        }
        
        return info, workbook
    except Exception as e:
        logger.error(f"Excel bilgi alma hatası: {str(e)}")
        return None, None

def search_in_column_c(search_value, workbook, sheet_name):
    """C sütununda arama yap"""
    try:
        # Sheet kontrolü
        if sheet_name not in workbook.sheetnames:
            logger.warning(f"'{sheet_name}' sayfası bulunamadı. Mevcut sayfalar: {workbook.sheetnames}")
            return None, f"❌ '{sheet_name}' sayfası bulunamadı!\nMevcut sayfalar: {', '.join(workbook.sheetnames)}"
        
        sheet = workbook[sheet_name]
        logger.info(f"Arama yapılan sayfa: {sheet_name}, Satır sayısı: {sheet.max_row}")
        
        results = []
        row_count = 0
        
        # C sütunu = 3. sütun (index 2)
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, values_only=True):
            row_count += 1
            if len(row) > 2 and row[2] is not None:  # C sütunu kontrolü
                if str(row[2]).lower() == str(search_value).lower():
                    # İstenen sütunları al: D(3), E(4), H(7), I(8), J(9), K(10), L(11), M(12)
                    columns = {
                        'D': row[3] if len(row) > 3 else None,
                        'E': row[4] if len(row) > 4 else None,
                        'H': row[7] if len(row) > 7 else None,
                        'I': row[8] if len(row) > 8 else None,
                        'J': row[9] if len(row) > 9 else None,
                        'K': row[10] if len(row) > 10 else None,
                        'L': row[11] if len(row) > 11 else None,
                        'M': row[12] if len(row) > 12 else None
                    }
                    
                    result_text = f"📌 **Satır {row_count}** (C: {row[2]})\n"
                    for col, val in columns.items():
                        if val is not None:
                            result_text += f"   {col}: {val}\n"
                        else:
                            result_text += f"   {col}: Boş\n"
                    
                    results.append(result_text)
                    
                    if len(results) >= 5:
                        results.append("⚠️ *Sadece ilk 5 sonuç gösteriliyor*")
                        break
        
        logger.info(f"Tarama tamamlandı: {row_count} satır, {len(results)} sonuç bulundu")
        
        if not results:
            return None, None
        
        return results, None
    
    except Exception as e:
        logger.error(f"Arama hatası: {str(e)}", exc_info=True)
        return None, f"❌ Arama hatası: {str(e)}"

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/start komutu - Excel bilgilerini ve test verilerini gönder"""
    user_id = update.effective_user.id
    logger.info(f"Start komutu alındı - Kullanıcı: {user_id}")
    
    await update.message.reply_text("🔍 *Excel Bot Başlatılıyor...*\nExcel dosyası kontrol ediliyor...", parse_mode="Markdown")
    
    # Excel bilgilerini al
    info, workbook = get_excel_info()
    
    if not info:
        await update.message.reply_text("❌ Excel dosyasına erişilemiyor! URL'yi kontrol edin.")
        return
    
    # Excel bilgilerini göster
    sheets_list = "\n".join([f"• {sheet}" for sheet in info['sheets']])
    sheet_status = f"✅ Aranan sayfa: *{SHEET_NAME}*" if SHEET_NAME in info['sheets'] else f"❌ Aranan sayfa *{SHEET_NAME}* bulunamadı!"
    
    info_text = f"""📊 *Excel Dosya Bilgileri*

• Dosya URL: {EXCEL_URL[:50]}...
• Toplam Sayfa: {info['total_sheets']}
• Sayfalar: 
{sheets_list}
{sheet_status}

🔍 *Arama Ayarları*
• Aranan Sütun: *C sütunu*
• Gösterilen Sütunlar: D, E, H, I, J, K, L, M
"""
    
    await update.message.reply_text(info_text, parse_mode="Markdown")
    
    # Eğer TX Detail List sayfası varsa, ilk 2 satırı test olarak gönder
    if SHEET_NAME in info['sheets'] and workbook:
        try:
            sheet = workbook[SHEET_NAME]
            
            test_text = f"📝 *'{SHEET_NAME}' Sayfası Test Verileri (İlk 2 Satır)*\n\n"
            
            # İlk 2 satırı göster (1 ve 2. satırlar)
            for row_num in range(1, min(3, sheet.max_row + 1)):
                row_data = []
                for col in sheet[row_num]:
                    if col.value:
                        row_data.append(str(col.value))
                    else:
                        row_data.append("Boş")
                
                test_text += f"*Satır {row_num}:*\n"
                test_text += f"  A: {row_data[0] if len(row_data) > 0 else 'Boş'}\n"
                test_text += f"  B: {row_data[1] if len(row_data) > 1 else 'Boş'}\n"
                test_text += f"  C: {row_data[2] if len(row_data) > 2 else 'Boş'}\n"
                test_text += f"  D: {row_data[3] if len(row_data) > 3 else 'Boş'}\n"
                test_text += f"  E: {row_data[4] if len(row_data) > 4 else 'Boş'}\n"
                test_text += "  ---\n"
            
            await update.message.reply_text(test_text, parse_mode="Markdown")
            
            # Örnek arama önerisi
            if sheet.max_row > 2 and len(sheet[3]) > 2 and sheet[3][2].value:
                sample_value = sheet[3][2].value  # 3. satır, C sütunu
                await update.message.reply_text(
                    f"💡 *Örnek Arama:*\n`{sample_value}`\n\n"
                    f"Bu değeri aratmayı deneyebilirsiniz.",
                    parse_mode="Markdown"
                )
        
        except Exception as e:
            logger.error(f"Test verisi gönderme hatası: {str(e)}")
            await update.message.reply_text(f"⚠️ Test verileri alınamadı: {str(e)}")
    
    # Kullanım talimatları
    usage_text = """🔍 *Nasıl Kullanılır?*

Sadece aramak istediğiniz değeri yazın.
Bot *C sütununda* arama yapacak ve eşleşen satırların D, E, H, I, J, K, L, M sütunlarını gösterecek.

📝 *Komutlar:*
/start - Botu başlat ve Excel kontrolü yap
/check - Excel dosyasını tekrar kontrol et
/help - Yardım al
"""
    await update.message.reply_text(usage_text, parse_mode="Markdown")

async def check_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/check komutu - Excel dosyasını kontrol et"""
    await update.message.reply_text("🔍 Excel dosyası kontrol ediliyor...")
    
    info, workbook = get_excel_info()
    
    if not info:
        await update.message.reply_text("❌ Excel dosyasına erişilemiyor!")
        return
    
    sheets_list = "\n".join([f"• {sheet}" for sheet in info['sheets']])
    sheet_found = "✅" if SHEET_NAME in info['sheets'] else "❌"
    
    check_text = f"""📊 *Excel Kontrol Sonucu*

• URL: {EXCEL_URL[:60]}...
• Sayfa sayısı: {info['total_sheets']}
• Sayfalar:
{sheets_list}

• Aranan sayfa '{SHEET_NAME}': {sheet_found}

{'✅ Sayfa bulundu! Arama yapabilirsiniz.' if SHEET_NAME in info['sheets'] else '❌ Sayfa bulunamadı! SHEET_NAME değişkenini kontrol edin.'}
"""
    await update.message.reply_text(check_text, parse_mode="Markdown")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/help komutu"""
    help_text = """📖 *Yardım Menüsü*

*Nasıl arama yaparım?*
Aramak istediğiniz değeri doğrudan yazın.

*Arama Detayları:*
• Aranan Sütun: **C sütunu** (Tam eşleşme)
• Gösterilen Sütunlar: D, E, H, I, J, K, L, M

*Örnek:*
Excel'de C sütununda "KIPP00GHC01CP101" varsa, bu değeri yazın.

*Komutlar:*
/start - Botu başlat ve Excel bilgilerini göster
/check - Excel dosyasını kontrol et
/help - Bu yardım menüsü

*Not:* Arama büyük/küçük harf duyarlı DEĞİLDİR.
"""
    await update.message.reply_text(help_text, parse_mode="Markdown")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Mesajları işle - C sütununda arama yap"""
    search_text = update.message.text.strip()
    user_id = update.effective_user.id
    
    if not search_text:
        await update.message.reply_text("⚠️ Lütfen bir arama değeri girin.")
        return
    
    if len(search_text) < 2:
        await update.message.reply_text("⚠️ En az 2 karakter girmelisiniz.")
        return
    
    logger.info(f"🔍 Arama: '{search_text}' - Kullanıcı: {user_id}")
    
    waiting_msg = await update.message.reply_text(f"🔍 *'{search_text}'* aranıyor...\n📊 Sayfa: {SHEET_NAME}\n📍 Sütun: C", parse_mode="Markdown")
    
    try:
        # Excel dosyasını indir
        response = requests.get(EXCEL_URL, timeout=30)
        response.raise_for_status()
        
        workbook = openpyxl.load_workbook(BytesIO(response.content), data_only=True)
        
        # C sütununda ara
        results, error = search_in_column_c(search_text, workbook, SHEET_NAME)
        
        await waiting_msg.delete()
        
        if error:
            await update.message.reply_text(error, parse_mode="Markdown")
        elif results:
            # Sonuçları gönder
            await update.message.reply_text(f"✅ *'{search_text}' için {len([r for r in results if '⚠️' not in r])} sonuç bulundu:*", parse_mode="Markdown")
            
            for result in results:
                if len(result) > 4000:
                    for i in range(0, len(result), 3500):
                        await update.message.reply_text(result[i:i+3500], parse_mode="Markdown")
                else:
                    await update.message.reply_text(result, parse_mode="Markdown")
        else:
            await update.message.reply_text(f"❌ *'{search_text}'* için C sütununda eşleşme bulunamadı.", parse_mode="Markdown")
    
    except requests.exceptions.RequestException as e:
        await waiting_msg.delete()
        logger.error(f"Excel indirme hatası: {str(e)}")
        await update.message.reply_text(f"❌ Excel dosyası indirilemedi! URL'yi kontrol edin.\nHata: {str(e)}")
    
    except Exception as e:
        await waiting_msg.delete()
        logger.error(f"Mesaj işleme hatası: {str(e)}", exc_info=True)
        await update.message.reply_text(f"⚠️ Bir hata oluştu: {str(e)}")

async def error_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    logger.error(f"Telegram hatası: {context.error}", exc_info=True)
    if update and update.effective_message:
        await update.effective_message.reply_text("⚠️ Beklenmeyen bir hata oluştu.")

def main():
    if not TOKEN:
        logger.error("❌ TELEGRAM_BOT_TOKEN bulunamadı!")
        return
    
    logger.info("🤖 Excel Bot başlatılıyor...")
    logger.info(f"📊 Excel URL: {EXCEL_URL}")
    logger.info(f"📄 Aranan Sayfa: {SHEET_NAME}")
    logger.info(f"🔍 Aranan Sütun: C")
    
    application = Application.builder().token(TOKEN).build()
    
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("check", check_command))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    application.add_error_handler(error_handler)
    
    logger.info("✅ Bot başarıyla başlatıldı!")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
