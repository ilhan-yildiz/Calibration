import os
import logging
from flask import Flask, jsonify
import threading
import openpyxl
from openpyxl.styles import PatternFill
import requests
from io import BytesIO
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes
from datetime import datetime
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

def get_column_headers(workbook, sheet_name):
    """2. satırdaki kolon başlıklarını al"""
    try:
        sheet = workbook[sheet_name]
        headers = {}
        for col_idx in range(1, sheet.max_column + 1):
            header_value = sheet.cell(2, col_idx).value
            if header_value:
                headers[col_idx] = header_value
        return headers
    except:
        return {}

def search_in_column_c_partial(search_value, workbook, sheet_name):
    """C sütununda kısmi eşleşme arama (3. satırdan itibaren)"""
    try:
        if sheet_name not in workbook.sheetnames:
            return None, f"❌ '{sheet_name}' sayfası bulunamadı!\nMevcut sayfalar: {', '.join(workbook.sheetnames)}"
        
        sheet = workbook[sheet_name]
        headers = get_column_headers(workbook, sheet_name)
        results = []
        search_lower = str(search_value).lower().strip()
        
        for row_idx, row in enumerate(sheet.iter_rows(min_row=3, values_only=True), 3):
            if len(row) > 2 and row[2] is not None:
                cell_value = str(row[2]).strip()
                if search_lower in cell_value.lower():
                    # C ve D sütunlarındaki verileri al
                    c_value = row[2] if len(row) > 2 else None
                    d_value = row[3] if len(row) > 3 else None
                    
                    result_text = f"📌 *Satır {row_idx}*\n"
                    result_text += f"📌 *{headers.get(3, 'C Sütunu')}:* {c_value}\n"
                    result_text += f"📅 *{headers.get(4, 'D Sütunu')}:* {d_value if d_value else 'Belirtilmemiş'}\n"
                    
                    results.append(result_text)
                    
                    if len(results) >= 10:
                        break
        
        if not results:
            return None, None
        return results, None
        
    except Exception as e:
        return None, f"Arama hatası: {str(e)}"

def search_calibration_date(search_value, workbook, sheet_name):
    """C sütununda arama yap, C ve D sütunlarındaki verileri göster"""
    try:
        if sheet_name not in workbook.sheetnames:
            return None, f"❌ '{sheet_name}' sayfası bulunamadı!"
        
        sheet = workbook[sheet_name]
        headers = get_column_headers(workbook, sheet_name)
        results = []
        search_lower = str(search_value).lower().strip()
        
        for row_idx, row in enumerate(sheet.iter_rows(min_row=3, values_only=True), 3):
            if len(row) > 2 and row[2] is not None:
                cell_value = str(row[2]).strip()
                if search_lower in cell_value.lower():
                    c_value = row[2] if len(row) > 2 else None
                    d_value = row[3] if len(row) > 3 else None
                    
                    result_text = f"📌 *{headers.get(3, 'Ekipman Kodu')}:* {c_value}\n"
                    result_text += f"📅 *{headers.get(4, 'Kalibrasyon Tarihi')}:* {d_value if d_value else 'Belirtilmemiş'}\n"
                    result_text += f"🔍 *Satır:* {row_idx}\n"
                    result_text += f"━━━━━━━━━━━━━━━━━━━━━\n"
                    
                    results.append(result_text)
                    
                    if len(results) >= 10:
                        break
        
        if not results:
            return None, None
        return results, None
        
    except Exception as e:
        return None, f"Arama hatası: {str(e)}"

def update_calibration_date(equipment_code, new_date, workbook, sheet_name):
    """Excel'de kalibrasyon tarihini güncelle"""
    try:
        if sheet_name not in workbook.sheetnames:
            return False, f"Sayfa bulunamadı: {sheet_name}"
        
        sheet = workbook[sheet_name]
        search_lower = str(equipment_code).lower().strip()
        found = False
        
        for row_idx, row in enumerate(sheet.iter_rows(min_row=3, max_row=sheet.max_row, values_only=False), 3):
            if len(row) > 2 and row[2].value is not None:
                cell_value = str(row[2].value).strip()
                if cell_value.lower() == search_lower:
                    # D sütununu güncelle (4. sütun)
                    date_cell = sheet.cell(row_idx, 4)
                    date_cell.value = new_date
                    
                    # Yeşil renklendir
                    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                    date_cell.fill = green_fill
                    
                    found = True
                    break
        
        if not found:
            return False, f"Ekipman kodu bulunamadı: {equipment_code}"
        
        return True, "Kalibrasyon tarihi güncellendi"
        
    except Exception as e:
        return False, f"Hata: {str(e)}"

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("🔍 Excel bot başlatılıyor...\n\nExcel dosyası kontrol ediliyor...")
    
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
            
            headers = get_column_headers(workbook, SHEET_NAME)
            info_text += "📝 *Kolon Başlıkları (2. satır)*\n"
            for col_idx, header in list(headers.items())[:5]:
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                info_text += f"• {col_letter}: {header}\n"
        
        await update.message.reply_text(info_text, parse_mode="Markdown")
        
    except Exception as e:
        await update.message.reply_text(f"⚠️ Excel okuma hatası: {str(e)}")

async def search_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/ara komutu - C sütununda kısmi eşleşme arama"""
    if not context.args:
        await update.message.reply_text("❌ *Kullanım:* `/ara ARANACAK_DEGER`\n\nÖrnek: `/ara 12LAB20CF101`", parse_mode="Markdown")
        return
    
    search_text = " ".join(context.args).strip()
    
    await update.message.reply_text(f"🔍 '{search_text}' aranıyor... (C sütununda, kısmi eşleşme)")
    logger.info(f"Arama yapılıyor: '{search_text}'")
    
    try:
        response = requests.get(EXCEL_URL, timeout=30)
        response.raise_for_status()
        
        workbook = openpyxl.load_workbook(BytesIO(response.content), data_only=True)
        
        results, error = search_in_column_c_partial(search_text, workbook, SHEET_NAME)
        
        if error:
            await update.message.reply_text(error)
        elif results:
            await update.message.reply_text(f"✅ {len(results)} sonuç bulundu:\n\n" + "\n".join(results), parse_mode="Markdown")
        else:
            await update.message.reply_text(f"❌ '{search_text}' için C sütununda eşleşme bulunamadı.")
            
    except Exception as e:
        await update.message.reply_text(f"❌ Hata: {str(e)}")

async def tarih_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/tarih komutu - Kalibrasyon tarihi sorgulama (sadece C ve D sütunları)"""
    if not context.args:
        await update.message.reply_text("❌ *Kullanım:* `/tarih ARANACAK_DEGER`\n\nÖrnek: `/tarih 12LAB20CF101`\n\nBu komut sadece ekipman kodu ve kalibrasyon tarihini gösterir.", parse_mode="Markdown")
        return
    
    search_text = " ".join(context.args).strip()
    
    await update.message.reply_text(f"📅 '{search_text}' için kalibrasyon tarihi aranıyor...")
    
    try:
        response = requests.get(EXCEL_URL, timeout=30)
        workbook = openpyxl.load_workbook(BytesIO(response.content), data_only=True)
        
        results, error = search_calibration_date(search_text, workbook, SHEET_NAME)
        
        if error:
            await update.message.reply_text(error)
        elif results:
            await update.message.reply_text(f"✅ *Kalibrasyon Bilgileri*\n\n" + "\n".join(results), parse_mode="Markdown")
        else:
            await update.message.reply_text(f"❌ '{search_text}' için kayıt bulunamadı.")
            
    except Exception as e:
        await update.message.reply_text(f"❌ Hata: {str(e)}")

async def guncelle_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/guncelle komutu - Kalibrasyon tarihini güncelle"""
    if len(context.args) < 2:
        await update.message.reply_text("❌ *Kullanım:* `/guncelle EKIPMAN_KODU YENI_TARIH`\n\nÖrnek: `/guncelle 12LAB20CF101 2026-05-15`\n\nTarih formatı: YYYY-AA-GG", parse_mode="Markdown")
        return
    
    equipment_code = context.args[0].strip()
    new_date = context.args[1].strip()
    
    # Tarih formatını kontrol et
    try:
        datetime.strptime(new_date, "%Y-%m-%d")
    except:
        await update.message.reply_text("❌ Hatalı tarih formatı! Lütfen YYYY-AA-GG formatında girin.\nÖrnek: 2026-05-15")
        return
    
    await update.message.reply_text(f"✏️ '{equipment_code}' için kalibrasyon tarihi güncelleniyor: {new_date}")
    
    try:
        # Excel'i indir
        response = requests.get(EXCEL_URL, timeout=30)
        workbook = openpyxl.load_workbook(BytesIO(response.content), data_only=False)
        
        # Güncelle
        success, message = update_calibration_date(equipment_code, new_date, workbook, SHEET_NAME)
        
        if success:
            # Güncellenmiş dosyayı kaydet (not: GitHub'a yazmak için ek işlem gerekir)
            # Şimdilik sadece mesaj gösteriyoruz
            await update.message.reply_text(f"✅ {message}\n\n⚠️ **Not:** Değişiklikler şu an sadece geçici olarak yapıldı. GitHub'a kaydetmek için ek kod gerekir.", parse_mode="Markdown")
        else:
            await update.message.reply_text(f"❌ {message}")
            
    except Exception as e:
        await update.message.reply_text(f"❌ Hata: {str(e)}")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = """📖 *Yardım Menüsü*

*Komutlar:*

🔍 `/ara <deger>` - C sütununda **kısmi eşleşme** arama yapar
   Örnek: `/ara 12LAB20CF101`

📅 `/tarih <deger>` - Kalibrasyon tarihi sorgular (sadece C ve D sütunları)
   Örnek: `/tarih 12LAB20CF101`

✏️ `/guncelle <kod> <tarih>` - Kalibrasyon tarihini günceller
   Örnek: `/guncelle 12LAB20CF101 2026-05-15`
   Tarih formatı: YYYY-AA-GG

ℹ️ `/start` - Botu başlat ve Excel bilgilerini göster
🆘 `/help` - Bu yardım menüsü

*Özellikler:*
• Kolon isimleri 2. satırdan alınır
• Kısmi eşleşme yapar (büyük/küçük harf duyarsız)
• Tarih sorgulama özel format"""
    
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
    application.add_handler(CommandHandler("tarih", tarih_command))
    application.add_handler(CommandHandler("guncelle", guncelle_command))
    application.add_handler(CommandHandler("help", help_command))
    
    application.run_polling()

if __name__ == "__main__":
    main()
