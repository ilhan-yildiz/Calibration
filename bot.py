import os
import logging
from flask import Flask
import threading

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")

# Flask uygulaması
flask_app = Flask(__name__)

@flask_app.route('/')
def health_check():
    return "Bot is running!", 200

def run_http_server():
    port = int(os.environ.get('PORT', 10000))
    flask_app.run(host='0.0.0.0', port=port)

# HTTP sunucusunu başlat
threading.Thread(target=run_http_server, daemon=True).start()

# Telegram bot
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters

async def start(update: Update, context):
    await update.message.reply_text("Merhaba! Bot çalışıyor!")

async def echo(update: Update, context):
    await update.message.reply_text(f"Mesajınız: {update.message.text}")

def main():
    if not TOKEN:
        logger.error("TELEGRAM_BOT_TOKEN yok!")
        return
    
    logger.info("Bot başlatılıyor...")
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT, echo))
    
    logger.info("Bot çalışıyor...")
    app.run_polling()

if __name__ == "__main__":
    main()
