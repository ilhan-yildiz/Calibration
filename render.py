services:
  - type: web
    name: bot
    runtime: python
    buildCommand: pip install -r requirements.txt
    startCommand: python bot.py
    envVars:
      - key: TELEGRAM_BOT
        sync: false
      - key: EXCEL_URL
        sync: false
