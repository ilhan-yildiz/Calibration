import pandas as pd
import requests
import json
import os
import sys
from datetime import datetime

EXCEL_FILE = "your-excel-file.xlsx"  # Repodaki Excel dosyası

def search_excel(search_value):
    try:
        # Local Excel dosyasını oku
        df = pd.read_excel(EXCEL_FILE, sheet_name="TX Detail List", header=None, engine="openpyxl")
        
        # C sütunu (index 2) ara
        mask = df[2].astype(str).str.lower() == str(search_value).lower()
        results = df[mask]
        
        if results.empty:
            return None
        
        # İstenen sütunlar: D(3), E(4), H(7), I(8), J(9), K(10), L(11), M(12)
        columns_needed = [3, 4, 7, 8, 9, 10, 11, 12]
        column_names = ['D', 'E', 'H', 'I', 'J', 'K', 'L', 'M']
        
        output_lines = []
        for _, row in results.iterrows():
            values = []
            for i, col in enumerate(columns_needed):
                val = str(row[col]) if pd.notna(row[col]) else "-"
                values.append(f"{column_names[i]}: {val}")
            output_lines.append(" | ".join(values))
        
        return "\n\n".join(output_lines)
    
    except Exception as e:
        return f"Hata: {str(e)}"

def send_telegram_message(chat_id, text, bot_token):
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": text,
        "parse_mode": "HTML"
    }
    response = requests.post(url, json=payload)
    return response.json()

def main():
    # GitHub Actions'dan gelen veriyi oku
    if len(sys.argv) < 4:
        print("Hata: Yeterli parametre yok")
        return
    
    chat_id = sys.argv[1]
    message_text = sys.argv[2]
    bot_token = sys.argv[3]
    
    print(f"Aranıyor: {message_text}")
    result = search_excel(message_text)
    
    if result:
        send_telegram_message(chat_id, f"✅ <b>Bulundu:</b>\n\n{result}", bot_token)
    else:
        send_telegram_message(chat_id, f"❌ '{message_text}' için eşleşme bulunamadı.", bot_token)

if __name__ == "__main__":
    main()
