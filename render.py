from bot import flask_app, main
import threading
import os

if __name__ == "__main__":
    # Bot'u ayrı thread'de başlat
    bot_thread = threading.Thread(target=main, daemon=True)
    bot_thread.start()
    
    # Flask uygulamasını başlat
    port = int(os.environ.get('PORT', 10000))
    flask_app.run(host='0.0.0.0', port=port)
