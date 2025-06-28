import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from flask import Flask, request
from dotenv import load_dotenv
import telegram

# Muat variabel dari file .env
load_dotenv()
print("Memulai aplikasi bot...")

# ==================================================
# ♠️ Inisialisasi
# ==================================================
app = Flask(__name__)

# Inisialisasi Bot Telegram
try:
    bot_token = os.getenv('TELEGRAM_TOKEN')
    if not bot_token:
        print("GAGAL: Pastikan TELEGRAM_TOKEN ada di file .env")
    bot = telegram.Bot(token=bot_token)
    print("Bot Telegram berhasil diinisialisasi.")
except Exception as e:
    print(f"GAGAL menginisialisasi bot: {e}")

# Mengatur koneksi ke Google Sheets
try:
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(os.getenv('SS_ID'))
    print(f"Berhasil terhubung ke Google Sheet: {spreadsheet.title}")
except Exception as e:
    print(f"GAGAL terhubung ke Google Sheets: {e}")

# ==================================================
# ♠️ Rute / Endpoint
# ==================================================
@app.route('/')
def index():
    return "Bot server is running."

@app.route('/webhook', methods=['POST'])
def webhook_handler():
    update_data = request.get_json()
    update = telegram.Update.de_json(update_data, bot)

    try:
        message = update.message or update.edited_message
        if message and message.text:
            chat_id = message.chat.id
            text = message.text
            print(f"Pesan diterima dari Chat ID {chat_id}: {text}")

            # --- LOGIKA UTAMA AKAN DISIMPAN DI SINI ---
            # Untuk sekarang, kita hanya akan membalas pesan yang masuk
            bot.send_message(chat_id=chat_id, text=f"Server menerima: '{text}'")

    except Exception as e:
        print(f"Error memproses pesan: {e}")

    return 'ok', 200

# ==================================================
# ♠️ Titik Mulai Aplikasi
# ==================================================
if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)