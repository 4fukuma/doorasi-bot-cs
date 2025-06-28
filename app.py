# -*- coding: utf-8 -*-

import os
import gspread
import telegram
import logging
import regex as re
import json
from datetime import datetime, timedelta, date
from calendar import month_name
from oauth2client.service_account import ServiceAccountCredentials
from flask import Flask, request
from dotenv import load_dotenv
from apscheduler.schedulers.background import BackgroundScheduler
from collections import defaultdict

# ==================================================
# ♠️ Inisialisasi & Konfigurasi Global
# ==================================================
# Muat environment variables dari file .env
load_dotenv()

# Konfigurasi logging untuk debugging yang lebih baik
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Inisialisasi Aplikasi Web Flask
app = Flask(__name__)

# Konfigurasi dari file .env
TOKEN = os.getenv('TELEGRAM_TOKEN')
SS_ID = os.getenv('SS_ID')
ADMIN_ID = os.getenv('ADMIN_ID')
SALES_GRP_ID = os.getenv('SALES_GRP_ID')
SALES_THREAD_ID = os.getenv('SALES_THREAD_ID')
AGENT_GROUP_ID = os.getenv('AGENT_GROUP_ID')
OLD_TRANSFER_GROUP_ID = os.getenv('OLD_TRANSFER_GROUP_ID')
AGENT_NOTIF_GROUP_ID = os.getenv('AGENT_NOTIF_GROUP_ID')
AGENT_NOTIF_THREAD_ID = os.getenv('AGENT_NOTIF_THREAD_ID')

# Lookup Ekspedisi (Sama seperti di JS)
SHIP_LOOKUP = {
    "id": "ID Express", "idx": "ID Express", "id express": "ID Express",
    "sap": "SAP Logistic", "ninja": "Ninja Xpress", "jne": "JNE",
    "jnt": "J&T Express", "j&t": "J&T Express", "spx": "SPX"
}

# Path file untuk menyimpan ID pesan (pengganti PropertiesService)
MESSAGE_ID_STORE_PATH = 'message_id_store.json'

# Inisialisasi Bot Telegram
try:
    bot = telegram.Bot(token=TOKEN)
    logger.info("Bot Telegram berhasil diinisialisasi.")
except Exception as e:
    logger.error(f"Gagal menginisialisasi bot: {e}", exc_info=True)

# Koneksi ke Google Sheets
try:
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SS_ID)
    closing_sheet = spreadsheet.worksheet("Closing")
    closing_mp_sheet = spreadsheet.worksheet("Closing MP")
    agen_sheet = spreadsheet.worksheet("AGEN")
    logger.info(f"Berhasil terhubung ke Google Sheet: {spreadsheet.title}")
except Exception as e:
    logger.error(f"GAGAL terhubung ke Google Sheets: {e}", exc_info=True)

# Cache sederhana untuk message ID yang sudah diproses
processed_messages = {}


# ==================================================
# ♠️ Fungsi Inti & Helper
# ==================================================
def send_msg(chat_id, text, thread_id=None, parse_mode=None):
    """Mengirim pesan teks ke chat/thread tertentu."""
    try:
        bot.send_message(chat_id=chat_id, text=text, message_thread_id=thread_id, parse_mode=parse_mode)
    except Exception as e:
        logger.error(f"Gagal mengirim pesan ke {chat_id}: {e}")

def fwd_msg(to_id, from_id, msg_id):
    """Meneruskan (forward) pesan."""
    try:
        bot.forward_message(chat_id=to_id, from_chat_id=from_id, message_id=msg_id)
    except Exception as e:
        logger.error(f"Gagal forward pesan {msg_id} ke {to_id}: {e}")
        
def delete_msg(chat_id, message_id):
    """Menghapus pesan."""
    try:
        bot.delete_message(chat_id=chat_id, message_id=message_id)
        logger.info(f"Berhasil menghapus pesan {message_id} dari chat {chat_id}")
    except Exception as e:
        logger.warning(f"Gagal menghapus pesan {message_id} dari chat {chat_id}: {e}")

def get_num(val_str):
    """Mengambil angka dari string (misal: 'Rp 50k')."""
    if not val_str: return 0
    match = re.search(r'Rp?\s*([\d.,]+)\s*k?', str(val_str), re.IGNORECASE)
    if not match: return 0
    cleaned = match.group(1).replace('.', '').replace(',', '')
    num = float(cleaned)
    if 'k' in match.group(0).lower():
        num *= 1000
    return int(num)

def format_phone_number(phone):
    """Membersihkan dan memformat nomor HP ke format 62."""
    if not phone: return ""
    cleaned = re.sub(r'\D', '', str(phone))
    if cleaned.startswith("0"):
        return "62" + cleaned[1:]
    if not cleaned.startswith("62"):
        return "62" + cleaned
    return cleaned

def format_date(dt_obj, fmt):
    """Memformat objek datetime (padanan formatDate GAS)."""
    # Hanya implementasi format yang dibutuhkan
    if fmt == "dd/MM/yyyy":
        return dt_obj.strftime("%d/%m/%Y")
    if fmt == "EEEE, dd MMMM yyyy":
        days = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
        months = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
        return f"{days[dt_obj.weekday()]}, {dt_obj.day} {months[dt_obj.month]} {dt_obj.year}"
    return dt_obj.strftime(fmt)

# ==================================================
# ♠️ Logika Pemrosesan Pesanan (Parsing, Validasi, dll)
# ==================================================
def parse_order(text):
    lines = text.split('\n')
    data = {}
    current_key = ""
    notes_lines = []
    
    is_mp = bool(re.search(r'SALES\s+\d+\s*-\s*(SHOPEE|LAZADA|TIKTOK|TIK TOK)', text, re.IGNORECASE))
    
    if is_mp and len(lines) > 1:
        data['order_id'] = lines[1].strip()
        lines.pop(1)
    
    for line in lines:
        if not line.strip(): continue
        
        if ":" in line:
            key, val = [x.strip() for x in line.split(":", 1)]
            key = key.lower()
            current_key = key
            
            if key == "doorasi":
                data['qty_box'] = int(m.group(1)) if (m := re.search(r'(\d+)\s+Box', val, re.I)) else 0
                data['qty_sachet'] = int(m.group(1)) if (m := re.search(r'(\d+)\s+Sachet', val, re.I)) else 0
                data['price'] = get_num(val)
            elif key == "sku": data['sku'] = val
            elif key == "total pembayaran": data['total pembayaran'] = get_num(val)
            elif key == "ongkir": data['ongkir'] = get_num(val)
            elif key == "ekspedisi":
                parts = [p.strip() for p in val.split("-", 1)]
                data['ekspedisi'] = SHIP_LOOKUP.get(parts[0].lower(), parts[0])
                data['pembayaran'] = parts[1] if len(parts) > 1 else "-"
            elif key == "nama": data['nama'] = val
            elif key == "no hp": data['no hp'] = format_phone_number(val)
            elif key == "alamat jalan": data['alamat'] = val
            elif key == "desa/kelurahan": data['kelurahan'] = val
            elif key == "kecamatan": data['kecamatan'] = val
            elif key == "kab/kota": data['kota/kab'] = val
            elif key == "kode pos": data['kode pos'] = val
            else: data[key] = val
        else:
            if current_key in ["alamat jalan", "desa/kelurahan", "kecamatan", "kab/kota"]:
                data[current_key] += "\n" + line.strip()
            else: notes_lines.append(line.strip())

    if not is_mp and notes_lines:
        data['notes'] = "\n".join(notes_lines)
    
    for k in ['qty_box', 'qty_sachet', 'total pembayaran', 'ongkir']: data.setdefault(k, 0)
    for k in ['ekspedisi', 'pembayaran', 'kelurahan', 'kecamatan', 'kota/kab', 'kode pos', 'notes', 'nama', 'no hp', 'alamat', 'sku']: data.setdefault(k, "")
    return data

def validate_order(text):
    errs = []
    required = ["Doorasi:", "SKU:", "Ongkir:", "Total Pembayaran:", "Nama:", "No HP:", "Alamat Jalan:", "Desa/Kelurahan:"]
    for s in required:
        if s.lower() not in text.lower(): errs.append(f"🚨 Missing '{s}'")
    
    data = parse_order(text)
    if not data['sku']: errs.append("🚨 Invalid SKU format")
    sku = data['sku']
    if sku.startswith("DRSBOX-") and data.get('qty_box') != int(sku.split("-")[1]): errs.append("🚨 Box qty mismatch")
    elif sku.startswith("DRSA-") and data.get('qty_sachet') != int(sku.split("-")[1]): errs.append("🚨 Sachet qty mismatch")
    
    if data['pembayaran'].upper() not in ["COD", "TRANSFER"]:
        errs.append("🚨 Invalid payment method. Must be COD or TRANSFER.")
        
    return "\n".join(errs) if errs else None

# ==================================================
# ♠️ Interaksi Sheet & Data Processing
# ==================================================
def is_valid_agent_code(notes):
    if not isinstance(notes, str) or not notes: return True
    match = re.search(r'^Agen\s+[\w\s]+#\d+', notes, re.IGNORECASE)
    if not match: return True
    agent_code_from_note = match.group(0).strip()
    try:
        codes = agen_sheet.col_values(2)[1:] # Ambil semua dari kolom B, kecuali header
        return agent_code_from_note in [c.strip() for c in codes]
    except Exception as e:
        logger.error(f"Gagal membaca sheet AGEN: {e}")
        return True

def is_dup_phone(phone, addr, sheet):
    try:
        today_str = datetime.now().strftime("%d/%m/%Y")
        all_data = sheet.get_all_records()
        for row in all_data:
            if str(row.get("TANGGAL INPUT", "")).startswith(today_str):
                row_phone = format_phone_number(row.get("WHATSAPP"))
                row_addr = str(row.get("ALAMAT", "")).strip().lower()
                if row_phone == phone or row_addr == addr.lower():
                    return f"Duplicate: {row.get('WHATSAPP')}"
        return False
    except Exception as e:
        logger.error(f"Error saat cek duplikat: {e}")
        return False

def get_combined_stats(cs_name=None):
    stats = {'invoices': 0, 'box': 0, 'sachet': 0}
    today_str = datetime.now().strftime("%d/%m/%Y")
    for sheet in [closing_sheet, closing_mp_sheet]:
        try:
            for row in sheet.get_all_records():
                if str(row.get("TANGGAL INPUT", "")).startswith(today_str):
                    if not cs_name or row.get("CUSTOMER SERVICE", "") == cs_name:
                        stats['invoices'] += 1
                        stats['box'] += int(row.get("QTY BOX", 0) or 0)
                        stats['sachet'] += int(row.get("QTY SACHET", 0) or 0)
        except Exception as e:
            logger.error(f"Gagal mengambil stats dari {sheet.title}: {e}")
    return stats

# ==================================================
# ♠️ Fungsi Notifikasi & Laporan (Terjadwal)
# ==================================================
def _load_message_ids():
    """Membaca ID pesan dari file JSON."""
    if not os.path.exists(MESSAGE_ID_STORE_PATH):
        return {}
    with open(MESSAGE_ID_STORE_PATH, 'r') as f:
        return json.load(f)

def _save_message_ids(data):
    """Menyimpan ID pesan ke file JSON."""
    with open(MESSAGE_ID_STORE_PATH, 'w') as f:
        json.dump(data, f)

def send_available_agents():
    """Mengirim daftar agen yang tersedia (padanan sendAvailableAgents)."""
    logger.info("Menjalankan tugas: Mengirim daftar agen.")
    try:
        agents = agen_sheet.col_values(2)[1:]
        agents = sorted([agent.strip() for agent in agents if agent and agent.strip()])
        if not agents:
            send_msg(AGENT_NOTIF_GROUP_ID, "🚨 Tidak ada agen yang tersedia di sheet 'AGEN'!", AGENT_NOTIF_THREAD_ID)
            return

        date_str = format_date(datetime.now(), "EEEE, dd MMMM yyyy")
        header = f"📋 Daftar Agen Tersedia\n📅 Tanggal: {date_str}\n\n"
        agent_lines = [f"{i+1}. {agent}" for i, agent in enumerate(agents)]
        total = f"\nTotal: {len(agents)} Agen"
        
        # Hapus pesan lama
        store = _load_message_ids()
        old_ids = store.get('agent_notif_ids', [])
        for msg_id in old_ids:
            delete_msg(AGENT_NOTIF_GROUP_ID, msg_id)

        # Kirim pesan baru (untuk simpel, tidak dipecah seperti di GAS)
        full_message = header + "\n".join(agent_lines) + total
        msg = bot.send_message(chat_id=AGENT_NOTIF_GROUP_ID, text=full_message, message_thread_id=AGENT_NOTIF_THREAD_ID)
        
        # Simpan ID pesan baru
        store['agent_notif_ids'] = [msg.message_id]
        _save_message_ids(store)
    except Exception as e:
        logger.error(f"Gagal mengirim daftar agen: {e}", exc_info=True)
        send_msg(ADMIN_ID, f"Bot Error (Send Agents): {e}")
        
def send_sales_report():
    """Mengirim laporan penjualan harian/mingguan/bulanan (padanan sendSalesReport)."""
    logger.info("Menjalankan tugas: Mengirim laporan penjualan.")
    try:
        today = date.today()
        week_start = today - timedelta(days=6) # 7 hari terakhir termasuk hari ini
        month_start = today.replace(day=1)

        daily_stats, weekly_stats, monthly_stats = defaultdict(lambda: {'b': 0, 's': 0, 'i': 0}), defaultdict(lambda: {'b': 0, 's': 0, 'i': 0}), defaultdict(lambda: {'b': 0, 's': 0, 'i': 0})
        
        for sheet in [closing_sheet, closing_mp_sheet]:
            for row in sheet.get_all_records():
                try:
                    row_date = datetime.strptime(str(row.get("TANGGAL", "")), "%d/%m/%Y").date()
                    cs = str(row.get("CUSTOMER SERVICE", "Unknown")).replace("DOORASI ", "").strip()
                    box = int(row.get("QTY BOX", 0) or 0)
                    sachet = int(row.get("QTY SACHET", 0) or 0)

                    if row_date == today:
                        daily_stats[cs]['b'] += box; daily_stats[cs]['s'] += sachet; daily_stats[cs]['i'] += 1
                    if week_start <= row_date <= today:
                        weekly_stats[cs]['b'] += box; weekly_stats[cs]['s'] += sachet; weekly_stats[cs]['i'] += 1
                    if month_start <= row_date <= today:
                        monthly_stats[cs]['b'] += box; monthly_stats[cs]['s'] += sachet; monthly_stats[cs]['i'] += 1
                except (ValueError, TypeError): continue
        
        def format_rank_section(title, stats_dict, period_str=""):
            if not stats_dict: return f"{title} {period_str}\nTidak ada data.\n\n"
            
            total = {'b': 0, 's': 0, 'i': 0}
            ranked_list = []
            for cs, data in stats_dict.items():
                total_box = data['b'] + data['s'] // 5
                rem_sachet = data['s'] % 5
                total['b'] += total_box; total['s'] += rem_sachet; total['i'] += data['i']
                ranked_list.append({'n': cs, 'b': total_box, 's': rem_sachet, 'i': data['i']})
            
            ranked_list.sort(key=lambda x: (x['b'], x['s']), reverse=True)
            
            lines = [f"{i+1}. {d['n']} | {d['b']} Box - {d['s']} Sachet ({d['i']} Inv)" for i, d in enumerate(ranked_list)]
            total_line = f"TOTAL: {total['i']} Invoice | {total['b']} Box | {total['s']} Sachet"
            return f"{title} {period_str}\n" + "\n".join(lines) + f"\n{total_line}\n\n"

        today_str = format_date(datetime.now(), "dd/MM/yyyy")
        msg = f"🏆 Laporan Penjualan CS\n📅 Tanggal: {format_date(datetime.now(), 'EEEE, dd MMMM yyyy')}\n\n"
        msg += format_rank_section("▶︎ Daily", daily_stats, f"({today_str})")
        msg += format_rank_section("▶︎ Mingguan (7 Hari Terakhir)", weekly_stats, f"({format_date(week_start, 'dd/MM/yyyy')} - {today_str})")
        msg += format_rank_section(f"▶︎ Bulanan ({month_name[today.month]})", monthly_stats)

        # Hapus pesan lama dan kirim yang baru
        store = _load_message_ids()
        today_key = f"sales_report_{today.isoformat()}"
        old_id = store.get(today_key)
        if old_id: delete_msg(SALES_GRP_ID, old_id)

        sent_msg = bot.send_message(chat_id=SALES_GRP_ID, text=msg.strip(), message_thread_id=SALES_THREAD_ID)
        store[today_key] = sent_msg.message_id
        _save_message_ids(store)

    except Exception as e:
        logger.error(f"Gagal mengirim laporan penjualan: {e}", exc_info=True)
        send_msg(ADMIN_ID, f"Bot Error (Sales Report): {e}")

def send_closing_reminder():
    now = datetime.now()
    msg = ""
    if now.hour == 11: msg = "⏰ Segera input SEMUA Closingan Pagi ini sebelum Jam 12:00 Siang!"
    elif now.hour == 14: msg = "⏰ Segera input SEMUA Closingan Siang ini sebelum Jam 15:00 Sore!"
    elif now.hour == 18: msg = "⏰ Segera input SEMUA Closingan Sore ini sebelum Jam 19:00 Malam!"
    else: return
        
    topic_remind = [1, 1875, 5838, 19334]
    for thread_id in topic_remind:
        send_msg(SALES_GRP_ID, msg, thread_id)
        logger.info(f"Mengirim reminder ke grup {SALES_GRP_ID}, thread {thread_id}")

# ==================================================
# ♠️ Webhook Utama & Konfirmasi
# ==================================================
@app.route('/webhook', methods=['POST'])
def webhook_handler():
    update_data = request.get_json()
    update = telegram.Update.de_json(update_data, bot)
    message = update.message or update.edited_message
    if not message: return 'ok', 200

    msg_id, chat_id, thread_id = message.message_id, message.chat.id, message.message_thread_id
    user = message.from_user
    cs_name = f"{user.first_name} {user.last_name or ''}".strip()
    text = message.text or message.caption or ""

    if is_msg_processed(msg_id):
        logger.warning(f"Pesan {msg_id} sudah diproses, diabaikan.")
        return 'ok', 200

    if not text or not text.lower().startswith("sales"):
        if message.photo and "transfer" in text.lower():
             send_msg(chat_id, "✅ Bukti transfer diterima. Kirim detail pesanan dengan format SALES yang valid.", thread_id)
        return 'ok', 200

    logger.info(f"Memproses pesanan dari {cs_name} (Chat ID: {chat_id})")
    mark_msg_processed(msg_id) # Tandai setelah filter awal, sebelum proses berat
    
    try:
        platform_match = re.search(r'SALES\s+\d+\s*-\s*(SHOPEE|LAZADA|TIKTOK|TIK TOK)', text, re.IGNORECASE)
        if platform_match:
            process_mp_order(chat_id, thread_id, cs_name, text, platform_match.group(1).upper())
        else:
            process_regular_order(chat_id, thread_id, cs_name, text, message)
        
        send_confirmation(chat_id, thread_id, cs_name, parse_order(text), message)
    except Exception as e:
        logger.error(f"Error fatal di webhook: {e}", exc_info=True)
        send_msg(ADMIN_ID, f"Bot Error: {e}")

    return 'ok', 200

def process_mp_order(chat_id, thread_id, cs_name, text, platform):
    result = parse_order(text)
    row_data = [
        result.get('order_id') or f"INV-MP-{int(datetime.now().timestamp())}",
        format_date(datetime.now(), "dd/MM/yyyy"), cs_name, result.get('nama'), result.get('no hp'),
        result.get('alamat'), result.get('kelurahan'), result.get('kecamatan'), result.get('kota/kab'),
        result.get('kode pos'), "DOORASI", result.get('sku'), result.get('qty_box'), result.get('qty_sachet'),
        result.get('total pembayaran'), result.get('ongkir'), result.get('ekspedisi'),
        platform, result.get('notes'), datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    ]
    closing_mp_sheet.append_row(row_data, value_input_option='USER_ENTERED')
    logger.info(f"Pesanan MP {result.get('nama')} berhasil disimpan.")

def process_regular_order(chat_id, thread_id, cs_name, text, message):
    error = validate_order(text)
    if error:
        send_msg(chat_id, error, thread_id)
        raise ValueError(f"Pesanan dari {cs_name} gagal validasi: {error}")

    result = parse_order(text)
    notes, phone, addr = result.get('notes', ""), result.get('no hp'), result.get('alamat')
    
    if "#ro" not in text.lower():
        if dup_check := is_dup_phone(phone, addr, closing_sheet):
            send_msg(chat_id, f"🚨 {dup_check} Silakan periksa kembali.", thread_id)
            raise ValueError(f"Pesanan duplikat terdeteksi: {dup_check}")
    
    if "agen" in notes.lower() and not is_valid_agent_code(notes):
        send_msg(chat_id, "🚨 Kode Agen tidak valid/terdaftar.", thread_id)
        raise ValueError(f"Kode agen tidak valid: {notes}")
    
    row_data = [
        f"INV-{int(datetime.now().timestamp())}",
        format_date(datetime.now(), "dd/MM/yyyy"), cs_name, result.get('nama'), phone, addr,
        result.get('kelurahan'), result.get('kecamatan'), result.get('kota/kab'), result.get('kode pos'), 
        "DOORASI", result.get('sku'), result.get('qty_box'), result.get('qty_sachet'),
        result.get('total pembayaran'), result.get('ongkir'), result.get('ekspedisi'),
        result.get('pembayaran'), notes, datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    ]
    closing_sheet.append_row(row_data, value_input_option='USER_ENTERED')
    logger.info(f"Pesanan reguler {result.get('nama')} berhasil disimpan.")

def send_confirmation(chat_id, thread_id, cs_name, result, message):
    stats = get_combined_stats(cs_name)
    msg_header = f"{cs_name} ★ {stats['invoices']} INVOICE - {stats['box']} Box - {stats['sachet']} Sachet"
    pay_method = result.get('pembayaran', "").upper()
    notes = result.get('notes', "").lower()
    
    if "agen" in notes:
        user_msg = f"{msg_header}\n\n🥷🏻 Data ter-supply untuk {result.get('notes')}"
        send_msg(chat_id, user_msg, thread_id)
        if pay_method == "TRANSFER" and message.photo:
            fwd_msg(AGENT_GROUP_ID, chat_id, message.message_id)
    else:
        if pay_method == "TRANSFER":
            user_msg = f"{msg_header}\n\n🏧 Orderan Transfer {result.get('nama')} diterima"
            send_msg(chat_id, user_msg, thread_id)
            if message.photo:
                fwd_msg(OLD_TRANSFER_GROUP_ID, chat_id, message.message_id)
        else:
            user_msg = f"{msg_header}\n\n✅ Success! {result.get('nama')} berhasil diinput"
            send_msg(chat_id, user_msg, thread_id)

# ==================================================
# ♠️ Titik Mulai Aplikasi & Scheduler
# ==================================================
if __name__ == "__main__":
    # Inisialisasi Scheduler
    scheduler = BackgroundScheduler(daemon=True, timezone='Asia/Jakarta')
    scheduler.add_job(send_closing_reminder, 'cron', hour='11,14,18', minute='0')
    scheduler.add_job(send_sales_report, 'cron', hour='21', minute='0') # Laporan penjualan jam 9 malam
    scheduler.add_job(send_available_agents, 'cron', hour='8', minute='0') # Notif agen jam 8 pagi
    scheduler.start()
    
    logger.info("Scheduler berhasil dimulai.")
    
    # Menjalankan server web Flask
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
