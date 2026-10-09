import os
import random
import sqlite3
import threading
from flask import Flask
from openpyxl import load_workbook
import telebot
from telebot import types

# --- BOT BİLGİLERİ ---
TOKEN = os.environ.get("BOT_TOKEN", "8899348135:AAH2G6dvKUSaSXuNI0p1PT3_JDCMgjRXaGo")
bot = telebot.TeleBot(TOKEN)

DB_FILE = "english_bot.db"
EXCEL_FILE = "translated.xlsx"

# --- FLASK WEB SUNUCUSU (RENDER & UPTIMEROBOT İÇİN) ---
app = Flask(__name__)

@app.route('/')
def home():
    return "English Coach AI Bot Aktif!"

@app.route('/health')
def health():
    return "OK", 200

def run_flask():
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)

# --- VERİTABANI İŞLEMLERİ ---
def init_database():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS words (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            english TEXT UNIQUE,
            turkish TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS user_stats (
            user_id INTEGER PRIMARY KEY,
            correct_count INTEGER DEFAULT 0,
            wrong_count INTEGER DEFAULT 0
        )
    """)
    
    cursor.execute("SELECT COUNT(*) FROM words")
    count = cursor.fetchone()[0]
    
    if count == 0 and os.path.exists(EXCEL_FILE):
        try:
            wb = load_workbook(EXCEL_FILE, read_only=True)
            sheet = wb.active
            rows = list(sheet.iter_rows(values_only=True))
            for row in rows[1:]:
                if row and len(row) >= 2:
                    eng = str(row[0]).strip() if row[0] is not None else ""
                    tr = str(row[1]).strip() if row[1] is not None else ""
                    if eng and tr and eng.lower() != 'the oxford':
                        cursor.execute("INSERT OR IGNORE INTO words (english, turkish) VALUES (?, ?)", (eng, tr))
            conn.commit()
            print("Kelimeler basariyla veritabanina aktarildi.")
        except Exception as e:
            print(f"Excel hatasi: {e}")
            
    conn.close()

def get_quiz_question():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT english, turkish FROM words ORDER BY RANDOM() LIMIT 4")
    rows = cursor.fetchall()
    conn.close()
    
    if len(rows) < 4:
        return None, None, []
    
    target_eng, target_tr = rows[0]
    options = [r[1] for r in rows]
    random.shuffle(options)
    return target_eng, target_tr, options

def update_user_score(user_id: int, is_correct: bool):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("INSERT OR IGNORE INTO user_stats (user_id, correct_count, wrong_count) VALUES (?, 0, 0)", (user_id,))
    if is_correct:
        cursor.execute("UPDATE user_stats SET correct_count = correct_count + 1 WHERE user_id = ?", (user_id,))
    else:
        cursor.execute("UPDATE user_stats SET wrong_count = wrong_count + 1 WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()

def get_user_score(user_id: int):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT correct_count, wrong_count FROM user_stats WHERE user_id = ?", (user_id,))
    res = cursor.fetchone()
    conn.close()
    if res:
        return res[0], res[1]
    return 0, 0

# --- BOT KOMUTLARI ---
@bot.message_handler(commands=['start'])
def send_welcome(message):
    name = message.from_user.first_name
    text = (
        f"Selam {name}! 🇬🇧\n\n"
        "Oxford 3000 kelimelik İngilizce koçuna hoş geldin!\n\n"
        "📌 Komutlar:\n"
        "🔹 /kelime - Yeni bir test sorusu sorar\n"
        "🔹 /skor - İstatistiklerini gösterir\n"
        "🔹 /yardim - Bilgi verir"
    )
    bot.reply_to(message, text)

@bot.message_handler(commands=['kelime'])
def ask_word(message):
    eng, tr, options = get_quiz_question()
    if not eng:
        bot.reply_to(message, "Kelimeler yükleniyor, lütfen birazdan tekrar dene.")
        return

    markup = types.InlineKeyboardMarkup(row_width=1)
    for opt in options:
        status = "C" if opt == tr else "W"
        # Callback verisi sınırı için güvenli veri
        cb_data = f"{status}|{eng[:15]}|{tr[:20]}"
        markup.add(types.InlineKeyboardButton(opt, callback_data=cb_data))

    bot.send_message(message.chat.id, f"🎯 **{eng}** kelimesinin Türkçe karşılığı nedir?", reply_markup=markup, parse_mode="Markdown")

@bot.callback_query_handler(func=lambda call: True)
def handle_answer(call):
    parts = call.data.split("|", 2)
    status = parts[0]
    eng = parts[1]
    tr = parts[2]
    user_id = call.from_user.id

    if status == "C":
        update_user_score(user_id, True)
        bot.edit_message_text(
            f"✅ **Doğru! Harika!**\n\n🇬🇧 **{eng}** = 🇹🇷 **{tr}**\n\nYeni kelime için 👉 /kelime",
            chat_id=call.message.chat.id,
            message_id=call.message.message_id,
            parse_mode="Markdown"
        )
    else:
        update_user_score(user_id, False)
        bot.edit_message_text(
            f"❌ **Yanlış cevap!**\n\n🇬🇧 **{eng}** kelimesinin doğrusu: 🇹🇷 **{tr}**\n\nYeni kelime için 👉 /kelime",
            chat_id=call.message.chat.id,
            message_id=call.message.message_id,
            parse_mode="Markdown"
        )

@bot.message_handler(commands=['skor'])
def show_score(message):
    c, w = get_user_score(message.from_user.id)
    total = c + w
    rate = f"%{int((c / total) * 100)}" if total > 0 else "%0"
    msg = (
        f"📊 **Senin İstatistiklerin**\n\n"
        f"✅ Doğru: {c}\n"
        f"❌ Yanlış: {w}\n"
        f"📈 Başarı Oranı: {rate}\n"
        f"🎯 Toplam Çözülen: {total}"
    )
    bot.reply_to(message, msg, parse_mode="Markdown")

@bot.message_handler(commands=['yardim'])
def send_help(message):
    bot.reply_to(message, "Yeni bir kelime çözmek için /kelime yazabilirsin!")

if __name__ == "__main__":
    init_database()
    
    # Web sunucusunu başlat
    t = threading.Thread(target=run_flask)
    t.daemon = True
    t.start()
    
    print("Bot basariyla baslatildi...")
    bot.infinity_polling()
