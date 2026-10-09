import os
import random
import sqlite3
import threading
from flask import Flask
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    CallbackQueryHandler,
    ContextTypes,
    MessageHandler,
    filters
)

# --- BOT BİLGİLERİ ---
TOKEN = "8899348135:AAH2G6dvKUSaSXuNI0p1PT3_JDCMgjRXaGo"
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

# --- VERİTABANI OLUŞTURMA VE KELİME YÜKLEME ---
def init_database():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    # Kelimeler tablosu
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS words (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            english TEXT UNIQUE,
            turkish TEXT
        )
    """)
    
    # Kullanıcı istatistik tablosu
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS user_stats (
            user_id INTEGER PRIMARY KEY,
            correct_count INTEGER DEFAULT 0,
            wrong_count INTEGER DEFAULT 0
        )
    """)
    
    # Eğer kelimeler henüz yüklenmediyse Excel'den aktar
    cursor.execute("SELECT COUNT(*) FROM words")
    count = cursor.fetchone()[0]
    
    if count == 0 and os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
            for _, row in df.iterrows():
                eng = str(row['English']).strip()
                tr = str(row['Türkçe']).strip()
                if eng and tr and eng.lower() != 'the oxford':
                    cursor.execute("INSERT OR IGNORE INTO words (english, turkish) VALUES (?, ?)", (eng, tr))
            conn.commit()
            print("Kelimeler basariyla veritabanina yuklendi.")
        except Exception as e:
            print(f"Excel aktarim hatasi: {e}")
            
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
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_name = update.effective_user.first_name
    text = (
        f"Selam {user_name}! 🇬🇧\n\n"
        "Oxford 3000 kelimelik İngilizce botuna hoş geldin!\n\n"
        "Neler yapabilirsin?\n"
        "🔹 /kelime - Yeni bir kelime testi çözer\n"
        "🔹 /skor - Doğru/Yanlış durumunu görürsün\n"
        "🔹 /yardim - Bilgi alırsın"
    )
    await update.message.reply_text(text)

async def kelime_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    eng, tr, options = get_quiz_question()
    if not eng:
        await update.message.reply_text("Kelimeler hazırlanıyor veya bulunamadı.")
        return

    # Şıkları buton yap
    keyboard = []
    for opt in options:
        status = "C" if opt == tr else "W"
        keyboard.append([InlineKeyboardButton(opt, callback_data=f"{status}:{eng}:{tr}")])
    
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text(
        f"🎯 **{eng}** kelimesinin Türkçe karşılığı nedir?",
        reply_markup=reply_markup,
        parse_mode="Markdown"
    )

async def handle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    user_id = query.from_user.id
    status, eng, tr = query.data.split(":", 2)
    
    if status == "C":
        update_user_score(user_id, True)
        await query.edit_message_text(
            f"✅ **Doğru! Harikasın!**\n\n"
            f"🇬🇧 **{eng}** = 🇹🇷 **{tr}**\n\n"
            f"Yeni kelime için 👉 /kelime",
            parse_mode="Markdown"
        )
    else:
        update_user_score(user_id, False)
        await query.edit_message_text(
            f"❌ **Yanlış cevap.**\n\n"
            f"🇬🇧 **{eng}** kelimesinin doğrusu: 🇹🇷 **{tr}**\n\n"
            f"Yeni kelime için 👉 /kelime",
            parse_mode="Markdown"
        )

async def skor_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    c, w = get_user_score(user_id)
    total = c + w
    rate = f"%{int((c / total) * 100)}" if total > 0 else "%0"
    
    msg = (
        f"📊 **Senin İstatistiklerin**\n\n"
        f"✅ Doğru: {c}\n"
        f"❌ Yanlış: {w}\n"
        f"📈 Başarı Oranı: {rate}\n"
        f"🎯 Toplam Çözülen: {total}"
    )
    await update.message.reply_text(msg, parse_mode="Markdown")

async def yardim_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "💡 **Kullanım Kılavuzu:**\n\n"
        "/kelime yazarak kelime testine başlayabilirsin.\n"
        "/skor yazarak güncel durumunu görebilirsin."
    )
    await update.message.reply_text(text, parse_mode="Markdown")

def main():
    # 1. Veritabanını hazırla
    init_database()
    
    # 2. Flask web sunucusunu arka planda başlat
    flask_thread = threading.Thread(target=run_flask)
    flask_thread.daemon = True
    flask_thread.start()
    
    # 3. Telegram botunu başlat
    application = ApplicationBuilder().token(TOKEN).build()
    
    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(CommandHandler("kelime", kelime_command))
    application.add_handler(CommandHandler("skor", skor_command))
    application.add_handler(CommandHandler("yardim", yardim_command))
    application.add_handler(CallbackQueryHandler(handle_callback))
    
    print("Bot basariyla calisiyor...")
    application.run_polling()

if __name__ == "__main__":
    main()
