import os
import random
import sqlite3
import threading
from flask import Flask
from openpyxl import load_workbook
import telebot
from telebot import types

# --- BOT VE ÇEVRESEL AYARLAR ---
TOKEN = os.environ.get("BOT_TOKEN", "8899348135:AAH2G6dvKUSaSXuNI0p1PT3_JDCMgjRXaGo")
bot = telebot.TeleBot(TOKEN)

DB_FILE = "english_bot.db"
EXCEL_FILE = "translated.xlsx"

# Aktif sınav durumlarını tutan bellek sözlüğü: {user_id: {"q_idx": 0, "score": 0}}
active_placement_tests = {}

# --- FLASK WEB SUNUCUSU (RENDER UPTIME) ---
app = Flask(__name__)

@app.route('/')
def home():
    return "English Coach AI Platform Aktif!"

@app.route('/health')
def health():
    return "OK", 200

def run_flask():
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)

# --- STATİK VERİLER: GRAMER, KATEGORİLER VE SEVİYE SINAVI ---

PLACEMENT_QUESTIONS = [
    {"q": "She _______ from Spain.", "options": ["is", "are", "am", "be"], "a": "is", "lvl": "A1"},
    {"q": "I _______ a movie yesterday night.", "options": ["watched", "watch", "watching", "have watched"], "a": "watched", "lvl": "A1"},
    {"q": "There aren't _______ apples left in the basket.", "options": ["any", "some", "much", "an"], "a": "any", "lvl": "A2"},
    {"q": "He was cooking dinner while she _______ a book.", "options": ["was reading", "reads", "has read", "is reading"], "a": "was reading", "lvl": "A2"},
    {"q": "If it rains tomorrow, we _______ the match.", "options": ["will cancel", "canceled", "cancel", "would cancel"], "a": "will cancel", "lvl": "B1"},
    {"q": "I have lived in this city _______ 2018.", "options": ["since", "for", "in", "during"], "a": "since", "lvl": "B1"},
    {"q": "By the time the police arrived, the thieves _______.", "options": ["had escaped", "escaped", "have escaped", "escape"], "a": "had escaped", "lvl": "B2"},
    {"q": "If I _______ you, I would take that offer immediately.", "options": ["were", "am", "will be", "had been"], "a": "were", "lvl": "B2"},
    {"q": "Seldom _______ such an extraordinary performance.", "options": ["have I seen", "I saw", "I have seen", "saw I"], "a": "have I seen", "lvl": "C1"},
    {"q": "The manager insisted that the report _______ submitted by noon.", "options": ["be", "is", "was", "will be"], "a": "be", "lvl": "C1"}
]

CATEGORY_DATA = {
    "travel": {
        "title": "✈️ Seyahat & Yurtdışı Tatili",
        "words": [
            ("passport", "pasaport"), ("boarding pass", "biniş kartı"), ("customs", "gümrük"),
            ("baggage claim", "bagaj alım yeri"), ("flight attendant", "uçuş görevlisi"),
            ("delay", "gecikme"), ("currency exchange", "döviz bürosu"), ("reservation", "rezervasyon"),
            ("departure", "kalkış / gidiş"), ("arrival", "varış")
        ],
        "sentences": [
            ("Where is the baggage claim area?", "Bagaj alım alanı nerede?"),
            ("Here is my passport and boarding pass.", "Buyrun pasaportum ve biniş kartım."),
            ("I would like to check in for my flight.", "Uçuşum için check-in yaptırmak istiyorum."),
            ("I am here for vacation for one week.", "Bir haftalığına tatil için buradayım."),
            ("Could you call a taxi for me, please?", "Benim için bir taksi çağırabilir misiniz lütfen?")
        ]
    },
    "business": {
        "title": "💼 İş Dünyası & Mülakat",
        "words": [
            ("deadline", "son teslim tarihi"), ("meeting", "toplantı"), ("negotiate", "müzakere etmek"),
            ("colleague", "iş arkadaşı"), ("salary", "maaş"), ("contract", "sözleşme"),
            ("promotion", "terfi"), ("schedule", "çalışma takvimi"), ("client", "müşteri")
        ],
        "sentences": [
            ("We need to meet the deadline tomorrow.", "Yarınki son teslim tarihine yetişmemiz gerekiyor."),
            ("Let's schedule a meeting for Friday.", "Cuma günü için bir toplantı planlayalım."),
            ("I am responsible for client communication.", "Müşteri iletişiminden ben sorumluyum."),
            ("Could you send me the contract via email?", "Sözleşmeyi bana e-posta ile gönderebilir misiniz?")
        ]
    },
    "food": {
        "title": "🍽 Restoran & Yeme-İçme",
        "words": [
            ("menu", "menü"), ("bill / check", "hesap"), ("waiter", "garson"),
            ("delicious", "lezzetli"), ("spicy", "baharatlı / acı"), ("dessert", "tatlı"),
            ("beverage", "içecek"), ("tip", "bahşiş"), ("order", "sipariş")
        ],
        "sentences": [
            ("Could we have the bill, please?", "Hesabı alabilir miyiz lütfen?"),
            ("What do you recommend for dinner?", "Akşam yemeği için ne önerirsiniz?"),
            ("I would like to order a steak with salad.", "Salata ile birlikte biftek sipariş etmek istiyorum."),
            ("Is this dish spicy?", "Bu yemek acı mı?")
        ]
    },
    "daily": {
        "title": "🏫 Günlük Yaşam & Sosyal",
        "words": [
            ("neighborhood", "mahalle / çevre"), ("groceries", "market alışverişi"),
            ("commute", "işe/okula gidip gelmek"), ("appointment", "randevu"),
            ("hobby", "hobi"), ("weather", "hava durumu")
        ],
        "sentences": [
            ("What time does the bus leave?", "Otobüs saat kaçta kalkıyor?"),
            ("I need to buy some groceries today.", "Bugün biraz market alışverişi yapmam gerek."),
            ("How was your day at university?", "Üniversitedeki günün nasıl geçti?")
        ]
    }
}

GRAMMAR_LESSONS = {
    "A1": (
        "📘 **A1 Gramer: To Be (am / is / are)**\n\n"
        "İngilizcede eylem bildirmeyen, durum ve kimlik belirten cümlelerin temel direğidir.\n\n"
        "• **I** ➔ **am** (I am a student.)\n"
        "• **He / She / It** ➔ **is** (He is tired.)\n"
        "• **You / We / They** ➔ **are** (They are ready.)\n\n"
        "💡 **Kısa Özet:** Bir eylem (koşmak, gitmek) yoksa, sadece durum/sıfat anlatıyorsan 'am/is/are' yapıştır geç!"
    ),
    "A2": (
        "📗 **A2 Gramer: Past Simple (Geçmiş Zaman -ed / V2)**\n\n"
        "Geçmişte yaşanmış, bitmiş ve net bir zamanı olan olaylar için kullanılır.\n\n"
        "• **Düzenli Fiiller:** Fiilin sonuna `-ed` gelir.\n"
        "  Örnek: *I watched TV yesterday.* (Dün TV izledim.)\n"
        "• **Düzensiz Fiiller:** Tamamen şekil değiştirir.\n"
        "  *go ➔ went*, *see ➔ saw*, *eat ➔ ate*\n\n"
        "💡 **Soru ve Olumsuzda:** 'did / didn't' devreye girer ve fiil hemen ilk haline (V1) döner!\n"
        "  *I didn't go.* (Gittim değil, gitmedim.)"
    ),
    "B1": (
        "📙 **B1 Gramer: Present Perfect Tense (have/has + V3)**\n\n"
        "Geçmişte olmuş ama **ne zaman olduğu önemli olmayan**, etkisi hala devam eden olaylar.\n\n"
        "• **I / You / We / They** ➔ **have + V3**\n"
        "• **He / She / It** ➔ **has + V3**\n\n"
        "Örnek: *I have lost my passport.* (Pasaportumu kaybettim — zamanı belli değil ama şu an elimde yok!)\n"
        "💡 **İpucu:** 'Since' (den beri) ve 'For' (boyunca) kelimeleri bu zamanın en iyi dostudur."
    ),
    "B2": (
        "📕 **B2 Gramer: Conditionals (Koşul Cümleleri - If Clauses)**\n\n"
        "Hayali durumlar, varsayımlar ve pişmanlıklar:\n\n"
        "• **Type 2 (Hayali/Şimdiki Zaman):**\n"
        "  *If I were rich, I would travel the world.* (Zengin olsaydım dünyayı gezerdim - şu an zengin değilim)\n"
        "• **Type 3 (Geçmiş Pişmanlık):**\n"
        "  *If I had studied harder, I would have passed the exam.* (Daha çok çalışsaydım geçerdim - çalışmadım ve kaldım)"
    )
}

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
        CREATE TABLE IF NOT EXISTS user_profiles (
            user_id INTEGER PRIMARY KEY,
            level TEXT DEFAULT 'A1',
            correct_count INTEGER DEFAULT 0,
            wrong_count INTEGER DEFAULT 0
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS unknown_words (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            english TEXT,
            turkish TEXT,
            UNIQUE(user_id, english)
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
            print("Oxford 3000 kelimeleri veritabanina yuklendi.")
        except Exception as e:
            print(f"Excel aktarim hatasi: {e}")
            
    conn.close()

def get_user_profile(user_id):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("INSERT OR IGNORE INTO user_profiles (user_id, level, correct_count, wrong_count) VALUES (?, 'A1', 0, 0)", (user_id,))
    conn.commit()
    cursor.execute("SELECT level, correct_count, wrong_count FROM user_profiles WHERE user_id = ?", (user_id,))
    res = cursor.fetchone()
    conn.close()
    return res

def set_user_level(user_id, level):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("UPDATE user_profiles SET level = ? WHERE user_id = ?", (level, user_id))
    conn.commit()
    conn.close()

def update_score(user_id, is_correct):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    if is_correct:
        cursor.execute("UPDATE user_profiles SET correct_count = correct_count + 1 WHERE user_id = ?", (user_id,))
    else:
        cursor.execute("UPDATE user_profiles SET wrong_count = wrong_count + 1 WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()

def add_unknown_word(user_id, eng, tr):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("INSERT OR IGNORE INTO unknown_words (user_id, english, turkish) VALUES (?, ?, ?)", (user_id, eng, tr))
    conn.commit()
    conn.close()

def remove_unknown_word(user_id, eng):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM unknown_words WHERE user_id = ? AND english = ?", (user_id, eng))
    conn.commit()
    conn.close()

def get_unknown_words_count(user_id):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM unknown_words WHERE user_id = ?", (user_id,))
    count = cursor.fetchone()[0]
    conn.close()
    return count

def get_random_unknown_quiz(user_id):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT english, turkish FROM unknown_words WHERE user_id = ? ORDER BY RANDOM() LIMIT 1", (user_id,))
    target = cursor.fetchone()
    if not target:
        conn.close()
        return None, None, []
    
    # 3 yanlis sik icin genel havuzdan sec
    cursor.execute("SELECT turkish FROM words WHERE english != ? ORDER BY RANDOM() LIMIT 3", (target[0],))
    distractors = [r[0] for r in cursor.fetchall()]
    conn.close()
    
    options = distractors + [target[1]]
    random.shuffle(options)
    return target[0], target[1], options

def get_general_quiz():
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

# --- ARAYÜZ VE MENÜLER ---

def get_main_keyboard():
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    keyboard.add(
        types.KeyboardButton("📚 Kelime Çöz"),
        types.KeyboardButton("🧠 Zayıf Kelimelerim")
    )
    keyboard.add(
        types.KeyboardButton("🗂 Kategorik Kelimeler"),
        types.KeyboardButton("📖 Gramer Dersleri")
    )
    keyboard.add(
        types.KeyboardButton("🎓 Seviye Belirleme Sınavı"),
        types.KeyboardButton("📊 Profilim & Skor")
    )
    return keyboard

# --- KOMUTLAR VE MESAJ YÖNLENDİRİCİ ---

@bot.message_handler(commands=['start'])
def handle_start(message):
    user_name = message.from_user.first_name
    get_user_profile(message.from_user.id)
    welcome_text = (
        f"Selam **{user_name}**! 🇬🇧\n\n"
        "AI Destekli İngilizce Koçuna hoş geldin! Aşağıdaki interaktif menüyü kullanarak hemen öğrenmeye başlayabilirsin."
    )
    bot.send_message(message.chat.id, welcome_text, reply_markup=get_main_keyboard(), parse_mode="Markdown")

@bot.message_handler(func=lambda msg: True)
def handle_text_menu(message):
    uid = message.from_user.id
    txt = message.text

    if txt == "📚 Kelime Çöz":
        send_quiz_question(message.chat.id, is_unknown_mode=False)

    elif txt == "🧠 Zayıf Kelimelerim":
        count = get_unknown_words_count(uid)
        if count == 0:
            bot.send_message(
                message.chat.id,
                "🎉 **Harika durumdasın!**\n\nHaznede hiç bilmediğin kelime yok. Normal kelime testleri çözerek eksiklerini keşfedebilirsin!",
                parse_mode="Markdown"
            )
        else:
            send_quiz_question(message.chat.id, is_unknown_mode=True)

    elif txt == "🗂 Kategorik Kelimeler":
        markup = types.InlineKeyboardMarkup(row_width=2)
        for cat_key, cat_val in CATEGORY_DATA.items():
            markup.add(
                types.InlineKeyboardButton(f"{cat_val['title']} (Kelimeler)", callback_data=f"catw:{cat_key}"),
                types.InlineKeyboardButton(f"🗣 Kalıp Cümleler", callback_data=f"cats:{cat_key}")
            )
        bot.send_message(message.chat.id, "📌 **Çalışmak istediğin alanı seç:**", reply_markup=markup, parse_mode="Markdown")

    elif txt == "📖 Gramer Dersleri":
        markup = types.InlineKeyboardMarkup(row_width=2)
        markup.add(
            types.InlineKeyboardButton("A1: Temel Başlangıç (To Be)", callback_data="grm:A1"),
            types.InlineKeyboardButton("A2: Geçmiş Zaman (Past Simple)", callback_data="grm:A2"),
            types.InlineKeyboardButton("B1: Present Perfect Tense", callback_data="grm:B1"),
            types.InlineKeyboardButton("B2: Koşul Cümleleri (If Clauses)", callback_data="grm:B2")
        )
        bot.send_message(message.chat.id, "📘 **Öğrenmek istediğin gramer seviyesini seç:**", reply_markup=markup, parse_mode="Markdown")

    elif txt == "🎓 Seviye Belirleme Sınavı":
        active_placement_tests[uid] = {"q_idx": 0, "score": 0}
        send_placement_question(message.chat.id, uid)

    elif txt == "📊 Profilim & Skor":
        lvl, c, w = get_user_profile(uid)
        un_count = get_unknown_words_count(uid)
        total = c + w
        rate = f"%{int((c / total) * 100)}" if total > 0 else "%0"
        
        msg = (
            f"👤 **Öğrenci Profili**\n\n"
            f"🎯 Mevcut Seviyen: **{lvl}**\n"
            f"✅ Doğru Cevaplar: **{c}**\n"
            f"❌ Yanlış Cevaplar: **{w}**\n"
            f"📈 Genel Başarı: **{rate}**\n"
            f"🧠 Tekrar Listendeki Kelimeler: **{un_count} adet**"
        )
        bot.send_message(message.chat.id, msg, parse_mode="Markdown")

# --- KELİME TESTİ FONKSİYONU ---

def send_quiz_question(chat_id, is_unknown_mode=False):
    if is_unknown_mode:
        eng, tr, options = get_random_unknown_quiz(chat_id)
        mode_prefix = "U" # Unknown mode
    else:
        eng, tr, options = get_general_quiz()
        mode_prefix = "G" # General mode

    if not eng:
        bot.send_message(chat_id, "Kelime havuzu hazırlanıyor, lütfen tekrar deneyin.")
        return

    markup = types.InlineKeyboardMarkup(row_width=1)
    for opt in options:
        status = "C" if opt == tr else "W"
        cb = f"Q:{mode_prefix}:{status}:{eng[:14]}:{tr[:14]}"
        markup.add(types.InlineKeyboardButton(opt, callback_data=cb))
    
    # Bilmiyorum butonu
    idk_cb = f"Q:{mode_prefix}:IDK:{eng[:14]}:{tr[:14]}"
    markup.add(types.InlineKeyboardButton("🤷‍♂️ Bilmiyorum", callback_data=idk_cb))

    title = "🧠 **[Zayıf Kelime Tekrarı]**\n" if is_unknown_mode else "🎯 "
    bot.send_message(chat_id, f"{title}**{eng}** kelimesinin Türkçe karşılığı nedir?", reply_markup=markup, parse_mode="Markdown")

# --- SEVİYE TESPİT SINAVI FONKSİYONU ---

def send_placement_question(chat_id, user_id):
    state = active_placement_tests.get(user_id)
    if not state:
        return
    
    idx = state["q_idx"]
    if idx >= len(PLACEMENT_QUESTIONS):
        # Sınav bitti
        score = state["score"]
        if score <= 2:
            final_lvl = "A1"
        elif score <= 4:
            final_lvl = "A2"
        elif score <= 7:
            final_lvl = "B1"
        elif score <= 9:
            final_lvl = "B2"
        else:
            final_lvl = "C1"
            
        set_user_level(user_id, final_lvl)
        del active_placement_tests[user_id]
        
        bot.send_message(
            chat_id,
            f"🏁 **Seviye Tespit Sınavı Tamamlandı!**\n\n"
            f"📊 Doğru Sayın: **{score} / 10**\n"
            f"🎖 Belirlenen Seviyen: **{final_lvl}**\n\n"
            f"Tebrikler! Seviyen sistemde güncellendi. Artık bu seviyeye göre pratik yapabilirsin.",
            parse_mode="Markdown"
        )
        return

    q_data = PLACEMENT_QUESTIONS[idx]
    markup = types.InlineKeyboardMarkup(row_width=2)
    for opt in q_data["options"]:
        cb = f"
