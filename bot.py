import os
import random
import sqlite3
import threading
from flask import Flask
from openpyxl import load_workbook
import telebot
from telebot import types
import google.generativeai as genai

# --- BOT VE ÇEVRESEL AYARLAR ---
TOKEN = os.environ.get("BOT_TOKEN", "8899348135:AAH2G6dvKUSaSXuNI0p1PT3_JDCMgjRXaGo")
GEMINI_KEY = os.environ.get("GEMINI_API_KEY", "")

bot = telebot.TeleBot(TOKEN)

# --- İZİN VERİLEN KULLANICILAR (WHITELIST) ---
# Kardeşin de @userinfobot'tan ID'sini alınca buradaki 0 yerine onun ID'sini yazabilirsin:
ALLOWED_USERS = [8022316805, 0]

def is_authorized(user_id):
    return user_id in ALLOWED_USERS

# Gemini Yapılandırması
if GEMINI_KEY:
    genai.configure(api_key=GEMINI_KEY)
    ai_model = genai.GenerativeModel("gemini-1.5-flash")
else:
    ai_model = None

DB_FILE = "english_bot.db"
EXCEL_FILE = "translated.xlsx"

# Kullanıcı durumları
active_placement_tests = {}
ai_chat_users = set()

# --- FLASK WEB SUNUCUSU (RENDER & UPTIMEROBOT) ---
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

# --- STATİK VERİLER: SINAV, KATEGORİLER VE GRAMER ---

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
        "title": "✈️ Seyahat & Tatil",
        "words": [
            ("passport", "pasaport"), ("boarding pass", "biniş kartı"), ("customs", "gümrük"),
            ("baggage claim", "bagaj alım yeri"), ("delay", "gecikme"), ("reservation", "rezervasyon")
        ],
        "sentences": [
            ("Where is the baggage claim area?", "Bagaj alım alanı nerede?"),
            ("Here is my passport and boarding pass.", "Buyrun pasaportum ve biniş kartım."),
            ("I would like to check in for my flight.", "Uçuşum için check-in yaptırmak istiyorum."),
            ("Could you call a taxi for me, please?", "Benim için bir taksi çağırabilir misiniz lütfen?")
        ]
    },
    "business": {
        "title": "💼 İş Dünyası",
        "words": [
            ("deadline", "son teslim tarihi"), ("meeting", "toplantı"), ("negotiate", "müzakere etmek"),
            ("salary", "maaş"), ("contract", "sözleşme"), ("client", "müşteri")
        ],
        "sentences": [
            ("We need to meet the deadline tomorrow.", "Yarınki son teslim tarihine yetişmemiz gerekiyor."),
            ("Let's schedule a meeting for Friday.", "Cuma günü için bir toplantı planlayalım."),
            ("Could you send me the contract via email?", "Sözleşmeyi bana e-posta ile gönderebilir misiniz?")
        ]
    },
    "food": {
        "title": "🍽 Restoran & Yemek",
        "words": [
            ("menu", "menü"), ("bill", "hesap"), ("waiter", "garson"),
            ("delicious", "lezzetli"), ("spicy", "acılı / baharatlı"), ("dessert", "tatlı")
        ],
        "sentences": [
            ("Could we have the bill, please?", "Hesabı alabilir miyiz lütfen?"),
            ("What do you recommend for dinner?", "Akşam yemeği için ne önerirsiniz?"),
            ("Is this dish spicy?", "Bu yemek acı mı?")
        ]
    },
    "daily": {
        "title": "🏫 Günlük Yaşam",
        "words": [
            ("neighborhood", "mahalle"), ("groceries", "market alışverişi"),
            ("commute", "işe/okula gidip gelmek"), ("appointment", "randevu")
        ],
        "sentences": [
            ("What time does the bus leave?", "Otobüs saat kaçta kalkıyor?"),
            ("How was your day?", "Günün nasıl geçti?")
        ]
    }
}

GRAMMAR_LESSONS = {
    "A1": (
        "📘 **A1: To Be (am / is / are)**\n\n"
        "İngilizcede eylem bildirmeyen, durum ve kimlik bildiren cümlelerde kullanılır.\n\n"
        "• I ➔ am\n"
        "• He / She / It ➔ is\n"
        "• You / We / They ➔ are\n\n"
        "Örnek: *She is a doctor.* (O bir doktordur.)"
    ),
    "A2": (
        "📗 **A2: Past Simple (Geçmiş Zaman)**\n\n"
        "Geçmişte olup bitmiş olaylar için kullanılır.\n\n"
        "• Düzenli: Fiile -ed gelir (watch ➔ watched)\n"
        "• Düzensiz: Şekil değişir (go ➔ went, see ➔ saw)\n\n"
        "Örnek: *I went to school yesterday.* (Dün okula gittim.)"
    ),
    "B1": (
        "📙 **B1: Present Perfect Tense**\n\n"
        "Geçmişte olmuş ama zamanı net olmayan veya etkisi süren durumlar.\n\n"
        "• have/has + V3 (fiilin 3. hali)\n\n"
        "Örnek: *I have visited Paris.* (Paris'i ziyaret ettim.)"
    ),
    "B2": (
        "📕 **B2: Conditionals (Koşul Cümleleri)**\n\n"
        "• Type 2 (Hayali): *If I were you, I would study.* (Senin yerinde olsaydım çalışırdım.)\n"
        "• Type 3 (Pişmanlık): *If I had known, I would have helped.* (Bilseydim yardım ederdim.)"
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
            print("Kelimeler veritabanina yuklendi.")
        except Exception as e:
            print("Excel yukleme hatasi:", e)
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

# --- KLAVYE MENÜSÜ ---

def get_main_keyboard():
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    keyboard.add(
        types.KeyboardButton("🤖 AI Koç (Soru Sor / Cümle Düzelt)")
    )
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

# --- KOMUT VE MESAJ YÖNLENDİRME ---

@bot.message_handler(commands=['start'])
def handle_start(message):
    uid = message.from_user.id
    if not is_authorized(uid):
        bot.send_message(message.chat.id, "⛔ **Erişim Engellendi:** Bu bot sadece yetkili kullanıcılara özeldir.", parse_mode="Markdown")
        return

    get_user_profile(uid)
    ai_chat_users.discard(uid)
    welcome_text = (
        f"Selam **{message.from_user.first_name}**! 🇬🇧\n\n"
        "AI Destekli İngilizce Koçuna hoş geldin! Aşağıdaki butonları kullanarak hemen öğrenmeye başlayabilirsin."
    )
    bot.send_message(message.chat.id, welcome_text, reply_markup=get_main_keyboard(), parse_mode="Markdown")

@bot.message_handler(func=lambda msg: True)
def handle_text_menu(message):
    uid = message.from_user.id
    if not is_authorized(uid):
        return

    txt = message.text

    if txt == "🤖 AI Koç (Soru Sor / Cümle Düzelt)":
        ai_chat_users.add(uid)
        bot.send_message(
            message.chat.id,
            "🤖 **AI İngilizce Koçun Dinliyor!**\n\n"
            "Bana aklına gelen her şeyi sorabilirsin:\n"
            "• İngilizce yazdığın bir cümleyi atıp *'Doğru mu?'* diyebilirsin.\n"
            "• Kelimeler arasındaki farkları sorabilirsin.\n"
            "• Gramer kuralı açıklatabilirsin.\n\n"
            "✍️ *Lütfen sorun veya cümleni doğrudan buraya yaz:*",
            parse_mode="Markdown"
        )
        return

    menu_buttons = [
        "📚 Kelime Çöz", "🧠 Zayıf Kelimelerim", "🗂 Kategorik Kelimeler",
        "📖 Gramer Dersleri", "🎓 Seviye Belirleme Sınavı", "📊 Profilim & Skor"
    ]
    if uid in ai_chat_users and txt not in menu_buttons:
        if not ai_model:
            bot.send_message(message.chat.id, "⚠️ AI anahtarı tanımlanmamış. Render Environment ayarlarını kontrol edin.")
            return

        bot.send_chat_action(message.chat.id, 'typing')
        prompt = (
            "Sen uzman, teşvik edici ve samimi bir İngilizce öğretmenisin. "
            "Kullanıcının yazdığı İngilizce cümleyi veya soruyu analiz et. "
            "Eğer cümlede gramer veya kelime hatası varsa doğrusunu göster ve nedenini Türkçe olarak net açıkla. "
            "Eğer cümle hatasızsa tebrik et ve alternatif kullanımlar öner. "
            f"Kullanıcının mesajı: {txt}"
        )
        try:
            response = ai_model.generate_content(prompt)
            bot.reply_to(message, response.text)
        except Exception as e:
            bot.reply_to(message, f"AI yanıt verirken bir sorun oluştu: {e}")
        return

    ai_chat_users.discard(uid)

    if txt == "📚 Kelime Çöz":
        send_quiz_question(message.chat.id, is_unknown_mode=False)

    elif txt == "🧠 Zayıf Kelimelerim":
        count = get_unknown_words_count(uid)
        if count == 0:
            bot.send_message(message.chat.id, "🎉 **Harika!** Zayıf kelimeler listende henüz kelime yok.", parse_mode="Markdown")
        else:
            send_quiz_question(message.chat.id, is_unknown_mode=True)

    elif txt == "🗂 Kategorik Kelimeler":
        markup = types.InlineKeyboardMarkup(row_width=2)
        for cat_key, cat_val in CATEGORY_DATA.items():
            markup.add(
                types.InlineKeyboardButton(f"{cat_val['title']} (Kelimeler)", callback_data=f"cw:{cat_key}"),
                types.InlineKeyboardButton("🗣 Cümleler", callback_data=f"cs:{cat_key}")
            )
        bot.send_message(message.chat.id, "📌 **Çalışmak istediğin alanı seç:**", reply_markup=markup, parse_mode="Markdown")

    elif txt == "📖 Gramer Dersleri":
        markup = types.InlineKeyboardMarkup(row_width=2)
        markup.add(
            types.InlineKeyboardButton("A1: To Be", callback_data="gm:A1"),
            types.InlineKeyboardButton("A2: Past Simple", callback_data="gm:A2"),
            types.InlineKeyboardButton("B1: Present Perfect", callback_data="gm:B1"),
            types.InlineKeyboardButton("B2: Conditionals", callback_data="gm:B2")
        )
        bot.send_message(message.chat.id, "📘 **Gramer Seviyesini Seç:**", reply_markup=markup, parse_mode="Markdown")

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
            f"🎯 Seviyen: **{lvl}**\n"
            f"✅ Doğru: **{c}**\n"
            f"❌ Yanlış: **{w}**\n"
            f"📈 Başarı: **{rate}**\n"
            f"🧠 Tekrar Bekleyen Kelimeler: **{un_count} adet**"
        )
        bot.send_message(message.chat.id, msg, parse_mode="Markdown")

# --- KELİME TESTİ ---

def send_quiz_question(chat_id, is_unknown_mode=False):
    if is_unknown_mode:
        eng, tr, options = get_random_unknown_quiz(chat_id)
        mode = "U"
    else:
        eng, tr, options = get_general_quiz()
        mode = "G"

    if not eng:
        bot.send_message(chat_id, "Kelime bulunamadı, lütfen tekrar deneyin.")
        return

    markup = types.InlineKeyboardMarkup(row_width=1)
    for opt in options:
        status = "C" if opt == tr else "W"
        cb = f"Q|{mode}|{status}|{eng[:12]}|{tr[:12]}"
        markup.add(types.InlineKeyboardButton(opt, callback_data=cb))
    
    idk_cb = f"Q|{mode}|IDK|{eng[:12]}|{tr[:12]}"
    markup.add(types.InlineKeyboardButton("🤷‍♂️ Bilmiyorum", callback_data=idk_cb))

    prefix = "🧠 **[Zayıf Kelime]**\n" if is_unknown_mode else "🎯 "
    bot.send_message(chat_id, f"{prefix}**{eng}** kelimesinin Türkçe karşılığı nedir?", reply_markup=markup, parse_mode="Markdown")

# --- SEVİYE SINAVI ---

def send_placement_question(chat_id, user_id):
    state = active_placement_tests.get(user_id)
    if not state:
        return
    idx = state["q_idx"]
    if idx >= len(PLACEMENT_QUESTIONS):
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
            f"🏁 **Sınav Bitti!**\n\nDoğru: **{score}/10**\nBelirlenen Seviyen: **{final_lvl}**",
            parse_mode="Markdown"
        )
        return

    q_data = PLACEMENT_QUESTIONS[idx]
    markup = types.InlineKeyboardMarkup(row_width=2)
    for opt in q_data["options"]:
        cb = f"PT|{opt}|{q_data['a']}"
        markup.add(types.InlineKeyboardButton(opt, callback_data=cb))

    bot.send_message(
        chat_id,
        f"📝 **Soru {idx + 1}/10 ({q_data['lvl']}):**\n\n{q_data['q']}",
        reply_markup=markup,
        parse_mode="Markdown"
    )

# --- BUTON CEVAPLARI ---

@bot.callback_query_handler(func=lambda call: True)
def handle_callbacks(call):
    if not is_authorized(call.from_user.id):
        bot.answer_callback_query(call.id, "⛔ Yetkisiz erişim!", show_alert=True)
        return

    bot.answer_callback_query(call.id)
    data = call.data
    uid = call.from_user.id
    cid = call.message.chat.id
    mid = call.message.message_id

    if data.startswith("Q|"):
        _, mode, status, eng, tr = data.split("|", 4)
        if status == "C":
            update_score(uid, True)
            extra = ""
            if mode == "U":
                remove_unknown_word(uid, eng)
                extra = "\n✨ *Bu kelimeyi öğrendin, listeden kaldırıldı!*"
            bot.edit_message_text(f"✅ **Doğru!**\n\n🇬🇧 **{eng}** = 🇹🇷 **{tr}**{extra}", chat_id=cid, message_id=mid, parse_mode="Markdown")
        elif status == "W":
            update_score(uid, False)
            add_unknown_word(uid, eng, tr)
            bot.edit_message_text(f"❌ **Yanlış!**\n\n🇬🇧 **{eng}** = 🇹🇷 **{tr}**\n\n📌 *Zayıf kelimeler listene eklendi.*", chat_id=cid, message_id=mid, parse_mode="Markdown")
        elif status == "IDK":
            add_unknown_word(uid, eng, tr)
            bot.edit_message_text(f"💡 **Doğrusu:**\n\n🇬🇧 **{eng}** = 🇹🇷 **{tr}**\n\n📌 *Zayıf kelimeler listene eklendi.*", chat_id=cid, message_id=mid, parse_mode="Markdown")

    elif data.startswith("PT|"):
        _, chosen, ans = data.split("|", 2)
        state = active_placement_tests.get(uid)
        if state:
            if chosen == ans:
                state["score"] += 1
            state["q_idx"] += 1
            bot.delete_message(cid, mid)
            send_placement_question(cid, uid)

    elif data.startswith("cw:"):
        k = data.split(":")[1]
        cat = CATEGORY_DATA.get(k)
        if cat:
            txt = f"{cat['title']} - **Kelimeler**:\n\n"
            for eng, tr in cat['words']:
                txt += f"• **{eng}**: {tr}\n"
            bot.send_message(cid, txt, parse_mode="Markdown")

    elif data.startswith("cs:"):
        k = data.split(":")[1]
        cat = CATEGORY_DATA.get(k)
        if cat:
            txt = f"{cat['title']} - **Kalıp Cümleler**:\n\n"
            for eng, tr in cat['sentences']:
                txt += f"💬 **{eng}**\n↳ *{tr}*\n\n"
            bot.send_message(cid, txt, parse_mode="Markdown")

    elif data.startswith("gm:"):
        lvl = data.split(":")[1]
        lesson = GRAMMAR_LESSONS.get(lvl)
        if lesson:
            bot.send_message(cid, lesson, parse_mode="Markdown")

# --- BAŞLATMA ---

if __name__ == "__main__":
    init_database()
    t = threading.Thread(target=run_flask)
    t.daemon = True
    t.start()
    print("Bot basariyla baslatildi...")
    bot.infinity_polling()
