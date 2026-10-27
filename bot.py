import asyncio
import json
import logging
import os
import sqlite3
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    BufferedInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    WebAppInfo,
)
import openpyxl

# Сюда вставьте ваш токен бота от BotFather
TOKEN = "9026519587:AAG0-XkQ7fvGu4H4agfoSDV4YuZNv9KffLY"  
# Ссылка на ваш рабочий Mini App (GitHub Pages)
WEB_APP_URL = "https://htmlpreview.github.io/?https://github.com/dilshodbekshermatov799-debug/maktab-miniapp/blob/main/index.html"

bot = Bot(token=TOKEN)
dp = Dispatcher()
DB_NAME = "school_bot.db"

class MediaState(StatesGroup):
    waiting_for_hw_media = State()
    waiting_for_club_media = State()

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    # Таблица для домашних заданий (сохраняет класс вместо ID ученика)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS homework_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_name TEXT,
            teacher_telegram_id INTEGER,
            status TEXT,
            file_unique_id TEXT,
            date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    # Таблица для кружков
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS club_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_telegram_id INTEGER,
            club_name TEXT,
            attendance_count INTEGER,
            file_unique_id TEXT,
            date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()

@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📱 Hisobot panelini ochish", web_app=WebAppInfo(url=WEB_APP_URL))],
        [InlineKeyboardButton(text="📊 Excel Hisobotni olish", callback_data="get_report")]
    ])
    await message.answer(
        "Assalomu alaykum! Maktab nazorat tizimiga xush kelibsiz.\nHisobot topshirish uchun quyidagi tugmani bosing:",
        reply_markup=kb
    )

@dp.message(F.web_app_data)
async def handle_web_app_data(message: types.Message, state: FSMContext):
  try:
    data = json.loads(message.web_app_data.data)
  except Exception:
    await message.answer("❌ Xatolik: Ma'lumotni o'qib bo'lmadi.")
    return

  if data.get("type") == "homework":
    class_name = data.get("class_name")
    status = data.get("status")

    await state.update_data(class_name=class_name, status=status)

    if status in ["Bajardi", "Chala"]:
      await message.answer(
          f"📸 {class_name} sinfi uchun rasm yoki videoni yuklang (Tekshirish"
          " uchun majburiy):"
      )
      await state.set_state(MediaState.waiting_for_hw_media)
    else:
      conn = sqlite3.connect(DB_NAME)
      cursor = conn.cursor()
      cursor.execute(
          "INSERT INTO homework_logs (class_name, teacher_telegram_id, status)"
          " VALUES (?, ?, ?)",
          (class_name, message.from_user.id, "Bajarmadi"),
      )
      conn.commit()
      conn.close()
      await message.answer(
          f"❌ {class_name} sinfining vazifasi 'Bajarmadi' deb saqlandi!"
      )

  elif data.get("type") == "club":
    club_name = data.get("club_name")
    count = data.get("count")

    await state.update_data(club_name=club_name, count=count)
    await message.answer(
        f"📸 {club_name} to'garagidan **rasm yoki video** yuboring:"
    )
    await state.set_state(MediaState.waiting_for_club_media)

@dp.message(MediaState.waiting_for_hw_media, F.photo | F.video)
async def save_hw_media(message: types.Message, state: FSMContext):
    file_uid = message.photo[-1].file_unique_id if message.photo else message.video.file_unique_id
    user_data = await state.get_data()

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM homework_logs WHERE file_unique_id = ?", (file_uid,))
    if cursor.fetchone():
        conn.close()
        await message.answer("⚠️ Bu rasm/video allaqachon ishlatilgan! Boshqasini yuboring.")
        return

    cursor.execute(
        "INSERT INTO homework_logs (class_name, teacher_telegram_id, status, file_unique_id) VALUES (?, ?, ?, ?)",
        (user_data["class_name"], message.from_user.id, user_data["status"], file_uid)
    )
    conn.commit()
    conn.close()
    await message.answer("✅ Sinf vazifasi va rasm muvaffaqiyatli saqlandi!")
    await state.clear()

@dp.message(MediaState.waiting_for_club_media, F.photo | F.video)
async def save_club_media(message: types.Message, state: FSMContext):
    file_uid = message.photo[-1].file_unique_id if message.photo else message.video.file_unique_id
    user_data = await state.get_data()

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM club_logs WHERE file_unique_id = ?", (file_uid,))
    if cursor.fetchone():
        conn.close()
        await message.answer("⚠️ Bu rasm/video allaqachon ishlatilgan! Boshqasini yuboring.")
        return

    cursor.execute(
        "INSERT INTO club_logs (teacher_telegram_id, club_name, attendance_count, file_unique_id) VALUES (?, ?, ?, ?)",
        (message.from_user.id, user_data["club_name"], user_data["count"], file_uid)
    )
    conn.commit()
    conn.close()
    await message.answer("✅ To'garak hisoboti va rasm saqlandi!")
    await state.clear()

@dp.callback_query(F.data == "get_report")
async def generate_excel(callback: types.CallbackQuery):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    wb = openpyxl.Workbook()

    # Лист домашний заданий (теперь выводит Класс вместо ID ученика)
    ws1 = wb.active
    ws1.title = "Uyga vazifa"
    ws1.append(["Sinf", "O'qituvchi ID", "Holati", "Sana"])
    cursor.execute("SELECT class_name, teacher_telegram_id, status, date FROM homework_logs")
    for row in cursor.fetchall():
        ws1.append(row)

    # Лист кружков
    ws2 = wb.create_sheet(title="To'garak")
    ws2.append(["To'garak", "Qatnashchilar", "Sana"])
    cursor.execute("SELECT club_name, attendance_count, date FROM club_logs")
    for row in cursor.fetchall():
        ws2.append(row)

    file_path = "hisobot.xlsx"
    wb.save(file_path)
    conn.close()

    if os.path.exists(file_path):
        doc = BufferedInputFile(open(file_path, "rb").read(), filename="Maktab_Hisobot.xlsx")
        await callback.message.answer_document(doc, caption="📊 Excel Hisobot tayyor!")
        os.remove(file_path)
    else:
        await callback.message.answer("❌ Xatolik: Hisobot fayli yaratilmadi!")

async def main():
    init_db()
    logging.basicConfig(level=logging.INFO)
    print("Bot ishga tushdi (Faqat sinflar bo'yicha)...")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())