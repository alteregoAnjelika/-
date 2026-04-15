import os
import logging
from datetime import datetime

import openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)

# логирование
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# токен
BOT_TOKEN = os.getenv("BOT_TOKEN", "8321107522:AAE4K0TFueHRnCDgBDSWfEILv9dMltrQHrc")

# ---------- загрузка данных из excel ----------

def load_arcana_data():
    """
    Читаем файл arcana.xlsx, вкладку 'сериалы'.
    A = номер аркана
    B = название аркана
    C = персонаж
    D = описание
    """
    data = {}

    # ищем файл рядом с bot.py
    base_dir = os.path.dirname(os.path.abspath(__file__))
    filepath = os.path.join(base_dir, "arcana.xlsx")

    wb = openpyxl.load_workbook(filepath, read_only=True)
    ws = wb["сериалы"]

    for row in ws.iter_rows(min_row=2, values_only=True):
        # пропускаем пустые строки
        if not row[0]:
            continue

        number = int(row[0])       # A — номер аркана
        name = str(row[1]).strip() # B — название
        character = str(row[2]).strip()  # C — персонаж
        description = str(row[3]).strip()  # D — описание

        data[number] = {
            "name": name,
            "character": character,
            "description": description,
        }

    wb.close()
    logger.info(f"Загр��жено {len(data)} арканов из Excel")
    return data


# глобальная переменная с данными
ARCANA = load_arcana_data()


# ---------- расчёт аркана ----------

def calculate_arcanum(day: int) -> int:
    """
    День рождения -> главный аркан (1-22).
    Если день <= 22, аркан = день.
    Если день > 22, складываем цифры.
    22 остаётся как 22.
    """
    if day <= 22:
        return day

    # складываем цифры дня
    result = sum(int(d) for d in str(day))

    # на всякий случай, если вдруг > 22 (не должно быть для дней 1-31)
    while result > 22:
        result = sum(int(d) for d in str(result))

    return result


def parse_date(text: str):
    """
    Пытаемся распарсить дату из текста пользователя.
    Поддерживаем форматы:
    ДД.ММ.ГГГГ, ДД/ММ/ГГГГ, ДД-ММ-ГГГГ,
    ДД.ММ, ДД ММ ГГГГ и просто число (день).
    """
    text = text.strip()

    # если просто число — считаем это днём
    if text.isdigit():
        day = int(text)
        if 1 <= day <= 31:
            return day
        return None

    # пробуем разные форматы
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y", "%d.%m", "%d %m %Y"):
        try:
            dt = datetime.strptime(text, fmt)
            return dt.day
        except ValueError:
            continue

    return None


# ---------- формируем красивый ответ ----------

def format_response(day: int) -> str:
    arcanum = calculate_arcanum(day)

    if arcanum not in ARCANA:
        return (
            f"Хм, для дня {day} получился аркан {arcanum}, "
            f"но у меня нет данных по нему. Проверь файл Excel 🤔"
        )

    info = ARCANA[arcanum]

    text = (
        f"🎂 День рождения: {day} число\n"
        f"🔮 Твой главный аркан: {arcanum} — {info['name']}\n\n"
        f"🎬 Твой персонаж: {info['character']}\n\n"
        f"{info['description']}\n\n"
        f"Узнал себя? 😏 Перешли другу — пусть тоже узнает, кто он!"
    )

    return text


# ---------- хэндлеры бота ----------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome = (
        "Привет! 👋\n\n"
        "Я расскажу, кто ты из сериалов «Друзья», "
        "«Отчаянные домохозяйки» и «Зачарованные» "
        "по твоей дате рождения.\n\n"
        "Просто отправь мне дату в любом формате:\n"
        "• 15.03.1990\n"
        "• 15/03/1990\n"
        "• 15-03-1990\n"
        "• или просто число дня, например: 15\n\n"
        "Поехали! 🚀"
    )
    await update.message.reply_text(welcome)


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text

    day = parse_date(text)

    if day is None:
        await update.message.reply_text(
            "Не могу разобрать дату 😅\n\n"
            "Отправь в формате ДД.ММ.ГГГГ (например, 15.03.1990) "
            "или просто число дня рождения (например, 15)."
        )
        return

    response = format_response(day)
    await update.message.reply_text(response)


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "🔮 Как пользоваться ботом:\n\n"
        "1. Отправь свою дату рождения\n"
        "2. Получи свой аркан и персонажа из сериала\n"
        "3. Перешли друзьям!\n\n"
        "Форматы даты: 15.03.1990, 15/03/1990, "
        "или просто число дня (15).\n\n"
        "Команды:\n"
        "/start — начать\n"
        "/help — эта справка\n"
        "/table — таблица всех арканов"
    )
    await update.message.reply_text(text)


async def table_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Показывает список всех арканов и персонажей."""
    lines = ["🔮 Все 22 аркана:\n"]

    for num in sorted(ARCANA.keys()):
        info = ARCANA[num]
        lines.append(f"{num}. {info['name']} — {info['character']}")

    await update.message.reply_text("\n".join(lines))


async def error_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    logger.error(f"Ошибка: {context.error}")


# ---------- запуск ----------

def main():
    app = Application.builder().token(BOT_TOKEN).build()

    # команды
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("table", table_command))

    # любое текстовое сообщение — пробуем как дату
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    # ошибки
    app.add_error_handler(error_handler)

    logger.info("Бот запущен!")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
