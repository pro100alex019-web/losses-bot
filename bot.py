"""
Telegram-бот расчёта потерь электроэнергии.
Приказ Минэнерго РФ № 326 от 30.12.2008.

Поток:
  /start → инструкция
  → загрузка фото/скана схемы + комментарий
  → GPT-4o Vision анализирует схему
  → бот отправляет Excel с предзаполненными данными
  → пользователь проверяет и загружает Excel обратно
  → бот рассчитывает константы C, A, B
  → предлагает рассчитать потери за конкретный месяц
  → пользователь вводит месяц + W_а → результат
"""

import io
import logging
import os

from dotenv import load_dotenv
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, BotCommand
from telegram.ext import (
    Application, ApplicationBuilder, CallbackQueryHandler,
    CommandHandler, ContextTypes, ConversationHandler,
    MessageHandler, filters,
)

from excel_generator import generate_excel
from excel_reader import read_constants_from_excel
from losses_calc import (
    NetworkConstants, calculate_monthly_losses,
    parse_month_and_consumption, MONTH_HOURS,
)
from losses_handler import format_constants
from scheme_parser import parse_scheme, format_parse_summary

load_dotenv()
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
ADMIN_ID = int(os.getenv("ADMIN_ID", "0"))

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# ── Состояния разговора ──
(
    WAITING_SCHEME_COMMENT,   # ждём комментарий перед схемой
    WAITING_SCHEME_IMAGE,     # ждём изображение схемы
    WAITING_EXCEL_BACK,       # ждём проверенный Excel обратно
    ASKING_MONTHLY,           # спрашиваем хочет ли расчёт за месяц
    WAITING_MONTH_INPUT,      # ждём ввод месяца + W_а
    NEXT_MONTH_QUESTION,      # спрашиваем ещё один месяц?
) = range(6)


# ── Клавиатуры ──

def kb_start_calc():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("📐 Начать расчёт потерь", callback_data="start_calc"),
    ]])

def kb_confirm_scheme():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Верно, сгенерировать Excel", callback_data="scheme_ok"),
        InlineKeyboardButton("🔄 Загрузить другую схему",    callback_data="scheme_retry"),
    ]])

def kb_excel_sent():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("📎 Загрузить проверенный Excel", callback_data="upload_excel"),
    ]])

def kb_monthly():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Да, рассчитать за месяц", callback_data="month_yes"),
        InlineKeyboardButton("🏁 Нет, завершить",          callback_data="month_no"),
    ]])

def kb_next_month():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("📅 Ещё один месяц", callback_data="month_yes"),
        InlineKeyboardButton("🏁 Завершить",       callback_data="month_no"),
    ]])


# ── Хэндлеры ──

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "⚡ *Бот расчёта потерь электроэнергии*\n\n"
        "Рассчитываю технические потери в электросетях по "
        "*Приказу Минэнерго РФ № 326 от 30.12.2008*.\n\n"
        "📋 *Что умею:*\n"
        "• Распознаю однолинейную схему (фото или скан)\n"
        "• Определяю марки кабелей, длины линий, марки трансформаторов\n"
        "• Генерирую Excel-шаблон с расчётом констант формул\n"
        "• Рассчитываю потери за любой месяц по показаниям счётчика\n\n"
        "📌 *Формулы (Приказ 326):*\n"
        "Линии: `ΔW = C × W²а / T`\n"
        "Трансформаторы: `ΔW = A × T + B × W²а / T`\n\n"
        "Нажми кнопку чтобы начать:",
        parse_mode="Markdown",
        reply_markup=kb_start_calc(),
    )


async def start_calc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    await q.message.reply_text(
        "📕 *Как подготовить схему — инструкция*\n\n"
        "🔴 *Шаг 1. Обведи нужные элементы КРАСНОЙ линией:*\n"
        "• Нарисуй поверх схемы красным маркером/ручкой\n"
        "• Линия должна проходить по всем участкам, потери в которых нужно считать\n"
        "• Захвати символы трансформаторов, если они входят в расчёт\n"
        "• Чем толще и ярче линия — тем точнее распознавание\n\n"
        "❗️ *Особенности обводки:*\n"
        "• Красная линия должна визуально пересекать участок линии или символ ТП\n"
        "• Не используй красные чернила если сама схема нарисована красным — "
        "возьми другой яркий цвет (зелёный, синий)\n"
        "• Если на схеме несколько ветвей — обведи только нужные\n"
        "• Стрелка или замкнутый контур вокруг элемента тоже работает\n\n"
        "📝 *Шаг 2. Проверь читаемость подписей рядом с обведёнными элементами:*\n"
        "• Марка кабеля/провода (ААБ2п 3×120, СИП 4×95, ЗАС-50…)\n"
        "• Длина участка (в метрах или км)\n"
        "• Марка трансформатора (ТМ-400/10, ТМГ-630/10…)\n\n"
        "⚠️ Если что-то не подписано — допиши в комментарии при загрузке.\n\n"
        "✏️ *Напиши комментарий* — уточни что обведено или добавь пропущенные данные "
        "(или напиши «всё указано на схеме»):",
        parse_mode="Markdown",
    )
    return WAITING_SCHEME_COMMENT


async def receive_comment(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["comment"] = update.message.text
    await update.message.reply_text(
        "📎 Теперь загрузи *фото или скан* однолинейной схемы.\n\n"
        "_Поддерживаются форматы: JPG, PNG, PDF (первая страница)_",
        parse_mode="Markdown",
    )
    return WAITING_SCHEME_IMAGE


async def receive_scheme_image(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Получаем изображение (фото или документ)
    if update.message.photo:
        file_obj = await update.message.photo[-1].get_file()
    elif update.message.document:
        file_obj = await update.message.document.get_file()
    else:
        await update.message.reply_text("❌ Нужно изображение схемы (фото или файл).")
        return WAITING_SCHEME_IMAGE

    await update.message.reply_text("🔍 Анализирую схему...")

    buf = io.BytesIO()
    await file_obj.download_to_memory(buf)
    image_bytes = buf.getvalue()
    context.user_data["scheme_bytes"] = image_bytes

    comment = context.user_data.get("comment", "все элементы")

    # GPT Vision анализ
    scheme_data = parse_scheme(image_bytes, comment, OPENAI_API_KEY)

    if scheme_data.get("error") or (
        not scheme_data.get("lines") and not scheme_data.get("transformers")
    ):
        await update.message.reply_text(
            "⚠️ Не удалось распознать схему.\n\n"
            "Возможные причины:\n"
            "• Изображение нечёткое или слишком мелкое\n"
            "• На схеме нет подписей к элементам\n\n"
            "Попробуй загрузить другое изображение:",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("🔄 Загрузить другую схему", callback_data="scheme_retry"),
            ]]),
        )
        return WAITING_SCHEME_IMAGE

    context.user_data["scheme_data"] = scheme_data
    summary = format_parse_summary(scheme_data)

    await update.message.reply_text(
        summary + "\n\n_Всё верно? Сгенерировать Excel-шаблон?_",
        parse_mode="Markdown",
        reply_markup=kb_confirm_scheme(),
    )
    return WAITING_EXCEL_BACK


async def scheme_ok(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    scheme_data = context.user_data.get("scheme_data", {})

    await q.message.reply_text("📊 Генерирую Excel-шаблон...")

    try:
        excel_bytes = generate_excel(scheme_data)
    except Exception as e:
        logger.error(f"Excel генерация: {e}")
        await q.message.reply_text("❌ Ошибка генерации Excel. Попробуй ещё раз.")
        return WAITING_EXCEL_BACK

    await q.message.reply_document(
        document=io.BytesIO(excel_bytes),
        filename="расчёт_потерь_326.xlsx",
        caption=(
            "📎 *Excel-шаблон с данными из схемы*\n\n"
            "Что нужно сделать:\n"
            "1️⃣ Открой файл и перейди на лист *«1. Параметры сети»*\n"
            "2️⃣ Проверь все жёлтые ячейки\n"
            "3️⃣ Исправь ошибки распознавания (марки, длины, мощности)\n"
            "4️⃣ Загрузи исправленный файл обратно сюда\n\n"
            "_Листы 2-3 заполнятся автоматически по твоим данным_"
        ),
        parse_mode="Markdown",
    )
    await q.message.reply_text(
        "⬆️ Загрузи проверенный Excel-файл:",
        reply_markup=kb_excel_sent(),
    )
    return WAITING_EXCEL_BACK


async def scheme_retry(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    await q.message.reply_text(
        "📎 Загрузи другое изображение схемы:",
    )
    return WAITING_SCHEME_IMAGE


async def receive_excel_back(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc or not doc.file_name.endswith(".xlsx"):
        await update.message.reply_text(
            "❌ Нужен файл *Excel (.xlsx)*.",
            parse_mode="Markdown",
        )
        return WAITING_EXCEL_BACK

    await update.message.reply_text("⏳ Читаю данные из Excel...")

    file = await doc.get_file()
    buf = io.BytesIO()
    await file.download_to_memory(buf)

    constants = read_constants_from_excel(buf.getvalue())

    if not constants or not constants.has_data():
        await update.message.reply_text(
            "⚠️ Не удалось извлечь данные из файла.\n"
            "Убедись что загружаешь файл *расчёт_потерь_326.xlsx* "
            "с заполненными данными на листе «1. Параметры сети».",
            parse_mode="Markdown",
        )
        return WAITING_EXCEL_BACK

    context.user_data["constants"] = constants

    # Показываем константы
    await update.message.reply_text(
        format_constants(constants),
        parse_mode="Markdown",
    )

    # Спрашиваем о расчёте за месяц
    await update.message.reply_text(
        "💬 *Хотите рассчитать потери за конкретный месяц?*\n\n"
        "Для этого нужно:\n"
        "• месяц\n"
        "• объём переданной электроэнергии W_а (кВт·ч) "
        "по входному счётчику за этот месяц\n\n"
        "⚠️ *W_а* — показания счётчика на вводе 10 кВ (весь отпуск объекта), "
        "не сумма счётчиков 0,4 кВ.",
        parse_mode="Markdown",
        reply_markup=kb_monthly(),
    )
    return ASKING_MONTHLY


async def month_yes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    await q.message.reply_text(
        "📅 Введи *месяц* и *W_а* в одном сообщении.\n\n"
        "Примеры:\n"
        "`Январь 120000`\n"
        "`март 85 000 кВт·ч`\n"
        "`октябрь 115000`\n\n"
        "_Число часов T для выбранного месяца подставится автоматически._",
        parse_mode="Markdown",
    )
    return WAITING_MONTH_INPUT


async def month_no(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    await q.message.reply_text(
        "✅ Расчёт завершён.\n\n"
        "Константы сохранены в Excel (листы 2–3).\n"
        "Для ежемесячного расчёта используй формулы:\n\n"
        "• Линии: `ΔW_лин = C × W²а / T`\n"
        "• Трансформаторы: `ΔW_тр = A × T + B × W²а / T`\n\n"
        "_Для нового расчёта: /start_",
        parse_mode="Markdown",
    )
    context.user_data.clear()
    return ConversationHandler.END


async def receive_month_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    parsed = parse_month_and_consumption(update.message.text)
    if not parsed:
        await update.message.reply_text(
            "❌ Не распознал. Введи в формате: `Январь 120000`",
            parse_mode="Markdown",
        )
        return WAITING_MONTH_INPUT

    month, T, W_a = parsed
    constants: NetworkConstants = context.user_data.get("constants")

    result = calculate_monthly_losses(constants, month, T, W_a)

    await update.message.reply_text(
        result.format_report(),
        parse_mode="Markdown",
    )
    await update.message.reply_text(
        "📅 Рассчитать потери за ещё один месяц?",
        reply_markup=kb_next_month(),
    )
    return NEXT_MONTH_QUESTION


async def month_again(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    await q.message.reply_text(
        "Введи следующий месяц и W_а:\n`Февраль 98000`",
        parse_mode="Markdown",
    )
    return WAITING_MONTH_INPUT


async def cmd_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("Отменено. /start — начать заново.")
    return ConversationHandler.END


# ── Запуск ──

async def post_init(app: Application):
    await app.bot.set_my_commands([
        BotCommand("start",  "⚡ Начать расчёт потерь"),
        BotCommand("cancel", "❌ Отменить"),
    ])


def main():
    app = (
        ApplicationBuilder()
        .token(TELEGRAM_TOKEN)
        .post_init(post_init)
        .build()
    )

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", cmd_start),
            CallbackQueryHandler(start_calc, pattern="^start_calc$"),
        ],
        states={
            WAITING_SCHEME_COMMENT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_comment),
            ],
            WAITING_SCHEME_IMAGE: [
                MessageHandler(filters.PHOTO | filters.Document.ALL, receive_scheme_image),
                CallbackQueryHandler(scheme_retry, pattern="^scheme_retry$"),
            ],
            WAITING_EXCEL_BACK: [
                MessageHandler(
                    filters.Document.MimeType(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ) | filters.Document.FileExtension("xlsx"),
                    receive_excel_back,
                ),
                CallbackQueryHandler(scheme_ok,    pattern="^scheme_ok$"),
                CallbackQueryHandler(scheme_retry, pattern="^scheme_retry$"),
                CallbackQueryHandler(lambda u, c: WAITING_EXCEL_BACK,
                                     pattern="^upload_excel$"),
            ],
            ASKING_MONTHLY: [
                CallbackQueryHandler(month_yes, pattern="^month_yes$"),
                CallbackQueryHandler(month_no,  pattern="^month_no$"),
            ],
            WAITING_MONTH_INPUT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_month_input),
            ],
            NEXT_MONTH_QUESTION: [
                CallbackQueryHandler(month_again, pattern="^month_yes$"),
                CallbackQueryHandler(month_no,    pattern="^month_no$"),
            ],
        },
        fallbacks=[CommandHandler("cancel", cmd_cancel)],
        per_message=False,
    )

    app.add_handler(conv)
    app.add_handler(CommandHandler("start",  cmd_start))
    app.add_handler(CommandHandler("cancel", cmd_cancel))

    logger.info("Бот потерь запущен...")
    app.run_polling()


if __name__ == "__main__":
    main()
