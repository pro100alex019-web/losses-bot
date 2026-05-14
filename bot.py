"""
Telegram-бот расчёта потерь электроэнергии.
Приказ Минэнерго РФ № 326 от 30.12.2008.

Состояния хранятся в context.user_data["state"].
Никакого ConversationHandler — простая и надёжная state-машина.
"""

import io
import logging
import os

from dotenv import load_dotenv
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, BotCommand
from telegram.ext import (
    Application, ApplicationBuilder, CallbackQueryHandler,
    CommandHandler, ContextTypes, MessageHandler, filters,
)

from excel_generator import generate_excel
from excel_reader import read_constants_from_excel
from losses_calc import (
    NetworkConstants, calculate_monthly_losses,
    parse_month_and_consumption,
)
from losses_handler import format_constants
from scheme_parser import parse_scheme, format_parse_summary

load_dotenv()
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
ADMIN_ID       = int(os.getenv("ADMIN_ID", "0"))

logging.basicConfig(
    format="%(asctime)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# ── Состояния ──────────────────────────────────────────────────
S_IDLE           = "idle"
S_WAIT_COMMENT   = "wait_comment"
S_WAIT_IMAGE     = "wait_image"
S_WAIT_EXCEL     = "wait_excel"
S_WAIT_MONTH     = "wait_month"
S_WAIT_MORE      = "wait_more"


def state(ctx):
    return ctx.user_data.get("state", S_IDLE)

def set_state(ctx, s):
    ctx.user_data["state"] = s
    logger.info(f"State → {s}")


# ── Клавиатуры ─────────────────────────────────────────────────

def kb_start():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("📐 Начать расчёт потерь", callback_data="cb_start"),
    ]])

def kb_confirm():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Верно → сгенерировать Excel", callback_data="cb_ok"),
        InlineKeyboardButton("🔄 Другая схема",                callback_data="cb_retry"),
    ]])

def kb_monthly():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Да, рассчитать за месяц", callback_data="cb_month_yes"),
        InlineKeyboardButton("🏁 Нет, завершить",          callback_data="cb_month_no"),
    ]])

def kb_next_month():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("📅 Ещё месяц", callback_data="cb_month_yes"),
        InlineKeyboardButton("🏁 Завершить", callback_data="cb_month_no"),
    ]])


# ── /start ─────────────────────────────────────────────────────

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    set_state(ctx, S_IDLE)
    await update.message.reply_text(
        "⚡ *Бот расчёта потерь электроэнергии*\n\n"
        "Рассчитываю технические потери в электросетях по "
        "*Приказу Минэнерго РФ № 326 от 30.12.2008*.\n\n"
        "📋 *Что умею:*\n"
        "• Распознаю однолинейную схему (фото / скан)\n"
        "• Нахожу обведённые красным элементы\n"
        "• Генерирую Excel с константами формул C, A, B\n"
        "• Считаю потери за любой месяц по показаниям счётчика\n\n"
        "Нажми кнопку чтобы начать:",
        parse_mode="Markdown",
        reply_markup=kb_start(),
    )


# ── Кнопка «Начать» ────────────────────────────────────────────

async def cb_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    set_state(ctx, S_WAIT_COMMENT)
    await update.callback_query.message.reply_text(
        "📕 *Как подготовить схему*\n\n"
        "🔴 *Шаг 1 — Обведи нужные элементы красной линией:*\n"
        "• Нарисуй поверх схемы красным маркером/ручкой\n"
        "• Линия должна проходить по участкам, где считаем потери\n"
        "• Захвати символы трансформаторов если они входят в расчёт\n"
        "• Чем толще и ярче — тем точнее распознавание\n\n"
        "❗️ *Важно:*\n"
        "• Схема сама не должна быть красной — иначе возьми другой цвет\n"
        "• Подписи рядом с обведёнными элементами должны быть читаемы:\n"
        "  марка кабеля (ААБ2п 3×120, СИП 4×95…), длина, марка ТП\n\n"
        "✏️ *Шаг 2 — Напиши комментарий:*\n"
        "Уточни что обведено или добавь данные которых нет на схеме\n"
        "_(или напиши: «всё указано на схеме»)_",
        parse_mode="Markdown",
    )


# ── Входящие сообщения ─────────────────────────────────────────

async def on_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    s = state(ctx)
    msg = update.message

    # Комментарий перед загрузкой схемы
    if s == S_WAIT_COMMENT:
        if msg.text:
            ctx.user_data["comment"] = msg.text
            set_state(ctx, S_WAIT_IMAGE)
            await msg.reply_text(
                "📎 Отлично! Теперь загрузи схему как *файл* (не как фото!)\n\n"
                "⚠️ *Telegram сжимает обычные фото — качество падает!*\n"
                "Отправь через: Скрепка 📎 → *Файл* → выбери изображение\n"
                "_(JPG, PNG, PDF — любое разрешение)_",
                parse_mode="Markdown",
            )
        else:
            await msg.reply_text("Напиши текстовый комментарий.")
        return

    # Изображение схемы
    if s == S_WAIT_IMAGE:
        if msg.photo or msg.document:
            await _process_image(msg, ctx)
        else:
            await msg.reply_text(
                "📎 Нужно *фото или документ* со схемой (не текст).",
                parse_mode="Markdown",
            )
        return

    # Проверенный Excel
    if s == S_WAIT_EXCEL:
        doc = msg.document
        if doc and (doc.file_name or "").endswith(".xlsx"):
            await _process_excel(msg, ctx)
        else:
            await msg.reply_text(
                "📎 Нужен файл *Excel (.xlsx)*.",
                parse_mode="Markdown",
            )
        return

    # Ввод месяца + W_а
    if s == S_WAIT_MONTH:
        if msg.text:
            await _process_month(msg, ctx)
        return

    # Любое сообщение в idle — предлагаем начать
    if s == S_IDLE:
        await msg.reply_text(
            "Нажми /start чтобы начать расчёт.",
        )


async def _process_image(msg, ctx):
    """Скачать изображение → GPT Vision → показать результат."""
    set_state(ctx, S_IDLE)   # пока ждём GPT — блокируем повторные нажатия
    await msg.reply_text("🔍 Анализирую схему, это займёт ~15–30 секунд...")

    try:
        if msg.photo:
            file_obj = await msg.photo[-1].get_file()
        else:
            file_obj = await msg.document.get_file()
        buf = io.BytesIO()
        await file_obj.download_to_memory(buf)
        image_bytes = buf.getvalue()
    except Exception as e:
        logger.error(f"Скачивание: {e}")
        await msg.reply_text("❌ Не удалось скачать файл. Попробуй ещё раз /start")
        return

    comment = ctx.user_data.get("comment", "все элементы")
    try:
        scheme_data = parse_scheme(image_bytes, comment, OPENAI_API_KEY)
    except Exception as e:
        logger.error(f"parse_scheme: {e}")
        set_state(ctx, S_WAIT_IMAGE)
        await msg.reply_text(
            "❌ Ошибка при анализе схемы. Попробуй загрузить снова."
        )
        return

    if not scheme_data.get("lines") and not scheme_data.get("transformers"):
        set_state(ctx, S_WAIT_IMAGE)
        await msg.reply_text(
            "⚠️ Не удалось распознать элементы схемы.\n\n"
            "Проверь:\n"
            "• Изображение чёткое и не смазанное?\n"
            "• Красная обводка хорошо видна?\n"
            "• Подписи к элементам читаемы?\n\n"
            "Загрузи другое изображение или нажми /start",
        )
        return

    ctx.user_data["scheme_data"] = scheme_data
    try:
        summary = format_parse_summary(scheme_data)
    except Exception as e:
        logger.error(f"format_parse_summary: {e}")
        summary = f"✅ Распознано: {len(scheme_data.get('lines',[]))} линий, {len(scheme_data.get('transformers',[]))} трансформаторов"

    set_state(ctx, S_WAIT_EXCEL)
    await msg.reply_text(
        summary + "\n\n_Всё верно? Сгенерировать Excel?_",
        parse_mode="Markdown",
        reply_markup=kb_confirm(),
    )


async def _process_excel(msg, ctx):
    """Прочитать Excel → рассчитать константы → спросить о месяце."""
    await msg.reply_text("⏳ Читаю данные из Excel...")

    file = await msg.document.get_file()
    buf  = io.BytesIO()
    await file.download_to_memory(buf)

    constants = read_constants_from_excel(buf.getvalue())
    if not constants or not constants.has_data():
        await msg.reply_text(
            "⚠️ Не удалось прочитать данные из файла.\n"
            "Убедись что загружаешь именно файл *расчёт_потерь_326.xlsx* "
            "с заполненными данными на листе «1. Параметры сети».\n\n"
            "Попробуй ещё раз или нажми /start",
            parse_mode="Markdown",
        )
        return

    ctx.user_data["constants"] = constants
    set_state(ctx, S_WAIT_MORE)

    # Отправляем константы без Markdown — имена элементов могут содержать спецсимволы
    await msg.reply_text(format_constants(constants))
    await msg.reply_text(
        "💬 Рассчитать потери за конкретный месяц?\n\n"
        "Нужно: месяц + W_а (кВт·ч) по входному счётчику на вводе 10 кВ.\n"
        "⚠️ W_а — показания входного счётчика 10 кВ, не сумма счётчиков 0.4 кВ.",
        reply_markup=kb_monthly(),
    )


async def _process_month(msg, ctx):
    """Рассчитать потери за указанный месяц."""
    parsed = parse_month_and_consumption(msg.text)
    if not parsed:
        await msg.reply_text(
            "❌ Не распознал. Введи в формате:\n"
            "`Январь 120000`  или  `март 85000 кВт·ч`",
            parse_mode="Markdown",
        )
        return

    month, T, W_a = parsed
    constants: NetworkConstants = ctx.user_data.get("constants")
    result = calculate_monthly_losses(constants, month, T, W_a)

    await msg.reply_text(result.format_report(), parse_mode="Markdown")
    set_state(ctx, S_WAIT_MORE)
    await msg.reply_text(
        "📅 Рассчитать потери за ещё один месяц?",
        reply_markup=kb_next_month(),
    )


# ── Callback-кнопки ────────────────────────────────────────────

async def on_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q   = update.callback_query
    await q.answer()
    data = q.data

    if data == "cb_start":
        set_state(ctx, S_WAIT_COMMENT)
        await q.message.reply_text(
            "📕 *Как подготовить схему*\n\n"
            "🔴 *Шаг 1 — Обведи нужные элементы красной линией:*\n"
            "• Нарисуй поверх схемы красным маркером/ручкой\n"
            "• Линия должна проходить по участкам, где считаем потери\n"
            "• Захвати символы трансформаторов если они входят в расчёт\n"
            "• Чем толще и ярче — тем точнее распознавание\n\n"
            "❗️ *Важно:*\n"
            "• Схема сама не должна быть красной — иначе возьми другой цвет\n"
            "• Подписи рядом с обведёнными элементами должны быть читаемы:\n"
            "  марка кабеля (ААБ2п 3×120, СИП 4×95…), длина, марка ТП\n\n"
            "✏️ *Шаг 2 — Напиши комментарий:*\n"
            "Уточни что обведено или добавь данные которых нет на схеме\n"
            "_(или напиши: «всё указано на схеме»)_",
            parse_mode="Markdown",
        )

    elif data == "cb_ok":
        scheme_data = ctx.user_data.get("scheme_data", {})
        await q.message.reply_text("📊 Генерирую Excel...")
        try:
            excel_bytes = generate_excel(scheme_data)
        except Exception as e:
            logger.error(f"Excel: {e}")
            await q.message.reply_text("❌ Ошибка генерации. /start")
            return
        set_state(ctx, S_WAIT_EXCEL)
        await q.message.reply_document(
            document=io.BytesIO(excel_bytes),
            filename="расчёт_потерь_326.xlsx",
            caption=(
                "📎 *Excel-шаблон с данными из схемы*\n\n"
                "1️⃣ Открой → лист «1. Параметры сети»\n"
                "2️⃣ Проверь все жёлтые ячейки\n"
                "3️⃣ Исправь ошибки распознавания\n"
                "4️⃣ Загрузи исправленный файл сюда ⬆️"
            ),
            parse_mode="Markdown",
        )

    elif data == "cb_retry":
        set_state(ctx, S_WAIT_IMAGE)
        await q.message.reply_text("📎 Загрузи другое изображение схемы:")

    elif data == "cb_month_yes":
        set_state(ctx, S_WAIT_MONTH)
        await q.message.reply_text(
            "📅 Введи месяц и объём потребления:\n\n"
            "Примеры:\n"
            "`Январь 120000`\n"
            "`март 85 000 кВт·ч`\n\n"
            "⚠️ W_а — показания входного счётчика на вводе 10 кВ "
            "(весь отпуск объекта, не сумма 0.4 кВ счётчиков).",
            parse_mode="Markdown",
        )

    elif data == "cb_month_no":
        set_state(ctx, S_IDLE)
        await q.message.reply_text(
            "✅ Расчёт завершён.\n\n"
            "Константы в Excel (листы 2–3):\n"
            "• Линии: `ΔW = C × W²а / T`\n"
            "• Трансформаторы: `ΔW = A×T + B×W²а/T`\n\n"
            "Для нового расчёта: /start",
            parse_mode="Markdown",
        )
        ctx.user_data.clear()


# ── /cancel ────────────────────────────────────────────────────

async def cmd_cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    set_state(ctx, S_IDLE)
    await update.message.reply_text("Отменено. /start — начать заново.")


# ── Запуск ─────────────────────────────────────────────────────

async def post_init(app: Application):
    await app.bot.set_my_commands([
        BotCommand("start",  "⚡ Начать расчёт потерь"),
        BotCommand("cancel", "❌ Отменить"),
    ])


def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).post_init(post_init).build()

    app.add_handler(CommandHandler("start",  cmd_start))
    app.add_handler(CommandHandler("cancel", cmd_cancel))
    app.add_handler(CallbackQueryHandler(on_callback))
    app.add_handler(MessageHandler(filters.ALL & ~filters.COMMAND, on_message))

    logger.info("Бот потерь запущен...")
    app.run_polling()


if __name__ == "__main__":
    main()
