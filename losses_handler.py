"""
Telegram ConversationHandler — расчёт потерь за месяц.

Состояния:
  SHOW_CONSTANTS   — вывод констант, предложение рассчитать за месяц
  WAITING_MONTH    — ожидание ввода месяца и W_а от пользователя
  NEXT_MONTH       — после расчёта: рассчитать ещё один месяц?
"""

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    CallbackQueryHandler, CommandHandler, ContextTypes,
    ConversationHandler, MessageHandler, filters,
)
from losses_calc import (
    NetworkConstants, calculate_monthly_losses,
    parse_month_and_consumption, MONTH_HOURS,
)

SHOW_CONSTANTS, WAITING_MONTH, NEXT_MONTH = range(3)


# ─── Клавиатуры ───────────────────────────────────────────────

def kb_calc_month():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Да, рассчитать за месяц", callback_data="calc_month_yes"),
        InlineKeyboardButton("🏁 Нет, завершить",          callback_data="calc_month_no"),
    ]])

def kb_next_month():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("📅 Рассчитать ещё месяц",   callback_data="calc_month_again"),
        InlineKeyboardButton("🏁 Завершить расчёт",        callback_data="calc_month_no"),
    ]])


# ─── Вспомогательная: форматировать константы ──────────────────

def _esc(text: str) -> str:
    """Экранирует спецсимволы Telegram Markdown v1 в динамических данных."""
    for ch in ["*", "_", "`", "["]:
        text = text.replace(ch, "\\" + ch)
    return text


def format_constants(c: NetworkConstants) -> str:
    sep = "─" * 28
    lines = [
        "📐 *Константы формул потерь:*\n",
        sep,
        "⚡ *Линии (ВЛ + КЛ):*",
        f"   ΔW_лин = `{c.C:.4E}` × W²а / T",
    ]
    if c.lines:
        for el in c.lines:
            lines.append(f"   • {_esc(str(el['name']))}: C = `{el['C']:.4E}`")
    lines += [
        "",
        "🔌 *Трансформаторы:*",
        f"   ΔW_тр = `{c.A:.3f}` × T  +  `{c.B:.4E}` × W²а / T",
        f"   где  A (х.х.) = `{c.A:.3f}` кВт,  B (нагруз.) = `{c.B:.4E}`",
    ]
    if c.transformers:
        for tr in c.transformers:
            lines.append(
                f"   • {_esc(str(tr['name']))}: A={tr['A']:.3f}, B={tr['B']:.4E}"
            )
    lines += [
        sep,
        "📌 _Константы рассчитаны из параметров сети и не меняются._",
        "_Подставляй W_а (кВт·ч) и T (ч) для любого месяца._",
    ]
    return "\n".join(lines)


# ─── Хэндлер: вход — показать константы и спросить о расчёте ──

async def show_constants_and_ask(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    constants: NetworkConstants,
) -> int:
    """
    Вызывается после того как бот рассчитал константы из шаблона.
    Показывает константы и предлагает рассчитать потери за месяц.
    """
    context.user_data["constants"] = constants

    await update.message.reply_text(
        format_constants(constants),
        parse_mode="Markdown",
    )
    await update.message.reply_text(
        "💬 *Хотите рассчитать потери за конкретный месяц?*\n\n"
        "Для этого нужно указать:\n"
        "• месяц\n"
        "• объём потребления электроэнергии W_а (кВт·ч) за этот месяц\n\n"
        "_Данные берутся из показаний счётчиков за расчётный период._",
        parse_mode="Markdown",
        reply_markup=kb_calc_month(),
    )
    return WAITING_MONTH


async def ask_month_yes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    await q.message.reply_text(
        "📅 Введи *месяц* и *объём потребления* в одном сообщении.\n\n"
        "Примеры:\n"
        "`Январь 120000`\n"
        "`март 85 000 кВт·ч`\n"
        "`февраль 92500`\n\n"
        "⚠️ *Важно:* W_а — это показания *входного счётчика* на уровне 10 кВ (весь отпуск объекта), не сумма счётчиков 0,4 кВ.\n\n"
        "_Число часов T для выбранного месяца будет подставлено автоматически._",
        parse_mode="Markdown",
    )
    return WAITING_MONTH


async def ask_month_no(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    await q.message.reply_text(
        "✅ Расчёт завершён.\n\n"
        "Константы формул сохранены в шаблоне Excel (Лист 4).\n"
        "Используй их для ежемесячного расчёта потерь:\n\n"
        "• Линии:          `ΔW = C × W²а / T`\n"
        "• Трансформаторы: `ΔW = A × T + B × W²а / T`",
        parse_mode="Markdown",
    )
    context.user_data.clear()
    return ConversationHandler.END


async def receive_month_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    parsed = parse_month_and_consumption(text)

    if not parsed:
        await update.message.reply_text(
            "❌ Не удалось распознать месяц или потребление.\n\n"
            "Введи в формате: `Январь 120000`\n"
            "_(месяц и число кВт·ч через пробел)_",
            parse_mode="Markdown",
        )
        return WAITING_MONTH

    month, T, W_a = parsed
    constants: NetworkConstants = context.user_data.get("constants")

    if not constants or not constants.has_data():
        await update.message.reply_text(
            "⚠️ Константы не найдены. Начни расчёт заново."
        )
        return ConversationHandler.END

    result = calculate_monthly_losses(constants, month, T, W_a)
    context.user_data["last_result"] = result

    await update.message.reply_text(
        result.format_report(),
        parse_mode="Markdown",
    )
    await update.message.reply_text(
        "📅 Рассчитать потери за ещё один месяц?",
        reply_markup=kb_next_month(),
    )
    return NEXT_MONTH


async def calc_month_again(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    await q.message.reply_text(
        "Введи следующий месяц и потребление:\n"
        "`Февраль 98000`",
        parse_mode="Markdown",
    )
    return WAITING_MONTH


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("Расчёт отменён.")
    return ConversationHandler.END


# ─── ConversationHandler ───────────────────────────────────────

def get_losses_monthly_handler() -> ConversationHandler:
    return ConversationHandler(
        entry_points=[],   # вход — внешний вызов show_constants_and_ask()
        states={
            WAITING_MONTH: [
                CallbackQueryHandler(ask_month_yes,  pattern="^calc_month_yes$"),
                CallbackQueryHandler(ask_month_no,   pattern="^calc_month_no$"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_month_input),
            ],
            NEXT_MONTH: [
                CallbackQueryHandler(calc_month_again, pattern="^calc_month_again$"),
                CallbackQueryHandler(ask_month_no,     pattern="^calc_month_no$"),
            ],
        },
        fallbacks=[
            CommandHandler("cancel", cancel),
            CallbackQueryHandler(ask_month_no, pattern="^calc_month_no$"),
        ],
        per_message=False,
    )
