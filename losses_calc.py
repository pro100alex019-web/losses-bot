"""
Модуль расчёта потерь электроэнергии за месяц.
Используется в Telegram-боте электросетевых потерь.

После вывода констант формул (C, A, B) бот спрашивает пользователя:
  → хочет ли рассчитать потери за конкретный месяц?
  → если да — запрашивает месяц и W_а (кВт·ч)
  → рассчитывает и выводит детальный результат
"""

from dataclasses import dataclass, field
from typing import Optional

# Часов в каждом месяце (невисокосный год)
MONTH_HOURS = {
    "январь": 744, "февраль": 672, "март": 744,
    "апрель": 720, "май": 744, "июнь": 720,
    "июль": 744, "август": 744, "сентябрь": 720,
    "октябрь": 744, "ноябрь": 720, "декабрь": 744,
    # Сокращения
    "янв": 744, "фев": 672, "мар": 744, "апр": 720,
    "июн": 720, "июл": 744, "авг": 744, "сен": 720,
    "окт": 744, "ноя": 720, "дек": 744,
}


@dataclass
class NetworkConstants:
    """Константы формул, рассчитанные из параметров сети."""
    # Линии: ΔW_лин = C × W²а / T
    C: float = 0.0          # суммарная константа линий

    # Трансформаторы: ΔW_тр = A × T + B × W²а / T
    A: float = 0.0          # постоянная составляющая (х.х.), кВт
    B: float = 0.0          # нагрузочная составляющая

    # Детализация по элементам (для отчёта)
    lines: list = field(default_factory=list)       # [{name, C}, ...]
    transformers: list = field(default_factory=list) # [{name, A, B}, ...]

    def has_data(self) -> bool:
        return self.C > 0 or self.A > 0 or self.B > 0


@dataclass
class MonthlyResult:
    """Результат расчёта потерь за месяц."""
    month: str
    T: int          # часов в месяце
    W_a: float      # расход эл. энергии, кВт·ч

    # Составляющие потерь
    dW_lines: float = 0.0   # потери в линиях
    dW_xx: float = 0.0      # потери х.х. трансформаторов
    dW_load: float = 0.0    # нагрузочные потери трансформаторов
    dW_total: float = 0.0   # итого
    pct: float = 0.0        # % от W_а

    def format_report(self) -> str:
        """Формирует текстовый отчёт для Telegram."""
        sep = "─" * 30
        dW_tr = self.dW_xx + self.dW_load

        lines = [
            f"⚡ Расчёт потерь за {self.month.capitalize()}",
            f"T = {self.T} ч  |  W_а = {self.W_a:,.0f} кВт·ч",
            sep,
            "Формулы (Приказ 326):",
            f"  ΔW_лин = C × W²а / T  =  {self.dW_lines:,.1f} кВт·ч",
            f"  ΔW_хх  = A × T        =  {self.dW_xx:,.1f} кВт·ч",
            f"  ΔW_н   = B × W²а / T  =  {self.dW_load:,.1f} кВт·ч",
            f"  ΔW_тр  = ΔW_хх + ΔW_н =  {dW_tr:,.1f} кВт·ч",
            sep,
            f"ΔW = ΔW_лин + ΔW_тр",
            f"ΔW = {self.dW_lines:,.1f} + {dW_tr:,.1f} = {self.dW_total:,.1f} кВт·ч",
            sep,
            f"Уровень потерь: {self.dW_total:,.1f} / {self.W_a:,.0f} × 100 = {self.pct:.2f}%",
        ]

        if self.pct < 5:
            lines.append("✅ Норма (< 5%)")
        elif self.pct < 10:
            lines.append("⚠️ Повышенные (5–10%) — рекомендуется анализ")
        else:
            lines.append("🔴 Высокие (> 10%) — требуется проверка сети")

        lines.append("\nРасчёт по Приказу Минэнерго № 326 от 30.12.2008")
        return "\n".join(lines)


def parse_month(text: str) -> Optional[tuple[str, int]]:
    """
    Разбирает строку вида 'Январь 120000' или 'март 85000'.
    Возвращает (название_месяца, T_часов) или None.
    """
    parts = text.strip().lower().split()
    for part in parts:
        if part in MONTH_HOURS:
            return part, MONTH_HOURS[part]
    return None


def parse_consumption(text: str) -> Optional[float]:
    """Извлекает число из строки (потребление кВт·ч)."""
    import re
    cleaned = re.sub(r"[^\d\s.,]", " ", text)
    nums = re.findall(r"\d[\d\s]*(?:[.,]\d+)?", cleaned)
    for n in nums:
        try:
            v = float(n.replace(" ", "").replace(",", "."))
            if v > 0:
                return v
        except ValueError:
            continue
    return None


def parse_month_and_consumption(text: str) -> Optional[tuple[str, int, float]]:
    """
    Разбирает строку с месяцем и потреблением.
    Примеры: 'Январь 120000', 'март 85 000 кВт·ч', 'февраль, 92500'
    Возвращает (месяц, T, W_a) или None.
    """
    month_result = parse_month(text)
    if not month_result:
        return None
    month, T = month_result

    # Убираем слово месяца из строки перед поиском числа
    import re
    text_no_month = re.sub(month[:3], "", text.lower(), flags=re.IGNORECASE)
    consumption = parse_consumption(text_no_month) or parse_consumption(text)
    if not consumption:
        return None

    return month, T, consumption


def calculate_monthly_losses(
    constants: NetworkConstants,
    month: str,
    T: int,
    W_a: float,
) -> MonthlyResult:
    """
    Рассчитывает потери за месяц по формулам Приказа 326.

    Линии:         ΔW_лин = C × W²а / T
    Трансформаторы: ΔW_тр = A × T + B × W²а / T
    """
    result = MonthlyResult(month=month, T=T, W_a=W_a)

    result.dW_lines = constants.C * (W_a ** 2) / T if T > 0 else 0
    result.dW_xx   = constants.A * T
    result.dW_load = constants.B * (W_a ** 2) / T if T > 0 else 0
    result.dW_total = result.dW_lines + result.dW_xx + result.dW_load
    result.pct = (result.dW_total / W_a * 100) if W_a > 0 else 0

    return result
