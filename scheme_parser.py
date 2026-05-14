"""
Анализ однолинейной схемы через GPT-4o Vision.
Возвращает структурированные данные по линиям и трансформаторам.
"""

import base64
import json
import logging
import re

from openai import OpenAI

logger = logging.getLogger(__name__)

PARSE_PROMPT = """Ты — эксперт по электрическим сетям.
На изображении — однолинейная схема электроснабжения.
Пользователь обвёл красной линией элементы, потери в которых нужно рассчитать.
Пользователь указал: {user_comment}

ЕСЛИ на схеме есть красная (или другая выделяющая) обводка — извлекай данные ТОЛЬКО по обведённым элементам.
ЕСЛИ обводки нет — извлекай все элементы схемы.

Верни ТОЛЬКО валидный JSON без пояснений:
{
  "lines": [
    {
      "name": "обозначение/название участка",
      "type": "ВЛ или КЛ",
      "mark": "марка провода/кабеля (например АС-50, ААБ2п 3×120, СИП 4×95)",
      "r0": удельное_сопротивление_Ом_км_или_null,
      "length": длина_км,
      "voltage": напряжение_кВ,
      "note": "доп. пометки или пустая строка"
    }
  ],
  "transformers": [
    {
      "name": "обозначение ТП/КТП",
      "mark": "марка (ТМ-400/10, ТМГ-630/10 и т.д.)",
      "P0": потери_хх_кВт_или_null,
      "Pk": потери_кз_кВт_или_null,
      "Sn": мощность_кВА,
      "n": количество_штук,
      "note": ""
    }
  ],
  "source_voltage": напряжение_питания_кВ,
  "schema_description": "краткое описание схемы 1-2 предложения",
  "unrecognized": ["список элементов которые не удалось распознать"]
}

Правила:
- r0 = null если марка неизвестна (заполним из справочника)
- P0/Pk = null если не указаны на схеме (заполним из паспортных данных)
- Для ВЛ и КЛ 10 кВ voltage=10, для 0.4 кВ voltage=0.38
- Если длина указана в метрах — переведи в км
- n = 1 если не указано иное
"""

# Справочник r0 по маркам (для автозаполнения когда GPT не знает)
R0_REFERENCE = {
    "зас-50": 0.650, "зас-35": 0.898,
    "а-25": 1.280, "а-35": 0.898, "а-50": 0.640, "а-70": 0.457, "а-95": 0.337,
    "ас-50": 0.640, "ас-70": 0.422, "ас-95": 0.306, "ас-120": 0.249,
    "сип 4х50": 0.641, "сип 4х70": 0.443, "сип 4х95": 0.326, "сип 4х120": 0.258,
    "сип 3х50": 0.641, "сип 3х70": 0.443, "сип 3х95": 0.326,
    "ааб2п 3х50": 0.620, "ааб2п 3х70": 0.443, "ааб2п 3х95": 0.325,
    "ааб2п 3х120": 0.258, "ааб2п 3х150": 0.206, "ааб2п 3х185": 0.167,
    "аашв 3х50": 0.620, "аашв 3х95": 0.325, "аашв 3х120": 0.258, "аашв 3х150": 0.206,
    "апвбшв 4х50": 0.641, "апвбшв 4х95": 0.326, "апвбшв 4х120": 0.258,
    "ввг 4х50": 0.387, "ввг 4х95": 0.194,
}

# Паспортные данные трансформаторов
TR_REFERENCE = {
    "тм-25":   {"P0": 0.135, "Pk": 0.600},
    "тм-40":   {"P0": 0.175, "Pk": 0.880},
    "тм-63":   {"P0": 0.245, "Pk": 1.280},
    "тм-100":  {"P0": 0.365, "Pk": 1.970},
    "тм-160":  {"P0": 0.540, "Pk": 2.650},
    "тм-250":  {"P0": 0.820, "Pk": 3.700},
    "тм-400":  {"P0": 0.950, "Pk": 5.900},
    "тм-630":  {"P0": 1.310, "Pk": 7.600},
    "тм-1000": {"P0": 1.750, "Pk": 11.50},
    "тм-1600": {"P0": 2.650, "Pk": 16.50},
    "тмг-100": {"P0": 0.290, "Pk": 1.970},
    "тмг-160": {"P0": 0.390, "Pk": 2.650},
    "тмг-250": {"P0": 0.530, "Pk": 3.700},
    "тмг-400": {"P0": 0.660, "Pk": 5.900},
    "тмг-630": {"P0": 0.900, "Pk": 7.600},
    "тмн-1000":{"P0": 1.800, "Pk": 11.60},
    "тмн-1600":{"P0": 2.700, "Pk": 16.50},
    "тмн-2500":{"P0": 3.850, "Pk": 23.50},
}


def _lookup_r0(mark: str) -> float | None:
    """Поиск r0 по марке провода/кабеля."""
    mark_low = mark.lower().replace("×", "x").replace(" ", "")
    for key, val in R0_REFERENCE.items():
        if key.replace(" ", "").replace("х", "x") in mark_low:
            return val
    return None


def _lookup_tr(mark: str) -> dict | None:
    """Поиск паспортных данных трансформатора."""
    mark_low = mark.lower()
    for key, val in TR_REFERENCE.items():
        if key in mark_low:
            return val
    return None


def parse_scheme(
    image_bytes: bytes,
    user_comment: str,
    openai_key: str,
) -> dict:
    """
    Анализирует изображение схемы через GPT-4o Vision.
    Возвращает словарь с данными по линиям и трансформаторам.
    """
    client = OpenAI(api_key=openai_key)

    b64 = base64.b64encode(image_bytes).decode()
    prompt = PARSE_PROMPT.format(user_comment=user_comment or "не указан")

    try:
        resp = client.chat.completions.create(
            model="gpt-4o",
            messages=[{
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
                ],
            }],
            temperature=0,
            max_tokens=2000,
        )
        raw = resp.choices[0].message.content.strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        data = json.loads(raw)
    except Exception as e:
        logger.error(f"GPT Vision ошибка: {e}")
        return {"lines": [], "transformers": [], "error": str(e)}

    # Дозаполняем r0 из справочника если GPT не нашёл
    for line in data.get("lines", []):
        if not line.get("r0") and line.get("mark"):
            line["r0"] = _lookup_r0(line["mark"])

    # Дозаполняем P0/Pk трансформаторов
    for tr in data.get("transformers", []):
        if tr.get("mark"):
            ref = _lookup_tr(tr["mark"])
            if ref:
                if not tr.get("P0"):
                    tr["P0"] = ref["P0"]
                if not tr.get("Pk"):
                    tr["Pk"] = ref["Pk"]

    return data


def format_parse_summary(data: dict) -> str:
    """Форматирует результат распознавания для показа пользователю."""
    lines_out = []
    lines_out.append("🔍 *Результат распознавания схемы:*\n")

    if data.get("schema_description"):
        lines_out.append(f"_{data['schema_description']}_\n")

    lines_out.append(f"⚡ *Линии ({len(data.get('lines', []))} уч-ков):*")
    for l in data.get("lines", []):
        r0_str = f"r₀={l['r0']}" if l.get("r0") else "r₀=❓"
        lines_out.append(
            f"  • {l.get('name','')} [{l.get('type','')}] {l.get('mark','')} "
            f"L={l.get('length','')} км, U={l.get('voltage','')} кВ, {r0_str}"
        )

    lines_out.append(f"\n🔌 *Трансформаторы ({len(data.get('transformers', []))} шт):*")
    for t in data.get("transformers", []):
        p0_str = f"ΔP₀={t['P0']}" if t.get("P0") else "ΔP₀=❓"
        pk_str = f"ΔPк={t['Pk']}" if t.get("Pk") else "ΔPк=❓"
        lines_out.append(
            f"  • {t.get('name','')} {t.get('mark','')} "
            f"Sн={t.get('Sn','')} кВА, n={t.get('n',1)}, {p0_str}, {pk_str}"
        )

    if data.get("unrecognized"):
        lines_out.append("\n⚠️ *Не распознано:*")
        for u in data["unrecognized"]:
            lines_out.append(f"  • {u}")

    return "\n".join(lines_out)
