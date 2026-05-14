"""
Читает проверенный Excel-шаблон и извлекает константы C, A, B.
"""

import io
import logging
import re

import pandas as pd
from losses_calc import NetworkConstants

logger = logging.getLogger(__name__)

# Справочник паспортных данных трансформаторов — дублируем из scheme_parser
# чтобы excel_reader был независимым
TR_REF = {
    "тм-25": (0.135, 0.600), "тм-40": (0.175, 0.880),
    "тм-63": (0.245, 1.280), "тм-100": (0.365, 1.970),
    "тм-160": (0.540, 2.650), "тм-250": (0.820, 3.700),
    "тм-400": (0.950, 5.900), "тм-630": (1.310, 7.600),
    "тм-1000": (1.750, 11.50), "тм-1600": (2.650, 16.50),
    "тмг-100": (0.290, 1.970), "тмг-160": (0.390, 2.650),
    "тмг-250": (0.530, 3.700), "тмг-400": (0.660, 5.900),
    "тмг-630": (0.900, 7.600), "тмг-1000": (1.400, 11.50),
    "тмн-1000": (1.800, 11.60), "тмн-1600": (2.700, 16.50),
    "тмн-2500": (3.850, 23.50), "тмн-4000": (5.800, 33.50),
    "ктп-400": (0.950, 5.900), "ктп-630": (1.310, 7.600),
    "ктп-1000": (1.750, 11.50),
}


def _lookup_tr_ref(mark: str):
    """Ищет P0, Pk по марке (нечёткий поиск)."""
    if not mark or mark in ("nan", ""):
        return None, None
    # Нормализуем: убираем пробелы, переводим в нижний регистр, / → -
    norm = re.sub(r"[\s/]", "-", mark.lower())
    norm = re.sub(r"-+", "-", norm)   # двойные дефисы → один
    for key, (p0, pk) in TR_REF.items():
        # Ищем ключ в нормализованной строке
        if key in norm or key.replace("-", "") in norm.replace("-", ""):
            return p0, pk
    return None, None


def _safe_float(val) -> float | None:
    """Безопасно конвертирует значение в float, None если пусто/NaN."""
    try:
        f = float(val)
        return None if f != f else f   # NaN проверка
    except Exception:
        return None


def read_constants_from_excel(file_bytes: bytes) -> NetworkConstants | None:
    """
    Читает Excel (лист 1 — Параметры сети), пересчитывает константы C, A, B.
    Возвращает NetworkConstants или None при ошибке.
    """
    try:
        buf = io.BytesIO(file_bytes)
        df = pd.read_excel(buf, sheet_name=0, header=None)
    except Exception as e:
        logger.error(f"Ошибка чтения Excel: {e}")
        return None

    logger.info(f"Excel прочитан: {df.shape[0]} строк, {df.shape[1]} колонок")

    # ── Общие коэффициенты (строки Excel 6-8 = pandas 5-7, колонка C = pandas 2)
    try:
        k2f     = _safe_float(df.iloc[5, 2]) or 1.10
        kk      = _safe_float(df.iloc[6, 2]) or 0.99
        cos_phi = _safe_float(df.iloc[7, 2]) or 0.90
    except Exception:
        k2f, kk, cos_phi = 1.10, 0.99, 0.90

    cos2 = cos_phi ** 2
    logger.info(f"Коэффициенты: K²ф={k2f}, k_k={kk}, cosφ={cos_phi}")

    # ── Линии (Excel строки 12-21 = pandas 11-20)
    # Структура: col B(1)=name, C(2)=type, D(3)=mark, E(4)=r0, F(5)=L, G(6)=U
    lines_list = []
    C_total = 0.0

    for i in range(10):
        row_idx = 11 + i
        try:
            name = str(df.iloc[row_idx, 1]).strip()
            r0   = _safe_float(df.iloc[row_idx, 4])
            L    = _safe_float(df.iloc[row_idx, 5])
            U    = _safe_float(df.iloc[row_idx, 6])
        except Exception as e:
            logger.debug(f"Линия row{row_idx}: исключение {e}")
            continue

        if not name or name == "nan":
            continue
        if r0 is None or L is None or U is None or U == 0:
            logger.debug(f"Линия '{name}': пропуск (r0={r0}, L={L}, U={U})")
            continue

        C = k2f * kk * r0 * L / (U ** 2 * cos2 * 1000)
        lines_list.append({"name": name, "C": C})
        C_total += C
        logger.info(f"Линия '{name}': r0={r0}, L={L}, U={U} → C={C:.6E}")

    # ── Трансформаторы (Excel строки 25-32 = pandas 24-31)
    # Структура: col B(1)=name, C(2)=mark, D(3)=P0, E(4)=Pk, F(5)=Sn, G(6)=n
    trs_list = []
    A_total = 0.0
    B_total = 0.0

    for i in range(8):
        row_idx = 24 + i
        try:
            name = str(df.iloc[row_idx, 1]).strip()
            mark = str(df.iloc[row_idx, 2]).strip()
            P0   = _safe_float(df.iloc[row_idx, 3])
            Pk   = _safe_float(df.iloc[row_idx, 4])
            Sn   = _safe_float(df.iloc[row_idx, 5])
            n_val = _safe_float(df.iloc[row_idx, 6])
        except Exception as e:
            logger.debug(f"ТР row{row_idx}: исключение {e}")
            continue

        if not name or name == "nan":
            continue
        if Sn is None or Sn == 0:
            logger.debug(f"ТР '{name}': пропуск (Sn={Sn})")
            continue

        n = int(n_val) if n_val else 1

        # Если P0 или Pk не заполнены — ищем по справочнику через марку
        if P0 is None or Pk is None:
            ref_p0, ref_pk = _lookup_tr_ref(mark)
            if P0 is None:
                P0 = ref_p0
            if Pk is None:
                Pk = ref_pk
            if P0 is None or Pk is None:
                # Пробуем найти по имени (может содержать марку)
                ref_p0, ref_pk = _lookup_tr_ref(name)
                if P0 is None: P0 = ref_p0
                if Pk is None: Pk = ref_pk

        if P0 is None or Pk is None:
            logger.warning(
                f"ТР '{name}' (марка: '{mark}'): P0/Pk не найдены ни в Excel, "
                f"ни в справочнике. Добавь данные вручную в Excel."
            )
            continue

        A = P0 * n
        B = Pk * n / (Sn ** 2 * cos2)
        trs_list.append({"name": name, "A": A, "B": B})
        A_total += A
        B_total += B
        logger.info(
            f"ТР '{name}': P0={P0}, Pk={Pk}, Sn={Sn}, n={n} "
            f"→ A={A:.3f}, B={B:.6E}"
        )

    logger.info(
        f"Итого: линий={len(lines_list)}, C={C_total:.6E} | "
        f"ТР={len(trs_list)}, A={A_total:.3f}, B={B_total:.6E}"
    )

    if C_total == 0 and A_total == 0:
        logger.error("Все константы нулевые — данные не прочитаны")
        return None

    return NetworkConstants(
        C=C_total, A=A_total, B=B_total,
        lines=lines_list, transformers=trs_list,
    )
