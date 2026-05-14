"""
Читает проверенный Excel-шаблон и извлекает константы C, A, B.
"""

import io
import logging

import pandas as pd
from losses_calc import NetworkConstants

logger = logging.getLogger(__name__)


def read_constants_from_excel(file_bytes: bytes) -> NetworkConstants | None:
    """
    Читает Excel (лист 'Параметры сети'), пересчитывает константы.
    Возвращает NetworkConstants или None при ошибке.
    """
    try:
        buf = io.BytesIO(file_bytes)
        # Лист 1: параметры сети
        df = pd.read_excel(buf, sheet_name=0, header=None)
    except Exception as e:
        logger.error(f"Ошибка чтения Excel: {e}")
        return None

    # Читаем общие коэффициенты (строки 5-8, колонка C = индекс 2)
    try:
        k2f    = float(df.iloc[5, 2])   # K²ф
        kk     = float(df.iloc[6, 2])   # k_k
        cos_phi = float(df.iloc[7, 2])  # cos φ
    except Exception:
        k2f, kk, cos_phi = 1.10, 0.99, 0.90

    cos2 = cos_phi ** 2

    # Читаем линии (строки 11-21, колонки B-G)
    # Заголовок на строке 10 (индекс 10), данные с 11 (индекс 11)
    lines_list = []
    C_total = 0.0

    for i in range(10):  # до 10 линий
        row_idx = 11 + i
        try:
            name = str(df.iloc[row_idx, 1]).strip()
            r0   = float(df.iloc[row_idx, 4])
            L    = float(df.iloc[row_idx, 5])
            U    = float(df.iloc[row_idx, 6])
        except Exception:
            continue
        if not name or name in ("nan", "") or r0 != r0 or L != L:
            continue
        C = k2f * kk * r0 * L / (U ** 2 * cos2 * 1000)
        lines_list.append({"name": name, "C": C})
        C_total += C

    # Читаем трансформаторы (строки 24-32, колонки B-G)
    trs_list = []
    A_total = 0.0
    B_total = 0.0

    for i in range(8):  # до 8 трансформаторов
        row_idx = 24 + i
        try:
            name = str(df.iloc[row_idx, 1]).strip()
            P0   = float(df.iloc[row_idx, 3])
            Pk   = float(df.iloc[row_idx, 4])
            Sn   = float(df.iloc[row_idx, 5])
            n    = int(float(df.iloc[row_idx, 6]))
        except Exception:
            continue
        if not name or name in ("nan", "") or P0 != P0 or Pk != Pk:
            continue
        A = P0 * n
        B = Pk * n / (Sn ** 2 * cos2)
        trs_list.append({"name": name, "A": A, "B": B})
        A_total += A
        B_total += B

    if C_total == 0 and A_total == 0:
        return None

    return NetworkConstants(
        C=C_total, A=A_total, B=B_total,
        lines=lines_list, transformers=trs_list,
    )
