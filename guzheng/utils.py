# -*- coding: utf-8 -*-
"""
文字處理工具函式。
純函式，無副作用，不依賴其他模組。
"""
from __future__ import annotations
import re
from typing import List


def clean_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def split_lines(value: str) -> List[str]:
    """將儲存格多行文字拆成非空字串清單。"""
    if not value:
        return []
    parts = [x.strip() for x in str(value).replace("\r", "\n").split("\n")]
    return [p for p in parts if p]


def split_names(value: str) -> List[str]:
    """支援逗號、頓號、換行分隔的姓名字串 → 清單。"""
    if not value:
        return []
    text = str(value).replace("\n", ",").replace("、", ",").replace("，", ",")
    parts = [x.strip() for x in text.split(",")]
    return [p for p in parts if p]


def normalize_raw_name(name: str) -> str:
    """
    統一姓名格式，用於比對與查表：
    - 含英文 → 取英文段
    - 純中文三字以上 → 去姓取名（後兩字）
    - 兩字中文 → 原樣保留
    """
    if not name:
        return ""

    text = str(name).strip()
    text = text.replace("（", "(").replace("）", ")")
    text = re.sub(r"\s+", "", text)

    if re.search(r"[A-Za-z]", text):
        m = re.search(r"[A-Za-z][A-Za-z0-9_ -]*", text)
        if m:
            return m.group(0).strip()

    chinese = re.sub(r"[^一-鿿]", "", text)
    if len(chinese) >= 3:
        return chinese[-2:]
    return chinese or text


def normalize_stand_full(stand: str) -> str:
    """將各種寫法的箏架名稱統一成標準名稱（例：黑Ａ架 → 黑A架）。"""
    s = clean_text(stand).replace(" ", "")
    s = s.replace("黑Ａ架", "黑A架").replace("Ａ", "A")
    mapping = {
        "黑架":  "黑架",
        "紅架":  "紅架",
        "綠架":  "綠架",
        "A架":   "A架",
        "T架":   "T架",
        "黑A架": "黑A架",
        "國風架": "國風架",
    }
    return mapping.get(s, s)


def stand_for_performer(stand: str) -> str:
    """將完整箏架名稱轉為表演者欄的簡短顯示（例：黑架 → 黑）。"""
    s = normalize_stand_full(stand)
    mapping = {
        "黑架":  "黑",
        "紅架":  "紅",
        "綠架":  "綠",
        "A架":   "A架",
        "T架":   "T架",
        "黑A架": "黑A架",
        "國風架": "國風架",
    }
    return mapping.get(s, s)


def strip_position_suffix(text: str) -> str:
    """移除古箏名稱中的位置括號，例如「古箏一號（左）」→「古箏一號」。"""
    t = clean_text(text)
    t = re.sub(r"（[^）]*）", "", t)
    t = re.sub(r"\([^)]*\)", "", t)
    return t.strip()


def dedupe_keep_order(items: List[str]) -> List[str]:
    """去除重複元素，保留原始順序。"""
    seen = set()
    out = []
    for item in items:
        if item not in seen:
            seen.add(item)
            out.append(item)
    return out
