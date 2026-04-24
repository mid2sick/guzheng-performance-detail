# -*- coding: utf-8 -*-
"""
Excel 讀取層。
每個函式對應來源 xlsx 的一個工作表，回傳 Python 原生型別或 models。
"""
from __future__ import annotations
import math
from collections import defaultdict
from typing import Dict, List, Optional, Tuple

from .models import SongAsset, SongPeople
from .utils import (
    clean_text,
    dedupe_keep_order,
    normalize_raw_name,
    normalize_stand_full,
    split_lines,
    split_names,
)


# ---------------------------------------------------------
# 曲目工作表
# ---------------------------------------------------------

def read_song_order_and_assets(wb) -> Tuple[List[str], Dict[str, SongAsset]]:
    """
    從「曲目」工作表讀取曲目順序與舞台物件資訊。
    row 2 = 表頭，row 3 起 = 資料。
    """
    ws = wb["曲目"]
    header = {clean_text(ws.cell(2, c).value): c for c in range(1, ws.max_column + 1)}

    col_song      = header["曲目名稱"]
    col_performers = header["表演者"]
    col_stands    = header["慣用箏架"]
    col_guzheng   = header["古箏"]

    song_order: List[str] = []
    song_assets: Dict[str, SongAsset] = {}

    for r in range(3, ws.max_row + 1):
        song = clean_text(ws.cell(r, col_song).value)
        if not song:
            continue

        guzheng_names = split_lines(clean_text(ws.cell(r, col_guzheng).value))
        stands        = [normalize_stand_full(x) for x in split_lines(clean_text(ws.cell(r, col_stands).value))]
        performers    = split_lines(clean_text(ws.cell(r, col_performers).value))

        guzheng_count  = len(guzheng_names) or len(stands) or len(performers)
        wood_chair_count = guzheng_count

        song_order.append(song)
        song_assets[song] = SongAsset(
            song=song,
            guzheng_names=guzheng_names,
            stands=stands,
            guzheng_count=guzheng_count,
            wood_chair_count=wood_chair_count,
        )

    return song_order, song_assets


def read_song_duration_minutes(wb) -> Dict[str, int]:
    """
    從「曲目」工作表讀取「曲目時長」欄（分鐘，無條件進位）。
    """
    ws = wb["曲目"]
    header = {clean_text(ws.cell(2, c).value): c for c in range(1, ws.max_column + 1)}

    col_song     = header["曲目名稱"]
    col_duration = header["曲目時長"]

    result: Dict[str, int] = {}
    for r in range(3, ws.max_row + 1):
        song = clean_text(ws.cell(r, col_song).value)
        if not song:
            continue
        raw = ws.cell(r, col_duration).value
        if raw is None or str(raw).strip() == "":
            continue
        try:
            result[song] = math.ceil(float(raw))
        except Exception:
            continue

    return result


# ---------------------------------------------------------
# 明細工作表
# ---------------------------------------------------------

def read_detail_people(wb) -> Dict[str, List[Tuple[str, str, str]]]:
    """
    從「明細」讀取每首曲目的古箏表演者清單。
    回傳 {曲目: [(姓名, 古箏編號, 箏架類型), ...]}
    """
    ws = wb["明細"]
    header = {clean_text(ws.cell(2, c).value): c for c in range(1, ws.max_column + 1)}

    col_song   = header["曲目"]
    col_name   = header["表演者"]
    col_guzheng = header["古箏"]
    col_stand  = header["慣用箏架 (from 表演者)"]

    result: Dict[str, List[Tuple[str, str, str]]] = defaultdict(list)
    for r in range(3, ws.max_row + 1):
        song = clean_text(ws.cell(r, col_song).value)
        if not song:
            continue
        name    = normalize_raw_name(clean_text(ws.cell(r, col_name).value))
        guzheng = clean_text(ws.cell(r, col_guzheng).value)
        stand   = normalize_stand_full(clean_text(ws.cell(r, col_stand).value))
        result[song].append((name, guzheng, stand))

    return result


# ---------------------------------------------------------
# 演出分配工作表
# ---------------------------------------------------------

def _detect_role(role_cell: str) -> Optional[str]:
    """
    將聲部欄文字轉為內部 role key。
    只接受明確的單一角色；舊格式或不明格式印出警告並回傳 None。
    """
    role = clean_text(role_cell).replace(" ", "")

    if role == "低音提琴":
        return "bass"
    if role == "鋼琴伴奏":
        return "piano"
    if role == "打擊":
        return "percussion"

    if role in ("打擊/鋼琴伴奏", "鋼琴伴奏/打擊"):
        print(f"[警告] 發現未拆分舊格式聲部：{role}")
        return None

    if role:
        print(f"[警告] 無法辨識的聲部：{role}")
    return None


def read_performance_roles(wb) -> Dict[str, Dict[str, List[str]]]:
    """
    從「演出分配」讀取打擊 / 鋼琴 / 低音提琴的表演者。
    回傳 {曲目: {"percussion": [...], "piano": [...], "bass": [...]}}
    """
    ws = wb["演出分配"]
    header = {clean_text(ws.cell(2, c).value): c for c in range(1, ws.max_column + 1)}

    col_song   = header["曲目"]
    col_people = header["表演者"]
    col_role   = header["聲部"]

    result: Dict[str, Dict[str, List[str]]] = defaultdict(lambda: {
        "percussion": [], "piano": [], "bass": [],
    })

    for r in range(3, ws.max_row + 1):
        song = clean_text(ws.cell(r, col_song).value)
        if not song:
            continue
        role_key = _detect_role(ws.cell(r, col_role).value)
        if not role_key:
            continue
        people = [normalize_raw_name(x) for x in split_names(clean_text(ws.cell(r, col_people).value))]
        result[song][role_key].extend(people)

    for song, role_map in result.items():
        for k, v in role_map.items():
            role_map[k] = dedupe_keep_order(v)

    return result


# ---------------------------------------------------------
# 後台工作分配工作表
# ---------------------------------------------------------

def read_backstage_staff(wb) -> Dict[str, str]:
    """
    從「表演後台工作分配表」讀取後台人員與其角色。
    回傳 {姓名: role_key}，role_key 為
    "general_core" / "mic" / "control" / "curtain" / "tuning" 其中之一。
    """
    ws = wb["表演後台工作分配表"]
    header = {clean_text(ws.cell(2, c).value): c for c in range(1, ws.max_column + 1)}

    col_name = header["姓名"]
    col_role = header["後台角色"]

    result: Dict[str, str] = {}
    for r in range(3, ws.max_row + 1):
        raw_name = clean_text(ws.cell(r, col_name).value)
        if not raw_name:
            continue
        name = normalize_raw_name(raw_name)
        role = clean_text(ws.cell(r, col_role).value)
        result[name] = role

    return result
