# -*- coding: utf-8 -*-
"""
時間軸建構。
給定曲目順序與各曲時長，算出每個曲目列和換場列的時間區間。
"""
from __future__ import annotations
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple

# key 型別別名，方便閱讀
TransitionKey = Tuple[Optional[str], Optional[str], str]


def fmt_range(start_dt: datetime, end_dt: datetime) -> str:
    return f"{start_dt.strftime('%H:%M')}~{end_dt.strftime('%H:%M')}"


def build_timeline_maps(
    song_order:           List[str],
    duration_minutes_map: Dict[str, int],
    start_time_str:       str,
    opening_minutes:      int,
    intermission_after:   Optional[int],
    transition_minutes:   int,
    intermission_minutes: int,
) -> Tuple[Dict[str, str], Dict[TransitionKey, str]]:
    """
    回傳：
      song_time_map       — {曲目: "HH:MM~HH:MM"}
      transition_time_map — {(prev, next, row_type): "HH:MM~HH:MM"}

    row_type 說明：
      "pre_opening"         第一首前置列（開場前準備）
      "normal_transition"   一般換場
      "before_intermission" 中場前最後一首後置列（只有下箏，不安排換場）
      "intermission"        中場休息列
      "pre_second_half"     下半場第一首前置列
      "last_teardown"       最後一首拆台列
    """
    song_time_map:       Dict[str, str]              = {}
    transition_time_map: Dict[TransitionKey, str]    = {}

    current_dt = datetime.strptime(start_time_str, "%H:%M")

    # 開場前置列
    if song_order:
        opening_end = current_dt + timedelta(minutes=opening_minutes)
        transition_time_map[(None, song_order[0], "pre_opening")] = fmt_range(current_dt, opening_end)
        current_dt = opening_end

    for idx, song in enumerate(song_order):
        # 曲目時間
        song_end = current_dt + timedelta(minutes=duration_minutes_map.get(song, 0))
        song_time_map[song] = fmt_range(current_dt, song_end)
        current_dt = song_end

        # 最後一首：只留拆台列，不再往後排換場
        if idx == len(song_order) - 1:
            transition_time_map[(song, None, "last_teardown")] = ""
            continue

        next_song = song_order[idx + 1]

        # 中場休息：在第 intermission_after 首結束後插入
        if intermission_after is not None and (idx + 1) == intermission_after:
            transition_time_map[(song, next_song, "before_intermission")] = ""

            intermission_end = current_dt + timedelta(minutes=intermission_minutes)
            transition_time_map[(song, next_song, "intermission")] = fmt_range(current_dt, intermission_end)
            current_dt = intermission_end

            transition_time_map[(None, next_song, "pre_second_half")] = ""
            continue

        # 一般換場
        transition_end = current_dt + timedelta(minutes=transition_minutes)
        transition_time_map[(song, next_song, "normal_transition")] = fmt_range(current_dt, transition_end)
        current_dt = transition_end

    return song_time_map, transition_time_map
