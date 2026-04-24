# -*- coding: utf-8 -*-
"""
資料組合層。
接收 readers 回傳的原始資料，合併或重建成完整的 domain objects。
"""
from __future__ import annotations
from typing import Dict, List, Tuple

from .models import SongAsset, SongPeople
from .utils import clean_text, normalize_stand_full


def build_song_people(
    detail_people: Dict[str, List[Tuple[str, str, str]]],
    role_people:   Dict[str, Dict[str, List[str]]],
) -> Dict[str, SongPeople]:
    """
    合併明細（古箏表演者）與演出分配（其他聲部）成 SongPeople。
    """
    songs = set(detail_people.keys()) | set(role_people.keys())
    result: Dict[str, SongPeople] = {}

    for song in songs:
        guzheng_players = detail_people.get(song, [])
        role_map        = role_people.get(song, {})
        result[song] = SongPeople(
            guzheng_players=guzheng_players,
            percussion_players=role_map.get("percussion", []),
            piano_players=role_map.get("piano", []),
            bass_players=role_map.get("bass", []),
        )

    return result


def rebuild_song_assets_from_detail(
    song_order:             List[str],
    song_assets_from_sheet: Dict[str, SongAsset],
    detail_people:          Dict[str, List[Tuple[str, str, str]]],
) -> Dict[str, SongAsset]:
    """
    優先用「明細」工作表的古箏/箏架資料重建 SongAsset，
    若某首在明細中沒有資料，才退回「曲目」工作表的原始資料。
    明細是箏架分配的最終依據。
    """
    rebuilt: Dict[str, SongAsset] = {}

    for song in song_order:
        rows = detail_people.get(song, [])

        if rows:
            guzheng_names: List[str] = []
            stands:        List[str] = []

            for _, guzheng, stand in rows:
                if g := clean_text(guzheng):
                    guzheng_names.append(g)
                if s := normalize_stand_full(clean_text(stand)):
                    stands.append(s)

            guzheng_count = len(guzheng_names) or len(stands)

            rebuilt[song] = SongAsset(
                song=song,
                guzheng_names=guzheng_names,
                stands=stands,
                guzheng_count=guzheng_count,
                wood_chair_count=guzheng_count,
            )
        else:
            rebuilt[song] = song_assets_from_sheet[song]

    return rebuilt
