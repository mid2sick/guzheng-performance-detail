# -*- coding: utf-8 -*-
"""
核心資料結構。
所有其他模組都依賴這裡，這裡不依賴任何其他模組。
"""
from __future__ import annotations
from dataclasses import dataclass
from typing import List, Tuple


@dataclass
class SongAsset:
    """一首曲目的舞台物件資訊（古箏、箏架、木椅）。"""
    song: str
    guzheng_names: List[str]    # 該曲使用的古箏編號清單
    stands: List[str]           # 對應的箏架類型清單
    guzheng_count: int
    wood_chair_count: int


@dataclass
class SongPeople:
    """一首曲目的所有表演者，依樂器分類。"""
    guzheng_players: List[Tuple[str, str, str]]  # (姓名, 古箏編號, 箏架類型)
    percussion_players: List[str]
    piano_players: List[str]
    bass_players: List[str]

    def all_names(self) -> set[str]:
        """回傳此曲所有表演者姓名集合（用於判斷誰正在上台）。"""
        names = {name for name, _, _ in self.guzheng_players}
        names |= set(self.percussion_players)
        names |= set(self.piano_players)
        names |= set(self.bass_players)
        return names
