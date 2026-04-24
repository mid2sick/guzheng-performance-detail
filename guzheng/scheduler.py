# -*- coding: utf-8 -*-
"""
換場任務分配演算法。

核心流程（generate_transition）：
  left   — 上一首表演者自己下箏 / 下樂器
  middle — 後台人員調整箏架、木椅、鋼琴椅
  right  — 後台人員或下一首表演者上箏 / 上樂器

人員池（PoolPicker）以 workload 平衡指派，並追蹤：
  - 每人累積工作量上限 2
  - 上台類任務（上箏 / 上架）上限 1
  - middle 與 right 的人員互斥（避免同一人兩邊跑）
"""
from __future__ import annotations
from collections import Counter, defaultdict
from typing import Dict, List, Optional, Tuple

from .config import PRIMARY_CURTAIN, BACKUP_CURTAIN, STAND_TYPES
from .models import SongAsset, SongPeople
from .utils import dedupe_keep_order, normalize_stand_full, strip_position_suffix


# =============================================================
# TaskBucket — 收集每個人的任務，最後輸出成顯示字串
# =============================================================

class TaskBucket:
    def __init__(self):
        self._tasks: Dict[str, List[str]] = defaultdict(list)

    def add(self, person: str, task: str):
        if person:
            self._tasks[person].append(task)

    def dump(self) -> str:
        if not self._tasks:
            return ""
        return "\n".join(
            f"【{person}】" + "、".join(tasks)
            for person, tasks in self._tasks.items()
        )


# =============================================================
# PoolPicker — 從後台人員池中挑人，並追蹤工作量與上台衝突
# =============================================================

class PoolPicker:
    """
    人員池管理與任務指派。

    四個池（優先序由高到低）：
      next_helpers — 下一首古箏表演者（換場期間可進中台，但快要上台）
      core         — general_core 後台人員
      mic          — mic 後台人員
      control      — control 後台人員

    兩個互斥集合：
      blocked_for_right  — 已在 middle 指派，不再進 right 區
      blocked_for_middle — 已在 right 指派，不再進 middle 區
    """

    def __init__(
        self,
        next_helpers: List[str],
        core:         List[str],
        mic:          List[str],
        control:      List[str],
    ):
        self.next_helpers = next_helpers[:]
        self.core         = core[:]
        self.mic          = mic[:]
        self.control      = control[:]

        self.workload:            Dict[str, int] = defaultdict(int)
        self.up_guzheng_count:    Dict[str, int] = defaultdict(int)
        self.up_action_count:     Dict[str, int] = defaultdict(int)
        self.middle_heavy_workers: set[str]      = set()
        self.blocked_for_right:    set[str]      = set()
        self.blocked_for_middle:   set[str]      = set()

    # ---------- 工作量追蹤 ----------

    def add_workload(self, person: Optional[str], amount: int = 1):
        if person:
            self.workload[person] += amount

    def mark_middle_heavy(self, person: Optional[str]):
        if person:
            self.middle_heavy_workers.add(person)

    def mark_middle_assigned(self, person: Optional[str]):
        """已在 middle 指派 → 不可再進 right。"""
        if person:
            self.blocked_for_right.add(person)

    def mark_right_assigned(self, person: Optional[str]):
        """已在 right 指派 → 不可再進 middle。"""
        if person:
            self.blocked_for_middle.add(person)

    def mark_up_action(self, person: Optional[str]):
        """上台類任務（上架 / 上箏 / 上木椅）每人最多 1 次。"""
        if person:
            self.up_action_count[person] += 1

    # ---------- 內部選人工具 ----------

    @staticmethod
    def _unique(items: List[str]) -> List[str]:
        return list(dict.fromkeys(items))

    def _can_take_workload(self, person: str, amount: int = 1) -> bool:
        return self.workload[person] + amount <= 2

    def _can_take_up_action(self, person: str) -> bool:
        return self.up_action_count[person] < 1

    def _least_loaded(
        self,
        candidates:          List[str],
        required_workload:   int  = 1,
        require_up_action:   bool = False,
    ) -> Optional[str]:
        eligible = [
            p for p in self._unique(candidates)
            if p
            and self._can_take_workload(p, required_workload)
            and (not require_up_action or self._can_take_up_action(p))
        ]
        if not eligible:
            return None
        min_load = min(self.workload[p] for p in eligible)
        return next(p for p in eligible if self.workload[p] == min_load)

    def _pick_by_priority(
        self,
        pools:               List[List[str]],
        required_workload:   int  = 1,
        require_up_action:   bool = False,
    ) -> Optional[str]:
        """依池的優先序選工作量最少且符合條件的人。"""
        for pool in pools:
            person = self._least_loaded(
                self._unique([p for p in pool if p]),
                required_workload=required_workload,
                require_up_action=require_up_action,
            )
            if person:
                return person
        return None

    def _pick_from_pool_capped(
        self,
        pool:              List[str],
        cap:               int,
        exclude:           set = frozenset(),
        require_up_action: bool = False,
    ) -> Optional[str]:
        """從 pool 中選人，限制同一人最多搬 cap 台箏。"""
        candidates = [
            p for p in self._unique(pool)
            if p not in exclude
            and self.up_guzheng_count[p] < cap
            and self._can_take_workload(p, 1)
            and (not require_up_action or self._can_take_up_action(p))
        ]
        if not candidates:
            return None
        person = self._least_loaded(candidates, require_up_action=require_up_action)
        if person:
            self.up_guzheng_count[person] += 1
        return person

    # ---------- middle 任務選人 ----------

    def _middle_pools(self) -> List[List[str]]:
        """回傳「可進 middle」的四個池（已排除 blocked_for_middle 的人）。"""
        excl = self.blocked_for_middle
        return [
            [p for p in self.next_helpers if p not in excl],
            [p for p in self.core         if p not in excl],
            [p for p in self.mic          if p not in excl],
            [p for p in self.control      if p not in excl],
        ]

    def pick_for_stand_adjust(self, required_workload: int = 1) -> Optional[str]:
        """調整箏架：next_helpers → core → mic → control。"""
        return self._pick_by_priority(self._middle_pools(), required_workload=required_workload)

    def pick_for_stand_down(self, required_workload: int = 1) -> Optional[str]:
        """下架 / 下木椅 / 下鋼琴椅：只用 core → mic → control。"""
        excl = self.blocked_for_middle
        return self._pick_by_priority(
            [
                [p for p in self.core    if p not in excl],
                [p for p in self.mic     if p not in excl],
                [p for p in self.control if p not in excl],
            ],
            required_workload=required_workload,
        )

    def pick_for_chair_adjust(self, required_workload: int = 1) -> Optional[str]:
        """調整木椅：next_helpers → core → control（不含 mic）。"""
        excl = self.blocked_for_middle
        return self._pick_by_priority(
            [
                [p for p in self.next_helpers if p not in excl],
                [p for p in self.core         if p not in excl],
                [p for p in self.control      if p not in excl],
            ],
            required_workload=required_workload,
        )

    def pick_for_middle_up(self, required_workload: int = 1) -> Optional[str]:
        """上架 / 上木椅 / 上鋼琴椅（上台類）：需要 up_action 名額。"""
        return self._pick_by_priority(self._middle_pools(), required_workload=required_workload, require_up_action=True)

    def pick_specific_for_up_stand(
        self,
        stand_type:            str,
        next_player_stand_map: Dict[str, str],
    ) -> Optional[str]:
        """
        上架優先：若下一首表演者中有人剛好用這種架子，讓他自己來上。
        """
        excl = self.blocked_for_middle
        candidates = [
            p for p in self.next_helpers
            if p not in excl
            and next_player_stand_map.get(p) == stand_type
            and self._can_take_workload(p, 1)
            and self._can_take_up_action(p)
        ]
        return self._least_loaded(candidates, require_up_action=True)

    # ---------- right 任務選人 ----------

    def pick_for_up_guzheng(self) -> Optional[str]:
        """
        上箏規則：
          第一輪（每人最多 1 台）：control → core → next_helpers → mic
          第二輪（每人最多 2 台，core 不參與）：control → next_helpers → mic
          每輪都先盡量避開 middle_heavy_workers，再逐步放寬排除條件。
        """
        pools_r1 = [self.control, self.core, self.next_helpers, self.mic]
        pools_r2 = [self.control, self.next_helpers, self.mic]

        for pools, cap in [(pools_r1, 1), (pools_r2, 2)]:
            for exclude in [
                self.middle_heavy_workers | self.blocked_for_right,
                self.blocked_for_right,
                set(),
            ]:
                for pool in pools:
                    person = self._pick_from_pool_capped(pool, cap, exclude=exclude, require_up_action=True)
                    if person:
                        return person

        return None


# =============================================================
# 輔助函式
# =============================================================

def stand_counter(stands: List[str]) -> Counter:
    out: Counter = Counter()
    for s in stands:
        key = normalize_stand_full(s)
        if key:
            out[key] += 1
    return out


def build_next_player_stand_map(next_people: Optional[SongPeople]) -> Dict[str, str]:
    """建立「下一首古箏表演者 → 他自己要用的箏架類型」對照表。"""
    if not next_people:
        return {}
    return {
        name: normalize_stand_full(stand)
        for name, _, stand in next_people.guzheng_players
        if name and stand
    }


def build_support_lists(
    prev_people:     SongPeople,
    next_people:     SongPeople,
    backstage_roles: Dict[str, str],
) -> Tuple[List[str], List[str], List[str], List[str], List[str]]:
    """
    從後台人員中篩出可用的人員池。
    凡在前一首或後一首表演者名單中的人，一律排除（他們正在台上或準備上台）。
    回傳 (next_helpers, core, mic, control, curtain)
    """
    on_stage = prev_people.all_names() | next_people.all_names()

    next_helpers = [name for name, _, _ in next_people.guzheng_players]
    next_helpers = dedupe_keep_order(next_helpers)

    core, mic, control, curtain, tuning = [], [], [], [], []

    for person, role in backstage_roles.items():
        if person in on_stage:
            continue
        if   role == "general_core": core.append(person)
        elif role == "mic":          mic.append(person)
        elif role == "control":      control.append(person)
        elif role == "curtain":      curtain.append(person)
        elif role == "tuning":       tuning.append(person)

    return next_helpers, core, mic, control, curtain


def choose_curtain_person(
    prev_names:      set[str],
    next_names:      set[str],
    primary_curtain: str,
    backup_curtain:  str,
) -> Optional[str]:
    """
    選布幕操作人員：
      1. 主要布幕人員不在台上 → 選他
      2. 主要在台上 → 改派支援人員
      3. 兩人都在台上 → 仍回傳支援者作保底
    """
    on_stage = prev_names | next_names
    if primary_curtain and primary_curtain not in on_stage:
        return primary_curtain
    if backup_curtain and backup_curtain not in on_stage:
        return backup_curtain
    return backup_curtain or primary_curtain


# =============================================================
# generate_transition — 主要換場任務分配函式
# =============================================================

def generate_transition(
    prev_song:             Optional[str],
    next_song:             Optional[str],
    prev_assets:           Optional[SongAsset],
    next_assets:           Optional[SongAsset],
    prev_people:           Optional[SongPeople],
    next_people:           Optional[SongPeople],
    backstage_roles:       Dict[str, str],
    is_pre_row:            bool = False,
    is_before_intermission: bool = False,
    is_last_song_teardown: bool = False,
) -> Tuple[str, str, str]:
    """
    計算換場列的三欄內容。

    left   — 表演者自己下箏 / 下樂器
    middle — 後台人員調整箏架、木椅、鋼琴椅、操作布幕
    right  — 後台人員上箏、特殊樂器上台

    特殊情況：
      is_pre_row            → 前置列，全空白
      is_before_intermission → 中場前，只有 left（下箏），不換場
      is_last_song_teardown  → 最後一首拆台，布幕改為「敬禮時按降幕」
    """
    if is_pre_row:
        return "", "", ""
    if not prev_assets or not prev_people:
        return "", "", ""

    left   = TaskBucket()
    middle = TaskBucket()
    right  = TaskBucket()

    # ── left：上一首表演者各自下自己的樂器 ──────────────────
    for name, guzheng, _ in prev_people.guzheng_players:
        if guzheng:
            left.add(name, f"下{guzheng}")
    for name in prev_people.piano_players:
        left.add(name, "下電鋼琴")
    for name in prev_people.percussion_players:
        left.add(name, "下打擊樂器")
    for name in prev_people.bass_players:
        left.add(name, "下低音提琴")

    if is_before_intermission:
        return left.dump(), "", ""

    # ── 建立人員池 ────────────────────────────────────────────
    if next_people is None:
        next_people = SongPeople([], [], [], [])

    next_helpers, core, mic, control, _ = build_support_lists(prev_people, next_people, backstage_roles)
    picker = PoolPicker(next_helpers, core, mic, control)

    next_player_stand_map = build_next_player_stand_map(next_people)

    # ── 布幕 ──────────────────────────────────────────────────
    curtain_person = choose_curtain_person(
        prev_names=prev_people.all_names(),
        next_names=next_people.all_names(),
        primary_curtain=PRIMARY_CURTAIN,
        backup_curtain=BACKUP_CURTAIN,
    )

    curtain_pending = bool(curtain_person)   # 待用於第一個調整任務
    if curtain_person:
        curtain_task = "敬禮時按降幕" if is_last_song_teardown else "敬禮按降幕，東西就位按升幕"
        middle.add(curtain_person, curtain_task)

    # ── middle：箏架 ──────────────────────────────────────────
    prev_stands = stand_counter(prev_assets.stands)
    next_stands = stand_counter(next_assets.stands if next_assets else [])

    for stand_type in STAND_TYPES:
        A = prev_stands.get(stand_type, 0)
        B = next_stands.get(stand_type, 0)
        if A == 0 and B == 0:
            continue

        # 下架
        for _ in range(max(A - B, 0)):
            person = picker.pick_for_stand_down()
            if person:
                middle.add(person, f"下{stand_type}*1")
                picker.add_workload(person, 1)
                picker.mark_middle_heavy(person)
                picker.mark_middle_assigned(person)

        # 調整架（留在原位但需重新定位）
        remain = min(A, B)
        while remain > 0:
            chunk = min(2, remain)
            if curtain_pending and curtain_person:
                person = curtain_person
                curtain_pending = False
            else:
                person = picker.pick_for_stand_adjust(required_workload=chunk)
            if person:
                middle.add(person, f"調整{stand_type}*{chunk}")
                picker.add_workload(person, chunk)
                picker.mark_middle_heavy(person)
                picker.mark_middle_assigned(person)
            remain -= chunk

        # 上架（上台類）
        for _ in range(max(B - A, 0)):
            person = picker.pick_specific_for_up_stand(stand_type, next_player_stand_map)
            if person is None:
                person = picker.pick_for_middle_up()
            if person:
                middle.add(person, f"上{stand_type}*1")
                picker.add_workload(person, 1)
                picker.mark_middle_heavy(person)
                picker.mark_middle_assigned(person)
                picker.mark_up_action(person)

    # ── middle：木椅 ─────────────────────────────────────────
    prev_chairs = prev_assets.wood_chair_count
    next_chairs = next_assets.wood_chair_count if next_assets else 0

    if prev_chairs > next_chairs:
        remain = prev_chairs - next_chairs
        while remain > 0:
            chunk = min(2, remain)
            person = picker.pick_for_stand_down(required_workload=chunk)
            if person:
                middle.add(person, f"下木椅*{chunk}")
                picker.add_workload(person, chunk)
                picker.mark_middle_heavy(person)
                picker.mark_middle_assigned(person)
            remain -= chunk

    elif next_chairs > prev_chairs:
        remain = next_chairs - prev_chairs
        while remain > 0:
            chunk = min(2, remain)
            person = picker.pick_for_middle_up(required_workload=chunk)
            if person:
                middle.add(person, f"上木椅*{chunk}")
                picker.add_workload(person, chunk)
                picker.mark_middle_heavy(person)
                picker.mark_middle_assigned(person)
                picker.mark_up_action(person)
            remain -= chunk

    remain = min(prev_chairs, next_chairs)
    while remain > 0:
        chunk = min(2, remain)
        if curtain_pending and curtain_person:
            person = curtain_person
            curtain_pending = False
        else:
            person = picker.pick_for_chair_adjust(required_workload=chunk)
        if person:
            middle.add(person, f"調整木椅*{chunk}")
            picker.add_workload(person, chunk)
            picker.mark_middle_heavy(person)
            picker.mark_middle_assigned(person)
        remain -= chunk

    # ── middle：鋼琴椅 ────────────────────────────────────────
    prev_has_piano = bool(prev_people.piano_players)
    next_has_piano = bool(next_people.piano_players)

    if next_has_piano and not prev_has_piano:
        person = picker.pick_for_middle_up()
        if person:
            middle.add(person, "上鋼琴椅*1")
            picker.add_workload(person, 1)
            picker.mark_middle_heavy(person)
            picker.mark_middle_assigned(person)
            picker.mark_up_action(person)

    if prev_has_piano and not next_has_piano:
        person = picker.pick_for_stand_down()
        if person:
            middle.add(person, "下鋼琴椅*1")
            picker.add_workload(person, 1)
            picker.mark_middle_heavy(person)
            picker.mark_middle_assigned(person)

    # ── right：特殊樂器上台 ───────────────────────────────────
    for name in next_people.percussion_players:
        right.add(name, "上打擊樂器")
        picker.mark_right_assigned(name)
        picker.mark_up_action(name)

    for name in next_people.bass_players:
        right.add(name, "上低音提琴")
        picker.mark_right_assigned(name)
        picker.mark_up_action(name)

    for name in next_people.piano_players:
        right.add(name, "上電鋼琴")
        picker.mark_right_assigned(name)
        picker.mark_up_action(name)

    # ── right：上箏 ───────────────────────────────────────────
    for _, guzheng, _ in next_people.guzheng_players:
        instrument = strip_position_suffix(guzheng)
        person = picker.pick_for_up_guzheng()
        if person:
            right.add(person, f"上{instrument}")
            picker.add_workload(person, 1)
            picker.mark_right_assigned(person)
            picker.mark_up_action(person)

    return left.dump(), middle.dump(), right.dump()
