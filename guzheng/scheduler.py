# -*- coding: utf-8 -*-
"""
換場任務分配演算法。

核心流程（generate_transition）：
  right  — 後台人員先搶佔上箏 / 上樂器（G 欄，優先分配）
  left   — 上一首表演者自己下箏 / 下樂器（E 欄，自助）
  middle — 下一場表演者自己上/調整箏架與木椅；後台人員下架、下木椅、鋼琴椅（F 欄）

人員池（PoolPicker）以 workload 平衡指派，並追蹤：
  - 每人累積工作量上限 2
  - 上台類任務（上箏）上限 1
  - G 與 F 的人員互斥（已分 G 的人不進 F，反之亦然）
  - 已做「下」類任務的人不再被指派上鋼琴椅
  - 下架一次最多 2 組（workload 計 1）
  - 下椅子一次全下（workload 計 1）
  - F 人手不足時輸出【空缺】
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
        self._vacancies: List[str] = []

    def add(self, person: str, task: str):
        if person:
            self._tasks[person].append(task)

    def add_vacancy(self, task: str):
        self._vacancies.append(task)

    def people(self) -> set:
        return set(self._tasks.keys())

    def dump(self) -> str:
        parts = [
            f"【{person}】" + "、".join(tasks)
            for person, tasks in self._tasks.items()
        ]
        parts += [f"【空缺】{t}" for t in self._vacancies]
        return "\n".join(parts)


# =============================================================
# PoolPicker — 從後台人員池中挑人，並追蹤工作量與上台衝突
# =============================================================

class PoolPicker:
    """
    人員池管理與任務指派。

    四個池（優先序由高到低）：
      next_helpers — 下一首古箏表演者（換場期間負責上和調整自己的箏架、椅子，準備上台）
      core         — general_core 後台人員
      mic          — mic 後台人員
      control      — control 後台人員

    兩個互斥集合：
      blocked_for_right  — 已在 F 負責調整箏架木椅，不再進 G 上箏
      blocked_for_middle — 已在 G 負責上箏，不再進 F 調整箏架木椅

    down_workers：已執行下架 / 下椅子 / 下鋼琴椅，不可再執行另一次下架 / 下椅子，也不可接上鋼琴椅。
    curtain_workers：布幕人員，排除於一般 F 選人，完成布幕任務後不再接受其他任務。
    """

    def __init__(
        self,
        # TODO: next_helpers 理論上只負責上/調整自己的箏架、椅子，可能不需要放入 PoolPicker 作為其他任務的候選人
        next_helpers: List[str],
        core:         List[str],
        mic:          List[str],
        control:      List[str],
    ):
        self.next_helpers = next_helpers[:]
        self.core         = core[:]
        self.mic          = mic[:]
        self.control      = control[:]

        self.workload:             Dict[str, int] = defaultdict(int)
        self.up_guzheng_count:     Dict[str, int] = defaultdict(int)
        self.up_action_count:      Dict[str, int] = defaultdict(int)
        self.blocked_for_right:    set[str]       = set()
        self.blocked_for_middle:   set[str]       = set()
        self.down_workers:         set[str]       = set()
        self.curtain_workers:      set[str]       = set()

    # ---------- 工作量追蹤 ----------

    def add_workload(self, person: Optional[str], amount: int = 1):
        if person:
            self.workload[person] += amount

    def mark_middle_assigned(self, person: Optional[str]):
        """已在 F 指派 → 不可再進 G。"""
        if person:
            self.blocked_for_right.add(person)

    def mark_right_assigned(self, person: Optional[str]):
        """已在 G 指派 → 不可再進 F。"""
        if person:
            self.blocked_for_middle.add(person)

    def mark_up_action(self, person: Optional[str]):
        """上台類任務（上架 / 上箏 / 上木椅）每人最多 1 次。"""
        if person:
            self.up_action_count[person] += 1

    def mark_down_action(self, person: Optional[str]):
        """下架 / 下椅子 / 下鋼琴椅 → 不可再執行另一次下架 / 下椅子，也不可接上鋼琴椅。"""
        if person:
            self.down_workers.add(person)

    def mark_curtain_assigned(self, person: Optional[str]):
        """布幕人員 → 排除於一般 F 選人，只能透過 curtain_pending 機制取得一個附加任務。"""
        if person:
            self.curtain_workers.add(person)

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
        candidates:        List[str],
        required_workload: int  = 1,
        require_up_action: bool = False,
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
        pools:             List[List[str]],
        required_workload: int  = 1,
        require_up_action: bool = False,
    ) -> Optional[str]:
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

    # ---------- F（middle）任務選人 ----------

    def pick_for_stand_down(self, required_workload: int = 1) -> Optional[str]:
        """下架 / 下木椅 / 下鋼琴椅：只用 core → mic → control；已執行過下類任務的人不再選。"""
        excl = self.blocked_for_middle | self.down_workers | self.curtain_workers
        return self._pick_by_priority(
            [
                [p for p in self.core    if p not in excl],
                [p for p in self.mic     if p not in excl],
                [p for p in self.control if p not in excl],
            ],
            required_workload=required_workload,
        )

    def pick_for_piano_up(self) -> Optional[str]:
        """上鋼琴椅：排除已做過下架 / 下椅子的人，也排除布幕人員。"""
        excl = self.blocked_for_middle | self.down_workers | self.curtain_workers
        return self._pick_by_priority(
            [
                [p for p in self.next_helpers if p not in excl],
                [p for p in self.core         if p not in excl],
                [p for p in self.mic          if p not in excl],
                [p for p in self.control      if p not in excl],
            ],
            required_workload=1,
            require_up_action=True,
        )

    # ---------- G（right）任務選人 ----------

    def pick_for_up_guzheng(self) -> Optional[str]:
        """control → core → next_helpers → mic，blocked_for_right 為硬限制。"""
        for pool in [self.control, self.core, self.next_helpers, self.mic]:
            person = self._pick_from_pool_capped(
                pool, 1, exclude=self.blocked_for_right, require_up_action=True
            )
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


def build_support_lists(
    prev_people:     SongPeople,
    next_people:     SongPeople,
    backstage_roles: Dict[str, str],
) -> Tuple[List[str], List[str], List[str], List[str], List[str]]:
    on_stage = prev_people.all_names() | next_people.all_names()

    next_helpers = dedupe_keep_order(
        [name for name, _, _ in next_people.guzheng_players]
    )

    core, mic, control, curtain = [], [], [], []

    for person, role in backstage_roles.items():
        if person in on_stage:
            continue
        if   role == "general_core": core.append(person)
        elif role == "mic":          mic.append(person)
        elif role == "control":      control.append(person)
        elif role == "curtain":      curtain.append(person)

    return next_helpers, core, mic, control, curtain


def choose_curtain_person(
    prev_names:      set[str],
    next_names:      set[str],
    primary_curtain: str,
    backup_curtain:  str,
) -> Optional[str]:
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
    prev_song:              Optional[str],
    next_song:              Optional[str],
    prev_assets:            Optional[SongAsset],
    next_assets:            Optional[SongAsset],
    prev_people:            Optional[SongPeople],
    next_people:            Optional[SongPeople],
    backstage_roles:        Dict[str, str],
    is_pre_row:             bool = False,
    is_before_intermission: bool = False,
    is_last_song_teardown:  bool = False,
    debug:                  bool = False,
) -> Tuple[str, str, str, str]:
    """
    計算換場列的四欄內容。

    E（left）   — 表演者自己下箏 / 下樂器
    F（middle） — 下一場表演者優先自己上/調整箏架、木椅；後台人員處理剩餘工作
    G（right）  — 後台人員上箏、特殊樂器上台
    I（idle）   — 該時間段無任務的後台人員

    分配順序：
      1. left — 上一場表演者各自下樂器
      2. middle（優先）— 下一場表演者自己上或調整自己的箏架、木椅
      3. right — 後台人員上箏；特殊樂器表演者自帶
      4. middle（剩餘）— 下箏架、下木椅、下鋼琴椅（後台）
    F 人手不足時輸出【空缺】。
    """
    if is_pre_row:
        return "", "", "", ""
    if not prev_assets or not prev_people:
        return "", "", "", ""

    _dbg: any = (lambda msg: print(f"[debug] {msg}")) if debug else (lambda _: None)

    if debug:
        p = prev_song or "（開場）"
        n = next_song or "（結束）"
        print(f"\n[debug] ══ 換場 {p} → {n} ══")

    left   = TaskBucket()
    middle = TaskBucket()
    right  = TaskBucket()

    # ── E：上一首表演者各自下自己的樂器 ─────────────────────────
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
        on_stage = prev_people.all_names() | (next_people.all_names() if next_people else set())
        _IDLE_EXCLUDE_ROLES = {"tuning", "general_other"}
        idle_str = "\n".join(sorted(
            p for p, r in backstage_roles.items()
            if p not in on_stage and r not in _IDLE_EXCLUDE_ROLES
        ))
        return left.dump(), "", "", idle_str

    # ── 建立人員池 ───────────────────────────────────────────────
    if next_people is None:
        next_people = SongPeople([], [], [], [])

    next_helpers, core, mic, control, _ = build_support_lists(
        prev_people, next_people, backstage_roles
    )
    picker = PoolPicker(next_helpers, core, mic, control)

    if debug:
        print(f"[debug] 人員池 | next_helpers={next_helpers} core={core} mic={mic} control={control}")

    # ── 布幕（F 欄第一筆，獨立處理）────────────────────────────
    curtain_person = choose_curtain_person(
        prev_names=prev_people.all_names(),
        next_names=next_people.all_names(),
        primary_curtain=PRIMARY_CURTAIN,
        backup_curtain=BACKUP_CURTAIN,
    )
    if curtain_person:
        curtain_task = (
            "敬禮時按降幕"
            if is_last_song_teardown
            else "敬禮按降幕，東西就位按升幕"
        )
        middle.add(curtain_person, curtain_task)
        picker.mark_curtain_assigned(curtain_person)  # 排除於一般 F 選人
        picker.mark_middle_assigned(curtain_person)   # 布幕人員不可再進 G
        picker.add_workload(curtain_person, 1)
        _dbg(f"[布幕] {curtain_person} ← {curtain_task}")

    # ── 計算箏架與木椅的調整/上/下需求 ──────────────────────────
    prev_stands = stand_counter(prev_assets.stands)
    next_stands = stand_counter(next_assets.stands if next_assets else [])
    prev_chairs = prev_assets.wood_chair_count
    next_chairs = next_assets.wood_chair_count if next_assets else 0

    all_stand_types = set(list(prev_stands.keys()) + list(next_stands.keys()))
    # 各類型箏架的剩餘調整/上/下數量（表演者優先步驟後會更新）
    rem_adj  = {t: min(prev_stands.get(t, 0), next_stands.get(t, 0)) for t in all_stand_types}
    rem_up   = {t: max(next_stands.get(t, 0) - prev_stands.get(t, 0), 0) for t in all_stand_types}
    rem_down = {t: max(prev_stands.get(t, 0) - next_stands.get(t, 0), 0) for t in all_stand_types}

    rem_adj_chairs  = min(prev_chairs, next_chairs)
    rem_up_chairs   = max(next_chairs - prev_chairs, 0)
    rem_down_chairs = max(prev_chairs - next_chairs, 0)

    # ════════════════════════════════════════════════════════════
    # F（middle）優先 — 下一場表演者自己上或調整自己的箏架、木椅
    # ════════════════════════════════════════════════════════════
    for name, guzheng, stand in next_people.guzheng_players:
        if not name:
            continue
        stand_type = normalize_stand_full(stand) if stand else ""

        # 箏架任務：有剩餘調整名額優先調整，否則上架
        if stand_type and rem_adj.get(stand_type, 0) > 0:
            middle.add(name, f"調整{stand_type}*1")
            rem_adj[stand_type] -= 1
            picker.mark_middle_assigned(name)
            picker.add_workload(name, 1)
            _dbg(f"[F-表演者] 調整{stand_type}*1 → {name}")
        elif stand_type and rem_up.get(stand_type, 0) > 0:
            middle.add(name, f"上{stand_type}*1")
            rem_up[stand_type] -= 1
            picker.mark_middle_assigned(name)
            picker.mark_up_action(name)
            picker.add_workload(name, 1)
            _dbg(f"[F-表演者] 上{stand_type}*1 → {name}")

        # 木椅任務：有剩餘調整名額優先調整，否則上椅
        if rem_adj_chairs > 0:
            middle.add(name, "調整木椅*1")
            rem_adj_chairs -= 1
            picker.add_workload(name, 1)
            _dbg(f"[F-表演者] 調整木椅*1 → {name}")
        elif rem_up_chairs > 0:
            middle.add(name, "上木椅*1")
            rem_up_chairs -= 1
            picker.mark_up_action(name)
            picker.add_workload(name, 1)
            _dbg(f"[F-表演者] 上木椅*1 → {name}")

    # ════════════════════════════════════════════════════════════
    # G（right）— 特殊樂器表演者自帶 + 後台人員上箏
    # ════════════════════════════════════════════════════════════

    for name in next_people.percussion_players:
        right.add(name, "上打擊樂器")
        picker.mark_right_assigned(name)
        picker.mark_up_action(name)
        _dbg(f"[G] 上打擊樂器 → {name}（表演者自帶）")

    for name in next_people.bass_players:
        right.add(name, "上低音提琴")
        picker.mark_right_assigned(name)
        picker.mark_up_action(name)
        _dbg(f"[G] 上低音提琴 → {name}（表演者自帶）")

    for name in next_people.piano_players:
        right.add(name, "上電鋼琴")
        picker.mark_right_assigned(name)
        picker.mark_up_action(name)
        _dbg(f"[G] 上電鋼琴 → {name}（表演者自帶）")

    for _, guzheng, _ in next_people.guzheng_players:
        instrument = strip_position_suffix(guzheng)
        person = picker.pick_for_up_guzheng()
        if person:
            right.add(person, f"上{instrument}")
            picker.add_workload(person, 1)
            picker.mark_right_assigned(person)
            picker.mark_up_action(person)
            _dbg(f"[G] 上{instrument} → {person} (workload:{picker.workload[person]})")
        else:
            _dbg(f"[G] 上{instrument} → （無可用人員，跳過）")

    # ════════════════════════════════════════════════════════════
    # F（middle）剩餘 — 箏架下、木椅下、鋼琴椅
    # ════════════════════════════════════════════════════════════

    for stand_type in STAND_TYPES:
        # 下架：一次最多 2 組，workload 計 1
        need_down = rem_down.get(stand_type, 0)
        while need_down > 0:
            chunk = min(2, need_down)
            person = picker.pick_for_stand_down()
            if person:
                middle.add(person, f"下{stand_type}*{chunk}")
                picker.add_workload(person, 1)
                picker.mark_middle_assigned(person)
                picker.mark_down_action(person)
                _dbg(f"[F] 下{stand_type}*{chunk} → {person} (workload:{picker.workload[person]})")
            else:
                middle.add_vacancy(f"下{stand_type}*{chunk}")
                _dbg(f"[F] 下{stand_type}*{chunk} → 【空缺】")
            need_down -= chunk

    # ── 木椅（剩餘）────────────────────────────────────────────

    # 下椅子：每人一次最多下 3 張，workload 計 1
    while rem_down_chairs > 0:
        chunk = min(3, rem_down_chairs)
        person = picker.pick_for_stand_down()
        if person:
            middle.add(person, f"下木椅*{chunk}")
            picker.add_workload(person, 1)
            picker.mark_middle_assigned(person)
            picker.mark_down_action(person)
            _dbg(f"[F] 下木椅*{chunk} → {person} (workload:{picker.workload[person]})")
        else:
            middle.add_vacancy(f"下木椅*{chunk}")
            _dbg(f"[F] 下木椅*{chunk} → 【空缺】")
        rem_down_chairs -= chunk

    # ── 鋼琴椅 ───────────────────────────────────────────────────
    prev_has_piano = bool(prev_people.piano_players)
    next_has_piano = bool(next_people.piano_players)

    if next_has_piano and not prev_has_piano:
        # 上鋼琴椅：排除下架 / 下椅子人員（rule 5）
        person = picker.pick_for_piano_up()
        if person:
            middle.add(person, "上鋼琴椅*1")
            picker.add_workload(person, 1)
            picker.mark_middle_assigned(person)
            picker.mark_up_action(person)
            _dbg(f"[F] 上鋼琴椅*1 → {person} (workload:{picker.workload[person]})")
        else:
            middle.add_vacancy("上鋼琴椅*1")
            _dbg(f"[F] 上鋼琴椅*1 → 【空缺】")

    if prev_has_piano and not next_has_piano:
        person = picker.pick_for_stand_down()
        if person:
            middle.add(person, "下鋼琴椅*1")
            picker.add_workload(person, 1)
            picker.mark_middle_assigned(person)
            picker.mark_down_action(person)
            _dbg(f"[F] 下鋼琴椅*1 → {person} (workload:{picker.workload[person]})")
        else:
            middle.add_vacancy("下鋼琴椅*1")
            _dbg(f"[F] 下鋼琴椅*1 → 【空缺】")

    on_stage = prev_people.all_names() | next_people.all_names()
    busy = middle.people() | right.people()
    _IDLE_EXCLUDE_ROLES = {"tuning", "general_other"}
    idle_str = "\n".join(sorted(
        p for p, r in backstage_roles.items()
        if p not in on_stage and p not in busy and r not in _IDLE_EXCLUDE_ROLES
    ))
    return left.dump(), middle.dump(), right.dump(), idle_str
