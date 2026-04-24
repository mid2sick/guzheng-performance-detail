# 古箏公演細流產生器

自動從 Excel 資料整合檔讀取曲目、表演者、後台人員資訊，計算換場任務分配與時間軸，輸出完整的公演細流 Excel 檔。

## 執行方式

```bash
python generate_guzheng_flow.py [--source 古箏社公演_資料整合檔_0317.xlsx] [--output 細流第XX版.xlsx] [--debug]
```

所有可調整參數（檔案路徑、時間設定、布幕人員等）集中在 `guzheng/config.py`。

---

## 專案結構

```
guzheng-performance-detail/
├── generate_guzheng_flow.py   # 執行入口
└── guzheng/                   # 核心套件
    ├── config.py
    ├── models.py
    ├── utils.py
    ├── readers.py
    ├── builders.py
    ├── timeline.py
    ├── scheduler.py
    ├── styles.py
    └── writer.py
```

---

## 模組說明

### `generate_guzheng_flow.py` — 執行入口

串接所有步驟的主程式。依序呼叫讀檔、建立資料結構、計算時間軸、分配任務、輸出 Excel，本身不包含任何業務邏輯。修改執行流程時從這裡開始看。

---

### `guzheng/config.py` — 使用者設定區

所有可調整的常數集中於此，是唯一需要在換場前修改的檔案：

| 設定項目 | 說明 |
|---|---|
| `SOURCE_XLSX` / `OUTPUT_XLSX` | 來源與輸出檔案路徑 |
| `INTERMISSION_AFTER` | 中場休息插在第幾首之後 |
| `PRIMARY_CURTAIN` / `BACKUP_CURTAIN` | 主要與備用布幕操作人員 |
| `START_TIME` | 演出開始時間 |
| `OPENING_MINUTES` | 開場前置時間（分鐘） |
| `DEFAULT_TRANSITION_MINUTES` | 一般換場時間（分鐘） |
| `INTERMISSION_MINUTES` | 中場休息時間（分鐘） |
| `SONG_COLOR_MAP` | 各曲目的古箏定位顏色標記 |
| `STAND_TYPES` | 箏架類型清單（決定換場時的處理順序） |

**依賴：** 無。

---

### `guzheng/models.py` — 資料結構

定義兩個核心 dataclass，是整個專案的資料骨架：

#### `SongAsset`
記錄一首曲目的**舞台物件資訊**：使用哪些古箏、哪些箏架（含數量）、需要幾張木椅。換場時用來計算要上/下/調整哪些東西。

#### `SongPeople`
記錄一首曲目的**所有表演者**，依樂器類型分群：
- `guzheng_players`：`(姓名, 古箏編號, 箏架類型)` tuple 清單
- `percussion_players` / `piano_players` / `bass_players`：其他聲部姓名清單
- `all_names()`：回傳此曲全部表演者姓名集合，用於判斷誰在換場期間正在台上

**依賴：** 無。

---

### `guzheng/utils.py` — 文字工具函式

純函式，無副作用，專門處理來自 Excel 的髒資料：

| 函式 | 說明 |
|---|---|
| `clean_text` | 去除 `None` 與前後空白 |
| `split_lines` | 儲存格多行文字 → 字串清單 |
| `split_names` | 支援逗號 / 頓號 / 換行的姓名字串 → 清單 |
| `normalize_raw_name` | 統一姓名格式（含英文取英文段；中文三字以上去姓） |
| `normalize_stand_full` | 統一箏架名稱寫法（例：`黑Ａ架` → `黑A架`） |
| `stand_for_performer` | 完整箏架名稱 → 表演者欄簡短顯示（例：`黑架` → `黑`） |
| `strip_position_suffix` | 移除古箏名稱的位置括號（例：`古箏一號（左）` → `古箏一號`） |
| `dedupe_keep_order` | 去除重複元素，保留原始順序 |

**依賴：** 無。

---

### `guzheng/readers.py` — Excel 讀取層

每個函式對應來源 xlsx 的一個工作表，將原始儲存格資料轉為 Python 物件。讀取邏輯與後續處理分開，方便替換資料來源格式。

| 函式 | 工作表 | 回傳 |
|---|---|---|
| `read_song_order_and_assets` | 曲目 | 曲目順序清單 + `SongAsset` 字典 |
| `read_song_duration_minutes` | 曲目 | `{曲目: 分鐘數}` |
| `read_detail_people` | 明細 | `{曲目: [(姓名, 古箏, 箏架), ...]}` |
| `read_performance_roles` | 演出分配 | `{曲目: {role_key: [姓名, ...]}}` |
| `read_backstage_staff` | 表演後台工作分配表 | `{姓名: role_key}` |

`role_key` 為 `"general_core"` / `"mic"` / `"control"` / `"curtain"` / `"tuning"` 其中之一。

**依賴：** `models`、`utils`。

---

### `guzheng/builders.py` — 資料組合層

接收 `readers` 回傳的原始資料，合併或重建成完整的 domain 物件：

#### `build_song_people`
將明細（古箏表演者）與演出分配（打擊 / 鋼琴 / 低音提琴）合併成每首曲目的 `SongPeople`。

#### `rebuild_song_assets_from_detail`
優先使用「明細」工作表的箏架資料重建 `SongAsset`，若某首曲目在明細中無資料才退回「曲目」工作表的原始資料。**明細是箏架分配的最終依據。**

**依賴：** `models`、`utils`。

---

### `guzheng/timeline.py` — 時間軸建構

從演出開始時間線性往後推算，產生每個列的時間區間字串（格式：`HH:MM~HH:MM`）。

#### `build_timeline_maps`
回傳兩個字典：
- `song_time_map`：`{曲目: 時間區間}`
- `transition_time_map`：`{(prev_song, next_song, row_type): 時間區間}`

`row_type` 共六種：`"pre_opening"`、`"normal_transition"`、`"before_intermission"`、`"intermission"`、`"pre_second_half"`、`"last_teardown"`。

**依賴：** 無。

---

### `guzheng/scheduler.py` — 換場任務分配演算法

專案中最複雜的模組，純業務邏輯，不涉及任何 Excel 操作。

#### `TaskBucket`
收集每個人的任務字串，最後輸出為 `【姓名】任務A、任務B` 格式的顯示文字。

#### `PoolPicker`
後台人員池管理與任務指派，對於同一個時間段的核心約束：
- 若是已經上了  **1 個** 古箏，就不能再做其他任務，也不可以上 **2 次** 以上的古箏
- 上一場的表演者自己下自己的古箏，不可以再有其他任務
- 下一場的表演者優先自己上自己的箏架和木椅
- 若是上一場的箏架/木椅和下一場的箏架/木椅有部分重複，則下一場的表演者可以不用上箏架/木椅而是調整該箏架/木椅。例如：上一場有 2 個紅架、2 張木椅，下一場需要 1 個紅架、2 個黑架、三張木椅，則使用紅架的那位表演者要在  `middle` 「調整紅架\*1、調整木椅\*1」，使用黑架的其中一位表演者要「上黑架\*1、調整木椅\*1」，另一位使用黑架的表演者要「上黑架\*1、、上木椅\*1」
- 下木椅一次最多下 **3 張** 
- 下箏架或是調整箏架一次最多 **2 副** 
- 分配完下箏 `left` 後，先分配下一場表演者自己上或調整自己的箏架、木椅 `middle`，然後分配上箏 `right`，最後分配剩下的工作 `middle`
- 如果缺少人力，則輸出【空缺】[工作名稱]
- `middle` 與 `right` 兩區人員互斥（已指派 middle 者不進 right，反之亦然）
- `middle_heavy_workers`：已在中台做過工作的人，優先避免讓他們再去搬箏

**`pick_for_up_guzheng` 選人邏輯：**
兩個輪次（第一輪每人最多 1 台、第二輪最多 2 台），每輪各有三段遞進的排除條件（先嚴格後放寬），共最多 12 次嘗試。

#### `generate_transition`
換場任務分配的主要函式，回傳三個欄位的文字：

| 欄位 | 任務來源 | 說明 |
|---|---|---|
| `left` | 上一首表演者自填 | 各自下自己的箏 / 樂器 |
| `middle` | `PoolPicker` 指派 | 調整箏架、木椅、鋼琴椅；布幕操作 |
| `right` | `PoolPicker` 指派 | 上箏；特殊樂器（打擊、低音提琴、電鋼琴）上台 |

三個特殊旗標控制行為：`is_pre_row`（全空白）、`is_before_intermission`（只有 left）、`is_last_song_teardown`（布幕改為謝幕降幕）。

**依賴：** `models`、`utils`、`config`。

---

### `guzheng/styles.py` — Excel 視覺樣式層

**純呈現，零業務邏輯。** 所有顏色、字體、框線的定義與套用函式都在這裡。若要調整視覺風格，只需改此檔。

| 項目 | 說明 |
|---|---|
| 色票常數 | `_COLOR_HEADER`（表頭）、`_COLOR_SONG`（曲目列）、`_COLOR_INTERMISSION`（中場/結尾） |
| `header_fill` / `song_fill` / `intermission_fill` | 建立對應的 `PatternFill` 物件 |
| `thin_border` | 建立底部細框線 |
| `style_cell` | 套用對齊、字體、填色到單一儲存格 |
| `style_row` | 對整列所有欄套用相同樣式 |
| `apply_global_font_size` | 全 workbook 統一字體大小，保留粗體 / 斜體等其他屬性 |

**依賴：** 無業務模組。

---

### `guzheng/writer.py` — 輸出工作簿組裝層

唯一同時了解業務資料與 Excel 格式的模組，負責把兩者接合起來。

輸出列的排列順序：
```
[前置換場列]      ← 開場準備，全空白
[曲目列]          ← 藍色底，表演時間 + 曲目 + 表演人員
[換場列]          ← 下箏 / 換架椅 / 上箏
...（重複）
[中場前換場列]    ← 只有 left（下箏）
[中場休息列]      ← 灰色底，顯示時間區間
[下半場前置列]    ← 空白
...（繼續）
[最後換場列]      ← 拆台用
[結尾區塊]        ← 謝幕、老師致詞等固定文字
```

**依賴：** `models`、`scheduler`、`styles`、`config`、`utils`。

---

## 資料流程圖

```
SOURCE_XLSX
  ├─ 曲目          → read_song_order_and_assets  ─┐
  │                  read_song_duration_minutes    │
  ├─ 明細          → read_detail_people           ─┼─ builders ─→ song_people
  ├─ 演出分配      → read_performance_roles       ─┘             song_assets
  └─ 表演後台工作  → read_backstage_staff ──────────→ backstage_roles

song_order + duration_minutes_map ──→ timeline ──→ song_time_map
                                                    transition_time_map

song_people + song_assets + backstage_roles ──→ scheduler (generate_transition)
                                                  └─ PoolPicker（工作量平衡）

全部資料 ──→ writer ──→ styles ──→ OUTPUT_XLSX
```
