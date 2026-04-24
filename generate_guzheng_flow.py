# -*- coding: utf-8 -*-
"""
古箏公演細流產生器 — 執行入口。

執行方式：
    python generate_guzheng_flow.py --source 資料整合檔.xlsx --output 細流輸出.xlsx

未指定路徑時，預設值來自 guzheng/config.py。
其他可調整參數請至 guzheng/config.py 修改。
"""
import argparse
from openpyxl import load_workbook

from guzheng.config import (
    SOURCE_XLSX, OUTPUT_XLSX,
    START_TIME, OPENING_MINUTES, DEFAULT_TRANSITION_MINUTES,
    INTERMISSION_MINUTES, INTERMISSION_AFTER,
    DEFAULT_FONT_SIZE, DEFAULT_FONT_NAME,
)
from guzheng.readers import (
    read_song_order_and_assets,
    read_song_duration_minutes,
    read_detail_people,
    read_performance_roles,
    read_backstage_staff,
)
from guzheng.builders import build_song_people, rebuild_song_assets_from_detail
from guzheng.timeline import build_timeline_maps
from guzheng.writer  import build_output_workbook
from guzheng.styles  import apply_global_font_size


def parse_args():
    parser = argparse.ArgumentParser(description="古箏公演細流產生器")
    parser.add_argument("--source", default=SOURCE_XLSX, help="資料整合檔 (.xlsx)")
    parser.add_argument("--output", default=OUTPUT_XLSX, help="輸出細流檔案路徑 (.xlsx)")
    parser.add_argument("--debug",  action="store_true",  help="印出每個換場的人員分配過程")
    return parser.parse_args()


def main():
    args = parse_args()

    print("1. 開始讀檔")
    src = load_workbook(args.source, data_only=False)

    print("2. 讀曲目與資產")
    song_order, song_assets_sheet = read_song_order_and_assets(src)

    print("3. 讀明細")
    detail_people = read_detail_people(src)

    print("4. 讀演出分配")
    role_people = read_performance_roles(src)

    print("5. 讀後台人員")
    backstage_roles = read_backstage_staff(src)

    print("6. 建 song_people")
    song_people = build_song_people(detail_people, role_people)

    print("6.1 用明細重建 song_assets")
    song_assets = rebuild_song_assets_from_detail(
        song_order=song_order,
        song_assets_from_sheet=song_assets_sheet,
        detail_people=detail_people,
    )

    print("7. 讀曲目時長")
    duration_minutes_map = read_song_duration_minutes(src)
    duration_map = {song: str(m) for song, m in duration_minutes_map.items()}

    print("8. 建時間軸")
    song_time_map, transition_time_map = build_timeline_maps(
        song_order=song_order,
        duration_minutes_map=duration_minutes_map,
        start_time_str=START_TIME,
        opening_minutes=OPENING_MINUTES,
        intermission_after=INTERMISSION_AFTER,
        transition_minutes=DEFAULT_TRANSITION_MINUTES,
        intermission_minutes=INTERMISSION_MINUTES,
    )

    print("9. 建輸出 workbook")
    out_wb = build_output_workbook(
        song_order=song_order,
        song_assets=song_assets,
        song_people=song_people,
        backstage_roles=backstage_roles,
        transition_time_map=transition_time_map,
        song_time_map=song_time_map,
        duration_map=duration_map,
        debug=args.debug,
    )

    print("10. 套字體")
    apply_global_font_size(out_wb, size=DEFAULT_FONT_SIZE, default_name=DEFAULT_FONT_NAME)

    print("11. 存檔")
    out_wb.save(args.output)
    print(f"已輸出：{args.output}")


if __name__ == "__main__":
    main()
