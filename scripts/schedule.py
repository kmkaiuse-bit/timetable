"""
schedule.py — 跑完整排程並「產生排班結果 Excel」。

用法：
    python scripts/schedule.py "路徑/你的.xlsx"

它會：
  1. 執行完整自動排班（日/時/室/老師）。
  2. 產生一個結果檔 <你的檔名>_排班結果.xlsx，
     排好的時間表在「Class Assignments」分頁，另有每日時間表與問題清單。
  3. 印出摘要（排到多少、有多少問題、結果檔位置）。

注意：verify.py 只報數字、不儲存；要拿到填好的時間表，用這個 schedule.py。
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))), "api"))

import timetable_scheduler as ts


def run(xlsx_path: str) -> None:
    ts.load_rules_config()                       # apply config/rules.json
    data = open(xlsx_path, "rb").read()
    output_bytes, stats = ts.run_v4_from_bytes(data)

    stem = os.path.splitext(xlsx_path)[0]
    out_path = f"{stem}_排班結果.xlsx"
    with open(out_path, "wb") as f:
        f.write(output_bytes)

    scheduled = stats.get("scheduled", "?")
    issues = stats.get("issues", []) or []
    print("=" * 56)
    print(f"輸入：{os.path.basename(xlsx_path)}")
    print(f"✅ 已排班：{scheduled} 班")
    print(f"⚠️  問題（排不到／需注意）：{len(issues)}")
    print("-" * 56)
    print(f"📄 排班結果已寫入：\n   {out_path}")
    print("   （排好的時間表在「Class Assignments」分頁；另有每日時間表分頁）")
    print("=" * 56)
    print("提示：想只看排到／排不到的數字，用 scripts/verify.py。")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print('用法：python scripts/schedule.py "路徑/你的.xlsx"')
        sys.exit(1)
    run(sys.argv[1])
