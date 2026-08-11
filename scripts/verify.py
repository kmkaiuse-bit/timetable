"""
verify.py — 跑一次排程並輸出指標報告。

用法：
    python scripts/verify.py "路徑/你的.xlsx"

輸出：目前套用的規則、排到/排不到、超額違規（應為 0）、各中心緊張程度。
這是 AI agent 每次更改後應該跑、並把結果回報給同事的「安全網」。
"""
import os
import sys
from collections import Counter, defaultdict
from io import BytesIO

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))), "api"))

import openpyxl
import timetable_scheduler as ts


def run(xlsx_path: str) -> None:
    applied = ts.load_rules_config()          # apply config/rules.json
    data = open(xlsx_path, "rb").read()
    wb = openpyxl.load_workbook(BytesIO(data), data_only=True, read_only=True)
    conn, _, _ = ts.build_db(wb)
    wb.close()

    ts.phase0_schedule_cadets(conn)
    ts.auto_assign_schedule(conn, cc_only=False)
    ts.cc_assign_schedule(conn)

    total = conn.execute("SELECT COUNT(*) FROM classes").fetchone()[0]
    placed = conn.execute(
        "SELECT COUNT(DISTINCT class_code) FROM schedule WHERE day IS NOT NULL"
    ).fetchone()[0]
    over = conn.execute("""
        SELECT c.code, c.student_count, r.code, r.capacity
        FROM schedule s
        JOIN classes c ON c.code = s.class_code
        JOIN rooms r ON r.code = s.room_code
        WHERE s.day IS NOT NULL AND c.student_count > r.capacity
    """).fetchall()
    un = conn.execute("""
        SELECT c.code, cg.centre FROM classes c
        JOIN class_groups cg ON cg.code = c.group_code
        WHERE c.code NOT IN (SELECT class_code FROM schedule WHERE day IS NOT NULL)
    """).fetchall()
    by_centre = Counter(r["centre"] for r in un)

    print("=" * 56)
    print(f"檔案：{os.path.basename(xlsx_path)}")
    print("目前套用的規則（來自 config/rules.json）：")
    print(f"   DAE 上課日       : {applied.get('dae_days', '（用預設）')}")
    print(f"   Cadet 上課日     : {applied.get('cadet_days', '（用預設）')}")
    print(f"   老師一週上限     : {applied.get('teacher_weekly_cap', '（用預設）')}")
    print("-" * 56)
    print(f"總班數           : {total}")
    print(f"✅ 排到           : {placed}")
    print(f"❌ 排不到         : {total - placed}")
    print(f"⚠️  超額（人數坐不下，應為 0）: {len(over)}")
    if over:
        for c, st, rm, cap in over:
            print(f"      {c}: {st} 人 > {rm} ({cap} 座)")
    if by_centre:
        print("排不到 — 各中心：")
        for ct, n in sorted(by_centre.items(), key=lambda x: -x[1]):
            print(f"      {ct}: {n}")
    print("=" * 56)
    conn.close()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print('用法：python scripts/verify.py "路徑/你的.xlsx"')
        sys.exit(1)
    run(sys.argv[1])
