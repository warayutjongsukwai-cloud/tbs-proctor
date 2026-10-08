#!/usr/bin/env python3
"""ตัวรันหลัก: ตารางสอน + รายชื่อ -> จัดคิว -> Excel.

ใช้งาน:
  python run.py --schedule ตารางสอน.xlsx --roster คิว.xlsx --sheet "กลางภาค1-69รังสิต2" \
                --type midterm --out ผลจัดคิว.xlsx
"""
from __future__ import annotations
import argparse
import statistics
from proctor_scheduler import (build_sessions, load_roster, solve,
                               feasibility_check, export, Problem, ExamType)
from proctor_scheduler.demand import build_sessions_from_queue


def main():
    ap = argparse.ArgumentParser(description="จัดเจ้าหน้าที่คุมสอบ TBS")
    ap.add_argument("--schedule", required=True, help="ไฟล์ตารางสอน .xlsx")
    ap.add_argument("--roster", required=True, help="ไฟล์คิวเดิม (ดึงรายชื่อ)")
    ap.add_argument("--sheet", required=True, help="ชื่อชีทรายชื่อในไฟล์คิว")
    ap.add_argument("--type", default="midterm", choices=["midterm", "final"])
    ap.add_argument("--campus", default="rangsit")
    ap.add_argument("--term", default="")
    ap.add_argument("--exempt", default="", help="ชื่อคนยกเว้น คั่นด้วย , ")
    ap.add_argument("--time", type=int, default=30, help="เวลา solve สูงสุด (วินาที)")
    ap.add_argument("--demand-from-queue", action="store_true",
                    help="อ่าน demand จริงจากชีทคิว (แทนการคำนวณจาก enrollment)")
    ap.add_argument("--out", default="ผลจัดคิว.xlsx")
    a = ap.parse_args()

    et = ExamType.MIDTERM if a.type == "midterm" else ExamType.FINAL
    exempt = {x.strip() for x in a.exempt.split(",") if x.strip()}

    if a.demand_from_queue:
        print("1) อ่าน demand จริงจากชีทคิว ...")
        sessions = build_sessions_from_queue(a.roster, a.sheet, et)
    else:
        print("1) อ่านตารางสอน + คำนวณ demand (per-section) ...")
        sessions = build_sessions(a.schedule, et, a.campus)
    tot = sum(s.demand for s in sessions)
    ny = sum(1 for s in sessions if s.tier.value == "yellow")
    print(f"   คาบ {len(sessions)} (เหลือง {ny}/เขียว {len(sessions)-ny}) · demand รวม {tot} คน-คาบ")

    print("2) โหลดรายชื่อเจ้าหน้าที่ ...")
    staff = load_roster(a.roster, a.sheet, a.campus, exempt_names=exempt)
    avail = [s for s in staff if not s.is_exempt]
    print(f"   {len(staff)} คน (ยกเว้น {len(staff)-len(avail)})")

    problem = Problem(staff=staff, sessions=sessions, campus=a.campus, term=a.term)

    print("3) ตรวจกำลังคน ...")
    ok, msg = feasibility_check(problem)
    print("   " + msg.replace("\n", "\n   "))
    if not ok:
        print("   !! คนไม่พอบางคาบ — จัดต่อได้แต่คาบนั้นจะไม่ครบ (แก้: เพิ่มคน/ลดห้อง)")

    print(f"4) รัน CP-SAT (<= {a.time}s) ...")
    res = solve(problem, time_limit_s=a.time)
    print(f"   สถานะ {res.status} · {res.solve_seconds:.1f}s")
    if not res.is_ok():
        print("   !! หาคำตอบไม่ได้ (คนไม่พอ/ติดเงื่อนไข)"); return

    pays = [res.pay[s.staff_id] for s in avail]
    tots = [res.total[s.staff_id] for s in avail]
    print(f"   เงิน: {min(pays)}-{max(pays)} (SD {statistics.pstdev(pays):.0f}) · "
          f"คิว: {min(tots)}-{max(tots)} (SD {statistics.pstdev(tots):.1f})")

    export(problem, res, a.out,
           title=f"คิวกรรมการคุมสอบ {a.campus} {a.term}")
    print(f"5) บันทึก -> {a.out}")


if __name__ == "__main__":
    main()
