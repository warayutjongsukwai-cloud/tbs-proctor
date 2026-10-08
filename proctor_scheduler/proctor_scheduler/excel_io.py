"""Export ผลการจัด -> Excel หน้าตาเหมือนชีทเดิม (เจ้าหน้าที่คุ้นมือ).

2 ชีท:
  1) ตารางคุมสอบ : grid คน x คาบ, 1=ถูกจัด, คอลัมน์ เหลือง/เขียว/รวม + เงิน
  2) สรุปภาระรายคน : เช็คความแฟร์ + ไฮไลต์คนมาก/น้อยผิดปกติ
"""
from __future__ import annotations
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from .models import Problem, PayTier, PAY_YELLOW, PAY_GREEN
from .engine import Result

YELLOW_FILL = PatternFill("solid", fgColor="FFFF99")   # คาบนอกเวลา
GREEN_FILL = PatternFill("solid", fgColor="92D050")    # คาบในเวลา
HEAD_FILL = PatternFill("solid", fgColor="D9D9D9")
HIGH_FILL = PatternFill("solid", fgColor="FFC7CE")     # ไฮไลต์ผิดปกติ
BOLD = Font(bold=True)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)


def export(problem: Problem, result: Result, path: str, title: str = "",
           allocations: dict | None = None) -> str:
    wb = Workbook()
    _grid_sheet(wb.active, problem, result, title)
    _summary_sheet(wb.create_sheet("สรุปภาระรายคน"), problem, result)
    if allocations:
        _room_plan_sheet(wb.create_sheet("แผนจัดห้อง"), problem, result, allocations)
    wb.save(path)
    return path


def _room_plan_sheet(ws, problem, result, allocations):
    """ชีทแผนจัดห้อง: คาบ | ตึก | ห้อง | วิชา | คนสอบ | คนคุม(จำนวน) | ผู้คุม.

    แจกคนลงห้องด้วย distribute_proctors (เกลี่ยไป SC เท่ากัน — A8).
    """
    from .planner import distribute_proctors
    sess = {s.session_id: s for s in problem.sessions}
    id2name = {s.staff_id: s.name for s in problem.staff}
    dist = distribute_proctors(problem.sessions, allocations, result.assignments)
    heads = ["วันที่", "ช่วง", "เวลา", "ตึก", "ห้อง", "วิชา", "คนสอบ", "คนคุม", "ผู้คุมสอบ"]
    for j, h in enumerate(heads, 1):
        c = ws.cell(1, j, h); c.font = BOLD; c.fill = HEAD_FILL; c.alignment = CENTER; c.border = THIN
    row = 2
    for sid in sorted(allocations.keys()):
        se = sess.get(sid)
        if se is None:
            continue
        fill = YELLOW_FILL if se.tier == PayTier.YELLOW else GREEN_FILL
        for a, names in dist.get(sid, []):
            disp = [id2name.get(n, n) for n in names]
            vals = [se.date, se.slot, f"{se.start_time}-{se.end_time}",
                    a.room.building, a.room.room_id, ", ".join(a.courses),
                    a.students, a.proctors, ", ".join(disp)]
            for j, v in enumerate(vals, 1):
                c = ws.cell(row, j, v); c.border = THIN
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                if j in (1, 2, 3, 4, 5):
                    c.fill = fill
            row += 1
    for col, w in zip("ABCDEFGHI", [13, 7, 12, 12, 10, 24, 7, 7, 30]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"


def _grid_sheet(ws, problem: Problem, result: Result, title: str):
    ws.title = "ตารางคุมสอบ"
    sessions = problem.sessions
    staff = [s for s in problem.staff if s.is_active and not s.is_exempt]
    assigned = result.assignments

    ws.cell(1, 1, title or f"คิวกรรมการคุมสอบ {problem.campus} {problem.term}").font = Font(bold=True, size=14)
    # header: No | ชื่อ | เหลืองคิว เงิน | เขียวคิว เงิน | รวมคิว เงิน | คาบ...
    r_date, r_slot, r_time = 3, 4, 5
    fixed = ["No.", "ชื่อ", "เหลือง\nคิว", "เหลือง\nเงิน", "เขียว\nคิว", "เขียว\nเงิน", "รวม\nคิว", "รวม\nเงิน"]
    for j, h in enumerate(fixed, start=1):
        c = ws.cell(r_slot, j, h); c.font = BOLD; c.fill = HEAD_FILL; c.alignment = CENTER; c.border = THIN
    base = len(fixed)
    for k, se in enumerate(sessions):
        col = base + 1 + k
        ws.cell(r_date, col, se.date).alignment = CENTER
        sc = ws.cell(r_slot, col, se.slot); sc.alignment = CENTER; sc.font = BOLD
        ws.cell(r_time, col, f"{se.start_time}-{se.end_time}").alignment = CENTER
        fill = YELLOW_FILL if se.tier == PayTier.YELLOW else GREEN_FILL
        for rr in (r_date, r_slot, r_time):
            ws.cell(rr, col).fill = fill; ws.cell(rr, col).border = THIN

    row = r_time + 1
    for i, s in enumerate(staff, start=1):
        sess_ids = set(assigned.get(s.staff_id, []))
        y = sum(1 for se in sessions if se.session_id in sess_ids and se.tier == PayTier.YELLOW)
        g = sum(1 for se in sessions if se.session_id in sess_ids and se.tier == PayTier.GREEN)
        vals = [i, s.name, y, y * PAY_YELLOW, g, g * PAY_GREEN, y + g, y * PAY_YELLOW + g * PAY_GREEN]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(row, j, v); c.border = THIN
            if j == 2:
                c.alignment = Alignment(horizontal="left")
            else:
                c.alignment = CENTER
        for k, se in enumerate(sessions):
            col = base + 1 + k
            c = ws.cell(row, col); c.border = THIN; c.alignment = CENTER
            if se.session_id in sess_ids:
                c.value = 1
                c.fill = YELLOW_FILL if se.tier == PayTier.YELLOW else GREEN_FILL
        row += 1
    # demand row (ตรวจ)
    ws.cell(row, 2, "รวมคน/คาบ").font = BOLD
    for k, se in enumerate(sessions):
        ws.cell(row, base + 1 + k, se.demand).alignment = CENTER
        ws.cell(row, base + 1 + k).font = BOLD

    ws.column_dimensions["B"].width = 16
    ws.freeze_panes = ws.cell(r_time + 1, 3)


def _summary_sheet(ws, problem: Problem, result: Result):
    staff = [s for s in problem.staff if s.is_active and not s.is_exempt]
    heads = ["No.", "ชื่อ", "ยอดยกมา(เงิน)", "เหลือง", "เขียว", "คิวรวม", "เงินเทอมนี้", "เงินสะสม"]
    for j, h in enumerate(heads, 1):
        c = ws.cell(1, j, h); c.font = BOLD; c.fill = HEAD_FILL; c.alignment = CENTER
    pays = [result.pay.get(s.staff_id, 0) for s in staff]
    mean = sum(pays) / len(pays) if pays else 0
    for i, s in enumerate(staff, 1):
        y = result.yellow.get(s.staff_id, 0) - s.carry_yellow
        tot = result.total.get(s.staff_id, 0) - (s.carry_yellow + s.carry_green)
        term_pay = y * PAY_YELLOW + (tot - y) * PAY_GREEN
        cum = result.pay.get(s.staff_id, 0)
        vals = [i, s.name, s.carry_pay, y, tot - y, tot, term_pay, cum]
        for j, v in enumerate(vals, 1):
            c = ws.cell(i + 1, j, v)
            c.alignment = Alignment(horizontal="left" if j == 2 else "center")
        if abs(cum - mean) > 0.15 * mean and mean > 0:        # เบี่ยงจากเฉลี่ย >15%
            ws.cell(i + 1, 8).fill = HIGH_FILL
    n = len(staff)
    ws.cell(n + 3, 2, "เฉลี่ยเงินสะสม").font = BOLD
    ws.cell(n + 3, 8, round(mean)).font = BOLD
    ws.cell(n + 4, 2, "ช่วง (สูงสุด-ต่ำสุด)").font = BOLD
    ws.cell(n + 4, 8, (max(pays) - min(pays)) if pays else 0).font = BOLD
    ws.column_dimensions["B"].width = 16
