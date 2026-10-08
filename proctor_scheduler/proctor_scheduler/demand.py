"""อ่านตารางสอน (.xlsx) -> สร้าง Session + คำนวณ demand คนคุมต่อคาบ.

Demand = ผลรวมคนคุมของทุก section ในคาบนั้น โดยแต่ละ section:
  แตกเป็นห้องละ <=65 คน (A14) แล้วนับคนคุม 2/3 ตามขนาด (A13).
"""
from __future__ import annotations
import re
import datetime
from openpyxl import load_workbook
from .models import Session, PayTier, ExamType, proctors_for, split_into_rooms

FACULTY_SHEETS = ['AC', 'BA', 'FN', 'MK', 'MOEH', 'OMส่งคณะ', 'mis', 'IBLT', 'RB ', 'BBM']
_SKIP = ('จัดสอบเอง', 'ติดต่อผู้สอน')


def _is_weekend(datestr: str) -> bool:
    s = datestr.strip()
    return s.startswith('อา.') or s.startswith('ส.')


def _slot_from_time(time_range: str) -> str:
    """แปลงช่วงเวลาเป็นชื่อคาบ."""
    t = time_range.replace(' ', '')
    h = t.split('-')[0].split('.')[0]
    try:
        h = int(h)
    except ValueError:
        return 'อื่นๆ'
    if h < 12:
        return 'เช้า'
    if h < 15:
        return 'เที่ยง'          # 12.xx
    return 'เย็น'                 # 15.xx+ (กลางภาค); ปลายภาคใช้ 'บ่าย' เรียกเองได้


def _tier(datestr: str, slot: str, exam_type: ExamType) -> PayTier:
    if exam_type == ExamType.MAKEUP:
        return PayTier.YELLOW                        # A11
    if _is_weekend(datestr):
        return PayTier.YELLOW                        # A1
    if slot == 'เย็น':
        return PayTier.YELLOW                        # A1 (คาบเย็นวันธรรมดา)
    return PayTier.GREEN


def _parse_exam_cell(raw) -> tuple[str, str, str] | None:
    """คืน (date, start, end) หรือ None ถ้าไม่ใช่คาบสอบ."""
    if raw is None:
        return None
    s = str(raw).strip()
    if any(k in s for k in _SKIP) or not s:
        return None
    parts = s.split('\n')
    date = parts[0].strip()
    tr = parts[1].strip() if len(parts) > 1 else ''
    tr = tr.replace('น.', '').replace('น', '').strip()
    if '-' in tr:
        start, end = [p.strip() for p in tr.split('-', 1)]
    else:
        start, end = tr, ''
    return date, start, end


def _enroll(v) -> int:
    if v is None:
        return 0
    m = re.findall(r'\d+', str(v))
    return sum(int(x) for x in m) if m else 0


def _scan_sheet(ws, which: str):
    """คืน list ของ (exam_cell_raw, enrollment) ต่อ section. which='mid'|'final'."""
    rows = list(ws.iter_rows(values_only=True))
    hdr = None
    for i, r in enumerate(rows):
        j = ' '.join(str(v) for v in r if v)
        if 'วันสอบกลางภาค' in j and 'จำนวนนักศึกษา' in j:
            hdr = i
            break
    if hdr is None:
        return []
    col = {}
    for k, v in enumerate(rows[hdr]):
        if v is None:
            continue
        sv = str(v)
        if 'วันสอบกลางภาค' in sv:
            col['mid'] = k
        elif 'วันสอบปลายภาค' in sv:
            col['final'] = k
        elif 'จำนวนนักศึกษา' in sv:
            col['enr'] = k
    want = col.get(which)
    if want is None:
        return []
    out = []
    for r in rows[hdr + 2:]:
        code = r[0] if r else None
        if not code or not re.match(r'^[A-Z]{2,4}\d', str(code).strip()):
            continue
        raw = r[want] if want < len(r) else None
        enr = _enroll(r[col['enr']]) if 'enr' in col and col['enr'] < len(r) else 0
        out.append((raw, enr))
    return out


def build_sessions(xlsx_path: str, exam_type: ExamType = ExamType.MIDTERM,
                   campus: str = 'rangsit') -> list[Session]:
    which = 'mid' if exam_type == ExamType.MIDTERM else 'final'
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    # key = (date, start, end) -> demand สะสม
    agg: dict[tuple, int] = {}
    meta: dict[tuple, tuple] = {}
    for name in FACULTY_SHEETS:
        if name not in wb.sheetnames:
            continue
        for raw, enr in _scan_sheet(wb[name], which):
            parsed = _parse_exam_cell(raw)
            if parsed is None or enr <= 0:
                continue
            date, start, end = parsed
            slot = _slot_from_time(start + '-' + end if end else start)
            key = (date, slot, start, end)
            proctors = sum(proctors_for(n) for n in split_into_rooms(enr))
            agg[key] = agg.get(key, 0) + proctors
            meta[key] = (date, slot, start, end)

    sessions = []
    for i, (key, dem) in enumerate(sorted(agg.items())):
        date, slot, start, end = key
        tier = _tier(date, slot, exam_type)
        sessions.append(Session(
            session_id=f"S{i+1:02d}",
            date=date, slot=slot, start_time=start, end_time=end,
            tier=tier, is_weekend=_is_weekend(date), demand=dem))
    return sessions


# ============================================================
# ทางเลือก: อ่าน Session + demand จาก "ชีทคิวเดิม" โดยตรง
# ใช้เมื่อ demand ถูกกำหนดจากตารางสอบ/การจัดห้องแล้ว (ตรงกับของจริง)
# แทนการเดา demand จาก enrollment (ซึ่งต้องรู้กฎรวม section ก่อน)
# ============================================================
def _weekend_q(datestr: str) -> bool:
    s = str(datestr)
    if s.startswith('อา.') or s.startswith('ส.'):
        return True
    if s[:1] in 'จอพศ':
        return False
    try:
        y, m, d = s[:10].split('-')
        return datetime.date(int(y) - 543, int(m), int(d)).weekday() >= 5
    except Exception:
        return False


def build_sessions_from_queue(xlsx_path: str, sheet: str,
                              exam_type: ExamType = ExamType.MIDTERM,
                              first_row: int = 7) -> list[Session]:
    """สร้าง Session จากชีทคิว: demand = จำนวนคนที่ถูกจัดจริงในแต่ละคอลัมน์คาบ."""
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet]
    srow = None
    for r in range(2, first_row):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if any(v in ('เช้า', 'เที่ยง', 'เย็น', 'บ่าย') for v in vals):
            srow = r
            break
    if srow is None:
        raise ValueError("หาแถว slot (เช้า/เที่ยง/..) ไม่เจอ")
    drow = srow - 1
    trow = srow + 1
    sess_cols = [c for c in range(1, ws.max_column + 1)
                 if ws.cell(srow, c).value in ('เช้า', 'เที่ยง', 'เย็น', 'บ่าย')]
    sessions = []
    cur = None
    for i, c in enumerate(sess_cols):
        d = ws.cell(drow, c).value
        if d:
            cur = str(d).replace('\n', ' ').strip()
        slot = ws.cell(srow, c).value
        tr = str(ws.cell(trow, c).value or '')
        start, end = (tr.split('-', 1) + [''])[:2] if '-' in tr else (tr, '')
        demand = sum(1 for r in range(first_row, ws.max_row + 1)
                     if (ws.cell(r, c).value is not None
                         and str(ws.cell(r, c).value).strip() not in ('', '0')))
        we = _weekend_q(cur)
        tier = PayTier.YELLOW if (exam_type == ExamType.MAKEUP or we or slot == 'เย็น') else PayTier.GREEN
        sessions.append(Session(
            session_id=f"S{i+1:02d}", date=cur, slot=slot,
            start_time=start.strip(), end_time=end.strip(),
            tier=tier, is_weekend=we, demand=demand))
    return sessions


# ============================================================
# ประเมิน demand จากตารางสอน (ใช้เมื่อยังไม่รู้คนสอบจริง)
# กฎผู้ใช้: ต่าง section = แยกห้อง, วิชาคนสอบ <10 รวมกันได้, >65 แตกห้อง
# เนื่องจากตารางสอนมีแค่ "ที่นั่งที่เปิด" ไม่ใช่คนสอบจริง จึงเป็น "ค่าประมาณ"
# ที่ผู้ใช้ปรับแก้ได้ภายหลัง (ค่าเริ่มต้นคนคุม/ห้อง = default_proctors)
# ============================================================
def estimate_sessions(xlsx_path: str, exam_type: ExamType = ExamType.MIDTERM,
                      default_proctors: int = 2) -> list[Session]:
    """ประเมิน Session + demand โดยนับ section ต่อคาบ.

    demand = จำนวนห้อง × default_proctors
    จำนวนห้อง = จำนวน section (วิชาที่นั่งเปิด >65 แตกเพิ่มห้อง ตาม A14)
    merge <10: ประเมินล่วงหน้าไม่ได้ (ไม่รู้คนสอบจริง) -> ไม่หักในขั้นประเมิน
    """
    which = 'mid' if exam_type == ExamType.MIDTERM else 'final'
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    agg: dict[tuple, int] = {}      # (date,slot,start,end) -> จำนวนห้อง
    for name in FACULTY_SHEETS:
        if name not in wb.sheetnames:
            continue
        for raw, enr in _scan_sheet(wb[name], which):
            parsed = _parse_exam_cell(raw)
            if parsed is None:
                continue
            date, start, end = parsed
            slot = _slot_from_time(start + '-' + end if end else start)
            key = (date, slot, start, end)
            rooms = max(1, len(split_into_rooms(enr)) if enr > 0 else 1)
            agg[key] = agg.get(key, 0) + rooms
    sessions = []
    for i, (key, rooms) in enumerate(sorted(agg.items())):
        date, slot, start, end = key
        tier = _tier(date, slot, exam_type)
        sessions.append(Session(
            session_id=f"S{i+1:02d}", date=date, slot=slot,
            start_time=start, end_time=end, tier=tier,
            is_weekend=_is_weekend(date), demand=rooms * default_proctors))
    return sessions
