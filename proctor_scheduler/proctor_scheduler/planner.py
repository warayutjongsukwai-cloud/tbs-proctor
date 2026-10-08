"""รวมขั้นตอน: ตารางสอน -> จัดห้อง (พบ>SC1>SC3) -> Session+demand + แผนห้อง."""
from __future__ import annotations
from dataclasses import dataclass
from .models import Session, PayTier, ExamType
from .exams import exams_per_session
from .rooms import load_rooms, allocate_session, RoomAllocation


@dataclass
class RoomPlan:
    sessions: list[Session]                         # มี demand = ผลรวมคนคุมของห้องในคาบ
    allocations: dict[str, list[RoomAllocation]]    # session_id -> ห้องที่จัด
    unplaced: dict[str, list]                       # session_id -> วิชาที่ห้องไม่พอ


def _tier(date, slot, et):
    if et == ExamType.MAKEUP:
        return PayTier.YELLOW
    if _is_weekend(date) or slot == 'เย็น':
        return PayTier.YELLOW
    return PayTier.GREEN


def _is_weekend(datestr: str) -> bool:
    s = str(datestr).strip()
    return s.startswith('อา.') or s.startswith('ส.')


def plan_rooms(schedule_xlsx: str, exam_type: ExamType = ExamType.MIDTERM,
               campus: str = "รังสิต") -> RoomPlan:
    per = exams_per_session(schedule_xlsx, exam_type)
    rooms = load_rooms(campus=campus)
    sessions, allocs, unplaced = [], {}, {}
    for i, (key, units) in enumerate(sorted(per.items())):
        date, slot, start, end = key
        sid = f"S{i+1:02d}"
        a, un = allocate_session([(u.label, u.students) for u in units], rooms)
        for x in a:
            x.session_key = sid
        demand = sum(x.proctors for x in a)
        sessions.append(Session(
            session_id=sid, date=date, slot=slot, start_time=start, end_time=end,
            tier=_tier(date, slot, exam_type), is_weekend=_is_weekend(date),
            demand=demand))
        allocs[sid] = a
        if un:
            unplaced[sid] = un
    return RoomPlan(sessions=sessions, allocations=allocs, unplaced=unplaced)


def distribute_proctors(sessions, allocations: dict, assignments: dict) -> dict:
    """แจกคนที่ engine จัด (ระดับคาบ) ลงห้อง โดยเกลี่ยการไป SC ให้เท่ากัน (A8).

    คืน dict[session_id] -> list[(RoomAllocation, [ชื่อผู้คุม...])]
    หลักการ: ห้อง SC ให้คนที่ยังไป SC น้อยก่อน; ห้อง พบ. ใครก็ได้.
    """
    sc_count: dict[str, int] = {}
    out: dict[str, list] = {}
    for se in sessions:
        sid = se.session_id
        allocs = allocations.get(sid, [])
        if not allocs:
            out[sid] = []
            continue
        # คนที่ถูกจัดในคาบนี้
        pool = [st for st, slist in assignments.items() if sid in slist]
        for st in pool:
            sc_count.setdefault(st, 0)
        # จัดห้อง พบ. ก่อน (ใครก็ได้) แล้ว SC (คนไป SC น้อยก่อน)
        pb = [a for a in allocs if a.room.building == "คณะพาณิชย์"]
        sc = [a for a in allocs if a.room.building != "คณะพาณิชย์"]
        assigned: set[str] = set()
        plan = []
        for a in pb:
            take = [st for st in pool if st not in assigned][:a.proctors]
            assigned.update(take); plan.append((a, take))
        for a in sc:
            cand = sorted([st for st in pool if st not in assigned],
                          key=lambda s: sc_count[s])
            take = cand[:a.proctors]
            for st in take:
                sc_count[st] += 1
            assigned.update(take); plan.append((a, take))
        out[sid] = plan
    return out


def sc_balance_report(distributed: dict, id2name: dict) -> dict[str, int]:
    """นับจำนวนครั้งที่แต่ละคนไป SC (สำหรับตรวจความสมดุล A8)."""
    cnt: dict[str, int] = {}
    for sid, plan in distributed.items():
        for a, names in plan:
            if a.room.building != "คณะพาณิชย์":
                for st in names:
                    cnt[id2name.get(st, st)] = cnt.get(id2name.get(st, st), 0) + 1
    return cnt
