"""CP-SAT engine จัดเจ้าหน้าที่คุมสอบระดับคาบ (session-level, MVP phase 1).

หลักการ (ยืนยันจาก backtest 1/69):
  HARD  H1 แต่ละคาบต้องได้คนครบตาม demand
        H2 คนที่ไม่ว่าง/ยกเว้นในคาบนั้น ห้ามจัด
        (ล็อก) assignment ที่ผู้ใช้ล็อกไว้ ต้องคงอยู่
  SOFT  เรียงความสำคัญตามที่ผู้ใช้ยืนยัน:
        1) เงินรวม (carry+ปัจจุบัน) เท่ากัน      <- สำคัญสุด
        2) คิวรวม / เหลืองสะสม เท่ากัน
  วิธี: minimize ผลรวมถ่วงน้ำหนักของ (ช่วงสูงสุด-ต่ำสุด) ของแต่ละมิติ
"""
from __future__ import annotations
from dataclasses import dataclass
from ortools.sat.python import cp_model
from .models import Problem, PayTier

# ลำดับคาบในวัน (เช้า→เที่ยง→บ่าย→เย็น) สำหรับกฎ "คาบติดกัน"
SLOT_ORDER = {"เช้า": 0, "เที่ยง": 1, "บ่าย": 1, "เย็น": 2}


@dataclass
class Result:
    status: str
    assignments: dict[str, list[str]]        # staff_id -> [session_id,...]
    pay: dict[str, int]
    yellow: dict[str, int]
    total: dict[str, int]
    solve_seconds: float
    objective: int
    gaps: int = 0              # จำนวนครั้งที่คุมแบบเว้นคาบกลางวัน (ยิ่งน้อยยิ่งดี)
    shortfalls: dict | None = None   # session_id -> จำนวนคนที่ยังขาด (โหมดจัดเท่าที่ทำได้)

    def is_ok(self) -> bool:
        return self.status in ("OPTIMAL", "FEASIBLE")


# น้ำหนัก soft (ปรับได้) — ลำดับตามที่ผู้ใช้ยืนยัน:
#   1 เงินเท่ากัน > 2 คิว/เหลือง > 3 คาบติดกัน > 4 เว้นวันหยุด
W_PAY_RANGE = 100
W_YELLOW_RANGE = 10
W_TOTAL_RANGE = 5
W_GAP = 8                 # penalty ต่อ 1 ช่องว่างกลางวัน (S3)
W_WEEKEND_RANGE = 3       # เกลี่ยวันหยุด ส-อา (S4)


W_SHORT = 10_000_000      # penalty ต่อ 1 ที่นั่งที่ขาด — ใหญ่กว่าทุก soft เพื่อให้เติมคนให้มากสุดก่อน

def solve(problem: Problem, time_limit_s: int = 30,
          w_pay: int = W_PAY_RANGE, w_yellow: int = W_YELLOW_RANGE,
          w_total: int = W_TOTAL_RANGE, w_gap: int = W_GAP,
          w_weekend: int = W_WEEKEND_RANGE, allow_partial: bool = False) -> Result:
    staff = problem.assignable_staff()
    sessions = problem.sessions
    sid = {s.staff_id: s for s in staff}
    unavail = problem.unavailable_map()
    locked = {(a.staff_id, a.session_id) for a in problem.locked}

    m = cp_model.CpModel()
    x = {}
    for s in staff:
        for se in sessions:
            v = m.NewBoolVar(f"x_{s.staff_id}_{se.session_id}")
            x[(s.staff_id, se.session_id)] = v
            if se.session_id in unavail.get(s.staff_id, set()):
                m.Add(v == 0)                       # H2
    for (st, se) in locked:                          # ล็อก manual override
        if (st, se) in x:
            m.Add(x[(st, se)] == 1)

    short = {}                                       # H1 ต้องการครบ (หรือขาดได้ถ้า allow_partial)
    for se in sessions:
        assigned = sum(x[(s.staff_id, se.session_id)] for s in staff)
        if allow_partial:
            sh = m.NewIntVar(0, se.demand, f"short_{se.session_id}")
            m.Add(assigned + sh == se.demand)        # เติมให้มากสุด ที่ขาดไปเก็บใน sh
            short[se.session_id] = sh
        else:
            m.Add(assigned == se.demand)

    # ตัวแปรรวมต่อคน (บวก carry เพื่อความแฟร์ข้ามเทอม — A2)
    pay, yq, tot = {}, {}, {}
    for s in staff:
        pay[s.staff_id] = s.carry_pay + sum(
            x[(s.staff_id, se.session_id)] * se.pay for se in sessions)
        yq[s.staff_id] = s.carry_yellow + sum(
            x[(s.staff_id, se.session_id)] for se in sessions if se.tier == PayTier.YELLOW)
        tot[s.staff_id] = (s.carry_yellow + s.carry_green) + sum(
            x[(s.staff_id, se.session_id)] for se in sessions)

    max_pay = sum(se.pay for se in sessions) + max((s.carry_pay for s in staff), default=0)
    pmin = m.NewIntVar(0, max_pay, "pmin"); pmax = m.NewIntVar(0, max_pay, "pmax")
    ymin = m.NewIntVar(0, len(sessions) + 99, "ymin"); ymax = m.NewIntVar(0, len(sessions) + 99, "ymax")
    tmin = m.NewIntVar(0, len(sessions) + 99, "tmin"); tmax = m.NewIntVar(0, len(sessions) + 99, "tmax")
    for s in staff:
        m.Add(pmin <= pay[s.staff_id]); m.Add(pmax >= pay[s.staff_id])
        m.Add(ymin <= yq[s.staff_id]);  m.Add(ymax >= yq[s.staff_id])
        m.Add(tmin <= tot[s.staff_id]); m.Add(tmax >= tot[s.staff_id])

    # ---- S3: คาบในวันต้องติดกัน ไม่เว้นกลางวัน ----
    # จัดกลุ่มคาบตามวัน แล้วเรียงตามลำดับคาบ; ช่องว่าง = คุมคาบก่อน+หลัง แต่ไม่คุมคาบกลาง
    by_day: dict[str, list] = {}
    for se in sessions:
        by_day.setdefault(se.date, []).append(se)
    for d in by_day:
        by_day[d].sort(key=lambda s: SLOT_ORDER.get(s.slot, 9))
    gap_vars = []
    for s in staff:
        for d, day_sess in by_day.items():
            if len(day_sess) < 3:
                continue                      # ต้องมี >=3 คาบถึงจะเกิดช่องว่างกลาง
            for i in range(1, len(day_sess) - 1):
                prev = x[(s.staff_id, day_sess[i - 1].session_id)]
                cur = x[(s.staff_id, day_sess[i].session_id)]
                nxt = x[(s.staff_id, day_sess[i + 1].session_id)]
                g = m.NewBoolVar(f"gap_{s.staff_id}_{d}_{i}")
                # g=1 เมื่อ prev=1 และ nxt=1 และ cur=0
                m.Add(g >= prev + nxt - cur - 1)
                m.Add(g <= prev); m.Add(g <= nxt); m.Add(g <= 1 - cur)
                gap_vars.append(g)
    total_gap = m.NewIntVar(0, len(gap_vars) if gap_vars else 0, "total_gap")
    m.Add(total_gap == (sum(gap_vars) if gap_vars else 0))

    # ---- S4: เกลี่ยวันหยุด ส-อา ให้ได้พักบ้าง (balance จำนวนคาบวันหยุด/คน) ----
    we_sessions = [se for se in sessions if se.is_weekend]
    wmin = m.NewIntVar(0, len(we_sessions), "wmin"); wmax = m.NewIntVar(0, len(we_sessions), "wmax")
    if we_sessions:
        for s in staff:
            wc = sum(x[(s.staff_id, se.session_id)] for se in we_sessions)
            m.Add(wmin <= wc); m.Add(wmax >= wc)

    short_pen = W_SHORT * sum(short.values()) if short else 0
    m.Minimize(w_pay * (pmax - pmin) + w_yellow * (ymax - ymin)
               + w_total * (tmax - tmin) + w_gap * total_gap
               + w_weekend * (wmax - wmin) + short_pen)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit_s
    solver.parameters.num_search_workers = 8
    st = solver.Solve(m)
    status = solver.StatusName(st)

    assignments: dict[str, list[str]] = {s.staff_id: [] for s in staff}
    pv, yv, tv = {}, {}, {}
    gaps = 0
    if status in ("OPTIMAL", "FEASIBLE"):
        for s in staff:
            for se in sessions:
                if solver.Value(x[(s.staff_id, se.session_id)]):
                    assignments[s.staff_id].append(se.session_id)
            pv[s.staff_id] = solver.Value(pay[s.staff_id])
            yv[s.staff_id] = solver.Value(yq[s.staff_id])
            tv[s.staff_id] = solver.Value(tot[s.staff_id])
        gaps = solver.Value(total_gap)
    sf = None
    if short and status in ("OPTIMAL", "FEASIBLE"):
        sf = {se: solver.Value(v) for se, v in short.items() if solver.Value(v) > 0}
    return Result(status, assignments, pv, yv, tv, solver.WallTime(),
                  int(solver.ObjectiveValue()) if status in ("OPTIMAL", "FEASIBLE") else -1,
                  gaps=gaps, shortfalls=sf)


def feasibility_check(problem: Problem) -> tuple[bool, str]:
    """เช็คก่อนจัด: กำลังคนพอไหม (กฎ SKILL ข้อ 2 — ห้ามจัดฝืนแล้วเงียบ)."""
    staff = problem.assignable_staff()
    unavail = problem.unavailable_map()
    msgs = []
    ok = True
    for se in problem.sessions:
        avail = sum(1 for s in staff if se.session_id not in unavail.get(s.staff_id, set()))
        if avail < se.demand:
            ok = False
            msgs.append(f"  คาบ {se.date} {se.slot}: ต้องการ {se.demand} คน แต่ว่างแค่ {avail} คน")
    if ok:
        return True, "กำลังคนเพียงพอทุกคาบ"
    return False, "คาบที่คนไม่พอ:\n" + "\n".join(msgs)


def detect_conflicts(problem: Problem, assignments: dict[str, list[str]]) -> list[str]:
    """ตรวจ conflict หลังแก้มือ: ชนวันไม่ว่าง, คาบเกิน/ขาด demand, ภาระเบี่ยง."""
    unavail = problem.unavailable_map()
    sess = {s.session_id: s for s in problem.sessions}
    msgs = []
    # ชนวันไม่ว่าง
    for sid, slist in assignments.items():
        for se in slist:
            if se in unavail.get(sid, set()):
                nm = next((s.name for s in problem.staff if s.staff_id == sid), sid)
                msgs.append(f"🔴 {nm} ถูกจัดคาบ {sess[se].date} {sess[se].slot} ที่ตั้งว่าไม่ว่าง")
    # demand ครบไหม
    cnt = {se.session_id: 0 for se in problem.sessions}
    for slist in assignments.values():
        for se in slist:
            if se in cnt:
                cnt[se] += 1
    for se in problem.sessions:
        if cnt[se.session_id] != se.demand:
            msgs.append(f"🟠 คาบ {se.date} {se.slot}: จัด {cnt[se.session_id]} คน "
                        f"(ต้องการ {se.demand})")
    return msgs
