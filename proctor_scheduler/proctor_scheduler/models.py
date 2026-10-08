"""โครงสร้างข้อมูลหลักของระบบจัดเจ้าหน้าที่คุมสอบ TBS.

ทุกกฎในที่นี้อิงเอกสารวิเคราะห์ (reverse-engineering) ที่ยืนยันจากไฟล์จริง:
- ค่าตอบแทน 2 ระดับ (A1): เหลือง 300 / เขียว 100
- เพดาน 65 คนสอบ/ห้อง (A14), คนคุม 2 (<=18) / 3 (19-65) (A13)
"""
from __future__ import annotations
from dataclasses import dataclass, field
from enum import Enum
import math

# ---- ค่าคงที่เชิงกฎ (ปรับได้ที่เดียว) ----
PAY_YELLOW = 300          # นอกเวลาราชการ (A1)
PAY_GREEN = 100           # ในเวลาราชการ (A1)
ROOM_STUDENT_CAP = 65     # เพดานคนสอบต่อห้อง (A14)
PROCTOR_SMALL_MAX = 18    # <=18 คน -> 2 คนคุม (A13)
PROCTORS_SMALL = 2
PROCTORS_LARGE = 3        # 19-65 คน -> 3 คนคุม (A13)


class PayTier(str, Enum):
    YELLOW = "yellow"     # 300 บาท
    GREEN = "green"       # 100 บาท


class ExamType(str, Enum):
    MIDTERM = "midterm"
    FINAL = "final"
    MAKEUP = "makeup"     # สอบชดเชย -> เหลืองเสมอ (A11)


def proctors_for(n_students: int) -> int:
    """จำนวนคนคุมตามจำนวนคนสอบจริงในห้อง (กฎ A13)."""
    n = min(n_students, ROOM_STUDENT_CAP)
    return PROCTORS_SMALL if n <= PROCTOR_SMALL_MAX else PROCTORS_LARGE


def rooms_needed(enrollment: int) -> int:
    """จำนวนห้องที่วิชานี้ต้องใช้ (แตกห้องเมื่อเกิน 65 คน — A14)."""
    if enrollment <= 0:
        return 0
    return math.ceil(enrollment / ROOM_STUDENT_CAP)


def split_into_rooms(enrollment: int) -> list[int]:
    """แตกจำนวน นศ. เป็นห้องละ <=65 ให้สมดุล แล้วคืน list จำนวนคนต่อห้อง."""
    k = rooms_needed(enrollment)
    if k == 0:
        return []
    base, extra = divmod(enrollment, k)
    return [base + (1 if i < extra else 0) for i in range(k)]


@dataclass
class Staff:
    staff_id: str
    name: str                       # ชื่อที่ใช้แสดง (ชื่อเล่น)
    full_name: str = ""
    campus: str = "rangsit"
    program: str = ""               # BBA / IBMP / "" (ใช้ที่ท่าพระจันทร์)
    is_active: bool = True
    is_exempt: bool = False         # ขอยกเว้นทั้งเทอม (A6)
    carry_yellow: int = 0           # ยอดยกมา เหลือง (A2)
    carry_green: int = 0
    carry_pay: int = 0              # ยอดเงินสะสมยกมา (A2)
    note: str = ""


@dataclass
class Session:
    """คาบสอบ 1 ช่วง (เช่น อ.22 ก.ย. เช้า)."""
    session_id: str
    date: str                       # ข้อความวันที่ไทย เช่น 'อ.22 ก.ย.69'
    slot: str                       # เช้า / เที่ยง / บ่าย / เย็น
    start_time: str
    end_time: str
    tier: PayTier
    is_weekend: bool = False
    demand: int = 0                 # จำนวนคนคุมที่คาบนี้ต้องใช้

    @property
    def pay(self) -> int:
        return PAY_YELLOW if self.tier == PayTier.YELLOW else PAY_GREEN


@dataclass
class Unavailability:
    """วันไม่ว่าง/ขอยกเว้นรายครั้ง (A6) — ช่องสีดำในระบบเดิม."""
    staff_id: str
    session_id: str
    reason: str = ""


@dataclass
class Assignment:
    staff_id: str
    session_id: str
    is_locked: bool = False         # ผู้ใช้ล็อกไว้ (manual override)


@dataclass
class Problem:
    """ชุดข้อมูลครบสำหรับจัด 1 เทอม 1 ศูนย์."""
    staff: list[Staff] = field(default_factory=list)
    sessions: list[Session] = field(default_factory=list)
    unavailable: list[Unavailability] = field(default_factory=list)
    locked: list[Assignment] = field(default_factory=list)
    campus: str = "rangsit"
    term: str = ""

    def unavailable_map(self) -> dict[str, set[str]]:
        m: dict[str, set[str]] = {s.staff_id: set() for s in self.staff}
        for u in self.unavailable:
            m.setdefault(u.staff_id, set()).add(u.session_id)
        for s in self.staff:                 # คนยกเว้นทั้งเทอม = ไม่ว่างทุกคาบ
            if s.is_exempt:
                m[s.staff_id] = {ss.session_id for ss in self.sessions}
        return m

    def assignable_staff(self) -> list[Staff]:
        return [s for s in self.staff if s.is_active and not s.is_exempt]
