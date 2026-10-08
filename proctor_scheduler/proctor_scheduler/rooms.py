"""Room Master + ตัวจัดห้องสอบ (room allocation) ตามลำดับตึก พบ. > SC1 > SC3.

กฎ (ยืนยันจากผู้ใช้ 8 ต.ค.):
- เลือกห้องตามลำดับตึก: คณะพาณิชย์ (พบ.) ก่อน → SC1 → SC-3
- แต่ละ section = 1 ห้อง (วิชาเล็ก <10 คน รวมกันได้ — ทำใน exams.py)
- ห้องจุคนสอบ <=65 (A14), คนคุม 2 (<=18) / 3 (19-65) (A13)
"""
from __future__ import annotations
import csv
import os
from dataclasses import dataclass, field
from .models import proctors_for, ROOM_STUDENT_CAP

BUILDING_PRIORITY = {"คณะพาณิชย์": 0, "SC1": 1, "SC-3": 2}
_DATA = os.path.join(os.path.dirname(__file__), "data", "room_master.csv")


@dataclass
class Room:
    room_id: str
    building: str
    floor: str
    capacity: int            # ที่นั่งสอบใช้จริง (÷2, cap 65)
    priority: int = 99

    def proctors_for_n(self, n: int) -> int:
        return proctors_for(n)


@dataclass
class RoomAllocation:
    session_key: str
    room: Room
    courses: list[str]       # วิชาที่สอบในห้องนี้ (ปกติ 1, >1 ถ้ารวมวิชาเล็ก)
    students: int
    proctors: int


def load_rooms(path: str = _DATA, campus: str = "รังสิต") -> list[Room]:
    rooms = []
    with open(path, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            if r["ศูนย์"] != campus:
                continue
            cap = int(r["ความจุใช้จริง(cap65)"])
            b = r["อาคาร"]
            rooms.append(Room(room_id=r["ห้อง"], building=b, floor=r["ชั้น"],
                              capacity=cap, priority=BUILDING_PRIORITY.get(b, 99)))
    return rooms


def _room_order(rooms: list[Room]) -> list[Room]:
    """เรียงห้อง: ตึกตามลำดับความสำคัญ → ในตึกเรียงความจุน้อยไปมาก (ใช้ห้องเล็กก่อน)."""
    return sorted(rooms, key=lambda r: (r.priority, r.capacity))


def allocate_session(exams: list[tuple[str, int]], rooms: list[Room]
                     ) -> tuple[list[RoomAllocation], list[tuple[str, int]]]:
    """จัดห้องให้ 1 คาบ.

    exams: list ของ (course_label, students) — แต่ละตัว = 1 ห้อง (จัด/รวมมาแล้ว)
    คืน (allocations, unplaced) ; unplaced = ที่ห้องไม่พอ
    """
    pool = _room_order(rooms)
    used = set()
    allocs = []
    unplaced = []
    # section ใหญ่ก่อน (หาห้องยากกว่า)
    for label, n in sorted(exams, key=lambda e: -e[1]):
        n_eff = min(n, ROOM_STUDENT_CAP)
        pick = next((r for r in pool
                     if r.room_id not in used and r.capacity >= n_eff), None)
        if pick is None:
            unplaced.append((label, n))
            continue
        used.add(pick.room_id)
        allocs.append(RoomAllocation(
            session_key="", room=pick, courses=[label],
            students=n, proctors=pick.proctors_for_n(n)))
    # เรียงผลตามลำดับตึก/ห้อง เพื่ออ่านง่าย
    allocs.sort(key=lambda a: (a.room.priority, a.room.floor, a.room.room_id))
    return allocs, unplaced


def building_usage(allocs: list[RoomAllocation]) -> dict[str, int]:
    out: dict[str, int] = {}
    for a in allocs:
        out[a.room.building] = out.get(a.room.building, 0) + 1
    return out
