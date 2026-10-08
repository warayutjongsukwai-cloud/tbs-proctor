"""ระบบจัดเจ้าหน้าที่คุมสอบ TBS (proctor scheduler)."""
from .models import Staff, Session, Problem, ExamType, PayTier
from .demand import build_sessions, estimate_sessions
from .staff_io import load_roster
from .engine import solve, feasibility_check, detect_conflicts, Result
from .planner import plan_rooms, RoomPlan
from .rooms import load_rooms, allocate_session
from .excel_io import export

__all__ = ["Staff", "Session", "Problem", "ExamType", "PayTier",
           "build_sessions", "estimate_sessions", "load_roster", "solve",
           "feasibility_check", "detect_conflicts", "Result",
           "plan_rooms", "RoomPlan", "load_rooms", "allocate_session", "export"]
__version__ = "1.0.0"
