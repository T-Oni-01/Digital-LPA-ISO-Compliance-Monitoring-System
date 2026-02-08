from dataclasses import dataclass, field
from typing import List, Set

@dataclass
class Employee:
    first_name: str
    last_name: str
    role: str
    shift: int
    active: bool = True
    sections_audited: Set[str] = field(default_factory=set)  # Track sections audited this month

    def name(self):
        return f"{self.first_name} {self.last_name}"

@dataclass
class PairingRecord:
    emp_a: str
    emp_b: str
    month: int
    year: int


@dataclass
class LPAAssignment:
    section: str
    target_shift: int
    auditors: List[str]


@dataclass
class ISOAction:
    description: str
    owner: str
    due_date: str
    status: str  # Open / Closed