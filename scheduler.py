from itertools import combinations
from datetime import datetime
from models import LPAAssignment, PairingRecord
from utils import is_shift_compatible, months_between
import random

LPAS_PER_EMPLOYEE = 2
PAIRING_LOCK_MONTHS = 4


def paired_recently(a, b, history, month, year):
    for r in history:
        if {r.emp_a, r.emp_b} == {a, b}:
            if months_between(r.month, r.year, month, year) < PAIRING_LOCK_MONTHS:
                return True
    return False


def score_pair(e1, e2, history, month, year):
    score = 0
    if e1.role == e2.role:
        score += 10
    if e1.shift == e2.shift:
        score += 5
    if paired_recently(e1.name(), e2.name(), history, month, year):
        score += 100
    return score


def schedule_month(employees, sections, shifts, history, month, year):
    assignments = []
    count = {e.name(): 0 for e in employees if e.active}
    section_tracker = {e.name(): [] for e in employees if e.active}  # Track sections per employee

    # Create list of all section-shift combinations
    all_combinations = [(s, t) for s in sections for t in shifts]
    random.shuffle(all_combinations)  # Randomize to avoid patterns

    # First pass: assign as many 2-person audits as possible
    for section, target in all_combinations:
        # Skip if we already have enough assignments
        if all(count[e.name()] >= LPAS_PER_EMPLOYEE for e in employees if e.active):
            break

        pairs = []
        for e1, e2 in combinations(employees, 2):
            if not (e1.active and e2.active):
                continue
            if not is_shift_compatible(e1.shift, e2.shift, target):
                continue
            if count[e1.name()] >= LPAS_PER_EMPLOYEE or count[e2.name()] >= LPAS_PER_EMPLOYEE:
                continue

            # Calculate penalty for duplicate sections
            penalty = 0
            if section in section_tracker[e1.name()]:
                penalty += 15
            if section in section_tracker[e2.name()]:
                penalty += 15

            base_score = score_pair(e1, e2, history, month, year)
            total_score = base_score + penalty
            pairs.append((e1, e2, total_score))

        if not pairs:
            continue

        pairs.sort(key=lambda x: x[2])

        # Try to find a pair where at least one hasn't done this section
        selected_pair = None
        for e1, e2, score in pairs:
            if section not in section_tracker[e1.name()] or section not in section_tracker[e2.name()]:
                selected_pair = (e1, e2)
                break

        if not selected_pair:
            e1, e2, _ = pairs[0]
            selected_pair = (e1, e2)

        e1, e2 = selected_pair

        assignments.append(LPAAssignment(section, target, [e1.name(), e2.name()]))
        count[e1.name()] += 1
        count[e2.name()] += 1
        section_tracker[e1.name()].append(section)
        section_tracker[e2.name()].append(section)
        history.append(PairingRecord(e1.name(), e2.name(), month, year))

    # Second pass: handle any remaining with 3-person teams
    for section, target in all_combinations:
        # Find employees who still need assignments
        needing_assignments = [e for e in employees if e.active and count[e.name()] < LPAS_PER_EMPLOYEE]

        if len(needing_assignments) < 3:
            continue

        # Check if this section-shift is already assigned
        if any(a.section == section and a.target_shift == target for a in assignments):
            continue

        # Try to form a trio
        for trio in combinations(needing_assignments, 3):
            names = [e.name() for e in trio]
            shifts_ok = any(e.shift == target for e in trio)

            if shifts_ok:
                assignments.append(LPAAssignment(section, target, names))
                for name in names:
                    count[name] += 1
                    section_tracker[name].append(section)
                break

    # Third pass: make sure everyone has exactly 2 LPAs
    for employee in employees:
        if not employee.active:
            continue

        while count[employee.name()] < LPAS_PER_EMPLOYEE:
            # Try to add to existing assignment
            added = False
            for assignment in assignments:
                if employee.name() in assignment.auditors:
                    continue
                if len(assignment.auditors) >= 3:
                    continue
                if not is_shift_compatible(employee.shift, assignment.target_shift, assignment.target_shift):
                    continue

                assignment.auditors.append(employee.name())
                count[employee.name()] += 1
                section_tracker[employee.name()].append(assignment.section)
                added = True
                break

            if not added:
                # Create new assignment if needed
                for section, target in all_combinations:
                    if any(a.section == section and a.target_shift == target for a in assignments):
                        continue

                    # Find another employee
                    for other in employees:
                        if not other.active or other.name() == employee.name():
                            continue
                        if count[other.name()] >= LPAS_PER_EMPLOYEE:
                            continue
                        if not is_shift_compatible(employee.shift, other.shift, target):
                            continue

                        assignments.append(LPAAssignment(section, target, [employee.name(), other.name()]))
                        count[employee.name()] += 1
                        count[other.name()] += 1
                        section_tracker[employee.name()].append(section)
                        section_tracker[other.name()].append(section)
                        history.append(PairingRecord(employee.name(), other.name(), month, year))
                        added = True
                        break

                    if added:
                        break

            if not added:
                print(f"⚠ Warning: Could not assign second LPA to {employee.name()}")
                break

    # Print summary
    print("\n📊 Section Assignment Summary:")
    for e in employees:
        if e.active:
            sections = section_tracker[e.name()]
            unique_sections = set(sections)
            print(f"  {e.name()}: {len(sections)} LPAs, {len(unique_sections)} unique sections")

    return assignments