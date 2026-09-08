"""Task 3 -- English (DAE102) teacher assignment by five-week block.

Thirty weeks split into six blocks of five weeks (three per term), twenty hours
each, so every class carries 120 hours of English.  Each class needs one Net
teacher block per term; CS1 and CS2 need only one across the whole year because
Mr. Ivan Yuen holds 100 hours of each.

The two Net teachers are Ms. Loveleen Kaur and Mr. Peter Barrett.  Their hour
budgets leave no slack at all -- see the plan document -- so the Net placement is
solved exactly by backtracking rather than greedily, and the remaining blocks are
filled afterwards from the full-time staff first.
"""

from __future__ import annotations

import sys
from collections import defaultdict

from timetable_scheduler import _travel_mins

# --------------------------------------------------------------------------
# configuration -- all figures confirmed by Jo 2026-08-18
# --------------------------------------------------------------------------

BLOCK_HOURS = 20
BLOCKS_PER_TERM = 3
TERMS = (1, 2)
BLOCK_LABELS = ["T1 wk1-5", "T1 wk6-10", "T1 wk11-15",
                "T2 wk1-5", "T2 wk6-10", "T2 wk11-15"]

NET_TEACHERS = ["Ms. Loveleen Kaur", "Mr. Peter Barrett"]

# Hours each teacher may take on DAE102 across the whole year.
# Loveleen is 720 rather than Jo's 680 -- the top of the 640-720 range she gave
# in her first email, which is what makes the Net requirement add up.
HOUR_BUDGET = {
    "Ms. Loveleen Kaur": 720,
    "Mr. Peter Barrett": 360,
    "Ms. Hailey Wong": 660,
    "Ms. Cherry Ip": 420,
    "Mr. Ivan Yuen": 200,
}

# Ivan Yuen's 200 hours are earmarked for these two classes only.
RESTRICTED_TO = {"Mr. Ivan Yuen": {"DAE102_CS1", "DAE102_CS2"}}

# Classes that need only one Net block for the year, and the term it falls in.
HALF_NET = {"DAE102_CS1": 1, "DAE102_CS2": 2}

# Mr. Chris Hon has left and is no longer available for English (Jo, 2026-08-18).
# The 26-27 TeachingLoad sheet already omits him.
PT_POOL = ["Ms. Elise Ye", "Mr. Ray Leung", "Ms. Sasha Cheung",
           "Ms. Lee Kit Wan"]

ENGLISH_TRAVEL_CAP = 30   # minutes, English teachers only


def Issue(kind, subject, detail, action="", where=""):
    return {"kind": kind, "subject": subject, "detail": detail,
            "action": action, "where": where}


# --------------------------------------------------------------------------
# problem set-up
# --------------------------------------------------------------------------

def build_problem(class_rows):
    """Pick the DAE102 rows out of module A's output and index them by slot."""
    classes = {}
    for rec in class_rows:
        if not rec["code"].startswith("DAE102") or not rec.get("day"):
            continue
        classes[rec["code"]] = {
            "code": rec["code"],
            "day": rec["day"],
            "time": rec["time"],
            "start": rec["start"],
            "location": rec["location"],
            "centre": str(rec["room"]).split("-")[0].strip(),
            "lecturer1": rec["lecturers"][0] if rec["lecturers"] else None,
        }
    slots = defaultdict(list)
    for code, c in sorted(classes.items()):
        slots[(c["day"], c["time"])].append(code)
    return classes, dict(slots)


def _same_day_ok(classes, a, b):
    """Two classes a teacher holds in one block: same day means travel matters."""
    ca, cb = classes[a], classes[b]
    if ca["day"] != cb["day"]:
        return True
    if ca["start"] == cb["start"]:
        return False                       # literally the same slot
    return _travel_mins(ca["centre"], cb["centre"]) <= ENGLISH_TRAVEL_CAP


def _set_ok(classes, codes):
    codes = list(codes)
    for i in range(len(codes)):
        for j in range(i + 1, len(codes)):
            if not _same_day_ok(classes, codes[i], codes[j]):
                return False
    return True


# --------------------------------------------------------------------------
# phase 1 -- Net placement, solved exactly
# --------------------------------------------------------------------------

def _term_demand(slots, term):
    """Classes in each slot that need a Net block during this term."""
    out = {}
    for slot, codes in slots.items():
        out[slot] = [c for c in codes if HALF_NET.get(c, term) == term]
    return out


def solve_net_term(classes, slots, term, l_per_period, b_per_period, issues):
    """Assign the Net blocks for one term.

    Returns [{class: teacher}, ...] with one dict per block, or None.
    """
    demand = _term_demand(slots, term)
    slot_keys = sorted(demand)

    # Ms. Kaur takes one class from every slot in every block, so she covers
    # BLOCKS_PER_TERM classes per slot; Mr. Barrett covers whatever is left.
    n_b = {s: len(demand[s]) - BLOCKS_PER_TERM for s in slot_keys}
    short = [s for s, v in n_b.items() if v < 0]
    if short:
        issues.append(Issue(
            "net infeasible", f"term {term}",
            "these slots hold fewer classes than the blocks Ms. Kaur must fill: "
            + "; ".join(f"{d} {t}" for d, t in short)))
        return None
    if sum(n_b.values()) != b_per_period * BLOCKS_PER_TERM:
        issues.append(Issue(
            "net infeasible", f"term {term}",
            f"Mr. Barrett must cover {sum(n_b.values())} blocks but his budget "
            f"allows {b_per_period * BLOCKS_PER_TERM}"))
        return None

    loveleen, barrett = NET_TEACHERS
    remaining = {s: set(demand[s]) for s in slot_keys}
    b_left = dict(n_b)
    result = []

    def l_options(period_left):
        """Every travel-feasible way for Ms. Kaur to take one class per slot."""
        by_day = defaultdict(list)
        for s in slot_keys:
            by_day[s[0]].append(s)
        day_choices = []
        for day in sorted(by_day):
            day_slots = sorted(by_day[day], key=lambda s: s[1])
            combos = [[]]
            for s in day_slots:
                combos = [c + [x] for c in combos for x in sorted(remaining[s])]
            combos = [c for c in combos if _set_ok(classes, c)]
            if not combos:
                return None
            day_choices.append(combos)
        out = [[]]
        for choices in day_choices:
            out = [a + b for a in out for b in choices]
            if len(out) > 20000:
                break
        return out

    def b_options(taken):
        """Every travel-feasible triple Mr. Barrett can take this block."""
        usable = [s for s in slot_keys if b_left[s] > 0
                  and (remaining[s] - taken)]
        out = []

        def rec(i, chosen_slots, chosen):
            if len(chosen) == b_per_period:
                out.append(list(chosen))
                return
            if i >= len(usable):
                return
            # skip this slot
            rec(i + 1, chosen_slots, chosen)
            s = usable[i]
            for code in sorted(remaining[s] - taken):
                if _set_ok(classes, chosen + [code]):
                    chosen.append(code)
                    rec(i + 1, chosen_slots + [s], chosen)
                    chosen.pop()

        rec(0, [], [])
        return out

    def slot_of(code):
        return (classes[code]["day"], classes[code]["time"])

    def search(period):
        if period == BLOCKS_PER_TERM:
            return all(not v for v in remaining.values())
        # prune: every slot must still be able to place its remaining classes
        left = BLOCKS_PER_TERM - period
        for s in slot_keys:
            if len(remaining[s]) > left + min(b_left[s], left):
                return False

        for l_pick in (l_options(left) or []):
            if len(set(l_pick)) != len(slot_keys):
                continue
            taken = set(l_pick)
            for b_pick in b_options(taken):
                assignment = {c: loveleen for c in l_pick}
                assignment.update({c: barrett for c in b_pick})

                for c in l_pick:
                    remaining[slot_of(c)].discard(c)
                for c in b_pick:
                    s = slot_of(c)
                    remaining[s].discard(c)
                    b_left[s] -= 1
                result.append(assignment)

                if search(period + 1):
                    return True

                result.pop()
                for c in b_pick:
                    s = slot_of(c)
                    remaining[s].add(c)
                    b_left[s] += 1
                for c in l_pick:
                    remaining[slot_of(c)].add(c)
        return False

    if not search(0):
        issues.append(Issue("net infeasible", f"term {term}",
                            "no arrangement satisfies the 30-minute travel cap "
                            "together with the Net hour budgets"))
        return None
    return result


# --------------------------------------------------------------------------
# phase 2 -- fill the remaining blocks, full-time staff first
# --------------------------------------------------------------------------

def fill_remaining(classes, slots, net_blocks, issues):
    """Assign every non-Net block.  Returns (blocks, blocks_used_per_teacher).

    Full-time staff are loaded to their contracted hours first.  Each block gets
    a per-teacher quota of `hours still owed / blocks still to come`, so nobody
    is left stranded in the last block with hours they can no longer place.
    """
    blocks = [dict(b) for b in net_blocks]
    n_blocks = len(BLOCK_LABELS)

    budget = {t: HOUR_BUDGET[t] // BLOCK_HOURS for t in HOUR_BUDGET}
    used = defaultdict(int)
    for block in blocks:
        for teacher in block.values():
            used[teacher] += 1

    fulltime = ["Ms. Hailey Wong", "Ms. Cherry Ip"]
    restricted = [t for t in RESTRICTED_TO if t in budget]

    def allowed(teacher, code, bi):
        if teacher in RESTRICTED_TO and code not in RESTRICTED_TO[teacher]:
            return False
        if teacher in budget and used[teacher] >= budget[teacher]:
            return False
        held = [c for c, t in blocks[bi].items() if t == teacher]
        return _set_ok(classes, held + [code])

    def held_before(code, upto):
        return {t for bi in range(upto) for c, t in blocks[bi].items() if c == code}

    # Busiest slots first -- they have the fewest ways to be covered.
    order = sorted(classes, key=lambda c: (
        -len(slots[(classes[c]["day"], classes[c]["time"])]), c))

    for bi in range(n_blocks):
        blocks_left = n_blocks - bi

        # -- teachers tied to specific classes go first: those classes have
        #    exactly enough non-Net blocks to use up their hours, so anyone else
        #    taking one would strand them.
        for teacher in restricted:
            for code in sorted(RESTRICTED_TO[teacher]):
                if code not in blocks[bi] and allowed(teacher, code, bi):
                    blocks[bi][code] = teacher
                    used[teacher] += 1

        # -- full-time staff, most urgent first
        while True:
            pending = [t for t in fulltime if budget[t] - used[t] > 0]
            if not pending:
                break
            teacher = max(pending, key=lambda t: (budget[t] - used[t]) / blocks_left)
            need = budget[teacher] - used[teacher]
            quota = -(-need // blocks_left)          # ceiling division
            taken = 0
            candidates = [c for c in order if c not in blocks[bi]]
            candidates.sort(key=lambda c: (
                0 if classes[c]["lecturer1"] == teacher else 1,
                0 if teacher in held_before(c, bi) else 1,
                c))
            for code in candidates:
                if taken >= quota:
                    break
                if allowed(teacher, code, bi):
                    blocks[bi][code] = teacher
                    used[teacher] += 1
                    taken += 1
            if taken == 0:
                fulltime = [t for t in fulltime if t != teacher]
                if not fulltime:
                    break
            if all(budget[t] - used[t] <= 0 for t in fulltime):
                break
        fulltime = ["Ms. Hailey Wong", "Ms. Cherry Ip"]

        # -- part-time pool, keeping a class with the same teacher where possible
        for code in order:
            if code in blocks[bi]:
                continue
            named = classes[code]["lecturer1"]
            pool = ([named] if named in PT_POOL else []) +                    [t for t in PT_POOL if t != named]
            previous = held_before(code, bi)
            pool.sort(key=lambda t: (
                0 if t == named else 1,
                0 if t in previous else 1,
                used[t]))
            pick = next((t for t in pool if allowed(t, code, bi)), None)
            if pick is None:
                # Nobody is directly free.  Before giving up, see whether one
                # teacher can hand a single class to a colleague and take this
                # one instead -- the block is otherwise stranded on a technicality.
                if _repair_by_swap(classes, blocks, bi, code, budget, used):
                    continue
                issues.append(Issue(
                    "unfilled block", code,
                    f"{BLOCK_LABELS[bi]}: no teacher is free in this slot, "
                    "close enough to travel, and still within contracted hours",
                    action="Add an English teacher for this slot, or move the class"))
                continue
            blocks[bi][code] = pick
            used[pick] += 1

    for teacher, target in budget.items():
        if used[teacher] != target:
            issues.append(Issue(
                "budget mismatch", teacher,
                f"assigned {used[teacher] * BLOCK_HOURS} hrs against a target of "
                f"{target * BLOCK_HOURS} hrs"))
    return blocks, used


def _repair_by_swap(classes, blocks, bi, code, budget, used):
    """Free up a teacher for `code` by moving one of their classes elsewhere.

    Only a single hand-over is tried.  Anything deeper would be hard to explain
    to Jo, and if one swap is not enough the pool really is too small.
    """
    everyone = list(HOUR_BUDGET) + PT_POOL

    def free_for(teacher, target, skip=None):
        if teacher in RESTRICTED_TO and target not in RESTRICTED_TO[teacher]:
            return False
        held = [c for c, t in blocks[bi].items() if t == teacher and c != skip]
        return _set_ok(classes, held + [target])

    for teacher in everyone:
        if teacher in RESTRICTED_TO and code not in RESTRICTED_TO[teacher]:
            continue
        if teacher in budget and used[teacher] >= budget[teacher]:
            continue
        held = [c for c, t in blocks[bi].items() if t == teacher]
        blocking = [c for c in held if not _same_day_ok(classes, c, code)]
        if len(blocking) != 1:
            continue
        give_away = blocking[0]
        for other in everyone:
            if other == teacher:
                continue
            # `other` picks up an extra block, so they need room in their budget.
            if other in budget and used[other] >= budget[other]:
                continue
            if not free_for(other, give_away):
                continue
            blocks[bi][give_away] = other
            blocks[bi][code] = teacher
            used[other] += 1          # teacher swaps one class for another
            return True
    return False


# --------------------------------------------------------------------------
# entry point
# --------------------------------------------------------------------------

def assign(class_rows, issues):
    classes, slots = build_problem(class_rows)
    if not classes:
        issues.append(Issue("english", "-", "no DAE102 class has a placement"))
        return None

    l_budget = HOUR_BUDGET["Ms. Loveleen Kaur"] // BLOCK_HOURS
    b_budget = HOUR_BUDGET["Mr. Peter Barrett"] // BLOCK_HOURS
    n_blocks = len(BLOCK_LABELS)
    l_per_period, b_per_period = l_budget // n_blocks, b_budget // n_blocks

    if l_per_period != len(slots):
        issues.append(Issue(
            "net structure", "Ms. Loveleen Kaur",
            f"{l_budget} blocks over {n_blocks} periods is {l_per_period} per "
            f"period but there are {len(slots)} distinct English slots"))

    net_blocks = []
    for term in TERMS:
        solved = solve_net_term(classes, slots, term,
                                l_per_period, b_per_period, issues)
        if solved is None:
            return None
        net_blocks.extend(solved)

    blocks, used = fill_remaining(classes, slots, net_blocks, issues)
    return {"classes": classes, "slots": slots, "blocks": blocks, "used": used}
