# Capacity Shortage Record — 2026-08

**Date:** 2026-08-11
**Source file:** `Planning for Timetable V4 (8) +CCGroup.xlsx` (auto-assign, DAE Mon/Tue/Thu)
**Engine state:** after the 2026-08 correctness fixes (no over-capacity placements;
combine tries all rooms/valid days). See `PRD-timetable-scheduler.md` §8.

> **Note (Jo, 2026-08-11):** adding rooms is **not realistically possible**. This record
> is kept so the shortage is documented; the realistic levers are **combine small classes**,
> **relocate classes to a centre with spare room**, or **accept the unplaced list**.

## Result

126 classes, **24 unplaceable** — all are "a big-enough room exists but every slot on
Mon/Tue/Thu is full" (`SLOT_ROOM`). None are data errors; none are too large for their
centre's biggest room.

## Per-centre ceiling vs demand

Ceiling = rooms × 3 days × 2 four-hour blocks. (2-hour classes pack tighter, so "placed"
can exceed the 4-hour ceiling.)

| Centre | Rooms | 4h ceiling | Classes | Placed | Unplaced | Verdict |
|--------|-------|-----------|---------|--------|----------|---------|
| FL 粉嶺 | 1 | 6 | 12 | 6 | 6 | 🔴 Hard ceiling — 1 room can't hold 12 |
| TS 沙田 | 1 | 6 | 16 | 10 | 6 | 🔴 Hard ceiling — badly oversubscribed |
| WT 黃大仙 | 1 | 6 | 8 | 6 | 2 | 🔴 Hard ceiling |
| TM 屯門 | 3 | 18 | 20 | 18 | 2 | 🔴 Hard ceiling |
| CSW 長沙灣 | 6 | 36 | 25 | 19 | 6 | 🟡 Room exists, blocked by big classes |
| KT 觀塘 | 3 | 18 | 16 | 14 | 2 | 🟡 Room exists, packing-limited |

- 🔴 **FL / TS / WT / TM (16 classes)** — structurally short of rooms; only adding rooms
  or **relocating classes to another centre** can place them.
- 🟡 **CSW / KT (8 classes)** — rooms exist but large classes fill them; **combine** or
  better packing may recover some.

## Action buckets

### 🟢 Combine-worth-trying (7 small classes, ≤25 students, same subject can share a room)

`DAE103_TS1 (20)`, `DAE103_TS2 (20)`, `DAE103_TS3 (20)`, `DAE103_TS4 (22)`,
`DAE106_TS3 (20)`, `DAE106_TS4 (5)`, `DAE106_FL3 (23)`

Fill the **CC Group** column to merge these (same subject; merged total must fit one room),
or combine into a centre with a spare room — as the DAE102 TS→TM combine did.

### 🔴 Need a room / relocation (17 large classes, ≥30 students)

`DAE103_CS5 (45)`, `DAE103_WT1 (42)`, `DAE103_WT2 (34)`, `DAE103_CS6 (33)`,
`DAE103_FL1 (33)`, `DAE103_FL2 (33)`, `DAE103_TM1 (33)`, `DAE101_FL3 (32)`,
`DAE102_FL3 (32)`, `DAE103_FL3 (32)`, `DAE103_KT3 (32)`, `DAE103_KT4 (32)`,
`DAE103_TM2 (32)`, `DAE101_CS2 (31)`, `DAE102_CS2 (31)`, `DAE103_CS2 (31)`,
`DAE106_CS2 (31)`

Too many students to combine into one room — these need an additional room or relocation.

## Key observation — DAE103 dominates

**16 of the 24 unplaced are DAE103.** This subject has a section in almost every group, so
it drives the shortage everywhere. Prioritising DAE103's room arrangement (extra rooms,
relocation, or splitting sections across days) would move the needle most.

## Realistic options (rooms fixed at current count)

1. **Combine** the 7 small classes above (needs Jo's pairings).
2. **Relocate** some large classes to a centre with spare slots (CSW/KT have headroom).
3. **Accept** the residual unplaced list and handle manually.
