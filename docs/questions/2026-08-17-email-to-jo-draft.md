# Draft email to Jo — 2026-08-17

**Re:** T2026C Daily Timetable + English Net Teacher arrangement — a few points to confirm

---

Dear Jo,

Thank you for the detailed brief and the three files. I have gone through
`T2026C_WeeklyTimetable_V4 0817.xlsx`, `HKIT DAE Calendar_V1.xlsx` and the
`DailyTimetable_T2025C` sample, and I can see how the four tasks fit together.

Before I start building this into the AI workbuddy, may I confirm a few points?
Most of them I have already made an assumption on — if the assumption is right,
a simple "yes" is enough and I will proceed.

**A. Task 1 — filling Day / Time / Location in the `Classes` sheet**

1. **Source of the information.** My assumption is that you have already placed
   every class on the `Mon(Term1)` … `Sat(Term1)` sheets, and that Task 1 is to
   read those sheets and copy the result back into the `Classes` sheet — i.e. I
   should *not* re-schedule anything. Is that correct?

2. **CSW room mapping.** On the `Classroom Code` sheet, `CSW - C1` through
   `CSW - C6` all map to `CSW - C1` in the "Classroom in daily timetable"
   column. But the 2025 sample daily timetable does show `CSW - C2`, `CSW - C3`
   and so on. Should each CSW room keep its own code, or is `CSW - C1` really
   the correct single name for all six?

3. **Seven classes not yet on the Term 1 grid.** These codes appear in the
   `Classes` sheet but I cannot find them on any Term 1 day sheet:
   `DAE256_G`, `DAE256_H`, `DAE258_G`, `DAE258_H`, `DAE260_G`, `DAE260_H`,
   `DAE270`. Are they scheduled somewhere else, or still to be arranged?

4. **27 blank classroom cells.** In a few places a class sits on a row with no
   room name in column A — for example `DAE106_TS3` and `DAE106_TS4` on Monday,
   which appear just underneath the `TM - TM2` row but look as though they
   belong at `TS - TS`. My plan is to infer the room from the class code
   (e.g. `..._TS3` → `TS - TS`) and give you a list of every cell I inferred, so
   you can check them. Is that acceptable, or would you prefer to fill them in
   yourself first?

**B. Task 2 — the daily timetable export**

5. **Number of sessions.** For a 4-hour subject I am producing 15 lesson days
   × 2 sessions = **30 rows** per class, and for `DAE103` (Maths) and `DAE260`
   15 rows. In point 4 of your email you mention "T2026C would have 15
   session" for English — could you confirm that means 15 lesson *days*
   (30 sessions), and not 15 sessions in total?

6. **Programme label.** Should the Programme column read `DAE - FT2026C`
   (following `DAE - FT2025C` in last year's sample)?

**C. Task 3 — English Net teacher arrangement**

7. **Net teacher capacity.** I did a rough check before starting. If one block
   is 5 weeks × 4 hours = 20 hours, then covering all 28 English classes across
   both terms needs 28 × 2 × 20 = **1,120 Net hours**. Against that, the
   `TeachingLoad` sheet shows 360 hours budgeted for Mr. Peter Barrett and up to
   720 for Ms. Loveleen Kaur — 1,080 hours in total, so we would be roughly two
   blocks short. This matches the KT1 / WT2 gap we saw in the earlier run.

   Would you prefer that I: (a) leave those one or two classes without a Net
   block and flag them, (b) relax the 30-minute travel limit between centres so
   the two Net teachers can cover more, or (c) plan on the basis that additional
   Net hours will be approved? Also, please confirm the 20-hours-per-block
   figure — if a block is counted differently the whole calculation changes.

Once I have your answers I will build the export and send you a first version
for checking. Tasks 1, 2 and 4 are straightforward data processing, so those
should come back quickly; Task 3 will take a little longer as it needs the
scheduling logic.

Please let me know if I have misunderstood anything.

Best regards,
Steven
