# Working with the AI workbuddy on the DAE timetable

A short guide for Jo — 2026-08-18

## Before you start

You asked for a prompt you could use to arrange the timetable yourself. Here is
my answer, and it follows the same line I showed in the demo rather than a
single big prompt.

You may remember the split from then. Organising requirements and checking work
is where AI is genuinely strong, so that is what we give it. The arranging
itself goes to the small tool, because a timetable has to be accurate, has to be
redone every term, and cannot be wrong — and the tool gives the same answer from
the same input every time. You drive it by editing the input Excel. That has not
changed, and this note is the everyday version of it.

A single "arrange the T2026C timetable for me" prompt goes back the other way,
and it is worth being clear about why. Our timetable is not really a scheduling
puzzle. It is a series of judgement calls that happen to be written down in a
spreadsheet — nine centres, combine classes, Net teacher hours, travel between
campuses, room capacity that cannot be changed, and a dozen things that are true
only because you know them. No prompt carries all of that, so what comes back is
a complete-looking timetable built on guesses it never mentions.

The same files this week show what the other approach gets you. Every one of the
300 placements was read off your six day sheets, expanded into 4,494 daily rows,
and six genuine clashes came out — including two classrooms that are simply
blank — all of which you would otherwise have found after the timetable went
out. None of that needed judgement. All of it takes a person hours and is easy
to get slightly wrong.

So: you keep every decision, and hand over everything else. Five minutes to
read, and the prompts in the middle are meant to be copied straight into the
chat.

## The short version

1. **Don't ask it to build the timetable. Ask it to check, convert and explain.**
   You keep the judgement calls; it does the mechanical work.
2. **Always ask what it assumed.** An answer with no assumptions listed is the
   dangerous kind.
3. **Tell it never to relax a rule quietly.** Say so explicitly, every time.

## The tool and the chat do different jobs

Two things sit in front of you and they are not interchangeable.

**The small tool** takes your workbook and produces the files — the Day, Time
and Location for the Classes sheet, the daily timetable, the class view, the
teacher view. Same input, same output, every time. Run it whenever you have
changed the grid. It is not clever and that is exactly the point: you can hand
its output to a centre without wondering whether it improvised.

**The chat** is for everything the tool does not cover — checking a workbook
before you send it, asking what a change would break, working through one
awkward clash, or getting something into a format nobody has asked for before.
It is flexible, which is also why it should not be the thing producing your
final timetable.

A simple rule: **if you want a file, use the tool. If you want an answer, use the
chat.**

And in both cases the boundary is the same one from the demo — you decide where
classes go; everything else can be handed over, as long as it tells you what it
noticed on the way.

## The loop

```
   1  Give it the file and say what you want
                ↓
   2  Read the assumptions it lists  ←── this step is the whole point
                ↓
   3  Correct the assumptions, not the output
                ↓
   4  Ask it to run again
```

Step 3 matters. If a room is wrong in the output, don't ask it to change that
cell — find out *why* it chose that room, fix the input, and re-run. Otherwise
the same mistake comes back next time.

## Paste this at the top of any timetable request

These are the standing rules of DAE scheduling. The workbuddy has no way to know
them, and it will invent something plausible if you don't say:

> Context for the DAE timetable:
> - Only the `Mon(Term1)` … `Sat(Term1)` sheets belong to T2026C. The `Term2`
>   sheets are the next academic term — ignore them.
> - DAE classes run Monday, Tuesday and Thursday. Police cadet classes run
>   Monday, Wednesday and Friday.
> - A 4-hour subject is two sessions in a day. DAE103 Maths and DAE260 are
>   2 hours, so one session. Use the Loading column, not the subject code.
> - A term is 15 lesson days, so 30 topics for a 4-hour subject.
> - Combine classes share one room and one teacher. They sit on consecutive rows
>   with the classroom left blank on all but the first.
> - English teachers may travel at most 30 minutes between centres on the same
>   day.
> - Each class needs one Net teacher block per term. CS1 and CS2 need only one
>   for the whole year.
> - Classroom codes: CSW-C1 to C6 each keep their own code on the daily
>   timetable.

## Prompts you can copy

**Check my grid before I send it out**

> Please check this workbook for clashes before I circulate it. Look for three
> things: the same room used at the same time by two different teachers, one
> teacher in two rooms at once, and one class with two subjects at once. Treat
> "same room, same teacher" as a combine class and ignore those. Give me the
> sheet name and the cell for every problem you find. If you find nothing, tell
> me exactly what you checked.

**Redo the outputs after I've edited the grid**

> Here is the updated weekly workbook and the DAE calendar. Please redo the four
> outputs: Day/Time/Location for the Classes sheet, the daily timetable, the
> class view and the teacher view. Take my placement as given — do not
> re-schedule anything. List everything you had to assume in a separate tab.

**Try a change before committing to it**

> If I move DAE102_KT1 from Monday morning to Tuesday morning, what breaks?
> Check the room capacity at KT, whether the teacher is free at that time, and
> whether the English Net arrangement still works. Don't change anything yet —
> just tell me what the consequences are.

**Work on one specific problem**

> On `Fri(Term1)`, Mr. Philip Chan is booked at SSP-303 and SSP-201 at the same
> time — rows 30, 33 and 34. Show me which other time slots are free for
> DAE229_A in the same room, and which other teachers could take it instead.
> Don't pick one for me; show me the options with the trade-offs.

**Add classes when a new file arrives**

> Here is the police college file. Please add DAE256_G/H, DAE258_G/H and
> DAE260_G/H to the Term 1 grid and regenerate everything. Tell me if any of
> them clash with what is already placed, and don't move anything existing to
> make room — tell me instead.

**Re-run the English Net arrangement with different numbers**

> Redo the English Net teacher assignment with Ms. Kaur at 680 hours instead of
> 720, keeping the 30-minute travel limit. If some classes then cannot get a Net
> block, list them and explain why. Do not relax the travel rule to make the
> numbers come out even.

## Three phrases that change the answer

| Say this | Because otherwise |
|---|---|
| "Take my placement as given — do not re-schedule anything." | It will quietly produce its own timetable and you won't notice until later |
| "List everything you assumed." | The guesses stay invisible |
| "If you cannot satisfy this rule, tell me — do not relax it to make it fit." | It will bend a rule to give you a tidy answer |

The third is the important one. On the English Net arrangement it is the
difference between "here is your complete timetable" and "here is your
timetable, and these two classes cannot get a Net teacher because of the travel
limit — you need to decide."

## Leave these to yourself, not the AI

- Which class moves when there is a clash
- Whether a teacher will travel further than the rule allows
- Whether to accept a room that is over capacity
- Whether a class goes without a Net teacher
- Anything that has already been told to students

For all of these the right prompt ends with *"show me the options"*, not
*"fix it"*.

## If something looks wrong

Ask it to explain rather than to correct:

> Where did you get the classroom for DAE101_ST1 on Tuesday? Show me the cell
> you read it from.

You will usually find the input is wrong rather than the AI — and then you have
fixed it for good instead of patching one cell.
