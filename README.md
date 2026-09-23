# Examination Room Sheet Generator

Generates one exam room sheet (Word) per room, plus a master allocation
report (Excel), from:

1. A **Rooms Excel** file — Room No + Capacity
2. A **Students Word doc** — one block per course (Sem / Course name /
   Code / student table)

It fills each room to capacity in order, splitting a course across
rooms when it doesn't fully fit, and continuing that course's serial
numbers into the next room.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually http://localhost:8501).

## Deploy for free on Streamlit Community Cloud

1. Create a new **public or private GitHub repo** and push these three
   files (`app.py`, `engine.py`, `requirements.txt`) plus this README.
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with
   GitHub, click **New app**.
3. Pick your repo, branch `main`, and main file path `app.py`.
4. Click **Deploy**. You'll get a permanent `https://<something>.streamlit.app`
   link you can bookmark and reuse every exam session — no reinstalling
   anything.

Any time you `git push` an update, the deployed app updates itself
automatically within a minute or two.

## Input file formats

### Rooms Excel (`rooms.xlsx`)

One sheet, header row required:

| Room No | Capacity |
|---------|----------|
| 24      | 32       |
| 25      | 32       |
| 26      | 60       |

- Column names are matched case-insensitively (`Room No`, `Room No.`,
  `Room Number`, `Room` all work; same for `Capacity` / `Max Capacity`).
- Capacity is the literal number of students/seats in that room.
- Rooms are filled **in the order they appear** in this sheet, so put
  them in whatever priority order you want used first.

### Students Word doc (`students.docx`)

Repeat this block for every course, in the order you want them
allocated (courses are drawn from top to bottom, filling rooms as it
goes):

```
Sem: Sem 3

Course name: MINOR 1- UK3DSCSGE206 DRAMA AND ALANKARA

Code: S3M1SGK

| Sl. No | Name Of The Student | APAAR ID     | Department |
|--------|----------------------|--------------|------------|
| 1      | Anagha B Shenoy      | 999881648202 | Hindi      |
| 2      | Anugraha Gireesh     | 786054768564 | English    |
| ...    | ...                  | ...          | ...        |
```

- The `Sem:`, `Course name:`, `Code:` lines can be bold/plain, and
  `Course name:` can wrap across more than one line — the parser
  re-joins it automatically.
- The table's header row can be in any column order; it's matched by
  keyword (`name`, `apaar`, `department`/`dept`). The **original
  Sl.No column in the Word doc is ignored** — the generator always
  renumbers using `Code` + a running number, continuing across rooms
  if a course is split.
- `Code` becomes the room-sheet Sl.No prefix (e.g. `S3M1SGK01`,
  `S3M1SGK02`, ...). Make sure every course has a unique code.

## What you get back

- `Room_<no>.docx` for every room that received at least one student —
  formatted like the original sample: college name, exam title,
  date/day/session/room row, one heading + table per course present in
  that room, and a totals/signature footer.
- `Master_Allocation.xlsx` with:
  - **Room Summary** — capacity, allocated count, free seats, and which
    courses/counts are in each room
  - **Student Allocation** — every student with their final room and
    Sl.No, one row each (useful for records / verifying counts)
  - **Warnings** (only if present) — flags if total capacity was less
    than the total number of students, so some students couldn't be
    seated

## Notes / limitations of this first version

- The allocator does a simple **greedy fill in the given order** — it
  does not try to optimize for "avoid leaving 1-2 empty seats" or
  "always pair exactly one Sem X with one Sem Y." If you want a
  specific pairing priority (e.g. always try to pair two different
  semesters before mixing two courses from the same semester), that's
  a small change to `allocate()` in `engine.py` — flag it and it can be
  added.
- Bench-level seating (who sits next to whom) is intentionally **not**
  encoded in the sheet — the invigilator arranges seating from the
  printed list, same as your original process.
- All processing happens locally in the browser session / Streamlit
  server; no student data is sent anywhere else.
