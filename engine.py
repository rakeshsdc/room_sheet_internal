"""
Room Sheet Generator - core engine
====================================
Parses:
  1. A Rooms Excel file      -> list of (room_no, capacity)
  2. A Students Word doc     -> list of course blocks, each with
                                 Sem / Course name / Code / student table

Allocates students to rooms (greedy fill, splitting a course across
rooms when it doesn't fit in one), then generates:
  - one Word "room sheet" per room, formatted like the sample template
  - one master Excel workbook summarising the whole allocation

No network access is used anywhere in this module.
"""

from __future__ import annotations

import io
import re
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm as rl_cm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table as RLTable, TableStyle, Paragraph as RLParagraph,
    Spacer, PageBreak,
)


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class Student:
    name: str
    apaar_id: str
    department: str


@dataclass
class Course:
    sem: str            # e.g. "Sem 3"
    course_name: str    # e.g. "MINOR 1- UK3DSCSGE206 DRAMA AND ALANKARA"
    code: str            # e.g. "S3M1SGK"
    students: list[Student] = field(default_factory=list)
    cursor: int = 0       # index of next un-allocated student

    @property
    def remaining(self) -> int:
        return len(self.students) - self.cursor

    @property
    def code_width(self) -> int:
        # zero-pad width for Sl.No numbering, min 2 digits
        return max(2, len(str(len(self.students))))


@dataclass
class Room:
    room_no: str
    capacity: int


@dataclass
class AllocatedChunk:
    course: Course
    start_index: int   # 0-based index into course.students
    count: int

    @property
    def students(self) -> list[Student]:
        return self.course.students[self.start_index:self.start_index + self.count]

    @property
    def start_num(self) -> int:
        return self.start_index + 1

    def sl_no(self, offset: int) -> str:
        n = self.start_num + offset
        return f"{self.course.code}{str(n).zfill(self.course.code_width)}"


@dataclass
class RoomAllocation:
    room: Room
    chunks: list[AllocatedChunk] = field(default_factory=list)

    @property
    def total_allocated(self) -> int:
        return sum(c.count for c in self.chunks)

    @property
    def by_sem_counts(self) -> dict[str, int]:
        out: dict[str, int] = {}
        for c in self.chunks:
            out[c.course.sem] = out.get(c.course.sem, 0) + c.count
        return out


# ---------------------------------------------------------------------------
# Parsing: Rooms Excel
# ---------------------------------------------------------------------------

def parse_rooms_excel(path: str | Path) -> list[Room]:
    """
    Expects a worksheet with a header row containing columns named
    (case-insensitive, whitespace-tolerant) 'Room No' and 'Capacity'.
    Any other columns are ignored. Uses the first (active) sheet.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Rooms Excel file is empty.")

    header = [str(h).strip().lower() if h is not None else "" for h in rows[0]]

    def find_col(*names: str) -> int:
        for name in names:
            for i, h in enumerate(header):
                if h == name:
                    return i
        raise ValueError(
            f"Could not find a column named any of {names} in the Rooms Excel "
            f"header row {rows[0]!r}."
        )

    room_col = find_col("room no", "room no.", "room number", "room")
    cap_col = find_col("capacity", "max capacity", "maximum capacity")

    rooms: list[Room] = []
    for r in rows[1:]:
        if r is None or r[room_col] is None:
            continue
        room_no = str(r[room_col]).strip()
        cap_raw = r[cap_col]
        if cap_raw is None:
            continue
        capacity = int(cap_raw)
        rooms.append(Room(room_no=room_no, capacity=capacity))

    if not rooms:
        raise ValueError("No valid room rows found in the Rooms Excel file.")
    return rooms


# ---------------------------------------------------------------------------
# Parsing: Students Word doc
# ---------------------------------------------------------------------------

_SEM_RE = re.compile(r"sem\s*[:\-]?\s*(.+)", re.IGNORECASE)
_COURSE_RE = re.compile(r"course\s*name\s*[:\-]?\s*(.+)", re.IGNORECASE)
_CODE_RE = re.compile(r"code\s*[:\-]?\s*(.+)", re.IGNORECASE)


def _clean(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _table_to_students(table) -> list[Student]:
    """
    Expects a header row then rows of: Sl.No | Name | APAAR ID | Department
    Column order is detected from the header text so minor re-ordering
    in the source doc doesn't break parsing.
    """
    rows = table.rows
    if not rows:
        return []
    header_cells = [_clean(c.text).lower() for c in rows[0].cells]

    def find(*keywords: str) -> int:
        for i, h in enumerate(header_cells):
            if any(k in h for k in keywords):
                return i
        return -1

    name_col = find("name")
    apaar_col = find("apaar")
    dept_col = find("department", "dept")

    students = []
    for row in rows[1:]:
        cells = [_clean(c.text) for c in row.cells]
        if not any(cells):
            continue
        name = cells[name_col] if name_col >= 0 and name_col < len(cells) else ""
        apaar = cells[apaar_col] if apaar_col >= 0 and apaar_col < len(cells) else ""
        dept = cells[dept_col] if dept_col >= 0 and dept_col < len(cells) else ""
        if not name:
            continue
        students.append(Student(name=name, apaar_id=apaar, department=dept))
    return students


def parse_students_docx(path: str | Path) -> list[Course]:
    """
    Walks the document body top-to-bottom. Paragraph text since the last
    table is buffered and scanned for 'Sem:', 'Course name:', 'Code:'
    lines (labels may be split across bold/non-bold runs, so we match on
    the concatenated paragraph text). When a table is hit, it is parsed
    as that course's student list and the buffer is reset.
    """
    document = docx.Document(str(path))
    courses: list[Course] = []

    buffer_lines: list[str] = []

    def flush_course(table) -> None:
        text_blob = "\n".join(buffer_lines)
        sem_match = _SEM_RE.search(text_blob)
        code_match = _CODE_RE.search(text_blob)
        # Course name may wrap across multiple paragraph lines; join every
        # line from the "Course name:" line up to (not including) the
        # "Code:" line.
        course_name = ""
        collecting = False
        name_parts = []
        for line in buffer_lines:
            cm = _COURSE_RE.search(line)
            if cm:
                collecting = True
                name_parts.append(cm.group(1))
                continue
            if collecting:
                if _CODE_RE.search(line):
                    break
                if line.strip():
                    name_parts.append(line.strip())
        course_name = _clean(" ".join(name_parts))

        sem = _clean(sem_match.group(1)) if sem_match else ""
        code = _clean(code_match.group(1)).replace(" ", "") if code_match else ""

        students = _table_to_students(table)
        if students:
            courses.append(Course(sem=sem, course_name=course_name, code=code, students=students))

    for child in document.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            para = docx.text.paragraph.Paragraph(child, document)
            if para.text.strip():
                buffer_lines.append(para.text)
        elif tag == "tbl":
            table = docx.table.Table(child, document)
            flush_course(table)
            buffer_lines = []
        # sectPr and anything else is ignored

    if not courses:
        raise ValueError(
            "No course blocks found in the Word document. Each block must have "
            "'Sem:', 'Course name:', 'Code:' lines followed by a student table."
        )
    return courses


def sort_course_students(courses: list[Course]) -> None:
    """
    Sorts each course's student list ascending by Department, then by Name
    within the same department. Mutates the courses in place. Call this
    right after parse_students_docx() and before allocate() -- everything
    downstream (Sl.No numbering, room allocation splits, room sheets, the
    master Excel, and the course-wise PDF) simply follows whatever order
    course.students is in, so sorting once here is all that's needed.
    """
    for course in courses:
        course.students.sort(key=lambda s: (s.department.strip().lower(), s.name.strip().lower()))


# ---------------------------------------------------------------------------
# Allocation
# ---------------------------------------------------------------------------

def _distribute_balanced(capacity: int, group_remaining: dict) -> dict:
    """
    Split `capacity` seats across the groups in `group_remaining` (a dict of
    group_key -> seats still needed) as evenly as possible ("max-min fair"),
    redistributing any leftover from a group that runs out of students to
    the other still-active groups. Mutates nothing; returns a dict of
    group_key -> seats allocated to that group in this room.
    """
    remaining = dict(group_remaining)  # local working copy
    allocated = {g: 0 for g in remaining}
    capacity_left = capacity

    active = [g for g in remaining if remaining[g] > 0]
    while capacity_left > 0 and active:
        share = capacity_left // len(active)
        if share == 0:
            # Fewer seats left than active groups: hand out one at a time,
            # in group order, until capacity runs out.
            for g in active:
                if capacity_left == 0:
                    break
                allocated[g] += 1
                remaining[g] -= 1
                capacity_left -= 1
            active = [g for g in remaining if remaining[g] > 0]
            continue

        for g in list(active):
            take = min(share, remaining[g])
            allocated[g] += take
            remaining[g] -= take
            capacity_left -= take
            if remaining[g] == 0:
                active.remove(g)
        # Any leftover (because a group had fewer students than its equal
        # share) loops back around and gets re-split among the groups that
        # are still active.

    return allocated


ORPHAN_THRESHOLD = 2  # the "+2 / -2 flexibility" the user asked for


def allocate(rooms: list[Room], courses: list[Course]) -> tuple[list[RoomAllocation], int]:
    """
    Fills rooms with, as close as possible, an EQUAL number of students from
    each semester present (e.g. Sem 3 and Sem 5), and -- within a room --
    from exactly ONE course per semester wherever possible (mixing more than
    one course of the *same* semester into a room only happens as a last
    resort, when a semester has no partner semester left to pair with).

    To avoid stranding a tiny handful of students from a course alone in a
    different room from the rest of their course-mates, a room's target
    seat count is allowed to flex by up to ORPHAN_THRESHOLD (2) seats: if a
    course would otherwise leave a leftover of 1-2 students to trail into
    the next room, those 1-2 are pulled into the current room instead so
    the whole course finishes together, even if that puts the room slightly
    over its nominal capacity.

    Returns (allocations, unallocated_student_count).
    """
    # Group courses by semester, preserving the order semesters and courses
    # first appear in the Word document.
    semester_order: list[str] = []
    semester_courses: dict[str, list[Course]] = {}
    for c in courses:
        if c.sem not in semester_courses:
            semester_courses[c.sem] = []
            semester_order.append(c.sem)
        semester_courses[c.sem].append(c)

    cursor_idx = {sem: 0 for sem in semester_order}

    def current_course(sem: str) -> Course | None:
        courses_list = semester_courses[sem]
        idx = cursor_idx[sem]
        while idx < len(courses_list) and courses_list[idx].remaining <= 0:
            idx += 1
        cursor_idx[sem] = idx
        return courses_list[idx] if idx < len(courses_list) else None

    def semester_remaining(sem: str) -> int:
        return sum(c.remaining for c in semester_courses[sem])

    allocations: list[RoomAllocation] = []

    for room in rooms:
        active = [s for s in semester_order if semester_remaining(s) > 0]
        if not active:
            break  # every student already allocated; stop, leaving later rooms unused

        ra = RoomAllocation(room=room)
        capacity = room.capacity

        if len(active) == 1:
            # Only one semester has students left -- no partner semester to
            # pair with, so (as explicitly permitted as a last resort) this
            # room may span more than one course of that same semester.
            sem = active[0]
            need = capacity
            while need > 0:
                course = current_course(sem)
                if course is None:
                    break
                take = min(need, course.remaining)
                leftover = course.remaining - take
                if 0 < leftover <= ORPHAN_THRESHOLD:
                    take = course.remaining  # finish the course rather than strand 1-2 students
                chunk = AllocatedChunk(course=course, start_index=course.cursor, count=take)
                ra.chunks.append(chunk)
                course.cursor += take
                need -= take
        else:
            # Two or more semesters active: pair up their CURRENT course
            # only (never a second course of the same semester) and split
            # the room's capacity between them as evenly as possible.
            current = {s: current_course(s) for s in active if current_course(s) is not None}
            avail = {s: current[s].remaining for s in current}
            share = _distribute_balanced(capacity, avail)

            take: dict[str, int] = {}
            for s in current:
                t = min(share.get(s, 0), avail[s])
                leftover = avail[s] - t
                if 0 < leftover <= ORPHAN_THRESHOLD:
                    t = avail[s]  # absorb the small remainder, finishing this course here
                take[s] = t

            # Rare case: absorbing a small remainder on more than one side
            # at once could push the room past its +2 tolerance. If so,
            # trim back the extra (beyond each side's natural share) from
            # whichever side has the smallest overshoot first.
            overflow = sum(take.values()) - (capacity + ORPHAN_THRESHOLD)
            if overflow > 0:
                extras = sorted(
                    ((s, take[s] - min(share.get(s, 0), avail[s])) for s in current),
                    key=lambda item: item[1],
                )
                for s, extra in extras:
                    if overflow <= 0:
                        break
                    if extra <= 0:
                        continue
                    cut = min(extra, overflow)
                    take[s] -= cut
                    overflow -= cut

            for s, course in current.items():
                t = take.get(s, 0)
                if t <= 0:
                    continue
                chunk = AllocatedChunk(course=course, start_index=course.cursor, count=t)
                ra.chunks.append(chunk)
                course.cursor += t

        allocations.append(ra)

    unallocated = sum(c.remaining for c in courses)
    return allocations, unallocated


# ---------------------------------------------------------------------------
# Word output: one room sheet per room
# ---------------------------------------------------------------------------

MIN_ROW_HEIGHT_CM = 0.8


def _set_cell_text(cell, text, bold=False, size=10, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    if align:
        p.alignment = align
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)


def _set_row_min_height(row, cm_height: float = MIN_ROW_HEIGHT_CM) -> None:
    """Force a minimum row height (Word will still grow it for wrapped text)."""
    trPr = row._tr.get_or_add_trPr()
    trHeight = trPr.makeelement(qn("w:trHeight"), {
        qn("w:val"): str(Cm(cm_height).twips),
        qn("w:hRule"): "atLeast",
    })
    trPr.append(trHeight)


def _repeat_as_header_row(row) -> None:
    """Mark a table row to repeat at the top of every page the table spans."""
    trPr = row._tr.get_or_add_trPr()
    tblHeader = trPr.makeelement(qn("w:tblHeader"), {qn("w:val"): "true"})
    trPr.append(tblHeader)


def _prevent_row_split(row) -> None:
    """Stop a single row's content being split across two pages."""
    trPr = row._tr.get_or_add_trPr()
    cantSplit = trPr.makeelement(qn("w:cantSplit"), {})
    trPr.append(cantSplit)


def _keep_with_next(paragraph) -> None:
    """Glue this paragraph to whatever follows it, so they stay on one page."""
    paragraph.paragraph_format.keep_with_next = True


def _set_fixed_column_widths(table, widths_cm: list[float]) -> None:
    """
    Switch a table to FIXED layout with explicit column widths, applied to
    both the table grid and every individual cell (Word requires both to
    reliably honor fixed widths and, importantly, to stop re-flowing the
    table in a way that can override explicit row heights).
    """
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = tblPr.makeelement(qn("w:tblLayout"), {qn("w:type"): "fixed"})
    tblPr.append(layout)

    for i, w in enumerate(widths_cm):
        table.columns[i].width = Cm(w)
    for row in table.rows:
        for i, w in enumerate(widths_cm):
            row.cells[i].width = Cm(w)


def _strip_cell_borders(cell) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.makeelement(qn("w:tcBorders"), {})
    for edge in ("top", "left", "bottom", "right"):
        el = tcPr.makeelement(qn(f"w:{edge}"), {qn("w:val"): "none"})
        borders.append(el)
    tcPr.append(borders)


def _keep_cell_content_together(cell) -> None:
    for p in cell.paragraphs:
        _keep_with_next(p)


def build_room_sheet_docx(
    college_name: str,
    exam_title: str,
    date_str: str,
    day_str: str,
    session_label: str,   # e.g. "FN 10:30 am to 11:30 am"
    allocation: RoomAllocation,
    out_path: str | Path,
) -> None:
    d = docx.Document()

    # Narrow margins so the sheet is compact like the sample
    section = d.sections[0]
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)

    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(college_name)
    run.bold = True
    run.font.size = Pt(14)

    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(exam_title)
    run.bold = True
    run.font.size = Pt(12)

    # Date/day | session | room number row
    info_table = d.add_table(rows=1, cols=3)
    cells = info_table.rows[0].cells
    _set_cell_text(cells[0], f"{date_str}  {day_str}", bold=True)
    _set_cell_text(cells[1], session_label, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(cells[2], f"Room - {allocation.room.room_no}", bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
    # strip table borders for this info row (visually just three cells of text)
    tbl = info_table._tbl
    tblPr = tbl.tblPr
    borders = tblPr.makeelement(qn("w:tblBorders"), {})
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = tblPr.makeelement(qn(f"w:{edge}"), {qn("w:val"): "none"})
        borders.append(el)
    tblPr.append(borders)

    d.add_paragraph()

    # 5 columns: Sl.No / APAAR ID / Name / Department / Signature.
    # Fixed widths matching the content width (page width minus margins),
    # so Word uses a FIXED table layout instead of AutoFit — this is what
    # makes the explicit row heights below actually stick in desktop Word.
    content_width_cm = section.page_width.cm - section.left_margin.cm - section.right_margin.cm
    col_widths = [2.6, 3.2, 5.0, 3.4, 3.2]
    scale = content_width_cm / sum(col_widths)
    col_widths = [w * scale for w in col_widths]

    last_table = None
    last_chunk_is_final = False
    for chunk_idx, chunk in enumerate(allocation.chunks):
        course = chunk.course
        heading = d.add_paragraph()
        run = heading.add_run(f"{course.sem} {course.course_name}")
        run.bold = True
        run.font.size = Pt(11)

        table = d.add_table(rows=1, cols=5)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        header_row = table.rows[0]
        hdr = header_row.cells
        for i, text in enumerate(["Sl.No.", "APAAR ID", "Name", "Department", "Signature"]):
            _set_cell_text(hdr[i], text, bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
        _repeat_as_header_row(header_row)      # header repeats on every page this table spans
        _set_row_min_height(header_row)
        _prevent_row_split(header_row)

        for offset, student in enumerate(chunk.students):
            row_cells = table.add_row().cells
            _set_cell_text(row_cells[0], chunk.sl_no(offset), size=10)
            _set_cell_text(row_cells[1], student.apaar_id, size=10)
            _set_cell_text(row_cells[2], student.name, size=10)
            _set_cell_text(row_cells[3], student.department, size=10)
            _set_cell_text(row_cells[4], "", size=10)
            data_row = table.rows[-1]
            _set_row_min_height(data_row)      # rows at least 0.8 cm tall
            _prevent_row_split(data_row)        # a row's own content never splits across pages

        is_final_chunk = chunk_idx == len(allocation.chunks) - 1
        if is_final_chunk:
            # Totals + signature are appended as extra rows of THIS SAME
            # table (merged across all 5 columns) rather than a separate
            # table. Word keeps rows together far more reliably within one
            # table than across a table -> paragraph -> table boundary, so
            # this is what actually stops the totals line from being
            # orphaned alone at the top of a new page.
            total = allocation.total_allocated
            if len(allocation.chunks) > 1:
                parts = ",".join(
                    f"{ch.course.sem}-{ch.course.code}={ch.count}" for ch in allocation.chunks
                )
                total_str = f"Total {total} ({parts})"
            else:
                total_str = f"Total - {total}"

            totals_row = table.add_row()
            totals_cell = totals_row.cells[0]
            for c in totals_row.cells[1:]:
                totals_cell = totals_cell.merge(c)
            totals_cell.text = ""
            tp = totals_cell.paragraphs[0]
            tp.paragraph_format.tab_stops.add_tab_stop(Cm(content_width_cm * 0.45), WD_TAB_ALIGNMENT.LEFT)
            tp.paragraph_format.tab_stops.add_tab_stop(Cm(content_width_cm * 0.85), WD_TAB_ALIGNMENT.LEFT)
            run = tp.add_run(f"{total_str}\tPresent -\tAbsent -")
            run.bold = True
            run.font.size = Pt(10)
            _strip_cell_borders(totals_cell)
            _set_row_min_height(totals_row)
            _prevent_row_split(totals_row)

            sig_row = table.add_row()
            sig_cell = sig_row.cells[0]
            for c in sig_row.cells[1:]:
                sig_cell = sig_cell.merge(c)
            sig_cell.text = ""
            sp = sig_cell.paragraphs[0]
            sp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = sp.add_run("Signature(s) of Invigilator(s)")
            run.bold = True
            run.font.size = Pt(10)
            _strip_cell_borders(sig_cell)
            _set_row_min_height(sig_row)
            _prevent_row_split(sig_row)

            # Glue the last 3 student rows + totals row to each other (the
            # signature row, being last, needs nothing after it to stick to).
            rows = table.rows
            glue_rows = rows[-min(4, len(rows) - 1):-1]  # last few data/totals rows, excluding sig row
            for r in glue_rows:
                for cell in r.cells:
                    _keep_cell_content_together(cell)
        else:
            d.add_paragraph()

        _set_fixed_column_widths(table, col_widths)
        last_table = table

    d.save(str(out_path))


# ---------------------------------------------------------------------------
# Excel output: master allocation report
# ---------------------------------------------------------------------------

def build_master_excel(
    allocations: list[RoomAllocation],
    unallocated: int,
    out_path: str | Path,
) -> None:
    wb = openpyxl.Workbook()

    bold = Font(bold=True)
    center = Alignment(horizontal="center")
    thin = Side(style="thin")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # --- Sheet 1: Room summary ---
    ws1 = wb.active
    ws1.title = "Room Summary"
    headers = ["Room No", "Capacity", "Allocated", "Free Seats", "Courses (Sem - Code - Course - Count)"]
    ws1.append(headers)
    for c in ws1[1]:
        c.font = bold
        c.alignment = center
        c.border = border

    for ra in allocations:
        course_desc = "; ".join(
            f"{ch.course.sem} - {ch.course.code} - {ch.course.course_name} ({ch.count})"
            for ch in ra.chunks
        )
        row = [
            ra.room.room_no,
            ra.room.capacity,
            ra.total_allocated,
            ra.room.capacity - ra.total_allocated,
            course_desc,
        ]
        ws1.append(row)

    for col, width in zip("ABCDE", [12, 12, 12, 12, 90]):
        ws1.column_dimensions[col].width = width

    # --- Sheet 2: Full student allocation ---
    ws2 = wb.create_sheet("Student Allocation")
    headers2 = ["Room No", "Sl.No", "Sem", "Course Code", "Course Name", "Name", "APAAR ID", "Department"]
    ws2.append(headers2)
    for c in ws2[1]:
        c.font = bold
        c.alignment = center
        c.border = border

    for ra in allocations:
        for chunk in ra.chunks:
            for offset, student in enumerate(chunk.students):
                ws2.append([
                    ra.room.room_no,
                    chunk.sl_no(offset),
                    chunk.course.sem,
                    chunk.course.code,
                    chunk.course.course_name,
                    student.name,
                    student.apaar_id,
                    student.department,
                ])

    for col, width in zip("ABCDEFGH", [10, 12, 10, 14, 45, 25, 16, 20]):
        ws2.column_dimensions[col].width = width

    # --- Sheet 3: Warnings ---
    if unallocated > 0:
        ws3 = wb.create_sheet("Warnings")
        ws3.append(["Warning"])
        ws3["A1"].font = bold
        ws3.append([
            f"{unallocated} student(s) could not be allocated - total room capacity "
            f"was less than the total number of students. Add more rooms or increase capacity."
        ])
        ws3.column_dimensions["A"].width = 100

    wb.save(str(out_path))


# ---------------------------------------------------------------------------
# PDF output: single consolidated course-wise room list (for students)
# ---------------------------------------------------------------------------

def build_course_wise_room_list_pdf(
    college_name: str,
    exam_title: str,
    date_str: str,
    day_str: str,
    session_label: str,
    courses: list[Course],
    allocations: list[RoomAllocation],
    out_path: str | Path,
) -> None:
    """
    One PDF, one course per page (a course that needs more than a page just
    continues, with the column header repeated): Sl.No / Name / APAAR ID /
    Department / Room No, so a student can look up their own room number
    without needing to see the room-wise sheets at all.
    """
    # Work out which room each individual student ended up in. Built from
    # `allocations` (not `courses`) so this works correctly even when
    # `courses` is a filtered subset (e.g. only one semester's courses) --
    # a defaultdict means chunks belonging to courses outside that subset
    # are simply harmless extra entries, never a KeyError.
    room_of: dict[int, dict[int, str]] = defaultdict(dict)
    for ra in allocations:
        for chunk in ra.chunks:
            for offset in range(chunk.count):
                idx = chunk.start_index + offset
                room_of[id(chunk.course)][idx] = ra.room.room_no

    styles = {
        "college": ParagraphStyle("college", fontName="Helvetica-Bold", fontSize=14,
                                   alignment=TA_CENTER, spaceAfter=4),
        "exam": ParagraphStyle("exam", fontName="Helvetica-Bold", fontSize=11,
                                alignment=TA_CENTER, spaceAfter=4),
        "info": ParagraphStyle("info", fontName="Helvetica-Bold", fontSize=9,
                                alignment=TA_LEFT, spaceAfter=10),
        "course": ParagraphStyle("course", fontName="Helvetica-Bold", fontSize=10,
                                  alignment=TA_LEFT, spaceAfter=6),
        "cell": ParagraphStyle("cell", fontName="Helvetica", fontSize=9, leading=11),
        "cell_center": ParagraphStyle("cell_center", fontName="Helvetica", fontSize=9,
                                       leading=11, alignment=TA_CENTER),
        "cell_hdr": ParagraphStyle("cell_hdr", fontName="Helvetica-Bold", fontSize=9,
                                    leading=11, alignment=TA_CENTER),
    }

    doc = SimpleDocTemplate(
        str(out_path), pagesize=A4,
        leftMargin=1.8 * rl_cm, rightMargin=1.8 * rl_cm,
        topMargin=1.5 * rl_cm, bottomMargin=1.5 * rl_cm,
    )

    content_width = A4[0] - doc.leftMargin - doc.rightMargin
    col_widths = [w / 17.4 * content_width for w in (2.6, 3.2, 5.4, 3.4, 2.8)]

    story = []
    for i, course in enumerate(courses):
        if i > 0:
            story.append(PageBreak())
        story.append(RLParagraph(college_name, styles["college"]))
        story.append(RLParagraph(exam_title, styles["exam"]))
        story.append(RLParagraph(f"{date_str}  {day_str}  &nbsp;&nbsp;&nbsp;  {session_label}", styles["info"]))
        story.append(RLParagraph(f"{course.sem}  {course.course_name}", styles["course"]))

        header = ["Sl.No.", "APAAR ID", "Name", "Department", "Room No."]
        data = [[RLParagraph(h, styles["cell_hdr"]) for h in header]]
        room_map = room_of[id(course)]
        for idx, student in enumerate(course.students):
            sl_no = f"{course.code}{str(idx + 1).zfill(course.code_width)}"
            room_no = room_map.get(idx, "Not allotted")
            data.append([
                RLParagraph(sl_no, styles["cell_center"]),
                RLParagraph(student.apaar_id, styles["cell_center"]),
                RLParagraph(student.name, styles["cell"]),
                RLParagraph(student.department, styles["cell"]),
                RLParagraph(room_no, styles["cell_center"]),
            ])

        table = RLTable(data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.75, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(table)

    doc.build(story)


def _semester_short_label(sem: str) -> str:
    """'Sem 3' -> 'S3'; falls back to a filename-safe version of the label
    itself if no number is found in it."""
    m = re.search(r"(\d+)", sem)
    if m:
        return f"S{m.group(1)}"
    return _sanitize_filename_part(sem)


def _sanitize_filename_part(text: str) -> str:
    text = re.sub(r"\s+", "_", text.strip())
    text = re.sub(r"[^A-Za-z0-9_\-]", "", text)
    return text or "file"


def build_course_wise_room_list_pdfs_by_semester(
    college_name: str,
    exam_title: str,
    date_str: str,
    day_str: str,
    session_label: str,
    courses: list[Course],
    allocations: list[RoomAllocation],
    out_dir: str | Path,
) -> list[Path]:
    """
    Same content as build_course_wise_room_list_pdf(), but split into one
    PDF per semester (e.g. all Sem 3 courses in one file, all Sem 5 courses
    in another), named "<SemesterShortLabel>_<date>_<session>.pdf" -- e.g.
    S3_23-09-2026_FN.pdf, S5_23-09-2026_FN.pdf. Semesters are emitted in
    the order they first appear in the Word document. Returns the list of
    file paths written, in that same order.
    """
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    semester_order: list[str] = []
    grouped: dict[str, list[Course]] = {}
    for c in courses:
        if c.sem not in grouped:
            grouped[c.sem] = []
            semester_order.append(c.sem)
        grouped[c.sem].append(c)

    date_part = _sanitize_filename_part(date_str)
    # Filenames use just the leading token of the session label (e.g. "FN"
    # out of "FN 10:30 am to 11:30 am") -- the full label still appears in
    # the PDF's own heading text.
    session_first_word = session_label.strip().split()[0] if session_label.strip() else "session"
    session_part = _sanitize_filename_part(session_first_word)

    paths: list[Path] = []
    for sem in semester_order:
        short = _semester_short_label(sem)
        fname = f"{short}_{date_part}_{session_part}.pdf"
        out_path = out_dir / fname
        build_course_wise_room_list_pdf(
            college_name=college_name,
            exam_title=exam_title,
            date_str=date_str,
            day_str=day_str,
            session_label=session_label,
            courses=grouped[sem],
            allocations=allocations,
            out_path=out_path,
        )
        paths.append(out_path)

    return paths
