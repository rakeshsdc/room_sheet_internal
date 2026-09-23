"""
Streamlit UI for the Examination Room Sheet Generator.

Run locally with:  streamlit run app.py
Deploy for free on Streamlit Community Cloud by pointing it at this
file in your GitHub repo.
"""

import io
import zipfile
from pathlib import Path
import tempfile

import streamlit as st

from engine import (
    parse_rooms_excel,
    parse_students_docx,
    allocate,
    build_room_sheet_docx,
    build_master_excel,
)

st.set_page_config(page_title="Exam Room Sheet Generator", page_icon="📝", layout="centered")

st.title("📝 Examination Room Sheet Generator")
st.caption(
    "Upload the rooms list and the course-wise student list, fill in the exam "
    "details, and generate one room sheet (Word) per room plus a master "
    "allocation report (Excel)."
)

with st.sidebar:
    st.header("1. Exam details")
    college_name = st.text_input("College name", "Sanatana Dharma College Alappuzha")
    exam_title = st.text_input("Exam title", "FYUGP Internal Examination September 2026")
    date_str = st.text_input("Date (DD-MM-YYYY)", "23-09-2026")
    day_str = st.text_input("Day", "Wednesday")
    session_label = st.text_input("Session / time", "FN 10:30 am to 11:30 am")

st.header("2. Upload files")
col1, col2 = st.columns(2)
with col1:
    rooms_file = st.file_uploader("Rooms Excel (Room No, Capacity)", type=["xlsx"])
with col2:
    students_file = st.file_uploader(
        "Students Word doc (Sem / Course name / Code + table per course)",
        type=["docx"],
    )

st.markdown(
    "**Rooms Excel format:** a header row with columns `Room No` and "
    "`Capacity`, one room per row.\n\n"
    "**Students Word doc format:** for each course, a block of text with "
    "`Sem: <...>`, `Course name: <...>`, `Code: <...>` lines, followed "
    "immediately by a table with columns Sl.No / Name Of The Student / "
    "APAAR ID / Department."
)

generate = st.button("Generate room sheets", type="primary", disabled=not (rooms_file and students_file))

if generate:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        rooms_path = tmp_path / "rooms.xlsx"
        students_path = tmp_path / "students.docx"
        rooms_path.write_bytes(rooms_file.getvalue())
        students_path.write_bytes(students_file.getvalue())

        try:
            rooms = parse_rooms_excel(rooms_path)
            courses = parse_students_docx(students_path)
        except ValueError as e:
            st.error(str(e))
            st.stop()

        allocations, unallocated = allocate(rooms, courses)

        if unallocated > 0:
            st.warning(
                f"⚠️ {unallocated} student(s) could not be allocated — total room "
                f"capacity is less than the total number of students. Add more "
                f"rooms, or increase capacity, and try again. The files below "
                f"only cover the students that WERE allocated."
            )
        else:
            st.success("All students allocated successfully.")

        # Build outputs in-memory
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        room_files = []
        for ra in allocations:
            if ra.total_allocated == 0:
                continue
            fname = f"Room_{ra.room.room_no}.docx"
            fpath = out_dir / fname
            build_room_sheet_docx(
                college_name=college_name,
                exam_title=exam_title,
                date_str=date_str,
                day_str=day_str,
                session_label=session_label,
                allocation=ra,
                out_path=fpath,
            )
            room_files.append(fpath)

        master_path = out_dir / "Master_Allocation.xlsx"
        build_master_excel(allocations, unallocated, master_path)

        # Zip the room sheets together for one-click download
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in room_files:
                zf.write(f, arcname=f.name)
        zip_buffer.seek(0)

        st.header("3. Download")
        st.download_button(
            "⬇️ Download all room sheets (.zip of .docx)",
            data=zip_buffer,
            file_name="room_sheets.zip",
            mime="application/zip",
        )
        st.download_button(
            "⬇️ Download master allocation report (.xlsx)",
            data=master_path.read_bytes(),
            file_name="Master_Allocation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("Allocation summary")
        table_rows = [
            {
                "Room No": ra.room.room_no,
                "Capacity": ra.room.capacity,
                "Allocated": ra.total_allocated,
                "Free Seats": ra.room.capacity - ra.total_allocated,
                "Courses": ", ".join(
                    f"{c.course.sem}-{c.course.code} ({c.count})" for c in ra.chunks
                ),
            }
            for ra in allocations
            if ra.total_allocated > 0
        ]
        st.dataframe(table_rows, use_container_width=True)
