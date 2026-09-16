from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import Any, Dict, Tuple

import streamlit as st

from compare_names import collect_rsvp_people, mark_matches


st.set_page_config(page_title="WGII Attendance Checker", page_icon="✅", layout="wide")


def process_uploaded_files(
    attendee_bytes: bytes,
    attendee_name: str,
    rsvp_bytes: bytes,
    rsvp_name: str,
    sheet_name: str,
) -> Tuple[bytes, Dict[str, Any]]:
    """Run name comparison on uploaded files and return marked workbook bytes and stats."""
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        attendee_path = temp_path / attendee_name
        rsvp_path = temp_path / rsvp_name
        output_path = temp_path / f"{Path(attendee_name).stem}_marked.xlsx"

        attendee_path.write_bytes(attendee_bytes)
        rsvp_path.write_bytes(rsvp_bytes)

        rsvp_people = collect_rsvp_people(rsvp_path, first_col=7, last_col=8)
        total_rows, yes_rows, maybe_rows, first_hits, last_hits, unmatched_count = mark_matches(
            source_path=attendee_path,
            output_path=output_path,
            rsvp_people=rsvp_people,
            sheet_name=sheet_name,
            status_col=12,
            maybe_names_col=13,
            review_col=14,
        )

        output_bytes = output_path.read_bytes()

    stats = {
        "processed_rows": total_rows,
        "yes_rows": yes_rows,
        "maybe_rows": maybe_rows,
        "first_hits": first_hits,
        "last_hits": last_hits,
        "unmatched_rsvp": unmatched_count,
    }
    return output_bytes, stats


def init_state() -> None:
    defaults = {
        "attendee_bytes": None,
        "attendee_name": None,
        "rsvp_bytes": None,
        "rsvp_name": None,
        "output_bytes": None,
        "output_name": None,
        "stats": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


init_state()

st.title("WGII Attendance Checker")
st.caption("Upload attendee and RSVP files, then download the marked workbook.")

sheet_name = st.text_input("Attendee sheet name", value="Full List", help="Default: Full List")

tab_attendee, tab_rsvp, tab_download = st.tabs(
    ["Upload Attendee List", "Upload LAM RSVP List", "Download Marked List"]
)

with tab_attendee:
    attendee_file = st.file_uploader(
        "Upload AR7 attendee workbook (.xlsx)",
        type=["xlsx"],
        key="attendee_uploader",
    )
    if attendee_file is not None:
        st.session_state.attendee_bytes = attendee_file.getvalue()
        st.session_state.attendee_name = attendee_file.name
        st.success(f"Loaded attendee file: {attendee_file.name}")

with tab_rsvp:
    rsvp_file = st.file_uploader(
        "Upload LAM RSVP workbook (.xlsx)",
        type=["xlsx"],
        key="rsvp_uploader",
    )
    if rsvp_file is not None:
        st.session_state.rsvp_bytes = rsvp_file.getvalue()
        st.session_state.rsvp_name = rsvp_file.name
        st.success(f"Loaded RSVP file: {rsvp_file.name}")

with tab_download:
    st.write("Step 1 status:", "ready" if st.session_state.attendee_bytes else "missing attendee file")
    st.write("Step 2 status:", "ready" if st.session_state.rsvp_bytes else "missing RSVP file")

    can_generate = bool(st.session_state.attendee_bytes and st.session_state.rsvp_bytes)

    if st.button("Generate Marked List", type="primary", disabled=not can_generate):
        try:
            output_bytes, stats = process_uploaded_files(
                attendee_bytes=st.session_state.attendee_bytes,
                attendee_name=st.session_state.attendee_name or "attendee.xlsx",
                rsvp_bytes=st.session_state.rsvp_bytes,
                rsvp_name=st.session_state.rsvp_name or "rsvp.xlsx",
                sheet_name=sheet_name,
            )
            st.session_state.output_bytes = output_bytes
            base_name = Path(st.session_state.attendee_name or "AR7_LAM3_Attendee_List").stem
            st.session_state.output_name = f"{base_name}_marked.xlsx"
            st.session_state.stats = stats
            st.success("Marked workbook generated.")
        except Exception as exc:
            st.session_state.output_bytes = None
            st.session_state.output_name = None
            st.session_state.stats = None
            st.error(f"Failed to generate marked workbook: {exc}")

    if st.session_state.stats:
        col1, col2, col3 = st.columns(3)
        col1.metric("Processed", st.session_state.stats["processed_rows"])
        col2.metric("Yes", st.session_state.stats["yes_rows"])
        col3.metric("Maybe", st.session_state.stats["maybe_rows"])

        col4, col5, col6 = st.columns(3)
        col4.metric("First Name Hits", st.session_state.stats["first_hits"])
        col5.metric("Last Name Hits", st.session_state.stats["last_hits"])
        col6.metric("Unmatched RSVP", st.session_state.stats["unmatched_rsvp"])

    if st.session_state.output_bytes:
        st.download_button(
            label="Download Marked Workbook",
            data=st.session_state.output_bytes,
            file_name=st.session_state.output_name or "AR7 LAM3 Attendee List_marked.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if st.button("Clear Uploaded Files"):
        for key in [
            "attendee_bytes",
            "attendee_name",
            "rsvp_bytes",
            "rsvp_name",
            "output_bytes",
            "output_name",
            "stats",
        ]:
            st.session_state[key] = None
        st.rerun()
