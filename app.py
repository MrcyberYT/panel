import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import io

st.set_page_config(page_title="Student Attendance & Behavior Tracker", layout="wide")

# Default configurations
PERIODS = ["P1", "P2", "P3", "P4", "Elective", "P5"]
ATTENDANCE_OPTIONS = ["", "Present", "Absent", "Late"]

# Initialize Session State
if "behaviors" not in st.session_state:
    st.session_state.behaviors = ["Excellent", "Good", "On Task", "Warning - Talking", "Disruptive", "Needs Improvement"]

if "classes" not in st.session_state:
    st.session_state.classes = {
        "Class 10A": ["Alice Smith", "Bob Jones", "Charlie Brown", "Diana Prince"],
        "Class 11B": ["Ethan Hunt", "Fiona Gallagher", "George Clark", "Hannah Abbott"],
        "Elective Art": ["Alice Smith", "George Clark", "Ian Malcolm", "Julia Roberts"]
    }

if "data" not in st.session_state:
    st.session_state.data = {}
    for cls, students in st.session_state.classes.items():
        st.session_state.data[cls] = []
        for idx, name in enumerate(students, start=1):
            row = {"Student ID": f"STU-{idx:03d}", "Student Name": name}
            for p in PERIODS:
                row[f"{p} Att"] = ""
                row[f"{p} Behav"] = ""
            row["Notes"] = ""
            st.session_state.data[cls].append(row)

st.title("🍎 Student Attendance & Behavior Tracker")

# Sidebar Controls
st.sidebar.header("Class & Behavior Controls")

# Manage Classes
new_class = st.sidebar.text_input("Add New Class Name")
if st.sidebar.button("Add Class") and new_class:
    if new_class not in st.session_state.classes:
        st.session_state.classes[new_class] = []
        st.session_state.data[new_class] = []
        st.rerun()

# Select Active Class
active_class = st.sidebar.selectbox("Select Class to View/Edit", list(st.session_state.classes.keys()))

# Manage Students in Active Class
st.sidebar.subheader(f"Students in {active_class}")
new_student = st.sidebar.text_input("Add New Student Name")
if st.sidebar.button("Add Student") and new_student:
    idx = len(st.session_state.data[active_class]) + 1
    new_row = {"Student ID": f"STU-{idx:03d}", "Student Name": new_student}
    for p in PERIODS:
        new_row[f"{p} Att"] = ""
        new_row[f"{p} Behav"] = ""
    new_row["Notes"] = ""
    st.session_state.data[active_class].append(new_row)
    st.rerun()

# Manage Behaviors
st.sidebar.subheader("Behavior Options")
new_behavior = st.sidebar.text_input("Add Custom Behavior Choice")
if st.sidebar.button("Add Behavior") and new_behavior:
    if new_behavior not in st.session_state.behaviors:
        st.session_state.behaviors.append(new_behavior)
        st.rerun()

# Main Interactive Data Table
st.subheader(f"Editing: {active_class}")

if active_class in st.session_state.data:
    df = pd.DataFrame(st.session_state.data[active_class])

    # Configure drop-down options for columns
    column_config = {
        "Student ID": st.column_config.TextColumn(disabled=True),
        "Student Name": st.column_config.TextColumn(disabled=True)
    }
    for p in PERIODS:
        column_config[f"{p} Att"] = st.column_config.SelectboxColumn(options=ATTENDANCE_OPTIONS)
        column_config[f"{p} Behav"] = st.column_config.SelectboxColumn(options=[""] + st.session_state.behaviors)

    edited_df = st.data_editor(df, column_config=column_config, num_rows="dynamic", use_container_width=True)
    st.session_state.data[active_class] = edited_df.to_dict("records")

# Export to Formatted Excel Function
def generate_excel():
    wb = openpyxl.Workbook()
    first = True
    
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    center_align = Alignment(horizontal="center", vertical="center")

    for cls_name, records in st.session_state.data.items():
        ws = wb.active if first else wb.create_sheet(title=cls_name)
        ws.title = cls_name
        first = False

        if records:
            headers = list(records[0].keys())
            ws.append(headers)

            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align

            for row_idx, row_data in enumerate(records, start=2):
                row_vals = list(row_data.values())
                ws.append(row_vals)
                for col_idx in range(1, len(headers) + 1):
                    ws.cell(row=row_idx, column=col_idx).border = thin_border

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

st.sidebar.markdown("---")
excel_data = generate_excel()
st.sidebar.download_button(
    label="📥 Download Excel File (.xlsx)",
    data=excel_data,
    file_name="Student_Attendance_and_Behavior_Tracker.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)