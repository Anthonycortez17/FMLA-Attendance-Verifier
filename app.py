import io
import re
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="HRIS Leave Claim Automation",
    page_icon="📊",
    layout="wide"
)

# Custom Styling
st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    .stButton>button {
        width: 100%;
        background-color: #0f172a;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        height: 3em;
    }
    .stButton>button:hover {
        background-color: #1e293b;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 HRIS FMLA & Leave Claim Automation")
st.caption("Paste the combined Unum Email Request and SAP Attendance Table below to generate an audit report and email reply.")

# Main Input Section
text_input = st.text_area(
    "Paste Combined Unum Request + SAP Text Here:",
    height=250,
    placeholder="Employee Name: Alex I Contreras\nEmployee ID#: 06425427\n...\n\n| Pers.No.|Employee/app.name..."
)

def parse_sap_pipe_table(text):
    lines = text.split('\n')
    data_rows = []
    for line in lines:
        line_str = line.strip()
        if line_str.startswith('|') and not line_str.startswith('|*') and 'Pers.No.' not in line_str:
            parts = [p.strip() for p in line_str.split('|')[1:-1]]
            if len(parts) >= 11:
                data_rows.append(parts[:11])
    if not data_rows:
        return None
    headers = ['Pers.No.', 'Name', 'Period', 'Date', 'TmType', 'TimeTyText', 'Number', 'Cost Ctr', 'PSubarea', 'Subarea', 'Cost Ctr Ref']
    df = pd.DataFrame(data_rows, columns=headers)
    df['Number'] = pd.to_numeric(df['Number'].str.replace(',', ''), errors='coerce')
    return df

def process_combined_text(raw_text):
    # 1. Parse Unum Email
    emp_name_match = re.search(r"Employee Name:\s*(.+)", raw_text)
    emp_id_match = re.search(r"Employee ID#:\s*(\d+)", raw_text)
    leave_num_match = re.search(r"Leave Number:\s*(\d+)", raw_text)
    dates_match = re.search(r"(\d{1,2}/\d{1,2}/\d{2,4})\s+through\s+(\d{1,2}/\d{1,2}/\d{2,4})", raw_text)

    emp_name = emp_name_match.group(1).strip() if emp_name_match else "Employee"
    emp_id = emp_id_match.group(1).strip() if emp_id_match else "000000"
    leave_num = leave_num_match.group(1).strip() if leave_num_match else "N/A"
    
    start_date_str = dates_match.group(1) if dates_match else "07/24/2025"
    end_date_str = dates_match.group(2) if dates_match else "07/23/2026"

    clean_emp_id = str(emp_id).lstrip('0')
    last_name = emp_name.split()[-1]
    output_filename = f"{clean_emp_id} - {last_name}.xlsx"

    def parse_dt(d_str):
        for fmt in ("%m/%d/%Y", "%m/%d/%y"):
            try: return datetime.datetime.strptime(d_str, fmt).date()
            except ValueError: pass
        return datetime.date.today()

    req_start = parse_dt(start_date_str)
    req_end = parse_dt(end_date_str)

    # 2. Extract SAP Table
    df = parse_sap_pipe_table(raw_text)
    if df is None or df.empty:
        st.error("Could not parse SAP table. Ensure table rows start with '|' pipe symbols.")
        return None, None, None

    df['Date_dt'] = pd.to_datetime(df['Date']).dt.date

    # 3. Missing Boundary Check
    missing_dates = []
    if req_start not in df['Date_dt'].values:
        missing_dates.append(req_start.strftime("%m/%d/%Y"))
    if req_end not in df['Date_dt'].values:
        missing_dates.append(req_end.strftime("%m/%d/%Y"))

    df_filtered = df[(df['Date_dt'] >= req_start) & (df['Date_dt'] <= req_end)].copy()

    # 4. Excel Generation
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "FMLA Hours Log"
    ws_out.views.sheetView[0].showGridLines = True

    NAVY_PRIMARY = "0F172A"
    ZEBRA_FILL = PatternFill(start_color="F8FAFC", fill_type="solid")
    WHITE_FILL = PatternFill(start_color="FFFFFF", fill_type="solid")
    GOLD_TOTAL_FILL = PatternFill(start_color="FEF08A", fill_type="solid")
    CARD_BG_FILL = PatternFill(start_color="F1F5F9", fill_type="solid")
    ALERT_BG_FILL = PatternFill(start_color="FEF2F2", fill_type="solid")

    FONT_BANNER = Font(name="Segoe UI", size=16, bold=True, color="FFFFFF")
    FONT_HEADER = Font(name="Segoe UI", size=12, bold=True, color="FFFFFF")
    FONT_BODY = Font(name="Segoe UI", size=11, color="1E293B")
    FONT_BOLD = Font(name="Segoe UI", size=11, bold=True, color="0F172A")
    FONT_TOTAL = Font(name="Segoe UI", size=12, bold=True, color="0F172A")
    FONT_KPI_VAL = Font(name="Segoe UI", size=26, bold=True, color="1E3A8A")
    FONT_KPI_LBL = Font(name="Segoe UI", size=9, bold=True, color="64748B")
    FONT_ALERT_TITLE = Font(name="Segoe UI", size=11, bold=True, color="991B1B")
    FONT_ALERT_BODY = Font(name="Segoe UI", size=11, bold=True, italic=True, color="B91C1C")

    BORDER_SUBTLE = Side(border_style="thin", color="E2E8F0")
    BORDER_GRID = Border(left=BORDER_SUBTLE, right=BORDER_SUBTLE, top=BORDER_SUBTLE, bottom=BORDER_SUBTLE)
    BORDER_TOTAL = Border(top=Side(border_style="medium", color="0F172A"), bottom=Side(border_style="double", color="0F172A"), left=BORDER_SUBTLE, right=BORDER_SUBTLE)
    BORDER_CARD = Border(left=Side(border_style="thick", color="2563EB"), right=BORDER_SUBTLE, top=BORDER_SUBTLE, bottom=BORDER_SUBTLE)

    ws_out.merge_cells("A1:K2")
    ws_out["A1"] = "  FMLA / STD HOURS WORKED VERIFICATION REPORT"
    ws_out["A1"].font = FONT_BANNER
    ws_out["A1"].fill = PatternFill(start_color=NAVY_PRIMARY, fill_type="solid")
    ws_out["A1"].alignment = Alignment(horizontal="left", vertical="center")

    start_table_row = 5
    table_headers = ['Pers.No.', 'Name', 'Period', 'Date', 'TmType', 'TimeTyText', 'Number', 'Cost Ctr', 'PSubarea', 'Subarea', 'Cost Ctr Ref']
    for col_i, h_txt in enumerate(table_headers, start=1):
        cell = ws_out.cell(row=start_table_row, column=col_i, value=h_txt)
        cell.fill = PatternFill(start_color=NAVY_PRIMARY, fill_type="solid")
        cell.font = FONT_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER_GRID

    ws_out.row_dimensions[start_table_row].height = 32

    filtered_vals = df_filtered.iloc[:, :11].values
    for idx, r_vals in enumerate(filtered_vals, start=start_table_row + 1):
        ws_out.row_dimensions[idx].height = 22
        row_fill = ZEBRA_FILL if idx % 2 == 0 else WHITE_FILL
        for col_i in range(1, 12):
            val = r_vals[col_i - 1]
            if col_i in [1, 3, 5, 10, 11]:
                try: val = int(float(val))
                except: pass
            elif col_i == 7:
                try: val = float(val)
                except: pass

            cell = ws_out.cell(row=idx, column=col_i, value=val)
            cell.font = FONT_BODY
            cell.fill = row_fill
            cell.border = BORDER_GRID

            if col_i in [1, 3, 5, 10, 11]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_i == 4:
                cell.number_format = "MM/DD/YYYY"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_i == 7:
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")

    total_row = start_table_row + len(filtered_vals) + 1
    ws_out.cell(row=total_row, column=6, value="Total Hours:").font = FONT_TOTAL
    ws_out.cell(row=total_row, column=6).border = BORDER_TOTAL
    sum_cell = ws_out.cell(row=total_row, column=7, value=f"=SUM(G6:G{total_row-1})")
    sum_cell.font = FONT_TOTAL
    sum_cell.number_format = "#,##0.00"
    sum_cell.fill = GOLD_TOTAL_FILL
    sum_cell.border = BORDER_TOTAL

    excel_buffer = io.BytesIO()
    wb_out.save(excel_buffer)
    excel_buffer.seek(0)

    # 5. Build Email
    total_hours_val = df_filtered['Number'].sum()
    missing_text = f" and did not work on {' and '.join(missing_dates)}" if missing_dates else ""

    email_response = f"""Per your request. Please see attached report to determine eligible hours. 

Employee logged {total_hours_val:,.2f} total hours worked{missing_text}. 

Direct any questions to your supervisor.

Thank you,

Anthony Cortez
Benefits/HRIS Analyst"""

    return email_response, excel_buffer, output_filename

# Process Action
if st.button("🚀 Process Claim & Generate Report"):
    if not text_input.strip():
        st.warning("Please paste text before clicking process.")
    else:
        email_txt, excel_data, filename = process_combined_text(text_input)
        if email_txt:
            st.success("Claim audit complete!")
            st.subheader("✉️ Ready-to-Send Email Reply")
            st.code(email_txt, language="text")
            
            st.subheader("📥 Executive Excel Report")
            st.download_button(
                label=f"Download {filename}",
                data=excel_data,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
