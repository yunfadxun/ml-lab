import datetime
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st
from github import Github

# --- PAGE SETUP & CONSTANTS ---
st.set_page_config(
    page_title="SEAR Measles & Rubella Monthly Reporting", layout="wide"
)

LABS = [
    "Aceh",
    "Bandung",
    "Banjarbaru",
    "Banjarnegara",
    "BBLKM Jakarta",
    "BBLKM Yogyakarta",
    "Binomika Jakarta",
    "Jakarta",
    "Labkesda Jakarta",
    "Makassar",
    "Palembang",
    "Surabaya",
    "Yogyakarta",
]

MONTHS = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]

COLUMNS = [
    "Month",  # Col 1
    "Number of Serum Samples Received",  # Col 2
    "Number Tested for Measles",  # Col 3
    "Measles IgM Positive",  # Col 4
    "Measles IgM Negative",  # Col 5
    "Measles IgM Equivocal",  # Col 6
    "Number Tested for Rubella",  # Col 7
    "Rubella IgM Positive",  # Col 8
    "Rubella IgM Negative",  # Col 9
    "Rubella IgM Equivocal",  # Col 10
    "Number of Samples Negative for Measles and Rubella",  # Col 11
    "Number of samples with results within 7 days",  # Col 12
    "Number of samples with Measles results within 4 days",  # Col 13
    "Number of samples with Rubella results within 4 days",  # Col 14
    "Number of Urine Samples Received",  # Col 16
    "Urine Measles Isolation Positive",  # Col 17
    "Urine Measles Isolation Negative",  # Col 18
    "Number of Throat Swab Received",  # Col 19
    "Throat Swab Measles Isolation Positive",  # Col 20
    "Throat Swab Measles Isolation Negative",  # Col 21
]


def get_empty_df():
  """Creates a default template dataframe with 12 months initialized to 0."""
  df = pd.DataFrame(columns=COLUMNS)
  df["Month"] = MONTHS
  for col in COLUMNS[1:]:
    df[col] = 0
  return df


# --- GITHUB UTILITIES ---
def get_github_repo():
  try:
    token = st.secrets["GITHUB_TOKEN"]
    repo_name = st.secrets["GITHUB_REPO"]  # Format: "owner/repo"
    g = Github(token)
    return g.get_repo(repo_name)
  except Exception as e:
    st.error(f"GitHub Auth Error: {e}")
    return None


def load_master_excel(file_name, repo):
  """Loads master workbook data (rows starting from row 7/8)."""
  try:
    file_content = repo.get_contents(file_name)
    excel_file = pd.ExcelFile(io.BytesIO(file_content.decoded_content))

    data_dict = {}
    date_dict = {}

    for sheet in excel_file.sheet_names:
      # Read full raw sheet to extract date header and table rows
      raw_df = pd.read_excel(excel_file, sheet_name=sheet, header=None)

      # Extract Date of Report from Cell I3 (row index 2, col index 8) if present
      report_date = None
      try:
        val = raw_df.iloc[2, 8]
        report_date = (
            pd.to_datetime(val).date() if pd.notna(val) else datetime.date.today()
        )
      except Exception:
        report_date = datetime.date.today()
      date_dict[sheet] = report_date

      # Extract monthly data table (Rows starting from Row 7 index 6)
      df_data = pd.read_excel(
          excel_file, sheet_name=sheet, skiprows=6, nrows=12
      )
      # Drop formula total row if present
      df_data = df_data.iloc[:12, : len(COLUMNS)]
      df_data.columns = COLUMNS
      data_dict[sheet] = df_data

    return data_dict, date_dict, file_content.sha
  except Exception:
    # If file doesn't exist yet, return initial defaults
    default_data = {lab: get_empty_df() for lab in LABS}
    default_dates = {lab: datetime.date.today() for lab in LABS}
    return default_data, default_dates, None


def generate_exact_formatted_workbook(
    data_dict, date_dict, report_title_type="CBMS"
):
  """Generates an openpyxl Workbook matching the WHO SEARO template layout."""
  wb = openpyxl.Workbook()
  # Remove default sheet
  wb.remove(wb.active)

  thin = Side(border_style="thin", color="000000")
  thick = Side(border_style="medium", color="000000")
  border_box = Border(left=thin, right=thin, top=thin, bottom=thin)

  header_fill = PatternFill(
      start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
  )

  for lab_name in LABS:
    ws = wb.create_sheet(title=lab_name)
    ws.views.sheetView[0].showGridLines = True

    # 1. Header Title (Row 1)
    title_str = (
        f"SEAR - Measles Laboratory Network - Monthly Reporting"
        f" ({report_title_type})"
    )
    ws.merge_cells("A1:U1")
    ws["A1"] = title_str
    ws["A1"].font = Font(name="Calibri", size=11, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    # 2. Lab Name & Date of Report (Row 3)
    ws["A3"] = "Laboratory:"
    ws["A3"].font = Font(bold=True)

    ws.merge_cells("B3:G3")
    ws["B3"] = f"BB LABKESMAS {lab_name.upper()}"
    ws["B3"].font = Font(bold=True, italic=True)
    ws["B3"].border = Border(
        left=thick, right=thick, top=thick, bottom=thick
    )  # Bordered text box

    ws["I3"] = "Date of Report:"
    ws["I3"].font = Font(bold=True)

    ws.merge_cells("K3:M3")
    report_date = date_dict.get(lab_name, datetime.date.today())
    ws["K3"] = report_date.strftime("%d-%b-%Y")
    ws["K3"].alignment = Alignment(horizontal="center")
    ws["K3"].border = Border(
        left=thick, right=thick, top=thick, bottom=thick
    )  # Bordered date box

    # 3. Column Headers (Row 5 - 6)
    headers_map = [
        ("A5:A6", "Month"),
        ("B5:B6", "Number of Serum\nSamples Received"),
        ("C5:C6", "Number\nTested for\nMeasles"),
        ("D5:F5", "Measles IgM Results"),
        ("D6:D6", "Positive"),
        ("E6:E6", "Negative"),
        ("F6:F6", "Equivocal"),
        ("G5:G6", "Number\nTested for\nRubella"),
        ("H5:J5", "Rubella IgM Results"),
        ("H6:H6", "Positive"),
        ("I6:I6", "Negative"),
        ("J6:J6", "Equivocal"),
        (
            "K5:K6",
            "Number of\nSamples\nNegative for\nMeasles and\nRubella",
        ),
        ("L5:L6", "Number of\nsamples with\nresults within\n7 days"),
        ("M5:M6", "Number of\nsamples with\nMeasles results\nwithin 4 days"),
        ("N5:N6", "Number of\nsamples with\nRubella results\nwithin 4 days"),
        ("O5:O6", "Number of\nUrine Samples\nReceived"),
        ("P5:Q5", "Measles Isolation"),
        ("P6:P6", "Positive"),
        ("Q6:Q6", "Negative"),
        ("R5:R6", "Number of\nThroat Swab\nReceived"),
        ("S5:T5", "Measles Isolation"),
        ("S6:S6", "Positive"),
        ("T6:T6", "Negative"),
    ]

    for cell_range, text in headers_map:
      if ":" in cell_range:
        start, end = cell_range.split(":")
        if start != end:
          ws.merge_cells(cell_range)
        cell = ws[start]
      else:
        cell = ws[cell_range]

      cell.value = text
      cell.font = Font(name="Calibri", size=9, bold=True)
      cell.alignment = Alignment(
          horizontal="center", vertical="center", wrap_text=True
      )
      cell.fill = header_fill

    # Apply borders across headers
    for r in range(5, 7):
      for c in range(1, 21):
        ws.cell(row=r, column=c).border = border_box

    # 4. Column Numbers (Row 7)
    col_numbers = [
        1,
        2,
        3,
        4,
        5,
        6,
        7,
        8,
        9,
        10,
        11,
        12,
        13,
        14,
        16,
        17,
        18,
        19,
        20,
        21,
    ]
    for c_idx, num in enumerate(col_numbers, start=1):
      cell = ws.cell(row=7, column=c_idx, value=f"Column: {num}" if c_idx==1 else num)
      cell.font = Font(
          name="Calibri", size=9, italic=True, color="FF0000" if c_idx==1 else "000000"
      )
      cell.alignment = Alignment(horizontal="center", vertical="center")
      cell.border = border_box

    # 5. Populate Monthly Data Rows (Rows 8 to 19)
    df = data_dict.get(lab_name, get_empty_df())

    for i, row in df.iterrows():
      row_num = 8 + i
      ws.cell(row=row_num, column=1, value=row["Month"]).alignment = Alignment(
          horizontal="center"
      )

      for c_idx, col_name in enumerate(COLUMNS[1:], start=2):
        val = row[col_name]
        cell = ws.cell(
            row=row_num, column=c_idx, value=int(val) if pd.notna(val) else 0
        )
        cell.alignment = Alignment(horizontal="center")

      # Add borders
      for c_idx in range(1, 21):
        ws.cell(row=row_num, column=c_idx).border = border_box

    # 6. Total Row (Row 20) with Formulas
    ws.cell(row=20, column=1, value="Total").font = Font(bold=True)
    ws.cell(row=20, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=20, column=1).border = border_box

    for c_idx in range(2, 21):
      col_letter = get_column_letter(c_idx)
      cell = ws.cell(row=20, column=c_idx, value=f"=SUM({col_letter}8:{col_letter}19)")
      cell.font = Font(bold=True)
      cell.alignment = Alignment(horizontal="center")
      cell.border = border_box

  output = io.BytesIO()
  wb.save(output)
  return output.getvalue()


# --- APP INTERFACE ---
st.title("🔬 SEAR Measles & Rubella Monthly Report Manager")

repo = get_github_repo()

# Step 1: Report Type
report_type = st.radio(
    "1. Select Report Type:",
    options=["CBMS (Case-Based Measles Surveillance)", "OB (Outbreak)"],
    horizontal=True,
)

file_name = "master_cbms.xlsx" if "CBMS" in report_type else "master_ob.xlsx"
title_code = "CBMS" if "CBMS" in report_type else "Outbreak"

# Step 2: Choose Laboratory Site
selected_lab = st.selectbox("2. Select Laboratory / Site:", options=LABS)

if repo:
  data_dict, date_dict, sha = load_master_excel(file_name, repo)

  # Step 3: Date Picker for "Date of Report"
  current_date = date_dict.get(selected_lab, datetime.date.today())
  selected_date = st.date_input(
      f"3. Select Date of Report for {selected_lab}:", value=current_date
  )

  # Step 4: Data Entry Table
  st.subheader(
      f"4. Input Data starting at Row 8 (Jan–Dec) for **{selected_lab}**"
  )

  lab_df = data_dict.get(selected_lab, get_empty_df())

  # Interactive table editor
  edited_df = st.data_editor(
      lab_df, num_rows="fixed", use_container_width=True, hide_index=True
  )

  # Save Button
  if st.button("💾 Save & Commit to GitHub Master File", type="primary"):
    data_dict[selected_lab] = edited_df
    date_dict[selected_lab] = selected_date

    with st.spinner("Generating formatted master Excel file and saving to GitHub..."):
      excel_bytes = generate_exact_formatted_workbook(
          data_dict, date_dict, report_title_type=title_code
      )

      commit_msg = (
          f"Update {selected_lab} ({title_code}) report date {selected_date}"
      )
      if sha:
        repo.update_file(file_name, commit_msg, excel_bytes, sha)
      else:
        repo.create_file(file_name, commit_msg, excel_bytes)

    st.success(
        f"Successfully updated sheet **'{selected_lab}'** with header date"
        f" **{selected_date.strftime('%d-%b-%Y')}** in **{file_name}**!"
    )

  st.divider()

  # Download Current Master File Button
  st.subheader("📥 Export Master File")
  export_bytes = generate_exact_formatted_workbook(
      data_dict, date_dict, report_title_type=title_code
  )

  st.download_button(
      label=f"Download Master File ({file_name})",
      data=export_bytes,
      file_name=file_name,
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )
