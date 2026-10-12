import datetime
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Measles & Rubella Multi-Lab File Merger", layout="wide"
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
  """Generates a default empty template dataframe."""
  df = pd.DataFrame(columns=COLUMNS)
  df["Month"] = MONTHS
  for col in COLUMNS[1:]:
    df[col] = 0
  return df


def detect_report_type_from_sheet(sheet_name, row0_val):
  """Detects whether a sheet contains CBMS or OB (Outbreak) data based on sheet name or row 0 header string."""
  str_check = f"{str(sheet_name)} {str(row0_val)}".lower()
  if "outbreak" in str_check or " ob" in str_check or "_ob" in str_check:
    return "OB"
  return "CBMS"


def parse_uploaded_file_sheet(excel_file, sheet_name):
  """Parses a specific sheet from an Excel File into a 20-column DataFrame and extracts its date."""
  try:
    # Read raw content for header inspection
    raw_df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

    # Extract report date from cell I3 (row index 2, col index 8)
    report_date = datetime.date.today()
    try:
      val = raw_df.iloc[2, 8]
      if pd.notna(val):
        report_date = pd.to_datetime(val).date()
    except Exception:
      pass

    # Extract monthly table rows (skip top headers)
    df_data = pd.read_excel(
        excel_file, sheet_name=sheet_name, skiprows=6, nrows=12
    )

    # Align columns dynamically (handles source files with 14 vs 20 columns)
    num_cols = min(df_data.shape[1], len(COLUMNS))
    df_data = df_data.iloc[:12, :num_cols].copy()
    df_data.columns = COLUMNS[:num_cols]

    # Fill any missing columns with 0
    for col in COLUMNS:
      if col not in df_data.columns:
        df_data[col] = 0

    df_data = df_data[COLUMNS]
    return df_data, report_date
  except Exception as e:
    st.error(f"Error reading sheet '{sheet_name}': {e}")
    return get_empty_df(), datetime.date.today()


def parse_entire_uploaded_file(file_obj, fallback_type="CBMS"):
  """Reads an uploaded file and returns a list of tuples: [(detected_type, df_data, report_date), ...]."""
  results = []
  try:
    excel_file = pd.ExcelFile(file_obj)
    sheet_names = excel_file.sheet_names

    for sheet in sheet_names:
      # Read row 0 to check header text
      raw_df = pd.read_excel(excel_file, sheet_name=sheet, header=None, nrows=1)
      row0_val = raw_df.iloc[0, 0] if not raw_df.empty else ""

      detected_type = detect_report_type_from_sheet(sheet, row0_val)

      # If there is only 1 sheet in the workbook and auto-detection is neutral, use UI fallback choice
      if len(sheet_names) == 1 and "cbms" not in str(sheet).lower():
        detected_type = fallback_type

      df_data, report_date = parse_uploaded_file_sheet(excel_file, sheet)
      results.append((detected_type, df_data, report_date, sheet))

  except Exception as e:
    st.error(f"Error reading file {getattr(file_obj, 'name', '')}: {e}")

  return results


def build_master_workbook(data_dict, date_dict, report_title_type="CBMS"):
  """Creates formatted openpyxl workbook containing separate sheets for each lab."""
  wb = openpyxl.Workbook()
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
    ws.merge_cells("A1:U1")
    ws["A1"] = (
        f"SEAR - Measles Laboratory Network - Monthly Reporting"
        f" ({report_title_type})"
    )
    ws["A1"].font = Font(name="Calibri", size=11, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    # 2. Lab Name & Date of Report (Row 3)
    ws["A3"] = "Laboratory:"
    ws["A3"].font = Font(bold=True)

    ws.merge_cells("B3:G3")
    ws["B3"] = f"BB LABKESMAS {lab_name.upper()}"
    ws["B3"].font = Font(bold=True, italic=True)
    ws["B3"].border = Border(left=thick, right=thick, top=thick, bottom=thick)

    ws["I3"] = "Date of Report:"
    ws["I3"].font = Font(bold=True)

    ws.merge_cells("K3:M3")
    report_date = date_dict.get(lab_name, datetime.date.today())
    ws["K3"] = report_date.strftime("%d-%b-%Y")
    ws["K3"].alignment = Alignment(horizontal="center")
    ws["K3"].border = Border(left=thick, right=thick, top=thick, bottom=thick)

    # 3. Merged Column Headers (Rows 5–6)
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

    for r in range(5, 7):
      for c in range(1, 21):
        ws.cell(row=r, column=c).border = border_box

    # 4. Column Numbering (Row 7)
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
      cell = ws.cell(
          row=7, column=c_idx, value=f"Column: {num}" if c_idx == 1 else num
      )
      cell.font = Font(
          name="Calibri",
          size=9,
          italic=True,
          color="FF0000" if c_idx == 1 else "000000",
      )
      cell.alignment = Alignment(horizontal="center", vertical="center")
      cell.border = border_box

    # 5. Data Rows (Rows 8 to 19)
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

      for c_idx in range(1, 21):
        ws.cell(row=row_num, column=c_idx).border = border_box

    # 6. Total Formula Row (Row 20)
    ws.cell(row=20, column=1, value="Total").font = Font(bold=True)
    ws.cell(row=20, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=20, column=1).border = border_box

    for c_idx in range(2, 21):
      col_letter = get_column_letter(c_idx)
      cell = ws.cell(
          row=20, column=c_idx, value=f"=SUM({col_letter}8:{col_letter}19)"
      )
      cell.font = Font(bold=True)
      cell.alignment = Alignment(horizontal="center")
      cell.border = border_box

  output = io.BytesIO()
  wb.save(output)
  return output.getvalue()


# --- STREAMLIT APP INTERFACE ---
st.markdown(
    "<h2 style='font-weight: normal; font-size: 22px; color: #2C3E50;'>Measles &"
    " Rubella Multi-Lab Report Merger</h2>",
    unsafe_allow_html=True,
)
st.caption(
    "Upload files below to combine laboratory reports into formatted master"
    " sheets."
)

cbms_data_dict = {lab: get_empty_df() for lab in LABS}
cbms_date_dict = {lab: datetime.date.today() for lab in LABS}

ob_data_dict = {lab: get_empty_df() for lab in LABS}
ob_date_dict = {lab: datetime.date.today() for lab in LABS}

# --- 15 UPLOAD SLOTS INTERFACE ---
st.markdown("### Upload Files & Configure Metadata")

for slot in range(1, 16):
  with st.expander(f"Slot {slot}: Upload & Configure File", expanded=(slot == 1)):
    col1, col2, col3, col4 = st.columns([2, 1.5, 2, 2])

    with col1:
      file_obj = st.file_uploader(
          f"Upload File #{slot}", type=["xls", "xlsx"], key=f"file_{slot}"
      )

    with col2:
      report_type = st.radio(
          f"Type #{slot}",
          options=["Auto Detect / CBMS", "OB (Outbreak)"],
          key=f"type_{slot}",
      )

    with col3:
      default_lab_idx = (slot - 1) % len(LABS)
      lab_site = st.selectbox(
          f"Lab Site #{slot}",
          options=LABS,
          index=default_lab_idx,
          key=f"lab_{slot}",
      )

    with col4:
      report_date = st.date_input(
          f"Report Date #{slot}",
          value=datetime.date.today(),
          key=f"date_{slot}",
      )

    # Process file if uploaded
    if file_obj is not None:
      fallback = "OB" if "OB" in report_type else "CBMS"
      parsed_results = parse_entire_uploaded_file(
          file_obj, fallback_type=fallback
      )

      for dtype, df_data, p_date, sheet_name in parsed_results:
        # Use explicit date picked in UI if available, else parsed date
        final_date = report_date if report_date else p_date

        if dtype == "CBMS":
          cbms_data_dict[lab_site] = df_data
          cbms_date_dict[lab_site] = final_date
          st.info(
              f"Slot {slot} (Sheet '{sheet_name}'): Parsed as CBMS for {lab_site}"
          )
        else:
          ob_data_dict[lab_site] = df_data
          ob_date_dict[lab_site] = final_date
          st.info(
              f"Slot {slot} (Sheet '{sheet_name}'): Parsed as OB (Outbreak)"
              f" for {lab_site}"
          )

st.divider()

# --- MERGE & DOWNLOAD SECTION ---
st.markdown("### Generate & Download Master Excel Files")

col_left, col_right = st.columns(2)

with col_left:
  st.markdown("#### Master CBMS File")
  cbms_bytes = build_master_workbook(
      cbms_data_dict, cbms_date_dict, report_title_type="CBMS"
  )
  st.download_button(
      label="Download Master CBMS File",
      data=cbms_bytes,
      file_name="master_cbms.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )

with col_right:
  st.markdown("#### Master OB (Outbreak) File")
  ob_bytes = build_master_workbook(
      ob_data_dict, ob_date_dict, report_title_type="Outbreak"
  )
  st.download_button(
      label="Download Master OB File",
      data=ob_bytes,
      file_name="master_ob.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )
