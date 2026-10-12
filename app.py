import datetime
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="SEAR Measles & Rubella Multi-Lab File Merger", layout="wide"
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


def parse_uploaded_file(file_obj):
  """Extracts report date and 12-month data table from an uploaded Excel file."""
  try:
    raw_df = pd.read_excel(file_obj, header=None)

    # Extract report date from cell I3 (row index 2, col index 8)
    report_date = datetime.date.today()
    try:
      val = raw_df.iloc[2, 8]
      if pd.notna(val):
        report_date = pd.to_datetime(val).date()
    except Exception:
      pass

    # Extract monthly data (Rows starting from Row 8 / index 7 or Row 7 / index 6)
    # Check if header offset matches template
    df_data = pd.read_excel(file_obj, skiprows=6, nrows=12)
    df_data = df_data.iloc[:12, : len(COLUMNS)]
    df_data.columns = COLUMNS
    return df_data, report_date
  except Exception as e:
    st.error(f"Error reading file {file_obj.name}: {e}")
    return get_empty_df(), datetime.date.today()


def build_master_workbook(data_dict, date_dict, report_title_type="CBMS"):
  """Creates formatted workbook containing separate sheets for each lab."""
  wb = openpyxl.Workbook()
  wb.remove(wb.active)  # Remove default active sheet

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

    # Apply header borders
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

      for c
