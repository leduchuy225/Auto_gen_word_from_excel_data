import pandas as pd
from docx import Document
import os
from datetime import datetime
import re
from num2words import num2words

excel_sheet_name = "M07B"
excel_file_name = "GQVL.xlsx"
word_file_name = "Hop_dong_TD.docx"

column_mappings = {
    'stt': 'STT',
    'hop_dong_tin_dung_so': 'Hợp đồng tín dụng số',
    'ten_ben_cho_vay': 'Tên bên cho vay',
    'dia_chi_ben_cho_vay': 'Địa chỉ bên cho vay',
    'dien_thoai_ben_cho_vay': 'Điện thoại bên cho vay',
    'ho_va_ten_nguoi_dai_dien': 'Họ và tên người đại diện',
    'chuc_vu': 'Chức vụ',
    'ho_ten_nguoi_vay': 'Họ tên người vay',
    'nam_sinh': 'Năm sinh',
    'tuoi': 'Tuổi',
    'so_can_cuoc': 'Số Căn cước',
    'ngay_cap': 'Ngày cấp',
    'noi_cap': 'Nơi cấp',
    'noi_cu_tru': 'Nơi cư trú',
    'dien_thoai': 'Điện thoại',
    'so_tien_cho_vay_dong': 'Số tiền cho vay (đồng)',
    'thoi_gian_cho_vay_thang': 'Thời gian cho vay (tháng)',
    'han_tra_no_cuoi_cung': 'Hạn trả nợ cuối cùng',
    'muc_dich_su_dung_tien_vay': 'Mục đích sử dụng tiền vay',
    'ky_han_tra_no_goc_so_thang_ky': 'Kỳ hạn trả nợ gốc (số tháng/kỳ)',
    'so_tien_tra_no_goc_cho_moi_ky_han_dong': 'Số tiền trả nợ gốc cho mỗi kỳ hạn (đồng)',
    'ngay_bat_dau_tra_goc': 'Ngày bắt đầu trả gốc'
}

convert_str_field = ['so_can_cuoc', 'dien_thoai', 'dien_thoai_ben_cho_vay']

# Output folder
output_dir = "generated_docs"
os.makedirs(output_dir, exist_ok=True)

today = datetime.today()
current_date = f"ngày {today.day:02d} tháng {today.month:02d} năm {today.year}"

# Read Excel file (first sheet by default)
df = pd.read_excel(excel_file_name, sheet_name=excel_sheet_name,
                   engine='openpyxl', converters={column_mappings[field]: str for field in convert_str_field})

print(df.dtypes)


def evaluate_placeholder(expr: str):
  match = re.match(r'{{\s*(\w+)\((\w+)\)\s*}}', expr)
  if not match:
    return None

  func_name, arg_name = match.groups()

  return func_name, arg_name


def generate_ky_han_tra_no_goc(start_date, end_date, amount):
  start_date = datetime.strptime(start_date, "%d/%m/%Y")
  end_date = datetime.strptime(end_date, "%d/%m/%Y")

  # Create the list of yearly dates
  dates = []
  current = start_date
  while current <= end_date:
    dates.append(current)
    current = current.replace(year=current.year + 1)

  result_lines = []

  for date in dates:
    amount_str = f"{amount:,.0f}"  # Format number only
    result_lines.append(
        f"\t- Ngày {date.strftime('%d/%m/%Y')}, số tiền {amount_str} đồng.")

  # Combine all into one string
  final_output = "\n".join(result_lines)

  return final_output


# For each row in the Excel file
for index, row in df.iterrows():
  # Load template
  doc = Document(word_file_name)

  replacements = {}

  for var, col in column_mappings.items():
    if col in row and pd.notna(row[col]):
      value = row[col]

      # Check if it's a number (int or float), format with digit grouping
      if isinstance(value, (int, float)) and "tiền" in col:
        formatted_value = f"{value:,.0f}"
      else:
        formatted_value = str(value)

      replacements[f"{{{{{var}}}}}"] = formatted_value

  # Replace placeholders in paragraphs
  for paragraph in doc.paragraphs:
    for run in paragraph.runs:
      if f"{{current_date}}" in run.text:
        run.text = run.text.replace("{{current_date}}", str(current_date))

      if f"{{ky_han_tra_no_goc}}" in run.text:
        run.text = run.text.replace("{{ky_han_tra_no_goc}}",
                                    generate_ky_han_tra_no_goc(row[column_mappings['ngay_bat_dau_tra_goc']], row[column_mappings['han_tra_no_cuoi_cung']], row[column_mappings['so_tien_tra_no_goc_cho_moi_ky_han_dong']]))

      # Replace function placeholders
      func_placeholders = re.findall(r'{{\s*\w+\(\w+\)\s*}}', run.text)
      for ph in func_placeholders:
        func_name, arg_name = evaluate_placeholder(ph)
        if func_name == 'num2words':
          run.text = run.text.replace(
              ph, num2words(row[column_mappings[arg_name]], lang='vi').capitalize() + " đồng")

      for key, value in replacements.items():
        if key in run.text:
          run.text = run.text.replace(key, value)

  for table in doc.tables:
    for row_obj in table.rows:
      for cell in row_obj.cells:
        for p in cell.paragraphs:
          for key, value in replacements.items():
            if key in p.text:
              p.text = p.text.replace(key, value)

  # Save with a unique name (e.g., by borrower name or row number)
  borrower = str(row[column_mappings['ho_ten_nguoi_vay']]
                 ).replace(" ", "_").replace("/", "_")

  filename = f"{borrower}.docx"
  doc.save(os.path.join(output_dir, filename))

print("All documents generated successfully.")
