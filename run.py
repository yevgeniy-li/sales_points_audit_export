import os
import pandas
import openpyxl
import re

from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Side, Font

from mylib.gservice import GService

def download_photo(link_photo, service_drive, target_file_name, dir_path):
    if not (pandas.isna(link_photo) or link_photo == ""):
        file_id = link_photo.split("id=")[1]
        file_info = service_drive.files().get(fileId=file_id, fields="name, fileExtension").execute()
        file_name = "{0}.{1}".format(target_file_name, file_info["fileExtension"])
        file_path = os.path.join(dir_path, file_name)
        if not os.path.exists(file_path):
            gservice.download_media(service_drive, file_id, file_path)
        return file_name
    return None

def remove_illegal_symbol(unfiltred_string):
    filter_pattern = r"[0-9a-zA-Zа-яА-Я_\+-=, \.\(\)]+"
    filtred_string = unfiltred_string.replace("/", "_")
    filtred_string_parts = re.findall(filter_pattern, filtred_string)
    filtred_string = "".join(filtred_string_parts)
    return filtred_string

if __name__ == "__main__":
    token = os.path.join("token", "python-trade-points-audit.json")
    SPREADSHEET_ID = "1pQE5VS2Vf69oHbo6P-Xs2BVIHlBDCrT_jfjv7smg-f4"
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly"
    ]

    AUDIT_CODE = "YTHNI09"
    AUDIT_NAME = "Алматы 2020-09-02 - 2020-09-04"

    temp_dir = ".\\temp"
    if not os.path.exists(temp_dir):
        os.mkdir(temp_dir)

    gservice = GService(SCOPES)
    gservice.auth(token)
    spreadsheets = gservice.get_spreadsheets()
    service_drive = gservice.get_service_drive()

    sheets = spreadsheets.get(spreadsheetId=SPREADSHEET_ID).execute()
    range_name = sheets["sheets"][0]["properties"]["title"]
    sheet_rows = spreadsheets.values().get(spreadsheetId=SPREADSHEET_ID, range=range_name).execute()
    values = sheet_rows.get("values", [])
    if len(values) <= 1:  # первая строка - заголовок
        raise Exception("No data found!")

    header = values[0]
    df_data = pandas.DataFrame(values)
    df_data.columns = header
    df_data = df_data.drop(0, axis="index")

    col_names_code ={
        "{1}": "AuditorName",
        "{3}": "AuditCode",
        "{5}": "SalesPointTitle",
        "{6}": "Address",
        "{7}": "SalesChannel",
        "{42}": "Comment",

        "{4}": "link_FacadePhoto",
        "{9}": "link_ShelfPhoto_general_vodka",
        "{10}": "link_ShelfPhoto_closeup_vodka",
        "{35}": "link_ShelfPhoto_general_cognac",
        "{36}": "link_ShelfPhoto_closeup_cognac"
    }

    columns_new_names = {
        "Отметка времени": "Date"
    }

    for column_name in df_data.columns:
        for column_code in col_names_code:
            if str(column_code) in str(column_name):
                columns_new_names[column_name] = col_names_code[column_code]

    df_data = df_data.rename(columns=columns_new_names)

    df_data = df_data[df_data["AuditCode"] == AUDIT_CODE]
    df_data = df_data[df_data["Comment"].notna()]
    df_data["AuditName"] = AUDIT_NAME
    df_data = df_data.head(10)

    wb = openpyxl.Workbook()
    ws = wb.active

    header_row = 1
    first_row = 2
    first_col = 1

    # columns width
    ws.column_dimensions[get_column_letter(first_col + 0)].width = 35
    ws.column_dimensions[get_column_letter(first_col + 1)].width = 20
    ws.column_dimensions[get_column_letter(first_col + 2)].width = 12
    ws.column_dimensions[get_column_letter(first_col + 3)].width = 20
    ws.column_dimensions[get_column_letter(first_col + 4)].width = 40
    ws.column_dimensions[get_column_letter(first_col + 5)].width = 50
    # cell styles
    alignment_left_top_wrap = Alignment(horizontal='left', vertical='top', wrap_text=True)
    alignment_left_top = Alignment(horizontal='left', vertical='top')
    # cell border
    side_thin = Side(style="thin", color="000000")
    border_all = Border(left=side_thin, right=side_thin, top=side_thin, bottom=side_thin)
    # cell font
    font_bold = Font(bold=True)
    # header
    ws.cell(header_row, first_col + 0).value = "Аудит"
    ws.cell(header_row, first_col + 1).value = "Аудитор"
    ws.cell(header_row, first_col + 2).value = "Категория точки"
    ws.cell(header_row, first_col + 3).value = "Точка"
    ws.cell(header_row, first_col + 4).value = "Адрес"
    ws.cell(header_row, first_col + 5).value = "Комментарий"
    ws.cell(header_row, first_col + 6).value = "Фото фасада"
    # =ГИПЕРССЫЛКА("Фото\Фото.jpg";"Ссылка на фото")

    for row in ws["A1:F1"]:
        for cell in row:
            cell.alignment = alignment_left_top_wrap
            cell.border = border_all
            cell.font = font_bold

    i = first_row
    for row in df_data.iterrows():
        row_data = row[1]
        ws.cell(i, first_col + 0).value = row_data["AuditName"]
        ws.cell(i, first_col + 1).value = row_data["AuditorName"]
        ws.cell(i, first_col + 2).value = row_data["SalesChannel"]
        ws.cell(i, first_col + 3).value = row_data["SalesPointTitle"]
        ws.cell(i, first_col + 4).value = row_data["Address"]
        ws.cell(i, first_col + 5).value = row_data["Comment"]
        # photo name parts
        filtred_sales_point = remove_illegal_symbol(row_data["SalesPointTitle"])
        filtred_address = remove_illegal_symbol(row_data["Address"])
        # download photo
        dir_name = filtred_sales_point.strip() + "_" + filtred_address.strip()
        photo_name = "FacadePhoto_" + dir_name
        sales_point_dir = os.path.join(temp_dir, dir_name)
        if not os.path.exists(sales_point_dir):
            os.mkdir(sales_point_dir)
        file_name = download_photo(row_data["link_FacadePhoto"], service_drive, photo_name, sales_point_dir)
        # cell formula
        ws.cell(i, first_col + 6).value = "=ГИПЕРССЫЛКА(\"{0}\\{1}\"; \"Фото фасада\")".format(dir_name, file_name)
        print(ws.cell(i, first_col + 6).value)

        i += 1

    cells_range = "A{0}:{1}{2}".format(first_row, get_column_letter(6), i-1)
    for row in ws[cells_range]:
        for cell in row:
            cell.alignment = alignment_left_top_wrap
            cell.border = border_all
    wb.save(".\\temp\\test.xlsx")
