import os
import re

import pandas
import openpyxl

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
    filter_pattern = r"[0-9a-zA-Zа-яА-Я \_\(\)\+\-\=\.\,]+"
    # filtred_string = unfiltred_string.replace("/", "_")
    filtred_string_parts = re.findall(filter_pattern, unfiltred_string)
    filtred_string = "".join(filtred_string_parts)
    return filtred_string.strip()

if __name__ == "__main__":
    token = os.path.join("token", "python-trade-points-audit.json")
    SPREADSHEET_ID = "1pQE5VS2Vf69oHbo6P-Xs2BVIHlBDCrT_jfjv7smg-f4"
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly"
    ]

    gservice = GService(SCOPES)
    gservice.auth(token)
    spreadsheets = gservice.get_spreadsheets()
    service_gdrive = gservice.get_service_drive()

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
        "{51}": "TaskVodka",
        "{52}": "TaskCognac",
        "{53}": "TaskWater",
        "{59}": "TaskWine",

        "{4}": "link_FacadePhoto",
        "{9}": "link_ShelfPhoto_general_vodka",
        "{10}": "link_ShelfPhoto_closeup_vodka",
        "{35}": "link_ShelfPhoto_general_cognac",
        "{36}": "link_ShelfPhoto_closeup_cognac",
        "{61}": "link_ShelfPhoto_general_wine",
        "{62}": "link_ShelfPhoto_closeup_wine",
        "{65}": "link_PhotoRefrigerator",
    }

    columns_new_names = {
        "Отметка времени": "Date"
    }

    AUDIT_CODE = "PLMNB10"
    AUDIT_NAME = "Алматы 2020-10-12 - 2020-10-31"

    # dir_root = "D:\\python\\sales_points_audit_export\\temp"
    dir_root = "\\\\aw.com\\cloud\\REPORTS\\АУДИТ\\Almaty"
    dir_audit = os.path.join(dir_root, AUDIT_NAME)
    if not os.path.exists(dir_audit):
        os.mkdir(dir_audit)
    dir_audit_photos = os.path.join(dir_audit, "Photo")
    if not os.path.exists(dir_audit_photos):
        os.mkdir(dir_audit_photos)

    print("Выгрузка аудита: %s" % AUDIT_NAME)

    for column_name in df_data.columns:
        for column_code in col_names_code:
            if str(column_code) in str(column_name):
                columns_new_names[column_name] = col_names_code[column_code]

    df_data = df_data.rename(columns=columns_new_names)

    df_data = df_data[df_data["AuditCode"] == AUDIT_CODE]
    df_data = df_data[df_data["Comment"].notna()]
    df_data["AuditName"] = AUDIT_NAME
    # df_data = df_data.head(5)

    # удаление дублей
    df_data['row_key'] = df_data['AuditCode'] + '-' + df_data['SalesPointTitle'] + '-' + df_data['Address']
    df_data_unique = []
    for row_key in df_data['row_key'].unique():
        df_data_unique.append(df_data[df_data['row_key'] == row_key].iloc[0])
    df_data = pandas.DataFrame(df_data_unique)

    df_data['Date'] = pandas.to_datetime(df_data['Date'], format='%d.%m.%Y %H:%M:%S')

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "sheet1"

    header_row = 1
    first_row = 2
    first_col = 1

    xl_columns = [
        {
            "header_text": "Аудит",
            "value_column": "AuditName",
            "width": 12,
            "is_hyperlink": False,
            "file_name": None
        },
        {
            "header_text": "Дата",
            "value_column": "Date",
            "width": 10,
            "is_hyperlink": False,
            "file_name": None
        },
        {
            "header_text": "Аудитор",
            "value_column": "AuditorName",
            "width": 10,
            "is_hyperlink": False,
            "file_name": None
        },
        {
            "header_text": "Категория точки",
            "value_column": "SalesChannel",
            "width": 7,
            "is_hyperlink": False,
            "file_name": None
        },
        {
            "header_text": "Точка",
            "value_column": "SalesPointTitle",
            "width": 16,
            "is_hyperlink": False,
            "file_name": None
        },
        {
            "header_text": "Адрес",
            "value_column": "Address",
            "width": 19,
            "is_hyperlink": False,
            "file_name": None
        },
        {
            "header_text": "Комментарий",
            "value_column": "Comment",
            "width": 21,
            "is_hyperlink": False,
            "file_name": None
        },
        {
            "header_text": "Задача водка",
            "value_column": "TaskVodka",
            "width": 21,
            "is_hyperlink": False,
            "file_name": None
        },
        {
            "header_text": "Задача коньяк",
            "value_column": "TaskCognac",
            "width": 21,
            "is_hyperlink": False,
            "file_name": None
        },
        {
            "header_text": "Задача вода и напитки",
            "value_column": "TaskWater",
            "width": 21,
            "is_hyperlink": False,
            "file_name": None
        },
        {
            "header_text": "Задача вино",
            "value_column": "TaskWine",
            "width": 21,
            "is_hyperlink": False,
            "file_name": None
        },
        {
            "header_text": "Папка",
            "value_column": None,
            "width": 6,
            "is_hyperlink": True,
            "file_name": None
        },
        {
            "header_text": "Фото фасада",
            "value_column": "link_FacadePhoto",
            "width": 6,
            "is_hyperlink": True,
            "file_name": "facade_photo"
        },
        {
            "header_text": "Фото водочной полки",
            "value_column": "link_ShelfPhoto_general_vodka",
            "width": 6,
            "is_hyperlink": True,
            "file_name": "vodka_general"
        },
        {
            "header_text": "Фото водочной полки крупным планом",
            "value_column": "link_ShelfPhoto_closeup_vodka",
            "width": 6,
            "is_hyperlink": True,
            "file_name": "vodka_closeup"
        },
        {
            "header_text": "Фото коньячной полки",
            "value_column": "link_ShelfPhoto_general_cognac",
            "width": 6,
            "is_hyperlink": True,
            "file_name": "cognac_general"
        },
        {
            "header_text": "Фото коньячной полки крупным планом",
            "value_column": "link_ShelfPhoto_closeup_cognac",
            "width": 6,
            "is_hyperlink": True,
            "file_name": "cognac_closeup"
        },
        {
            "header_text": "Фото винной полки",
            "value_column": "link_ShelfPhoto_general_wine",
            "width": 6,
            "is_hyperlink": True,
            "file_name": "wine_general"
        },
        {
            "header_text": "Фото винной полки крупным планом",
            "value_column": "link_ShelfPhoto_closeup_wine",
            "width": 6,
            "is_hyperlink": True,
            "file_name": "wine_closeup"
        },
        {
            "header_text": "Фото ХО Азия Су",
            "value_column": "link_PhotoRefrigerator",
            "width": 6,
            "is_hyperlink": True,
            "file_name": "refrigerator_awc"
        }
    ]

    # cell styles
    alignment_header = Alignment(horizontal='center', vertical='center', wrap_text=True)
    alignment_cell = Alignment(horizontal='left', vertical='top', wrap_text=True)
    # cell border
    side_thin = Side(style="thin", color="000000")
    border_all = Border(left=side_thin, right=side_thin, top=side_thin, bottom=side_thin)
    # cell font
    font_bold = Font(bold=True)

    for col_ind, col_inf in enumerate(xl_columns):
        # columns width
        ws.column_dimensions[get_column_letter(first_col + col_ind)].width = col_inf["width"]
        cell = ws.cell(header_row, first_col + col_ind)
        cell.value = col_inf["header_text"]
        cell.alignment = alignment_header
        cell.border = border_all
        cell.font = font_bold

    ws.freeze_panes = "A2"

    i = first_row
    len_df_data = len(df_data.index)
    for row in df_data.iterrows():
        row_data = row[1]
        # indicator
        print("({0} / {1}) {2}".format(i - 1, len_df_data, row_data["SalesPointTitle"]))
        # photo name parts
        filtred_sales_point = remove_illegal_symbol(row_data["SalesPointTitle"])
        filtred_address = remove_illegal_symbol(row_data["Address"])
        dir_sales_point_name = filtred_sales_point + "_" + filtred_address
        sales_point_dir = os.path.join(dir_audit_photos, dir_sales_point_name)
        if not os.path.exists(sales_point_dir):
            os.mkdir(sales_point_dir)
        for col_ind, col_inf in enumerate(xl_columns):
            cell = ws.cell(i, first_col + col_ind)
            if col_inf["is_hyperlink"]:
                if col_inf["value_column"] is not None:
                    photo_file_name = download_photo(row_data[col_inf["value_column"]], service_gdrive, col_inf["file_name"], sales_point_dir)
                    if photo_file_name is not None:
                        photo_link = os.path.join("Photo", dir_sales_point_name, photo_file_name)
                        cell.value = "=HYPERLINK(\"{0}\", \"X\")".format(photo_link)
                else:
                    cell.value = "=HYPERLINK(\"{0}\", \"X\")".format(os.path.join("Photo", dir_sales_point_name))
                cell.style = 'Hyperlink'
                cell.alignment = alignment_cell
                cell.border = border_all
            else:
                cell.value = row_data[col_inf["value_column"]]
                cell.alignment = alignment_cell
                cell.border = border_all
        i += 1

    excel_file_name = "Задачи.xlsx"
    wb.save(os.path.join(dir_audit, excel_file_name))
