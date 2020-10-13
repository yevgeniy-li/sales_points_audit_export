import math
# from PIL import Image
# from PIL.ExifTags import TAGS, GPSTAGS

# def get_file_info(filename):
#     exif = Image.open(filename)._getexif()
#     if exif is not None:
#         file_info = {}
#         for key in exif.keys():
#             name = TAGS.get(key)
#             if not name is None:
#                 if name == "GPSInfo":
#                     file_info["GPSInfo"] = {}
#                     for gps_key in exif[key].keys():
#                         gps_name = GPSTAGS.get(gps_key)
#                         if not gps_name is None:
#                             file_info["GPSInfo"][gps_name] = exif[key][gps_key]
#                 else:
#                     file_info[name] = exif[key]
#         return file_info
#     return None

# def get_decimal_coordinates(gps_info):
#     def convert_gps(val, ref):
#         sign = 1
#         if ref in ["S","W"]:
#             sign = -1
#         return (val[0] + val[1] / 60.0 + val[2] / 3600.0) * sign
#     latitude = round(convert_gps(gps_info["GPSLatitude"], gps_info["GPSLatitudeRef"]), 7)
#     longitude = round(convert_gps(gps_info["GPSLongitude"], gps_info["GPSLongitudeRef"]), 7)
#     if math.isnan(latitude) or math.isnan(longitude):
#         raise Exception("Метка GPS пустая!")
#         # return {
#         #     "GPSLatitude": None,
#         #     "GPSLongitude": None,
#         # }
#     else:
#         return {
#             "GPSLatitude": latitude,
#             "GPSLongitude": longitude,
#         }

# def get_gps_from_photo(filename):
#     file_info = get_file_info(filename)
#     if file_info is None:
#         raise Exception("Метаданные фото не найдены!")
#     if file_info.get("GPSInfo") is None:
#         raise Exception("Метка GPS не найдена!")
#         # print("Метка GPS не обнаружена!")
#         # return {
#         #     "GPSLatitude": None,
#         #     "GPSLongitude": None,
#         # }
#     return get_decimal_coordinates(file_info["GPSInfo"])

def haversine_distance(point1, point2):
    lat1, lon1 = point1
    lat2, lon2 = point2
    radius = 6371  # km

    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2) * math.sin(dlat / 2) + math.cos(
        math.radians(lat1)
    ) * math.cos(math.radians(lat2)) * math.sin(dlon / 2) * math.sin(
        dlon / 2
    )
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    d = radius * c

    return d
    