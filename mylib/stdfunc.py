
import json

def read_json(file_path):
    with open(file_path, "r", encoding="utf-8") as read_file:
        data = json.load(read_file)
    return data
