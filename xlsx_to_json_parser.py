import json
import os
import xlsxwriter
import pandas as pd

# Parse a singular JSON file from file path into a XLSX format
def parse_json_file(file_path):
    print(f"Processing {file_path}")
    workbook_name = file_path.split("/")[-1].replace(".json", ".xlsx")
    print(workbook_name)
    workbook = xlsxwriter.Workbook(f"json_out/{workbook_name}")
    worksheet = workbook.add_worksheet()
    row = 0
    column = 0
    
    with open(file_path, "r", encoding="utf-8") as f:
        x = json.loads(f.read())

    for obj, d in x.items():
        column = 0
        for k, v in d['_fields'].items():
            worksheet.write(row, 0, d['_id'])
            worksheet.write(row, 1, k)
            worksheet.write(row, 2, v['value'])
            row += 1

    workbook.close()
    print("Done!")

# Parse a JSON file by batch (folder)
def parse_json_by_folder(folder_path):
    for r, d, f in os.walk(folder_path):
        for name in f:
            if ".json" in name:
                file_path = os.path.join(r, name)
                parse_json_file(file_path)

# Reconstruct a JSON file from XLSX file of the same name
def reconstruct_json(file_path):
    print(f"Processing {file_path}")
    workbook_name = file_path.replace(".json", ".xlsx")
    output_name = file_path.replace(".json", "_output.json")
    
    try:
        r = pd.read_excel(f"{workbook_name}", header=None)
        # r = [id, key, source, target]
        x = zip(r[0], r[1], r[2], r[3])
    
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.loads(f.read())

            for id, key, source, target in x:
                data[id]['_fields'][key]['value'] = target

            with open(output_name, "w", encoding="utf-8") as g:
                g.write(json.dumps(data, ensure_ascii=False))

    except FileNotFoundError:
        print(f"{workbook_name} not found.")

# Reconstruct a JSON file from XLSX file of the same name (folder, batch process)
def reconstruct_json_by_folder(folder_path):
    for r, d, f in os.walk(folder_path):
        for name in f:
            if ".json" in name and not "_output" in name:
                file_path = os.path.join(r, name)
                reconstruct_json(file_path)
