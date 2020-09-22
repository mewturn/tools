import json
import os
import xlsxwriter

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
    
def parse_json_by_folder(folder_path):
    for r, d, f in os.walk(folder_path):
        for name in f:
            if ".json" in name:
                file_path = os.path.join(r, name)
                parse_json_file(file_path)
                
def reconstruct_json(file_path):
    print(f"Processing {file_path}")
    xlsx_folder = "insightjsonmilton/"
    workbook_name = file_path.replace(".json", ".xlsx")
    workbook_name = workbook_name.split("/")[-1]
    workbook_name = f"{xlsx_folder}/{workbook_name}"
    output_name = file_path.replace(".json", "_output.json")
    output_name = output_name.split("/")[-1]
    output_name = f"{xlsx_folder}/{output_name}"
    
    try:
        r = pd.read_excel(f"{workbook_name}", header=None)
        x = zip(r[2], r[3])
    
        with open(file_path, "r", encoding="utf-8") as f:
            data = f.read()

            for i, j in x:
                source = '"' + str(i) + '"'
                target = '"' + str(j) + '"'
                data = data.replace(source, target)

            with open(output_name, "w", encoding="utf-8") as g:
                g.write(data)

    except FileNotFoundError:
        print("XLSX file not found")

def reconstruct_json_by_folder(folder_path):
    for r, d, f in os.walk(folder_path):
        for name in f:
            if ".json" in name and not "_output" in name:
                file_path = os.path.join(r, name)
                reconstruct_json(file_path)
