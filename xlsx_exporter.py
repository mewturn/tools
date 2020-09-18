def xlsx_export(source_file, active_sheet_name):
    print(f"Processing {source_file}")
    data = get_data_dict_from_xlsx(source_file, active_sheet_name)
    write_data_to_xlsx(data)

def get_data_dict_from_xlsx(source_file, active_sheet_name, start_row=2):
    wb = load_workbook(filename=source_file)
    ws = wb[active_sheet_name]
    
    data_dict = dict()
    
    # start_row = 2 for cases with a title row
    for i in range(start_row, ws.max_row+1):
        glossary_id = ws[f"B{i}"].value
        source_text = ws[f"D{i}"].value
        target_text = ws[f"F{i}"].value
        
        if glossary_id in data_dict:
            data_dict[glossary_id]['terms'].append(tuple((source_text, target_text))) 
        else:
            data_dict[glossary_id] = {
                "lang": ws[f"M{i}"].value,
                "client_id": ws[f"G{i}"].value,
                "master_id": ws[f"H{i}"].value,
                "glossary_name": ws[f"J{i}"].value,
                "am_id": ws[f"K{i}"].value,
                "am_name": ws[f"L{i}"].value,
                "terms": [tuple((source_text, target_text))]
            }

    return data_dict

def write_data_to_xlsx(data):
    for glossary_id, details in data.items():
        file_dir = "glossary_export"
        filename = f"glossary{glossary_id}_client{details['client_id']}_{details['glossary_name']}_{details['lang']}_{details['am_name']}.xlsx"
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Glossary"
        ws['A1'].value = "id"
        ws['B1'].value = "source_text"
        ws['C1'].value = "target_text"
        ws['D1'].value = "comments"
        ws['E1'].value = "last_updated"
        
        row_num = 2
        for source_text, target_text in details['terms']:
            ws[f"B{row_num}"].value = source_text
            ws[f"C{row_num}"].value = target_text
            row_num += 1
        try:
            wb.save(f"{file_dir}/{filename}")
        except Exception as e:
            print(f"Issue with filename (likely invalid filename). Filename: {filename}")
