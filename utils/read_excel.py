from openpyxl import load_workbook

def read_excel_range(file_path, cell_range):
    workbook = load_workbook(file_path, data_only=True)
    sheet = workbook.active

    name_data = []
    id_data = []

    rows = sheet[cell_range]
    col_offset = 1  # C 열을 읽어오기 위한 열의 offset

    for row in rows:
        name_data.append(row[0].value)
        student_id = row[1].value
        if student_id and len(student_id) >= 4:
            id_data.append(student_id[2:4] + "학번")
        else:
            id_data.append("")

    return name_data, id_data