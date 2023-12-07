import openpyxl

def fill_excel_with_data(data_arrays, file_name):
    try:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["counter", "author_id", "author_surname", "author_initials"])

        row = 2
        for data_array in data_arrays:
            col = 1
            for element in data_array:
                sheet.cell(row=row, column=col).value = element
                col += 1
                if col > 4:
                    col = 1
                    row += 1
            workbook.save(file_name)
    except Exception as e:
        print(f"An error occurred: {e}")

