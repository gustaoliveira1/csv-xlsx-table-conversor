import os
import pandas as pd
import openpyxl


def main():
    files_path = listFilesPath()

    for workbook_path in files_path:
        wb = openpyxl.Workbook()

        if isinstance(workbook_path, str):
            dataframe = pd.read_csv(workbook_path)
            worksheet = wb.active

            populateWorksheet(dataframe, worksheet)

            filename = extractFilenameFromUrl(workbook_path)

        else:
            for worksheet_path in workbook_path:
                dataframe = pd.read_csv(worksheet_path)
                worksheet = wb.create_sheet("Nova Planilha")

                populateWorksheet(dataframe, worksheet)

            filename = extractFilenameFromUrl(workbook_path[0])

            # delete a worksheet that is initialized as a workbook
            sheet = wb.active
            wb.remove(sheet)

        wb.save(f"out/{filename}.xlsx")


def populateWorksheet(dataframe, worksheet):
    head = dataframe.head(0).columns.to_list()

    worksheet.append(list(map(str.capitalize, head)))

    for _, row in dataframe.iterrows():
        new_row = [row[column] for column in head]
        worksheet.append(new_row)

    formatTableCells(worksheet)


def listFilesPath():
    files_directory = []
    data_path = os.listdir("./data")

    for workbook in data_path:
        file_path = os.path.join("./data", workbook)

        if file_path.endswith(".csv"):
            files_directory.append(file_path)

        else:
            worksheets = []
            workbook_files = os.listdir(file_path)

            for file in workbook_files:
                worksheet_path = os.path.join(file_path, file)
                worksheets.append(worksheet_path)

            files_directory.append(worksheets)

    return files_directory


def formatTableCells(sheet):
    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        sheet.column_dimensions[column_letter].width = adjusted_width


def extractFilenameFromUrl(url):
    index = url.find("data/") + len("data/")
    filename = url[index:].split("/")[0]
    filename = filename.split(".")[0]

    return filename


main()
