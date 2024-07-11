import os
from docx import Document
from openpyxl import Workbook

PATH = './data/'
OUTPUT_PATH = './output/res.xlsx'


def extract_project_info(docs):
    projects = []
    for doc in docs:
        project_name = ''
        link = ''

        for table in doc.tables:
            for row in table.rows:
                for idx, cell in enumerate(row.cells):
                    if 'https://pt.2035.university/project/' in cell.text:
                        link = cell.text.strip()

                    if 'Название стартап-проекта*' in cell.text:
                        if idx + 1 < len(row.cells):
                            project_name = row.cells[idx + 1].text.strip()

            if project_name:
                current_project = {'Название стартап-проекта': project_name, 'Ссылка': link}
                projects.append(current_project)
                project_name = ''
                link = ''
                print(current_project)

    return projects


def process_dir(path):
    docs = []

    for root, _, files in os.walk(path):
        for file in files:
            if file.endswith('.docx'):
                word_path = os.path.join(root, file)
                if isinstance(word_path, str):
                    docs.append(Document(word_path))

    return extract_project_info(docs)


def save_to_excel(data, excel_filename):
    wb = Workbook()
    ws = wb.active

    headers = ['Название стартап-проекта', 'Ссылка']
    ws.append(headers)

    for item in data:
        row = [item['Название стартап-проекта'], item['Ссылка']]
        ws.append(row)

    wb.save(excel_filename)


if __name__ == '__main__':
    save_to_excel(process_dir(PATH), OUTPUT_PATH)
