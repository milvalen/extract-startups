import os
from docx import Document
from docx.table import Table
from openpyxl import Workbook

PATH = './data/'
OUTPUT_PATH = './output/res.xlsx'
SEPARATOR = 'ПАСПОРТ СТАРТАП-ПРОЕКТА'
BREAK = '-textbreak-'
NAME_KEY = 'Название стартап-проекта*'


def extract_startups(text):
    startups = []
    print(len(text.split(SEPARATOR)))
    for passport in text.split(SEPARATOR):
        startup = {'Название стартап-проекта': '', 'Ссылка': ''}

        passport_parts = passport.split(BREAK)
        for part in passport_parts:
            if 'https://pt.2035.university/project/' in ''.join(part):
                startup['Ссылка'] = ''.join(part).replace(' ', '').split('\t')[0]
            elif NAME_KEY in ''.join(part):
                startup['Название стартап-проекта'] = passport_parts[
                    passport_parts.index(NAME_KEY) + 1]
        if startup['Название стартап-проекта']:
            print(startup)
            startups.append(startup)

    return startups


def iter_block_items(doc):
    for block in doc.element.body:
        yield block


def extract_text_in_order(docs):
    full_text = []
    for doc in docs:
        for block in iter_block_items(doc):
            if type(block).__name__ == 'CT_P':
                full_text.append(block.text.replace(
                    'Паспорт стартап-проекта', SEPARATOR) + BREAK)
            elif type(block).__name__ == 'CT_Tbl':
                table_text = []
                table = Table(block, doc)
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            table_text.append(paragraph.text + BREAK)
                full_text.append(''.join(table_text))

    return extract_startups(''.join(full_text))


def process_dir(path):
    docs = []

    for root, _, files in os.walk(path):
        for file in files:
            print(file)
            if file.endswith('.docx'):
                word_path = os.path.join(root, file)
                if isinstance(word_path, str):
                    docs.append(Document(word_path))

    return extract_text_in_order(docs)


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
