import os
from docx import Document
from docx.table import Table
from openpyxl import Workbook

PATH = './data/'
OUTPUT_PATH = './output/res.xlsx'


def extract_startups(text):
    startups = []
    for passport in text.split('ПАСПОРТ СТАРТАП-ПРОЕКТА'):
        startup = {'Название стартап-проекта': '', 'Ссылка': ''}
        for text in passport.split('-textbreak-'):
            if 'https://pt.2035.university/project/' in ''.join(text):
                startup['Ссылка'] = ''.join(text).replace(' ', '').split('\t')[0]
            elif 'Название стартап-проекта*' in ''.join(text):
                startup['Название стартап-проекта'] = passport.split('-textbreak-')[
                    passport.split('-textbreak-').index('Название стартап-проекта*') + 1]
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
                full_text.append(block.text + '-textbreak-')
            elif type(block).__name__ == 'CT_Tbl':
                table_text = []
                table = Table(block, doc)
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            table_text.append(paragraph.text + '-textbreak-')
                full_text.append(''.join(table_text))

    return extract_startups(''.join(full_text))


def process_dir(path):
    docs = []

    for root, _, files in os.walk(path):
        for file in files:
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
