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

    if len(text.split(SEPARATOR)) < 2:
        text = text.replace('Тема стартап-проекта*', SEPARATOR)
    print(len(text.split(SEPARATOR)))
    for passport in text.split(SEPARATOR):
        startup = {'Название стартап-проекта': '', 'Ссылка': ''}

        passport_parts = passport.split(BREAK)
        for part in passport_parts:
            stripped_part = ''.join(part).replace(' ', '')
            if 'https://pt.2035' in stripped_part:
                startup['Ссылка'] = (
                    stripped_part
                    .replace('\xad', '-')
                    .replace('е', 'e')
                    .replace('Т', 'T')
                    .replace('у', 'y')
                    .replace('о', 'o')
                    .replace('р', 'p')
                    .replace('А', 'A')
                    .replace('а', 'a')
                    .replace('Н', 'H')
                    .replace('О', 'O')
                    .replace('Р', 'P')
                    .replace('К', 'K')
                    .replace('Х', 'X')
                    .replace('x', 'х')
                    .replace('В', 'B')
                    .replace('М', 'M')
                    .replace('г', 'r')
                    .replace('Р', 'P')
                    .replace('С', 'C')
                    .replace('с', 'c')
                    .replace('Е', 'E')
                    .replace('ь', 'b')
                    .replace('т', 'T')
                    .replace('п', 'n')
                    .replace('У', 'Y')
                    .replace('З', '3')
                    .replace('з', '3')
                    .replace('ш', 'w')
                    .replace('Ш', 'W')
                    .replace('Ы', 'bl')
                    .replace('ы', 'bl')
                    .replace('ц', 'u')
                    .replace('Ц', 'U')
                    .replace('Ь', 'b')
                    .split('\t'))[0]

                next_line = passport_parts[passport_parts.index(part) + 1]
                if next_line and not next_line[0].isdigit():
                    startup['Ссылка'] += next_line
            elif NAME_KEY in ''.join(part):
                name_key_index = passport_parts.index(NAME_KEY)
                name_shift = 1
                while not passport_parts[name_key_index + name_shift]:
                    name_shift += 1
                startup['Название стартап-проекта'] = passport_parts[name_key_index + name_shift]

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
            if type(block).__name__ == 'CT_Tbl':
                table_text = []
                table = Table(block, doc)
                for row in table.rows:
                    prior_tc = None
                    for cell in row.cells:
                        if cell._tc == prior_tc:
                            continue
                        prior_tc = cell._tc
                        for paragraph in cell.paragraphs:
                            table_text.append(paragraph.text + BREAK)
                            print(paragraph.text)

                full_text.append(''.join(table_text))
            elif block.text:
                print(block.text)
                full_text.append(block.text.replace(
                    'Паспорт стартап-проекта', SEPARATOR) + BREAK)

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
