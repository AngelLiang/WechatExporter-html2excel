import os
import re
import glob
from bs4 import BeautifulSoup
from openpyxl import Workbook

ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')


INPUT_PATH = glob.glob('input_data/*.html')
OUTPUT_PATH = 'output'

os.makedirs(OUTPUT_PATH, exist_ok=True)

for filepath in INPUT_PATH:
    html_doc = open(filepath, encoding='utf8')
    soup = BeautifulSoup(html_doc, 'html.parser')
    all_data = soup.find_all('span', class_='dspname left')

    wb = Workbook()
    ws = wb.active
    for item in all_data:
        # 名称
        # print(f'{item.text} {item.next_sibling}')
        content = item.parent.next_sibling.next_sibling.find('span')
        text = ''
        if content:
            text = content.text
            text = ILLEGAL_CHARACTERS_RE.sub(r'', text)
        ws.append([item.next_sibling, item.text, text])

    name = os.path.basename(filepath)
    output_filepath = os.path.join(OUTPUT_PATH, name+'.xlsx')
    wb.save(output_filepath)
