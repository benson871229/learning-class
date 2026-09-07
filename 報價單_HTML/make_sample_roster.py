# -*- coding: utf-8 -*-
"""產生「報名清單」範例 Excel（批次匯入用），供測試與示範。

每一列是一筆課程，同一位學生多門課就多列，工具會依姓名自動合併成一份單據。
執行：python3 make_sample_roster.py [學生人數]
輸出：範例資料/報名清單.xlsx
"""
import html
import random
import sys
import zipfile
from pathlib import Path

OUT = Path(__file__).parent / "範例資料" / "報名清單.xlsx"
HEAD = ['學生姓名', '課程名稱', '日期起訖', '學費', '教材費', '扣除金額', '開立日期', '備註']
SUR = '陳林黃張李王吳劉蔡楊許鄭謝洪郭曾廖賴徐周'
GIV = ['楷元', '曉薇', '冠廷', '雅婷', '宗翰', '怡君', '家豪', '淑芬', '俊傑', '詩涵',
       '柏翰', '思妤', '承恩', '宜蓁', '建宏', '佩珊', '品睿', '若瑄', '子軒', '欣怡']
COURSES = [('課輔班', 6000, 2000), ('美語班', 3000, 2500), ('數學班', 4000, 500),
           ('作文班', 2800, 300), ('理化班', 4500, 400)]


def col(i):
    s = ''
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def sheet_xml(rows):
    out = []
    for r, row in enumerate(rows, 1):
        cells = []
        for c, v in enumerate(row):
            ref = f'{col(c)}{r}'
            if isinstance(v, int) and not isinstance(v, bool):
                cells.append(f'<c r="{ref}"><v>{v}</v></c>')
            else:
                cells.append(f'<c r="{ref}" t="inlineStr"><is>'
                             f'<t xml:space="preserve">{html.escape(str(v))}</t></is></c>')
        out.append(f'<row r="{r}">{"".join(cells)}</row>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<sheetData>{"".join(out)}</sheetData></worksheet>')


def build(n_students=70, seed=7):
    random.seed(seed)
    rows = [HEAD]
    for i in range(n_students):
        name = SUR[i % len(SUR)] + GIV[(i * 7) % len(GIV)] + ('' if i < 20 else str(i))
        for cname, tuition, material in random.sample(COURSES, random.randint(1, 3)):
            rows.append([name, cname, '115/9/1-115/9/30', tuition, material,
                         1000 if random.random() < .25 else 0, '115/9/1',
                         '英文班每月計8堂課' if cname == '美語班' else ''])

    S = lambda x: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + x
    wb = S('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
           ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           '<sheets><sheet name="報名清單" sheetId="1" r:id="rId1"/></sheets></workbook>')
    wbrels = S('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
               'relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
    ct = S('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
           '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
           '<Default Extension="xml" ContentType="application/xml"/>'
           '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-'
           'officedocument.spreadsheetml.sheet.main+xml"/>'
           '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-'
           'officedocument.spreadsheetml.worksheet+xml"/></Types>')
    rels = S('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
             'relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', rels)
        z.writestr('xl/workbook.xml', wb)
        z.writestr('xl/_rels/workbook.xml.rels', wbrels)
        z.writestr('xl/worksheets/sheet1.xml', sheet_xml(rows))
    print(f'{OUT} → {n_students} 位學生、{len(rows) - 1} 筆課程列')


if __name__ == '__main__':
    build(int(sys.argv[1]) if len(sys.argv) > 1 else 70)
