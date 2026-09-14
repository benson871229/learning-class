# -*- coding: utf-8 -*-
"""
把使用者提供的 Word 原稿轉成 docxtemplater 範本。

刻意「沿用原稿再置換」而不是重畫：字體、框線、退費規定長文、三聯版面
全部原封不動保留，只把範例值換成佔位符。

輸入（範本/原始檔/）        輸出（範本/）
    繳費單_原稿.docx     →   繳費單範本.docx    A5 橫式，課程列可重複
    收據_壽豐路_原稿.docx →   收據範本.docx      A4 直式三聯，分校資訊為變數

執行：python3 build_templates.py
"""
from pathlib import Path
import copy

from docx import Document

BASE = Path(__file__).parent
SRC = BASE / "範本" / "原始檔"
OUT = BASE / "範本"


# ── 置換工具 ──────────────────────────────────────────────

def para_text(p):
    return "".join(r.text for r in p.runs)


def replace_in_para(p, old, new):
    """在段落中置換文字，保留第一個相關 run 的格式。

    文字常被 Word 拆散在多個 run（例如注音、拼字檢查造成的分割），
    因此比對用整段文字，改寫時把結果塞回第一個涉及的 run、清空其餘。
    """
    runs = p.runs
    if not runs:
        return False
    full = "".join(r.text for r in runs)
    idx = full.find(old)
    if idx == -1:
        return False

    # 找出涉及的 run 區間
    starts, pos = [], 0
    for r in runs:
        starts.append(pos)
        pos += len(r.text)
    end = idx + len(old)
    first = last = None
    for i, r in enumerate(runs):
        s, e = starts[i], starts[i] + len(r.text)
        if first is None and e > idx:
            first = i
        if s < end:
            last = i
    head = runs[first].text[: idx - starts[first]]
    tail = runs[last].text[end - starts[last]:]
    runs[first].text = head + new + tail
    for i in range(first + 1, last + 1):
        runs[i].text = ""
    return True


def replace_everywhere(doc, old, new):
    """整份文件（含表格、巢狀表格）置換，回傳置換次數。"""
    n = 0
    for p in doc.paragraphs:
        while replace_in_para(p, old, new):
            n += 1
    for t in doc.tables:
        n += _replace_in_table(t, old, new)
    return n


def _replace_in_table(t, old, new):
    n = 0
    for row in t.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                while replace_in_para(p, old, new):
                    n += 1
            for sub in cell.tables:
                n += _replace_in_table(sub, old, new)
    return n


def set_cell(cell, text, like=None):
    """把儲存格內容換成 text。空白格沒有 run 可改，就複製 like 的格式新建一個。"""
    p = cell.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
        return
    src = None
    if like is not None and like.paragraphs and like.paragraphs[0].runs:
        src = like.paragraphs[0].runs[0]
    if src is not None:
        new_r = copy.deepcopy(src._element)
        p._p.append(new_r)
        p.runs[-1].text = text
    else:
        p.add_run(text)


def drop_row(table, row):
    row._tr.getparent().remove(row._tr)


# ── 繳費單 ────────────────────────────────────────────────

def build_payment():
    doc = Document(SRC / "繳費單_原稿.docx")

    replace_everywhere(doc, "王大明", "{student_name}")

    t = doc.tables[0]
    rows = t.rows
    # 版面：0 表頭 / 1、2 兩筆範例課程 / 3 空白列 / 4 小計 / 5 備註
    data, sample2, subtotal, note = rows[1], rows[2], rows[4], rows[5]

    # 課程列改成可重複列：docxtemplater 的 {#courses} … {/courses} 放在同列頭尾格
    vals = ["{#courses}{name}", "{date}", "{tuition}", "{material}",
            "{deduction}", "{total}{/courses}"]
    for cell, v in zip(data.cells, vals):
        set_cell(cell, v, like=data.cells[0])

    drop_row(t, sample2)          # 第二筆範例交給迴圈產生

    # 注意：row.cells 會把合併儲存格重複展開。「小計」那格 gridSpan=2，
    # 因此 cells[0] 與 cells[1] 都指向它，四個金額格是 cells[2:6]。
    for cell, v in zip(subtotal.cells[2:6],
                       ["{sum_tuition}", "{sum_material}",
                        "{sum_deduction}", "{sum_total}"]):
        set_cell(cell, v, like=subtotal.cells[2])

    for p in note.cells[0].paragraphs:
        replace_in_para(p, "英文班每月計8堂課", "{note}")

    out = OUT / "繳費單範本.docx"
    doc.save(out)
    print(f"繳費單範本 → {out}")


def _el(parent, tag, **attrs):
    """在 parent 底下取得或建立子元素（tag 用 w: 前綴短名）。"""
    from docx.oxml.ns import qn
    q = qn(tag)
    el = parent.find(q)
    if el is None:
        el = parent.makeelement(q, {})
        parent.append(el)
    for k, v in attrs.items():
        el.set(qn(k), v)
    return el


def fit_receipt_on_one_page(doc):
    """讓三聯確實印在同一頁。

    原稿的三聯表格約 266mm，加上表格前的空段落（12pt，約 5.5mm）就會超過
    A4 版心的 271mm，只差零點幾毫米——任何一點內容變動都會把第三聯擠到第二頁。
    而「班別」欄僅 57.5mm（約 12 字），一旦學生報兩門以上，課程名稱串接後就會
    折行，三聯各多一行又是 16mm，必定跑版。

    這裡做三件事把餘裕拉回來，都不動原稿的視覺設計：
      1. 表格前後的空段落縮到 1pt
      2. 退費規定（6.5pt 細明文字）改為固定行高，省下每聯約 3mm
      3. 「班別」欄加寬、姓名欄相應縮窄，讓常見的課程串接維持一行
    另外為每一列加上 cantSplit，萬一真的溢出也是整聯移動，不會從中間被切斷。
    """
    from docx.oxml.ns import qn
    from docx.shared import Mm

    body = doc.element.body

    # ① 表格前後的空段落縮到 1pt（w:sz 以半點為單位，故 val=2）
    for p in body.findall(qn('w:p')):
        if not ''.join(t.text or '' for t in p.iter(qn('w:t'))).strip():
            pPr = _el(p, 'w:pPr')
            _el(pPr, 'w:spacing', **{'w:after': '0', 'w:before': '0',
                                     'w:line': '20', 'w:lineRule': 'exact'})
            _el(_el(pPr, 'w:rPr'), 'w:sz', **{'w:val': '2'})

    table = doc.tables[0]

    for row in table.rows:
        # ② 整列不跨頁
        _el(_el(row._tr, 'w:trPr'), 'w:cantSplit')

        # 直接操作底層 w:tc：row.cells 會把合併儲存格依所跨欄數重複展開，
        # 這個表格用了 gridSpan，取 row.cells 會拿到 9 格而不是實際的 4 格。
        tcs = row._tr.findall(qn('w:tc'))
        texts = [''.join(t.text or '' for t in tc.iter(qn('w:t'))) for tc in tcs]

        # ③ 退費規定：固定行高，避免 6.5pt 細字被行距撐開
        for tc, txt in zip(tcs, texts):
            if '補習班於學生繳納費用後' in txt:
                for para in tc.findall(qn('w:p')):
                    pPr = _el(para, 'w:pPr')
                    _el(pPr, 'w:spacing', **{'w:after': '0', 'w:before': '0',
                                             'w:line': '140', 'w:lineRule': 'exact'})

        # ④ 三聯之間的分隔空列從 7.8mm 收到 4mm。裁切線仍看得出來，
        #    但多出的 7.6mm 足以吸收「班別」萬一折行時多出的高度。
        if not any(t.strip() for t in texts):
            th = _el(_el(row._tr, 'w:trPr'), 'w:trHeight')
            th.set(qn('w:val'), str(int(Mm(4).twips)))
            th.set(qn('w:hRule'), 'exact')

        # ⑤ 「班別」欄加寬、姓名欄相應縮窄，維持整列仍是 185mm
        if texts and texts[0].strip() == '學生姓名' and len(tcs) == 4:
            for tc, mm in zip(tcs, (30, 55, 25, 75)):
                tcPr = _el(tc, 'w:tcPr')
                _el(tcPr, 'w:tcW', **{'w:w': str(int(Mm(mm).twips)), 'w:type': 'dxa'})


def unwrap_header_shapes(path):
    """讓頁首的浮動標籤不要把本文往下推。

    頁首那四個「第N聯」「NO.」標籤是絕對定位的 VML 圖形，卻帶著
    <w10:wrap type="square"/>（文繞圖），Word 會讓本文避開它們，
    於是三聯整個被往下擠，頁面上方留下一大片空白。

    要改成 type="none"（維持浮動、不繞排），不能把 <w10:wrap> 整個刪掉——
    少了這個元素，Word 會把圖形當成內嵌物件，頁首反而被撐高，內容更容易溢出。
    """
    import zipfile, shutil, tempfile, os
    zin = zipfile.ZipFile(path)
    tmp = path.with_suffix('.tmp')
    n = 0
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('word/header'):
                xml = data.decode('utf8')
                n += xml.count('<w10:wrap type="square"/>')
                xml = xml.replace('<w10:wrap type="square"/>',
                                  '<w10:wrap type="none"/>')
                data = xml.encode('utf8')
            zo.writestr(item, data)
    zin.close()
    shutil.move(str(tmp), str(path))
    return n


# ── 收據 ──────────────────────────────────────────────────

def build_receipt():
    doc = Document(SRC / "收據_壽豐路_原稿.docx")

    # 由長至短置換，避免短字串先吃掉長字串的一部分
    pairs = [
        ("高雄市私立Fun學院文理短期補習班壽豐分班", "{branch_title}"),
        ("高雄市楠梓區壽豐路302號", "{branch_addr}"),
        ("自115年9月1日起至115年9月30日止", "{period}"),
        ("新台幣伍仟陆佰元整", "新台幣{amount_cn}元整"),
        ("5600元", "{amount}元"),
        ("王大明", "{student_name}"),
        ("美語班", "{class_name}"),
    ]
    for old, new in pairs:
        n = replace_everywhere(doc, old, new)
        print(f"   {old[:22]:24s} → {new:18s} 置換 {n} 處")

    fit_receipt_on_one_page(doc)
    print("   已套用單頁排版調整（縮空段落／固定細字行高／加寬班別欄／列不跨頁）")

    out = OUT / "收據範本.docx"
    doc.save(out)
    n = unwrap_header_shapes(out)
    print(f"   頁首 {n} 個浮動標籤改為不繞排（原本會把本文往下推約 40mm）")
    print(f"收據範本 → {out}")


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    build_payment()
    build_receipt()
