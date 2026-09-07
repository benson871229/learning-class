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

    out = OUT / "收據範本.docx"
    doc.save(out)
    print(f"收據範本 → {out}")


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    build_payment()
    build_receipt()
