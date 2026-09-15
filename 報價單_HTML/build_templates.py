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
import re

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
    n = normalize_fonts(out)
    print(f"   字型統一：中文{CN_FONT}／英數{EN_FONT}（{n} 處）")
    n = fix_element_order(out)
    print(f"   子元素順序修正 {n} 處")
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

    真正的元凶是**行網格**：sectPr 的 <w:docGrid w:type="lines" w:linePitch="360"/>
    把每一行鎖在 18pt，不管字級多小。退費規定那格是 6.5pt、十幾行的細字，
    照字級算約 9mm，被網格撐成 89mm——三聯就是 267mm，單這一項就爆掉一頁。
    只縮字級完全沒有用，因為行高根本不看字級。

    第二個元凶是表格的浮動定位（tblpPr / tblpYSpec="center"）：三聯表格被當成
    浮動物件並垂直置中，只要比版心高一點點就整塊被擠到下一頁，上方還留一大片
    空白。改成一般表格後它會從版心頂端往下排，224mm 正好單頁。

    這裡做的事，都不動原稿的視覺設計：
      0. 關掉行網格（docGrid 改 default，並逐段加 snapToGrid=0）
      0-2. 表格改回非浮動，水平置中改用 w:jc
      1. 表格前後的空段落縮到 1pt
      2. 退費規定（6.5pt 細明文字）改為固定行高
      3. 「班別」欄加寬、姓名欄相應縮窄，讓常見的課程串接維持一行
    另外為每一列加上 cantSplit，萬一真的溢出也是整聯移動，不會從中間被切斷。
    """
    from docx.oxml.ns import qn
    from docx.shared import Mm

    body = doc.element.body

    # ⓪ 關掉行網格。docGrid 是「每行固定 linePitch」，字級縮再小行高也不變；
    #    type="default" 代表沒有網格，行高才會跟著字級走。
    #    另外逐段補 snapToGrid=0，避免任何樣式層級又把網格打開。
    for dg in body.iter(qn('w:docGrid')):
        dg.set(qn('w:type'), 'default')
    for para in body.iter(qn('w:p')):
        _el(_el(para, 'w:pPr'), 'w:snapToGrid', **{'w:val': '0'})

    # ① 表格前後的空段落縮到 1pt（w:sz 以半點為單位，故 val=2）
    for p in body.findall(qn('w:p')):
        if not ''.join(t.text or '' for t in p.iter(qn('w:t'))).strip():
            pPr = _el(p, 'w:pPr')
            _el(pPr, 'w:spacing', **{'w:after': '0', 'w:before': '0',
                                     'w:line': '20', 'w:lineRule': 'exact'})
            _el(_el(pPr, 'w:rPr'), 'w:sz', **{'w:val': '2'})

    table = doc.tables[0]

    # ⓪-2 把三聯表格從「浮動」改回一般表格。
    #     原稿的 tblPr 帶 <w:tblpPr tblpYSpec="center">，表格被當成浮動物件
    #     並在頁面上「垂直置中」。這種表格一旦比版心高一點就整塊往下一頁擠，
    #     上方還會留下一大片空白——實測 2 頁、頂端空 66mm。
    #     改成一般表格後，它就從版心頂端老實往下排：224mm，剛好單頁。
    #     表格寬 185mm 幾乎等於版心寬，所以水平置中改用 w:jc 保留即可。
    tblPr = table._tbl.find(qn('w:tblPr'))
    for tblp in tblPr.findall(qn('w:tblpPr')):
        tblPr.remove(tblp)
    _el(tblPr, 'w:jc', **{'w:val': 'center'})

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


def move_labels_into_body(path):
    """把「第N聯」「NO.」六個標籤從頁首／頁尾搬進本文，並改成相對頁面定位。

    這是收據一直跑版的主因。那六個標籤是絕對定位的浮動文字方塊，但它們
    錨定在頁首／頁尾的段落上（relativeFrom="paragraph"）。Word 與 LibreOffice
    都會把頁首撐到足以容納它們——最下面那個離頁首起點 296.7pt，頁首因此高達
    115mm，本文被壓到只剩 155mm，三聯無論如何縮都放不下。

    實測：把 headerReference／footerReference 拿掉，同一份檔案立刻從 2 頁變 1 頁、
    表格從 129.4mm 回到 13.3mm。但標籤不能跟著消失，所以搬進本文、改成
    relativeFrom="page" 的絕對座標，位置照舊、卻不再撐開任何東西。

    只把 style 改成 page 是不夠的（試過，仍是 2 頁）——錨點本身必須離開頁首。
    """
    import zipfile, shutil, tempfile
    from lxml import etree

    EMU = 12700.0                      # 1pt
    PAGE_H, LEFT_MARGIN = 841.89, 36.0    # A4 高、左邊界 720 twips

    zin = zipfile.ZipFile(str(path))
    parts = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    order = [i.filename for i in zin.infolist()]
    zin.close()

    doc = etree.fromstring(parts["word/document.xml"])
    nsmap = doc.nsmap
    W = "{%s}" % nsmap["w"]
    WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"

    body = doc.find(W + "body")
    sect = body.find(W + "sectPr")
    pgMar = sect.find(W + "pgMar")
    hdr_base = float(pgMar.get(W + "header") or 0) / 20.0
    ftr_base = PAGE_H - float(pgMar.get(W + "footer") or 0) / 20.0

    found = []
    for name, base in (("word/header1.xml", hdr_base),
                       ("word/footer1.xml", ftr_base)):
        if name not in parts:
            continue
        root = etree.fromstring(parts[name])
        for run in root.iter(W + "r"):
            anchor = None
            for a in run.iter(WP + "anchor"):
                anchor = a
                break
            if anchor is None:
                continue
            posv = anchor.find(WP + "positionV")
            off = posv.find(WP + "posOffset")
            y = base + float(off.text) / EMU
            txt = "".join(t.text or "" for t in run.iter(W + "t"))
            found.append({"run": run, "y": y, "no": "NO." in txt,
                          "text": txt, "anchor": anchor})
        parts[name] = etree.tostring(root, xml_declaration=True,
                                     encoding="UTF-8", standalone=True)

    if not found:
        return 0

    # 依原本的上下順序決定各屬於第幾聯：兩組（標籤／號碼）各自由上往下排
    copies = {}
    for kind in (False, True):
        group = sorted([f for f in found if f["no"] is kind], key=lambda f: f["y"])
        for i, f in enumerate(group):
            f["copy"] = i
    copies = max((f["copy"] for f in found), default=0) + 1

    # 錨點放在「該聯抬頭那一列」的儲存格裡，垂直方向相對於該段落。
    # 不用頁面絕對座標：列高宣告的是最小值，實際會被內容撐大（實測每聯
    # 比宣告多約 7mm），用絕對座標會一聯比一聯偏，第三聯會掉到上一聯的簽名列。
    # 錨在自己那一列上，就永遠跟著自己那一聯走。
    tbl = body.find(W + "tbl")
    hosts = []
    for tr in tbl.findall(W + "tr"):
        txt = "".join(t.text or "" for t in tr.iter(W + "t"))
        if "{branch_title}" in txt:
            tc = tr.find(W + "tc")
            hosts.append(tc.find(W + "p"))
    if len(hosts) < copies:
        hosts = [body.find(W + "p")] * copies

    for f in sorted(found, key=lambda f: (f["copy"], f["no"])):
        x = LEFT_MARGIN + _shape_left(f["anchor"], WP)
        y = -2.0 + (21.3 if f["no"] else 0.0)      # 相對抬頭段落
        _repos_anchor(f["anchor"], WP, x, y, EMU)
        _repos_vml(f["run"], W, x, y)
        hosts[f["copy"]].append(f["run"])

    # 錨點已經不在頁首／頁尾，參照拿掉，本文才會從版心頂端開始
    for ref in list(sect):
        if ref.tag in (W + "headerReference", W + "footerReference"):
            sect.remove(ref)

    parts["word/document.xml"] = etree.tostring(
        doc, xml_declaration=True, encoding="UTF-8", standalone=True)

    tmp = Path(tempfile.mkstemp(suffix=".docx")[1])
    with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zo:
        for name in order:
            zo.writestr(name, parts[name])
    shutil.move(str(tmp), str(path))
    return len(found)


def _shape_left(anchor, WP):
    """取出原本的水平位移（pt）。原本相對於文字欄，加上左邊界就是頁面座標。"""
    posh = anchor.find(WP + "positionH")
    off = posh.find(WP + "posOffset")
    return float(off.text) / 12700.0


def _repos_anchor(anchor, WP, x, y, EMU):
    from lxml import etree
    # 水平用頁面座標（標籤要落在版心右緣，不受儲存格限制），
    # 垂直相對於所在段落，才會跟著那一聯移動。
    anchor.set("layoutInCell", "0")
    for tag, rel, val in ((WP + "positionH", "page", x),
                          (WP + "positionV", "paragraph", y)):
        el = anchor.find(tag)
        el.set("relativeFrom", rel)
        off = el.find(WP + "posOffset")
        off.text = str(int(round(val * EMU)))
    # 改成完全不繞排。原本是 wrapSquare：搬進本文後，本文會繞著這六個標籤排，
    # 第三聯因此被撐開一百多 mm，照樣溢出第二頁。
    for kid in list(anchor):
        if kid.tag.startswith(WP + "wrap"):
            anchor.remove(kid)
    extent = anchor.find(WP + "extent")
    idx = list(anchor).index(extent) + 1 if extent is not None else len(anchor)
    eff = anchor.find(WP + "effectExtent")
    if eff is not None:
        idx = list(anchor).index(eff) + 1
    anchor.insert(idx, etree.SubElement(anchor, WP + "wrapNone"))


def _repos_vml(run, W, x, y):
    """VML 後備圖形也要一起改，否則舊版 Word 會畫在別的位置。"""
    import re as _re
    for el in run.iter():
        st = el.get("style")
        if not st or "position:absolute" not in st:
            continue
        st = _re.sub(r'margin-left:-?[\d.]+pt', 'margin-left:%.2fpt' % x, st)
        st = _re.sub(r'margin-top:-?[\d.]+pt', 'margin-top:%.2fpt' % y, st)
        st = st.replace("mso-position-horizontal-relative:text",
                        "mso-position-horizontal-relative:page")
        st = st.replace("mso-position-vertical-relative:text",
                        "mso-position-vertical-relative:paragraph")
        el.set("style", st)
    W10 = "{urn:schemas-microsoft-com:office:word}wrap"
    for el in run.iter(W10):
        el.set("type", "none")


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
    n = move_labels_into_body(out)
    print(f"   {n} 個「第N聯／NO.」標籤搬進本文並改為相對頁面定位")
    n = normalize_fonts(out)
    print(f"   字型統一：中文{CN_FONT}／英數{EN_FONT}（{n} 處）")
    n = fix_element_order(out)
    print(f"   子元素順序修正 {n} 處")
    print(f"收據範本 → {out}")


# ── 字型正規化 ────────────────────────────────────────────

CN_FONT = "標楷體"      # 中文
EN_FONT = "Times New Roman"   # 英文、數字、其他符號

_RFONTS_RE = re.compile(r'<w:rFonts\b([^>]*?)/>')
_ATTR_RE = re.compile(r'([\w:]+)="([^"]*)"')


def _rfonts_tag(attrs_text):
    """重寫單一 <w:rFonts/>：中文走標楷體，其餘走 Times New Roman。

    主題字型（asciiTheme 等）優先權高於明寫的字型名，因此必須拿掉，
    否則 Word 仍會用 theme1.xml 裡的字型。w:hint 保留，它決定
    全形標點這類「兩邊都說得通」的字元要算中文還是英文。
    """
    attrs = dict(_ATTR_RE.findall(attrs_text))
    hint = attrs.get("w:hint")
    out = ['w:ascii="%s"' % EN_FONT, 'w:hAnsi="%s"' % EN_FONT,
           'w:eastAsia="%s"' % CN_FONT, 'w:cs="%s"' % EN_FONT]
    if hint:
        out.append('w:hint="%s"' % hint)
    return "<w:rFonts %s/>" % " ".join(out)


def normalize_fonts(path):
    """把 docx 裡所有字型設定統一成：中文標楷體、英數符號 Times New Roman。

    三個地方都要改，缺一就會有文字漏網：
      document.xml  逐一 run 上的字型
      styles.xml    docDefaults 與各樣式（沒寫 rFonts 的 run 繼承這裡）
      header/footer 頁首頁尾
    """
    import zipfile, shutil, tempfile
    tmp = Path(tempfile.mkstemp(suffix=".docx")[1])
    zin = zipfile.ZipFile(str(path))
    count = 0
    with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zin.infolist():
            data = zin.read(item.filename)
            name = item.filename
            if (name == "word/document.xml" or name == "word/styles.xml"
                    or name.startswith("word/header")
                    or name.startswith("word/footer")):
                xml = data.decode("utf8")
                xml, n = _RFONTS_RE.subn(
                    lambda m: _rfonts_tag(m.group(1)), xml)
                count += n
                if name == "word/styles.xml":
                    xml, k = _ensure_default_rfonts(xml)
                    count += k
                data = xml.encode("utf8")
            zo.writestr(item, data)
    zin.close()
    shutil.move(str(tmp), str(path))
    return count


def _ensure_default_rfonts(xml):
    """docDefaults 若根本沒寫 rFonts，補一個，讓沒指定字型的文字也吃到。"""
    if "<w:rPrDefault>" not in xml:
        return xml, 0
    head, sep, tail = xml.partition("<w:rPrDefault>")
    if "<w:rFonts" in tail.split("</w:rPrDefault>")[0]:
        return xml, 0          # 已經有了，前面的 subn 改過
    tag = _rfonts_tag('w:hint="eastAsia"')
    block = tail.split("</w:rPrDefault>")[0]
    if "<w:rPr>" in block:
        tail = tail.replace("<w:rPr>", "<w:rPr>" + tag, 1)
    elif "<w:rPr/>" in block:
        tail = tail.replace("<w:rPr/>", "<w:rPr>" + tag + "</w:rPr>", 1)
    else:
        tail = "<w:rPr>" + tag + "</w:rPr>" + tail
    return head + sep + tail, 1


# ── OOXML 子元素順序 ──────────────────────────────────────

# WordprocessingML 的 schema 是 xsd:sequence，子元素順序寫錯，Word 會直接
# 忽略放錯位置的那個元素（LibreOffice 則照收）。先前收據的 w:spacing 被放到
# w:rPr 後面、w:cantSplit 被放到 w:trHeight 後面，於是所有壓縮行高的設定
# 在 Word 裡等於沒寫——排版量測全對，實際列印卻是兩頁，原因就在這裡。
CHILD_ORDER = {
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
            "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd",
            "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
            "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
            "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
            "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
            "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl",
            "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"],
    "rPr": ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
            "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
            "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
            "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
            "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
            "eastAsianLayout", "specVanish", "oMath", "rPrChange"],
    "trPr": ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore",
             "wAfter", "cantSplit", "trHeight", "tblHeader", "tblCellSpacing",
             "jc", "hidden", "ins", "del", "trPrChange"],
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders",
             "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign",
             "hideMark", "cellIns", "cellDel", "cellMerge", "tcPrChange"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
              "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
              "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
              "tblCellMar", "tblLook", "tblCaption", "tblDescription",
              "tblPrChange"],
    "sectPr": ["headerReference", "footerReference", "footnotePr", "endnotePr",
               "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType",
               "pgNumType", "cols", "formProt", "vAlign", "noEndnote",
               "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid",
               "printerSettings", "sectPrChange"],
}

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _reorder(el):
    """遞迴把子元素排回 schema 順序。含未知元素的節點一律不動，以免弄巧成拙。"""
    n = 0
    for child in el:
        n += _reorder(child)
    tag = el.tag
    if not isinstance(tag, str) or not tag.startswith(W_NS):
        return n
    order = CHILD_ORDER.get(tag[len(W_NS):])
    if order is None:
        return n
    kids = list(el)
    names = []
    for k in kids:
        if not isinstance(k.tag, str) or not k.tag.startswith(W_NS):
            return n                      # 有非 w: 元素，不碰
        names.append(k.tag[len(W_NS):])
    if any(name not in order for name in names):
        return n                          # 有沒列在表裡的元素，不碰
    idx = [order.index(name) for name in names]
    if idx == sorted(idx):
        return n
    for k in sorted(kids, key=lambda k: order.index(k.tag[len(W_NS):])):
        el.append(k)                      # append 會搬移既有節點
    return n + 1


def fix_element_order(path):
    """把 docx 各部分的子元素順序修正到 Word 能接受的樣子。"""
    import zipfile, shutil, tempfile
    from lxml import etree
    tmp = Path(tempfile.mkstemp(suffix=".docx")[1])
    zin = zipfile.ZipFile(str(path))
    fixed = 0
    with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zin.infolist():
            data = zin.read(item.filename)
            name = item.filename
            if (name in ("word/document.xml", "word/styles.xml")
                    or name.startswith("word/header")
                    or name.startswith("word/footer")):
                root = etree.fromstring(data)
                fixed += _reorder(root)
                data = etree.tostring(root, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
            zo.writestr(item, data)
    zin.close()
    shutil.move(str(tmp), str(path))
    return fixed


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    build_payment()
    build_receipt()
