# -*- coding: utf-8 -*-
"""
把 範本/報價單範本.docx 內嵌成 lib/template-b5.js（base64）。

為什麼要內嵌？
    直接用瀏覽器開啟本機 HTML（file://）時，瀏覽器基於安全限制不允許用
    fetch 讀取旁邊的檔案。把預設範本內嵌成 JS 就能「雙擊即用」，
    完全不需要架伺服器。

想換版面仍然自由：頁面上的「更換範本」可以載入你自己用 Word 改過的 .docx；
或改完範本後重跑本程式，更新內建的預設範本。
"""
import base64
from pathlib import Path

BASE = Path(__file__).parent
SRC = BASE / "範本" / "報價單範本.docx"
DST = BASE / "lib" / "template-b5.js"


def main():
    data = SRC.read_bytes()
    b64 = base64.b64encode(data).decode("ascii")

    lines = [b64[i:i + 120] for i in range(0, len(b64), 120)]
    body = "\n".join(f"  '{ln}' +" for ln in lines[:-1])
    body += f"\n  '{lines[-1]}';"

    js = (
        "/*!\n"
        " * template-b5.js ── 內建的預設 Word 範本（B5 橫式 257x182mm，docxtemplater 語法）\n"
        " * 由 embed_template.py 從 範本/報價單範本.docx 自動產生，請勿手動編輯。\n"
        " * 想改版面：用 Word 改 範本/報價單範本.docx 後重跑 embed_template.py，\n"
        " * 或直接在頁面上用「更換範本」載入你自己的 .docx。\n"
        " */\n"
        "var TEMPLATE_B5_BASE64 =\n"
        f"{body}\n"
    )
    DST.write_text(js, encoding="utf-8")
    print(f"已內嵌範本：{DST}（{len(data):,} bytes → base64 {len(b64):,} 字元）")


if __name__ == "__main__":
    main()
