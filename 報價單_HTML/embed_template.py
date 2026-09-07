# -*- coding: utf-8 -*-
"""
把 範本/ 裡的 Word 範本轉成 base64，內嵌成 lib/templates.js。

瀏覽器用 file:// 開啟時無法 fetch 本機檔案（CSP 也封鎖），
因此範本必須內嵌進 JS 才能免安裝、雙擊即用。

改完範本後執行：python3 embed_template.py
"""
import base64
from pathlib import Path

BASE = Path(__file__).parent
OUT = BASE / "lib" / "templates.js"

TEMPLATES = [
    ("PAYMENT", "繳費單範本.docx", "繳費單（A5 橫式）"),
    ("RECEIPT", "收據範本.docx", "收據（A4 直式三聯）"),
]


def chunks(s, n=100):
    return [s[i:i + n] for i in range(0, len(s), n)]


def main():
    parts = [
        "/*!\n"
        " * templates.js ── 內建的 Word 範本（docxtemplater 語法）\n"
        " * 由 embed_template.py 從 範本/*.docx 自動產生，請勿手動編輯。\n"
        " * 想改版面：用 Word 改 範本/ 裡的檔案後重跑 embed_template.py，\n"
        " * 或直接在頁面上用「更換 Word 範本」載入。\n"
        " */\n"
    ]
    for var, filename, desc in TEMPLATES:
        path = BASE / "範本" / filename
        raw = path.read_bytes()
        b64 = base64.b64encode(raw).decode("ascii")
        lines = ",\n  ".join(f"'{c}'" for c in chunks(b64))
        parts.append(f"\n/* {desc} ── {filename}，{len(raw):,} bytes */\n"
                     f"var TEMPLATE_{var}_BASE64 = [\n  {lines}\n].join('');\n")
        print(f"  {desc:22s} {len(raw):>7,} bytes → base64 {len(b64):,} 字元")

    OUT.write_text("".join(parts), encoding="utf8")
    print(f"已寫入 {OUT}（{OUT.stat().st_size:,} bytes）")


if __name__ == "__main__":
    main()
