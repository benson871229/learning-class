# -*- coding: utf-8 -*-
"""
套版程式：把一份報價資料灌進 範本/報價單範本.docx，輸出可再編輯的 Word。

計算規則（對照原稿）
    每列  總計 = 學費 + 教材費 − 扣除金額
    小計  各欄加總；金額為 0 的欄位留白（不印 0）
    扣除金額 以負號顯示（例：1000 → -1000）

可獨立當函式呼叫（給 Flask / 資料庫用），也可直接跑 JSON 產生範例。
"""
from __future__ import annotations

import json
from pathlib import Path

from docxtpl import DocxTemplate

BASE = Path(__file__).parent
TEMPLATE = BASE / "範本" / "報價單範本.docx"


def _num(v) -> int:
    """把可能為 None / '' / 字串的金額轉成整數，空值視為 0。"""
    if v in (None, "", "-"):
        return 0
    return int(round(float(v)))


def _money(n: int) -> str:
    """金額顯示：0 → 留白；其餘印整數（保留原稿無千分位風格）。"""
    return "" if n == 0 else f"{n}"


def _deduct(n: int) -> str:
    """扣除金額顯示：0 → 留白；其餘加負號。"""
    return "" if n == 0 else f"-{n}"


def build_context(quote: dict) -> dict:
    """把報價資料整理成範本需要的 context（含各列/小計的顯示字串）。"""
    courses_ctx = []
    sum_tuition = sum_material = sum_deduction = sum_total = 0

    for c in quote.get("courses", []):
        tuition = _num(c.get("tuition"))
        material = _num(c.get("material"))
        deduction = _num(c.get("deduction"))
        total = tuition + material - deduction

        sum_tuition += tuition
        sum_material += material
        sum_deduction += deduction
        sum_total += total

        courses_ctx.append({
            "name": c.get("name", ""),
            "date": c.get("date", ""),
            "tuition": _money(tuition),
            "material": _money(material),
            "deduction": _deduct(deduction),
            "total": _money(total),
        })

    return {
        "student_name": quote.get("student_name", ""),
        "courses": courses_ctx,
        "sum_tuition": _money(sum_tuition),
        "sum_material": _money(sum_material),
        "sum_deduction": _deduct(sum_deduction),
        "sum_total": _money(sum_total),
        "note": quote.get("note", ""),
    }


def render(quote: dict, out_path: str | Path) -> Path:
    """產生 Word。回傳輸出檔路徑。"""
    doc = DocxTemplate(TEMPLATE)
    doc.render(build_context(quote))
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    return out_path


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="產生報價單 Word")
    ap.add_argument("json", nargs="?", help="報價資料 JSON 檔（省略則用內建範例）")
    ap.add_argument("-o", "--out", help="輸出 .docx 路徑")
    args = ap.parse_args()

    if args.json:
        quote = json.loads(Path(args.json).read_text(encoding="utf-8"))
    else:
        quote = {
            "student_name": "陳楷元",
            "courses": [
                {"name": "課輔班", "date": "115/6/1-115/6/30", "tuition": 6000},
                {"name": "美語班", "date": "115/6/1-115/6/30", "tuition": 3000, "deduction": 1000},
            ],
            "note": "英文班每月計8堂課",
        }

    name = quote.get("student_name", "報價單")
    out = args.out or BASE / "輸出" / f"報價單_{name}.docx"
    path = render(quote, out)
    print(f"已產生：{path}")
