# -*- coding: utf-8 -*-
"""建立資料庫並灌入範例資料（對照原稿：陳楷元 / 課輔班 / 美語班）。"""
import db


def run():
    db.init_db()

    if db.list_students() or db.list_courses():
        print("資料庫已有資料，略過 seed。")
        return

    # 課程／價目表
    db.add_course(name="課輔班", default_tuition=6000, default_material=0, note="")
    db.add_course(name="美語班", default_tuition=3000, default_material=0,
                  note="英文班每月計8堂課")
    db.add_course(name="數學班", default_tuition=4000, default_material=500, note="")

    # 學生名冊
    sid = db.add_student(name="陳楷元", grade="國小三年級", parent_name="陳先生",
                         phone="0912-345-678", email="", address="", note="")

    # 一張範例報價（對照原稿）
    qid = db.create_quote(
        student_id=sid, student_name="陳楷元",
        issue_date="115/6/1", note="英文班每月計8堂課",
        items=[
            {"name": "課輔班", "date_range": "115/6/1-115/6/30", "tuition": 6000},
            {"name": "美語班", "date_range": "115/6/1-115/6/30", "tuition": 3000,
             "deduction": 1000},
        ],
    )
    print(f"seed 完成：學生 {sid}、報價 {qid}")


if __name__ == "__main__":
    run()
