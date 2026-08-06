# -*- coding: utf-8 -*-
"""
本機報價管理介面（Flask，開在 localhost）。

啟動：
    cd 報價系統
    python3 app.py
然後用瀏覽器開 http://127.0.0.1:5000

功能：學生名冊、課程／價目表、開立報價單（產出 Word）、報價歷史、繳費狀態。
只跑在自己的電腦，不對外開放。
"""
from __future__ import annotations

import io
from datetime import datetime

from flask import (Flask, render_template, request, redirect, url_for,
                   send_file, flash)

import db
import quote_docx

app = Flask(__name__)
app.secret_key = "buxiban-local-tool"

PAYMENT_STATUSES = ["未繳", "部分繳", "已繳"]
PAYMENT_METHODS = ["", "現場繳費", "線上匯款"]


@app.route("/")
def index():
    quotes = db.list_quotes()
    students = db.list_students()
    unpaid = [q for q in quotes if q["payment_status"] != "已繳"]
    revenue = sum(q["sum_total"] for q in quotes if q["payment_status"] == "已繳")
    return render_template("index.html", quotes=quotes[:8], n_students=len(students),
                           n_quotes=len(quotes), n_unpaid=len(unpaid), revenue=revenue)


# ─────────────────── 學生名冊 ───────────────────
@app.route("/students")
def students():
    return render_template("students.html", students=db.list_students())


@app.route("/students/new", methods=["GET", "POST"])
@app.route("/students/<int:sid>/edit", methods=["GET", "POST"])
def student_form(sid=None):
    student = db.get_student(sid) if sid else None
    if request.method == "POST":
        f = {k: request.form.get(k, "").strip() for k in
             ("name", "grade", "parent_name", "phone", "email", "address", "note")}
        if not f["name"]:
            flash("請填寫學生姓名", "danger")
        else:
            if sid:
                db.update_student(sid, **f)
                flash("已更新學生資料", "success")
            else:
                db.add_student(**f)
                flash("已新增學生", "success")
            return redirect(url_for("students"))
    return render_template("student_form.html", student=student)


@app.route("/students/<int:sid>/delete", methods=["POST"])
def student_delete(sid):
    db.delete_student(sid)
    flash("已刪除學生", "success")
    return redirect(url_for("students"))


# ─────────────────── 課程／價目表 ───────────────────
@app.route("/courses")
def courses():
    return render_template("courses.html", courses=db.list_courses())


@app.route("/courses/new", methods=["GET", "POST"])
@app.route("/courses/<int:cid>/edit", methods=["GET", "POST"])
def course_form(cid=None):
    course = db.get_course(cid) if cid else None
    if request.method == "POST":
        f = {
            "name": request.form.get("name", "").strip(),
            "default_tuition": request.form.get("default_tuition") or 0,
            "default_material": request.form.get("default_material") or 0,
            "note": request.form.get("note", "").strip(),
            "active": 1 if request.form.get("active") else 0,
        }
        if not f["name"]:
            flash("請填寫課程名稱", "danger")
        else:
            if cid:
                db.update_course(cid, **f)
                flash("已更新課程", "success")
            else:
                db.add_course(**f)
                flash("已新增課程", "success")
            return redirect(url_for("courses"))
    return render_template("course_form.html", course=course)


@app.route("/courses/<int:cid>/delete", methods=["POST"])
def course_delete(cid):
    db.delete_course(cid)
    flash("已刪除課程", "success")
    return redirect(url_for("courses"))


# ─────────────────── 開立報價單 ───────────────────
@app.route("/quotes/new", methods=["GET", "POST"])
def quote_new():
    if request.method == "POST":
        student_id = request.form.get("student_id") or None
        student_name = request.form.get("student_name", "").strip()
        issue_date = request.form.get("issue_date", "").strip()
        note = request.form.get("note", "").strip()

        names = request.form.getlist("item_name")
        dates = request.form.getlist("item_date")
        tuitions = request.form.getlist("item_tuition")
        materials = request.form.getlist("item_material")
        deductions = request.form.getlist("item_deduction")

        items = []
        for i, nm in enumerate(names):
            if not nm.strip():
                continue
            items.append({
                "name": nm.strip(),
                "date_range": dates[i].strip() if i < len(dates) else "",
                "tuition": tuitions[i] or 0,
                "material": materials[i] or 0,
                "deduction": deductions[i] or 0,
            })

        if not student_name:
            flash("請填寫學生姓名", "danger")
        elif not items:
            flash("請至少新增一筆課程", "danger")
        else:
            qid = db.create_quote(
                student_id=int(student_id) if student_id else None,
                student_name=student_name, issue_date=issue_date, note=note,
                items=items,
                payment_status=request.form.get("payment_status", "未繳"),
                payment_method=request.form.get("payment_method") or None,
            )
            flash("報價單已建立", "success")
            return redirect(url_for("quote_view", qid=qid))

    return render_template("quote_new.html",
                           students=db.list_students(),
                           courses=db.list_courses(only_active=True),
                           today=datetime.now().strftime("%Y/%m/%d"),
                           statuses=PAYMENT_STATUSES, methods=PAYMENT_METHODS)


# ─────────────────── 報價歷史 ───────────────────
@app.route("/quotes")
def quotes():
    return render_template("quotes.html", quotes=db.list_quotes(),
                           statuses=PAYMENT_STATUSES)


@app.route("/quotes/<int:qid>")
def quote_view(qid):
    q = db.get_quote(qid)
    if not q:
        flash("找不到報價單", "danger")
        return redirect(url_for("quotes"))
    return render_template("quote_view.html", q=q, statuses=PAYMENT_STATUSES,
                           methods=PAYMENT_METHODS)


@app.route("/quotes/<int:qid>/docx")
def quote_docx_download(qid):
    q = db.get_quote(qid)
    if not q:
        flash("找不到報價單", "danger")
        return redirect(url_for("quotes"))
    ctx = db.quote_to_docx_ctx(qid)
    doc = quote_docx.DocxTemplate(quote_docx.TEMPLATE)
    doc.render(quote_docx.build_context(ctx))
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    fname = f"報價單_{q['student_name']}_{q['quote_no']}.docx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/quotes/<int:qid>/payment", methods=["POST"])
def quote_payment(qid):
    db.update_payment(qid,
                      payment_status=request.form.get("payment_status", "未繳"),
                      payment_method=request.form.get("payment_method") or None,
                      receipt_no=request.form.get("receipt_no", "").strip() or None)
    flash("已更新繳費狀態", "success")
    return redirect(url_for("quote_view", qid=qid))


@app.route("/quotes/<int:qid>/delete", methods=["POST"])
def quote_delete(qid):
    db.delete_quote(qid)
    flash("已刪除報價單", "success")
    return redirect(url_for("quotes"))


if __name__ == "__main__":
    db.init_db()
    app.run(debug=True, port=5000)
