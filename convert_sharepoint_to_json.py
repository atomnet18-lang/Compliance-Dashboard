#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convert_sharepoint_to_json.py
แปลงข้อมูลที่ผู้ปฏิบัติงานกรอกใน SharePoint List (export เป็น .xlsx/.csv)
ให้เป็น data.json สำหรับ Dashboard Compliance 2569 (GitHub Pages)

โครงสร้าง List ที่คาดหวัง (1 แถว = 1 กิจกรรม) คอลัมน์:
  AID         : รหัสกิจกรรม เช่น 1-01   (KEY — ห้ามแก้)
  StartMonth  : เดือนเริ่ม   เช่น ม.ค.70  (ต้องตรงกับ months)
  EndMonth    : เดือนสิ้นสุด เช่น มี.ค.70
  Status      : ยังไม่เริ่ม | กำลังดำเนินการ | เสร็จสิ้น
  Note        : หมายเหตุภายใน (ไม่ถูกเผยแพร่ — ตัดทิ้งตอนแปลง)

วิธีใช้:
  python3 convert_sharepoint_to_json.py <list_export.xlsx|csv> [data.json]
สคริปต์จะอ่าน template โครงสร้างจาก data.json เดิม (ชื่อกิจกรรม/ผู้รับผิดชอบ/months)
แล้วเติมเฉพาะ start/end/status จาก List — ไม่นำ Note (field ภายใน) ออกเผยแพร่
"""
import sys, json, os, datetime

STATUS_MAP = {
    "ยังไม่เริ่ม": "none", "none": "none", "": "none",
    "กำลังดำเนินการ": "prog", "กำลังทำ": "prog", "prog": "prog",
    "เสร็จสิ้น": "done", "เสร็จ": "done", "done": "done",
}

def load_rows(path):
    ext = os.path.splitext(path)[1].lower()
    rows = []
    if ext in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        header = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
        idx = {h: i for i, h in enumerate(header)}
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r is None or all(v is None for v in r):
                continue
            rows.append({h: (r[idx[h]] if idx[h] < len(r) else None) for h in header})
    else:
        import csv
        with open(path, encoding="utf-8-sig") as f:
            for d in csv.DictReader(f):
                rows.append(d)
    return rows

def main():
    if len(sys.argv) < 2:
        print("usage: python3 convert_sharepoint_to_json.py <list_export.xlsx|csv> [data.json]")
        sys.exit(1)
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else "data.json"

    here = os.path.dirname(os.path.abspath(__file__))
    template = json.load(open(os.path.join(here, "data.json"), encoding="utf-8"))
    months = template["months"]
    midx = {m: i for i, m in enumerate(months)}

    rows = load_rows(src)
    by_aid = {}
    for r in rows:
        aid = str(r.get("AID", "")).strip()
        if not aid:
            continue
        by_aid[aid] = r

    PUBLIC = ("aid", "name", "resp", "indent")  # fields ที่เปิดเผยได้
    warnings = []
    for s in template["steps"]:
        for a in s["acts"]:
            # เก็บเฉพาะ field เปิดเผยได้ แล้วเติม start/end/status
            for k in list(a.keys()):
                if k not in PUBLIC:
                    a.pop(k, None)
            a.setdefault("start", None); a.setdefault("end", None); a.setdefault("status", "none")
            row = by_aid.get(a["aid"])
            if not row:
                continue
            sm = str(row.get("StartMonth", "") or "").strip()
            em = str(row.get("EndMonth", "") or "").strip()
            st = str(row.get("Status", "") or "").strip()
            a["status"] = STATUS_MAP.get(st, "none")
            if sm and sm not in midx:
                warnings.append(f"{a['aid']}: StartMonth '{sm}' ไม่ตรงกับ months")
            if em and em not in midx:
                warnings.append(f"{a['aid']}: EndMonth '{em}' ไม่ตรงกับ months")
            s_i = midx.get(sm); e_i = midx.get(em)
            if s_i is not None and e_i is not None:
                if e_i < s_i:
                    s_i, e_i = e_i, s_i
                a["start"], a["end"] = s_i, e_i
            elif s_i is not None:
                a["start"] = a["end"] = s_i

    template["updated"] = datetime.datetime.now().strftime("%d/%m/") + str(datetime.datetime.now().year + 543)
    json.dump(template, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

    done = sum(1 for s in template["steps"] for a in s["acts"] if a["status"] == "done")
    prog = sum(1 for s in template["steps"] for a in s["acts"] if a["status"] == "prog")
    print(f"เขียน {out} แล้ว | เสร็จ {done} · กำลังทำ {prog}")
    for w in warnings:
        print("  ⚠", w)

if __name__ == "__main__":
    main()
