#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_data.py — แปลงไฟล์บันทึกผล Excel เป็น data.json สำหรับ Dashboard Compliance (รูปแบบ CG)

แหล่งข้อมูลเดียว:
    Compliance_Tracking_2569.xlsx   (ชีต "บันทึกผล")
    - คอลัมน์ เริ่มตามแผน / สิ้นสุดตามแผน  -> เป้าหมายสะสม % รายเดือน
    - คอลัมน์ สถานะ / เริ่มจริง / สิ้นสุดจริง -> ผลการดำเนินงานจริง % รายเดือน
    - คอลัมน์ หมายเหตุ -> รายละเอียดที่แสดงในหน้า dashboard

ใช้ไลบรารีมาตรฐานของ Python เท่านั้น (ไม่ต้อง pip install อะไรเลย)

วิธีรัน:
    python3 build_data.py
"""
import json, os, sys, re, zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(HERE, "Compliance_Tracking_2569.xlsx")
SHEET_NAME = "บันทึกผล"
OUT  = os.path.join(HERE, "data.json")
UPDATED   = "20/06/2569"   # วันที่ปรับปรุง (แก้ได้)
CUR_MONTH = 5              # เดือนที่ dashboard เปิดค่าเริ่มต้น (0=ม.ค.69 ... 5=มิ.ย.69)

MONTHS = ['ม.ค.69','ก.พ.69','มี.ค.69','เม.ย.69','พ.ค.69','มิ.ย.69','ก.ค.69','ส.ค.69',
          'ก.ย.69','ต.ค.69','พ.ย.69','ธ.ค.69','ม.ค.70','ก.พ.70','มี.ค.70']
N = len(MONTHS)
MIDX = {m: i for i, m in enumerate(MONTHS)}
STATUS_MAP = {'เสร็จสิ้น':'done','เสร็จ':'done',
              'กำลังดำเนินการ':'prog','กำลังดำเนิน':'prog',
              'ยังไม่เริ่ม':'none','ยังไม่เริ่มงาน':'none','':'none'}
warnings = []

# ---------- ตัวอ่าน .xlsx แบบไม่พึ่งไลบรารีภายนอก ----------
NS_MAIN = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
NS_REL  = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
NS_PKG  = '{http://schemas.openxmlformats.org/package/2006/relationships}'

def col_of(ref):
    m = re.match(r'([A-Z]+)(\d+)', ref)
    letters, row = m.group(1), int(m.group(2))
    c = 0
    for ch in letters:
        c = c * 26 + (ord(ch) - 64)
    return c, row

def read_sheet(path, sheet_name):
    """คืน list ของ row โดยแต่ละ row เป็น dict {col_index: value}"""
    with zipfile.ZipFile(path) as z:
        # shared strings
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            sst = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in sst.findall(NS_MAIN + 'si'):
                shared.append(''.join(t.text or '' for t in si.iter(NS_MAIN + 't')))
        # หา target ของชีตตามชื่อ
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rid = None
        for s in wb.find(NS_MAIN + 'sheets'):
            if s.get('name') == sheet_name:
                rid = s.get(NS_REL + 'id'); break
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        target = None
        for rel in rels:
            if rel.get('Id') == rid:
                target = rel.get('Target'); break
        if target is None:
            target = 'worksheets/sheet1.xml'
        if target.startswith('/'):
            target = target[1:]               # path สัมบูรณ์ในแพ็กเกจ เช่น /xl/worksheets/sheet1.xml
        elif not target.startswith('xl/'):
            target = 'xl/' + target           # path สัมพัทธ์กับโฟลเดอร์ xl/
        sheet = ET.fromstring(z.read(target))
        rows = {}
        for c in sheet.iter(NS_MAIN + 'c'):
            ref = c.get('r')
            if not ref:
                continue
            ci, ri = col_of(ref)
            t = c.get('t')
            v = c.find(NS_MAIN + 'v')
            if t == 's':
                val = shared[int(v.text)] if v is not None else ''
            elif t == 'inlineStr':
                isn = c.find(NS_MAIN + 'is')
                val = ''.join(x.text or '' for x in isn.iter(NS_MAIN + 't')) if isn is not None else ''
            else:
                val = v.text if v is not None else ''
            rows.setdefault(ri, {})[ci] = (val or '').strip()
        return [rows[k] for k in sorted(rows)]

# ---------- คำนวณ ----------
def ramp(s, e):
    if s is None or e is None:
        return [0]*N
    if e < s: e = s
    out = []
    for i in range(N):
        if i < s: out.append(0)
        elif i >= e: out.append(100)
        else: out.append(round((i - s + 1)/(e - s + 1)*100))
    return out

def actual_ramp(status, s, e):
    a = [0]*N
    if status == 'none' or s is None:
        return a
    if e is None or e < s: e = s
    for i in range(N):
        if i < s:
            a[i] = 0
        elif status == 'done':
            a[i] = 100 if i >= e else round((i - s + 1)/(e - s + 1)*100)
        else:  # prog: ไต่ถึงเดือนปัจจุบัน ไม่ถึง 100
            a[i] = min(99, round((i - s + 1)/(e - s + 1)*100)) if i <= CUR_MONTH else 0
    return a

def midx(label, what):
    if not label:
        return None
    if label not in MIDX:
        warnings.append("%s '%s' ไม่อยู่ในรายชื่อเดือน (ตรวจการสะกด)" % (what, label))
        return None
    return MIDX[label]

def main():
    if not os.path.exists(XLSX):
        print("ไม่พบไฟล์:", XLSX); sys.exit(1)
    rows = read_sheet(XLSX, SHEET_NAME)
    if not rows:
        print("ชีต '%s' ว่าง" % SHEET_NAME); sys.exit(1)

    header = rows[0]
    col = {}  # ชื่อหัว -> index
    for ci, name in header.items():
        col[name] = ci
    def get(row, name):
        return row.get(col.get(name, -1), "")

    plans = {}
    order = []
    aid = 0
    for row in rows[1:]:
        step = get(row, "ขั้นตอน"); id_ = get(row, "รหัส")
        if not id_:
            continue
        aid += 1
        name = get(row, "กิจกรรม"); unit = get(row, "ผู้รับผิดชอบ")
        ps = midx(get(row, "เริ่มตามแผน"), "เริ่มตามแผน")
        pe = midx(get(row, "สิ้นสุดตามแผน"), "สิ้นสุดตามแผน")
        status = STATUS_MAP.get(get(row, "สถานะ"), 'none')
        as_ = midx(get(row, "เริ่มจริง"), "เริ่มจริง")
        ae = midx(get(row, "สิ้นสุดจริง"), "สิ้นสุดจริง")
        note = get(row, "หมายเหตุ / ผลการดำเนินงาน")
        # ถ้ามีสถานะแต่ไม่กรอกช่วงจริง ใช้ช่วงตามแผนแทน
        if status != 'none' and as_ is None:
            as_, ae = ps, pe
        t = ramp(ps, pe)
        a = actual_ramp(status, as_, ae)
        notes = [""]*N
        if note:
            pos = ae if ae is not None else (as_ if as_ is not None else CUR_MONTH)
            if pos is None: pos = CUR_MONTH
            notes[min(pos, N-1)] = note

        key = "s" + re.sub(r'\D', '', step or id_.split('.')[0])
        if key not in plans:
            plans[key] = {"name": step or ("ขั้นที่ " + id_.split('.')[0]),
                          "full": step or "", "kpi": "ดำเนินการครบทุกกิจกรรมตามแผน",
                          "budget": "", "cum": [], "acts": []}
            order.append(key)
        plans[key]["acts"].append({
            "id": id_, "aid": "A%02d" % aid, "name": name, "unit": unit, "w": 1,
            "start": (ps if ps is not None else 0), "end": (pe if pe is not None else N-1),
            "status": status, "t": t, "a": a, "notes": notes,
        })

    # เติม full/cum/budget
    for key in plans:
        p = plans[key]
        sn = re.sub(r'\D', '', key)
        if not p["full"]:
            p["full"] = p["name"]
        elif not p["full"].startswith("ขั้นที่"):
            p["full"] = p["name"]
        p["budget"] = "%d กิจกรรม" % len(p["acts"])
        p["cum"] = [round(sum(x["t"][i] for x in p["acts"])/len(p["acts"]), 1) for i in range(N)]

    out = {
        "updated": UPDATED, "curMonth": CUR_MONTH,
        "source": "งานด้าน Compliance · ฝ่ายบริหารความเสี่ยงองค์กร (ฝบส.) การไฟฟ้านครหลวง",
        "title": "แผนการกำกับดูแลการปฏิบัติตามกฎ ระเบียบ (Compliance) ประจำปี 2569",
        "months": MONTHS,
        "plans": {k: plans[k] for k in order},
    }
    with open(OUT, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=2)

    tot = sum(len(p["acts"]) for p in plans.values())
    done = sum(1 for p in plans.values() for a in p["acts"] if a["status"] == "done")
    prog = sum(1 for p in plans.values() for a in p["acts"] if a["status"] == "prog")
    print("✓ สร้าง data.json สำเร็จ")
    print("  ขั้นตอน %d · กิจกรรม %d · เสร็จ %d · กำลังดำเนินการ %d · ยังไม่เริ่ม %d"
          % (len(plans), tot, done, prog, tot - done - prog))
    if warnings:
        print("\n⚠ คำเตือน:")
        for w in dict.fromkeys(warnings):
            print("   -", w)

if __name__ == "__main__":
    main()
