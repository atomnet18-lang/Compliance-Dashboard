#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_data.py — แปลงไฟล์บันทึกผลจริง Excel เป็น data.json สำหรับ Dashboard Compliance

แหล่งข้อมูล:  CP_2569_บันทึกผลจริง_template.xlsx   (ชีต "บันทึกผล" — แผนเดียว แบ่ง 5 ขั้นตอนด้วยแถบหัวกลุ่ม)
    - แถวหัวกลุ่ม "ขั้นที่ N · ..."     -> แบ่งขั้นตอน
    - คอลัมน์ น้ำหนัก                    -> ค่าถ่วงน้ำหนักของกิจกรรม
    - คอลัมน์ ม.ค.69..มี.ค.70            -> ผลจริง "ร้อยละสะสม" รายเดือน (ผู้ปฏิบัติงานกรอก)
    - คอลัมน์ สิ่งที่ดำเนินการ <เดือน>   -> หมายเหตุรายเดือน (แสดงบน dashboard)
    - คอลัมน์ ปัญหา/อุปสรรค              -> ใช้ภายใน ไม่ใส่ลง data.json (ไม่เผยแพร่)

เป้าหมายตามแผน (target) ฝังไว้ในสคริปต์ -> คำนวณเป็น "เป้าหมายสะสม %" รายเดือนแบบไต่เชิงเส้น
ใช้ไลบรารีมาตรฐานของ Python เท่านั้น (ไม่ต้อง pip install)   วิธีรัน:  python3 build_data.py
"""
import json, os, sys, re, zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(HERE, "CP_2569_บันทึกผลจริง_template.xlsx")
SHEET = "บันทึกผล"
OUT  = os.path.join(HERE, "data.json")
import datetime as _dt
_t=_dt.date.today()
UPDATED   = "%02d/%02d/%d" % (_t.day, _t.month, _t.year+543)  # วันที่แปลงไฟล์ (อัตโนมัติ)
CUR_MONTH = None   # None = เลือกเดือนล่าสุดที่มีข้อมูลผลจริงให้อัตโนมัติ

MONTHS = ['ม.ค.69','ก.พ.69','มี.ค.69','เม.ย.69','พ.ค.69','มิ.ย.69','ก.ค.69','ส.ค.69',
          'ก.ย.69','ต.ค.69','พ.ย.69','ธ.ค.69','ม.ค.70','ก.พ.70','มี.ค.70']
N = len(MONTHS)
COL_W = 4                 # คอลัมน์ น้ำหนัก
COL_M0 = 5                # คอลัมน์ผลจริงเดือนแรก (E)
COL_NOTE0 = 5 + N + 1     # คอลัมน์หมายเหตุเดือนแรก (ข้าม ปัญหา/อุปสรรค ที่คอลัมน์ 5+N)

PLAN = {  # ขั้น -> [(เดือนเริ่ม, เดือนสิ้นสุด) ตามลำดับกิจกรรม]
 "1": [(0,12),(0,11),(0,11)],
 "2": [(0,11),(0,5),(0,11)],
 "3": [(0,12),(0,12),(0,12),(0,12)],
 "4": [(4,9),(5,11),(5,11),(9,12)],
 "5": [(6,11),(6,12),(11,12)],
}
STEP_TITLE = {
 "1":"ทบทวนและปรับปรุงทะเบียนกฎหมาย กฎ ระเบียบที่สำคัญที่เกี่ยวข้องกับการดำเนินงาน",
 "2":"ระบุและประเมินปัจจัยความเสี่ยงที่อาจมีผลกระทบต่อองค์กรที่ทำให้ไม่สามารถปฏิบัติตามกฎ ระเบียบ และพิจารณากำหนด/ปรับปรุงการควบคุม",
 "3":"สื่อสารความรู้ สร้างความตระหนัก และพัฒนาบุคลากร",
 "4":"สอบทานและประเมินการปฏิบัติตามกฎ ระเบียบ",
 "5":"ทบทวนและนำเสนอผลการทบทวนนโยบาย แนวทางการปฏิบัติ และแผนงานประจำปี ต่อคณะกรรมการที่เกี่ยวข้อง",
}
warnings = []

# ---------- ตัวอ่าน .xlsx แบบไม่พึ่งไลบรารีภายนอก ----------
NS_MAIN='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
NS_REL ='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

def col_of(ref):
    m=re.match(r'([A-Z]+)(\d+)',ref); letters,row=m.group(1),int(m.group(2)); c=0
    for ch in letters: c=c*26+(ord(ch)-64)
    return c,row

def read_sheet(path, sheet_name):
    with zipfile.ZipFile(path) as z:
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall(NS_MAIN+'si'):
                shared.append(''.join(t.text or '' for t in si.iter(NS_MAIN+'t')))
        wb=ET.fromstring(z.read('xl/workbook.xml'))
        rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relmap={r.get('Id'):r.get('Target') for r in rels}
        target=None
        for s in wb.find(NS_MAIN+'sheets'):
            if s.get('name')==sheet_name:
                target=relmap.get(s.get(NS_REL+'id')); break
        if target is None:
            print("ไม่พบชีต '%s'"%sheet_name); sys.exit(1)
        target=target[1:] if target.startswith('/') else ('xl/'+target if not target.startswith('xl/') else target)
        sheet=ET.fromstring(z.read(target)); rows={}
        for c in sheet.iter(NS_MAIN+'c'):
            ref=c.get('r')
            if not ref: continue
            ci,ri=col_of(ref); t=c.get('t'); v=c.find(NS_MAIN+'v')
            if t=='s': val=shared[int(v.text)] if v is not None else ''
            elif t=='inlineStr':
                isn=c.find(NS_MAIN+'is'); val=''.join(x.text or '' for x in isn.iter(NS_MAIN+'t')) if isn is not None else ''
            else: val=v.text if v is not None else ''
            rows.setdefault(ri,{})[ci]=(val or '').strip()
        return [rows.get(k,{}) for k in range(1,(max(rows) if rows else 0)+1)]

def ramp(s,e):
    if e<s: e=s
    return [0 if i<s else (100 if i>=e else round((i-s+1)/(e-s+1)*100)) for i in range(N)]

def num(x):
    try: return max(0,min(100,round(float(x))))
    except: return 0

def main():
    if not os.path.exists(XLSX):
        print("ไม่พบไฟล์:",XLSX); sys.exit(1)
    rows=read_sheet(XLSX, SHEET)
    hr=next((i for i,r in enumerate(rows) if any(v=='ลำดับ' for v in r.values())), None)
    if hr is None:
        print("ไม่พบหัวตาราง (คำว่า 'ลำดับ')"); sys.exit(1)

    plans={}; order=[]; aid=0; cur=None; seq=0
    for row in rows[hr+1:]:
        a1=row.get(1,"")
        if not a1: continue
        mg=re.match(r'^\s*ขั้นที่\s*(\d+)', a1)
        if mg:                                   # แถวหัวกลุ่มขั้นตอน
            cur=mg.group(1); seq=0
            key="s"+cur
            if key not in plans:
                plans[key]={"name":"ขั้นที่ %s"%cur,"full":"ขั้นที่ %s · %s"%(cur,STEP_TITLE.get(cur,"")),
                            "kpi":"ดำเนินการครบทุกกิจกรรมตามแผน","budget":"","cum":[],"acts":[]}
                order.append(key)
            continue
        if cur is None: continue
        if not re.match(r'^\d+\.\d+', a1): continue   # ข้ามแถวที่ไม่ใช่กิจกรรม
        seq+=1
        name=row.get(2,""); unit=row.get(3,""); w=num(row.get(COL_W,0)) or 1
        a=[num(row.get(COL_M0+mi,0)) for mi in range(N)]
        notes=[row.get(COL_NOTE0+mi,"") for mi in range(N)]
        pw=PLAN.get(cur,[])
        if seq-1 < len(pw): s,e=pw[seq-1]
        else: s,e=0,N-1; warnings.append("ขั้น %s ลำดับ %d ไม่มีช่วงแผน ใช้เต็มช่วง"%(cur,seq))
        t=ramp(s,e)
        mx=max(a) if a else 0
        status='done' if mx>=100 else ('prog' if mx>0 else 'none')
        aid+=1
        plans["s"+cur]["acts"].append({"id":a1,"aid":"A%02d"%aid,"name":name,"unit":unit,
            "w":w,"start":s,"end":e,"status":status,"t":t,"a":a,"notes":notes})

    for key in list(plans):
        p=plans[key]
        if not p["acts"]:
            warnings.append("ขั้น %s ไม่มีกิจกรรม"%key); del plans[key]; order.remove(key); continue
        tw=sum(x["w"] for x in p["acts"])
        p["cum"]=[round(sum(x["t"][i]*x["w"] for x in p["acts"])/tw,1) for i in range(N)]
        p["budget"]="%d กิจกรรม"%len(p["acts"])

    cur=CUR_MONTH
    if cur is None:
        last=-1
        for p in plans.values():
            for a in p["acts"]:
                for i,v in enumerate(a["a"]):
                    if v and v>0 and i>last: last=i
        cur=last if last>=0 else 0
    out={"updated":UPDATED,"curMonth":cur,
         "source":"งานด้าน Compliance · ฝ่ายบริหารความเสี่ยงองค์กร (ฝบส.) การไฟฟ้านครหลวง",
         "title":"แผนการกำกับดูแลการปฏิบัติตามกฎ ระเบียบ (Compliance) ประจำปี 2569",
         "months":MONTHS,"plans":{k:plans[k] for k in order}}
    with open(OUT,"w",encoding="utf-8") as fh:
        json.dump(out,fh,ensure_ascii=False,indent=2)

    tot=sum(len(p["acts"]) for p in plans.values())
    done=sum(1 for p in plans.values() for a in p["acts"] if a["status"]=="done")
    prog=sum(1 for p in plans.values() for a in p["acts"] if a["status"]=="prog")
    print("✓ สร้าง data.json สำเร็จ")
    print("  ขั้นตอน %d · กิจกรรม %d · เสร็จ %d · กำลังดำเนินการ %d · ยังไม่เริ่ม %d"%(len(plans),tot,done,prog,tot-done-prog))
    if warnings:
        print("\n⚠ คำเตือน:")
        for w in dict.fromkeys(warnings): print("   -",w)

if __name__=="__main__":
    main()
