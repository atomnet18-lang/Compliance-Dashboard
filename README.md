# Dashboard Compliance 2569 — workflow ระดับ A (กึ่งอัตโนมัติ)

ผู้ปฏิบัติงานกรอกผ่าน **SharePoint List** (อยู่ในระบบ M365 ของ กฟน.) → แอดมินแปลงเป็น `data.json` ด้วยสคริปต์ → push ขึ้น GitHub → หน้า Dashboard (GitHub Pages) อ่านข้อมูลล่าสุดมาแสดงแบบ Gantt

```
ผู้ปฏิบัติงาน (ทุกหน่วยงาน / ฝบส. / ฝกม. ฯลฯ)
   ↓ เลือกเดือนเริ่ม–สิ้นสุด + สถานะ ของแต่ละกิจกรรม
SharePoint List  (M365 — ฐานข้อมูลกลาง คุมสิทธิ์รายฝ่าย)
   ↓ export .xlsx/.csv  →  convert_sharepoint_to_json.py  (แอดมินรัน / อนุมัติ)
data.json  บน GitHub repo
   ↓ fetch()
GitHub Pages  (หน้า Dashboard public — อ่านอย่างเดียว)
```

---

## ไฟล์ในชุดนี้

| ไฟล์ | หน้าที่ |
|---|---|
| `index.html` | หน้า Dashboard Gantt (อ่าน `data.json` แบบ fetch — อ่านอย่างเดียว) |
| `data.json` | ข้อมูลที่แสดง (โครงสร้าง 5 ขั้น 34 กิจกรรม + ช่วงเวลา/สถานะ) |
| `convert_sharepoint_to_json.py` | แปลง List export → `data.json` (ตัด field ภายในออก) |
| `Compliance_SharePoint_List_Seed.xlsx` | ข้อมูลตั้งต้น 34 แถวสำหรับนำเข้า SharePoint List พร้อม dropdown |
| `list_export_template.csv` | ฟอร์ม CSV เปล่าไว้ทดสอบสคริปต์ |

---

## ขั้นที่ 1 — สร้าง SharePoint List

นำเข้า `Compliance_SharePoint_List_Seed.xlsx` เป็น List ใหม่ (SharePoint → New → List → From Excel) คอลัมน์:

| คอลัมน์ | ชนิด | หมายเหตุ |
|---|---|---|
| **AID** | Single line | รหัสกิจกรรม เช่น `1-01` — **KEY ห้ามแก้/ห้ามลบแถว** |
| Step | Single line | ขั้นที่ (อ้างอิงเฉยๆ) |
| Activity | Multiple lines | ชื่อกิจกรรม (อ้างอิงเฉยๆ) |
| Resp | Single line | ผู้รับผิดชอบ |
| **StartMonth** | Choice | ผู้ปฏิบัติงานเลือก — ค่าต้องตรงกับ `months` (`ม.ค.69` … `มี.ค.70`) |
| **EndMonth** | Choice | เดือนสิ้นสุด |
| **Status** | Choice | `ยังไม่เริ่ม` / `กำลังดำเนินการ` / `เสร็จสิ้น` |
| Note | Multiple lines | หมายเหตุภายใน — **ไม่ถูกเผยแพร่** (สคริปต์ตัดทิ้ง) |

ตั้งสิทธิ์ให้แต่ละฝ่ายแก้ได้เฉพาะแถวของตน (item-level permission หรือแยก View ตาม Resp)

> คอลัมน์ Choice ของ StartMonth/EndMonth ให้ใส่ตัวเลือก 15 เดือนตามลำดับใน `data.json` → `months`

## ขั้นที่ 2 — รอบรายงาน: แปลงเป็น data.json

แอดมิน export List เป็น Excel/CSV แล้วรัน:

```bash
python3 convert_sharepoint_to_json.py <ไฟล์ที่ export.xlsx> data.json
```

สคริปต์จะ: จับคู่แถวด้วย AID, แปลงชื่อเดือนเป็น index, แปลงสถานะไทย→โค้ด, **ตัดคอลัมน์ Note ออก** และเตือนหากมีชื่อเดือนพิมพ์ผิด (ไม่ตรง `months`)

## ขั้นที่ 3 — push ขึ้น GitHub + เปิด Pages

```bash
git init
git add index.html data.json
git commit -m "เริ่มต้น Compliance Dashboard 2569"
git branch -M main
git remote add origin https://github.com/<username>/compliance-dashboard-2569.git
git push -u origin main
```

Settings → Pages → Source = Deploy from a branch → `main` / root → Save
เผยแพร่ที่ `https://<username>.github.io/compliance-dashboard-2569/`

แต่ละรอบถัดไป แปลง `data.json` ใหม่ แล้ว `git add data.json && git commit -m "อัปเดต <รอบ>" && git push` — หน้าเว็บอัปเดตเอง

---

## ทดสอบบนเครื่องก่อน push

เปิด `index.html` ตรงๆ จะติด CORS (อ่าน data.json ไม่ได้) ต้องรันผ่านเซิร์ฟเวอร์:

```bash
cd <โฟลเดอร์ที่มี index.html กับ data.json>
python -m http.server 8000
# เปิด http://localhost:8000
```

---

## ธรรมาภิบาลข้อมูล (สำคัญ)

GitHub Pages เปิด public — สคริปต์จึงดึงเฉพาะ field ที่เปิดเผยได้ (AID, ชื่อกิจกรรม, ผู้รับผิดชอบ, ช่วงเวลา, สถานะ) และ **ตัด Note (หมายเหตุภายใน) ออกทุกครั้ง** แอดมินควรเป็นผู้ "อนุมัติ push" ทุกรอบในเฟสแรก

## ต่อยอด (เฟสถัดไป — ระดับ B)

ใช้ **Power Automate** ดึงข้อมูลจาก SharePoint List ตามตารางเวลา → จัดรูปเป็น `data.json` → เขียนเข้า GitHub ผ่าน GitHub API (HTTP action, PUT `/contents/data.json`) โดยอัตโนมัติ ไม่ต้องมีคนกลางรันสคริปต์/push ด้วยมือ

---
จัดทำโดยงานด้าน Compliance · ฝ่ายบริหารความเสี่ยงองค์กร (ฝบส.) การไฟฟ้านครหลวง
