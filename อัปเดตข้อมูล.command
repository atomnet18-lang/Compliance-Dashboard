#!/bin/bash
# ดับเบิลคลิกไฟล์นี้เพื่อแปลงข้อมูลจาก Excel เป็น data.json
# จากนั้นเปิด GitHub Desktop เพื่อ Commit + Push
cd "$(dirname "$0")"
echo "==============================================="
echo " อัปเดตข้อมูล Dashboard Compliance"
echo "==============================================="
python3 build_data.py
status=$?
echo ""
if [ $status -eq 0 ]; then
  echo "✓ สร้าง data.json เรียบร้อย"
  echo "  ขั้นต่อไป: เปิด GitHub Desktop -> ใส่ข้อความ -> Commit to main -> Push origin"
else
  echo "✗ เกิดข้อผิดพลาด — ตรวจว่าได้บันทึกไฟล์ Excel และติดตั้ง Python 3 แล้ว"
fi
echo ""
read -p "กด Enter เพื่อปิดหน้าต่างนี้..."
