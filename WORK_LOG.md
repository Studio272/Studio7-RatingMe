# 📋 WORK LOG - Studio 7 Dashboard

บันทึกการทำงานทุกครั้ง - **อ่านก่อนทำงานต่อทุกครั้ง**

---

## 🔖 สรุปโครงสร้างไฟล์หลัก

| ไฟล์ | คำอธิบาย |
|------|----------|
| `dashboard.html` | ไฟล์หลัก Dashboard (HTML + CSS + JS ทั้งหมดอยู่ในไฟล์เดียว ~6000 บรรทัด) |
| `check.js` | Script ตรวจสอบ |
| `index.html` | หน้า redirect |
| `WORK_LOG.md` | ไฟล์บันทึกการทำงาน (ไฟล์นี้) |

## 🏗️ โครงสร้างหลักใน dashboard.html

| ส่วน | ตำแหน่ง (โดยประมาณ) |
|------|---------------------|
| CSS Styles | บรรทัด 15-1250 |
| HTML - Upload View | บรรทัด 1251-1350 |
| HTML - Dashboard View | บรรทัด 1351-1900 |
| HTML - Tab: ข้อมูลยอดขาย | `#tab-sales` |
| HTML - Tab: Rating Demo | `#tab-rating` (บรรทัด ~1716) |
| HTML - Tab: Staff | `#tab-staff` |
| JS - Data Parsing | บรรทัด ~2400-2650 |
| JS - renderDashboard() | บรรทัด ~3400-4050 |
| JS - Rating Branch Summary | บรรทัด ~3751-3844 |
| JS - renderRDStaffTable() | บรรทัด ~5888-5962 |
| JS - populateRDStaffData() | บรรทัด ~5984-5998 |

## 📊 ข้อมูลสำคัญ

### Data Fields (Rating records)
- `feedback1` = ตอบกลับ Discovery 
- `feedback2` = ตอบกลับ Demo Topic
- `rating` = คะแนนความพึงพอใจ (1-5)
- `topic` = หัวข้อ Demo
- `staffName`, `staffId`, `branch` = ข้อมูลพนักงาน/สาขา

### การตรวจสอบ feedback สำเร็จ/ไม่สำเร็จ
- **สำเร็จ**: ค่าไม่ว่าง, ไม่ใช่ '-', ไม่ใช่ 'N/A'
- **ไม่สำเร็จ**: ค่าว่าง หรือ '-' หรือ 'N/A'

---

## 📝 บันทึกการทำงาน

### 2026-05-19 | ปรับปรุงหน้า Rating Demo
**ผู้ทำ**: Antigravity AI  
**สิ่งที่เปลี่ยน**:

1. **ส่วน "ผล RATING DEMO" (KPI Cards)**
   - เพิ่ม KPI Card: %สำเร็จ ตอบกลับ Discovery (`id="kpi-disc-pct"`)
   - เพิ่ม KPI Card: ตอบกลับ Demo Topic (`id="kpi-demotopic-success"`)
   - ปรับ grid layout รองรับ 7 cards (flex-wrap)
   - ตำแหน่ง HTML: บรรทัด ~1726-1769
   - ตำแหน่ง JS: บรรทัด ~3696-3706

2. **ส่วน "สรุปรวมสาขา (Rating Demo)"**
   - ปรับตารางเป็น 7 คอลัมน์: สาขา, พนักงาน(%เทียบDemo), Unit(%เทียบDemo), Demo, ตอบกลับ Demo Topic, %ไม่สำเร็จ, ความพึงพอใจ
   - เพิ่ม Dropdown sorting มากไปน้อย / น้อยไปมาก ทุกหัวข้อ
   - เพิ่มเงื่อนไขสีแดง: ถ้า %พนักงานเทียบDemo ต่ำกว่าภาพรวม
   - ตำแหน่ง HTML: บรรทัด ~1781-1788
   - ตำแหน่ง JS: บรรทัด ~3751-3844

3. **ส่วน "👤 สรุปข้อมูลรายพนักงาน"**
   - ปรับตารางให้แสดงข้อมูลเหมือนสรุปรวมสาขา
   - เพิ่ม Dropdown sorting แต่ละหัวข้อ
   - ตำแหน่ง HTML: บรรทัด ~1791-1810
   - ตำแหน่ง JS: บรรทัด ~5888-5998

4. **สร้างไฟล์ WORK_LOG.md**
   - บันทึกการทำงานทุกครั้ง
   - อ่านก่อนทำงานต่อทุกครั้ง

### 2026-05-19 13:12 | เปลี่ยนชื่อคอลัมน์ + แก้สูตร AVG/PIA
**สิ่งที่เปลี่ยน**:
- เปลี่ยน %พนง.เทียบDemo → **AVG/PIA** = Demo/พนักงาน (จาก staff/demo% เป็น demo/staff)
- เปลี่ยนชื่อคอลัมน์: Unit ขาย→**Sale Unit**, %Unit/Demo→**Att%Unit**, จำนวน Demo→**Demo**, ตอบกลับ Demo Topic→**สำเร็จ**, %ไม่สำเร็จ Demo Topic→**ไม่สำเร็จ**
- แก้ทั้ง 2 ตาราง: สรุปรวมสาขา + สรุปข้อมูลรายพนักงาน
- ตำแหน่ง: renderRatingBranchTable() ~line 6081, renderRDStaffTable() ~line 5983
