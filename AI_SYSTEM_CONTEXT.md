# AI System Context - IPAD JV

ไฟล์นี้เป็นบันทึกบริบทสำหรับ AI/ผู้ดูแลที่จะแก้ระบบต่อในอนาคต ให้อ่านไฟล์นี้ก่อนแก้ไขโค้ดหรือข้อมูลทุกครั้ง

## ภาพรวมระบบ

โปรเจกต์นี้เป็น Google Apps Script สำหรับระบบบริหารจัดการไอแพด โรงเรียนจักราชวิทยา

ไฟล์หลัก:
- `Code.gs` - backend Google Apps Script, จัดการชีต, auth, CRUD, import/export, upload รูป, issue workflow
- `Index.html` - frontend ทั้งหน้าแรก, dashboard, admin pages, modal, login, status app
- `appsscript.json` - manifest/scopes/web app config
- `.clasp.json` - scriptId และ rootDir สำหรับ push ด้วย clasp

คำสั่ง push:
```powershell
$nodeRoot = Join-Path $env:LOCALAPPDATA 'Programs\node-portable'
& (Join-Path $nodeRoot 'clasp.cmd') push --force
```

บัญชี clasp ที่เคยใช้:
- `wisanuforever15@gmail.com`

ข้อจำกัดสิทธิ์:
- บัญชีนี้ push/create version ได้
- แต่ deploy/redeploy web app ไม่ได้ เพราะไม่ใช่บัญชีใน domain เดียวกับ owner
- ถ้าต้อง deploy ให้ใช้บัญชีเจ้าของสคริปต์หรือบัญชีในโดเมนโรงเรียน

## ชีตหลัก

`iPad_Data`
- 1: ID
- 2: Serial
- 3: AssetCode / รหัสครุภัณฑ์
- 4: Prefix
- 5: FirstName
- 6: LastName
- 7: PersonCode
- 8: Position
- 9: Grade
- 10: Room
- 11: Status
- 12: BorrowDate
- 13: ReturnDate
- 14: Notes

`Database`
- ใช้เป็น master serial database
- หลัง repair ล่าสุด มี 1,299 rows
- โครงล่าสุด:
  - Serial
  - ID
  - Asset Number
  - Borrower Name

`Students`
- มีคอลัมน์รูปนักเรียนที่ column 8
- โครงที่ควรรักษา:
  - ลำดับ, รหัสนักเรียน, คำนำหน้า, ชื่อ, นามสกุล, ระดับชั้น, ห้อง, รูปนักเรียน

## สถานะข้อมูลล่าสุด

มีการนำข้อมูลจากไฟล์:
- `C:\Users\Kru_E\Downloads\Grid_Report_export_WS_855dce45-ae50-4f65-b1b5-8307e045ae9c.csv`
- `C:\Users\Kru_E\Downloads\พิมพ์เลขคุรุภัณฑ์ Ipad - Database.csv`

ผลสรุปล่าสุด:
- ไฟล์ Grid มี 1,299 เครื่อง
- SC = 1,174
- TC = 125
- ไม่มี serial ซ้ำในไฟล์ Grid
- ไม่มี ID ซ้ำในไฟล์ Grid
- ไฟล์ `พิมพ์เลขคุรุภัณฑ์ Ipad - Database.csv` ใช้เฉพาะ `serial_no` และ `asset_no`

หลัง repair ล่าสุดจากชีตจริง:
- Database rows = 1,299
- Database matched = 1,299
- Database missing = 0
- Database ID mismatch = 0
- Database asset mismatch = 0
- Database extra = 0
- iPad_Data rows = 1,179
- iPad_Data ID mismatch = 0
- iPad_Data asset mismatch = 0
- มี serial จากไฟล์ Database ประมาณ 120 รายการที่ยังไม่อยู่ใน `iPad_Data`

หลัก mapping:
- `ID` ใน `iPad_Data` และ `Database` ใช้รหัส `SC...` หรือ `TC...`
- `AssetCode` / `Asset Number` ใช้เลขครุภัณฑ์จริงจาก `asset_no` เช่น `จว. 13.1.2/1/68`
- ห้ามนำ `จักราชวิทยา_SC875` ไปใส่เป็นรหัสครุภัณฑ์

## การแก้ที่ทำไปแล้ว

### Upload รูปนักเรียน

ปรับให้ระบบข้ามนักเรียนที่มีรูปแล้ว:
- frontend preview แสดงสถานะพร้อมอัปโหลด / ข้ามเพราะมีรูปแล้ว / ไม่พบรหัส
- backend `uploadStudentPhoto()` ตรวจ column 8 ก่อน ถ้ามีรูปแล้ว return `skipped: true`
- ไม่ลบ/เขียนทับไฟล์รูปเดิม

### Admin login

เพิ่ม flow:
- ถ้า email ปัจจุบันอยู่ใน `Users` และ role ไม่ใช่ `user` เข้า admin ได้เลย
- ถ้าไม่มีสิทธิ์ จะเปิดหน้า login ผู้ดูแล
- default admin password:
  - user: `admin`
  - pass: `112233`
- เมื่อ login สำเร็จ จะ upsert email ปัจจุบันเข้า `Users` เป็น admin/super admin

### หน้าแรก UX/UI

หน้าแรกปรับเป็นเมนูหลัก 4 ช่องสี่เหลี่ยม:
- ตรวจสอบสถานะ
- แจ้งปัญหา
- ครู / นักเรียน
- ลงทะเบียน

ลิงก์รอง:
- คู่มือการใช้งาน
- สำหรับผู้ดูแลระบบ

ลด animation/transition:
- ลด `--dur` เป็น `.12s`
- modal animation สั้นลง
- เอา loading overlay ออกจากการเปลี่ยนหน้า admin เพื่อให้รู้สึกเร็วขึ้น

อัปเดตล่าสุด:
- หน้าแรกใช้ header แบบใหม่ มีกรอบโลโก้ badge และชื่อโรงเรียนวางซ้ายเพื่อให้อ่านง่ายขึ้น
- พื้นหลังหน้าแรกเป็นสีอ่อนพร้อม grid texture บาง ๆ
- เมนู 4 ช่องมี accent bar สีแยกตามประเภทงาน
- เมนูครู/นักเรียนใช้โทนเขียว, ตรวจสอบสถานะใช้โทนน้ำเงิน, แจ้งปัญหาใช้โทนเหลือง, ลงทะเบียนใช้โทนเทา
- คู่มือและผู้ดูแลระบบยังเป็นลิงก์รองขนาดเล็กด้านล่าง

### ID แบบ SC/TC

แก้ `addIPad()` ให้ไม่พังเมื่อ ID ไม่ใช่ตัวเลข:
- เพิ่ม `nextIPadId_(sheet)`
- ถ้าเพิ่ม iPad ใหม่โดยไม่มี `d.id` ระบบจะสร้าง fallback ID แบบ numeric ถ้าไม่มี numeric ID เหลืออยู่

## สิ่งที่ต้องระวัง

1. อย่าเปิด endpoint `doPost` ชั่วคราวค้างไว้
   - ก่อนหน้านี้เคยเพิ่ม endpoint เพื่อ update/audit Database
   - หลังใช้งานต้องถอดออกและ push แล้ว
   - หลัง push ต้องให้ owner deploy New version เพื่อปิด endpoint ใน web app จริง

2. `clasp run` ใช้ไม่ได้
   - error: `Script function not found. Please make sure script is deployed as API executable.`
   - ห้ามพึ่ง `clasp run` สำหรับงานข้อมูล

3. Sheets API ตรงถูกปิดใน Google-provided OAuth project
   - การเรียก Sheets REST API เคยได้ 403 service disabled
   - บัญชีปัจจุบันไม่มีสิทธิ์เปิด API

4. การ deploy
   - บัญชี `wisanuforever15@gmail.com` redeploy ไม่ได้
   - ต้องให้เจ้าของ/บัญชีใน domain กด deploy เอง

5. หากต้องอัปเดตข้อมูลชีตอีก
   - ให้สร้าง endpoint ชั่วคราวแบบมี secret เฉพาะงาน
   - push
   - ให้ owner deploy New version
   - ยิง payload
   - audit
   - ถอด endpoint
   - push
   - ให้ owner deploy New version อีกครั้งเพื่อปิด endpoint

## ไฟล์ช่วยตรวจ/แปลงที่สร้างไว้

- `Database_SC_TC_from_grid.csv`
- `Serial_ID_SC_TC_map.csv`
- `asset_no_missing_sc_tc_id.csv`
- `ipad_db_audit_result.json`
- `ipad_db_audit_after_repair.json`

ไฟล์เหล่านี้เป็นผลการวิเคราะห์/ตรวจสอบ ไม่ใช่ source code หลัก

## แนวทางสำหรับ AI รอบถัดไป

เมื่อผู้ใช้ให้แก้ระบบ:
1. อ่าน `AI_SYSTEM_CONTEXT.md` ก่อน
2. อ่านเฉพาะส่วนเกี่ยวข้องใน `Code.gs` และ `Index.html`
3. ตรวจ syntax ด้วย Node:
```powershell
$nodeRoot = Join-Path $env:LOCALAPPDATA 'Programs\node-portable'
@'
const fs = require('fs');
new Function(fs.readFileSync('Code.gs','utf8'));
const html = fs.readFileSync('Index.html','utf8');
for (const m of html.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/gi)) new Function(m[1]);
console.log('syntax OK');
'@ | & (Join-Path $nodeRoot 'node.exe') -
```
4. Push ด้วย `clasp.cmd push --force`
5. ถ้ามี web app behavior เปลี่ยน ต้องให้ owner deploy New version เอง
