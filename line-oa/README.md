# LINE OA Rich Menu

ไฟล์ชุดนี้ใช้สร้าง Rich Menu สำหรับ LINE Official Account ของระบบ iPad JV

## เมนู

- ตรวจสถานะ
- แจ้งปัญหา
- ยืม / คืน
- ขอแอพ
- คู่มือ
- เปิดแอป

## วิธีใช้

1. ไปที่ LINE Developers ของ OA นี้ แล้วสร้าง/คัดลอก Channel access token แบบ long-lived
2. ตั้ง token ใน PowerShell:

```powershell
[Environment]::SetEnvironmentVariable('LINE_CHANNEL_ACCESS_TOKEN', 'PASTE_TOKEN_HERE', 'User')
```

3. เปิด PowerShell ใหม่ แล้วรัน:

```powershell
.\line-oa\create-rich-menu.ps1
```

ถ้าต้องการสร้างรูปอย่างเดียว:

```powershell
.\line-oa\create-rich-menu.ps1 -SkipUpload
```

## ไฟล์

- `rich-menu.json` คือพื้นที่กดและ URL ของแต่ละเมนู
- `create-rich-menu.ps1` สร้าง `rich-menu.png`, สร้าง rich menu ผ่าน LINE API, อัปโหลดรูป, และตั้งเป็น default

Rich menu ที่สร้างผ่าน Messaging API จะต้องจัดการผ่าน API ตามกฎของ LINE ไม่ใช่แก้ instance เดิมผ่าน LINE Official Account Manager
