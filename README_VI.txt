BO CAI DONG BO DASHBOARD LEN CLOUD (BAN BACKEND)
================================================

Cac file trong bo nay:
- app.py: backend Cloud, giu API /api/data va bo sung API nhan Excel /api/upload.
- sync_to_cloud.py: chay tren may chu xuong, tu dong gui Excel khi file thay doi.
- sync_config.example.json: mau cau hinh cho may chu xuong.
- START_SYNC.bat: chay chuong trinh dong bo tren Windows.
- requirements.txt: thu vien Python can cai tren Cloud.

QUAN TRONG: Bo nay CHUA co index.html/CSS/JavaScript cua Dashboard goc. Can dua cac file giao dien goc len cung thu muc voi app.py tren Cloud. Khong dua file Excel len GitHub.

A. CAU HINH CLOUD (HUONG DAN THEO RENDER)
1. Tao GitHub repository rieng, upload app.py, requirements.txt va cac file giao dien Dashboard goc (index.html, CSS, JavaScript, hinh anh...).
2. Khong upload Daily Production.xlsx, sync_config.json, token, hoac file nhay cam len repository.
3. Tren Render chon New > Web Service va ket noi repository.
4. Build Command: pip install -r requirements.txt
5. Start Command: python app.py
6. Trong Environment Variables, them:
   DATA_DIR = /var/data
   SYNC_TOKEN = mot chuoi bi mat dai, ngau nhien (vi du tao 32-64 ky tu).
   Render tu cap bien PORT; khong can tu dat PORT.
7. Them Persistent Disk cho Web Service, Mount Path: /var/data.
   Dung goi tra phi co Persistent Disk; khong dung Free cho van hanh thuc te vi Free sleep va khong giu file upload qua restart.
8. Deploy. Ghi lai dia chi HTTPS onrender.com cua Dashboard.

B. CAU HINH MAY CHU XUONG
1. Copy sync_to_cloud.py, START_SYNC.bat va sync_config.example.json vao thu muc may chu xuong.
2. Copy sync_config.example.json thanh sync_config.json.
3. Mo sync_config.json bang Notepad va sua:
   dashboard_url: dia chi HTTPS tren Render
   sync_token: phai giong chinh xac gia tri SYNC_TOKEN tren Render
   excel_path: duong dan day du den Daily Production.xlsx tren may chu xuong
4. Cai Python 3 neu may chua co. Khong can cai them thu vien cho sync_to_cloud.py (chi dung thu vien co san cua Python).
5. Chay START_SYNC.bat. Chuong trinh se gui Excel lan dau, sau do kiem tra moi 10 giay va gui lai khi noi dung file thay doi.
6. De dong bo tu dong sau khi khoi dong Windows, dua shortcut START_SYNC.bat vao Startup cua tai khoan may chu xuong.

C. CAC DIEU KIEN GIAO DIEN
- Giao dien Dashboard goc can tiep tuc goi GET /api/data.
- Theo huong dan hien co, giao dien dang duoc thiet ke de tai lai du lieu moi moi 30 giay. Tuy nhien chua co index.html/JS trong bo file nhan duoc nen chua the xac minh hoac dieu chinh phan nay.
- Frontend va backend phai dung cung domain; khong can CORS.
- Link Cloud la cong khai theo lua chon cua chu he thong: ai co link deu xem duoc du lieu. API upload van duoc bao ve bang SYNC_TOKEN.

D. COT DU LIEU EXCEL
Backend giu cach doc hien tai: sheet dau tien; du lieu tu dong 3; A=No., B=Process, C=Person in charge, E=Performance, F=Remain work.
Giu ten file tren may chu xuong la Daily Production.xlsx va khong doi thu tu/cau truc cot neu khong cap nhat code.
