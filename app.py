from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse
from datetime import datetime
import json, re, os
from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent
XLSX = BASE / 'Daily Production.xlsx'
PORT = 8000

def read_jobs():
    if not XLSX.exists():
        return {'error': 'Không tìm thấy Daily Production.xlsx trong thư mục Dashboard.'}
    wb = load_workbook(XLSX, data_only=True, read_only=False)
    ws = wb[wb.sheetnames[0]]
    jobs=[]; current=None
    for row in range(3, ws.max_row+1):
        no = ws.cell(row,1).value
        process = ws.cell(row,2).value
        person = ws.cell(row,3).value
        performance = ws.cell(row,5).value
        remain = ws.cell(row,6).value
        if no is not None and (process is not None or remain is not None):
            if current: jobs.append(current)
            current={'no': no, 'process': str(process or '').strip(), 'person': str(person or '').strip(), 'performance_raw': str(performance or '').strip(), 'remain': []}
        if current and remain is not None and str(remain).strip():
            current['remain'].append(str(remain).strip())
    if current: jobs.append(current)
    for j in jobs:
        raw=j.pop('performance_raw')
        m=re.search(r'(\d+(?:[.,]\d+)?)\s*%', raw)
        if m: pct=float(m.group(1).replace(',','.'))
        elif 'hoàn thành' in raw.lower(): pct=100.0
        else: pct=0.0
        pct=max(0,min(100,pct)); j['progress']=pct
        j['status']='Hoàn thành' if pct>=100 else ('Chưa bắt đầu' if pct<=0 else 'Đang thực hiện')
        j['product']=j['remain'][0] if j['remain'] else '—'
        j['remaining']=j['remain'][1:] if len(j['remain'])>1 else []
        j['remain_text']=' · '.join(j['remaining'][:3]) if j['remaining'] else ('Không có mục còn lại' if pct>=100 else 'Chưa ghi công việc còn lại')
        j['person']=j['person'] or 'Chưa phân công'
    st=os.stat(XLSX)
    stamp=datetime.fromtimestamp(st.st_mtime).strftime('%d/%m/%Y %H:%M:%S')
    completed=sum(1 for j in jobs if j['progress']>=100)
    started=sum(1 for j in jobs if 0<j['progress']<100)
    not_started=sum(1 for j in jobs if j['progress']==0)
    avg=round(sum(j['progress'] for j in jobs)/len(jobs)) if jobs else 0
    return {'title': ws.cell(1,1).value or 'THEO DÕI TIẾN ĐỘ SẢN XUẤT', 'jobs':jobs,
            'total':len(jobs),'completed':completed,'started':started,'not_started':not_started,
            'average':avg,'updated':stamp,'source':XLSX.name}

class Handler(SimpleHTTPRequestHandler):
    def __init__(self,*args,**kwargs): super().__init__(*args,directory=str(BASE),**kwargs)
    def do_GET(self):
        if urlparse(self.path).path == '/api/data':
            payload=json.dumps(read_jobs(),ensure_ascii=False).encode('utf-8')
            self.send_response(200); self.send_header('Content-Type','application/json; charset=utf-8'); self.send_header('Cache-Control','no-store'); self.send_header('Content-Length',str(len(payload))); self.end_headers(); self.wfile.write(payload)
        else:
            if urlparse(self.path).path == '/': self.path='/index.html'
            super().do_GET()
    def log_message(self, fmt, *args):
        print('[%s] %s' % (self.log_date_time_string(), fmt % args))

if __name__ == '__main__':
    print('\nDASHBOARD TIEN DO SAN XUAT')
    print('Mo tren may chu: http://localhost:%d' % PORT)
    print('Nhan Ctrl+C de dung. May tinh nay can duoc giu bat khi TV dang xem.\n')
    ThreadingHTTPServer(('0.0.0.0',PORT),Handler).serve_forever()
