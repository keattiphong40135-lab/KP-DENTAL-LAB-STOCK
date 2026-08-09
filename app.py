from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import io
from datetime import datetime

app = Flask(__name__)

inventory_db = {}
logs = []
stats = {"import_count": 0, "export_count": 0, "total": 0}

DEPARTMENTS = {
    "แผนกติดแน่น": ["Crown", "Veneer", "Implant", "Bridge"],
    "แผนกถอดได้": ["Retainer ใส", "Retainer ลวด", "CD", "RPD"]
}

def get_department(job_type):
    for dept, jobs in DEPARTMENTS.items():
        if job_type in jobs:
            return dept
    return "แผนกติดแน่น"

@app.route('/')
def index():
    return render_template('index.html', logs=logs, stats=stats, current_dept='all', departments=DEPARTMENTS)

@app.route('/department/<path:dept_name>')
def department_page(dept_name):
    filtered_logs = [l for l in logs if l.get('department') == dept_name]
    return render_template('index.html', logs=filtered_logs, stats=stats, current_dept=dept_name, departments=DEPARTMENTS)

@app.route('/api/scan', methods=['POST'])
def scan_barcode():
    data = request.json
    barcode = data.get('barcode', '').strip()
    action = data.get('action') 
    forced_dept = data.get('department')
    
    if not barcode:
        return jsonify({"success": False, "message": "กรุณาสแกนบาร์โค้ด!"}), 400

    today_str = datetime.now().strftime('%Y-%m-%d')

    if barcode not in inventory_db:
        default_job = DEPARTMENTS[forced_dept][0] if forced_dept in DEPARTMENTS else "Crown"
        inventory_db[barcode] = {
            "barcode": barcode,
            "department": forced_dept if forced_dept in DEPARTMENTS else "แผนกติดแน่น",
            "job_type": default_job, 
            "arch": "U/L",
            "receive_date": today_str, 
            "due_date": "", 
            "material": "-", 
            "teeth_count": "1",
            "retainer_color": "-",
            "importer": "",
            "exporter": "",
            "export_date": today_str,
            "status": "รอนำเข้า"
        }
    
    item = inventory_db[barcode]
    return jsonify({
        "success": True, 
        "message": f"สแกนบาร์โค้ดสำเร็จ: {barcode}",
        "item": item,
        "action": action
    })

@app.route('/api/update-job', methods=['POST'])
def update_job():
    data = request.json
    barcode = data.get('barcode')
    action = data.get('action')
    
    if barcode not in inventory_db:
        return jsonify({"success": False, "message": "ไม่พบข้อมูลสินค้า"}), 404
        
    item = inventory_db[barcode]
    job_type = data.get('job_type')
    dept = get_department(job_type)
    
    receive_date = data.get('receive_date')
    due_date = data.get('due_date')
    material = data.get('material')
    teeth_count = data.get('teeth_count', '1')
    retainer_color = data.get('retainer_color', '-')
    importer = data.get('importer')
    exporter = data.get('exporter')
    export_date = data.get('export_date')
    
    if action == 'import':
        if not receive_date or not importer or not material or material == '-':
            return jsonify({"success": False, "message": "⚠️ กรุณากรอก วันรับงาน, วัสดุ และชื่อผู้นำเข้า ให้ครบถ้วนก่อนบันทึกนำเข้า!"}), 400
        item['status'] = 'กำลังทำ (In Progress)'
        stats['import_count'] += 1
        stats['total'] += 1
        log_text = "นำเข้าสินค้า"
        item['receive_date'] = receive_date
        item['importer'] = importer
        
    elif action == 'export':
        if not export_date or not exporter:
            return jsonify({"success": False, "message": "⚠️ กรุณากรอก วันที่สแกนงานออก และชื่อผู้นำออกงาน ให้ครบถ้วนก่อนบันทึกนำออก!"}), 400
        item['status'] = 'สินค้าออกแล้ว (Exported)'
        stats['export_count'] += 1
        stats['total'] += 1
        log_text = "นำออกสินค้า"
        item['exporter'] = exporter
        item['export_date'] = export_date
    else:
        log_text = "ตรวจสอบสินค้า"

    item['department'] = dept
    item['job_type'] = job_type
    item['arch'] = data.get('arch', 'U/L')
    item['due_date'] = due_date
    item['material'] = material
    item['teeth_count'] = teeth_count
    item['retainer_color'] = retainer_color
    if receive_date: item['receive_date'] = receive_date
    if importer: item['importer'] = importer
    
    existing_log = next((l for l in logs if l['barcode'] == barcode), None)
    log_data = {
        "action": log_text,
        "barcode": barcode,
        "department": item['department'],
        "job_type": item['job_type'],
        "arch": item['arch'],
        "receive_date": item['receive_date'],
        "due_date": item['due_date'],
        "material": item['material'],
        "teeth_count": item['teeth_count'],
        "retainer_color": item['retainer_color'],
        "importer": item['importer'],
        "exporter": item['exporter'],
        "export_date": item['export_date'],
        "status": item['status']
    }

    if existing_log:
        logs.remove(existing_log)
    logs.insert(0, log_data)
                
    return jsonify({
        "success": True, 
        "message": "บันทึกข้อมูลงานเรียบร้อย", 
        "logs": logs,
        "stats": stats,
        "item": item
    })

@app.route('/api/download-excel/<type>', methods=['GET'])
def download_excel(type):
    if not logs:
        df = pd.DataFrame(columns=["การทำงาน", "บาร์โค้ด", "แผนก", "ประเภทงาน", "ตำแหน่งฟัน", "จำนวนซี่", "สีรีเทนเนอร์", "วันรับงาน", "วันกำหนดออก", "วัสดุ", "ผู้นำเข้า", "ผู้นำออก", "วันที่ออก", "สถานะ"])
    else:
        if type == 'import':
            filtered_logs = [l for l in logs if l['action'] == 'นำเข้าสินค้า']
        elif type == 'export':
            filtered_logs = [l for l in logs if l['action'] == 'นำออกสินค้า']
        else:
            filtered_logs = logs
            
        df = pd.DataFrame(filtered_logs)
        df = df.rename(columns={
            "action": "การทำงาน",
            "barcode": "บาร์โค้ด",
            "department": "แผนก",
            "job_type": "ประเภทงาน",
            "arch": "ตำแหน่งฟัน",
            "teeth_count": "จำนวนซี่",
            "retainer_color": "สีรีเทนเนอร์",
            "receive_date": "วันรับงาน",
            "due_date": "วันกำหนดออก",
            "material": "วัสดุ",
            "importer": "ผู้นำเข้า",
            "exporter": "ผู้นำออก",
            "export_date": "วันที่ออก",
            "status": "สถานะ"
        })

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='รายงานข้อมูล')
    output.seek(0)
    
    filename_map = {
        'all': 'inventory_all_stats.xlsx',
        'import': 'inventory_import_stats.xlsx',
        'export': 'inventory_export_stats.xlsx'
    }
    
    return send_file(output, download_name=filename_map.get(type, 'inventory_stats.xlsx'), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)