from flask import Flask, request
from flask_cors import CORS
import xlrd, xlwt
from xlutils.copy import copy
import os
import re
from datetime import datetime

app = Flask(__name__)
CORS(app)

OUTPUT_PATH = r'C:\EXCEL_PYTHON'
FILE_NAME = 'EXCEL CERTIFICADOS.xls'
FULL_PATH = os.path.join(OUTPUT_PATH, FILE_NAME)

def get_mapping(lente_raw, trat_raw):
    # Regras de negócio para conversão de produtos
    lente = lente_raw.upper()
    tratamento = trat_raw.upper()

    if any(x in lente for x in ['1.59 TGNS CINZA', '1.59 FOTO CINZA']):
        return "LENTE OC SEG MULT 1.59", f"{tratamento} / FOTOSSENSIVEL" if tratamento else "FOTOSSENSIVEL"
    
    if any(x in lente for x in ['1.50 TGNS CINZA', '1.50 FOTO CINZA']):
        return "LENTE OC SEG MULT 1.50", f"{tratamento} / FOTOSSENSIVEL" if tratamento else "FOTOSSENSIVEL"

    defaults = {
        'LP PRO DESIGN 1.59 CLEAR': 'LENTE OC SEG VS 1.59',
        'VS GEN 1.59 CLEAR': 'LENTE OC SEG VS 1.59',
        'MF TECHNOPARK 1.59 CLEAR': 'LENTE OC SEG MULT 1.59',
        'MF TECHNOPARK 1.50 CLEAR': 'LENTE OC SEG MULT 1.50'
    }
    
    return defaults.get(lente, lente), tratamento

def process_xls_write(dados):
    if not os.path.exists(OUTPUT_PATH):
        os.makedirs(OUTPUT_PATH)

    if not os.path.exists(FULL_PATH):
        wb_init = xlwt.Workbook()
        ws_init = wb_init.add_sheet('Dados')
        headers = ['CLIENTE', 'EMPRESA', 'LENTE', 'TRATAMENTO', 'ESF OD', 'CIL OD', 'EIXO OD', 'ADD OD', 'DNP OD', 'ALT OD', 'ESF OE', 'CIL OE', 'EIXO OE', 'ADD OE', 'DNP OE', 'ALT OE', 'OS', 'DATA']
        for col, h in enumerate(headers): ws_init.write(0, col, h)
        wb_init.save(FULL_PATH)

    rb = xlrd.open_workbook(FULL_PATH, formatting_info=True)
    wb = copy(rb)
    ws = wb.get_sheet(0)
    
    raw_p = dados.get('paciente', '')
    cliente, empresa = (raw_p.split('/', 1) + ["NORTEL"])[:2] if '/' in raw_p else (raw_p, "NORTEL")
    
    lente_f, trat_f = get_mapping(dados.get('lente', ''), dados.get('tratamento', ''))
    
    os_id = int(re.sub(r'\D', '', str(dados.get('os', '0')))) or 0

    row_vals = [
        cliente.strip(), empresa.strip(), lente_f, trat_f,
        dados.get('esf_od'), dados.get('cil_od'), dados.get('eixo_od'), dados.get('add_od'),
        dados.get('dnp_od'), dados.get('alt_od'), dados.get('esf_oe'), dados.get('cil_oe'),
        dados.get('eixo_oe'), dados.get('add_oe'), dados.get('dnp_oe'), dados.get('alt_oe'),
        os_id, datetime.now().strftime("%d/%m/%Y")
    ]

    for i, val in enumerate(row_vals):
        ws.write(rb.sheet_by_index(0).nrows, i, val)

    wb.save(FULL_PATH)

@app.route('/salvar', methods=['POST'])
def api_endpoint():
    try:
        data = request.get_json(force=True)
        process_xls_write(data)
        return {"status": "success"}, 200
    except Exception as e:
        print(f"Runtime Error: {e}")
        return {"status": "error"}, 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)