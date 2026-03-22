import os
import io
import tempfile
from flask import Flask, request, send_file, jsonify
from werkzeug.utils import secure_filename
import pandas as pd
from btg_consolidador import consolidar

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB

ALLOWED_EXT = {'xlsx', 'xls'}

def allowed(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT

@app.route('/')
def index():
    return app.send_static_file('index.html')

@app.route('/processar', methods=['POST'])
def processar():
    try:
        extrato        = request.files.get('extrato')
        template       = request.files.get('template')
        nome_portfolio = request.form.get('nome_portfolio', 'JBL_Onshore').strip() or 'JBL_Onshore'

        if not extrato or not template:
            return jsonify({'erro': 'Envie os dois arquivos: extrato e planilha template.'}), 400
        if not allowed(extrato.filename) or not allowed(template.filename):
            return jsonify({'erro': 'Apenas arquivos .xlsx são suportados.'}), 400

        with tempfile.TemporaryDirectory() as tmpdir:
            ext_path = os.path.join(tmpdir, 'extrato.xlsx')
            tpl_path = os.path.join(tmpdir, 'template.xlsx')
            out_path = os.path.join(tmpdir, 'saida.xlsx')

            extrato.save(ext_path)
            template.save(tpl_path)

            consolidar(ext_path, tpl_path, out_path, nome_portfolio=nome_portfolio)

            with open(out_path, 'rb') as f:
                data = f.read()

        return send_file(
            io.BytesIO(data),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='comdinheiro_preenchida.xlsx'
        )

    except Exception as e:
        return jsonify({'erro': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=7860, debug=False)
