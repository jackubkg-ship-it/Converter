import os
from flask import Flask, request, send_file, render_template
from converter import convert_workbook

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "No file part", 400

    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    if not file.filename.lower().endswith(('.xlsx', '.xlsm')):
        return "Invalid file type. Only .xlsx and .xlsm are allowed.", 400

    totals = '1' in request.form.getlist('totals')

    try:
        output = convert_workbook(file, split_sheets=True, add_totals=totals)
        base = os.path.splitext(file.filename)[0]
        out_name = base + "_Avto326.xlsx"
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=out_name
        )
    except Exception as e:
        app.logger.exception("Conversion error")
        return "Conversion error: " + str(e), 500


if __name__ == '__main__':
    port = int(os.environ.get('FLASK_PORT', 5000))
    debug = os.environ.get('FLASK_DEBUG', '0') == '1'
    print(" * Starting Flask on http://localhost:" + str(port))
    app.run(debug=debug, host='0.0.0.0', port=port, use_reloader=False)