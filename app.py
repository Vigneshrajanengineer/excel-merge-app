from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from werkzeug.utils import secure_filename
app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
@app.route('/')
def index():
    return render_template('index.html')
@app.route('/merge', methods=['POST'])
def merge():
    sheet_name = request.form.get('sheet_name') or 0
    files = request.files.getlist('files')
    dfs = []
    for file in files:
        if file and file.filename.endswith(('.xlsx', '.xls')):
            filename = secure_filename(file.filename)
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(file_path)
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                dfs.append(df)
            except Exception as e:
                return f"Error reading {filename}: {e}"
    if not dfs:
        return "No valid Excel files uploaded."
    merged_df = pd.concat(dfs, ignore_index=True)
    output_path = os.path.join(UPLOAD_FOLDER, 'merged_output.xlsx')
    merged_df.to_excel(output_path, index=False)
    return send_file(output_path, as_attachment=True)

