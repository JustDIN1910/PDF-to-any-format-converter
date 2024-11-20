import os
import pdfplumber
import pandas as pd
from openpyxl import Workbook
from docx import Document
from pdf2image import convert_from_path
from flask import Flask, render_template, request, send_file, jsonify

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['pdf_file']
        output_format = request.form.get('output_format')

        if file and output_format:
            input_name = os.path.splitext(file.filename)[0]  # Get the name without extension
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)

            if output_format == 'excel':
                output_file = convert_pdf_to_excel(filepath, input_name)
            elif output_format == 'word':
                output_file = convert_pdf_to_word(filepath, input_name)
            elif output_format == 'text':
                output_file = convert_pdf_to_text(filepath, input_name)
            else:
                return jsonify({"error": "Invalid format selected."}), 400

            return send_file(output_file, as_attachment=True)
    return render_template('index.html')

def convert_pdf_to_excel(pdf_path, input_name):
    output_path = os.path.join(OUTPUT_FOLDER, f'{input_name}.xlsx')
    with pdfplumber.open(pdf_path) as pdf:
        all_text = []
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                df = pd.DataFrame(table[1:], columns=table[0])
                all_text.append(df)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for i, df in enumerate(all_text):
            df.to_excel(writer, index=False, sheet_name=f'Sheet{i+1}')
    return output_path

def convert_pdf_to_word(pdf_path, input_name):
    output_path = os.path.join(OUTPUT_FOLDER, f'{input_name}.docx')
    document = Document()
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                document.add_paragraph(text)
    document.save(output_path)
    return output_path

def convert_pdf_to_text(pdf_path, input_name):
    output_path = os.path.join(OUTPUT_FOLDER, f'{input_name}.txt')
    with pdfplumber.open(pdf_path) as pdf:
        with open(output_path, 'w') as f:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    f.write(text + '\n')
    return output_path

if __name__ == '__main__':
    app.run(debug=True)
