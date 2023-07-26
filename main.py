from flask import Flask, request, render_template
import zipfile
import subprocess
import sys
import os


app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'No file uploaded'

        file = request.files['file']
        if file.filename == '':
            return 'No file selected'

        if file and allowed_file(file.filename):
            file_path = 'uploads/' + file.filename
            file.save(file_path)
            document_xml_path = 'uploads/'
            document_rels_path = 'uploads/'
            image_folder_path = 'uploads/'
            extract_docx_contents(file_path, document_xml_path, document_rels_path, image_folder_path)
            subprocess.run([sys.executable, 'convert.py'])
            return 'document.xml, document.xml.rels, and images extracted and saved successfully'
    return render_template('upload.html')


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'doc', 'docx'}


def extract_docx_contents(docx_file_path, document_xml_path, document_rels_path, image_folder_path):
    with zipfile.ZipFile(docx_file_path, 'r') as zip_ref:
        zip_ref.extract('word/document.xml', document_xml_path)
        zip_ref.extract('word/_rels/document.xml.rels', document_rels_path)
        os.makedirs(image_folder_path, exist_ok=True)
        for item in zip_ref.namelist():
            if item.startswith('word/media/'):
                image_filename = os.path.basename(item)
                image_destination = os.path.join(image_folder_path, "media", image_filename)
                os.makedirs(os.path.dirname(image_destination), exist_ok=True)
                with zip_ref.open(item) as source, open(image_destination, "wb") as destination:
                    destination.write(source.read())


if __name__ == '__main__':
    app.run()