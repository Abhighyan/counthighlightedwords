from flask import Flask, request, render_template
from docx import Document
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)

# Allowable extensions for document upload
ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Function to count highlighted words
def count_highlighted_words(docx_file):
    doc = Document(docx_file)
    highlighted_words = 0
    for para in doc.paragraphs:
        for run in para.runs:
            if run.font.highlight_color:
                highlighted_words += len(run.text.split())
    return highlighted_words

# Route for homepage
@app.route('/')
def upload_file():
    return render_template('upload.html')

# Route to handle file upload and processing
@app.route('/upload', methods=['POST'])
def upload_and_count():
    if 'file' not in request.files:
        return "No file part"
    file = request.files['file']
    if file.filename == '':
        return "No selected file"
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file.save(os.path.join('.', filename))
        highlighted_word_count = count_highlighted_words(filename)
        os.remove(filename)  # Clean up uploaded file
        return f"Number of highlighted words: {highlighted_word_count}"
    return "Invalid file type. Please upload a .docx file."

if __name__ == '__main__':
    # Use the PORT environment variable provided by Railway
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
