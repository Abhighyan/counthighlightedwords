from flask import Flask, request, render_template, jsonify
from docx import Document
from textblob import TextBlob
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

ALLOWED_EXTENSIONS = {'docx'}
MAX_FILE_SIZE = 5 * 1024 * 1024  # 5 MB

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def count_highlighted_words(docx_file):
    doc = Document(docx_file)
    highlighted_words = 0
    full_word_count = 0
    highlighted_word_color_count = {}
    
    for para in doc.paragraphs:
        for run in para.runs:
            words = run.text.split()
            full_word_count += len(words)
            if run.font.highlight_color:
                highlighted_words += len(words)
                highlighted_word_color_count[run.font.highlight_color] = highlighted_word_color_count.get(run.font.highlight_color, 0) + len(words)
    
    return highlighted_words, highlighted_word_color_count, full_word_count

def perform_sentiment_analysis(docx_file):
    doc = Document(docx_file)
    full_text = ' '.join(para.text for para in doc.paragraphs)
    blob = TextBlob(full_text)
    return {
        "polarity": blob.sentiment.polarity,
        "subjectivity": blob.sentiment.subjectivity,
        "positive": sum(1 for word in full_text.split() if TextBlob(word).sentiment.polarity > 0),
        "negative": sum(1 for word in full_text.split() if TextBlob(word).sentiment.polarity < 0),
        "neutral": sum(1 for word in full_text.split() if TextBlob(word).sentiment.polarity == 0),
    }

@app.route('/')
def upload_file():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_and_count():
    if 'file' not in request.files:
        return jsonify(status='error', message="No file part")
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify(status='error', message="No selected file")
    
    if file and allowed_file(file.filename):
        if file.content_length > MAX_FILE_SIZE:
            return jsonify(status='error', message="File size exceeds the 5MB limit")
        
        filename = secure_filename(file.filename)
        file_path = os.path.join('/tmp', filename)
        file.save(file_path)
        
        try:
            highlighted_word_count, color_counts, full_word_count = count_highlighted_words(file_path)
            sentiment_data = perform_sentiment_analysis(file_path)
        except Exception as e:
            os.remove(file_path)
            return jsonify(status='error', message=f"Error processing file: {str(e)}")
        
        os.remove(file_path)
        
        color_percentage_details = "<br>".join([f"Highlighted in {color}: {count} words ({(count / full_word_count * 100):.2f}% of total)"
                                                 for color, count in color_counts.items()])
        highlighted_word_percentage = (highlighted_word_count / full_word_count * 100) if full_word_count else 0
        
        return jsonify(status='success', 
                       full_word_count=full_word_count, 
                       highlighted_word_count=highlighted_word_count,
                       highlighted_word_percentage=highlighted_word_percentage,
                       sentiment_data=sentiment_data,
                       color_percentage_details=color_percentage_details)
    
    return jsonify(status='error', message="Invalid file type. Please upload a .docx file.")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
