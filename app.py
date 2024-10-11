from flask import Flask, request, render_template, jsonify
from docx import Document
from textblob import TextBlob
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Allowable extensions for document upload
ALLOWED_EXTENSIONS = {'docx'}

# Check if the file is a valid .docx file
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Count the highlighted words in the .docx document
def count_highlighted_words(docx_file):
    doc = Document(docx_file)
    highlighted_words = 0
    full_word_count = 0
    highlighted_word_color_count = {}
    
    for para in doc.paragraphs:
        for run in para.runs:
            full_word_count += len(run.text.split())
            if run.font.highlight_color:
                highlighted_words += len(run.text.split())
                # Count highlighted words by color
                if run.font.highlight_color not in highlighted_word_color_count:
                    highlighted_word_color_count[run.font.highlight_color] = 0
                highlighted_word_color_count[run.font.highlight_color] += len(run.text.split())

    return highlighted_words, highlighted_word_color_count, full_word_count

# Perform sentiment analysis using TextBlob
def perform_sentiment_analysis(text):
    blob = TextBlob(text)
    polarity = blob.sentiment.polarity
    subjectivity = blob.sentiment.subjectivity
    return polarity, subjectivity

@app.route('/')
def upload_file():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_and_count():
    if 'file' not in request.files:
        return jsonify({"status": "error", "message": "No file part"})
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({"status": "error", "message": "No selected file"})
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file.save(os.path.join('.', filename))
        
        # Perform word count and highlight analysis
        highlighted_word_count, color_counts, full_word_count = count_highlighted_words(filename)
        
        os.remove(filename)  # Clean up uploaded file

        color_percentage_details = "\n".join([
            f"Highlighted in {color}: {count} words ({(count / full_word_count * 100):.2f}% of total)"
            for color, count in color_counts.items()
        ])

        return jsonify({
            "status": "success",
            "full_word_count": full_word_count,
            "highlighted_word_count": highlighted_word_count,
            "highlighted_word_percentage": (highlighted_word_count / full_word_count * 100),
            "color_percentage_details": color_percentage_details
        })

    return jsonify({"status": "error", "message": "Invalid file type. Please upload a .docx file."})


if __name__ == '__main__':
    # Use the PORT environment variable provided by Railway
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
