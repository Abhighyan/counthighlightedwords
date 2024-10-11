from flask import Flask, request, render_template
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
    # Check if the post request has the file part
    if 'file' not in request.files:
        return "No file part"
    
    file = request.files['file']
    
    # Check if the file is empty or no file is selected
    if file.filename == '':
        return "No selected file"
    
    # Process the file if it is allowed
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join('.', filename)
        file.save(filepath)
        
        # Count highlighted words
        highlighted_word_count, highlighted_word_color_count, full_word_count = count_highlighted_words(filepath)
        
        # Read file content for sentiment analysis
        doc = Document(filepath)
        text = "\n".join([para.text for para in doc.paragraphs])
        polarity, subjectivity = perform_sentiment_analysis(text)
        
        os.remove(filepath)  # Clean up uploaded file after processing
        
        # Prepare output
        output = (
            f"<h1>File Analysis Results</h1>"
            f"<h2>Sentiment Analysis:</h2>"
            f"<p><strong>Polarity:</strong> {polarity:.2f} (Scale from -1 to 1, where -1 is negative, 1 is positive)</p>"
            f"<p><strong>Subjectivity:</strong> {subjectivity:.2f} (Scale from 0 to 1, where 0 is very objective and 1 is very subjective)</p>"
            f"<br>"
            f"<h2>Highlighted Words Analysis:</h2>"
            f"<p><strong>Total Word Count:</strong> {full_word_count}</p>"
            f"<p><strong>Number of Highlighted Words:</strong> {highlighted_word_count} "
            f"({(highlighted_word_count/full_word_count)*100:.2f}% of total word count)</p>"
        )

        # Display highlighted words by color
        for color, count in highlighted_word_color_count.items():
            output += f"<p><strong>Words highlighted in {color}:</strong> {count}</p>"
        
        return output

    return "Invalid file type. Please upload a .docx file."

if __name__ == '__main__':
    # Use the PORT environment variable provided by Railway
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
