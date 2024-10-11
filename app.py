from flask import Flask, request, render_template
from docx import Document
from werkzeug.utils import secure_filename
from textblob import TextBlob
import os

app = Flask(__name__)

# Allowable extensions for document upload
ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Function to count highlighted words and colors
def count_highlighted_words(docx_file):
    doc = Document(docx_file)
    highlighted_words = 0
    full_word_count = 0
    highlight_colors = {}
    for para in doc.paragraphs:
        for run in para.runs:
            full_word_count += len(run.text.split())
            if run.font.highlight_color:
                highlighted_words += len(run.text.split())
                color = run.font.highlight_color
                if color not in highlight_colors:
                    highlight_colors[color] = len(run.text.split())
                else:
                    highlight_colors[color] += len(run.text.split())
    return highlighted_words, highlight_colors, full_word_count

# Function to perform sentiment analysis using TextBlob
def analyze_sentiment(docx_file):
    doc = Document(docx_file)
    full_text = ''
    for para in doc.paragraphs:
        full_text += para.text + ' '
    
    blob = TextBlob(full_text)
    polarity = blob.sentiment.polarity  # Range: -1 to 1, where -1 is negative, 0 is neutral, 1 is positive
    subjectivity = blob.sentiment.subjectivity  # Range: 0 to 1, where 0 is objective and 1 is subjective
    
    sentiment_analysis = (
        f"Polarity: {polarity} (Polarity measures how positive or negative the text is. -1 is very negative, 1 is very positive, and 0 is neutral.)<br>"
        f"Subjectivity: {subjectivity} (Subjectivity measures how subjective or objective the text is. 0 is very objective, 1 is very subjective.)<br>"
    )
    
    return sentiment_analysis

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
        
        # Get highlighted word count, highlight colors, and full word count
        highlighted_word_count, highlight_colors, full_word_count = count_highlighted_words(filename)
        
        # Get sentiment analysis
        sentiment_analysis = analyze_sentiment(filename)
        
        # Format the highlight count details
        highlight_color_details = ""
        for color, count in highlight_colors.items():
            percentage = (count / full_word_count) * 100 if full_word_count > 0 else 0
            highlight_color_details += f"Highlighted in {color}: {count} ({percentage:.2f}% of total word count)<br>"
        
        # Remove the uploaded file
        os.remove(filename)  # Clean up uploaded file
        
        # Check for division by zero when calculating percentages
        highlighted_percentage = (highlighted_word_count / full_word_count) * 100 if full_word_count > 0 else 0
        
        # Return the results with line breaks and formatting
        return (
            f"<h3>Sentiment Analysis:</h3>{sentiment_analysis}"
            "<br><br>"  # Add space between sections
            f"<h3>Highlight Count:</h3>"
            f"Number of highlighted words: {highlighted_word_count} ({highlighted_percentage:.2f}% of total word count)<br>"
            f"Total word count: {full_word_count}<br>"
            f"{highlight_color_details}"
        )
    return "Invalid file type. Please upload a .docx file."

if __name__ == '__main__':
    # Use the PORT environment variable provided by Railway
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
