import os
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from werkzeug.utils import secure_filename
import utils

app = Flask(__name__)
app.secret_key = "super_secret_key_for_flash_messages"

# Configuration
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__name__)), 'uploads')
OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__name__)), 'outputs')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max

ALLOWED_EXTENSIONS = {'pdf', 'docx'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file part')
        return redirect(url_for('index'))
    
    file = request.files['file']
    if file.filename == '':
        flash('No selected file')
        return redirect(url_for('index'))
    
    old_year = request.form.get('old_year')
    new_year = request.form.get('new_year')
    
    if not old_year or not new_year:
        flash('Please provide both the old year and the new year.')
        return redirect(url_for('index'))
        
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        output_filename = f"updated_{filename}"
        output_filepath = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        
        try:
            # Process based on extension
            ext = filename.rsplit('.', 1)[1].lower()
            if ext == 'docx':
                utils.replace_year_in_docx(filepath, output_filepath, old_year, new_year)
            elif ext == 'pdf':
                utils.replace_year_in_pdf(filepath, output_filepath, old_year, new_year)
                
            return send_file(output_filepath, as_attachment=True, download_name=output_filename)
            
        except Exception as e:
            flash(f'Error processing file: {str(e)}')
            return redirect(url_for('index'))
            
    else:
        flash('Allowed file types are docx, pdf')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
