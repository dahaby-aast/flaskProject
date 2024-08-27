from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Change this to a real secret key
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    filenames = os.listdir(UPLOAD_FOLDER)
    return render_template('index.html', filenames=filenames)

@app.route('/add_column', methods=['GET', 'POST'])
def add_column():
    if request.method == 'POST':
        files = request.files.getlist('files')
        if len(files) != 6:
            flash('Please upload exactly 6 files.', 'warning')
            return redirect(url_for('add_column'))

        prefix = request.form['prefix']
        locations = request.form.getlist('location')

        for file, location in zip(files, locations):
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(file_path)
                df = pd.read_excel(file_path)
                df['Location'] = location
                new_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{prefix}_{filename}")
                df.to_excel(new_file_path, index=False)
            else:
                flash('Invalid file type. Only .xlsx files are allowed.', 'danger')
                return redirect(url_for('add_column'))

        flash('Files processed successfully.', 'success')
        return redirect(url_for('index'))

    return render_template('add_column.html')

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.errorhandler(404)
def page_not_found(e):
    return render_template('error.html', error="404 Not Found"), 404

@app.errorhandler(500)
def internal_error(e):
    return render_template('error.html', error="500 Internal Server Error"), 500

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)
