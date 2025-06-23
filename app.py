from flask import Flask, request, render_template, send_file, flash, redirect, url_for
import os
from werkzeug.utils import secure_filename
import io

# Import a function from our logic file
from logic_handler import get_static_data_from_excel, process_excel_file

# --- Basic Flask App Setup ---
app = Flask(__name__)
# A secret key is needed for flashing messages
app.config['SECRET_KEY'] = 'your_super_secret_key_12345'
# Define the path for the data file
DATA_FILE_PATH = "Data.xlsx"

# --- Main Route to Display the Upload Page ---
@app.route('/', methods=['GET'])
def index():
    """Renders the main upload page."""
    # Load data for the dropdown menu
    static_data = get_static_data_from_excel(DATA_FILE_PATH)
    if static_data:
        chxd_list = static_data.get("listbox_data", [])
    else:
        chxd_list = []
        flash("Lỗi: Không thể đọc file Data.xlsx. Vui lòng kiểm tra lại file.", "danger")
        
    return render_template('index.html', chxd_list=chxd_list)

# --- Route to Handle File Processing ---
@app.route('/process', methods=['POST'])
def process():
    """Handles the file upload and processing."""
    # Check if a file was uploaded
    if 'file' not in request.files:
        flash('Không có file nào được tải lên.', 'warning')
        return redirect(url_for('index'))

    file = request.files['file']
    selected_chxd = request.form.get('chxd')

    # Check if the user selected a CHXD and uploaded a file
    if file.filename == '' or not selected_chxd:
        flash('Vui lòng chọn CHXD và tải lên file bảng kê.', 'warning')
        return redirect(url_for('index'))

    if file:
        try:
            # Read the content of the uploaded file into memory
            file_content = file.read()
            
            # Load static data needed for processing
            static_data = get_static_data_from_excel(DATA_FILE_PATH)
            if not static_data:
                raise ValueError("Không thể tải dữ liệu tĩnh từ Data.xlsx.")

            # Call the main processing function from our logic handler
            result_buffer = process_excel_file(file_content, static_data, selected_chxd)
            
            # Send the processed file back to the user for download
            return send_file(
                result_buffer,
                as_attachment=True,
                download_name='UpSSE.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        except Exception as e:
            # If any error occurs, flash the message and redirect to the main page
            flash(f'Đã xảy ra lỗi trong quá trình xử lý: {e}', 'danger')
            return redirect(url_for('index'))

    return redirect(url_for('index'))

# --- Run the App ---
if __name__ == '__main__':
    # This is for local development only. Render uses Gunicorn.
    app.run(debug=True)
