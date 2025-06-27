from flask import Flask, request, render_template, send_file, flash, redirect, url_for
import io
import zipfile

# Import the main processing function from our logic file
from logic_handler import get_static_data_from_excel, process_file_with_price_periods

# --- Basic Flask App Setup ---
app = Flask(__name__)
app.config['SECRET_KEY'] = 'your_super_secret_key_12345'
DATA_FILE_PATH = "Data.xlsx"

# --- Main Route to Display the Upload Page ---
@app.route('/', methods=['GET'])
def index():
    """Renders the main upload page."""
    try:
        static_data = get_static_data_from_excel(DATA_FILE_PATH)
        if static_data:
            chxd_list = static_data.get("listbox_data", [])
        else:
            chxd_list = []
            flash("Lỗi: Không thể đọc file Data.xlsx. Vui lòng kiểm tra lại file.", "danger")
    except Exception as e:
        chxd_list = []
        flash(f"Lỗi nghiêm trọng khi đọc Data.xlsx: {e}", "danger")
        
    return render_template('index.html', chxd_list=chxd_list)

# --- Route to Handle File Processing ---
@app.route('/process', methods=['POST'])
def process():
    """Handles the file upload and processing based on price periods."""
    if 'file' not in request.files:
        flash('Không có file nào được tải lên.', 'warning')
        return redirect(url_for('index'))

    file = request.files['file']
    selected_chxd = request.form.get('chxd')
    price_periods = request.form.get('price_periods')
    invoice_number = request.form.get('invoice_number', '').strip()

    if file.filename == '':
        flash('Vui lòng tải lên file bảng kê.', 'warning')
        return redirect(url_for('index'))

    if not selected_chxd:
        flash('Vui lòng chọn CHXD.', 'warning')
        return redirect(url_for('index'))

    if price_periods == '2' and not invoice_number:
        flash('Vui lòng nhập "Số hóa đơn đầu tiên của giá mới" khi chọn 2 giai đoạn giá.', 'warning')
        return redirect(url_for('index'))

    if file:
        try:
            file_content = file.read()
            
            static_data = get_static_data_from_excel(DATA_FILE_PATH)
            if not static_data:
                raise ValueError("Không thể tải dữ liệu tĩnh từ Data.xlsx.")

            # --- ***** START OF CHANGE: HANDLE RETURN VALUE ***** ---
            result = process_file_with_price_periods(
                uploaded_file_content=file_content, 
                static_data=static_data, 
                selected_chxd=selected_chxd,
                price_periods=price_periods,
                new_price_invoice_number=invoice_number
            )
            
            # Nếu kết quả là một dictionary, nghĩa là có 2 file cần nén lại
            if isinstance(result, dict):
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    if result.get('new'):
                        result['new'].seek(0)
                        zipf.writestr('UpSSE_gia_moi.xlsx', result['new'].read())
                    if result.get('old'):
                        result['old'].seek(0)
                        zipf.writestr('UpSSE_gia_cu.xlsx', result['old'].read())
                
                zip_buffer.seek(0)
                return send_file(
                    zip_buffer,
                    as_attachment=True,
                    download_name='UpSSE_2_giai_doan.zip',
                    mimetype='application/zip'
                )
            # Nếu không, trả về 1 file như bình thường
            else:
                return send_file(
                    result,
                    as_attachment=True,
                    download_name='UpSSE.xlsx',
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            # --- ***** END OF CHANGE ***** ---

        except ValueError as ve:
            flash(f'{ve}', 'danger')
            return redirect(url_for('index'))
        except Exception as e:
            flash(f'Đã xảy ra lỗi không xác định: {e}', 'danger')
            return redirect(url_for('index'))

    return redirect(url_for('index'))

# --- Run the App ---
if __name__ == '__main__':
    app.run(debug=True)
