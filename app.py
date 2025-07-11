import base64
import io
import os
import zipfile
from flask import Flask, flash, redirect, render_template, request, send_file, url_for, get_flashed_messages # Import get_flashed_messages
from openpyxl import load_workbook

# --- CÁC IMPORT CHO CÁC HANDLER ---
# Giả định bạn có file detector.py để nhận diện loại file
from detector import detect_report_type 
from hddt_handler import process_hddt_report
from pos_handler import process_pos_report
from doisoat_handler import perform_reconciliation
# START: THÊM IMPORT CHO THEKHO_HANDLER MỚI
from TheKho_handler import process_stock_card_data
# END: THÊM IMPORT CHO THEKHO_HANDLER MỚI

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'a_very_strong_and_unified_secret_key')

def get_chxd_list():
    """Đọc danh sách CHXD trực tiếp từ cột D của file Data_HDDT.xlsx."""
    chxd_list = []
    try:
        # Sử dụng file cấu hình chung
        wb = load_workbook("Data_HDDT.xlsx", data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=3, min_col=4, max_col=4, values_only=True):
            chxd_name = row[0]
            if chxd_name and isinstance(chxd_name, str) and chxd_name.strip():
                chxd_list.append(chxd_name.strip())
        chxd_list.sort()
        return chxd_list
    except FileNotFoundError:
        flash("Lỗi nghiêm trọng: Không tìm thấy file cấu hình Data_HDDT.xlsx!", "danger")
        return []
    except Exception as e:
        flash(f"Lỗi khi đọc file Data_HDDT.xlsx: {e}", "danger")
        return []

@app.route('/', methods=['GET'])
def index():
    """Hiển thị trang upload chính."""
    chxd_list = get_chxd_list()
    # THAY ĐỔI: Lấy active_tab từ query parameter nếu có, để duy trì tab khi redirect
    active_tab = request.args.get('active_tab', 'upsse') 
    return render_template('index.html', chxd_list=chxd_list, form_data={"active_tab": active_tab})

@app.route('/process', methods=['POST'])
def process():
    """Xử lý file tải lên cho chức năng UpSSE."""
    chxd_list = get_chxd_list()
    form_data = {
        "selected_chxd": request.form.get('chxd'),
        "price_periods": request.form.get('price_periods', '1'),
        "invoice_number": request.form.get('invoice_number', '').strip(),
        "confirmed_date": request.form.get('confirmed_date'),
        "encoded_file": request.form.get('file_content_b64')
    }
    
    try:
        if not form_data["selected_chxd"]:
            flash('Vui lòng chọn CHXD.', 'warning')
            return redirect(url_for('index', active_tab='upsse')) # Giữ tab active

        file_content = None
        if form_data["encoded_file"]:
            file_content = base64.b64decode(form_data["encoded_file"])
        elif 'file' in request.files and request.files['file'].filename != '':
            file_content = request.files['file'].read()
        else:
            flash('Vui lòng tải lên file Bảng kê.', 'warning')
            return redirect(url_for('index', active_tab='upsse')) # Giữ tab active

        report_type = detect_report_type(file_content)
        result = None

        if report_type == 'POS':
            result = process_pos_report(
                file_content_bytes=file_content,
                selected_chxd=form_data["selected_chxd"],
                price_periods=form_data["price_periods"],
                new_price_invoice_number=form_data["invoice_number"]
            )
        elif report_type == 'HDDT':
            result = process_hddt_report(
                file_content_bytes=file_content,
                selected_chxd=form_data["selected_chxd"],
                price_periods=form_data["price_periods"],
                new_price_invoice_number=form_data["invoice_number"],
                confirmed_date_str=form_data["confirmed_date"]
            )
        else:
            raise ValueError("Không thể tự động nhận diện loại Bảng kê. Vui lòng kiểm tra lại file Excel bạn đã tải lên.")

        if isinstance(result, dict) and result.get('choice_needed'):
            form_data["encoded_file"] = base64.b64encode(file_content).decode('utf-8')
            return render_template('index.html', chxd_list=chxd_list, date_ambiguous=True, date_options=result['options'], form_data=form_data)
        
        elif isinstance(result, dict) and ('old' in result or 'new' in result):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                if result.get('old'):
                    result['old'].seek(0)
                    zipf.writestr('UpSSE_gia_cu.xlsx', result['old'].read())
                if result.get('new'):
                    result['new'].seek(0)
                    zipf.writestr('UpSSE_gia_moi.xlsx', result['new'].read())
            zip_buffer.seek(0)
            flash('Xử lý Đồng bộ SSE thành công!', 'success') # Flash message
            return send_file(zip_buffer, as_attachment=True, download_name='UpSSE_2_giai_doan.zip', mimetype='application/zip')

        elif isinstance(result, io.BytesIO):
            result.seek(0)
            flash('Xử lý Đồng bộ SSE thành công!', 'success') # Flash message
            return send_file(result, as_attachment=True, download_name='UpSSE.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        else:
            raise ValueError("Hàm xử lý không trả về kết quả hợp lệ.")

    except ValueError as ve:
        flash(str(ve).replace('\n', '<br>'), 'danger')
        return render_template('index.html', chxd_list=chxd_list, form_data=form_data)
    except Exception as e:
        flash(f"Đã xảy ra lỗi không mong muốn: {e}", 'danger')
        return render_template('index.html', chxd_list=chxd_list, form_data=form_data)

@app.route('/reconcile', methods=['POST'])
def reconcile():
    """Xử lý file tải lên cho chức năng Đối soát (giữ nguyên)."""
    chxd_list = get_chxd_list()
    reconciliation_data = None
    try:
        selected_chxd = request.form.get('chxd')
        file_log_bom = request.files.get('file_log_bom')
        file_hddt = request.files.get('file_hddt')

        if not selected_chxd or not file_log_bom or not file_hddt:
            flash('Vui lòng chọn CHXD và tải lên đủ cả 2 file để đối soát.', 'warning')
            return redirect(url_for('index', active_tab='doisoat')) # Giữ tab active

        log_bom_bytes = file_log_bom.read()
        hddt_bytes = file_hddt.read()
        reconciliation_data = perform_reconciliation(log_bom_bytes, hddt_bytes, selected_chxd)
        
        if reconciliation_data:
             flash('Đối soát thành công!', 'success')
        else:
             flash('Không có dữ liệu trả về từ chức năng đối soát.', 'warning')

    except Exception as e:
        flash(f"Lỗi trong quá trình đối soát: {e}", 'danger')

    return render_template('index.html', 
                           chxd_list=chxd_list, 
                           reconciliation_data=reconciliation_data,
                           form_data={"active_tab": "doisoat"}) # Giữ tab active

# START: ROUTE MỚI CHO CHỨC NĂNG THẺ KHO TỰ ĐỘNG
@app.route('/process_stock_card', methods=['POST'])
def process_stock_card():
    """
    Xử lý file ảnh/PDF tải lên cho chức năng Thẻ kho tự động.
    Sử dụng Gemini API để trích xuất dữ liệu và tạo file Excel.
    """
    chxd_list = get_chxd_list()
    selected_chxd = request.form.get('chxd_thekho') # Lấy CHXD từ form thẻ kho

    try:
        if not selected_chxd:
            flash('Vui lòng chọn Cửa Hàng Xăng Dầu (CHXD) cho chức năng Thẻ kho.', 'warning')
            # Đảm bảo tab thẻ kho được hiển thị lại nếu có lỗi
            return render_template('index.html', chxd_list=chxd_list, form_data={"active_tab": "thekho"})

        uploaded_files = request.files.getlist('files[]')
        if not uploaded_files or all(f.filename == '' for f in uploaded_files):
            flash('Vui lòng tải lên ít nhất một file ảnh hoặc PDF.', 'warning')
            # Đảm bảo tab thẻ kho được hiển thị lại nếu có lỗi
            return render_template('index.html', chxd_list=chxd_list, form_data={"active_tab": "thekho"})

        # Gọi hàm xử lý chính trong TheKho_handler
        excel_buffer = process_stock_card_data(uploaded_files, selected_chxd)
        
        if excel_buffer:
            excel_buffer.seek(0)
            flash('Xử lý Thẻ kho tự động thành công!', 'success') # Flash message
            return send_file(
                excel_buffer,
                as_attachment=True,
                download_name='TheKho_TuDong.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            flash('Không có dữ liệu hợp lệ được trích xuất từ các file đã tải lên.', 'warning')
            # Đảm bảo tab thẻ kho được hiển thị lại nếu có lỗi
            return render_template('index.html', chxd_list=chxd_list, form_data={"active_tab": "thekho"})

    except ValueError as ve:
        # Bắt lỗi ValueError chi tiết từ TheKho_handler và hiển thị
        flash(str(ve).replace('\n', '<br>'), 'danger')
        # Đảm bảo tab thẻ kho được hiển thị lại nếu có lỗi
        return render_template('index.html', chxd_list=chxd_list, form_data={"active_tab": "thekho"})
    except Exception as e:
        flash(f"Đã xảy ra lỗi không mong muốn trong quá trình xử lý Thẻ kho: {e}", 'danger')
        # Đảm bảo tab thẻ kho được hiển thị lại nếu có lỗi
        return render_template('index.html', chxd_list=chxd_list, form_data={"active_tab": "thekho"})
# END: ROUTE MỚI CHO CHỨC NĂNG THẺ KHO TỰ ĐỘNG

# START: ROUTE MỚI ĐỂ XÓA FLASH MESSAGES
@app.route('/clear_flash_messages', methods=['GET'])
def clear_flash_messages():
    """
    Route này được gọi bởi JavaScript để xóa các thông báo flash trong session.
    """
    _ = get_flashed_messages() # Gọi hàm này để lấy và xóa tất cả messages
    return '', 204 # Trả về phản hồi rỗng với status code 204 (No Content)
# END: ROUTE MỚI ĐỂ XÓA FLASH MESSAGES

if __name__ == '__main__':
    # Cần có file detector.py trong cùng thư mục để chạy
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))

