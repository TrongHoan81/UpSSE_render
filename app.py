import base64
import io
import os
import zipfile
from flask import Flask, flash, redirect, render_template, request, send_file, url_for, get_flashed_messages, jsonify
from openpyxl import load_workbook
import re  # Import re for _clean_string_app
import pandas as pd  # Import pandas for excel date conversion in static data loading
from collections import defaultdict  # Import defaultdict
from datetime import datetime  # <-- THÊM: dùng để xử lý ngày cho tên file

# --- CÁC IMPORT CHO CÁC HANDLER ---
from detector import detect_report_type
from hddt_handler import process_hddt_report
from pos_handler import process_pos_report
from doisoat_handler import perform_reconciliation, _load_discount_data, _generate_discount_report_excel  # Import _load_discount_data and _generate_discount_report_excel
from TheKho_handler import process_stock_card_data

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'a_very_strong_and_unified_secret_key')

# --- HÀM TIỆN ÍCH CHO VIỆC NẠP DỮ LIỆU CẤU HÌNH ---
def _clean_string_app(s):
    """Làm sạch chuỗi, loại bỏ khoảng trắng thừa và ký tự '."""
    if s is None:
        return ""
    cleaned_s = str(s).strip()
    if cleaned_s.startswith("'"):
        cleaned_s = cleaned_s[1:]
    return re.sub(r'\s+', ' ', cleaned_s)

def _to_float_app(value):
    """Chuyển đổi giá trị sang float, xử lý các trường hợp lỗi."""
    if value is None:
        return 0.0
    try:
        return float(str(value).replace(',', '').strip())
    except (ValueError, TypeError):
        return 0.0

# --- CUSTOM JINJA2 FILTER ---
@app.template_filter('format_currency')
def format_currency_filter(value):
    """
    Định dạng số thành chuỗi tiền tệ có dấu phẩy phân cách hàng nghìn.
    Sử dụng cho hiển thị trong template Jinja2.
    """
    try:
        num = float(value)
        return f"{num:,.0f}"
    except (ValueError, TypeError):
        return "0"

# =========================
#   HÀM PHỤ CHO TÊN FILE
# =========================

def _sanitize_filename_piece(s: str) -> str:
    """
    Làm sạch chuỗi để đưa vào tên file: loại ký tự không hợp lệ cho tên file.
    Giữ lại ký tự Unicode (kể cả có dấu), khoảng trắng và dấu chấm giai đoạn.
    """
    if not s:
        return "Untitled"
    s = s.strip()
    # Loại bỏ các ký tự gây lỗi tên file phổ biến trên nhiều hệ thống
    return re.sub(r'[\\/:*?"<>|\r\n]+', '_', s)

def _parse_date_like_hddt(cell_val):
    """
    Bắt chước logic parse ngày của hddt_handler._parse_date_from_excel_cell để phục vụ đặt tên file.
    Trả về đối tượng date nếu parse được, ngược lại None.
    """
    if cell_val is None:
        return None
    if isinstance(cell_val, datetime):
        return cell_val.date()

    if isinstance(cell_val, (int, float)):
        try:
            return pd.to_datetime(float(cell_val), unit='D', origin='1899-12-30').date()
        except Exception:
            return None

    if isinstance(cell_val, str):
        date_str = cell_val.strip()
        fmts = ['%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d', '%Y-%m-%d', '%d/%m/%y', '%d-%m-%y']
        for fmt in fmts:
            try:
                return datetime.strptime(date_str, fmt).date()
            except ValueError:
                continue
    return None

def _parse_date_like_pos(cell_val):
    """
    Bắt chước logic parse ngày của pos_handler (_pos_parse_date) ở mức tối thiểu để đặt tên file.
    Trả về đối tượng date nếu parse được, ngược lại None.
    """
    if cell_val is None:
        return None
    if isinstance(cell_val, datetime):
        return cell_val.date()

    if isinstance(cell_val, (int, float)):
        try:
            return pd.to_datetime(float(cell_val), unit='D', origin='1899-12-30').date()
        except Exception:
            return None

    if isinstance(cell_val, str):
        date_str = cell_val.strip()
        fmts = [
            '%Y-%m-%d %H:%M:%S',
            '%Y-%m-%d',
            '%d-%m-%Y %H:%M:%S',
            '%d-%m-%Y',
            '%d/%m/%Y %H:%M:%S',
            '%d/%m/%Y',
        ]
        for fmt in fmts:
            try:
                return datetime.strptime(date_str, fmt).date()
            except ValueError:
                continue
    return None

def _extract_report_date_for_filename(file_bytes: bytes, report_type: str, confirmed_date_str: str | None) -> datetime.date:
    """
    Trả về ngày (date) dùng đặt tên file.
    - HDDT: nếu có confirmed_date_str (YYYY-MM-DD) thì dùng luôn.
            ngược lại quét cột ngày của bảng kê như hddt_handler để tìm 1 ngày duy nhất.
    - POS:  quét cột ngày trong bảng kê POS; nếu nhiều ngày, lấy ngày nhỏ nhất.
    Nếu không tìm được, fallback về ngày hôm nay (để không làm hỏng luồng tải file).
    """
    try:
        if report_type == 'HDDT':
            if confirmed_date_str:
                try:
                    return datetime.strptime(confirmed_date_str, '%Y-%m-%d').date()
                except Exception:
                    pass
            # Quét file để lấy ngày duy nhất
            wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
            ws = wb.active
            unique_dates = set()
            for row in ws.iter_rows(min_row=11, values_only=True):
                qty = _to_float_app(row[8] if len(row) > 8 else None)
                if qty > 0:
                    dt = _parse_date_like_hddt(row[20] if len(row) > 20 else None)
                    if dt:
                        unique_dates.add(dt)
            wb.close()
            if unique_dates:
                # file chuẩn là 1 ngày; nếu lỡ nhiều, lấy min để ổn định
                return min(unique_dates)
        elif report_type == 'POS':
            wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
            ws = wb.active
            unique_dates = set()
            # Dòng dữ liệu POS bắt đầu từ row 5 (tham chiếu đọc hiện tại)
            for row in ws.iter_rows(min_row=5, values_only=True):
                # cột ngày của POS trong handler là row[3]
                dt = _parse_date_like_pos(row[3] if len(row) > 3 else None)
                if dt:
                    unique_dates.add(dt)
            wb.close()
            if unique_dates:
                return min(unique_dates)
    except Exception:
        pass

    # Fallback an toàn
    return datetime.today().date()

def _make_base_filename(store_name: str, file_date: datetime.date) -> str:
    store = _sanitize_filename_piece(store_name)
    date_part = f"{file_date.day:02d}.{file_date.month:02d}.{file_date.year}"
    return f"{store}.{date_part}"

# --- HÀM NẠP TẤT CẢ DỮ LIỆU CẤU HÌNH TĨNH ---
def load_all_static_config_data():
    """
    Tải tất cả dữ liệu cấu hình từ các file Excel tĩnh một lần duy nhất.
    Trả về một dictionary chứa các cấu hình cho POS và HDDT handlers.
    """
    static_data = {}
    try:
        # Load Data_HDDT.xlsx
        wb_hddt = load_workbook("Data_HDDT.xlsx", data_only=True)
        ws_hddt = wb_hddt.active

        # Data for POS handler
        chxd_detail_map_pos = {}
        store_specific_x_lookup_pos = {}

        # Data for HDDT handler
        chxd_list_for_hddt = []
        tk_mk_map_hddt = {}
        khhd_map_hddt = {}
        chxd_to_khuvuc_map_hddt = {}
        vu_viec_map_hddt = {}

        vu_viec_headers = [_clean_string_app(cell.value) for cell in ws_hddt[2][4:9]]

        for row_idx in range(3, ws_hddt.max_row + 1):
            row_values = [cell.value for cell in ws_hddt[row_idx]]

            if len(row_values) > 11:
                chxd_name = _clean_string_app(row_values[3])
                if chxd_name:
                    # For POS handler
                    chxd_detail_map_pos[chxd_name] = {
                        'g5_val': row_values[9],
                        'h5_val': _clean_string_app(row_values[11]).lower(),
                        'f5_val_full': _clean_string_app(row_values[10]),
                        'b5_val': chxd_name
                    }
                    store_specific_x_lookup_pos[chxd_name] = {
                        "xăng e5 ron 92-ii": row_values[4],
                        "xăng ron 95-iii": row_values[5],
                        "dầu do 0,05s-ii": row_values[6],
                        "dầu do 0,001s-v": row_values[7]
                    }

                    # For HDDT handler
                    if chxd_name not in chxd_list_for_hddt:
                        chxd_list_for_hddt.append(chxd_name)

                    ma_kho = _clean_string_app(row_values[9])
                    khhd = _clean_string_app(row_values[10])
                    khu_vuc = _clean_string_app(row_values[11])

                    if ma_kho:
                        tk_mk_map_hddt[chxd_name] = ma_kho
                    if khhd:
                        khhd_map_hddt[chxd_name] = khhd
                    if khu_vuc:
                        chxd_to_khuvuc_map_hddt[chxd_name] = khu_vuc

                    vu_viec_map_hddt[chxd_name] = {}
                    vu_viec_data_row = row_values[4:9]
                    for i, header in enumerate(vu_viec_headers):
                        if header:
                            key = "Dầu mỡ nhờn" if i == len(vu_viec_headers) - 1 else header
                            vu_viec_map_hddt[chxd_name][key] = _clean_string_app(vu_viec_data_row[i])

        def get_lookup_pos_config(min_r, max_r, min_c=1, max_c=2):
            return {
                _clean_string_app(row[0]).lower(): row[1]
                for row in ws_hddt.iter_rows(min_row=min_r, max_row=max_r, min_col=min_c, max_col=min_c + 1, values_only=True)
                if row[0] and row[1] is not None
            }

        tmt_lookup_table_pos = {k: _to_float_app(v) for k, v in get_lookup_pos_config(10, 13).items()}

        static_data['pos_config'] = {
            "lookup_table": get_lookup_pos_config(4, 7),
            "tmt_lookup_table": tmt_lookup_table_pos,
            "s_lookup_table": get_lookup_pos_config(29, 31),
            "t_lookup_regular": get_lookup_pos_config(33, 35),
            "t_lookup_tmt": get_lookup_pos_config(48, 50),
            "v_lookup_table": get_lookup_pos_config(53, 55),
            "u_value": ws_hddt['B36'].value,
            "chxd_detail_map": chxd_detail_map_pos,
            "store_specific_x_lookup": store_specific_x_lookup_pos
        }

        def get_lookup_hddt_config(min_r, max_r, min_c=1, max_c=2):
            return {
                _clean_string_app(row[0]): row[1]
                for row in ws_hddt.iter_rows(min_row=min_r, max_row=max_r, min_col=min_c, max_col=min_c + 1, values_only=True)
                if row[0] and row[1] is not None
            }

        phi_bvmt_map_raw = get_lookup_hddt_config(10, 13)
        phi_bvmt_map_hddt = {_clean_string_app(k): _to_float_app(v) for k, v in phi_bvmt_map_raw.items()}

        static_data['hddt_config'] = {
            "DS_CHXD": chxd_list_for_hddt,
            "tk_mk": tk_mk_map_hddt,
            "khhd_map": khhd_map_hddt,
            "chxd_to_khuvuc_map": chxd_to_khuvuc_map_hddt,
            "vu_viec_map": vu_viec_map_hddt,
            "phi_bvmt_map": phi_bvmt_map_hddt,
            "tk_no_map": get_lookup_hddt_config(29, 31),
            "tk_doanh_thu_map": get_lookup_hddt_config(33, 35),
            "tk_thue_co_map": get_lookup_hddt_config(38, 40),
            "tk_gia_von_value": ws_hddt['B36'].value,
            "tk_no_bvmt_map": get_lookup_hddt_config(44, 46),
            "tk_dt_thue_bvmt_map": get_lookup_hddt_config(48, 50),
            "tk_gia_von_bvmt_value": ws_hddt['B51'].value,
            "tk_thue_co_bvmt_map": get_lookup_hddt_config(53, 55)
        }
        wb_hddt.close()

        # --- START: LOGIC CẬP NHẬT ---
        # Load MaHH.xlsx và đọc thêm cột "Loại hàng hóa"
        wb_mahh = load_workbook("MaHH.xlsx", data_only=True)
        ws_mahh = wb_mahh.active

        ma_hang_map = {}
        petroleum_products_list = []  # Danh sách để lưu các mặt hàng xăng dầu
        # Duyệt file MaHH.xlsx, đọc 4 cột (A, B, C, D)
        for r in ws_mahh.iter_rows(min_row=2, max_col=4, values_only=True):
            ten_hang = _clean_string_app(r[0])
            ma_hang = _clean_string_app(r[2])
            loai_hang = _clean_string_app(r[3])  # Đọc cột D - Loại hàng hóa

            if ten_hang and ma_hang:
                ma_hang_map[ten_hang] = ma_hang

            # Nếu loại hàng là "Xăng dầu", thêm vào danh sách
            if ten_hang and loai_hang.lower() == 'xăng dầu':
                petroleum_products_list.append(ten_hang)

        # Lưu cả hai map/list vào dictionary cấu hình
        static_data['hddt_config']["ma_hang_map"] = ma_hang_map
        static_data['hddt_config']["petroleum_products"] = petroleum_products_list
        # THÊM MỚI: Thêm danh sách mặt hàng xăng dầu cho POS handler
        static_data['pos_config']["petroleum_products"] = petroleum_products_list
        wb_mahh.close()
        # --- END: LOGIC CẬP NHẬT ---

        # Load DSKH.xlsx
        wb_dskh = load_workbook("DSKH.xlsx", data_only=True)
        static_data['hddt_config']["mst_to_makh_map"] = {
            _clean_string_app(r[2]): _clean_string_app(r[3])
            for r in wb_dskh.active.iter_rows(min_row=2, max_col=4, values_only=True)
            if r[2]
        }
        wb_dskh.close()

        # Load ChietKhau.xlsx for discount data
        try:
            with open("ChietKhau.xlsx", "rb") as f:
                discount_file_bytes = f.read()
            static_data['discount_data'] = _load_discount_data(discount_file_bytes)
        except FileNotFoundError:
            print("Cảnh báo: Không tìm thấy file 'ChietKhau.xlsx'. Chức năng chiết khấu sẽ không hoạt động.")
            static_data['discount_data'] = defaultdict(dict)
        except Exception as e:
            print(f"Lỗi khi tải file 'ChietKhau.xlsx': {e}. Chức năng chiết khấu có thể bị ảnh hưởng.")
            static_data['discount_data'] = defaultdict(dict)

        return static_data, None
    except FileNotFoundError as e:
        return None, f"Lỗi: Không tìm thấy file cấu hình. Chi tiết: {e.filename}"
    except Exception as e:
        return None, f"Lỗi khi đọc file cấu hình: {e}"

# Load static data once when the app starts
_global_static_config_data, _static_config_error = load_all_static_config_data()
if _static_config_error:
    print(f"Error loading static configuration data: {_static_config_error}")

def get_chxd_list():
    """
    Đọc danh sách CHXD và ký hiệu hóa đơn tương ứng từ dữ liệu cấu hình đã tải.
    """
    if _static_config_error:
        flash(_static_config_error, "danger")
        return []

    chxd_data = []
    for chxd_name, details in _global_static_config_data['pos_config']['chxd_detail_map'].items():
        chxd_data.append({
            'name': chxd_name,
            'symbol': details['f5_val_full']
        })
    chxd_data.sort(key=lambda x: x['name'])
    return chxd_data

@app.route('/', methods=['GET'])
def index():
    """Hiển thị trang upload chính."""
    chxd_list = get_chxd_list()
    active_tab = request.args.get('active_tab', 'upsse')
    return render_template('index.html', chxd_list=chxd_list, form_data={"active_tab": active_tab}, date_ambiguous=False)

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
        if _static_config_error:
            raise ValueError(_static_config_error)

        if not form_data["selected_chxd"]:
            flash('Vui lòng chọn CHXD.', 'warning')
            return redirect(url_for('index', active_tab='upsse'))

        file_content = None
        if form_data["encoded_file"]:
            file_content = base64.b64decode(form_data["encoded_file"])
        elif 'file' in request.files and request.files['file'].filename != '':
            file_content = request.files['file'].read()
        else:
            flash('Vui lòng tải lên file Bảng kê.', 'warning')
            return redirect(url_for('index', active_tab='upsse'))

        # Nhận diện loại bảng kê (POS/HDDT)
        report_type = detect_report_type(file_content)

        selected_chxd_symbol = None
        for chxd_info in chxd_list:
            if chxd_info['name'] == form_data["selected_chxd"]:
                selected_chxd_symbol = chxd_info['symbol']
                break

        if not selected_chxd_symbol:
            flash(f"Không tìm thấy ký hiệu hóa đơn cho CHXD '{form_data['selected_chxd']}'. Vui lòng kiểm tra file cấu hình Data_HDDT.xlsx.", 'danger')
            return redirect(url_for('index', active_tab='upsse'))

        # Gọi handler tạo dữ liệu như cũ (KHÔNG ĐỔI THUẬT TOÁN)
        if report_type == 'POS':
            result = process_pos_report(
                file_content_bytes=file_content,
                selected_chxd=form_data["selected_chxd"],
                price_periods=form_data["price_periods"],
                new_price_invoice_number=form_data["invoice_number"],
                static_data_pos=_global_static_config_data['pos_config'],
                selected_chxd_symbol=selected_chxd_symbol
            )
        elif report_type == 'HDDT':
            result = process_hddt_report(
                file_content_bytes=file_content,
                selected_chxd=form_data["selected_chxd"],
                price_periods=form_data["price_periods"],
                new_price_invoice_number=form_data["invoice_number"],
                confirmed_date_str=form_data["confirmed_date"],
                static_data_hddt=_global_static_config_data['hddt_config'],
                selected_chxd_symbol=selected_chxd_symbol
            )
        else:
            raise ValueError("Không thể tự động nhận diện loại Bảng kê. Vui lòng kiểm tra lại file Excel bạn đã tải lên.")

        # Nhánh cần người dùng xác nhận ngày (giữ nguyên)
        if isinstance(result, dict) and result.get('choice_needed'):
            form_data["encoded_file"] = base64.b64encode(file_content).decode('utf-8')
            return render_template('index.html', chxd_list=chxd_list, date_ambiguous=True, date_options=result['options'], form_data=form_data)

        # Tạo phần tên file mới dựa theo cửa hàng + ngày
        report_date = _extract_report_date_for_filename(file_content, report_type, form_data["confirmed_date"])
        base_filename = _make_base_filename(form_data["selected_chxd"], report_date)

        # Hai giai đoạn giá (trả về dict với 'old' và/hoặc 'new')
        if isinstance(result, dict) and ('old' in result or 'new' in result):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                if result.get('old'):
                    result['old'].seek(0)
                    zipf.writestr(f'{base_filename}_GiaCu.xlsx', result['old'].read())
                if result.get('new'):
                    result['new'].seek(0)
                    zipf.writestr(f'{base_filename}_GiaMoi.xlsx', result['new'].read())
            zip_buffer.seek(0)
            flash('Xử lý Đồng bộ SSE thành công!', 'success')
            # Giữ tên zip ngoài như cũ để tránh ảnh hưởng automation phía người dùng (chỉ đổi tên file bên trong)
            return send_file(zip_buffer, as_attachment=True, download_name='UpSSE_2_giai_doan.zip', mimetype='application/zip')

        # Một giai đoạn giá: trả về 1 file .xlsx
        elif isinstance(result, io.BytesIO):
            result.seek(0)
            flash('Xử lý Đồng bộ SSE thành công!', 'success')
            return send_file(
                result,
                as_attachment=True,
                download_name=f'{base_filename}.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            raise ValueError("Hàm xử lý không trả về kết quả hợp lệ.")

    except ValueError as ve:
        flash(str(ve).replace('\n', '<br>'), 'danger')
        return render_template('index.html', chxd_list=chxd_list, form_data=form_data, date_ambiguous=False)
    except Exception as e:
        flash(f"Đã xảy ra lỗi không mong muốn: {e}", 'danger')
        return render_template('index.html', chxd_list=chxd_list, form_data=form_data, date_ambiguous=False)

@app.route('/reconcile', methods=['POST'])
def reconcile():
    """Xử lý file tải lên cho chức năng Đối soát."""
    chxd_list_data = get_chxd_list()
    reconciliation_data = None
    try:
        if _static_config_error:
            raise ValueError(_static_config_error)

        selected_chxd_name = request.form.get('chxd')
        file_log_bom = request.files.get('file_log_bom')
        file_hddt = request.files.get('file_hddt')

        if not selected_chxd_name or not file_log_bom or not file_hddt:
            flash('Vui lòng chọn CHXD và tải lên đủ cả 2 file để đối soát.', 'warning')
            return redirect(url_for('index', active_tab='doisoat'))

        selected_chxd_symbol = None
        for chxd_info in chxd_list_data:
            if chxd_info['name'] == selected_chxd_name:
                selected_chxd_symbol = chxd_info['symbol']
                break

        if not selected_chxd_symbol:
            flash(f"Không tìm thấy ký hiệu hóa đơn cho CHXD '{selected_chxd_name}'.", 'danger')
            return redirect(url_for('index', active_tab='doisoat'))

        log_bom_bytes = file_log_bom.read()
        hddt_bytes = file_hddt.read()

        discount_data = _global_static_config_data.get('discount_data', defaultdict(dict))

        reconciliation_data = perform_reconciliation(
            log_bom_bytes,
            hddt_bytes,
            selected_chxd_name,
            selected_chxd_symbol,
            discount_data
        )

        if reconciliation_data:
            reconciliation_data['selected_chxd_name'] = selected_chxd_name
            flash('Đối soát thành công!', 'success')
        else:
            flash('Không có dữ liệu trả về từ chức năng đối soát.', 'warning')

    except Exception as e:
        flash(f"Lỗi trong quá trình đối soát: {e}", 'danger')

    return render_template('index.html',
                           chxd_list=chxd_list_data,
                           reconciliation_data=reconciliation_data,
                           form_data={"active_tab": "doisoat"},
                           date_ambiguous=False)

@app.route('/generate_discount_report', methods=['POST'])
def generate_discount_report():
    try:
        reconciliation_data_json = request.json
        if not reconciliation_data_json:
            raise ValueError("Không nhận được dữ liệu đối soát để tạo báo cáo.")

        discount_data = _global_static_config_data.get('discount_data', defaultdict(dict))
        excel_buffer = _generate_discount_report_excel(reconciliation_data_json, discount_data)

        if excel_buffer:
            excel_buffer.seek(0)
            return send_file(
                excel_buffer,
                as_attachment=True,
                download_name='BaoCaoChietKhau.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            raise ValueError("Không thể tạo báo cáo chiết khấu.")

    except ValueError as ve:
        flash(str(ve).replace('\n', '<br>'), 'danger')
        return jsonify({"status": "error", "message": str(ve)}), 400
    except Exception as e:
        flash(f"Đã xảy ra lỗi không mong muốn khi tạo báo cáo chiết khấu: {e}", 'danger')
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/process_stock_card', methods=['POST'])
def process_stock_card():
    """Xử lý file ảnh/PDF tải lên cho chức năng Thẻ kho tự động."""
    chxd_list = get_chxd_list()
    selected_chxd = request.form.get('chxd_thekho')

    try:
        if _static_config_error:
            raise ValueError(_static_config_error)

        if not selected_chxd:
            flash('Vui lòng chọn Cửa Hàng Xăng Dầu (CHXD) cho chức năng Thẻ kho.', 'warning')
            return render_template('index.html', chxd_list=chxd_list, form_data={"active_tab": "thekho"}, date_ambiguous=False)

        uploaded_files = request.files.getlist('files[]')
        if not uploaded_files or all(f.filename == '' for f in uploaded_files):
            flash('Vui lòng tải lên ít nhất một file ảnh hoặc PDF.', 'warning')
            return render_template('index.html', chxd_list=chxd_list, form_data={"active_tab": "thekho"}, date_ambiguous=False)

        excel_buffer = process_stock_card_data(uploaded_files, selected_chxd)

        if excel_buffer:
            excel_buffer.seek(0)
            flash('Xử lý Thẻ kho tự động thành công!', 'success')
            return send_file(
                excel_buffer,
                as_attachment=True,
                download_name='TheKho_TuDong.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            flash('Không có dữ liệu hợp lệ được trích xuất từ các file đã tải lên.', 'warning')
            return render_template('index.html', chxd_list=chxd_list, form_data={"active_tab": "thekho"}, date_ambiguous=False)

    except ValueError as ve:
        flash(str(ve).replace('\n', '<br>'), 'danger')
        return render_template('index.html', chxd_list=chxd_list, form_data={"active_tab": "thekho"}, date_ambiguous=False)
    except Exception as e:
        flash(f"Đã xảy ra lỗi không mong muốn trong quá trình xử lý Thẻ kho: {e}", 'danger')
        return render_template('index.html', chxd_list=chxd_list, form_data={"active_tab": "thekho"}, date_ambiguous=False)

@app.route('/clear_flash_messages', methods=['GET'])
def clear_flash_messages():
    """Route này được gọi bởi JavaScript để xóa các thông báo flash trong session."""
    _ = get_flashed_messages()
    return '', 204

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
