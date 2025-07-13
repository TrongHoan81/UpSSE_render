import io
import re
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook, Workbook

# --- Các hàm tiện ích nội bộ ---
def _clean_string_hddt(s):
    """Làm sạch chuỗi, loại bỏ khoảng trắng thừa và ký tự '."""
    if s is None: return ""
    cleaned_s = str(s).strip()
    if cleaned_s.startswith("'"): cleaned_s = cleaned_s[1:]
    return re.sub(r'\s+', ' ', cleaned_s)

def _to_float_hddt(value):
    """Chuyển đổi giá trị sang float, xử lý các trường hợp lỗi."""
    if value is None: return 0.0
    try:
        return float(str(value).replace(',', '').strip())
    except (ValueError, TypeError): return 0.0

def _format_tax_code_hddt(raw_vat_value):
    """Định dạng mã thuế."""
    if raw_vat_value is None: return ""
    try:
        s_value = str(raw_vat_value).replace('%', '').strip()
        f_value = float(s_value)
        if 0 < f_value < 1: f_value *= 100
        return f"{round(f_value):02d}"
    except (ValueError, TypeError): return ""
    
def _create_upsse_workbook_hddt():
    """Tạo một workbook Excel mới với các header chuẩn cho UpSSE."""
    headers = ["Mã khách", "Tên khách hàng", "Ngày", "Số hóa đơn", "Ký hiệu", "Diễn giải", "Mã hàng", "Tên mặt hàng", "Đvt", "Mã kho", "Mã vị trí", "Mã lô", "Số lượng", "Giá bán", "Tiền hàng", "Mã nt", "Tỷ giá", "Mã thuế", "Tk nợ", "Tk doanh thu", "Tk giá vốn", "Tk thuế có", "Cục thuế", "Vụ việc", "Bộ phận", "Lsx", "Sản phẩm", "Hợp đồng", "Phí", "Khế ước", "Nhân viên bán", "Tên KH(thuế)", "Địa chỉ (thuế)", "Mã số Thuế", "Nhóm Hàng", "Ghi chú", "Tiền thuế"]
    wb = Workbook()
    ws = wb.active
    for _ in range(4): ws.append([''] * len(headers))
    ws.append(headers)
    return wb

# --- Hàm tạo dòng BVMT (chỉ dùng cho hóa đơn riêng lẻ) ---
def _create_hddt_bvmt_row(original_row, phi_bvmt, static_data_hddt, khu_vuc):
    """Tạo dòng Thuế Bảo vệ Môi trường (BVMT) cho hóa đơn riêng lẻ."""
    bvmt_row = list(original_row)
    so_luong = _to_float_hddt(original_row[12])
    thue_suat = _to_float_hddt(original_row[17]) / 100.0 if original_row[17] else 0.0
    
    tien_hang_dong_bvmt = round(phi_bvmt * so_luong)
    tien_thue_dong_bvmt = round(tien_hang_dong_bvmt * thue_suat)

    bvmt_row[6], bvmt_row[7] = "TMT", "Thuế bảo vệ môi trường"
    bvmt_row[13] = phi_bvmt
    bvmt_row[14] = tien_hang_dong_bvmt
    bvmt_row[36] = tien_thue_dong_bvmt
    
    bvmt_row[18] = static_data_hddt.get('tk_no_bvmt_map', {}).get(khu_vuc)
    bvmt_row[19] = static_data_hddt.get('tk_dt_thue_bvmt_map', {}).get(khu_vuc)
    bvmt_row[20] = static_data_hddt.get('tk_gia_von_bvmt_value')
    bvmt_row[21] = static_data_hddt.get('tk_thue_co_bvmt_map', {}).get(khu_vuc)
    
    for i in [5, 31, 32, 33]: bvmt_row[i] = ''
    return bvmt_row

def _parse_date_from_excel_cell(date_val):
    """
    Attempts to parse a date value from an Excel cell, handling various formats.
    Returns a datetime.date object if successful, None otherwise.
    """
    if date_val is None:
        return None

    # 1. Already a datetime object (openpyxl's native parsing)
    if isinstance(date_val, datetime):
        return date_val.date()

    # 2. Excel serial number
    if isinstance(date_val, (int, float)):
        try:
            # Excel's epoch is 1899-12-30 for Windows, 1904-01-01 for Mac.
            # pandas defaults to 1899-12-30, which is common.
            # Handle potential float issues (e.g., 45998.0)
            return pd.to_datetime(float(date_val), unit='D', origin='1899-12-30').date()
        except (ValueError, TypeError):
            pass # Fall through to string parsing

    # 3. String formats
    if isinstance(date_val, str):
        date_str = date_val.strip()
        # Common date formats to try
        date_formats = [
            '%d/%m/%Y', # 13/07/2025
            '%d-%m-%Y', # 13-07-2025
            '%Y/%m/%d', # 2025/07/13
            '%Y-%m-%d', # 2025-07-13
            '%d/%m/%y', # 13/07/25
            '%d-%m-%y', # 13-07-25
        ]
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str, fmt).date()
            except ValueError:
                continue # Try next format
    
    return None # If no format matches

# --- Hàm xử lý chính ---
def _generate_upsse_from_hddt_rows(rows_to_process, static_data_hddt, selected_chxd, final_date, summary_suffix_map):
    """Tạo các dòng dữ liệu cho file UpSSE từ dữ liệu bảng kê HĐĐT."""
    upsse_wb = _create_upsse_workbook_hddt()
    ws = upsse_wb.active

    # Nếu không có dòng nào để xử lý, trả về một workbook rỗng với các header
    if not rows_to_process:
        print(f"DEBUG: Không có dòng nào để xử lý trong giai đoạn này. Trả về workbook rỗng.")
        output_buffer = io.BytesIO()
        upsse_wb.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer

    khu_vuc, ma_kho = static_data_hddt['chxd_to_khuvuc_map'].get(selected_chxd), static_data_hddt['tk_mk'].get(selected_chxd)
    tk_no, tk_doanh_thu, tk_gia_von, tk_thue_co = static_data_hddt['tk_no_map'].get(khu_vuc), static_data_hddt['tk_doanh_thu_map'].get(khu_vuc), static_data_hddt['tk_gia_von_value'], static_data_hddt['tk_thue_co_map'].get(khu_vuc)
    original_invoice_rows, bvmt_rows, summary_data = [], [], {}
    first_invoice_prefix_source = ""
    
    processed_row_count = 0 # Thêm biến đếm số dòng được xử lý
    for bkhd_row in rows_to_process:
        if _to_float_hddt(bkhd_row[8] if len(bkhd_row) > 8 else None) <= 0: 
            # print(f"DEBUG: Bỏ qua dòng có số lượng <= 0: {bkhd_row}") # Có thể bỏ comment để debug chi tiết hơn
            continue
        
        processed_row_count += 1 # Tăng biến đếm khi một dòng được xử lý
        ten_kh, ten_mat_hang = _clean_string_hddt(bkhd_row[3]), _clean_string_hddt(bkhd_row[6])
        is_anonymous, is_petrol = ("không lấy hóa đơn" in ten_kh.lower()), (ten_mat_hang in static_data_hddt['phi_bvmt_map'])
        
        # Xử lý hóa đơn riêng lẻ (KHÔNG THAY ĐỔI)
        if not is_anonymous or not is_petrol:
            new_upsse_row = [''] * 37
            new_upsse_row[9], new_upsse_row[1], new_upsse_row[31], new_upsse_row[2] = ma_kho, ten_kh, ten_kh, final_date
            so_hd_goc = str(bkhd_row[19] or '').strip()
            new_upsse_row[3] = f"HN{so_hd_goc[-6:]}" if selected_chxd == "Nguyễn Huệ" else f"{(str(bkhd_row[18] or '').strip())[-2:]}{so_hd_goc[-6:]}"
            new_upsse_row[4] = _clean_string_hddt(bkhd_row[17]) + _clean_string_hddt(bkhd_row[18])
            new_upsse_row[5], new_upsse_row[7], new_upsse_row[6] = f"Xuất bán hàng theo hóa đơn số {new_upsse_row[3]}", ten_mat_hang, static_data_hddt['ma_hang_map'].get(ten_mat_hang, '')
            new_upsse_row[8], new_upsse_row[12] = _clean_string_hddt(bkhd_row[10]), round(_to_float_hddt(bkhd_row[8]), 3)
            phi_bvmt = static_data_hddt['phi_bvmt_map'].get(ten_mat_hang, 0.0) if is_petrol else 0.0
            new_upsse_row[13] = _to_float_hddt(bkhd_row[9]) - phi_bvmt
            ma_thue = _format_tax_code_hddt(bkhd_row[14])
            new_upsse_row[17] = ma_thue
            thue_suat = _to_float_hddt(ma_thue) / 100.0 if ma_thue else 0.0
            tien_thue_goc, so_luong = _to_float_hddt(bkhd_row[15]), _to_float_hddt(bkhd_row[8])
            tien_thue_phi_bvmt = round(phi_bvmt * so_luong * thue_suat)
            new_upsse_row[36] = round(tien_thue_goc - tien_thue_phi_bvmt)
            new_upsse_row[14] = round(_to_float_hddt(bkhd_row[13]) if not is_petrol else _to_float_hddt(bkhd_row[16]) - tien_thue_goc - round(phi_bvmt * so_luong))
            new_upsse_row[18], new_upsse_row[19], new_upsse_row[20], new_upsse_row[21] = tk_no, tk_doanh_thu, tk_gia_von, tk_thue_co
            chxd_vu_viec_map = static_data_hddt['vu_viec_map'].get(selected_chxd, {})
            new_upsse_row[23] = chxd_vu_viec_map.get(ten_mat_hang, chxd_vu_viec_map.get("Dầu mỡ nhờn", ''))
            new_upsse_row[32], mst_khach_hang = _clean_string_hddt(bkhd_row[4]), _clean_string_hddt(bkhd_row[5])
            new_upsse_row[33] = mst_khach_hang
            ma_kh_fast = _clean_string_hddt(bkhd_row[2])
            new_upsse_row[0] = ma_kh_fast if ma_kh_fast and len(ma_kh_fast) < 12 else static_data_hddt['mst_to_makh_map'].get(mst_khach_hang, ma_kho)
            original_invoice_rows.append(new_upsse_row)
            if is_petrol: bvmt_rows.append(_create_hddt_bvmt_row(new_upsse_row, phi_bvmt, static_data_hddt, khu_vuc))
        
        # Gom dữ liệu khách vãng lai (KHÔNG THAY ĐỔI)
        else:
            if not first_invoice_prefix_source: first_invoice_prefix_source = str(bkhd_row[18] or '').strip()
            if ten_mat_hang not in summary_data:
                summary_data[ten_mat_hang] = {'sl': 0, 'thue': 0, 'phai_thu': 0, 'first_data': {'mau_so': _clean_string_hddt(bkhd_row[17]),'ky_hieu': _clean_string_hddt(bkhd_row[18]),'don_gia': _to_float_hddt(bkhd_row[9]),'vat_raw': bkhd_row[14]}}
            summary_data[ten_mat_hang]['sl'] += _to_float_hddt(bkhd_row[8])
            summary_data[ten_mat_hang]['thue'] += _to_float_hddt(bkhd_row[15])
            summary_data[ten_mat_hang]['phai_thu'] += _to_float_hddt(bkhd_row[16])
    
    # --- Tạo các dòng tổng hợp cho khách vãng lai ---
    prefix = first_invoice_prefix_source[-2:] if len(first_invoice_prefix_source) >= 2 else first_invoice_prefix_source
    for product, data in summary_data.items():
        summary_row = [''] * 37
        first_data = data['first_data']
        
        # --- START: LOGIC TÍNH TOÁN "PHÂN BỔ TỪ TỔNG" ---
        total_phai_thu = data['phai_thu']
        total_tien_thue_gtgt = data['thue']
        total_so_luong = data['sl']
        phi_bvmt_unit = static_data_hddt['phi_bvmt_map'].get(product, 0.0)
        ma_thue_str = _format_tax_code_hddt(first_data['vat_raw'])
        thue_suat = _to_float_hddt(ma_thue_str) / 100.0 if ma_thue_str else 0.0

        tien_hang_dong_bvmt = round(phi_bvmt_unit * total_so_luong)
        tien_thue_dong_bvmt = round(tien_hang_dong_bvmt * thue_suat)
        tien_thue_dong_goc = total_tien_thue_gtgt - tien_thue_dong_bvmt
        tien_hang_dong_goc = total_phai_thu - tien_hang_dong_bvmt - tien_thue_dong_bvmt - tien_thue_dong_goc
        # --- END: LOGIC TÍNH TOÁN "PHÂN BỔ TỪ TỔNG" ---
        
        # Điền dữ liệu cho dòng gốc
        summary_row[0], summary_row[1] = ma_kho, f"Khách hàng mua {product} không lấy hóa đơn"
        summary_row[31], summary_row[2] = summary_row[1], final_date
        summary_row[3] = f"{prefix}BK.{final_date.strftime('%d.%m')}.{summary_suffix_map.get(product, '')}"
        summary_row[4] = first_data['mau_so'] + first_data['ky_hieu']
        summary_row[5] = f"Xuất bán hàng theo hóa đơn số {summary_row[3]}"
        summary_row[7], summary_row[6], summary_row[8], summary_row[9] = product, static_data_hddt['ma_hang_map'].get(product, ''), "Lít", ma_kho
        summary_row[12] = round(total_so_luong, 3)
        summary_row[13] = first_data['don_gia'] - phi_bvmt_unit
        summary_row[17] = ma_thue_str
        summary_row[14] = tien_hang_dong_goc
        summary_row[36] = tien_thue_dong_goc
        summary_row[18], summary_row[19], summary_row[20], summary_row[21] = tk_no, tk_doanh_thu, tk_gia_von, tk_thue_co
        summary_row[23] = static_data_hddt['vu_viec_map'].get(selected_chxd, {}).get(product, '')
        original_invoice_rows.append(summary_row)
        
        # --- START: TẠO DÒNG BVMT THỦ CÔNG ĐỂ BẢO TOÀN GIÁ TRỊ ---
        bvmt_summary_row = list(summary_row)
        bvmt_summary_row[6], bvmt_summary_row[7] = "TMT", "Thuế bảo vệ môi trường"
        bvmt_summary_row[13] = phi_bvmt_unit
        bvmt_summary_row[18] = static_data_hddt.get('tk_no_bvmt_map', {}).get(khu_vuc)
        bvmt_summary_row[19] = static_data_hddt.get('tk_dt_thue_bvmt_map', {}).get(khu_vuc)
        bvmt_summary_row[20] = static_data_hddt.get('tk_gia_von_bvmt_value')
        bvmt_summary_row[21] = static_data_hddt.get('tk_thue_co_bvmt_map', {}).get(khu_vuc)
        # Gán chính xác các giá trị đã được phân bổ
        bvmt_summary_row[14] = tien_hang_dong_bvmt
        bvmt_summary_row[36] = tien_thue_dong_bvmt
        # Xóa các trường không cần thiết
        for i in [5, 31, 32, 33]: bvmt_summary_row[i] = ''
        bvmt_rows.append(bvmt_summary_row)
        # --- END: TẠO DÒNG BVMT THỦ CÔNG ---
    
    # --- Ghi ra file Excel ---
    for row_data in original_invoice_rows + bvmt_rows:
        ws.append(row_data)

    print(f"DEBUG: Số dòng hóa đơn gốc được thêm vào workbook: {len(original_invoice_rows)}")
    print(f"DEBUG: Số dòng BVMT được thêm vào workbook: {len(bvmt_rows)}")
    print(f"DEBUG: Tổng số dòng (sau khi lọc số lượng <= 0) được xử lý trong giai đoạn này: {processed_row_count}")

    for row_index in range(6, ws.max_row + 1):
        date_cell = ws[f'C{row_index}']
        if isinstance(date_cell.value, datetime):
            date_cell.number_format = 'dd/mm/yyyy'
        text_cell = ws[f'R{row_index}']
        text_cell.number_format = '@'
        
    output_buffer = io.BytesIO()
    upsse_wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer

# --- Khối lệnh điều phối chính ---
def process_hddt_report(file_content_bytes, selected_chxd, price_periods, new_price_invoice_number, confirmed_date_str=None, static_data_hddt=None, selected_chxd_symbol=None):
    """
    Xử lý bảng kê HĐĐT để tạo file UpSSE.
    Bao gồm xác thực CHXD dựa trên ký hiệu hóa đơn trong bảng kê HĐĐT.
    """
    # static_data_hddt và selected_chxd_symbol đã được truyền từ app.py
    if static_data_hddt is None:
        raise ValueError("Dữ liệu cấu hình tĩnh cho HDDT chưa được tải. Vui lòng kiểm tra cấu hình ứng dụng.")
    if selected_chxd_symbol is None:
        raise ValueError("Ký hiệu hóa đơn của CHXD chưa được cung cấp để xác thực.")

    bkhd_wb = load_workbook(io.BytesIO(file_content_bytes), data_only=True)
    bkhd_ws = bkhd_wb.active

    # --- BƯỚC XÁC THỰC KÝ HIỆU HÓA ĐƠN TỪ FILE HĐĐT (CỘT S, DÒNG 11 TRỞ ĐI) ---
    # Lấy 6 ký tự cuối của ký hiệu hóa đơn từ file cấu hình Data_HDDT.xlsx
    if len(selected_chxd_symbol) < 6:
        raise ValueError(f"Ký hiệu hóa đơn trong file cấu hình Data_HDDT.xlsx ('{selected_chxd_symbol}') quá ngắn để xác thực.")
    expected_invoice_symbol_suffix = selected_chxd_symbol[-6:].upper()

    has_at_least_one_valid_invoice_for_symbol_check = False
    
    # Duyệt qua cột S (index 18) từ dòng 11 để kiểm tra ký hiệu hóa đơn
    # Kiểm tra các dòng có số lượng > 0 để xác định đây là dòng hóa đơn thực tế
    # Giới hạn số dòng kiểm tra để tránh đọc toàn bộ file lớn không cần thiết
    max_rows_to_check = min(bkhd_ws.max_row, 100) # Check up to 100 rows or end of sheet
    for row_index, row_values in enumerate(bkhd_ws.iter_rows(min_row=11, max_row=max_rows_to_check, values_only=True), start=11):
        # Cột I (index 8) là số lượng
        quantity_val = _to_float_hddt(row_values[8] if len(row_values) > 8 else None)
        
        # Nếu số lượng <= 0, đây có thể là dòng tiêu đề, chân trang, hoặc dòng tổng cộng. Bỏ qua xác thực ký hiệu cho các dòng này.
        if quantity_val <= 0:
            continue

        # Nếu đến đây, đây là một dòng hóa đơn hợp lệ (có số lượng > 0)
        has_at_least_one_valid_invoice_for_symbol_check = True

        # Thực hiện xác thực ký hiệu hóa đơn
        if len(row_values) > 18 and row_values[18] is not None: # Cột S (index 18)
            actual_invoice_symbol_hddt = _clean_string_hddt(row_values[18])
            if len(actual_invoice_symbol_hddt) >= 6:
                if actual_invoice_symbol_hddt[-6:].upper() != expected_invoice_symbol_suffix:
                    # ĐÃ THAY ĐỔI: Bỏ phần ký hiệu mong muốn
                    raise ValueError("Bảng kê HĐĐT không phải của cửa hàng bạn chọn hoặc không tìm thấy ký hiệu hóa đơn hợp lệ.")
            else:
                # Dòng hóa đơn hợp lệ nhưng ký hiệu quá ngắn
                raise ValueError(f"Ký hiệu hóa đơn tại dòng {row_index} của bảng kê HDDT quá ngắn để xác thực.")
        else:
            # Dòng hóa đơn hợp lệ nhưng thiếu ký hiệu hóa đơn
            raise ValueError(f"Hóa đơn tại dòng {row_index} của bảng kê HDDT thiếu ký hiệu hóa đơn (cột S).")
    
    # Sau khi kiểm tra tất cả các dòng, nếu không tìm thấy bất kỳ dòng hóa đơn hợp lệ nào để xác thực ký hiệu.
    if not has_at_least_one_valid_invoice_for_symbol_check:
        raise ValueError("Không tìm thấy hóa đơn hợp lệ nào trong file Bảng kê HDDT để xác thực ký hiệu.")

    # --- Tiếp tục với logic phát hiện ngày tháng hiện có ---
    final_date = None
    if confirmed_date_str:
        final_date = datetime.strptime(confirmed_date_str, '%Y-%m-%d')
    else:
        unique_dates = set()
        for row in bkhd_ws.iter_rows(min_row=11, values_only=True):
            # Chỉ xử lý các dòng có số lượng (cột I, index 8) lớn hơn 0
            quantity_val = _to_float_hddt(row[8] if len(row) > 8 else None)
            if quantity_val > 0:
                date_val_from_cell = row[20] if len(row) > 20 else None # Cột U (index 20)
                parsed_date = _parse_date_from_excel_cell(date_val_from_cell)
                if parsed_date:
                    unique_dates.add(parsed_date)
                else:
                    # Ghi log cảnh báo nếu không phân tích được ngày tháng từ một dòng hợp lệ
                    print(f"WARNING: Could not parse date '{date_val_from_cell}' from row with quantity > 0. Row data: {row}")
        
        # Nếu sau khi duyệt tất cả các dòng, tập hợp unique_dates vẫn rỗng
        if not unique_dates:
            raise ValueError("Không tìm thấy dữ liệu hóa đơn hợp lệ nào trong file Bảng kê HDDT.")
        
        # Nếu tìm thấy nhiều hơn một ngày duy nhất, yêu cầu xác nhận hoặc báo lỗi
        if len(unique_dates) > 1:
            # Sắp xếp các ngày để đảm bảo thứ tự hiển thị nhất quán
            sorted_dates = sorted(list(unique_dates))
            options = []
            for d in sorted_dates:
                # Tạo 2 tùy chọn: một là ngày/tháng/năm, một là tháng/ngày/năm
                # Chỉ hiển thị 2 tùy chọn nếu chúng khác nhau
                date1_str = d.strftime('%d/%m/%Y')
                date2_str = d.strftime('%m/%d/%Y') # Format for potential ambiguity
                
                options.append({'text': date1_str, 'value': d.strftime('%Y-%m-%d')})
                if date1_str != date2_str: # Only add if it's genuinely ambiguous
                     # Tạo một đối tượng datetime mới để đảm bảo tháng và ngày được hoán đổi nếu cần
                    ambiguous_date_obj = datetime(d.year, d.day, d.month)
                    options.append({'text': ambiguous_date_obj.strftime('%d/%m/%Y'), 'value': ambiguous_date_obj.strftime('%Y-%m-%d')})
            
            # Loại bỏ các tùy chọn trùng lặp và sắp xếp lại
            unique_options = []
            seen_values = set()
            for opt in options:
                if opt['value'] not in seen_values:
                    unique_options.append(opt)
                    seen_values.add(opt['value'])
            
            # Sắp xếp lại các tùy chọn theo ngày tháng
            unique_options.sort(key=lambda x: datetime.strptime(x['value'], '%Y-%m-%d'))

            return {'choice_needed': True, 'options': unique_options}
        
        # Nếu chỉ có một ngày duy nhất
        the_date = unique_dates.pop()
        final_date = datetime(the_date.year, the_date.month, the_date.day) # Chuyển đổi về datetime object

    all_rows = list(bkhd_ws.iter_rows(min_row=11, values_only=True))
    print(f"DEBUG: Tổng số dòng đọc được từ file Excel (từ dòng 11): {len(all_rows)}")

    if price_periods == '1':
        suffix_map = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}
        print(f"DEBUG: Xử lý 1 giai đoạn giá.")
        return _generate_upsse_from_hddt_rows(all_rows, static_data_hddt, selected_chxd, final_date, suffix_map)
    else:
        print(f"DEBUG: Xử lý 2 giai đoạn giá.")
        if not new_price_invoice_number: raise ValueError("Vui lòng nhập 'Số hóa đơn đầu tiên của giá mới'.")
        split_index = -1
        for i, row in enumerate(all_rows):
            # Cột số hóa đơn là cột T (index 19)
            if str(row[19] or '').strip() == new_price_invoice_number:
                split_index = i
                break
        
        print(f"DEBUG: Số hóa đơn giá mới cần tìm: '{new_price_invoice_number}'")
        print(f"DEBUG: split_index tìm được: {split_index}")

        if split_index == -1: 
            raise ValueError(f"Không tìm thấy hóa đơn số '{new_price_invoice_number}'. Vui lòng đảm bảo số hóa đơn chính xác và có trong bảng kê.")
        
        # Lấy các dòng cho giai đoạn giá cũ (trước hoặc bằng split_index)
        # Bao gồm dòng hóa đơn giá mới trong giai đoạn giá cũ để nó được xử lý với giá cũ nếu cần
        # Hoặc chỉ lấy các dòng TRƯỚC split_index và dòng split_index là dòng đầu tiên của giá mới.
        # Dựa trên mô tả "Số hóa đơn đầu tiên của giá mới", logic hiện tại là:
        # all_rows[:split_index] -> giá cũ (không bao gồm hóa đơn giá mới)
        # all_rows[split_index:] -> giá mới (bao gồm hóa đơn giá mới)
        # Nếu muốn hóa đơn giá mới nằm trong giai đoạn giá cũ: all_rows[:split_index + 1]
        # Nếu muốn hóa đơn giá mới là ranh giới và không thuộc cả 2, hoặc thuộc giai đoạn mới: all_rows[:split_index] và all_rows[split_index:]
        # Giữ nguyên logic hiện tại: all_rows[:split_index] là cũ, all_rows[split_index:] là mới.
        # Điều này có nghĩa là hóa đơn đầu tiên của giá mới sẽ nằm trong file giá mới.

        rows_old_price = all_rows[:split_index]
        rows_new_price = all_rows[split_index:]

        print(f"DEBUG: Số dòng cho giai đoạn giá cũ (trước hóa đơn giá mới): {len(rows_old_price)}")
        print(f"DEBUG: Số dòng cho giai đoạn giá mới (từ hóa đơn giá mới trở đi): {len(rows_new_price)}")

        suffix_map_old = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}
        suffix_map_new = {"Xăng E5 RON 92-II": "5", "Xăng RON 95-III": "6", "Dầu DO 0,05S-II": "7", "Dầu DO 0,001S-V": "8"}
        
        result_old = _generate_upsse_from_hddt_rows(rows_old_price, static_data_hddt, selected_chxd, final_date, suffix_map_old)
        result_new = _generate_upsse_from_hddt_rows(rows_new_price, static_data_hddt, selected_chxd, final_date, suffix_map_new)
        
        output_dict = {}
        if result_old: 
            result_old.seek(0)
            output_dict['old'] = result_old
        if result_new: 
            result_new.seek(0)
            output_dict['new'] = result_new
        
        # Ensure both keys are present, even if the BytesIO object is for an empty workbook
        # This part ensures that even if a period has no data, an empty file is created.
        if 'old' not in output_dict:
            empty_wb_old = _create_upsse_workbook_hddt()
            empty_buffer_old = io.BytesIO()
            empty_wb_old.save(empty_buffer_old)
            empty_buffer_old.seek(0)
            output_dict['old'] = empty_buffer_old
        
        if 'new' not in output_dict:
            empty_wb_new = _create_upsse_workbook_hddt()
            empty_buffer_new = io.BytesIO()
            empty_wb_new.save(empty_buffer_new)
            empty_buffer_new.seek(0)
            output_dict['new'] = empty_buffer_new

        return output_dict
