import io
import re
from datetime import datetime
from openpyxl import load_workbook, Workbook

# --- CÁC HÀM TIỆN ÍCH ---
def _pos_to_float(value):
    """Chuyển đổi giá trị sang float, xử lý các trường hợp lỗi."""
    try:
        if isinstance(value, str):
            value = value.replace(",", "").strip()
        return float(value)
    except (ValueError, TypeError):
        return 0.0

def _pos_clean_string(s):
    """Làm sạch chuỗi, loại bỏ khoảng trắng thừa và ký tự '."""
    if s is None:
        return ""
    return re.sub(r'\s+', ' ', s).strip()

def _convert_to_datetime(date_input):
    """Chuyển đổi ngày tháng sang đối tượng datetime."""
    if isinstance(date_input, datetime):
        return date_input
    if isinstance(date_input, str):
        try:
            date_part = date_input.split(' ')[0]
            return datetime.strptime(date_part, '%d-%m-%Y')
        except (ValueError, TypeError):
            return date_input
    return date_input

# --- HÀM TẠO DÒNG TMT (CHỈ DÙNG CHO HÓA ĐƠN LẺ) ---
def _pos_create_tmt_row_for_individual(original_row, tmt_value, details):
    """Tạo dòng Thuế Bảo vệ Môi trường (TMT) cho hóa đơn riêng lẻ."""
    tmt_row = list(original_row)
    ma_thue_for_calc = _pos_to_float(original_row[17])
    tax_rate_decimal = ma_thue_for_calc / 100.0
    tmt_row[6], tmt_row[7], tmt_row[8] = "TMT", "Thuế bảo vệ môi trường", "Lít"
    tmt_row[9] = details['g5_val']
    tmt_row[13] = tmt_value
    
    # Áp dụng logic đồng bộ: round-then-calculate
    tien_hang_bvmt = round(tmt_value * _pos_to_float(original_row[12]))
    tien_thue_bvmt = round(tien_hang_bvmt * tax_rate_decimal)
    tmt_row[14] = tien_hang_bvmt
    tmt_row[36] = tien_thue_bvmt

    tmt_row[18] = details['s_lookup_table'].get(details['h5_val'], '')
    tmt_row[19] = details['t_lookup_tmt'].get(details['h5_val'], '')
    tmt_row[20], tmt_row[21] = details['u_value'], details['v_lookup_table'].get(details['h5_val'], '')
    tmt_row[31] = ""
    for idx in [5, 10, 11, 15, 16, 22, 24, 25, 26, 27, 28, 29, 30, 32, 33, 34, 35]:
        if idx < len(tmt_row): tmt_row[idx] = ''
    return tmt_row

# --- HÀM XỬ LÝ HÓA ĐƠN LẺ ---
def _pos_process_single_row(row, details, selected_chxd):
    """Xử lý một dòng dữ liệu hóa đơn riêng lẻ từ bảng kê POS."""
    upsse_row = [''] * 37
    try:
        ma_kh, ten_kh, ngay_hd_raw, so_ct, so_hd, dia_chi_goc, mst_goc, product_name, so_luong, don_gia_vat, tien_hang_source, tien_thue_source = \
        _pos_clean_string(str(row[4])), _pos_clean_string(str(row[5])), row[3], _pos_clean_string(str(row[1])), _pos_clean_string(str(row[2])), \
        _pos_clean_string(str(row[6])), _pos_clean_string(str(row[7])), _pos_clean_string(str(row[8])), _pos_to_float(row[10]), \
        _pos_to_float(row[11]), _pos_to_float(row[13]), _pos_to_float(row[14])
        ma_thue_percent = _pos_to_float(row[15]) if row[15] is not None else 8.0
    except IndexError:
        raise ValueError("Lỗi đọc cột từ file bảng kê POS. Vui lòng đảm bảo file có đủ các cột từ A đến P.")
    
    upsse_row[0] = ma_kh if ma_kh and len(ma_kh) <= 9 else details['g5_val']
    upsse_row[1] = ten_kh
    upsse_row[2] = _convert_to_datetime(ngay_hd_raw)
    if details['b5_val'] == "Nguyễn Huệ": upsse_row[3] = f"HN{so_hd[-6:]}"
    elif details['b5_val'] == "Mai Linh": upsse_row[3] = f"MM{so_hd[-6:]}"
    else: upsse_row[3] = f"{so_ct[-2:]}{so_hd[-6:]}"
    upsse_row[4] = f"1{so_ct}" if so_ct else ''
    upsse_row[5] = f"Xuất bán lẻ theo hóa đơn số {upsse_row[3]}"
    upsse_row[6] = details['lookup_table'].get(product_name.lower(), '')
    upsse_row[7], upsse_row[8] = product_name, "Lít"
    upsse_row[9] = details['g5_val']
    upsse_row[12] = so_luong
    tmt_value = details['tmt_lookup_table'].get(product_name.lower(), 0.0)
    tax_rate_decimal = ma_thue_percent / 100.0
    upsse_row[13] = round(don_gia_vat / (1 + tax_rate_decimal) - tmt_value, 2)
    
    # Áp dụng logic đồng bộ: round-then-calculate
    tien_hang_bvmt_le = round(tmt_value * so_luong)
    tien_thue_bvmt_le = round(tien_hang_bvmt_le * tax_rate_decimal)
    upsse_row[14] = tien_hang_source - tien_hang_bvmt_le
    upsse_row[36] = tien_thue_source - tien_thue_bvmt_le
    
    upsse_row[17] = f'{int(ma_thue_percent):02d}'
    upsse_row[18] = details['s_lookup_table'].get(details['h5_val'], '')
    upsse_row[19] = details['t_lookup_regular'].get(details['h5_val'], '')
    upsse_row[20] = details['u_value']
    upsse_row[21] = details['v_lookup_table'].get(details['h5_val'], '')
    upsse_row[23] = details['store_specific_x_lookup'].get(selected_chxd, {}).get(product_name.lower(), '')
    upsse_row[31] = upsse_row[1]
    upsse_row[32] = mst_goc
    upsse_row[33] = dia_chi_goc
    return upsse_row

# --- HÀM TẠO DÒNG TỔNG HỢP (LOGIC "TOP-DOWN" ĐỒNG BỘ) ---
def _pos_add_summary_row(original_source_rows, product_name, details, product_tax, selected_chxd, is_new_price_period=False):
    """Tạo dòng tổng hợp cho khách vãng lai (người mua không lấy hóa đơn)."""
    # 1. Gom các số tổng từ file POS gốc làm "chân lý"
    total_qty = sum(_pos_to_float(r[10]) for r in original_source_rows)
    total_tien_thue_source = sum(_pos_to_float(r[14]) for r in original_source_rows)
    total_phai_thu = sum(_pos_to_float(r[13]) + _pos_to_float(r[14]) for r in original_source_rows)
    
    tmt_value = details['tmt_lookup_table'].get(product_name.lower(), 0.0)
    tax_rate_decimal = product_tax / 100.0

    # 2. Tính các thành phần của dòng Thuế BVMT (LOGIC ĐỒNG BỘ VỚI HDDT_HANDLER)
    tien_hang_dong_bvmt = round(tmt_value * total_qty)
    tien_thue_dong_bvmt = round(tien_hang_dong_bvmt * tax_rate_decimal)

    # 3. Tính "Tiền thuế dòng gốc" bằng phép trừ để bảo toàn
    tien_thue_dong_goc = total_tien_thue_source - tien_thue_dong_bvmt

    # 4. Tính "Tiền hàng dòng gốc" bằng phép trừ để bảo toàn
    tien_hang_dong_goc = total_phai_thu - tien_hang_dong_bvmt - tien_thue_dong_bvmt - tien_thue_dong_goc

    # Bắt đầu điền dữ liệu cho dòng gốc
    new_row = [''] * 37
    sample_row = original_source_rows[0]
    ngay_hd_raw = sample_row[3]
    so_ct = _pos_clean_string(str(sample_row[1]))
    
    new_row[0] = details['g5_val']
    new_row[1] = f"Khách hàng mua {product_name} không lấy hóa đơn"
    new_row[2] = _convert_to_datetime(ngay_hd_raw)
    new_row[4] = f"1{so_ct}" if so_ct else ''
    
    value_E = _pos_clean_string(new_row[4])
    suffix_d_map = {"Xăng E5 RON 92-II": "5" if is_new_price_period else "1", "Xăng RON 95-III": "6" if is_new_price_period else "2", "Dầu DO 0,05S-II": "7" if is_new_price_period else "3", "Dầu DO 0,001S-V": "8" if is_new_price_period else "4"}
    suffix_d = suffix_d_map.get(product_name, "")
    date_part = ""
    dt_obj = new_row[2]
    if isinstance(dt_obj, datetime): date_part = f"{dt_obj.day:02d}{dt_obj.month:02d}"
    
    if details['b5_val'] == "Nguyễn Huệ": new_row[3] = f"HNBK{date_part}.{suffix_d}"
    elif details['b5_val'] == "Mai Linh": new_row[3] = f"MMBK{date_part}.{suffix_d}"
    else: new_row[3] = f"{value_E[-2:]}BK{date_part}.{suffix_d}"
    
    new_row[5] = f"Xuất bán lẻ theo hóa đơn số {new_row[3]}"
    new_row[6] = details['lookup_table'].get(product_name.lower(), '')
    new_row[7], new_row[8] = product_name, "Lít"
    new_row[9] = details['g5_val']
    new_row[12] = total_qty
    
    new_row[14] = tien_hang_dong_goc
    new_row[36] = tien_thue_dong_goc
    
    new_row[17] = f'{int(product_tax):02d}'
    new_row[18] = details['s_lookup_table'].get(details['h5_val'], '')
    new_row[19] = details['t_lookup_regular'].get(details['h5_val'], '')
    new_row[20] = details['u_value']
    new_row[21] = details['v_lookup_table'].get(details['h5_val'], '')
    new_row[23] = details['store_specific_x_lookup'].get(selected_chxd, {}).get(product_name.lower(), '')
    new_row[31] = f"Khách mua {product_name} không lấy hóa đơn"
    
    return new_row, tien_hang_dong_bvmt, tien_thue_dong_bvmt

# --- HÀM TẠO FILE UPPSSE ---
def _pos_generate_upsse_rows(source_data_rows, static_data_pos, selected_chxd, is_new_price_period=False):
    """Tạo các dòng dữ liệu cho file UpSSE từ dữ liệu POS."""
    chxd_details = static_data_pos["chxd_detail_map"].get(selected_chxd)
    if not chxd_details: raise ValueError(f"Không tìm thấy thông tin chi tiết cho CHXD: '{selected_chxd}'")
    details = {**static_data_pos, **chxd_details}
    final_rows, all_tmt_rows = [], []
    no_invoice_rows = {p: [] for p in ["Xăng E5 RON 92-II", "Xăng RON 95-III", "Dầu DO 0,05S-II", "Dầu DO 0,001S-V"]}
    product_tax_map = {}
    
    for row_idx, row in enumerate(source_data_rows):
        if not row or row[0] is None: continue
        try:
            ten_kh, product_name, ma_thue_percent = _pos_clean_string(str(row[5])), _pos_clean_string(str(row[8])), _pos_to_float(row[15]) if row[15] is not None else 8.0
        except IndexError: raise ValueError(f"Dòng {row_idx + 5} trong file bảng kê POS không đủ cột.")
        if product_name and product_name not in product_tax_map: product_tax_map[product_name] = ma_thue_percent
        if ten_kh == "Người mua không lấy hóa đơn" and product_name in no_invoice_rows:
            no_invoice_rows[product_name].append(row)
        else:
            upsse_row = _pos_process_single_row(row, details, selected_chxd)
            final_rows.append(upsse_row)
            tmt_value = details['tmt_lookup_table'].get(product_name.lower(), 0.0)
            so_luong = _pos_to_float(row[10])
            if tmt_value > 0 and so_luong > 0:
                all_tmt_rows.append(_pos_create_tmt_row_for_individual(upsse_row, tmt_value, details))

    for product, original_rows in no_invoice_rows.items():
        if original_rows:
            product_tax = product_tax_map.get(product, 8.0)
            summary_row, tien_hang_bvmt, tien_thue_bvmt = _pos_add_summary_row(
                original_rows, product, details, product_tax, selected_chxd, is_new_price_period
            )
            final_rows.append(summary_row)
            
            tmt_unit = details['tmt_lookup_table'].get(product.lower(), 0)
            if tmt_unit > 0 and _pos_to_float(summary_row[12]) > 0:
                tmt_summary = list(summary_row)
                tmt_summary[1] = summary_row[1]
                tmt_summary[6], tmt_summary[7] = "TMT", "Thuế bảo vệ môi trường"
                tmt_summary[13] = tmt_unit
                tmt_summary[18] = details['s_lookup_table'].get(details['h5_val'], '')
                tmt_summary[19] = details['t_lookup_tmt'].get(details['h5_val'], '')
                tmt_summary[20] = details['u_value']
                tmt_summary[21] = details['v_lookup_table'].get(details['h5_val'], '')
                tmt_summary[14] = tien_hang_bvmt
                tmt_summary[36] = tien_thue_bvmt
                for idx in [5, 31, 32, 33]: tmt_summary[idx] = ''
                all_tmt_rows.append(tmt_summary)

    final_rows.extend(all_tmt_rows)
    return final_rows

# --- HÀM TẠO FILE EXCEL ---
def _pos_create_excel_buffer(processed_rows):
    """Tạo một đối tượng BytesIO chứa file Excel từ các dòng dữ liệu đã xử lý."""
    if not processed_rows: return None
    output_wb = Workbook()
    output_ws = output_wb.active
    headers = ["Mã khách", "Tên khách hàng", "Ngày", "Số hóa đơn", "Ký hiệu", "Diễn giải", "Mã hàng", "Tên mặt hàng", "Đvt", "Mã kho", "Mã vị trí", "Mã lô", "Số lượng", "Giá bán", "Tiền hàng", "Mã nt", "Tỷ giá", "Mã thuế", "Tk nợ", "Tk doanh thu", "Tk giá vốn", "Tk thuế có", "Cục thuế", "Vụ việc", "Bộ phận", "Lsx", "Sản phẩm", "Hợp đồng", "Phí", "Khế ước", "Nhân viên bán", "Tên KH(thuế)", "Địa chỉ (thuế)", "Mã số Thuế", "Nhóm Hàng", "Ghi chú", "Tiền thuế"]
    for _ in range(4): output_ws.append([''] * len(headers))
    output_ws.append(headers)
    for r_data in processed_rows: output_ws.append(r_data)
    for row_index in range(6, output_ws.max_row + 1):
        date_cell = output_ws[f'C{row_index}']
        if isinstance(date_cell.value, datetime):
            date_cell.number_format = 'dd/mm/yyyy'
        text_cell = output_ws[f'R{row_index}']
        text_cell.number_format = '@'
    output_ws.column_dimensions['B'].width = 35
    output_ws.column_dimensions['C'].width = 12
    output_ws.column_dimensions['D'].width = 12
    output_buffer = io.BytesIO()
    output_wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer

# --- HÀM ĐIỀU PHỐI CHÍNH ---
def process_pos_report(file_content_bytes, selected_chxd, price_periods, new_price_invoice_number, static_data_pos, selected_chxd_symbol, **kwargs):
    """
    Xử lý bảng kê POS để tạo file UpSSE.
    Bao gồm xác thực CHXD dựa trên ký hiệu hóa đơn trong bảng kê POS và mã cửa hàng ở ô B5.
    """
    try:
        # static_data_pos đã được truyền từ app.py, không cần gọi _pos_get_static_data
        if static_data_pos is None:
            raise ValueError("Dữ liệu cấu hình tĩnh cho POS chưa được tải. Vui lòng kiểm tra cấu hình ứng dụng.")

        bkhd_wb = load_workbook(io.BytesIO(file_content_bytes), data_only=True)
        bkhd_ws = bkhd_wb.active
        
        # --- BƯỚC XÁC THỰC CHXD TỪ FILE BẢNG KÊ POS (CỘT B, DÒNG 5 TRỞ ĐI) ---
        if selected_chxd_symbol is None:
            raise ValueError("Ký hiệu hóa đơn của CHXD chưa được cung cấp để xác thực.")

        # Lấy 6 ký tự cuối của ký hiệu hóa đơn từ file cấu hình Data_HDDT.xlsx
        if len(selected_chxd_symbol) < 6:
            raise ValueError(f"Ký hiệu hóa đơn trong file cấu hình Data_HDDT.xlsx ('{selected_chxd_symbol}') quá ngắn để xác thực.")
        expected_invoice_symbol_suffix = selected_chxd_symbol[-6:].upper()
        
        # Duyệt qua cột B (index 1) từ dòng 5 để kiểm tra ký hiệu hóa đơn
        # Kiểm tra ít nhất 10 dòng đầu tiên có dữ liệu để xác định tính hợp lệ
        found_matching_symbol_in_pos_file = False
        # Giới hạn số dòng kiểm tra để tránh đọc toàn bộ file lớn không cần thiết
        max_rows_to_check = min(bkhd_ws.max_row, 100) # Check up to 100 rows or end of sheet
        for row_index, row_values in enumerate(bkhd_ws.iter_rows(min_row=5, max_row=max_rows_to_check, values_only=True), start=5):
            if len(row_values) > 1 and row_values[1] is not None: # Column B is index 1
                actual_invoice_symbol_pos = _pos_clean_string(row_values[1])
                if len(actual_invoice_symbol_pos) >= 6:
                    if actual_invoice_symbol_pos[-6:].upper() == expected_invoice_symbol_suffix:
                        found_matching_symbol_in_pos_file = True
                        break # Found a matching symbol, no need to check further
        
        if not found_matching_symbol_in_pos_file:
            # ĐÃ THAY ĐỔI: Bỏ phần ký hiệu mong muốn
            raise ValueError(f"Bảng kê POS không phải của cửa hàng bạn chọn hoặc không tìm thấy ký hiệu hóa đơn hợp lệ.")

        # --- KIỂM TRA MÃ CỬA HÀNG Ở Ô B5 (GIỮ NGUYÊN) ---
        chxd_details = static_data_pos["chxd_detail_map"].get(selected_chxd)
        if not chxd_details: 
            # Trường hợp này khó xảy ra nếu selected_chxd được lấy từ get_chxd_list()
            raise ValueError(f"Không tìm thấy thông tin chi tiết cho CHXD: '{selected_chxd}'. Vui lòng kiểm tra file cấu hình Data_HDDT.xlsx.")
        
        b5_bkhd = _pos_clean_string(str(bkhd_ws['B5'].value))
        f5_norm = _pos_clean_string(chxd_details['f5_val_full']) # Full symbol from Data_HDDT.xlsx (Column K)
        
        # So sánh 6 ký tự cuối của ký hiệu từ Data_HDDT.xlsx với giá trị ô B5 của bảng kê POS
        if f5_norm and len(f5_norm) >= 6 and f5_norm[-6:] != b5_bkhd:
            raise ValueError(f"Lỗi dữ liệu: Mã cửa hàng không khớp.\n- Mã trong Bảng kê POS (ô B5): '{b5_bkhd}'\n- Mã trong file cấu hình (6 ký tự cuối cột K): '{f5_norm[-6:]}'")
        
        # --- Tiếp tục với logic hiện có ---
        all_source_rows = list(bkhd_ws.iter_rows(min_row=5, values_only=True)) # Bắt đầu đọc dữ liệu từ dòng 5
        if price_periods == '1':
            processed_rows = _pos_generate_upsse_rows(all_source_rows, static_data_pos, selected_chxd, is_new_price_period=False)
            if not processed_rows: raise ValueError("Không có dữ liệu hợp lệ để xử lý trong file POS tải lên.")
            return _pos_create_excel_buffer(processed_rows)
        else:
            if not new_price_invoice_number: raise ValueError("Vui lòng nhập 'Số hóa đơn đầu tiên của giá mới' khi chọn 2 giai đoạn giá.")
            split_index = -1
            for i, row in enumerate(all_source_rows):
                # Kiểm tra cột C (index 2) cho số hóa đơn
                if len(row) > 2 and row[2] is not None and _pos_clean_string(str(row[2])) == new_price_invoice_number:
                    split_index = i
                    break
            if split_index == -1: raise ValueError(f"Không tìm thấy số hóa đơn '{new_price_invoice_number}' để chia giai đoạn giá.")
            old_price_rows, new_price_rows = all_source_rows[:split_index], all_source_rows[split_index:]
            buffer_new = _pos_create_excel_buffer(_pos_generate_upsse_rows(new_price_rows, static_data_pos, selected_chxd, is_new_price_period=True))
            buffer_old = _pos_create_excel_buffer(_pos_generate_upsse_rows(old_price_rows, static_data_pos, selected_chxd, is_new_price_period=False))
            if not buffer_new and not buffer_old: raise ValueError("Không có dữ liệu hợp lệ để xử lý trong file POS tải lên.")
            return {'new': buffer_new, 'old': buffer_old}
    except Exception as e:
        raise e

