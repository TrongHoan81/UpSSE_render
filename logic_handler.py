import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import NamedStyle
from datetime import datetime
import re
import io

# --- Các hàm trợ giúp (Không thay đổi) ---
def to_float(value):
    try:
        if isinstance(value, str):
            value = value.replace(",", "").strip()
        return float(value)
    except (ValueError, TypeError):
        return 0.0

def clean_string(s):
    if s is None:
        return ""
    return re.sub(r'\s+', ' ', str(s)).strip()

# --- Hàm đọc dữ liệu tĩnh (Không thay đổi) ---
def get_static_data_from_excel(file_path):
    try:
        wb = load_workbook(file_path, data_only=True)
        ws = wb.active
        
        chxd_detail_map = {}
        store_specific_x_lookup = {}
        
        for row_idx in range(4, ws.max_row + 1):
            row_values = [cell.value for cell in ws[row_idx]]
            if len(row_values) < 18: continue
            
            chxd_name = clean_string(row_values[10])
            if chxd_name:
                chxd_detail_map[chxd_name] = {
                    'g5_val': row_values[15],
                    'h5_val': clean_string(row_values[17]).lower(),
                    'f5_val_full': clean_string(row_values[16]),
                    'b5_val': chxd_name
                }
                store_specific_x_lookup[chxd_name] = {
                    "xăng e5 ron 92-ii": row_values[11], "xăng ron 95-iii": row_values[12],
                    "dầu do 0,05s-ii": row_values[13], "dầu do 0,001s-v": row_values[14]
                }
        
        listbox_data = list(chxd_detail_map.keys())
        
        def get_lookup(min_r, max_r, min_c=9, max_c=10):
            return {clean_string(row[0]).lower(): row[1] for row in ws.iter_rows(min_row=min_r, max_row=max_r, min_col=min_c, max_col=max_c, values_only=True) if row[0] and row[1]}

        tmt_lookup_table = {k: to_float(v) for k, v in get_lookup(10, 13).items()}

        wb.close()
        return {
            "listbox_data": listbox_data, "lookup_table": get_lookup(4, 7),
            "tmt_lookup_table": tmt_lookup_table, "s_lookup_table": get_lookup(29, 31),
            "t_lookup_regular": get_lookup(33, 35), "t_lookup_tmt": get_lookup(48, 50),
            "v_lookup_table": get_lookup(53, 55), "u_value": ws['J36'].value,
            "chxd_detail_map": chxd_detail_map, "store_specific_x_lookup": store_specific_x_lookup
        }
    except Exception as e:
        print(f"Error reading static data: {e}")
        return None

# --- Hàm xử lý logic lõi (được tái sử dụng) ---
# ******** START OF CHANGE 1 ********
# Add a new parameter is_new_price_period
def _generate_upsse_rows(source_data_rows, full_bkhd_ws, static_data, selected_chxd, is_new_price_period=False):
# ******** END OF CHANGE 1 ********
    chxd_details = static_data["chxd_detail_map"].get(selected_chxd)
    if not chxd_details:
        raise ValueError(f"Không tìm thấy thông tin chi tiết cho CHXD: '{selected_chxd}'")
    
    details = {**static_data, **chxd_details}

    vi_tri_cu_idx = [0, 1, 2, 3, 4, 5, 7, 6, 8, 10, 11, 13, 14, 16]
    intermediate_data = []
    for row in source_data_rows:
        if len(row) <= max(vi_tri_cu_idx): continue
        new_row = [row[i] for i in vi_tri_cu_idx]
        if new_row[3] and not isinstance(new_row[3], datetime):
            try:
                # Handle dd-mm-yyyy format from POS
                new_row[3] = datetime.strptime(str(new_row[3]).split(' ')[0], '%d-%m-%Y')
            except (ValueError, TypeError): pass
        if isinstance(new_row[3], datetime):
             new_row[3] = new_row[3].strftime('%Y-%m-%d')
        ma_kh = clean_string(str(new_row[4]))
        new_row.append("No" if not ma_kh or len(ma_kh) > 9 else "Yes")
        intermediate_data.append(new_row)

    if not intermediate_data:
        return []

    final_rows, all_tmt_rows = [], []
    no_invoice_rows = {p: [] for p in ["Xăng E5 RON 92-II", "Xăng RON 95-III", "Dầu DO 0,05S-II", "Dầu DO 0,001S-V"]}
    headers = ["Mã khách", "Tên khách hàng", "Ngày", "Số hóa đơn", "Ký hiệu", "Diễn giải", "Mã hàng", "Tên mặt hàng", "Đvt", "Mã kho", "Mã vị trí", "Mã lô", "Số lượng", "Giá bán", "Tiền hàng", "Mã nt", "Tỷ giá", "Mã thuế", "Tk nợ", "Tk doanh thu", "Tk giá vốn", "Tk thuế có", "Cục thuế", "Vụ việc", "Bộ phận", "Lsx", "Sản phẩm", "Hợp đồng", "Phí", "Khế ước", "Nhân viên bán", "Tên KH(thuế)", "Địa chỉ (thuế)", "Mã số Thuế", "Nhóm Hàng", "Ghi chú", "Tiền thuế"]
    
    for row in intermediate_data:
        upsse_row = [''] * len(headers)
        upsse_row[0] = clean_string(str(row[4])) if row[-1] == 'Yes' and pd.notna(row[4]) else details['g5_val']
        upsse_row[1], upsse_row[2] = clean_string(str(row[5])), row[3]
        b_orig, c_orig = clean_string(str(row[1])), clean_string(str(row[2]))
        
        if details['b5_val'] == "Nguyễn Huệ": upsse_row[3] = f"HN{c_orig[-6:]}"
        elif details['b5_val'] == "Mai Linh": upsse_row[3] = f"MM{c_orig[-6:]}"
        else: upsse_row[3] = f"{b_orig[-2:]}{c_orig[-6:]}"

        upsse_row[4] = f"1{b_orig}" if b_orig else ''
        upsse_row[5] = f"Xuất bán lẻ theo hóa đơn số {upsse_row[3]}"
        product_name = clean_string(str(row[8]))
        upsse_row[6] = details['lookup_table'].get(product_name.lower(), '')
        upsse_row[7], upsse_row[8] = product_name, "Lít"
        upsse_row[9] = details['g5_val']
        upsse_row[12] = to_float(row[9])
        tmt_value = details['tmt_lookup_table'].get(product_name.lower(), 0.0)
        upsse_row[13] = round(to_float(row[10]) / 1.1 - tmt_value, 2)
        upsse_row[14] = to_float(row[11]) - round(tmt_value * upsse_row[12])
        upsse_row[17] = 10
        upsse_row[18] = details['s_lookup_table'].get(details['h5_val'], '')
        upsse_row[19] = details['t_lookup_regular'].get(details['h5_val'], '')
        upsse_row[20] = details['u_value']
        upsse_row[21] = details['v_lookup_table'].get(details['h5_val'], '')
        upsse_row[23] = details['store_specific_x_lookup'].get(selected_chxd, {}).get(product_name.lower(), '')
        upsse_row[31] = upsse_row[1]
        upsse_row[32], upsse_row[33] = row[6], row[7]
        upsse_row[36] = to_float(row[12]) - round(upsse_row[12] * tmt_value * 0.1)

        if upsse_row[1] == "Người mua không lấy hóa đơn" and product_name in no_invoice_rows:
            no_invoice_rows[product_name].append(upsse_row)
        else:
            final_rows.append(upsse_row)
            if tmt_value > 0 and upsse_row[12] > 0:
                all_tmt_rows.append(create_tmt_row(upsse_row, tmt_value, details))

    for product, rows in no_invoice_rows.items():
        if rows:
            # ******** START OF CHANGE 2 ********
            # Pass the new parameter down to the summary function
            summary_row = add_summary_row(rows, full_bkhd_ws, product, details, is_new_price_period=is_new_price_period)
            # ******** END OF CHANGE 2 ********
            final_rows.append(summary_row)
            tmt_unit = details['tmt_lookup_table'].get(product.lower(), 0)
            tmt_summary = create_tmt_row(summary_row, tmt_unit, details)
            tmt_summary[1] = summary_row[1]
            all_tmt_rows.append(tmt_summary)

    final_rows.extend(all_tmt_rows)
    return final_rows

# --- Hàm điều phối chính ---
def process_file_with_price_periods(uploaded_file_content, static_data, selected_chxd, price_periods, new_price_invoice_number):
    try:
        bkhd_wb = load_workbook(io.BytesIO(uploaded_file_content), data_only=True)
        bkhd_ws = bkhd_wb.active
        
        chxd_details = static_data["chxd_detail_map"].get(selected_chxd)
        if not chxd_details: 
            raise ValueError(f"Không tìm thấy thông tin cho CHXD: '{selected_chxd}'")
        
        b5_bkhd = clean_string(str(bkhd_ws['B5'].value))

        f5_norm = clean_string(chxd_details['f5_val_full'])
        if f5_norm.startswith('1'): 
            f5_norm = f5_norm[1:]
        
        if f5_norm != b5_bkhd:
            error_message = (
                "Lỗi dữ liệu: Mã cửa hàng không khớp. Vui lòng kiểm tra lại.\n\n"
                f"   - Mã trong file Bảng kê tải lên (lấy từ ô B5): '{b5_bkhd}'\n"
                f"   - Mã trong file cấu hình Data.xlsx (cột Q):    '{f5_norm}'\n\n"
                "**Gợi ý**: Hãy đảm bảo mã seri ở ô B5 của file bảng kê khớp với mã trong file cấu hình."
            )
            raise ValueError(error_message)

        all_source_rows = list(bkhd_ws.iter_rows(min_row=5, values_only=True))
        
        old_price_rows = []
        new_price_rows = []
        
        if price_periods == '1':
            old_price_rows = all_source_rows
        else: # price_periods == '2'
            split_found = False
            invoice_col_idx = 2 # Column C is 'Số' (invoice number)
            for row in all_source_rows:
                if len(row) > invoice_col_idx and row[invoice_col_idx] is not None:
                    current_invoice = clean_string(str(row[invoice_col_idx]))
                    if not split_found and current_invoice == new_price_invoice_number:
                        split_found = True
                
                if split_found:
                    new_price_rows.append(row)
                else:
                    old_price_rows.append(row)
            
            if not split_found:
                raise ValueError(f"Không tìm thấy số hóa đơn '{new_price_invoice_number}' để chia giai đoạn giá.")

        # ******** START OF CHANGE 3 ********
        # Call the processing function with the correct flag for each period
        processed_rows_old = _generate_upsse_rows(old_price_rows, bkhd_ws, static_data, selected_chxd, is_new_price_period=False)
        processed_rows_new = _generate_upsse_rows(new_price_rows, bkhd_ws, static_data, selected_chxd, is_new_price_period=True)
        # ******** END OF CHANGE 3 ********
        
        combined_rows = processed_rows_old + processed_rows_new
        if not combined_rows:
            raise ValueError("Không có dữ liệu hợp lệ để xử lý trong file tải lên.")

        output_wb = Workbook()
        output_ws = output_wb.active
        headers = ["Mã khách", "Tên khách hàng", "Ngày", "Số hóa đơn", "Ký hiệu", "Diễn giải", "Mã hàng", "Tên mặt hàng", "Đvt", "Mã kho", "Mã vị trí", "Mã lô", "Số lượng", "Giá bán", "Tiền hàng", "Mã nt", "Tỷ giá", "Mã thuế", "Tk nợ", "Tk doanh thu", "Tk giá vốn", "Tk thuế có", "Cục thuế", "Vụ việc", "Bộ phận", "Lsx", "Sản phẩm", "Hợp đồng", "Phí", "Khế ước", "Nhân viên bán", "Tên KH(thuế)", "Địa chỉ (thuế)", "Mã số Thuế", "Nhóm Hàng", "Ghi chú", "Tiền thuế"]
        
        for _ in range(4): output_ws.append([''] * len(headers))
        output_ws.append(headers)
        for r_data in combined_rows: output_ws.append(r_data)

        date_style = NamedStyle(name="date_style", number_format='DD/MM/YYYY')
        for row_index in range(6, output_ws.max_row + 1):
            cell = output_ws[f'C{row_index}']
            if isinstance(cell.value, str) and '-' in cell.value:
                try:
                    date_obj = datetime.strptime(cell.value, '%Y-%m-%d')
                    cell.value = date_obj 
                    cell.style = date_style 
                except (ValueError, TypeError): pass
            elif isinstance(cell.value, datetime):
                cell.style = date_style

        output_ws.column_dimensions['B'].width = 35
        output_ws.column_dimensions['C'].width = 12
        output_ws.column_dimensions['D'].width = 12
        
        output_buffer = io.BytesIO()
        output_wb.save(output_buffer)
        output_buffer.seek(0)
        
        return output_buffer

    except Exception as e:
        print(f"Error during processing: {e}")
        raise e

# --- Các hàm phụ ---
# ******** START OF CHANGE 4 ********
# Add the new parameter to the function definition
def add_summary_row(data_product, bkhd_source_ws, product_name, details, is_new_price_period=False):
# ******** END OF CHANGE 4 ********
    headers = ["Mã khách", "Tên khách hàng", "Ngày", "Số hóa đơn", "Ký hiệu", "Diễn giải", "Mã hàng", "Tên mặt hàng", "Đvt", "Mã kho", "Mã vị trí", "Mã lô", "Số lượng", "Giá bán", "Tiền hàng", "Mã nt", "Tỷ giá", "Mã thuế", "Tk nợ", "Tk doanh thu", "Tk giá vốn", "Tk thuế có", "Cục thuế", "Vụ việc", "Bộ phận", "Lsx", "Sản phẩm", "Hợp đồng", "Phí", "Khế ước", "Nhân viên bán", "Tên KH(thuế)", "Địa chỉ (thuế)", "Mã số Thuế", "Nhóm Hàng", "Ghi chú", "Tiền thuế"]
    new_row = [''] * len(headers)
    new_row[0] = details['g5_val']
    new_row[1] = f"Khách hàng mua {product_name} không lấy hóa đơn"
    new_row[2] = data_product[0][2] if data_product else ""
    new_row[4] = data_product[0][4] if data_product else ""
    
    value_C = clean_string(new_row[2])
    value_E = clean_string(new_row[4])
    
    # ******** START OF CHANGE 5 ********
    # Use different suffix maps based on the price period
    if is_new_price_period:
        # Suffixes for the new price period (5, 6, 7, 8)
        suffix_d_map = {
            "Xăng E5 RON 92-II": "5", "Xăng RON 95-III": "6",
            "Dầu DO 0,05S-II": "7", "Dầu DO 0,001S-V": "8"
        }
    else:
        # Original suffixes for the old price period (1, 2, 3, 4)
        suffix_d_map = {
            "Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2",
            "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"
        }
    suffix_d = suffix_d_map.get(product_name, "")
    # ******** END OF CHANGE 5 ********

    date_part = ""
    if value_C and len(value_C) >= 10:
        try:
            dt_obj = datetime.strptime(value_C, '%Y-%m-%d')
            date_part = f"{dt_obj.day:02d}{dt_obj.month:02d}"
        except ValueError:
            pass 
    
    if details['b5_val'] == "Nguyễn Huệ": new_row[3] = f"HNBK{date_part}.{suffix_d}"
    elif details['b5_val'] == "Mai Linh": new_row[3] = f"MMBK{date_part}.{suffix_d}"
    else: new_row[3] = f"{value_E[-2:]}BK{date_part}.{suffix_d}"
    
    new_row[5] = f"Xuất bán lẻ theo hóa đơn số {new_row[3]}"
    new_row[6] = details['lookup_table'].get(product_name.lower(), '')
    new_row[7] = product_name
    new_row[8] = "Lít"
    new_row[9] = details['g5_val']
    
    total_qty = sum(to_float(r[12]) for r in data_product)
    new_row[12] = total_qty
    new_row[13] = max((to_float(r[13]) for r in data_product), default=0.0)

    tien_hang_hd = sum(to_float(r[13]) for r in bkhd_source_ws.iter_rows(min_row=5, values_only=True) if clean_string(str(r[5])) == "Người mua không lấy hóa đơn" and clean_string(str(r[8])) == product_name)
    tienthue_hd = sum(to_float(r[14]) for r in bkhd_source_ws.iter_rows(min_row=5, values_only=True) if clean_string(str(r[5])) == "Người mua không lấy hóa đơn" and clean_string(str(r[8])) == product_name)
    
    price_per_liter = {"Xăng E5 RON 92-II": 1900, "Xăng RON 95-III": 2000, "Dầu DO 0,05S-II": 1000, "Dầu DO 0,001S-V": 1000}.get(product_name, 0)
    new_row[14] = tien_hang_hd - round(total_qty * price_per_liter, 0)
    
    new_row[17] = 10
    new_row[18] = details['s_lookup_table'].get(details['h5_val'], '')
    new_row[19] = details['t_lookup_regular'].get(details['h5_val'], '')
    new_row[20] = details['u_value']
    new_row[21] = details['v_lookup_table'].get(details['h5_val'], '')
    new_row[23] = details['store_specific_x_lookup'].get(details['b5_val'], {}).get(product_name.lower(), '')
    new_row[31] = f"Khách mua {product_name} không lấy hóa đơn"
    new_row[36] = tienthue_hd - round(total_qty * price_per_liter * 0.1, 0)
    return new_row

def create_tmt_row(original_row, tmt_value, details):
    tmt_row = list(original_row)
    tmt_row[6], tmt_row[7], tmt_row[8] = "TMT", "Thuế bảo vệ môi trường", "Lít"
    tmt_row[9] = details['g5_val']
    tmt_row[13] = tmt_value
    tmt_row[14] = round(tmt_value * to_float(original_row[12]), 0)
    tmt_row[17] = 10
    tmt_row[18] = details['s_lookup_table'].get(details['h5_val'], '')
    tmt_row[19] = details['t_lookup_tmt'].get(details['h5_val'], '')
    tmt_row[20], tmt_row[21] = details['u_value'], details['v_lookup_table'].get(details['h5_val'], '')
    tmt_row[31] = ""
    tmt_row[36] = round(tmt_value * to_float(original_row[12]) * 0.1, 0)
    for idx in [5, 10, 11, 15, 16, 22, 24, 25, 26, 27, 28, 29, 30, 32, 33, 34, 35]:
        if idx < len(tmt_row): tmt_row[idx] = ''
    return tmt_row
