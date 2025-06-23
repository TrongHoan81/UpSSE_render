import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import NamedStyle
from datetime import datetime
import re
import io

# Các hàm này được sao chép và điều chỉnh từ file streamlit_app.py gốc của bạn.
# Chúng ta giữ nguyên logic cốt lõi mà bạn đã phát triển.

# --- Các hàm trợ giúp ---
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

# --- Hàm đọc dữ liệu tĩnh ---
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

# --- Các hàm xử lý dòng ---
def add_summary_row(data_product, source_df, product_name, details):
    headers = ["Mã khách", "Tên khách hàng", "Ngày", "Số hóa đơn", "Ký hiệu", "Diễn giải", "Mã hàng", "Tên mặt hàng", "Đvt", "Mã kho", "Mã vị trí", "Mã lô", "Số lượng", "Giá bán", "Tiền hàng", "Mã nt", "Tỷ giá", "Mã thuế", "Tk nợ", "Tk doanh thu", "Tk giá vốn", "Tk thuế có", "Cục thuế", "Vụ việc", "Bộ phận", "Lsx", "Sản phẩm", "Hợp đồng", "Phí", "Khế ước", "Nhân viên bán", "Tên KH(thuế)", "Địa chỉ (thuế)", "Mã số Thuế", "Nhóm Hàng", "Ghi chú", "Tiền thuế"]
    new_row = [''] * len(headers)
    new_row[0] = details['g5_val']
    new_row[1] = f"Khách hàng mua {product_name} không lấy hóa đơn"
    new_row[2] = data_product[0][2] if data_product else ""
    new_row[4] = data_product[0][4] if data_product else ""
    
    value_C = clean_string(new_row[2])
    value_E = clean_string(new_row[4])
    suffix_d = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}.get(product_name, "")
    
    if details['b5_val'] == "Nguyễn Huệ": new_row[3] = f"HNBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    elif details['b5_val'] == "Mai Linh": new_row[3] = f"MMBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    else: new_row[3] = f"{value_E[-2:]}BK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    
    new_row[5] = f"Xuất bán lẻ theo hóa đơn số {new_row[3]}"
    new_row[6] = details['lookup_table'].get(product_name.lower(), '')
    new_row[7] = product_name
    new_row[8] = "Lít"
    new_row[9] = details['g5_val']
    
    total_qty = sum(to_float(r[12]) for r in data_product)
    new_row[12] = total_qty
    new_row[13] = max((to_float(r[13]) for r in data_product), default=0.0)

    filtered_df = source_df[(source_df[5].astype(str) == "Người mua không lấy hóa đơn") & (source_df[8].astype(str) == product_name)]
    tien_hang_hd = filtered_df[13].apply(to_float).sum()
    tienthue_hd = filtered_df[14].apply(to_float).sum()
    
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

# --- Hàm xử lý chính ---
def process_excel_file(uploaded_file_content, static_data, selected_chxd):
    """
    Hàm chính để xử lý file Excel.
    Nhận nội dung file đã tải lên, dữ liệu tĩnh, và CHXD đã chọn.
    Trả về một đối tượng BytesIO chứa file Excel kết quả.
    """
    try:
        # Làm sạch file đầu vào bằng pandas-calamine
        df = pd.read_excel(io.BytesIO(uploaded_file_content), engine='calamine', header=None)
        cleaned_buffer = io.BytesIO()
        df.to_excel(cleaned_buffer, index=False, header=False, engine='openpyxl')
        cleaned_buffer.seek(0)

        source_df = pd.read_excel(cleaned_buffer, header=None, skiprows=4)
        cleaned_buffer.seek(0)

        # Bắt đầu xử lý logic chính
        chxd_details = static_data["chxd_detail_map"].get(selected_chxd)
        if not chxd_details:
            raise ValueError(f"Không tìm thấy thông tin chi tiết cho CHXD: '{selected_chxd}'")
        
        details = static_data | chxd_details

        vi_tri_cu_idx = [0, 1, 2, 3, 4, 5, 7, 6, 8, 10, 11, 13, 14, 16]
        intermediate_data = []
        for _, row_series in source_df.iterrows():
            row = row_series.tolist()
            if len(row) <= max(vi_tri_cu_idx): continue
            new_row = [row[i] for i in vi_tri_cu_idx]
            if pd.notna(new_row[3]):
                try:
                    new_row[3] = pd.to_datetime(new_row[3]).strftime('%Y-%m-%d')
                except (ValueError, TypeError): pass
            new_row.append("No" if pd.isna(new_row[4]) or len(clean_string(str(new_row[4]))) > 9 else "Yes")
            intermediate_data.append(new_row)

        if not intermediate_data:
            raise ValueError("Không có dữ liệu hợp lệ trong file bảng kê.")

        b2_bkhd = clean_string(str(intermediate_data[0][1]))
        f5_norm = clean_string(details['f5_val_full'])
        if f5_norm.startswith('1'): f5_norm = f5_norm[1:]
        if f5_norm != b2_bkhd:
            raise ValueError("Bảng kê hóa đơn không phải của cửa hàng bạn chọn.")

        headers = ["Mã khách", "Tên khách hàng", "Ngày", "Số hóa đơn", "Ký hiệu", "Diễn giải", "Mã hàng", "Tên mặt hàng", "Đvt", "Mã kho", "Mã vị trí", "Mã lô", "Số lượng", "Giá bán", "Tiền hàng", "Mã nt", "Tỷ giá", "Mã thuế", "Tk nợ", "Tk doanh thu", "Tk giá vốn", "Tk thuế có", "Cục thuế", "Vụ việc", "Bộ phận", "Lsx", "Sản phẩm", "Hợp đồng", "Phí", "Khế ước", "Nhân viên bán", "Tên KH(thuế)", "Địa chỉ (thuế)", "Mã số Thuế", "Nhóm Hàng", "Ghi chú", "Tiền thuế"]
        final_rows, all_tmt_rows = [[''] * len(headers) for _ in range(4)] + [headers], []
        no_invoice_rows = {p: [] for p in ["Xăng E5 RON 92-II", "Xăng RON 95-III", "Dầu DO 0,05S-II", "Dầu DO 0,001S-V"]}

        for row in intermediate_data:
            upsse_row = [''] * len(headers)
            upsse_row[0] = clean_string(str(row[4])) if row[-1] == 'Yes' and pd.notna(row[4]) else details['g5_val']
            upsse_row[1], upsse_row[2] = clean_string(str(row[5])), row[3]
            b_orig, c_orig = clean_string(str(row[1])), clean_string(str(row[2]))
            
            if c_orig and len(c_orig) >= 6:
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
                summary_row = add_summary_row(rows, source_df, product, details)
                final_rows.append(summary_row)
                tmt_unit = details['tmt_lookup_table'].get(product.lower(), 0)
                tmt_summary = create_tmt_row(summary_row, tmt_unit, details)
                tmt_summary[1] = summary_row[1]
                all_tmt_rows.append(tmt_summary)

        final_rows.extend(all_tmt_rows)
        
        output_wb = Workbook()
        output_ws = output_wb.active
        for r_data in final_rows: output_ws.append(r_data)

        date_style = NamedStyle(name="date_style", number_format='DD/MM/YYYY')
        for cell in output_ws['C']:
            if isinstance(cell.value, str):
                try: cell.value = datetime.strptime(cell.value, '%Y-%m-%d').date()
                except ValueError: pass
            if isinstance(cell.value, datetime):
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
        # Trả về lỗi để Flask có thể xử lý
        raise e
