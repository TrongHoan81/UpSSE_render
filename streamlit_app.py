import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import NamedStyle, Alignment
from datetime import datetime
import io
import os
import re # Import regex module

# --- Cấu hình trang Streamlit ---
st.set_page_config(layout="centered", page_title="Đồng bộ dữ liệu SSE")

# Đường dẫn đến các file cần thiết (giả định cùng thư mục với script)
LOGO_PATH = "Logo.png"
DATA_FILE_PATH = "Data.xlsx" # Tên chính xác của file dữ liệu

# Định nghĩa tiêu đề cho file UpSSE.xlsx
headers = ["Mã khách", "Tên khách hàng", "Ngày", "Số hóa đơn", "Ký hiệu", "Diễn giải", "Mã hàng", "Tên mặt hàng",
           "Đvt", "Mã kho", "Mã vị trí", "Mã lô", "Số lượng", "Giá bán", "Tiền hàng", "Mã nt", "Tỷ giá", "Mã thuế",
           "Tk nợ", "Tk doanh thu", "Tk giá vốn", "Tk thuế có", "Cục thuế", "Vụ việc", "Bộ phận", "Lsx", "Sản phẩm",
           "Hợp đồng", "Phí", "Khế ước", "Nhân viên bán", "Tên KH(thuế)", "Địa chỉ (thuế)", "Mã số Thuế",
           "Nhóm Hàng", "Ghi chú", "Tiền thuế"]

# --- Kiểm tra ngày hết hạn ứng dụng ---
expiration_date = datetime(2025, 6, 26)
current_date = datetime.now()

if current_date > expiration_date:
    st.error("Có lỗi khi chạy chương trình, vui lòng liên hệ tác giả để được hỗ trợ!")
    st.info("Nguyễn Trọng Hoàn - 0902069469")
    st.stop() # Dừng ứng dụng

# --- Hàm trợ giúp ---
def to_float(value):
    """Chuyển đổi giá trị sang float, trả về 0.0 nếu không thể chuyển đổi."""
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
@st.cache_data
def get_static_data_from_excel(file_path):
    """
    Đọc dữ liệu và xây dựng các bảng tra cứu từ Data.xlsx.
    """
    try:
        wb = load_workbook(file_path, data_only=True)
        ws = wb.active

        listbox_data = []
        chxd_detail_map = {}
        store_specific_x_lookup = {}
        
        for row_idx in range(4, ws.max_row + 1):
            row_values = [cell.value for cell in ws[row_idx]]

            if len(row_values) < 18: continue

            raw_chxd_name = row_values[10]
            if raw_chxd_name and clean_string(raw_chxd_name):
                chxd_name_str = clean_string(raw_chxd_name)
                
                if chxd_name_str and chxd_name_str not in listbox_data:
                    listbox_data.append(chxd_name_str)

                g5_val = row_values[15] if pd.notna(row_values[15]) else None
                f5_val_full = clean_string(row_values[16]) if pd.notna(row_values[16]) else ''
                h5_val = clean_string(row_values[17]).lower() if pd.notna(row_values[17]) else ''
                
                if f5_val_full:
                    chxd_detail_map[chxd_name_str] = {
                        'g5_val': g5_val, 'h5_val': h5_val,
                        'f5_val_full': f5_val_full, 'b5_val': chxd_name_str
                    }
                
                store_specific_x_lookup[chxd_name_str] = {
                    "xăng e5 ron 92-ii": row_values[11],
                    "xăng ron 95-iii":   row_values[12],
                    "dầu do 0,05s-ii":   row_values[13],
                    "dầu do 0,001s-v":   row_values[14]
                }
        
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
    except FileNotFoundError:
        st.error(f"Lỗi: Không tìm thấy file {file_path}. Vui lòng đảm bảo file tồn tại.")
        st.stop()
    except Exception as e:
        st.error(f"Lỗi không mong muốn khi đọc file Data.xlsx: {e}")
        st.exception(e)
        st.stop()

# --- Các hàm xử lý dòng ---
def add_tmt_summary_row(product_name_full, g5_val, s_lookup, t_lookup_tmt, v_lookup, u_val, h5_val, 
                        representative_date, representative_symbol, total_quantity_for_tmt, tmt_unit_value_for_summary, b5_val, customer_name_for_summary_row, x_lookup_for_store):
    new_tmt_row = [''] * len(headers)
    new_tmt_row[0], new_tmt_row[1], new_tmt_row[2] = g5_val, customer_name_for_summary_row, representative_date
    value_C, value_E = clean_string(representative_date), clean_string(representative_symbol)
    suffix_d = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}.get(product_name_full, "")
    if b5_val == "Nguyễn Huệ": new_tmt_row[3] = f"HNBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    elif b5_val == "Mai Linh": new_tmt_row[3] = f"MMBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    else: new_tmt_row[3] = f"{value_E[-2:]}BK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    new_tmt_row[4] = representative_symbol
    new_tmt_row[6], new_tmt_row[7], new_tmt_row[8] = "TMT", "Thuế bảo vệ môi trường", "Lít"
    new_tmt_row[9], new_tmt_row[12] = g5_val, total_quantity_for_tmt
    new_tmt_row[13] = tmt_unit_value_for_summary
    new_tmt_row[14] = round(to_float(total_quantity_for_tmt) * to_float(tmt_unit_value_for_summary), 0)
    new_tmt_row[17] = 10
    new_tmt_row[18] = s_lookup.get(h5_val, '')
    new_tmt_row[19] = t_lookup_tmt.get(h5_val, '')
    new_tmt_row[20], new_tmt_row[21] = u_val, v_lookup.get(h5_val, '')
    new_tmt_row[23] = x_lookup_for_store.get(product_name_full.lower(), '')
    new_tmt_row[31] = ""
    new_tmt_row[36] = round(to_float(total_quantity_for_tmt) * to_float(tmt_unit_value_for_summary) * 0.1, 0)
    for idx in [5,10,11,15,16,22,24,25,26,27,28,29,30,32,33,34,35]:
        if idx != 23 and idx < len(new_tmt_row): new_tmt_row[idx] = ''
    return new_tmt_row

def add_summary_row_for_no_invoice(data_for_summary_product, source_df, product_name, headers_list,
                    g5_val, b5_val, s_lookup, t_lookup, v_lookup, x_lookup_for_store, u_val, h5_val, common_lookup_table):
    new_row = [''] * len(headers_list)
    new_row[0], new_row[1] = g5_val, f"Khách hàng mua {product_name} không lấy hóa đơn"
    new_row[2] = data_for_summary_product[0][2] if data_for_summary_product else ""
    new_row[4] = data_for_summary_product[0][4] if data_for_summary_product else ""
    value_C, value_E = clean_string(new_row[2]), clean_string(new_row[4])
    suffix_d = {"Xăng E5 RON 92-II": "1", "Xăng RON 95-III": "2", "Dầu DO 0,05S-II": "3", "Dầu DO 0,001S-V": "4"}.get(product_name, "")
    if b5_val == "Nguyễn Huệ": new_row[3] = f"HNBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    elif b5_val == "Mai Linh": new_row[3] = f"MMBK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    else: new_row[3] = f"{value_E[-2:]}BK{value_C[-2:]}.{value_C[5:7]}.{suffix_d}"
    new_row[5] = f"Xuất bán lẻ theo hóa đơn số {new_row[3]}"
    new_row[6], new_row[7], new_row[8], new_row[9] = common_lookup_table.get(clean_string(product_name).lower(), ''), product_name, "Lít", g5_val
    new_row[10], new_row[11] = '', ''
    total_M = sum(to_float(r[12]) for r in data_for_summary_product)
    new_row[12] = total_M
    new_row[13] = max((to_float(r[13]) for r in data_for_summary_product), default=0.0)

    # Lọc DataFrame nguồn để tính tổng
    filtered_df = source_df[(source_df[5].astype(str) == "Người mua không lấy hóa đơn") & (source_df[8].astype(str) == product_name)]
    tien_hang_hd = filtered_df[13].apply(to_float).sum()
    TienthueHD_from_original_bkhd = filtered_df[14].apply(to_float).sum()
    
    price_per_liter = {"Xăng E5 RON 92-II": 1900, "Xăng RON 95-III": 2000, "Dầu DO 0,05S-II": 1000, "Dầu DO 0,001S-V": 1000}.get(product_name, 0)
    new_row[14] = tien_hang_hd - round(total_M * price_per_liter, 0)
    
    new_row[15], new_row[16], new_row[17] = '', '', 10
    new_row[18], new_row[19] = s_lookup.get(h5_val, ''), t_lookup.get(h5_val, '')
    new_row[20], new_row[21] = u_val, v_lookup.get(h5_val, '')
    new_row[22] = ''
    new_row[23] = x_lookup_for_store.get(clean_string(product_name).lower(), '')
    for i in range(24, 31): new_row[i] = ''
    new_row[31] = f"Khách mua {product_name} không lấy hóa đơn"
    new_row[32], new_row[33], new_row[34], new_row[35] = "", "", '', ''
    
    new_row[36] = TienthueHD_from_original_bkhd - round(total_M * price_per_liter * 0.1, 0) 
    return new_row

def create_per_invoice_tmt_row(original_row_data, tmt_value, g5_val, s_lookup, t_lookup_tmt, v_lookup, u_val, h5_val):
    tmt_row = list(original_row_data)
    tmt_row[6], tmt_row[7], tmt_row[8] = "TMT", "Thuế bảo vệ môi trường", "Lít"
    tmt_row[9] = g5_val
    tmt_row[13] = tmt_value
    tmt_row[14] = round(tmt_value * to_float(original_row_data[12]), 0)
    tmt_row[17] = 10
    tmt_row[18] = s_lookup.get(h5_val, '')
    tmt_row[19] = t_lookup_tmt.get(h5_val, '')
    tmt_row[20], tmt_row[21] = u_val, v_lookup.get(h5_val, '')
    tmt_row[31] = ""
    tmt_row[36] = round(tmt_value * to_float(original_row_data[12]) * 0.1, 0)
    for idx in [5, 10, 11, 15, 16, 22, 24, 25, 26, 27, 28, 29, 30, 32, 33, 34, 35]:
        if idx < len(tmt_row): tmt_row[idx] = ''
    return tmt_row

# --- Tải dữ liệu tĩnh ---
static_data = get_static_data_from_excel(DATA_FILE_PATH)

# --- Giao diện người dùng ---
col1, col2 = st.columns([1, 4]) 
with col1:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=140)
with col2:
    st.markdown("""
    <div style="display: flex; flex-direction: column; justify-content: center; align-items: center; height: 100%; padding-top: 10px;">
        <h2 style="color: red; font-weight: bold; margin-bottom: 0px; font-size: 24px;">CÔNG TY CỔ PHẦN XĂNG DẦU</h2>
        <h2 style="color: red; font-weight: bold; margin-top: 0px; font-size: 24px;">DẦU KHÍ NAM ĐỊNH</h2>
    </div>
    """, unsafe_allow_html=True)

st.title("Công cụ đồng bộ dữ liệu SSE")

st.info("Ứng dụng này có thể tự động xử lý các file Excel bị lỗi định dạng.")

selected_chxd = st.selectbox("Chọn CHXD:", options=[""] + static_data["listbox_data"])
uploaded_file = st.file_uploader("Tải lên file bảng kê hóa đơn (.xlsx)", type=["xlsx"])

# --- Xử lý chính ---
if st.button("Xử lý"):
    if not selected_chxd: st.warning("Vui lòng chọn một CHXD.")
    elif uploaded_file is None: st.warning("Vui lòng tải lên file bảng kê hóa đơn.")
    else:
        try:
            # --- BƯỚC 1: LÀM SẠCH FILE ---
            try:
                df = pd.read_excel(uploaded_file, engine='calamine', header=None)
                cleaned_buffer = io.BytesIO()
                df.to_excel(cleaned_buffer, index=False, header=False, engine='openpyxl')
                cleaned_buffer.seek(0)
            except Exception as e:
                st.error("Lỗi nghiêm trọng khi đọc và làm sạch file Excel.")
                st.error(f"Chi tiết: {e}")
                st.error("Hãy đảm bảo bạn đã thêm 'pandas-calamine' vào file requirements.txt.")
                st.stop()

            # --- BƯỚC 2: XỬ LÝ DỮ LIỆU TỪ FILE ĐÃ SẠCH ---
            chxd_details = static_data["chxd_detail_map"].get(selected_chxd)
            if not chxd_details:
                st.error(f"Không tìm thấy thông tin chi tiết cho CHXD: '{selected_chxd}'")
                st.stop()
            
            cleaned_buffer.seek(0)
            source_df = pd.read_excel(cleaned_buffer, header=None, skiprows=4)

            long_cells = [f"H{r_idx+5}" for r_idx, cell_val in enumerate(source_df[7]) if cell_val and len(str(cell_val)) > 128]
            if long_cells:
                st.error("Địa chỉ trên ô " + ', '.join(long_cells) + " quá dài, hãy điều chỉnh và thử lại.")
                st.stop()
            
            vi_tri_cu_idx = [0, 1, 2, 3, 4, 5, 7, 6, 8, 10, 11, 13, 14, 16]
            intermediate_data = []
            for _, row_series in source_df.iterrows():
                row = row_series.tolist()
                if len(row) <= max(vi_tri_cu_idx): continue
                new_row = [row[i] for i in vi_tri_cu_idx]
                if pd.notna(new_row[3]):
                    try:
                        date_val = pd.to_datetime(new_row[3]).strftime('%Y-%m-%d')
                        new_row[3] = date_val
                    except (ValueError, TypeError): pass
                ma_kh = new_row[4]
                new_row.append("No" if pd.isna(ma_kh) or len(clean_string(str(ma_kh))) > 9 else "Yes")
                intermediate_data.append(new_row)

            if not intermediate_data:
                st.error("Không có dữ liệu hợp lệ trong file bảng kê.")
                st.stop()
            
            b2_bkhd = clean_string(str(intermediate_data[0][1]))
            f5_norm = clean_string(chxd_details['f5_val_full'])
            if f5_norm.startswith('1'): f5_norm = f5_norm[1:]
            if f5_norm != b2_bkhd:
                st.error("Bảng kê hóa đơn không phải của cửa hàng bạn chọn.")
                st.stop()

            final_rows, all_tmt_rows = [[''] * len(headers) for _ in range(4)] + [headers], []
            no_invoice_rows = {p: [] for p in ["Xăng E5 RON 92-II", "Xăng RON 95-III", "Dầu DO 0,05S-II", "Dầu DO 0,001S-V"]}
            
            # --- Gộp các lookup tables vào chxd_details để truyền đi cho gọn ---
            details = static_data | chxd_details

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
                        all_tmt_rows.append(create_per_invoice_tmt_row(upsse_row, tmt_value, details['g5_val'], details['s_lookup_table'], details['t_lookup_tmt'], details['v_lookup_table'], details['u_value'], details['h5_val']))

            for product, rows in no_invoice_rows.items():
                if rows:
                    summary_row = add_summary_row_for_no_invoice(rows, source_df, product, headers, details['g5_val'], details['b5_val'], details['s_lookup_table'], details['t_lookup_regular'], details['v_lookup_table'], details['store_specific_x_lookup'].get(selected_chxd, {}), details['u_value'], details['h5_val'], details['lookup_table'])
                    final_rows.append(summary_row)
                    tmt_unit = details['tmt_lookup_table'].get(product.lower(), 0)
                    total_qty = sum(to_float(r[12]) for r in rows)
                    tmt_summary = add_tmt_summary_row(product, details['g5_val'], details['s_lookup_table'], details['t_lookup_tmt'], details['v_lookup_table'], details['u_value'], details['h5_val'], summary_row[2], summary_row[4], total_qty, tmt_unit, details['b5_val'], summary_row[1], details['store_specific_x_lookup'].get(selected_chxd, {}))
                    all_tmt_rows.append(tmt_summary)
            
            final_rows.extend(all_tmt_rows)
            
            output_wb = Workbook()
            output_ws = output_wb.active
            for r_data in final_rows: output_ws.append(r_data)

            date_style = NamedStyle(name="date_style", number_format='DD/MM/YYYY')
            for cell in output_ws['C']:
                if isinstance(cell.value, str):
                    try:
                        cell.value = datetime.strptime(cell.value, '%Y-%m-%d').date()
                    except ValueError: pass
                if isinstance(cell.value, datetime):
                    cell.style = date_style
            
            output_ws.column_dimensions['B'].width = 35
            output_ws.column_dimensions['C'].width = 12
            output_ws.column_dimensions['D'].width = 12
            
            output_buffer = io.BytesIO()
            output_wb.save(output_buffer)
            st.success("Đã tạo file UpSSE.xlsx thành công!")
            st.download_button("Tải xuống file UpSSE.xlsx", output_buffer.getvalue(), "UpSSE.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"Lỗi trong quá trình xử lý file: {e}")
            st.exception(e)
