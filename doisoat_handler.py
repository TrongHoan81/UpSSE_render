import io
import re
from collections import defaultdict
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook

# --- CÁC HÀM TIỆN ÍCH ---
def _clean_string(s):
    """Làm sạch chuỗi, loại bỏ khoảng trắng thừa và ký tự '."""
    if s is None: return ""
    cleaned_s = str(s).strip()
    if cleaned_s.startswith("'"): cleaned_s = cleaned_s[1:]
    return re.sub(r'\s+', ' ', cleaned_s)

def _to_float(value):
    """Chuyển đổi giá trị sang float, xử lý các trường hợp lỗi."""
    if value is None: return 0.0
    try:
        return float(str(value).replace(',', '').strip())
    except (ValueError, TypeError): return 0.0

def _format_number(num):
    """Định dạng số thành chuỗi có dấu phẩy phân cách hàng nghìn và 2 chữ số thập phân."""
    try:
        return f"{num:,.2f}"
    except (ValueError, TypeError):
        return "0.00"

def _excel_date_to_datetime(excel_date):
    """Chuyển đổi ngày tháng từ định dạng Excel sang đối tượng datetime."""
    if isinstance(excel_date, (int, float)):
        try:
            return pd.to_datetime(excel_date, unit='D', origin='1899-12-30').to_pydatetime()
        except Exception:
            return None
    elif isinstance(excel_date, datetime):
        return excel_date
    return None

# --- CÁC HÀM PHÂN TÍCH FILE ---

def _parse_hddt_file(hddt_bytes):
    """Phân tích dữ liệu từ file Bảng kê HĐĐT."""
    try:
        wb = load_workbook(io.BytesIO(hddt_bytes), data_only=True)
        ws = wb.active
        
        pos_invoices = []
        direct_petroleum_invoices = []
        other_invoices = []

        PETROLEUM_PRODUCTS = [
            'Xăng RON 95-III', 'Dầu DO 0,05S-II', 
            'Xăng E5 RON 92-II', 'Dầu DO 0,001S-V'
        ]
        
        for row_index, row_values in enumerate(ws.iter_rows(min_row=11, values_only=True), start=11):
            quantity = _to_float(row_values[8] if len(row_values) > 8 else None)
            if quantity <= 0:
                continue # Bỏ qua các dòng không có số lượng hoặc số lượng <= 0 (bao gồm dòng tổng)

            fkey = _clean_string(row_values[24] if len(row_values) > 24 else None)
            item_name = _clean_string(row_values[6] if len(row_values) > 6 else None)
            
            invoice_data = {
                'fkey': fkey,
                'item_name': item_name,
                'quantity': quantity,
                'total_amount': _to_float(row_values[16] if len(row_values) > 16 else None),
                'invoice_number': _clean_string(row_values[19] if len(row_values) > 19 else None),
                'invoice_date_raw': row_values[20] if len(row_values) > 20 else None,
                'invoice_symbol_hddt': _clean_string(row_values[18] if len(row_values) > 18 else None), # Lấy ký hiệu hóa đơn từ cột S (index 18)
                'source_row': row_index
            }

            if fkey and fkey.upper().startswith('POS'):
                pos_invoices.append(invoice_data)
            else:
                if item_name in PETROLEUM_PRODUCTS:
                    direct_petroleum_invoices.append(invoice_data)
                else:
                    other_invoices.append(invoice_data)

        return {
            'pos_invoices': pos_invoices,
            'direct_petroleum_invoices': direct_petroleum_invoices,
            'other_invoices': other_invoices
        }
    except Exception as e:
        raise ValueError(f"Lỗi khi đọc file Bảng kê HĐĐT: {e}")

def _parse_log_bom_file(log_bom_bytes):
    """Phân tích dữ liệu từ file Log Bơm (POS)."""
    try:
        wb = load_workbook(io.BytesIO(log_bom_bytes), data_only=True)
        ws = wb.active
        
        pump_logs = []
        # Các loại giao dịch cần xuất hóa đơn
        VALID_TRANSACTION_TYPES = ['bán lẻ', 'hợp đồng', 'khuyến mãi', 'trả trước']
        
        for row_index, row_values in enumerate(ws.iter_rows(min_row=10, values_only=True), start=10):
            transaction_type = _clean_string(row_values[7] if len(row_values) > 7 else None)
            
            # Kiểm tra xem loại giao dịch có nằm trong danh sách hợp lệ không
            if transaction_type.lower() not in VALID_TRANSACTION_TYPES:
                continue

            fkey = _clean_string(row_values[14] if len(row_values) > 14 else None)
            if not fkey:
                raise ValueError(
                    f"Lỗi nghiêm trọng tại dòng {row_index} của file Log Bơm: "
                    f"Giao dịch '{transaction_type}' bắt buộc phải có mã H.Đơn (FKEY) ở cột O nhưng lại bị trống."
                )

            item_name = _clean_string(row_values[3] if len(row_values) > 3 else None)
            quantity = _to_float(row_values[4] if len(row_values) > 4 else None)
            total_amount = _to_float(row_values[6] if len(row_values) > 6 else None)

            pump_logs.append({
                'fkey': fkey,
                'item_name': item_name,
                'quantity': quantity,
                'total_amount': total_amount,
                'source_row': row_index
            })
            
        if not pump_logs:
            raise ValueError("Không tìm thấy giao dịch nào cần xuất hóa đơn trong file Log Bơm.")
        return pump_logs
    except Exception as e:
        raise ValueError(f"Lỗi khi đọc file Log Bơm: {e}")


def perform_reconciliation(log_bom_bytes, hddt_bytes, selected_chxd_name, invoice_symbol_from_config):
    """
    Thực hiện đối soát dữ liệu giữa file Log Bơm (POS) và file Bảng kê HĐĐT.
    Bổ sung bước xác thực CHXD và ký hiệu hóa đơn.
    """
    try:
        # --- BƯỚC XÁC THỰC CHXD TỪ FILE LOG BƠM (POS) ---
        log_wb = load_workbook(io.BytesIO(log_bom_bytes), data_only=True)
        log_ws = log_wb.active
        
        # Đọc ô A2 (có thể là merged cell A-B-C-D-E2)
        pos_chxd_cell_value = log_ws['A2'].value
        if pos_chxd_cell_value:
            # Loại bỏ "CHXD " và làm sạch chuỗi để so sánh
            pos_chxd_name_extracted = _clean_string(str(pos_chxd_cell_value).replace("CHXD ", ""))
            if pos_chxd_name_extracted.lower() != selected_chxd_name.lower():
                raise ValueError("Bảng kê log bơm không phải của cửa hàng bạn chọn.")
        else:
            raise ValueError("Không tìm thấy thông tin CHXD trong file Log Bơm (ô A2 trống).")

        # --- BƯỚC XÁC THỰC KÝ HIỆU HÓA ĐƠN TỪ FILE HĐĐT ---
        hddt_wb = load_workbook(io.BytesIO(hddt_bytes), data_only=True)
        hddt_ws = hddt_wb.active
        
        # Lấy 6 ký tự cuối của ký hiệu hóa đơn từ file cấu hình
        # Đảm bảo ký hiệu từ config đủ dài để cắt
        if len(invoice_symbol_from_config) < 6:
            raise ValueError(f"Ký hiệu hóa đơn trong file cấu hình Data_HDDT.xlsx ('{invoice_symbol_from_config}') quá ngắn để xác thực.")
        expected_invoice_symbol_suffix = invoice_symbol_from_config[-6:].upper()

        has_at_least_one_valid_invoice_for_symbol_check = False

        # Duyệt qua cột S (index 18) từ dòng 11 để kiểm tra ký hiệu hóa đơn
        for row_index, row_values in enumerate(hddt_ws.iter_rows(min_row=11, values_only=True), start=11):
            # Kiểm tra xem dòng này có phải là một hóa đơn thực tế (có số lượng > 0) không
            # Cột I (index 8) là số lượng
            quantity_val = _to_float(row_values[8] if len(row_values) > 8 else None)
            
            # Nếu số lượng <= 0, đây có thể là dòng tiêu đề, chân trang, hoặc dòng tổng cộng. Bỏ qua xác thực ký hiệu cho các dòng này.
            if quantity_val <= 0:
                continue

            # Nếu đến đây, đây là một dòng hóa đơn hợp lệ (có số lượng > 0)
            has_at_least_one_valid_invoice_for_symbol_check = True

            # Thực hiện xác thực ký hiệu hóa đơn
            if len(row_values) > 18 and row_values[18] is not None: # Cột S (index 18)
                actual_invoice_symbol_hddt = _clean_string(row_values[18])
                if len(actual_invoice_symbol_hddt) >= 6:
                    if actual_invoice_symbol_hddt[-6:].upper() != expected_invoice_symbol_suffix:
                        raise ValueError("Bảng kê hddt không phải của cửa hàng bạn chọn.")
                else:
                    # Dòng hóa đơn hợp lệ nhưng ký hiệu quá ngắn
                    raise ValueError(f"Ký hiệu hóa đơn tại dòng {row_index} của bảng kê HDDT quá ngắn để xác thực.")
            else:
                # Dòng hóa đơn hợp lệ nhưng thiếu ký hiệu hóa đơn
                raise ValueError(f"Hóa đơn tại dòng {row_index} của bảng kê HDDT thiếu ký hiệu hóa đơn (cột S).")
        
        # Sau khi kiểm tra tất cả các dòng, nếu không tìm thấy bất kỳ dòng hóa đơn hợp lệ nào để xác thực ký hiệu.
        if not has_at_least_one_valid_invoice_for_symbol_check:
            raise ValueError("Không tìm thấy hóa đơn hợp lệ nào trong file Bảng kê HDDT để xác thực ký hiệu.")

        # Nếu các bước xác thực thành công, tiếp tục xử lý đối soát
        parsed_hddt_data = _parse_hddt_file(hddt_bytes)
        hddt_invoices = parsed_hddt_data['pos_invoices']
        log_bom_data = _parse_log_bom_file(log_bom_bytes)

        if not hddt_invoices:
             raise ValueError("Không tìm thấy hóa đơn nào có FKEY bắt đầu bằng 'POS' để tiến hành đối soát.")

        log_map = {log['fkey']: log for log in log_bom_data}
        hddt_map = {inv['fkey']: inv for inv in hddt_invoices}

        log_fkeys = set(log_map.keys())
        hddt_fkeys = set(hddt_map.keys())

        missing_invoices_fkeys = sorted(list(log_fkeys - hddt_fkeys))
        extra_invoices_fkeys = sorted(list(hddt_fkeys - log_fkeys))
        common_fkeys = sorted(list(log_fkeys.intersection(hddt_fkeys)))

        quantity_mismatches = []
        amount_mismatches = []

        for fkey in common_fkeys:
            log = log_map[fkey]
            inv = hddt_map[fkey]
            
            inv_date = _excel_date_to_datetime(inv['invoice_date_raw'])
            inv_date_str = inv_date.strftime('%d/%m/%Y') if inv_date else 'N/A'
            
            mismatch_info = {
                'fkey': fkey,
                'invoice_number': inv.get('invoice_number', 'N/A'),
                'invoice_date': inv_date_str
            }

            if abs(log['quantity'] - inv['quantity']) > 0.001:
                quantity_mismatches.append(mismatch_info)

            if abs(log['total_amount'] - inv['total_amount']) > 1:
                amount_mismatches.append(mismatch_info)
                
        item_summary = defaultdict(lambda: {'quantity': {'pos': 0, 'hddt': 0}, 'amount': {'pos': 0, 'hddt': 0}})
        for log in log_bom_data:
            item_summary[log['item_name']]['quantity']['pos'] += log['quantity']
            item_summary[log['item_name']]['amount']['pos'] += log['total_amount']
        for inv in hddt_invoices:
            item_summary[inv['item_name']]['quantity']['hddt'] += inv['quantity']
            item_summary[inv['item_name']]['amount']['hddt'] += inv['total_amount']
            
        final_item_summary = {}
        for name, data in item_summary.items():
            qty_diff = data['quantity']['pos'] - data['quantity']['hddt']
            amt_diff = data['amount']['pos'] - data['amount']['hddt']
            final_item_summary[name] = {
                'quantity': {'pos': _format_number(data['quantity']['pos']), 'hddt': _format_number(data['quantity']['hddt']), 'difference': _format_number(qty_diff), 'is_match': abs(qty_diff) < 0.001},
                'amount': {'pos': _format_number(data['amount']['pos']), 'hddt': _format_number(data['amount']['hddt']), 'difference': _format_number(amt_diff), 'is_match': abs(amt_diff) < 1}
            }
            
        count_diff = len(log_bom_data) - len(hddt_invoices)
        reconciliation_data = {
            'summary': {
                'pos_count': len(log_bom_data), 
                'hddt_count': len(hddt_invoices), 
                'is_match': count_diff == 0 and not missing_invoices_fkeys and not extra_invoices_fkeys,
                'difference': count_diff,
                'missing_fkeys': missing_invoices_fkeys,
                'extra_fkeys': extra_invoices_fkeys
            },
            'detailed_mismatches': {
                'quantities': quantity_mismatches,
                'amounts': amount_mismatches
            },
            'item_comparison': final_item_summary,
            'non_pos_invoices': {
                'direct_petroleum': parsed_hddt_data['direct_petroleum_invoices'],
                'others': parsed_hddt_data['other_invoices']
            }
        }
        
        return reconciliation_data

    except Exception as e:
        print(f"Lỗi trong quá trình đối soát: {e}")
        raise e
