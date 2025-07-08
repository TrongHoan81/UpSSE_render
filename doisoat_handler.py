import io
import re
from collections import defaultdict
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook

# --- CÁC HÀM TIỆN ÍCH ---
def _clean_string(s):
    """Làm sạch chuỗi, loại bỏ khoảng trắng thừa và dấu nháy đơn ở đầu."""
    if s is None: return ""
    cleaned_s = str(s).strip()
    if cleaned_s.startswith("'"): cleaned_s = cleaned_s[1:]
    return re.sub(r'\s+', ' ', cleaned_s)

def _to_float(value):
    """Chuyển đổi giá trị sang kiểu float, xử lý lỗi nếu có."""
    if value is None: return 0.0
    try:
        return float(str(value).replace(',', '').strip())
    except (ValueError, TypeError): return 0.0

def _format_number(num):
    """Định dạng số với dấu phẩy ngăn cách hàng nghìn."""
    try:
        return f"{num:,.2f}"
    except (ValueError, TypeError):
        return "0.00"

# --- CÁC HÀM PHÂN TÍCH FILE ---

def _parse_hddt_file(hddt_bytes):
    """
    Đọc file bảng kê HĐĐT, phân loại và trích xuất dữ liệu.
    """
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
                continue

            fkey = _clean_string(row_values[24] if len(row_values) > 24 else None)
            item_name = _clean_string(row_values[6] if len(row_values) > 6 else None)
            
            invoice_data = {
                'fkey': fkey,
                'item_name': item_name,
                'quantity': quantity,
                'total_amount': _to_float(row_values[16] if len(row_values) > 16 else None),
                'invoice_date_raw': row_values[20] if len(row_values) > 20 else None,
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
    """
    Đọc file log bơm, lọc các giao dịch bán hàng và trích xuất dữ liệu.
    """
    try:
        wb = load_workbook(io.BytesIO(log_bom_bytes), data_only=True)
        ws = wb.active
        
        pump_logs = []
        for row_index, row_values in enumerate(ws.iter_rows(min_row=10, values_only=True), start=10):
            transaction_type = _clean_string(row_values[7] if len(row_values) > 7 else None)
            if transaction_type.lower() not in ['bán lẻ', 'hợp đồng']:
                continue

            fkey = _clean_string(row_values[14] if len(row_values) > 14 else None)
            if not fkey:
                raise ValueError(
                    f"Lỗi nghiêm trọng tại dòng {row_index} của file Log Bơm: "
                    f"Giao dịch '{transaction_type}' bắt buộc phải có mã H.Đơn (FKEY) ở cột O nhưng lại bị trống."
                )

            time_str = row_values[1] if len(row_values) > 1 else None
            item_name = _clean_string(row_values[3] if len(row_values) > 3 else None)
            quantity = _to_float(row_values[4] if len(row_values) > 4 else None)
            total_amount = _to_float(row_values[6] if len(row_values) > 6 else None)

            transaction_time = None
            if isinstance(time_str, datetime):
                transaction_time = time_str
            elif isinstance(time_str, str):
                try:
                    transaction_time = datetime.strptime(time_str, '%d/%m/%Y %H:%M:%S')
                except ValueError:
                    try:
                        transaction_time = datetime.strptime(time_str, '%d/%m/%Y')
                    except ValueError:
                        pass

            pump_logs.append({
                'fkey': fkey,
                'item_name': item_name,
                'quantity': quantity,
                'total_amount': total_amount,
                'transaction_time': transaction_time,
                'source_row': row_index
            })
            
        if not pump_logs:
            raise ValueError("Không tìm thấy giao dịch 'Bán lẻ' hoặc 'Hợp đồng' nào trong file Log Bơm.")
        return pump_logs
    except Exception as e:
        raise ValueError(f"Lỗi khi đọc file Log Bơm: {e}")


def perform_reconciliation(log_bom_bytes, hddt_bytes, selected_chxd):
    """
    Thực hiện logic đối soát giữa file log bơm và file bảng kê HĐĐT.
    """
    try:
        # BƯỚC 1: Đọc và phân loại dữ liệu
        parsed_hddt_data = _parse_hddt_file(hddt_bytes)
        hddt_invoices = parsed_hddt_data['pos_invoices']
        log_bom_data = _parse_log_bom_file(log_bom_bytes)

        if not hddt_invoices:
             raise ValueError("Không tìm thấy hóa đơn nào có FKEY bắt đầu bằng 'POS' để tiến hành đối soát.")

        # BƯỚC 2: Chuẩn bị dữ liệu để so sánh
        log_fkeys = {log['fkey'] for log in log_bom_data}
        hddt_fkey_counts = defaultdict(int)
        hddt_fkeys_set = set()
        for inv in hddt_invoices:
            if inv['fkey']:
                hddt_fkey_counts[inv['fkey']] += 1
                hddt_fkeys_set.add(inv['fkey'])

        # BƯỚC 3: Tìm kiếm sự chênh lệch FKEY
        missing_invoices = sorted(list(log_fkeys - hddt_fkeys_set)) # Log có, HĐĐT không có
        extra_invoices_orphan = sorted(list(hddt_fkeys_set - log_fkeys)) # HĐĐT có, Log không có
        
        # Hóa đơn bị trùng FKEY (1 log xuất nhiều HĐ)
        duplicate_invoices = sorted([fkey for fkey, count in hddt_fkey_counts.items() if count > 1])
        
        # Tổng hợp các hóa đơn xuất thừa
        all_extra_invoices = sorted(list(set(extra_invoices_orphan + duplicate_invoices)))

        # BƯỚC 4: Tổng hợp và so sánh theo từng mặt hàng
        item_summary = defaultdict(lambda: {
            'quantity': {'pos': 0, 'hddt': 0},
            'amount': {'pos': 0, 'hddt': 0}
        })

        for log in log_bom_data:
            item_summary[log['item_name']]['quantity']['pos'] += log['quantity']
            item_summary[log['item_name']]['amount']['pos'] += log['total_amount']

        for inv in hddt_invoices:
            item_summary[inv['item_name']]['quantity']['hddt'] += inv['quantity']
            item_summary[inv['item_name']]['amount']['hddt'] += inv['total_amount']
            
        # Hoàn thiện dữ liệu so sánh theo mặt hàng
        final_item_summary = {}
        for name, data in item_summary.items():
            qty_diff = data['quantity']['pos'] - data['quantity']['hddt']
            amt_diff = data['amount']['pos'] - data['amount']['hddt']
            final_item_summary[name] = {
                'quantity': {
                    'pos': _format_number(data['quantity']['pos']),
                    'hddt': _format_number(data['quantity']['hddt']),
                    'difference': _format_number(qty_diff),
                    'is_match': abs(qty_diff) < 0.001 # Cho phép sai số nhỏ
                },
                'amount': {
                    'pos': _format_number(data['amount']['pos']),
                    'hddt': _format_number(data['amount']['hddt']),
                    'difference': _format_number(amt_diff),
                    'is_match': abs(amt_diff) < 1 # Cho phép sai số nhỏ (ví dụ 1 đồng)
                }
            }
            
        # BƯỚC 5: Tạo kết quả cuối cùng
        count_diff = len(log_bom_data) - len(hddt_invoices)
        reconciliation_data = {
            'count': {
                'pos': len(log_bom_data), 
                'hddt': len(hddt_invoices), 
                'is_match': count_diff == 0 and not missing_invoices and not all_extra_invoices,
                'difference': count_diff,
                'missing_pos': missing_invoices, # Log chưa xuất HĐ
                'extra_hddt': all_extra_invoices # HĐ thừa (mồ côi + trùng lặp)
            },
            'items': final_item_summary,
            'direct_petroleum_invoices': parsed_hddt_data['direct_petroleum_invoices'],
            'other_invoices': parsed_hddt_data['other_invoices']
        }
        
        return reconciliation_data

    except Exception as e:
        print(f"Lỗi trong quá trình đối soát: {e}")
        raise e
