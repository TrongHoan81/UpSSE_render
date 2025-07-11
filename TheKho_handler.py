import io
import os
import json
import base64
from datetime import datetime
import google.generativeai as genai
from PIL import Image
import fitz # PyMuPDF
from openpyxl import Workbook

# Cấu hình Gemini API Key
# KHÔNG NÊN HARDCODE API KEY TRONG MÔI TRƯỜNG SẢN XUẤT!
# Thay vào đó, hãy đặt biến môi trường GEMINI_API_KEY trên Render.
api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    # Đây là fallback cho môi trường dev cục bộ nếu bạn chưa đặt biến môi trường
    # Trong môi trường Render, biến môi trường sẽ được tự động cung cấp.
    print("Cảnh báo: Không tìm thấy GEMINI_API_KEY trong biến môi trường. Vui lòng đặt biến này.")
    # Bạn có thể tạm thời đặt API key của mình vào đây nếu muốn thử nghiệm cục bộ
    # api_key = "ĐIỀN API VÀ ĐÂY" # <--- API KEY CỦA BẠN ĐÃ ĐƯỢC TÍCH HỢP CỨNG Ở ĐÂY
                                                      # NHỚ XÓA TRƯỚC KHI TRIỂN KHAI LÊN RENDER!

if api_key:
    genai.configure(api_key=api_key)
    # CHỈNH SỬA: Chuyển từ 'gemini-pro-vision' sang 'gemini-1.5-flash'
    gemini_model = genai.GenerativeModel('gemini-1.5-flash')
else:
    gemini_model = None
    print("Gemini API không được cấu hình. Chức năng Thẻ kho sẽ không hoạt động.")


def _convert_pdf_to_images(pdf_bytes):
    """
    Chuyển đổi mỗi trang của file PDF thành một đối tượng PIL Image.
    """
    images = []
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            # Render trang thành pixel map (PNG format)
            pix = page.get_pixmap()
            img_bytes = pix.tobytes("png")
            # Chuyển đổi bytes thành PIL Image
            images.append(Image.open(io.BytesIO(img_bytes)))
        doc.close()
    except Exception as e:
        raise ValueError(f"Lỗi khi chuyển đổi PDF sang ảnh: {e}")
    return images

def _extract_data_from_image_with_gemini(image_content):
    """
    Gửi hình ảnh tới Gemini API để trích xuất dữ liệu thẻ kho.
    image_content có thể là PIL Image object hoặc bytes của ảnh.
    """
    if gemini_model is None:
        raise ValueError("Gemini API chưa được cấu hình. Vui lòng kiểm tra GEMINI_API_KEY.")

    # Prompt để trích xuất thông tin dưới dạng JSON
    prompt_text = """
    Extract the following information from the provided document (Phiếu xuất kho kiêm vận chuyển nội bộ) and return it as a JSON object.
    Ensure all numerical values are parsed as numbers (integers or floats).
    If a field is not found or cannot be reliably extracted, set its value to null.
    The 'Ngày tháng' field should be in 'DD/MM/YYYY' format if present.

    JSON Keys:
    {
      "ky_hieu": "Ký hiệu (Serial)",
      "so": "Số",
      "ngay_thang": "Ngày tháng (Ngày xuất/Ngày)",
      "ten_chxd": "Tên CHXD (trường 'của' trên ảnh)",
      "ten_vat_tu": "Tên vật tư, hàng hóa",
      "don_vi_tinh": "Đơn vị tính",
      "so_luong": "Số lượng",
      "nhiet_do_thuc_te": "Nhiệt độ thực tế",
      "ty_trong": "Tỷ trọng",
      "he_so_vcf": "Hệ số VCF",
      "so_luong_quy_ve_15_do_c": "Số lượng xuất quy về 15 độ C"
    }
    """
    
    try:
        # Gửi prompt và hình ảnh tới Gemini
        # Gemini API expects image as PIL.Image.Image or raw bytes for multimodal input
        response = gemini_model.generate_content([prompt_text, image_content])
        
        # Xử lý phản hồi từ Gemini
        # Đôi khi Gemini có thể thêm các ký tự markdown như ```json hoặc ```
        json_string = response.text.strip()
        if json_string.startswith("```json"):
            json_string = json_string[len("```json"):].strip()
        if json_string.endswith("```"):
            json_string = json_string[:-len("```")].strip()

        extracted_data = json.loads(json_string)
        return extracted_data
    except json.JSONDecodeError as e:
        # LOG RA PHẢN HỒI GỐC TỪ GEMINI KHI CÓ LỖI JSON
        print(f"Lỗi JSONDecodeError: {e}")
        print(f"Phản hồi thô từ Gemini: \n{response.text}")
        raise ValueError(f"Gemini trả về định dạng không hợp lệ. Chi tiết: {e}")
    except Exception as e:
        raise ValueError(f"Lỗi khi gọi Gemini API hoặc xử lý phản hồi: {e}")

def _validate_and_normalize_data(data, filename=""):
    """
    Kiểm tra và chuẩn hóa dữ liệu trích xuất.
    Đảm bảo các trường quan trọng có giá trị và định dạng đúng.
    """
    # Các trường bắt buộc và tên hiển thị của chúng
    required_fields_map = {
        "ky_hieu": "Ký hiệu (Serial)",
        "so": "Số",
        "ngay_thang": "Ngày tháng",
        "ten_chxd": "Tên CHXD",
        "ten_vat_tu": "Tên vật tư, hàng hóa",
        "so_luong": "Số lượng"
    }

    for field, display_name in required_fields_map.items():
        if data.get(field) is None or str(data.get(field)).strip() == "":
            raise ValueError(f"File '{filename}': Không trích xuất được thông tin quan trọng: '{display_name}'")
    
    # Chuẩn hóa Ngày tháng
    ngay_thang_str = str(data.get("ngay_thang", "")).strip()
    try:
        # Thử các định dạng ngày phổ biến
        if '/' in ngay_thang_str:
            data['ngay_thang_dt'] = datetime.strptime(ngay_thang_str, '%d/%m/%Y')
        elif '-' in ngay_thang_str:
            data['ngay_thang_dt'] = datetime.strptime(ngay_thang_str, '%d-%m-%Y')
        else:
            # Nếu không có dấu phân cách, thử định dạng YYYYMMDD hoặc DDMMYYYY
            if len(ngay_thang_str) == 8:
                try:
                    data['ngay_thang_dt'] = datetime.strptime(ngay_thang_str, '%Y%m%d')
                except ValueError:
                    data['ngay_thang_dt'] = datetime.strptime(ngay_thang_str, '%d%m%Y')
            else:
                raise ValueError("Định dạng ngày tháng không xác định.")
    except (ValueError, TypeError):
        raise ValueError(f"File '{filename}': Không thể chuẩn hóa định dạng ngày tháng '{ngay_thang_str}'. Vui lòng kiểm tra lại.")

    # Chuẩn hóa các giá trị số
    # Các trường là số nguyên (Số lượng, Số lượng quy về 15 độ C)
    integer_fields = ["so_luong", "so_luong_quy_ve_15_do_c"]
    # Các trường là số thập phân (Nhiệt độ, Tỷ trọng, Hệ số VCF)
    float_fields = ["nhiet_do_thuc_te", "ty_trong", "he_so_vcf"]

    all_numerical_fields = integer_fields + float_fields

    for field in all_numerical_fields:
        value = data.get(field)
        if value is not None:
            s_value = str(value).strip()
            if not s_value:
                data[field] = None
                continue

            try:
                # Xóa khoảng trắng nếu có
                s_value = s_value.replace(' ', '')

                if field in integer_fields:
                    # Đối với số nguyên: Xóa tất cả dấu chấm và dấu phẩy
                    s_value = s_value.replace('.', '') 
                    s_value = s_value.replace(',', '') 
                    data[field] = int(s_value) # Chuyển thẳng sang int
                elif field in float_fields:
                    # Đối với số thập phân:
                    # Ưu tiên dấu phẩy là dấu thập phân.
                    if ',' in s_value:
                        s_value = s_value.replace('.', '') # Xóa dấu chấm nếu có (tránh trường hợp 1.234,56)
                        s_value = s_value.replace(',', '.') # Thay thế dấu phẩy bằng dấu chấm
                    elif '.' in s_value:
                        # Nếu không có dấu phẩy nhưng có dấu chấm, coi dấu chấm là dấu thập phân
                        pass 
                    data[field] = float(s_value)
            except (ValueError, TypeError):
                print(f"Cảnh báo: File '{filename}', trường '{field}' có giá trị '{value}' không phải là số hợp lệ. Đặt về null.")
                data[field] = None # Đặt là None nếu không phải số

    return data

def _create_excel_buffer(extracted_records):
    """
    Tạo file Excel từ danh sách các bản ghi đã trích xuất.
    """
    if not extracted_records:
        return None

    wb = Workbook()
    ws = wb.active
    ws.title = "The_kho_tu_dong"

    # Định nghĩa headers cho Excel
    headers = [
        "Ký hiệu (Serial)", "Số", "Ngày tháng", "Tên CHXD",
        "Tên vật tư, hàng hóa", "Đơn vị tính", "Số lượng", "Nhiệt độ thực tế",
        "Tỷ trọng", "Hệ số VCF", "Số lượng xuất quy về 15 độ C"
    ]
    ws.append(headers)

    for record in extracted_records:
        row = [
            record.get("ky_hieu"),
            record.get("so"),
            record.get("ngay_thang_dt").strftime('%d/%m/%Y') if record.get("ngay_thang_dt") else record.get("ngay_thang"), # Định dạng lại ngày tháng
            record.get("ten_chxd"),
            record.get("ten_vat_tu"),
            record.get("don_vi_tinh"),
            record.get("so_luong"),
            record.get("nhiet_do_thuc_te"),
            record.get("ty_trong"),
            record.get("he_so_vcf"),
            record.get("so_luong_quy_ve_15_do_c")
        ]
        ws.append(row)

    # Tự động điều chỉnh độ rộng cột (tùy chọn)
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter # Get the column name
        for cell in col:
            try: # Necessary to avoid error on empty cells
                if cell.value is not None:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer

def process_stock_card_data(uploaded_files, selected_chxd):
    """
    Hàm điều phối chính để xử lý nhiều file ảnh/PDF và tạo Excel.
    """
    all_extracted_data = []
    processing_errors = [] # Danh sách để lưu trữ các lỗi cụ thể

    for file in uploaded_files:
        file_content_bytes = file.read()
        file_mimetype = file.mimetype
        filename = file.filename # Lấy tên file để đưa vào thông báo lỗi

        images_to_process = []
        if file_mimetype.startswith('image/'):
            images_to_process.append(Image.open(io.BytesIO(file_content_bytes)))
        elif file_mimetype == 'application/pdf':
            try:
                images_to_process.extend(_convert_pdf_to_images(file_content_bytes))
            except ValueError as ve:
                processing_errors.append(f"File '{filename}': {ve}")
                continue
        else:
            processing_errors.append(f"File '{filename}': Định dạng không hỗ trợ ({file_mimetype}). Chỉ chấp nhận ảnh hoặc PDF.")
            continue

        if not images_to_process:
            processing_errors.append(f"File '{filename}': Không thể trích xuất hình ảnh từ file.")
            continue

        for i, img in enumerate(images_to_process):
            try:
                # Thêm chỉ số trang nếu là PDF
                current_filename_info = f"{filename} (Trang {i+1})" if file_mimetype == 'application/pdf' else filename
                extracted_record = _extract_data_from_image_with_gemini(img)
                
                # Thêm CHXD đã chọn vào dữ liệu trích xuất
                extracted_record['ten_chxd'] = selected_chxd
                
                validated_data = _validate_and_normalize_data(extracted_record, filename=current_filename_info)
                all_extracted_data.append(validated_data)
            except ValueError as ve:
                # Nếu có lỗi validation hoặc trích xuất, thông báo nhưng không dừng toàn bộ quá trình
                processing_errors.append(f"Lỗi xử lý '{current_filename_info}': {ve}")
                continue
            except Exception as e:
                processing_errors.append(f"Lỗi không xác định khi xử lý '{current_filename_info}': {e}")
                continue

    if not all_extracted_data:
        # Nếu không có dữ liệu nào được trích xuất thành công
        if processing_errors:
            # Gộp tất cả các lỗi chi tiết lại
            raise ValueError("Không có dữ liệu hợp lệ nào được trích xuất từ các file đã tải lên. Chi tiết lỗi:\n" + "\n".join(processing_errors))
        else:
            raise ValueError("Không có dữ liệu hợp lệ nào được trích xuất từ các file đã tải lên. Vui lòng kiểm tra định dạng file và nội dung.")

    # Sắp xếp dữ liệu: theo Ngày tháng, sau đó theo Số (seri)
    # Chuyển đổi 'so' thành chuỗi để sắp xếp đúng cho cả số và ký tự
    all_extracted_data.sort(key=lambda x: (x.get('ngay_thang_dt', datetime.min), str(x.get('so', ''))))

    return _create_excel_buffer(all_extracted_data)

