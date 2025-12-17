# ==========================================
# CÀI ĐẶT THƯ VIỆN CẦN THIẾT:
# pip install streamlit pandas python-docx docx2pdf openpyxl pywin32
# ==========================================

import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
import os
import copy
import tempfile
import zipfile
import pythoncom
import traceback  # Thư viện để in chi tiết lỗi

# ================== CẤU HÌNH TRANG ==================
st.set_page_config(page_title="Tool Điền Đơn", layout="centered")

# Kích thước ảnh mong muốn (300x400 px quy đổi sang Inches)
IMG_W = Inches(90 / 96)
IMG_H = Inches(120 / 96)

# ================== HÀM XỬ LÝ DATE & IMAGE (AN TOÀN) ==================
def process_value_and_replace(paragraph_or_cell, placeholder, raw_value, img_w, img_h):
    final_text = str(raw_value).strip()
    
    # --- 1. Xử lý DATE ---
    # Mục tiêu: Đưa về dạng dd/mm/yyyy bất kể Excel đang lưu dạng gì
    try:
        if "-" in final_text or "/" in final_text:
            # Kiểm tra xem có phải dạng ISO (yyyy-mm-dd) do Excel tự sinh ra không
            # Nếu 4 ký tự đầu là số (VD: 2025) -> Là Năm -> dayfirst=False
            if len(final_text) >= 4 and final_text[0:4].isdigit() and "-" in final_text:
                dt = pd.to_datetime(final_text)
            else:
                # Ngược lại ưu tiên ngày trước tháng (VN)
                dt = pd.to_datetime(final_text, dayfirst=True)
            
            final_text = dt.strftime("%d/%m/%Y")
    except:
        # Nếu không phải ngày tháng hợp lệ thì giữ nguyên text
        pass

    # --- 2. Xử lý ẢNH ---
    # Kiểm tra xem text có phải đường dẫn file ảnh hợp lệ không
    if os.path.isfile(final_text) and final_text.lower().endswith((".png", ".jpg", ".jpeg")):
        try:
            # Xóa chữ giữ chỗ (placeholder) trước
            if hasattr(paragraph_or_cell, 'text'):
                if placeholder in paragraph_or_cell.text:
                    paragraph_or_cell.text = paragraph_or_cell.text.replace(placeholder, "")
            
            # Chèn ảnh vào
            run = None
            if hasattr(paragraph_or_cell, 'add_run'):
                run = paragraph_or_cell.add_run()
            elif hasattr(paragraph_or_cell, 'paragraphs'):
                # Trường hợp là Cell của bảng
                if len(paragraph_or_cell.paragraphs) > 0:
                    run = paragraph_or_cell.paragraphs[0].add_run()
                else:
                    # Nếu cell trống trơn chưa có paragraph nào
                    paragraph_or_cell.add_paragraph()
                    run = paragraph_or_cell.paragraphs[0].add_run()
            
            if run:
                run.add_picture(final_text, width=img_w, height=img_h)
                
        except Exception as e:
            # QUAN TRỌNG: Nếu ảnh lỗi (corrupt file), in lỗi ra console và điền text báo lỗi vào Word
            # Giúp chương trình không bị dừng đột ngột
            print(f"❌ Lỗi chèn ảnh {final_text}: {e}")
            if hasattr(paragraph_or_cell, 'add_run'):
                paragraph_or_cell.add_run(f" [LỖI FILE ẢNH] ")

    # --- 3. Xử lý TEXT THƯỜNG ---
    else:
        if hasattr(paragraph_or_cell, 'text'):
            paragraph_or_cell.text = paragraph_or_cell.text.replace(placeholder, final_text)

# ================== HÀM GIẢI NÉN VÀ MAP ẢNH TỪ ZIP ==================
def extract_images_and_map(zip_file_obj, temp_folder):
    """Giải nén zip và tạo dict {'tên file viết thường': 'đường dẫn full'}"""
    image_map = {}
    with zipfile.ZipFile(zip_file_obj, 'r') as zip_ref:
        zip_ref.extractall(temp_folder)
        for root, dirs, files in os.walk(temp_folder):
            for file in files:
                if file.lower().endswith(('.png', '.jpg', '.jpeg')):
                    full_path = os.path.join(root, file)
                    name_no_ext = os.path.splitext(file)[0]
                    image_map[name_no_ext.lower().strip()] = full_path
    return image_map

# ================== HÀM CHÍNH (CORE LOGIC) ==================
def generate_documents(excel_file, word_file, zip_file, match_col, img_placeholder, export_pdf):
    # Tạo thư mục tạm để chứa file kết quả và file giải nén
    main_temp_dir = tempfile.mkdtemp()
    
    try:
        # --- BƯỚC 1: Đọc Excel ---
        df = pd.read_excel(excel_file, dtype=str, keep_default_na=False)
        # Lọc bỏ các dòng trống hoàn toàn (rác Excel)
        df['temp_check'] = df.apply(lambda x: ''.join(x.values.astype(str)).strip(), axis=1)
        df = df[df['temp_check'] != '']
        df = df.drop(columns=['temp_check'])
        df.reset_index(drop=True, inplace=True)

        if df.empty:
            st.error("File Excel không có dữ liệu!")
            return None, None

        # --- BƯỚC 2: Xử lý Zip ảnh (nếu có) ---
        if zip_file and match_col and img_placeholder:
            images_temp_dir = os.path.join(main_temp_dir, "images_extracted")
            os.makedirs(images_temp_dir, exist_ok=True)
            
            # Giải nén
            img_map = extract_images_and_map(zip_file, images_temp_dir)
            
            # Hàm tìm đường dẫn ảnh dựa trên tên
            def get_image_path(person_name):
                clean_name = str(person_name).lower().strip()
                return img_map.get(clean_name, "") # Trả về path hoặc rỗng

            # Tạo cột mới trong Dataframe chứa đường dẫn ảnh
            df[img_placeholder] = df[match_col].apply(get_image_path)
            
            # Thống kê
            found_count = len(df[df[img_placeholder] != ""])
            st.info(f"Đã tìm thấy {found_count}/{len(df)} ảnh khớp tên trong file Zip.")

        # --- BƯỚC 3: Xử lý Word Template ---
        base_doc = Document(word_file)
        template_blocks = []
        for block in base_doc.element.body:
            template_blocks.append(copy.deepcopy(block))
        base_doc.element.body.clear()

        # Thanh tiến trình
        progress_bar = st.progress(0)
        status_text = st.empty()
        total_rows = len(df)

        # --- BƯỚC 4: Vòng lặp điền đơn ---
        for index, data_row in df.iterrows():
            progress_bar.progress((index + 1) / total_rows)
            status_text.text(f"Đang xử lý hồ sơ {index + 1}/{total_rows}...")
            
            # 4.1 Chèn block mẫu
            for block in template_blocks:
                base_doc.element.body.append(copy.deepcopy(block))

            # 4.2 Điền dữ liệu Paragraph
            for paragraph in base_doc.paragraphs:
                for col in df.columns:
                    placeholder = f"{{{{{col}}}}}"
                    if placeholder in paragraph.text:
                        raw_val = str(data_row[col])
                        process_value_and_replace(paragraph, placeholder, raw_val, IMG_W, IMG_H)

            # 4.3 Điền dữ liệu Table
            for table in base_doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for col in df.columns:
                            placeholder = f"{{{{{col}}}}}"
                            if placeholder in cell.text:
                                raw_val = str(data_row[col])
                                process_value_and_replace(cell, placeholder, raw_val, IMG_W, IMG_H)

            # 4.4 Ngắt trang (Trừ người cuối cùng)
            if index < total_rows - 1:
                # Dùng cách thêm run break để tránh lỗi trang trắng
                if base_doc.paragraphs:
                    base_doc.paragraphs[-1].add_run().add_break(WD_BREAK.PAGE)
                else:
                    base_doc.add_page_break()

        # --- BƯỚC 5: Lưu File ---
        output_word_path = os.path.join(main_temp_dir, "KetQua_DonDaDien.docx")
        base_doc.save(output_word_path)

        output_pdf_path = None
        if export_pdf:
            status_text.text("Đang chuyển đổi PDF (Vui lòng đợi)...")
            try:
                # Cần thiết cho thread streamlit
                pythoncom.CoInitialize() 
                output_pdf_path = os.path.join(main_temp_dir, "KetQua_DonDaDien.pdf")
                convert(output_word_path, output_pdf_path)
            except Exception as e:
                st.warning(f"Không thể xuất PDF. Lỗi: {e}")

        status_text.text("Hoàn tất!")
        progress_bar.empty()
        return output_word_path, output_pdf_path

    except Exception as e:
        st.error("Gặp lỗi nghiêm trọng trong quá trình xử lý!")
        st.error(f"Chi tiết: {e}")
        # In traceback để debug lỗi khó hiểu
        st.code(traceback.format_exc())
        return None, None

# ================== GIAO DIỆN STREAMLIT (UI) ==================

st.title("📄 Tool Điền Đơn Tự Động Pro")
st.markdown("---")

# Cột trái phải
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. File Excel dữ liệu (.xlsx)", type="xlsx")
with col2:
    uploaded_word = st.file_uploader("2. File Word mẫu (.docx)", type="docx")

st.markdown("---")
st.subheader("3. Cấu hình Ảnh thẻ (Tùy chọn)")
use_image_zip = st.checkbox("Tôi muốn upload folder ảnh nén (.zip) để điền tự động", value=False)

uploaded_zip = None
match_col = None
img_placeholder_name = ""

if use_image_zip:
    uploaded_zip = st.file_uploader("Upload file .zip chứa ảnh", type="zip")
    
    if uploaded_excel:
        try:
            # Đọc thử header Excel để cho user chọn cột
            df_preview = pd.read_excel(uploaded_excel, nrows=0)
            cols = df_preview.columns.tolist()
            
            c1, c2 = st.columns(2)
            with c1:
                match_col = st.selectbox("Cột tên trong Excel dùng để so khớp:", cols)
                st.caption("Ví dụ: Cột 'HO_TEN'. Code sẽ tìm file ảnh có tên giống hệt nội dung cột này.")
            with c2:
                img_placeholder_name = st.text_input("Mã giữ chỗ ảnh trong Word:", value="ANH_THE")
                st.caption("Ví dụ: Trong Word bạn để `{{ANH_THE}}`, thì điền vào đây là `ANH_THE`.")
        except:
            pass

st.markdown("---")
need_pdf = st.checkbox("Xuất thêm file PDF (Yêu cầu Server có cài MS Word)", value=False)

# Nút chạy
if st.button("🚀 BẮT ĐẦU XỬ LÝ", type="primary"):
    if uploaded_excel and uploaded_word:
        # Check logic zip
        if use_image_zip and not uploaded_zip:
            st.warning("Bạn chọn dùng ảnh Zip nhưng chưa upload file Zip!")
        else:
            with st.spinner("Đang xử lý dữ liệu..."):
                word_out, pdf_out = generate_documents(
                    uploaded_excel, 
                    uploaded_word, 
                    uploaded_zip, 
                    match_col, 
                    img_placeholder_name, 
                    need_pdf
                )
                
                if word_out:
                    st.success("✅ Xử lý thành công!")
                    
                    d1, d2 = st.columns(2)
                    with open(word_out, "rb") as f:
                        d1.download_button(
                            label="📥 Tải file Word (.docx)",
                            data=f,
                            file_name="KetQua_DonDaDien.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    
                    if pdf_out and os.path.exists(pdf_out):
                        with open(pdf_out, "rb") as f:
                            d2.download_button(
                                label="📥 Tải file PDF (.pdf)",
                                data=f,
                                file_name="KetQua_DonDaDien.pdf",
                                mime="application/pdf"
                            )
    else:

        st.error("Vui lòng upload đủ file Excel và Word mẫu!")
