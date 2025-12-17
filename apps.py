import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_BREAK
import os
import copy
import tempfile
import zipfile
import traceback 

# ================== CẤU HÌNH TRANG ==================
st.set_page_config(page_title="Tool Điền Đơn Online", layout="centered")

# Kích thước ảnh (Bạn có thể sửa lại số 90, 120 thành 300, 400 nếu muốn ảnh to hơn)
IMG_W = Inches(90 / 96)
IMG_H = Inches(120 / 96)

# ================== HÀM XỬ LÝ DATE & IMAGE ==================
def process_value_and_replace(paragraph_or_cell, placeholder, raw_value, img_w, img_h):
    final_text = str(raw_value).strip()
    
    # --- 1. Xử lý DATE ---
    try:
        if "-" in final_text or "/" in final_text:
            if len(final_text) >= 4 and final_text[0:4].isdigit() and "-" in final_text:
                dt = pd.to_datetime(final_text)
            else:
                dt = pd.to_datetime(final_text, dayfirst=True)
            final_text = dt.strftime("%d/%m/%Y")
    except:
        pass

    # --- 2. Xử lý ẢNH ---
    if os.path.isfile(final_text) and final_text.lower().endswith((".png", ".jpg", ".jpeg")):
        try:
            if hasattr(paragraph_or_cell, 'text'):
                if placeholder in paragraph_or_cell.text:
                    paragraph_or_cell.text = paragraph_or_cell.text.replace(placeholder, "")
            
            run = None
            if hasattr(paragraph_or_cell, 'add_run'):
                run = paragraph_or_cell.add_run()
            elif hasattr(paragraph_or_cell, 'paragraphs'):
                if len(paragraph_or_cell.paragraphs) > 0:
                    run = paragraph_or_cell.paragraphs[0].add_run()
                else:
                    paragraph_or_cell.add_paragraph()
                    run = paragraph_or_cell.paragraphs[0].add_run()
            
            if run:
                run.add_picture(final_text, width=img_w, height=img_h)
                
        except Exception as e:
            print(f"❌ Lỗi chèn ảnh {final_text}: {e}")
            if hasattr(paragraph_or_cell, 'add_run'):
                paragraph_or_cell.add_run(f" [LỖI ẢNH] ")

    # --- 3. Xử lý TEXT ---
    else:
        if hasattr(paragraph_or_cell, 'text'):
            paragraph_or_cell.text = paragraph_or_cell.text.replace(placeholder, final_text)

# ================== HÀM GIẢI NÉN ZIP ==================
def extract_images_and_map(zip_file_obj, temp_folder):
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

# ================== HÀM CHÍNH (Đã xóa PDF) ==================
def generate_documents(excel_file, word_file, zip_file, match_col, img_placeholder):
    main_temp_dir = tempfile.mkdtemp()
    
    try:
        # 1. Đọc Excel
        df = pd.read_excel(excel_file, dtype=str, keep_default_na=False)
        df['temp_check'] = df.apply(lambda x: ''.join(x.values.astype(str)).strip(), axis=1)
        df = df[df['temp_check'] != '']
        df = df.drop(columns=['temp_check'])
        df.reset_index(drop=True, inplace=True)

        if df.empty:
            st.error("File Excel không có dữ liệu!")
            return None

        # 2. Xử lý Zip ảnh
        if zip_file and match_col and img_placeholder:
            images_temp_dir = os.path.join(main_temp_dir, "images_extracted")
            os.makedirs(images_temp_dir, exist_ok=True)
            img_map = extract_images_and_map(zip_file, images_temp_dir)
            
            def get_image_path(person_name):
                clean_name = str(person_name).lower().strip()
                return img_map.get(clean_name, "")

            df[img_placeholder] = df[match_col].apply(get_image_path)
            found_count = len(df[df[img_placeholder] != ""])
            st.info(f"Đã tìm thấy {found_count}/{len(df)} ảnh khớp tên.")

        # 3. Chuẩn bị Word
        base_doc = Document(word_file)
        template_blocks = []
        for block in base_doc.element.body:
            template_blocks.append(copy.deepcopy(block))
        base_doc.element.body.clear()

        progress_bar = st.progress(0)
        status_text = st.empty()
        total_rows = len(df)

        # 4. Vòng lặp
        for index, data_row in df.iterrows():
            progress_bar.progress((index + 1) / total_rows)
            status_text.text(f"Đang xử lý hồ sơ {index + 1}/{total_rows}...")
            
            for block in template_blocks:
                base_doc.element.body.append(copy.deepcopy(block))

            for paragraph in base_doc.paragraphs:
                for col in df.columns:
                    placeholder = f"{{{{{col}}}}}"
                    if placeholder in paragraph.text:
                        raw_val = str(data_row[col])
                        process_value_and_replace(paragraph, placeholder, raw_val, IMG_W, IMG_H)

            for table in base_doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for col in df.columns:
                            placeholder = f"{{{{{col}}}}}"
                            if placeholder in cell.text:
                                raw_val = str(data_row[col])
                                process_value_and_replace(cell, placeholder, raw_val, IMG_W, IMG_H)

            # Ngắt trang
            if index < total_rows - 1:
                if base_doc.paragraphs:
                    base_doc.paragraphs[-1].add_run().add_break(WD_BREAK.PAGE)
                else:
                    base_doc.add_page_break()

        # 5. Lưu file
        output_word_path = os.path.join(main_temp_dir, "KetQua_DonDaDien.docx")
        base_doc.save(output_word_path)

        status_text.text("Hoàn tất!")
        progress_bar.empty()
        return output_word_path

    except Exception as e:
        st.error("Gặp lỗi nghiêm trọng!")
        st.error(f"Chi tiết: {e}")
        st.code(traceback.format_exc())
        return None

# ================== GIAO DIỆN STREAMLIT ==================

st.title("📄 Tool Điền Đơn Tự Động (Word Only)")
st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. File Excel dữ liệu (.xlsx)", type="xlsx")
with col2:
    uploaded_word = st.file_uploader("2. File Word mẫu (.docx)", type="docx")

st.markdown("---")
use_image_zip = st.checkbox("Tôi muốn upload folder ảnh nén (.zip)", value=False)

uploaded_zip = None
match_col = None
img_placeholder_name = ""

if use_image_zip:
    uploaded_zip = st.file_uploader("Upload file .zip chứa ảnh", type="zip")
    if uploaded_excel:
        try:
            df_preview = pd.read_excel(uploaded_excel, nrows=0)
            cols = df_preview.columns.tolist()
            c1, c2 = st.columns(2)
            with c1:
                match_col = st.selectbox("Cột tên khớp ảnh:", cols)
            with c2:
                img_placeholder_name = st.text_input("Mã giữ chỗ ảnh (VD: ANH_THE):", value="ANH_THE")
        except:
            pass

st.markdown("---")

if st.button("🚀 BẮT ĐẦU XỬ LÝ", type="primary"):
    if uploaded_excel and uploaded_word:
        if use_image_zip and not uploaded_zip:
            st.warning("Vui lòng upload file Zip ảnh!")
        else:
            with st.spinner("Đang xử lý..."):
                word_out = generate_documents(
                    uploaded_excel, 
                    uploaded_word, 
                    uploaded_zip, 
                    match_col, 
                    img_placeholder_name
                )
                
                if word_out:
                    st.success("✅ Thành công!")
                    with open(word_out, "rb") as f:
                        st.download_button(
                            label="📥 Tải file kết quả (.docx)",
                            data=f,
                            file_name="KetQua_DonDaDien.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
    else:
        st.error("Thiếu file Excel hoặc Word mẫu!")
