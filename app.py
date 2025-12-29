import streamlit as st
import pandas as pd
import zipfile
import os
import io
from PIL import Image
from fpdf import FPDF, XPos, YPos
from datetime import datetime
import fitz  # PyMuPDF 라이브러리 (PDF->이미지 변환용)

# ==============================
# 공통 설정
# ==============================

# PDF 생성용 폰트 경로 (절대 경로로 변환)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_REGULAR = os.path.join(BASE_DIR, "fonts", "NanumGothic.ttf")
FONT_BOLD = os.path.join(BASE_DIR, "fonts", "NanumGothicBold.ttf")

# ==============================
# 유틸리티 함수 (Tab 1 용)
# ==============================

def get_example_excel():
    output = io.BytesIO()
    example_df = pd.DataFrame({
        '이름': ['홍길동', '김철수'],
        'Module1': ['1,3,5', '2,4'],
        'Module2': ['2,6', '1,3']
    })
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        example_df.to_excel(writer, index=False)
    output.seek(0)
    return output

def extract_zip_to_dict(zip_file):
    m1_imgs, m2_imgs = {}, {}
    with zipfile.ZipFile(zip_file) as z:
        for file in z.namelist():
            if file.lower().endswith(('png', 'jpg', 'jpeg')):
                parts = file.split('/')
                if len(parts) < 2:
                    continue
                folder = parts[0].lower()
                q_num = os.path.splitext(os.path.basename(file))[0]
                with z.open(file) as f:
                    img = Image.open(f).convert("RGB")
                    if folder == "m1":
                        m1_imgs[q_num] = img
                    elif folder == "m2":
                        m2_imgs[q_num] = img
    return m1_imgs, m2_imgs

def create_pdf_with_korean_font():
    """한글 폰트가 등록된 새 FPDF 인스턴스 생성"""
    pdf = FPDF(orientation='L')
    pdf.set_margins(25.4, 20, 25.4)
    pdf.set_auto_page_break(auto=True, margin=20)
    
    # 폰트 등록 (subsetting=False로 TSI 경고 해결)
    pdf.add_font("nanum", "", FONT_REGULAR, uni=True)
    pdf.add_font("nanum", "B", FONT_BOLD, uni=True)
    pdf.set_font_subsetting(False)  # 서브셋 비활성화
    
    return pdf

def create_student_pdf(name, m1_imgs, m2_imgs, doc_title, output_dir):
    """학생별 오답노트 PDF 생성"""
    pdf = create_pdf_with_korean_font()
    
    pdf.add_page()
    pdf.set_font("nanum", style='B', size=10)
    pdf.cell(0, 8, text=f"<{name}_{doc_title}>", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    def add_images(title, images):
        img_est_height = 100
        if title == "<Module2>" and pdf.get_y() + 10 + (img_est_height if images else 0) > pdf.page_break_trigger:
            pdf.add_page()

        pdf.set_font("nanum", size=10)
        pdf.cell(0, 8, text=title, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        if images:
            for img in images:
                img_path = os.path.join(output_dir, f"temp_{datetime.now().timestamp()}.jpg")
                img.save(img_path)
                pdf.image(img_path, h=153)
                try:
                    os.remove(img_path)
                except Exception:
                    pass
                pdf.ln(8)
        else:
            pdf.cell(0, 8, text="오답 없음", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.ln(8)

    add_images("<Module1>", m1_imgs)
    add_images("<Module2>", m2_imgs)

    pdf_path = os.path.join(output_dir, f"{name}_{doc_title}.pdf")
    pdf.output(pdf_path)
    return pdf_path

# ==============================
# Streamlit UI
# ==============================
st.set_page_config(page_title="SAT 오답노트 & 캡쳐 생성기", layout="centered")
st.title("SAT 오답노트 & 캡쳐 생성기")

tab1, tab2 = st.tabs(["📝 오답노트 생성기", "🖼️ 캡쳐이미지 ZIP 생성기"])

# =========================================================
# 탭 1: 오답노트 생성기
# =========================================================
with tab1:
    # 폰트 파일 존재 확인
    fonts_ok = os.path.exists(FONT_REGULAR) and os.path.exists(FONT_BOLD)
    if not fonts_ok:
        st.error(f"⚠️ 폰트 파일을 찾을 수 없습니다.\n\n경로: {FONT_REGULAR}")
         
    st.header("📊 예시 엑셀 양식")
    with st.expander("예시 엑셀파일 열기"):
        st.dataframe(pd.read_excel(get_example_excel()))
    example = get_example_excel()
    st.download_button("📥 예시 엑셀파일 다운로드", example, file_name="예시_오답노트_양식.xlsx")

    st.header("📄 문서 제목 입력")
    doc_title = st.text_input("문서 제목 (예: [11월대비01RW])", value="[11월대비01RW]")

    st.header("📦 오답노트 파일 업로드")
    st.caption("M1, M2 폴더 포함된 ZIP 파일 업로드")
    img_zip = st.file_uploader("ZIP 파일", type="zip", key="zip_uploader_tab1")

    st.caption("오답노트 엑셀 파일 업로드 (.xlsx)")
    excel_file = st.file_uploader("XLSX 파일", type="xlsx", key="excel_uploader_tab1")

    generate = st.button("📎 오답노트 생성")

    if generate and img_zip and excel_file:
        if not fonts_ok:
            st.error("폰트 파일이 없어 PDF를 생성할 수 없습니다.")
        else:
            with st.spinner("오답노트 생성 중..."):
                try:
                    m1_imgs, m2_imgs = extract_zip_to_dict(img_zip)
                    df = pd.read_excel(excel_file)
                    
                    output_dir = os.path.join(BASE_DIR, "generated_pdfs")
                    os.makedirs(output_dir, exist_ok=True)
                    
                    generated_files = []

                    for _, row in df.iterrows():
                        if '이름' not in row or 'Module1' not in row or 'Module2' not in row:
                            continue
                            
                        name = row['이름']

                        if pd.isna(row['Module1']) or pd.isna(row['Module2']):
                            continue

                        m1_nums = [num.strip() for num in str(row['Module1']).split(',') if num.strip()] if pd.notna(row['Module1']) else []
                        m2_nums = [num.strip() for num in str(row['Module2']).split(',') if num.strip()] if pd.notna(row['Module2']) else []
                        
                        m1_list = [m1_imgs[num] for num in m1_nums if num in m1_imgs]
                        m2_list = [m2_imgs[num] for num in m2_nums if num in m2_imgs]
                        
                        if m1_list or m2_list:
                            pdf_path = create_student_pdf(name, m1_list, m2_list, doc_title, output_dir)
                            generated_files.append((name, pdf_path))

                    if generated_files:
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w") as zipf:
                            for name, path in generated_files:
                                zipf.write(path, os.path.basename(path))
                        zip_buffer.seek(0)

                        st.success(f"✅ 총 {len(generated_files)}개의 오답노트 PDF 생성 완료!")
                        st.download_button("📁 ZIP 파일 다운로드", zip_buffer, file_name=f"{doc_title}_오답노트_모음.zip", type="primary")
                        
                        # 개별 다운로드
                        st.markdown("---")
                        st.header("👁️ 개별 PDF 다운로드")
                        sorted_names = sorted([name for name, _ in generated_files])
                        selected = st.selectbox("학생 선택", sorted_names, index=0)
                        
                        if selected:
                            generated_dict = {name: path for name, path in generated_files}
                            selected_path = generated_dict[selected]
                            with open(selected_path, "rb") as f:
                                st.download_button(
                                    f"📄 {selected} PDF 다운로드", 
                                    f, 
                                    file_name=os.path.basename(selected_path), 
                                    type="secondary"
                                )
                    else:
                        st.warning("생성된 PDF가 없습니다. 엑셀 파일과 이미지를 확인해주세요.")

                except Exception as e:
                    st.error(f"오류 발생: {e}")
                    import traceback
                    st.code(traceback.format_exc())

# =========================================================
# 탭 2: 캡쳐이미지 ZIP 생성기
# =========================================================
with tab2:
    st.header("📄 PDF 파일 업로드")
    pdf_file = st.file_uploader("변환할 PDF 파일 업로드", type="pdf", key="pdf_uploader_tab2")
    st.info("PDF파일을 페이지별로 JPG이미지 파일로 변환합니다  \n  \n1페이지 당 1문제 기준으로 분리하여, 문제번호를 순서대로 부여하고  \n오답노트 생성기에 **연동가능한** 양식의 ZIP파일로 정리해줍니다")

    st.markdown("---")

    st.header("📖 페이지 범위 설정")
    
    st.subheader("Module 1 (M1) 설정")
    col1, col2 = st.columns(2)
    with col1:
        m1_start = st.number_input("M1 시작 페이지", min_value=1, value=3, key="m1_start")
    with col2:
        m1_end = st.number_input("M1 종료 페이지", min_value=1, value=29, key="m1_end")

    st.subheader("Module 2 (M2) 설정")
    col3, col4 = st.columns(2)
    with col3:
        m2_start = st.number_input("M2 시작 페이지", min_value=1, value=33, key="m2_start")
    with col4:
        m2_end = st.number_input("M2 종료 페이지", min_value=1, value=59, key="m2_end")

    st.markdown("---")

    st.header("⚙️ 변환 품질 설정")
    
    col5, col6 = st.columns(2)
    with col5:
        dpi = st.slider("해상도 (DPI)", min_value=150, max_value=600, value=300, step=75)
        st.caption("높을수록 선명하지만 변환 속도가 오래 걸리고 파일이 커집니다. (기본: 300)")
    with col6:
        jpg_quality = st.slider("JPG 압축 품질", min_value=75, max_value=100, value=95, step=5)
        st.caption("높을수록 원본에 가깝지만 파일이 커집니다. (기본: 95)")

    st.markdown("---")

    capture_button = st.button("🖼️ 캡쳐이미지 ZIP 생성", type="primary")

    if capture_button and pdf_file:
        
        def process_pages_to_zip(doc, start_page, end_page, zip_handle, folder_name, dpi_setting, quality_setting):
            start_idx = start_page - 1
            end_idx = end_page
            img_counter = 1
            
            if start_idx >= len(doc):
                st.warning(f"'{folder_name}' 시작 페이지({start_page})가 PDF 전체 페이지({len(doc)})보다 큽니다.")
                return 0
            if end_idx > len(doc):
                st.warning(f"'{folder_name}' 종료 페이지({end_page})가 PDF 전체 페이지({len(doc)})보다 큽니다.")
                end_idx = len(doc)
            if start_idx >= end_idx:
                st.warning(f"'{folder_name}' 시작 페이지가 종료 페이지보다 크거나 같습니다.")
                return 0

            for i in range(start_idx, end_idx):
                page = doc.load_page(i)
                pix = page.get_pixmap(dpi=dpi_setting) 
                
                img_data = pix.tobytes("ppm")
                img = Image.frombytes("RGB", [pix.width, pix.height], img_data)
                
                img_buffer = io.BytesIO()
                img.save(img_buffer, format="JPEG", quality=quality_setting)
                img_buffer.seek(0)
                
                file_name = f"{folder_name}/{img_counter}.jpg"
                zip_handle.writestr(file_name, img_buffer.read())
                
                img_counter += 1
                
            return img_counter - 1

        try:
            with st.spinner(f"PDF 페이지를 이미지로 변환 중... (DPI: {dpi}, 품질: {jpg_quality})"):
                pdf_bytes = pdf_file.getvalue()
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                
                zip_buffer_capture = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer_capture, "w", zipfile.ZIP_DEFLATED) as zf:
                    m1_count = process_pages_to_zip(doc, m1_start, m1_end, zf, "M1", dpi, jpg_quality)
                    m2_count = process_pages_to_zip(doc, m2_start, m2_end, zf, "M2", dpi, jpg_quality)
                
                doc.close()
                zip_buffer_capture.seek(0)

            st.success(f"✅ ZIP 생성 완료! (M1: {m1_count}장, M2: {m2_count}장)")
            
            original_name = os.path.splitext(pdf_file.name)[0]
            zip_name = f"{original_name}_캡쳐.zip"
            
            st.download_button(
                "📁 캡쳐 ZIP 파일 다운로드",
                zip_buffer_capture,
                file_name=zip_name,
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"오류 발생: {e}")
            import traceback
            st.code(traceback.format_exc())
