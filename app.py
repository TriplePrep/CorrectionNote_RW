import streamlit as st
import pandas as pd
import zipfile
import os
import io
from PIL import Image
from fpdf import FPDF
from datetime import datetime

# ==============================
# 공통 설정
# ==============================

# PDF 생성용 폰트 경로
FONT_REGULAR = "fonts/NanumGothic.ttf"
FONT_BOLD = "fonts/NanumGothicBold.ttf"
pdf_font_name = "NanumGothic"

if os.path.exists(FONT_REGULAR) and os.path.exists(FONT_BOLD):
    class KoreanPDF(FPDF):
        def __init__(self):
            # 'L'을 추가하여 PDF 방향을 가로 모드 (Landscape)로 설정
            super().__init__(orientation='L') 
            self.set_margins(25.4, 30, 25.4)  # 왼쪽, 위쪽, 오른쪽 (mm 단위)
            self.set_auto_page_break(auto=True, margin=25.4) # 자동 페이지 나누기 여백
            self.add_font(pdf_font_name, '', FONT_REGULAR, uni=True)
            self.add_font(pdf_font_name, 'B', FONT_BOLD, uni=True)
            self.set_font(pdf_font_name, size=10)
else:
    st.error("⚠️ 한글 PDF 생성을 위해 fonts 폴더에 NanumGothic.ttf 와 NanumGothicBold.ttf 모두 필요합니다.")

# ==============================
# 유틸리티 함수
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

# ==============================
# 유틸: PDF 생성 (수정된 부분)
# ==============================
def create_student_pdf(name, m1_imgs, m2_imgs, doc_title, output_dir):
    pdf = KoreanPDF()
    pdf.add_page()
    pdf.set_font(pdf_font_name, style='B', size=10)
    
    # 💡 1. 문서 제목 출력 (첫 번째 줄)
    pdf.cell(0, 8, txt=f"<{name}_{doc_title}>", ln=True) 

    def add_images(title, images):
        TARGET_WIDTH = 240  # 이미지 너비 고정 (가로 모드 A4 폭 최대 활용)
        CELL_HEIGHT = 8     # 제목 및 여백 높이 8mm 유지
        
        # 💡 2. Module 제목 출력 (두 번째 줄)
        # Module2 제목이 바닥에 걸릴 경우: 제목 높이(8) + 이미지 최소 추정 높이(100)를 고려하여 페이지 넘김
        if title == "<Module2>" and pdf.get_y() + CELL_HEIGHT + 100 > pdf.page_break_trigger:
            pdf.add_page()
            
        pdf.set_font(pdf_font_name, size=10)
        pdf.cell(0, CELL_HEIGHT, txt=title, ln=True) 
        
        if images:
            # 💡 3. 이미지를 출력하기 전, 첫 번째 이미지에 대한 페이지 넘김 체크 (핵심 수정)
            
            # 첫 번째 이미지의 높이 계산 (w=240mm 기준)
            first_img = images[0]
            original_width_px, original_height_px = first_img.size
            if original_height_px != 0: 
                aspect_ratio = original_width_px / original_height_px
                calculated_height_mm = TARGET_WIDTH / aspect_ratio
            else:
                calculated_height_mm = 100 # 안전 값
            
            # 현재 Y 위치 + 이미지 높이 + 아래 여백(8mm)이 페이지 한계를 넘는지 확인
            # **이미 이 페이지에 제목 두 줄이 출력되었으므로, 여기서 체크하여 첫 이미지 다음 페이지 넘김 방지**
            if pdf.get_y() + calculated_height_mm + CELL_HEIGHT > pdf.page_break_trigger:
                 pdf.add_page()
                 # 새 페이지 시작 시 모듈 제목 다시 출력
                 pdf.set_font(pdf_font_name, size=10)
                 pdf.cell(0, CELL_HEIGHT, txt=title + " (이어짐)", ln=True) 
            
            
            # 💡 4. 이미지 출력 루프
            for idx, img in enumerate(images):
                # 💡 첫 번째 이미지 이후의 페이지 넘김은 루프 안에서 처리
                if idx > 0:
                    original_width_px, original_height_px = img.size
                    if original_height_px != 0: 
                        aspect_ratio = original_width_px / original_height_px
                        calculated_height_mm = TARGET_WIDTH / aspect_ratio
                    else:
                        calculated_height_mm = 100
                    
                    if pdf.get_y() + calculated_height_mm + CELL_HEIGHT > pdf.page_break_trigger:
                         pdf.add_page()
                         # 새 페이지 시작 시 모듈 제목 다시 출력
                         pdf.set_font(pdf_font_name, size=10)
                         pdf.cell(0, CELL_HEIGHT, txt=title + " (이어짐)", ln=True) 
                         
                
                # 이미지 출력
                img_path = f"temp_{datetime.now().timestamp()}.jpg"
                img.save(img_path)
                pdf.image(img_path, w=TARGET_WIDTH) 
                
                try:
                    os.remove(img_path)
                except Exception:
                    pass
                    
                pdf.ln(CELL_HEIGHT) # 이미지 아래 여백 8mm 유지

        else:
            pdf.cell(0, CELL_HEIGHT, txt="오답 없음", ln=True)
            pdf.ln(CELL_HEIGHT)

    add_images("<Module1>", m1_imgs)
    add_images("<Module2>", m2_imgs)

    pdf_path = os.path.join(output_dir, f"{name}_{doc_title}.pdf")
    pdf.output(pdf_path)
    return pdf_path

# ==============================
# Streamlit UI (변경 없음)
# ==============================
st.set_page_config(page_title="SAT 오답노트 생성기", layout="centered")
st.title("📝 SAT 오답노트 생성기 (PDF 가로 모드)")

st.header("📊 예시 엑셀 양식")
with st.expander("예시 엑셀파일 열기"):
    example_df = pd.DataFrame({
        '이름': ['홍길동', '김철수'],
        'Module1': ['1,3,5', '2,4'],
        'Module2': ['2,6', '1,3']
    })
    st.dataframe(example_df)
example = get_example_excel()
st.download_button("📥 예시 엑셀파일 다운로드", example, file_name="예시_오답노트_양식.xlsx")

st.header("📄 문서 제목 입력")
doc_title = st.text_input("문서 제목 (예: 25 S2 SAT MATH 만점반 Mock Test1)", value="25 S2 SAT MATH 만점반 Mock Test1")

st.header("📦 오답노트 파일 업로드")
st.caption("M1, M2 폴더 포함된 ZIP 파일 업로드")
img_zip = st.file_uploader("", type="zip", key="zip_uploader")

st.caption("오답노트 엑셀 파일 업로드 (.xlsx)")
excel_file = st.file_uploader("", type="xlsx", key="excel_uploader")

generated_files = []
generate = st.button("📎 오답노트 생성", type="primary")

if generate and img_zip and excel_file:
    with st.spinner("오답노트 생성 중..."):
        try:
            m1_imgs, m2_imgs = extract_zip_to_dict(img_zip)
            df = pd.read_excel(excel_file)
            
            output_dir = "generated_pdfs"
            os.makedirs(output_dir, exist_ok=True)

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

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                for name, path in generated_files:
                    zipf.write(path, os.path.basename(path))
            zip_buffer.seek(0)

            st.success(f"✅ 총 {len(generated_files)}개의 오답노트 PDF 생성 완료! (가로 모드)")
            st.download_button("📁 ZIP 파일 다운로드", zip_buffer, file_name=f"{doc_title}_오답노트_모음.zip", type="secondary")

        except Exception as e:
            st.error(f"오류 발생: {e}")

if generated_files:
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

