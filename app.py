import streamlit as st
from pdf2docx import Converter
import tempfile
import os

st.set_page_config(
    page_title="PDF to Word Converter",
    page_icon="📄",
    layout="centered"
)

# ซ่อน Header และ Footer เพื่อความเรียบหรูเมื่อแสดงผล
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stButton>button {
        width: 100%;
        background-color: #2563EB;
        color: white;
        font-weight: bold;
        border-radius: 10px;
        padding: 0.6rem;
    }
    .stButton>button:hover {
        background-color: #1D4ED8;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📄 PDF to Word Converter")
st.write("แปลงไฟล์ PDF เป็น Word (.docx) พร้อมคงตาราง รูปภาพ และโครงสร้างเดิม")

# กล่องรับไฟล์ PDF
uploaded_file = st.file_uploader("เลือกไฟล์ PDF ที่ต้องการแปลง", type=["pdf"])

if uploaded_file is not None:
    file_size_mb = uploaded_file.size / (1024 * 1024)
    st.caption(f"ขนาดไฟล์: {file_size_mb:.2f} MB")
    
    if st.button("🚀 เริ่มแปลงเป็น Word"):
        with st.spinner("กำลังวิเคราะห์โครงสร้าง ตาราง และรูปภาพ..."):
            # 1. เขียนไฟล์ลงพื้นที่ชั่วคราว (Temp)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                tmp_pdf.write(uploaded_file.getvalue())
                pdf_path = tmp_pdf.name
                
            docx_path = pdf_path.replace(".pdf", ".docx")
            
            try:
                # 2. ประมวลผลแปลงไฟล์ด้วย pdf2docx
                cv = Converter(pdf_path)
                cv.convert(docx_path, start=0, end=None)
                cv.close()
                
                # 3. อ่านไฟล์ Word เข้า RAM เพื่อส่งให้ดาวน์โหลด
                with open(docx_path, "rb") as f:
                    docx_bytes = f.read()
                    
                st.success("✅ แปลงไฟล์สำเร็จเรียบร้อย!")
                
                out_name = uploaded_file.name.rsplit(".", 1)[0] + ".docx"
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ Word (.docx)",
                    data=docx_bytes,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาดในการแปลงไฟล์: {str(e)}")
            finally:
                # 4. ลบไฟล์ชั่วคราวออกจากระบบทันที
                if os.path.exists(pdf_path): 
                    os.remove(pdf_path)
                if os.path.exists(docx_path): 
                    os.remove(docx_path)
