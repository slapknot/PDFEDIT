import streamlit as st
from pdf2docx import Converter
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import tempfile
import os
import io

# 1. ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="PDF to Word Converter",
    page_icon="📄",
    layout="centered"
)

# 2. ซ่อน Header/Footer ของ Streamlit เพื่อความเนียนเมื่ออยู่ใน Iframe
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stButton>button {
        background-color: #2563EB;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.6rem;
    }
    .stButton>button:hover {
        background-color: #1D4ED8;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📄 PDF to Word Converter")
st.write("แปลงไฟล์ PDF เป็น Word พร้อมตัวเลือกรักษารูปภาพและตาราง")

# 3. อัปโหลดไฟล์ชั่วคราวลง RAM
uploaded_file = st.file_uploader("📂 ลากไฟล์ PDF มาวางที่นี่", type=["pdf"])

if uploaded_file is not None:
    # 4. เลือกโหมดการแปลง
    st.markdown("### ⚙️ เลือกรูปแบบการแปลงไฟล์")
    mode = st.radio(
        "เลือกโหมดที่เหมาะสมกับเอกสารของคุณ:",
        options=[
            "📝 โหมดเน้นข้อความ (พิมพ์แก้ไขต่อได้ - อาจสูญเสียกราฟิกซับซ้อน)",
            "🖼️ โหมดเน้นหน้าตา (ได้รูปภาพ/แปลน/ตารางมาครบ 100% - แต่แก้ข้อความไม่ได้)"
        ]
    )

    if st.button("🚀 เริ่มแปลงไฟล์เป็น Word", use_container_width=True):
        
        # เตรียมพื้นที่ชั่วคราว
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(uploaded_file.getvalue())
            pdf_path = tmp_pdf.name
            
        docx_path = pdf_path.replace(".pdf", ".docx")
        
        try:
            # ---------------------------------------------------------
            # โหมดที่ 1: เน้นข้อความ (pdf2docx)
            # ---------------------------------------------------------
            if "โหมดเน้นข้อความ" in mode:
                with st.spinner("กำลังแปลงเป็นข้อความ..."):
                    cv = Converter(pdf_path)
                    cv.convert(docx_path, start=0, end=None)
                    cv.close()

            # ---------------------------------------------------------
            # โหมดที่ 2: เน้นหน้าตา (PyMuPDF + python-docx)
            # ---------------------------------------------------------
            else:
                with st.spinner("กำลังดึงรูปภาพและกราฟิกระดับความละเอียดสูง..."):
                    doc = fitz.open(pdf_path)
                    word_doc = Document()
                    
                    # ลดขอบหน้ากระดาษ Word เพื่อให้รูปภาพเต็มแผ่น
                    for section in word_doc.sections:
                        section.top_margin = Inches(0.4)
                        section.bottom_margin = Inches(0.4)
                        section.left_margin = Inches(0.4)
                        section.right_margin = Inches(0.4)
                    
                    # เรนเดอร์แต่ละหน้าเป็นภาพ 300 DPI
                    for page_num in range(len(doc)):
                        page = doc[page_num]
                        zoom = 2.0  # เพิ่มความคมชัด 2 เท่า
                        mat = fitz.Matrix(zoom, zoom)
                        pix = page.get_pixmap(matrix=mat, alpha=False)
                        img_bytes = pix.tobytes("png")
                        
                        img_stream = io.BytesIO(img_bytes)
                        # แทรกรูปลง Word (กว้างสุดหน้ากระดาษ)
                        word_doc.add_picture(img_stream, width=Inches(7.2))
                        
                    word_doc.save(docx_path)
                    doc.close()

            # ---------------------------------------------------------
            # อ่านไฟล์ส่งให้ผู้ใช้ และลบไฟล์ขยะ
            # ---------------------------------------------------------
            with open(docx_path, "rb") as f:
                docx_bytes = f.read()
                
            st.success("✅ แปลงไฟล์สำเร็จเรียบร้อย!")
            
            out_name = uploaded_file.name.rsplit(".", 1)[0] + "_Converted.docx"
            st.download_button(
                label="📥 คลิกที่นี่เพื่อดาวน์โหลดไฟล์ Word",
                data=docx_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")
            
        finally:
            # คลีน RAM/Temp storage ทันทีเพื่อความปลอดภัย
            if os.path.exists(pdf_path): 
                os.remove(pdf_path)
            if os.path.exists(docx_path): 
                os.remove(docx_path)
