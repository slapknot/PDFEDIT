import streamlit as st
from pdf2docx import Converter
import os
import tempfile

# 1. ตั้งค่าหน้าเว็บให้พอดีกับกรอบ iframe
st.set_page_config(page_title="PDF to Word", layout="centered")

# 2. ซ่อนแถบเมนู Header และ Footer ของ Streamlit เพื่อให้เนียนกับเว็บหลัก
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# 3. ส่วนแสดงผล UI
st.subheader("📄 ระบบแปลงไฟล์ PDF เป็น MS Word")
st.write("อัปโหลดไฟล์ PDF ระบบจะดึงตัวอักษร ตาราง และรูปภาพออกมาเป็นไฟล์ .docx")

uploaded_file = st.file_uploader("ลากไฟล์ หรือ กดเพื่อเลือกไฟล์ PDF", type=["pdf"])

if uploaded_file is not None:
    if st.button("🚀 เริ่มการแปลงไฟล์", use_container_width=True):
        with st.spinner("กำลังประมวลผล... โปรดรอสักครู่ (ขึ้นอยู่กับขนาดไฟล์)"):
            
            # สร้างไฟล์ชั่วคราวเพื่อรับและส่งข้อมูล
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                tmp_pdf.write(uploaded_file.getvalue())
                pdf_path = tmp_pdf.name
                
            docx_path = pdf_path.replace(".pdf", ".docx")
            
            try:
                # เริ่มกระบวนการแปลงไฟล์ด้วย pdf2docx
                cv = Converter(pdf_path)
                cv.convert(docx_path)
                cv.close()
                
                st.success("✅ แปลงไฟล์สำเร็จแล้ว!")
                
                # สร้างปุ่มสำหรับดาวน์โหลดไฟล์ Word
                with open(docx_path, "rb") as file:
                    st.download_button(
                        label="📥 ดาวน์โหลดไฟล์ Word (.docx)",
                        data=file,
                        file_name=uploaded_file.name.replace(".pdf", ".docx"),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาดในการแปลงไฟล์: ไฟล์อาจถูกเข้ารหัส หรือเป็นภาพสแกน 100%")
            finally:
                # ล้างไฟล์ขยะในเซิร์ฟเวอร์
                if os.path.exists(pdf_path): os.remove(pdf_path)
                if os.path.exists(docx_path): os.remove(docx_path)