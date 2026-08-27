import streamlit as st
from pdf2docx import Converter
import fitz  # PyMuPDF
import tempfile
import os

st.set_page_config(page_title="PDF to Word (Pro)", page_icon="📄", layout="centered")

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
    .stButton>button { background-color: #10B981; color: white; font-weight: bold; border-radius: 8px; padding: 0.6rem; }
    .stButton>button:hover { background-color: #059669; color: white; }
    </style>
""", unsafe_allow_html=True)

st.title("📄 PDF to Word Converter (Pro Engine)")
st.write("ระบบวิเคราะห์พิกัด จัดหน้าเป๊ะ และรักษารูปภาพแบบ Pre-Rasterization")

uploaded_file = st.file_uploader("📂 ลากไฟล์ PDF มาวางที่นี่", type=["pdf"])

if uploaded_file is not None:
    if st.button("🚀 เริ่มแปลงไฟล์เป็น Word", use_container_width=True):
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(uploaded_file.getvalue())
            pdf_path = tmp_pdf.name
            
        optimized_pdf_path = pdf_path.replace(".pdf", "_optimized.pdf")
        docx_path = pdf_path.replace(".pdf", ".docx")
        
        try:
            # ========================================================
            # ขั้นตอนที่ 1: PyMuPDF แปลงกราฟิกซับซ้อนให้เป็นรูปภาพฝังตัว
            # ========================================================
            with st.spinner("1/2 กำลังสแกนพิกัด และฝังรูปภาพลงบนกระดาษ..."):
                doc = fitz.open(pdf_path)
                
                for page in doc:
                    rects_to_capture = []
                    
                    # 1.1 ค้นหา Raster Images
                    for img in page.get_image_info():
                        rects_to_capture.append(fitz.Rect(img["bbox"]))
                        
                    # 1.2 ค้นหา Vector Drawings (กราฟ, ลายเส้น, แปลน)
                    for d in page.get_drawings():
                        r = d["rect"]
                        # กรองเส้นขอบกระดาษและจุดเล็กเกินไปทิ้ง
                        if 10 < r.width < (page.rect.width * 0.95) and 10 < r.height < (page.rect.height * 0.95):
                            rects_to_capture.append(r)
                            
                    # 1.3 ยุบรวมพิกัดรูปภาพที่อยู่ติดกันให้เป็นกล่องเดียว (Clustering)
                    merged_rects = []
                    for r in rects_to_capture:
                        # ระยะ margin ในการรวมกลุ่ม (5 px)
                        r_exp = r + (-5, -5, 5, 5) 
                        intersecting = [i for i, mr in enumerate(merged_rects) if r_exp.intersects(mr)]
                        
                        if not intersecting:
                            merged_rects.append(r)
                        else:
                            first_idx = intersecting[0]
                            merged_rects[first_idx] = merged_rects[first_idx] | r
                            for i in reversed(intersecting[1:]):
                                merged_rects[first_idx] = merged_rects[first_idx] | merged_rects[i]
                                del merged_rects[i]
                                
                    # 1.4 ถ่ายรูปความละเอียดสูง ถมขาว และ "แปะรูปฝังลงไปใน PDF" ทันที
                    for rect in merged_rects:
                        rect = rect.intersect(page.rect)  # กันกรอบทะลุกระดาษ
                        if rect.width < 5 or rect.height < 5: 
                            continue
                        
                        try:
                            # ถ่ายรูปพิกัดนั้นแบบชัดๆ (ซูม 3 เท่าเพื่อความคมชัดของ Word)
                            pix = page.get_pixmap(clip=rect, matrix=fitz.Matrix(3.0, 3.0))
                            if pix.width == 0 or pix.height == 0:
                                continue
                                
                            img_bytes = pix.tobytes("png")
                            
                            # ลบเส้นกราฟิกและข้อความเดิมในพื้นที่นั้นทิ้งให้เกลี้ยง
                            page.draw_rect(rect, color=(1,1,1), fill=(1,1,1))
                            
                            # แทรกรูปภาพความละเอียดสูงกลับเข้าไปที่พิกัด (X, Y) เดิมเป๊ะๆ
                            page.insert_image(rect, stream=img_bytes)
                            
                        except Exception:
                            continue
                            
                # เซฟ PDF ตัวใหม่ที่รูปภาพพร้อมแปลงแล้ว
                doc.save(optimized_pdf_path)
                doc.close()

            # ========================================================
            # ขั้นตอนที่ 2: ใช้ pdf2docx สร้างไฟล์ Word พร้อม Layout
            # ========================================================
            with st.spinner("2/2 กำลังประกอบโครงสร้างตารางและตัวอักษรลง Word..."):
                # ใส่พารามิเตอร์เพื่อให้ pdf2docx รักษา Layout ให้มากที่สุด
                cv = Converter(optimized_pdf_path)
                cv.convert(docx_path, start=0, end=None)
                cv.close()

            # ========================================================
            # ส่งไฟล์กลับให้ผู้ใช้
            # ========================================================
            with open(docx_path, "rb") as f:
                docx_bytes = f.read()
                
            st.success("✅ แปลงไฟล์สำเร็จ! รูปแบบและตำแหน่งมีความแม่นยำสูง")
            
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Word (.docx)",
                data=docx_bytes,
                file_name=uploaded_file.name.rsplit(".", 1)[0] + "_HighFidelity.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")
            
        finally:
            for path in [pdf_path, optimized_pdf_path, docx_path]:
                if os.path.exists(path):
                    os.remove(path)
