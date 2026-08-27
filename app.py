import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, Inches
import tempfile
import os
import io

st.set_page_config(page_title="PDF to Word (X-Y Engine)", page_icon="📐", layout="centered")

st.title("📐 PDF to Word (Hybrid X-Y Engine)")
st.write("สแกนพิกัด X, Y จาก PDF แล้วเรียงลง Word ตามตำแหน่งบนลงล่างแบบวิศวกรรม")

uploaded_file = st.file_uploader("📂 ลากไฟล์ PDF มาวาง", type=["pdf"])

if uploaded_file is not None:
    if st.button("🚀 เริ่มสแกนและแปลงไฟล์"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(uploaded_file.getvalue())
            pdf_path = tmp_pdf.name
            
        docx_path = pdf_path.replace(".pdf", ".docx")
        
        try:
            with st.spinner("กำลังสแกนพิกัด X, Y และสกัดข้อมูล..."):
                # 1. เปิดไฟล์ PDF ด้วย PyMuPDF
                doc = fitz.open(pdf_path)
                word_doc = Document()
                
                # ตั้งค่าหน้ากระดาษ Word (Margin แคบๆ เพื่อให้พื้นที่วางตรงกับ PDF มากที่สุด)
                for section in word_doc.sections:
                    section.top_margin = Inches(0.5)
                    section.bottom_margin = Inches(0.5)
                    section.left_margin = Inches(0.5)
                    section.right_margin = Inches(0.5)

                # 2. วนลูปอ่านทีละหน้า
                for page_num in range(len(doc)):
                    if page_num > 0:
                        word_doc.add_page_break()
                        
                    page = doc[page_num]
                    word_doc.add_heading(f'--- หน้าที่ {page_num + 1} ---', level=2)
                    
                    # 🎯 หัวใจหลัก: ดึงข้อมูลเป็น "Block" ซึ่งจะแถมพิกัด (x0, y0, x1, y1) มาให้ด้วย
                    # block_type: 0 = ข้อความ, 1 = รูปภาพ
                    blocks = page.get_text("blocks")
                    
                    # 🎯 อัลกอริทึมจัดเรียง: เรียงตามแกน Y (บนลงล่าง) เป็นหลัก และแกน X (ซ้ายไปขวา) เป็นรอง
                    # block[1] คือ y0 (พิกัด Y ด้านบน), block[0] คือ x0 (พิกัด X ด้านซ้าย)
                    blocks.sort(key=lambda b: (b[1], b[0]))
                    
                    # 3. วางลง Word ตามลำดับพิกัดที่สแกนมา
                    for b in blocks:
                        x0, y0, x1, y1 = b[:4]
                        block_type = b[6]
                        
                        # --- กรณีที่ 1: ถ้าบล็อกนั้นคือ "ข้อความ" ---
                        if block_type == 0:
                            text = b[4].strip()
                            if text:
                                # วางข้อความลงไป
                                p = word_doc.add_paragraph(text)
                                # (ทางเทคนิคสามารถตั้ง Indent ซ้ายขวาตามค่า x0, x1 ได้ด้วย Pt(x0))
                                
                        # --- กรณีที่ 2: ถ้าบล็อกนั้นคือ "รูปภาพ" ---
                        elif block_type == 1:
                            # ครอปรูปภาพเฉพาะพิกัด (x0, y0, x1, y1) นั้นออกมาจาก PDF
                            rect = fitz.Rect(x0, y0, x1, y1)
                            # เรนเดอร์จุดนั้นให้กลายเป็นรูป (ซูม 2 เท่าเพื่อความคมชัด)
                            pix = page.get_pixmap(clip=rect, matrix=fitz.Matrix(2.0, 2.0))
                            img_bytes = pix.tobytes("png")
                            
                            # วางรูปลง Word
                            img_stream = io.BytesIO(img_bytes)
                            width_pt = x1 - x0
                            # แปลงความกว้างจากพิกเซลหน้าจอเป็นหน่วยนิ้วลง Word (โดยประมาณ)
                            word_doc.add_picture(img_stream, width=Pt(width_pt))
                
                # 4. บันทึกไฟล์ Word
                word_doc.save(docx_path)
                doc.close()

            # ส่งไฟล์กลับให้ผู้ใช้ดาวน์โหลด
            with open(docx_path, "rb") as f:
                docx_bytes = f.read()
                
            st.success("✅ สแกนและประกอบไฟล์สำเร็จ!")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Word",
                data=docx_bytes,
                file_name="XY_Hybrid_Converted.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
        finally:
            if os.path.exists(pdf_path): os.remove(pdf_path)
            if os.path.exists(docx_path): os.remove(docx_path)
