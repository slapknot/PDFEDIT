import streamlit as st
from pdf2docx import Converter
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches
import tempfile
import os
import io

st.set_page_config(page_title="PDF to Word", page_icon="📄", layout="centered")

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
    .stButton>button { background-color: #10B981; color: white; font-weight: bold; border-radius: 8px; padding: 0.6rem; }
    .stButton>button:hover { background-color: #059669; color: white; }
    </style>
""", unsafe_allow_html=True)

st.title("📄 PDF to Word (Pro Engine)")
st.write("แปลงไฟล์รักษารูปแบบ (Layout) และแทรกรูปภาพตรงตำแหน่งเดิม")

uploaded_file = st.file_uploader("📂 ลากไฟล์ PDF มาวางที่นี่", type=["pdf"])

if uploaded_file is not None:
    if st.button("🚀 เริ่มแปลงไฟล์ (ระบบวิเคราะห์พิกัด)", use_container_width=True):
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(uploaded_file.getvalue())
            pdf_path = tmp_pdf.name
            
        marked_pdf_path = pdf_path.replace(".pdf", "_marked.pdf")
        docx_path = pdf_path.replace(".pdf", ".docx")
        
        try:
            # ========================================================
            # ขั้นตอนที่ 1: PyMuPDF สแกนพิกัด, เก็บรูป, และฝังรหัสลับ
            # ========================================================
            with st.spinner("1/3 กำลังสแกนพิกัดและสกัดรูปภาพ..."):
                doc = fitz.open(pdf_path)
                img_dict = {}
                img_count = 0
                
                for page in doc:
                    blocks = page.get_text("blocks")
                    for b in blocks:
                        if b[6] == 1:  # ถ้าเป็น Block ประเภท "รูปภาพ"
                            rect = fitz.Rect(b[:4]) # พิกัด x0, y0, x1, y1
                            
                            # แคปเจอร์รูปภาพตรงพิกัดนั้น (ความละเอียด 2 เท่า)
                            pix = page.get_pixmap(clip=rect, matrix=fitz.Matrix(2.0, 2.0))
                            marker = f"[[IMG_{img_count}]]"
                            
                            # เก็บรูปไว้ใน RAM
                            img_dict[marker] = {
                                "bytes": pix.tobytes("png"),
                                "width": rect.width
                            }
                            
                            # ถมสีขาวทับรูปเดิม แล้วพิมพ์รหัสลับลงไปแทนที่
                            page.draw_rect(rect, color=(1,1,1), fill=(1,1,1))
                            page.insert_text((rect.x0, rect.y0 + 12), marker, fontsize=10, color=(0,0,0))
                            img_count += 1
                            
                # เซฟเป็น PDF ตัวใหม่ที่มีแต่ตัวหนังสือและรหัสลับ
                doc.save(marked_pdf_path)
                doc.close()

            # ========================================================
            # ขั้นตอนที่ 2: ใช้ pdf2docx จัด Format และ Layout
            # ========================================================
            with st.spinner("2/3 กำลังสร้างโครงสร้าง Word และจัด Layout..."):
                cv = Converter(marked_pdf_path)
                cv.convert(docx_path, start=0, end=None)
                cv.close()

            # ========================================================
            # ขั้นตอนที่ 3: เปิดไฟล์ Word หาตำแหน่งรหัสลับ แล้ววางรูปลงไป!
            # ========================================================
            with st.spinner("3/3 กำลังประกอบรูปภาพลงในตำแหน่งเดิม..."):
                word_doc = Document(docx_path)
                
                def replace_markers(paragraphs):
                    for p in paragraphs:
                        for marker, img_data in img_dict.items():
                            if marker in p.text:
                                # เคลียร์รหัสลับทิ้ง
                                p.text = "" 
                                # วางรูปภาพลงไปแทนที่
                                run = p.add_run()
                                img_stream = io.BytesIO(img_data["bytes"])
                                # คำนวณขนาดรูปให้พอดี (จำกัดกว้างสุด 6.5 นิ้ว ไม่ให้ล้นหน้า)
                                width_inch = min(img_data["width"] / 72.0, 6.5)
                                run.add_picture(img_stream, width=Inches(width_inch))

                # สแกนหาในข้อความปกติ
                replace_markers(word_doc.paragraphs)
                
                # สแกนหาในตาราง (pdf2docx มักใช้ตารางซ่อนเส้นเพื่อทำเลย์เอาต์)
                for table in word_doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            replace_markers(cell.paragraphs)
                            
                word_doc.save(docx_path)

            # ========================================================
            # สำเร็จ! ส่งไฟล์ Word กลับไปให้ผู้ใช้ดาวน์โหลด
            # ========================================================
            with open(docx_path, "rb") as f:
                docx_bytes = f.read()
                
            st.success(f"✅ แปลงไฟล์และแทรกรูปภาพสำเร็จ ({img_count} รูป)")
            
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Word (Layout + รูปภาพครบ)",
                data=docx_bytes,
                file_name=uploaded_file.name.rsplit(".", 1)[0] + "_Perfect.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")
            
        finally:
            # เคลียร์ไฟล์ขยะ
            for path in [pdf_path, marked_pdf_path, docx_path]:
                if os.path.exists(path):
                    os.remove(path)
