import streamlit as st
from pdf2image import convert_from_bytes
from pptx import Presentation
from pptx.util import Inches
import io

st.set_page_config(page_title="PDF to PPT")
st.title("🚀 Web PDF to PowerPoint")

uploaded_file = st.file_uploader("Upload your PDF", type="pdf")

if uploaded_file and st.button("Convert"):
    try:
        with st.spinner("Converting... this may take a minute."):
            # NOTICE: No 'poppler_path' here. The web server handles it!
            images = convert_from_bytes(uploaded_file.read(), dpi=100)
            
            prs = Presentation()
            for img in images:
                # Resize logic to stay under PowerPoint's 56-inch limit
                w, h = img.size
                sw, sh = w/100, h/100
                if sw > 50 or sh > 50:
                    scale = 50 / max(sw, sh)
                    sw, sh = sw * scale, sh * scale
                
                prs.slide_width = Inches(sw)
                prs.slide_height = Inches(sh)
                
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                img_io = io.BytesIO()
                img.save(img_io, format='JPEG', quality=80)
                slide.shapes.add_picture(img_io, 0, 0, width=Inches(sw), height=Inches(sh))
            
            ppt_io = io.BytesIO()
            prs.save(ppt_io)
            st.success("Done!")
            st.download_button("📥 Download PPTX", ppt_io.getvalue(), "converted.pptx")
    except Exception as e:
        st.error(f"Error: {e}")
