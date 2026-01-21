import streamlit as st
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import tempfile
import os
import pytesseract

# ===== CONFIG STREAMLIT =====
st.set_page_config(
    page_title="PDF → PPTX",
    layout="wide"
)

st.title("📄➡️📊 Conversor de PDF para PowerPoint")
st.write("Converte PDF mantendo layout fiel (imagem por slide)")

# ===== UPLOAD =====
uploaded_pdf = st.file_uploader(
    "Selecione o PDF",
    type=["pdf"]
)

if uploaded_pdf:
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = os.path.join(tmpdir, "entrada.pdf")

        with open(pdf_path, "wb") as f:
            f.write(uploaded_pdf.read())

        if st.button("Converter para PPTX"):
            with st.spinner("Convertendo... aguarde"):

                # 🔥 CONVERSÃO PDF → IMAGENS (POPPLER NO LINUX)
                pages = convert_from_path(
                    pdf_path,
                    dpi=300,
                    poppler_path="/usr/bin"
                )

                prs = Presentation()
                prs.slide_width = Inches(13.33)
                prs.slide_height = Inches(7.5)

                for img in pages:
                    slide = prs.slides.add_slide(prs.slide_layouts[6])

                    img_stream = BytesIO()
                    img.save(img_stream, format="PNG")
                    img_stream.seek(0)

                    slide.shapes.add_picture(
                        img_stream,
                        Inches(0),
                        Inches(0),
                        prs.slide_width,
                        prs.slide_height
                    )

                ppt_path = os.path.join(tmpdir, "saida.pptx")
                prs.save(ppt_path)

                with open(ppt_path, "rb") as f:
                    st.success("✅ Conversão concluída!")
                    st.download_button(
                        "⬇️ Baixar PPTX",
                        f,
                        file_name="convertido.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
