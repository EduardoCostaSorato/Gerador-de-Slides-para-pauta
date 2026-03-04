
import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from io import BytesIO

st.set_page_config(page_title="Gerador de Slides para pauta", layout="centered")

st.title("Gerador de Slides para pauta")
st.write("Envie a planilha Excel para gerar os slides automaticamente")

arquivo = st.file_uploader("Enviar Excel", type=["xlsx"])


def adicionar_texto(slide, texto, x, y, largura, tamanho, cor):

    caixa = slide.shapes.add_textbox(x, y, largura, Inches(1.2))

    tf = caixa.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER

    run = p.add_run()
    run.text = texto
    run.font.name = "Century Gothic"
    run.font.size = Pt(tamanho)
    run.font.bold = True
    run.font.color.rgb = cor


if arquivo:

    df = pd.read_excel(arquivo)
    df.columns = df.columns.str.strip().str.lower()

    st.success("Excel carregado com sucesso")

    if st.button("Gerar apresentação"):

        progress = st.progress(0)

        substituicoes = {
            "JOAO EGMONT LEONCIO LOPES": "JOÃO EGMONT",
            "HECTOR VALVERDE SANTANNA": "HÉCTOR VALVERDE",
            "RENATO RODOVALHO SCUSSEL": "RENATO SCUSSEL",
            " FERNANDO ANTÔNIO TAVERNARD LIMA": "FERNANDO TAVERNARD"
        }

        prs = Presentation("modelo.pptx")

        largura_slide = prs.slide_width
        slide_ref = prs.slides[1]

        img_ref = None

        for shape in slide_ref.shapes:
            if shape.shape_type == 13:
                img_ref = shape
                break

        layout = prs.slide_layouts[6]

        for i, row in df.iterrows():

            slide = prs.slides.add_slide(layout)

            if img_ref:
                img_stream = BytesIO(img_ref.image.blob)

                slide.shapes.add_picture(
                    img_stream,
                    img_ref.left,
                    img_ref.top,
                    img_ref.width,
                    img_ref.height
                )

            adicionar_texto(
                slide,
                "2ª Turma Cível",
                Inches(4.5),
                Inches(1.15),
                Inches(4),
                18,
                RGBColor(255,255,255)
            )

            proc = str(row["processo"]).split(".8")[0]

            num = str(row["numero"]).replace(".0","").strip()

            texto = f"{proc} ({num})"

            adicionar_texto(
                slide,
                texto,
                0,
                Inches(2.7),
                largura_slide,
                60,
                RGBColor(0,0,0)
            )

            nome_original = str(row["desembargador"]).upper().strip()
            nome = substituicoes.get(nome_original, nome_original)

            relator = f"RELATOR:\nDESEMBARGADOR {nome}"

            adicionar_texto(
                slide,
                relator,
                0,
                Inches(4.2),
                largura_slide,
                32,
                RGBColor(0,0,0)
            )

            progress.progress((i+1)/len(df))

        while len(prs.slides) > len(df):
            prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

        output = BytesIO()
        prs.save(output)

        st.success("Apresentação gerada!")

        st.download_button(
            label="Baixar PowerPoint",
            data=output.getvalue(),
            file_name="slides_pauta.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"

        )
