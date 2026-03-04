import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from io import BytesIO
import os

st.set_page_config(
    page_title="Gerador de Slides da Pauta",
    layout="centered"
)

st.title("Gerador de Slides da Pauta")

# =====================================
# CONFIGURAÇÕES EDITÁVEIS
# =====================================

st.subheader("Configurações")

turma = st.text_input(
    "Nome da Turma",
    value="2ª Turma Cível"
)

st.write("### Substituições de nomes")

col1, col2 = st.columns(2)

with col1:
    nome1 = st.text_input(
        "JOAO EGMONT LEONCIO LOPES",
        value="JOÃO EGMONT"
    )

    nome2 = st.text_input(
        "HECTOR VALVERDE SANTANNA",
        value="HÉCTOR VALVERDE"
    )

with col2:
    nome3 = st.text_input(
        "RENATO RODOVALHO SCUSSEL",
        value="RENATO SCUSSEL"
    )

    nome4 = st.text_input(
        "FERNANDO ANTÔNIO TAVERNARD LIMA",
        value="FERNANDO TAVERNARD"
    )

    nome5 = st.text_input(
        "ALVARO CIARLINI",
        value="ALVARO CIARLINI"
    )

substituicoes = {
    "JOAO EGMONT LEONCIO LOPES": nome1,
    "HECTOR VALVERDE SANTANNA": nome2,
    "RENATO RODOVALHO SCUSSEL": nome3,
    "FERNANDO ANTÔNIO TAVERNARD LIMA": nome4,
    "ALVARO CIARLINI": nome5
}

st.divider()

# =====================================
# UPLOAD
# =====================================

arquivo = st.file_uploader("Enviar Excel", type=["xlsx"])


# =====================================
# FUNÇÃO TEXTO
# =====================================

def adicionar_texto(slide, texto, x, y, largura, tamanho, cor):

    caixa = slide.shapes.add_textbox(x, y, largura, Inches(1.2))

    tf = caixa.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER

    run = p.add_run()
    run.text = str(texto)

    run.font.name = "Century Gothic"
    run.font.size = Pt(tamanho)
    run.font.bold = True
    run.font.color.rgb = cor


# =====================================
# PROCESSAR EXCEL
# =====================================

if arquivo:

    df = pd.read_excel(arquivo)
    df.columns = df.columns.str.strip().str.lower()

    st.success("Excel carregado")

    st.subheader("Pré-visualização da planilha")
    st.dataframe(df)

    colunas_necessarias = {"numero","processo","desembargador"}

    if not colunas_necessarias.issubset(df.columns):

        st.error("A planilha precisa ter: numero, processo, desembargador")
        st.stop()

    if st.button("Gerar apresentação"):

        if not os.path.exists("modelo.pptx"):
            st.error("Arquivo modelo.pptx não encontrado")
            st.stop()

        progress = st.progress(0)

        prs = Presentation("modelo.pptx")

        largura_slide = prs.slide_width
        slide_ref = prs.slides[1]

        img_ref = None

        for shape in slide_ref.shapes:
            if shape.shape_type == 13:
                img_ref = shape
                break

        layout = prs.slide_layouts[6]

        total = len(df)

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

            # Nome da turma
            adicionar_texto(
                slide,
                turma,
                Inches(3.5),
                Inches(1.15),
                Inches(4),
                18,
                RGBColor(255,255,255)
            )

            # Processo
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

            # Relator
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

            # =====================================
            # SEGREDO DE JUSTIÇA
            # =====================================

            if "segredo" in df.columns:

                segredo = str(row["segredo"]).upper().strip()

                if segredo == "SIM":

                    adicionar_texto(
                        slide,
                        "SEGREDO DE JUSTIÇA",
                        0,
                        Inches(5.3),
                        largura_slide,
                        40,
                        RGBColor(255,0,0)
                    )

            progress.progress((i+1)/total)

        # manter slide inicial
        slides_gerados = len(df)

        while len(prs.slides) > slides_gerados + 1:
            prs.slides._sldIdLst.remove(prs.slides._sldIdLst[1])

        output = BytesIO()
        prs.save(output)

        st.success("Apresentação gerada!")

        st.download_button(
            label="Baixar PowerPoint",
            data=output.getvalue(),
            file_name="slides_pauta.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
