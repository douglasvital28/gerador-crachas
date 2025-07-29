import streamlit as st
import pandas as pd
import zipfile
import tempfile
import fitz
import os

# Caminho do template embutido (no mesmo diretório)
TEMPLATE_PDF = "cracha_template.pdf"

st.set_page_config(page_title="Gerador de Crachás", page_icon="🪪")
st.title("🪪 Gerador de Crachás - Caravana")

# Upload do Excel
df_file = st.file_uploader("📥 Envie a planilha preenchida (Excel)", type=["xlsx"])

# Campos adicionais
horario = st.text_input("Horário de Retorno:")
plataforma = st.text_input("Plataforma:")
responsavel = st.text_input("Responsável:")
ddd = st.text_input("DDD:")
telefone = st.text_input("Telefone do Responsável:")

if st.button("⚙️ Gerar Crachás") and df_file:
    try:
        df = pd.read_excel(df_file)
        temp_zip = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
        with zipfile.ZipFile(temp_zip.name, 'w') as zipf:
            for idx, linha in df.iterrows():
                nome = str(linha.get("Nome Completo", "")).strip()
                unidade = str(linha.get("Unidade", "")).strip()
                possui_plano = str(linha.get("Possuí Plano de Saúde?", "")).strip().lower() in ["sim", "yes"]
                nome_plano = str(linha.get("Nome do Plano", "")).strip()
                telefone_plano = str(linha.get("Telefone do Plano", "")).strip()
                contato_nome = str(linha.get("Nome Contato de Emergência", "")).strip()
                contato_telefone = str(linha.get("Telefone Contato de Emergência", "")).strip()
                tipo_contato = str(linha.get("Tipo de Contato", "")).strip().lower()

                doc = fitz.open(TEMPLATE_PDF)
                page = doc[0]
                page.set_textbox(fitz.Rect(70, 75, 400, 100), nome)
                page.set_textbox(fitz.Rect(70, 110, 400, 135), unidade)
                page.set_textbox(fitz.Rect(155, 145, 400, 165), horario)
                page.set_textbox(fitz.Rect(155, 170, 400, 195), plataforma)
                page.set_textbox(fitz.Rect(230, 195, 500, 215), responsavel)
                page.set_textbox(fitz.Rect(145, 220, 500, 245), f"({ddd}) {telefone}")

                if possui_plano:
                    page.set_textbox(fitz.Rect(170, 245, 185, 255), "X")
                else:
                    page.set_textbox(fitz.Rect(240, 245, 255, 255), "X")

                page.set_textbox(fitz.Rect(110, 270, 400, 290), nome_plano)
                page.set_textbox(fitz.Rect(170, 295, 400, 315), telefone_plano)

                if tipo_contato == "familiar":
                    page.set_textbox(fitz.Rect(160, 320, 180, 330), "X")
                elif tipo_contato == "amigo":
                    page.set_textbox(fitz.Rect(260, 320, 280, 330), "X")

                page.set_textbox(fitz.Rect(85, 345, 400, 370), contato_nome)
                page.set_textbox(fitz.Rect(100, 370, 400, 395), contato_telefone)

                output_name = f"Cracha_{idx+1}_{nome.replace(' ', '_')}.pdf"
                temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                doc.save(temp_pdf.name)
                doc.close()

                zipf.write(temp_pdf.name, arcname=output_name)

        with open(temp_zip.name, "rb") as f:
            st.success("✅ Crachás gerados com sucesso!")
            st.download_button("📁 Baixar ZIP dos Crachás", f.read(), file_name="crachas.zip")

    except Exception as e:
        st.error(f"❌ Erro ao gerar crachás: {e}")
