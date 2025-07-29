import streamlit as st
import pandas as pd
import tempfile
import os
from PyPDF2 import PdfReader, PdfWriter

TEMPLATE_PATH = "cracha_template.pdf"

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
        output_writer = PdfWriter()

        for i in range(0, len(df), 4):
            grupo = df.iloc[i:i+4].reset_index(drop=True)

            with open(TEMPLATE_PATH, "rb") as f:
                reader = PdfReader(f)
                page = reader.pages[0]
                writer = PdfWriter()
                writer.add_page(page)

                data_dict = {}

                for idx, linha in grupo.iterrows():
                    sufixo = "" if idx == 0 else f" {idx+1}"

                    def safe(col):
                        v = linha.get(col, "")
                        return "" if pd.isna(v) else str(v)

                    data_dict.update({
                        f"NOME{sufixo}": safe("Nome Completo"),
                        f"IGREJA{sufixo}": safe("Unidade"),
                        f"HORÁRIO DE RETORNO{sufixo}": horario,
                        f"PLATAFORMA{sufixo}": plataforma,
                        f"RESPONSÁVEL DA CARAVANA{sufixo}": responsavel,
                        f"DDD DO RESP{sufixo}": ddd,
                        f"TELEFONE DO RESPONSÁVEL{sufixo}": telefone,
                        f"CONVÊNIO MÉDICO{sufixo}": safe("Nome do Plano"),
                        f"TELEFONE DO CONVÊNIO{sufixo}": safe("Telefone do Plano"),
                        f"NOME DO CONTATO{sufixo}": safe("Nome Contato de Emergência"),
                        f"TELEFONE DO CONTATO{sufixo}": safe("Telefone Contato de Emergência"),
                        f"POLTRONA{sufixo}": safe("Poltrona")
                    })

                    plano = str(linha.get("Possuí Plano de Saúde?", "")).strip().lower()
                    data_dict[f"Possui Convenio {idx+1 if idx > 0 else 1}"] = "yes" if plano in ["sim", "yes"] else "Off"

                    tipo = str(linha.get("Tipo de Contato", "")).strip().lower()
                    if tipo == "familiar":
                        data_dict[f"Tipo do Contato {idx+1 if idx > 0 else 1}"] = "familar"
                    elif tipo == "amigo":
                        data_dict[f"Tipo do Contato {idx+1 if idx > 0 else 1}"] = "amigo"

                writer.update_page_form_field_values(writer.pages[0], data_dict)

                temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                with open(temp_pdf.name, "wb") as tf:
                    writer.write(tf)

                filled_reader = PdfReader(temp_pdf.name)
                output_writer.add_page(filled_reader.pages[0])

        result_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        with open(result_pdf.name, "wb") as f:
            output_writer.write(f)

        with open(result_pdf.name, "rb") as f:
            st.success("✅ Crachás gerados com sucesso!")
            st.download_button("📁 Baixar PDF com os Crachás", f.read(), file_name="crachas_caravana.pdf")

    except Exception as e:
        st.error(f"❌ Erro ao gerar crachás: {e}")
