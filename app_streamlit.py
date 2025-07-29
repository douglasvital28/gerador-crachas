import streamlit as st
import pandas as pd
import tempfile
import os
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject, ArrayObject

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
        output_pdf = PdfWriter()

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

                    def val(col):
                        v = linha.get(col, "")
                        return "" if pd.isna(v) else str(v)

                    data_dict.update({
                        f"NOME{sufixo}".strip(): val("Nome Completo"),
                        f"IGREJA{sufixo}".strip(): val("Unidade"),
                        f"HORÁRIO DE RETORNO{sufixo}".strip(): horario,
                        f"PLATAFORMA{sufixo}".strip(): plataforma,
                        f"RESPONSÁVEL DA CARAVANA{sufixo}".strip(): responsavel,
                        f"DDD DO RESP{sufixo}".strip(): ddd,
                        f"TELEFONE DO RESPONSÁVEL{sufixo}".strip(): telefone,
                        f"CONVÊNIO MÉDICO{sufixo}".strip(): val("Nome do Plano"),
                        f"TELEFONE DO CONVÊNIO{sufixo}".strip(): val("Telefone do Plano"),
                        f"NOME DO CONTATO{sufixo}".strip(): val("Nome Contato de Emergência"),
                        f"TELEFONE DO CONTATO{sufixo}".strip(): val("Telefone Contato de Emergência"),
                        f"POLTRONA{sufixo}".strip(): val("Poltrona")
                    })

                    plano = str(linha.get("Possuí Plano de Saúde?", "")).strip().lower() in ["sim", "yes"]
                    data_dict[f"Possui Convenio {idx+1 if idx > 0 else 1}".strip()] = "yes" if plano else "Off"

                    tipo = str(linha.get("Tipo de Contato", "")).strip().lower()
                    if tipo == "familiar":
                        data_dict[f"Tipo do Contato {idx+1 if idx > 0 else 1}".strip()] = "familar"
                    elif tipo == "amigo":
                        data_dict[f"Tipo do Contato {idx+1 if idx > 0 else 1}".strip()] = "amigo"

                writer.update_page_form_field_values(writer.pages[0], data_dict)

                for j in writer.pages:
                    j[NameObject("/Annots")] = ArrayObject()

                temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                with open(temp_pdf.name, "wb") as tf:
                    writer.write(tf)
                with open(temp_pdf.name, "rb") as tf:
                    output_pdf.add_page(PdfReader(tf).pages[0])

        result_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        with open(result_pdf.name, "wb") as f:
            output_pdf.write(f)

        with open(result_pdf.name, "rb") as f:
            st.success("✅ Crachás gerados com sucesso!")
            st.download_button("📁 Baixar PDF com os Crachás", f.read(), file_name="crachas_caravana.pdf")

    except Exception as e:
        st.error(f"❌ Erro ao gerar crachás: {e}")