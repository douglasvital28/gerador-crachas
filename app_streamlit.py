import streamlit as st
import pandas as pd
import tempfile
import fitz
import os

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
        template = fitz.open(TEMPLATE_PATH)
        final_pdf = fitz.open()

        for i in range(0, len(df), 4):
            final_pdf.insert_pdf(template, from_page=0, to_page=0)
            page = final_pdf[-1]  # última página adicionada
            grupo = df.iloc[i:i+4].reset_index(drop=True)

            for idx, linha in grupo.iterrows():
                sufixo = "" if idx == 0 else f" {idx+1}"

                valores = {
                    f"NOME{sufixo}": linha.get("Nome Completo", ""),
                    f"IGREJA{sufixo}": linha.get("Unidade", ""),
                    f"HORÁRIO DE RETORNO{sufixo}": horario,
                    f"PLATAFORMA{sufixo}": plataforma,
                    f"RESPONSÁVEL DA CARAVANA{sufixo}": responsavel,
                    f"DDD DO RESP{sufixo}": ddd,
                    f"TELEFONE DO RESPONSÁVEL{sufixo}": telefone,
                    f"CONVÊNIO MÉDICO{sufixo}": linha.get("Nome do Plano", ""),
                    f"TELEFONE DO CONVÊNIO{sufixo}": linha.get("Telefone do Plano", ""),
                    f"NOME DO CONTATO{sufixo}": linha.get("Nome Contato de Emergência", ""),
                    f"TELEFONE DO CONTATO{sufixo}": linha.get("Telefone Contato de Emergência", ""),
                    f"POLTRONA{sufixo}": linha.get("Poltrona", "")
                }

                # Checkbox: possui plano
                possui_plano = str(linha.get("Possuí Plano de Saúde?", "")).strip().lower() in ["sim", "yes"]
                valores[f"Possui Convenio {idx+1 if idx > 0 else 1}"] = "yes" if possui_plano else "Off"

                # Checkbox: tipo de contato
                tipo = str(linha.get("Tipo de Contato", "")).strip().lower()
                if tipo == "familiar":
                    valores[f"Tipo do Contato {idx+1 if idx > 0 else 1}"] = "familar"
                elif tipo == "amigo":
                    valores[f"Tipo do Contato {idx+1 if idx > 0 else 1}"] = "amigo"

                # Preencher os campos da página
                for widget in page.widgets():
                    nome = widget.field_name.strip()
                    if nome in valores:
                        widget.field_value = str(valores[nome])
                        widget.update()

        output = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        final_pdf.save(output.name)
        final_pdf.close()

        with open(output.name, "rb") as f:
            st.success("✅ Crachás gerados com sucesso!")
            st.download_button("📁 Baixar PDF com os Crachás", f.read(), file_name="crachas_caravana.pdf")

    except Exception as e:
        st.error(f"❌ Erro ao gerar crachás: {e}")