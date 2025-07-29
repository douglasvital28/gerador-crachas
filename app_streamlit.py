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
            grupo = df.iloc[i:i+4]

            for idx, linha in grupo.iterrows():
                slot = idx % 4

                def write(rect, text):
                    page.insert_textbox(fitz.Rect(*rect), str(text), fontsize=10, fontname="helv")

                nome = linha.get("Nome Completo", "")
                unidade = linha.get("Unidade", "")
                plano = str(linha.get("Possuí Plano de Saúde?", "")).strip().lower() in ["sim", "yes"]
                nome_plano = linha.get("Nome do Plano", "")
                telefone_plano = linha.get("Telefone do Plano", "")
                nome_contato = linha.get("Nome Contato de Emergência", "")
                telefone_contato = linha.get("Telefone Contato de Emergência", "")
                tipo_contato = linha.get("Tipo de Contato", "").strip().lower()

                base_y = [75, 270, 470, 665][slot]  # Y base para cada crachá

                write((70, base_y, 400, base_y+20), nome)
                write((70, base_y+35, 400, base_y+55), unidade)
                write((155, base_y+70, 400, base_y+90), horario)
                write((155, base_y+95, 400, base_y+115), plataforma)
                write((230, base_y+120, 500, base_y+140), responsavel)
                write((145, base_y+145, 500, base_y+165), f"({ddd}) {telefone}")

                if plano:
                    write((170, base_y+170, 180, base_y+180), "X")  # Sim
                else:
                    write((240, base_y+170, 250, base_y+180), "X")  # Não

                write((110, base_y+195, 400, base_y+215), nome_plano)
                write((170, base_y+220, 400, base_y+240), telefone_plano)

                if tipo_contato == "familiar":
                    write((160, base_y+245, 180, base_y+255), "X")
                elif tipo_contato == "amigo":
                    write((260, base_y+245, 280, base_y+255), "X")

                write((85, base_y+270, 400, base_y+290), nome_contato)
                write((100, base_y+295, 400, base_y+315), telefone_contato)

        output = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        final_pdf.save(output.name)
        final_pdf.close()

        with open(output.name, "rb") as f:
            st.success("✅ Crachás gerados com sucesso!")
            st.download_button("📁 Baixar PDF com os Crachás", f.read(), file_name="crachas_caravana.pdf")

    except Exception as e:
        st.error(f"❌ Erro ao gerar crachás: {e}")
