import streamlit as st
import pandas as pd
import fitz
import tempfile
import os

TEMPLATE_PATH = "cracha_template.pdf"
PLANILHA_MODELO = "modelo_planilha.xlsx"

st.set_page_config(page_title="Gerador de Crachás", page_icon="🪪")
st.title("🪪 Gerador de Crachás - Caravanas ao Solo Sagrado de Guarapiranga")

# Botão para baixar a planilha modelo
with open(PLANILHA_MODELO, "rb") as modelo:
    st.download_button(
        label="📄 Baixar Planilha Modelo",
        data=modelo,
        file_name="modelo_planilha.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Instruções para o usuário
st.markdown("""
### 📝 Como utilizar este gerador:

1. 📄 Clique no botão acima **"Baixar Planilha Modelo"** para obter o Excel padrão
2. ✍️ Preencha a planilha com os dados dos caravanistas
3. 📤 Envie a planilha preenchida no campo abaixo
4. ⌨️ Preencha os campos padrão como horário, plataforma e responsável
5. ⚙️ Clique em **"Gerar Crachás"** para baixar o PDF pronto para impressão

---
ℹ️ Dica: os campos em branco na planilha serão deixados vazios no crachá.
""")

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
        df.columns = df.columns.str.strip()
        grupos = [df[i:i + 4] for i in range(0, len(df), 4)]

        def seguro(valor):
            return "" if pd.isna(valor) or valor is None else str(valor).strip()

        pdf_saida = fitz.open()

        for grupo_idx, grupo_df in enumerate(grupos):
            doc = fitz.open(TEMPLATE_PATH)
            page = doc[0]
            widgets = list(page.widgets())

            for i, (_, pessoa) in enumerate(grupo_df.iterrows()):
                sufixo = "" if i == 0 else f" {i+1}"
                campos = {
                    f"NOME{sufixo}": seguro(pessoa.get("Nome Completo")),
                    f"IGREJA{sufixo}": seguro(pessoa.get("Unidade")),
                    f"POLTRONA{sufixo}": seguro(pessoa.get("Poltrona")),
                    f"HORÁRIO DE RETORNO{sufixo}": horario,
                    f"PLATAFORMA{sufixo}": plataforma,
                    f"RESPONSÁVEL DA CARAVANA{sufixo}": responsavel,
                    f"DDD DO RESP{sufixo}": ddd,
                    f"TELEFONE DO RESPONSÁVEL{sufixo}": telefone,
                    f"CONVÊNIO MÉDICO{sufixo}": seguro(pessoa.get("Nome do Plano")),
                    f"TELEFONE DO CONVÊNIO{sufixo}": seguro(pessoa.get("Telefone do Plano")),
                    f"NOME DO CONTATO{sufixo}": seguro(pessoa.get("Nome Contato de Emergência")),
                    f"TELEFONE DO CONTATO{sufixo}": seguro(pessoa.get("Telefone Contato de Emergência"))
                }

                for widget in widgets:
                    nome = widget.field_name
                    if nome in campos:
                        widget.field_value = campos[nome]
                        widget.update()

            pdf_saida.insert_pdf(doc)
            doc.close()

        result_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        pdf_saida.save(result_pdf.name)
        pdf_saida.close()

        with open(result_pdf.name, "rb") as f:
            st.success("✅ Crachás gerados com sucesso!")
            st.download_button("📁 Baixar PDF com os Crachás", f.read(), file_name="crachas_caravana.pdf")

    except Exception as e:
        st.error(f"❌ Erro ao gerar crachás: {e}")
