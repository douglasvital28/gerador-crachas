
import fitz
import os
import pandas as pd

TEMPLATE = "crachá caravana - template.pdf"
EXCEL = "Planilha Input.xlsx"
PASTA_SAIDA = "crachas_preenchidos"
os.makedirs(PASTA_SAIDA, exist_ok=True)

# Perguntas ao usuário
print("Informe os valores padrão para todos os crachás:")
horario_retorno = input("⏰ Horário de Retorno (ex: 22h00): ").strip()
plataforma = input("🚌 Plataforma (ex: Rodoviária - 12): ").strip()
responsavel = input("👤 Responsável da Caravana: ").strip()
ddd_resp = input("📞 DDD do Responsável (ex: 21): ").strip()
telefone_resp = input("📱 Telefone do Responsável (ex: 99999-9999): ").strip()

df = pd.read_excel(EXCEL)
df.columns = df.columns.str.strip()
grupos = [df[i:i + 4] for i in range(0, len(df), 4)]

def seguro(valor):
    return "" if pd.isna(valor) or valor is None else str(valor).strip()

def preencher_pdf(grupo_df, grupo_idx):
    doc = fitz.open(TEMPLATE)
    page = doc[0]
    widgets = list(page.widgets())

    for i, (_, pessoa) in enumerate(grupo_df.iterrows()):
        sufixo = "" if i == 0 else f" {i+1}"

        campos = {
            f"NOME{sufixo}": seguro(pessoa.get("Nome Completo")),
            f"IGREJA{sufixo}": seguro(pessoa.get("Unidade")),
            f"POLTRONA{sufixo}": seguro(pessoa.get("Poltrona")),
            f"HORÁRIO DE RETORNO{sufixo}": horario_retorno,
            f"PLATAFORMA{sufixo}": plataforma,
            f"RESPONSÁVEL DA CARAVANA{sufixo}": responsavel,
            f"DDD DO RESP{sufixo}": ddd_resp,
            f"TELEFONE DO RESPONSÁVEL{sufixo}": telefone_resp,
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

    out_path = os.path.join(PASTA_SAIDA, f"Cracha_Grupo_{grupo_idx+1}.pdf")
    doc.save(out_path, incremental=False)
    doc.close()
    print(f"✅ Cracha salvo: {out_path}")

for idx, grupo in enumerate(grupos):
    preencher_pdf(grupo, idx)

print("\n🎉 Todos os crachás foram preenchidos com sucesso!")
