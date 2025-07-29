import fitz  # PyMuPDF
import pandas as pd
import os

def preencher_crachas_em_lote(
    excel_path,
    output_folder,
    pdf_modelo,
    horario,
    plataforma,
    responsavel,
    ddd,
    telefone
):
    df = pd.read_excel(excel_path)
    
    for idx, linha in df.iterrows():
        nome = str(linha.get("Nome Completo", "")).strip()
        unidade = str(linha.get("Unidade", "")).strip()
        possui_plano = str(linha.get("Possuí Plano de Saúde?", "")).strip().lower() in ["sim", "yes"]
        nome_plano = str(linha.get("Nome do Plano", "")).strip()
        telefone_plano = str(linha.get("Telefone do Plano", "")).strip()
        contato_nome = str(linha.get("Nome Contato de Emergência", "")).strip()
        contato_telefone = str(linha.get("Telefone Contato de Emergência", "")).strip()
        tipo_contato = str(linha.get("Tipo de Contato", "")).strip().lower()

        doc = fitz.open(pdf_modelo)
        page = doc[0]
        
        # Preencher os campos
        page.set_textbox(fitz.Rect(70, 75, 400, 100), nome)
        page.set_textbox(fitz.Rect(70, 110, 400, 135), unidade)
        page.set_textbox(fitz.Rect(155, 145, 400, 165), horario)
        page.set_textbox(fitz.Rect(155, 170, 400, 195), plataforma)
        page.set_textbox(fitz.Rect(230, 195, 500, 215), responsavel)
        page.set_textbox(fitz.Rect(145, 220, 500, 245), f"({ddd}) {telefone}")

        # Convênio médico (sim/não checkbox)
        if possui_plano:
            page.set_textbox(fitz.Rect(170, 245, 185, 255), "X")  # Sim
        else:
            page.set_textbox(fitz.Rect(240, 245, 255, 255), "X")  # Não

        page.set_textbox(fitz.Rect(110, 270, 400, 290), nome_plano)
        page.set_textbox(fitz.Rect(170, 295, 400, 315), telefone_plano)

        # Contato de emergência (familiar ou amigo)
        if tipo_contato == "familiar":
            page.set_textbox(fitz.Rect(160, 320, 180, 330), "X")
        elif tipo_contato == "amigo":
            page.set_textbox(fitz.Rect(260, 320, 280, 330), "X")

        page.set_textbox(fitz.Rect(85, 345, 400, 370), contato_nome)
        page.set_textbox(fitz.Rect(100, 370, 400, 395), contato_telefone)

        output_path = os.path.join(output_folder, f"Cracha_{idx+1}_{nome.replace(' ', '_')}.pdf")
        doc.save(output_path)
        doc.close()

    return True
