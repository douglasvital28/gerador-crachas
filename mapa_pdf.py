"""Mapa para conferência manual, com nomes completos e uma página por piso."""
from io import BytesIO
from xml.sax.saxutils import escape
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph
from mapa_onibus import MAPAS, ocupacao, numero_poltrona

ROTULOS = {43: 'Leito 43', 44: 'LD Semi-leito 44', 46: 'Leito 46', 56: 'G8 Misto 56'}


def gerar_mapa_pdf(tabela, modelo):
    if modelo not in MAPAS:
        raise ValueError('Selecione um modelo de ônibus válido.')
    lugares = ocupacao(tabela)
    if any(len(indices) > 1 for indices in lugares.values()):
        raise ValueError('Há poltronas repetidas. Corrija antes de imprimir o mapa.')
    if any(str(v).strip() and numero_poltrona(v) not in range(1, modelo+1) for v in tabela['Poltrona']):
        raise ValueError('Há poltronas fora do modelo ou inválidas. Corrija antes de imprimir o mapa.')
    saida = BytesIO()
    pdf = canvas.Canvas(saida, pagesize=A4)
    pdf.setTitle('Mapa preenchido do ônibus')
    largura, altura = A4

    def cabecalho(subtitulo):
        pdf.setFillColorRGB(.07, .24, .27)
        pdf.setFont('Helvetica-Bold', 17)
        pdf.drawString(30, altura-38, 'MAPA DO ÔNIBUS - ' + ROTULOS[modelo])
        pdf.setFont('Helvetica-Bold', 11)
        pdf.drawString(30, altura-59, subtitulo)
        pdf.setFillColorRGB(0, 0, 0)

    def texto(corpo, x, y, w, h):
        for tamanho in (11, 10, 9, 8, 7):
            p = Paragraph(escape(corpo), ParagraphStyle('Nome', fontName='Helvetica-Bold',
                          fontSize=tamanho, leading=tamanho+1, alignment=1, splitLongWords=True))
            _, ph = p.wrap(w, h)
            if ph <= h:
                p.drawOn(pdf, x, y+(h-ph)/2)
                return
        raise ValueError('Um nome é muito longo para o mapa impresso. Abrevie-o antes de gerar.')

    for piso, filas in MAPAS[modelo].items():
        tipo = 'Semi-leito' if modelo == 44 or (modelo == 56 and piso == 'Piso superior') else 'Leito'
        cabecalho(tipo.upper() if modelo == 44 else piso.upper() + ' - ' + tipo.upper())
        pdf.setFont('Helvetica', 9)
        pdf.drawCentredString(largura/2, altura-81, 'FRENTE DO ÔNIBUS')
        duplas = len(filas[0]) == 4
        corredor = 24
        w = (largura-60-corredor)/(4 if duplas else 3)
        rh = min(110, (altura-145)/len(filas))
        for linha, fila in enumerate(filas):
            y = altura-98-(linha+1)*rh
            for coluna, assento in enumerate(fila):
                x = 30+coluna*w+(corredor if coluna >= 2 else 0)
                if assento == '':
                    continue
                pdf.setStrokeColorRGB(.35, .4, .4)
                pdf.setFillColorRGB(.96, .97, .97) if isinstance(assento, str) else pdf.setFillColorRGB(1, 1, 1)
                largura_celula = 2*w if isinstance(assento, str) and coluna+1 < len(fila) and fila[coluna+1] == '' else w
                pdf.rect(x, y, largura_celula, rh, fill=1)
                pdf.setFillColorRGB(0, 0, 0)
                if isinstance(assento, str):
                    texto(assento.upper(), x+4, y+4, largura_celula-8, rh-8)
                else:
                    pdf.setFont('Helvetica-Bold', 12)
                    pdf.drawString(x+6, y+rh-16, f'{assento:02d}')
                    indices = lugares.get(assento, [])
                    nome = str(tabela.iloc[indices[0]]['Nome Completo']).upper() if indices else 'LIVRE'
                    texto(nome, x+5, y+4, w-10, rh-24)
        pdf.setFont('Helvetica', 8)
        pdf.drawString(30, 28, 'Conferência manual: nomes completos em destaque. Assentos sem passageiro: LIVRE.')
        pdf.showPage()
    pendentes = [str(p['Nome Completo']).upper() for _, p in tabela.iterrows() if not str(p['Poltrona']).strip()]
    for inicio in range(0, len(pendentes), 25):
        cabecalho('PASSAGEIROS SEM POLTRONA - ' + str(len(pendentes)) + ' NO TOTAL')
        for i, nome in enumerate(pendentes[inicio:inicio+25]):
            y = altura-95-i*28
            texto(nome, 30, y-20, largura-60, 25)
        pdf.showPage()
    pdf.save()
    return saida.getvalue()
