"""Lista de passageiros para conferência em papel."""
from io import BytesIO
import re
from xml.sax.saxutils import escape
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Flowable
from geracao import MAX_PARTICIPANTES, normalizar


class Caixa(Flowable):
    def __init__(self):
        super().__init__()
        self.width = self.height = 10

    def draw(self):
        self.canv.setLineWidth(0.7)
        self.canv.rect(0, 0, 10, 10, stroke=1, fill=0)


def ordenar_participantes(df, ordem):
    if ordem == 'Missão e nome':
        from padronizacao import ordenar_por_missao
        return ordenar_por_missao(df)
    if ordem not in ('Nome', 'Poltrona'):
        raise ValueError('Escolha ordenação por Nome ou Poltrona.')
    def chave(i):
        pessoa = df.iloc[i]
        nome = normalizar(pessoa['Nome Completo'])
        if ordem == 'Nome':
            return nome, i
        poltrona = str(pessoa['Poltrona']).strip()
        partes = tuple((0, int(p)) if p.isdigit() else (1, normalizar(p))
                       for p in re.split(r'(\d+)', poltrona) if p)
        return not bool(poltrona), partes, nome, i
    return df.iloc[sorted(range(len(df)), key=chave)].copy()


def gerar_lista(df, ordem='Nome', data_viagem=None, responsavel=''):
    if not 1 <= len(df) <= MAX_PARTICIPANTES:
        raise ValueError('Selecione entre 1 e 400 participantes para gerar a lista.')
    df = ordenar_participantes(df, ordem)
    estilo = ParagraphStyle('Celula', fontName='Helvetica', fontSize=9, leading=12)
    centro = ParagraphStyle('Centro', parent=estilo, alignment=TA_CENTER)
    resp = Paragraph(escape(str(responsavel).strip()) or '________________________________________', estilo)
    largura = A4[0] - 24 * mm
    _, altura_resp = resp.wrap(largura - 70, 1000)
    if altura_resp > 60:
        raise ValueError('Nome do responsável muito longo para o cabeçalho da lista. Abrevie o conteúdo.')
    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm,
                            topMargin=75 + altura_resp, bottomMargin=36,
                            title='Lista de conferência de passageiros', author='Gerador de Crachás')
    larguras = [v * mm for v in (11, 53, 32, 13, 12, 16, 49)]
    dados = [['Nº', 'Nome', 'Unidade', 'Poltrona', 'Ida', 'Retorno', 'Observações']]
    alturas = [24]
    for indice, (_, pessoa) in enumerate(df.iterrows(), 1):
        linha = [Paragraph(str(indice), centro)]
        altura = 32
        for coluna, campo in enumerate(('Nome Completo', 'Unidade', 'Poltrona'), 1):
            texto = Paragraph(escape(str(pessoa[campo])), centro if campo == 'Poltrona' else estilo)
            _, h = texto.wrap(larguras[coluna] - 10, 10000)
            altura = max(altura, h + 12)
            linha.append(texto)
        if altura > A4[1] - doc.topMargin - doc.bottomMargin - 50:
            raise ValueError(f'Texto muito longo para a lista no participante {indice}. Abrevie nome, unidade ou poltrona.')
        dados.append([*linha, Caixa(), Caixa(), ''])
        alturas.append(altura)
    tabela = Table(dados, colWidths=larguras, rowHeights=alturas, repeatRows=1, hAlign='LEFT')
    tabela.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#eeeeee')),
        ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#777777')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('ALIGN', (4, 1), (5, -1), 'CENTER'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
    ]))

    def cabecalho(canvas, documento):
        canvas.saveState()
        canvas.setFont('Helvetica-Bold', 15)
        canvas.drawString(12*mm, A4[1] - 30, 'Lista de conferência de passageiros')
        canvas.setFont('Helvetica', 9)
        data = data_viagem.strftime('%d/%m/%Y') if data_viagem else '____/____/________'
        canvas.drawString(12*mm, A4[1] - 47, f'Data da viagem: {data}     Total de passageiros: {len(df)}')
        canvas.drawString(12*mm, A4[1] - 64, 'Responsável:')
        resp.drawOn(canvas, 12*mm + 70, A4[1] - 56 - altura_resp)
        canvas.setFont('Helvetica', 8)
        canvas.drawString(12*mm, 20, 'Marque os quadrados na ida e no retorno. Use Observações para anotações.')
        canvas.drawRightString(A4[0] - 12*mm, 20, f'Página {documento.page}')
        canvas.restoreState()

    doc.build([tabela], onFirstPage=cabecalho, onLaterPages=cabecalho)
    return output.getvalue()
