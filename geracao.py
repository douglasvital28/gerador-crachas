"""Validação da planilha e geração de crachás sem arquivos temporários."""
from io import BytesIO
from pathlib import Path
import pandas as pd
import pymupdf as fitz

BASE = Path(__file__).resolve().parent
MAX_PARTICIPANTES = 400
MAX_UPLOAD_BYTES = 5 * 1024 * 1024
COLUNAS = (
    'Nome Completo', 'Unidade', 'Poltrona', 'Nome do Plano',
    'Telefone do Plano', 'Nome Contato de Emergência',
    'Telefone Contato de Emergência',
)


def ler_planilha(conteudo):
    if len(conteudo) > MAX_UPLOAD_BYTES:
        raise ValueError('A planilha deve ter no máximo 5 MB.')
    try:
        df = pd.read_excel(BytesIO(conteudo), dtype=str, keep_default_na=False)
    except Exception as exc:
        raise ValueError('Não foi possível ler o arquivo. Envie uma planilha Excel válida (.xlsx).') from exc
    df.columns = [str(c).strip() for c in df.columns]
    if df.columns.duplicated().any():
        raise ValueError('Existem colunas com nomes repetidos. Use a planilha modelo.')
    faltantes = [c for c in COLUNAS if c not in df.columns]
    if faltantes:
        raise ValueError('Colunas não encontradas: ' + ', '.join(faltantes) + '. Use a planilha modelo.')
    df = df.fillna('').map(lambda v: str(v).strip())
    df = df.loc[df.ne('').any(axis=1)]
    if df.empty:
        raise ValueError('A planilha não contém participantes. Preencha a planilha modelo.')
    sem_nome = df.index[df['Nome Completo'].eq('')]
    if len(sem_nome):
        raise ValueError('Preencha Nome Completo nas linhas: ' + ', '.join(str(i + 2) for i in sem_nome) + '.')
    if len(df) > MAX_PARTICIPANTES:
        raise ValueError(f'Gere até {MAX_PARTICIPANTES} participantes por vez. Divida a planilha em lotes menores.')
    return df.reset_index(drop=True)


def gerar_pdf(df, horario='', plataforma='', responsavel='', ddd='', telefone=''):
    if not 1 <= len(df) <= MAX_PARTICIPANTES:
        raise ValueError('Quantidade de participantes inválida.')
    with fitz.open() as saida:
        for inicio in range(0, len(df), 4):
            with fitz.open(BASE / 'cracha_template.pdf') as doc:
                page = doc[0]
                widgets = list(page.widgets())
                # O modelo contém uma opção marcada por padrão: limpar os slots vazios.
                for widget in widgets:
                    if widget.field_type == fitz.PDF_WIDGET_TYPE_TEXT:
                        widget.field_value = ''
                        widget.update()
                    else:
                        doc.xref_set_key(widget.xref, 'V', '/Off')
                        doc.xref_set_key(widget.xref, 'AS', '/Off')
                for i, (_, pessoa) in enumerate(df.iloc[inicio:inicio + 4].iterrows()):
                    sufixo = '' if i == 0 else f' {i+1}'
                    campos = {
                        f'NOME{sufixo}': pessoa['Nome Completo'],
                        f'IGREJA{sufixo}': pessoa['Unidade'],
                        f'POLTRONA{sufixo}': pessoa['Poltrona'],
                        f'HORÁRIO DE RETORNO{sufixo}': horario,
                        f'PLATAFORMA{sufixo}': plataforma,
                        f'RESPONSÁVEL DA CARAVANA{sufixo}': responsavel,
                        f'DDD DO RESP{sufixo}': ddd,
                        f'TELEFONE DO RESPONSÁVEL{sufixo}': telefone,
                        f'CONVÊNIO MÉDICO{sufixo}': pessoa['Nome do Plano'],
                        f'TELEFONE DO CONVÊNIO{sufixo}': pessoa['Telefone do Plano'],
                        f'NOME DO CONTATO{sufixo}': pessoa['Nome Contato de Emergência'],
                        f'TELEFONE DO CONTATO{sufixo}': pessoa['Telefone Contato de Emergência'],
                    }
                    existentes = {w.field_name for w in widgets}
                    if campos.keys() - existentes:
                        raise ValueError('O modelo PDF não contém todos os campos necessários.')
                    for widget in widgets:
                        if widget.field_name in campos:
                            valor = campos[widget.field_name]
                            # Helvetica é suportada pelo formulário e permite medir o texto.
                            widget.text_font = 'Helv'
                            largura = fitz.get_text_length(valor, fontname='helv', fontsize=12)
                            widget.text_fontsize = min(12, 12 * max(1, widget.rect.width - 6) / max(1, largura))
                            if widget.text_fontsize < 6:
                                raise ValueError(f'Texto muito longo no campo {widget.field_name}: revise ou abrevie o conteúdo.')
                            widget.field_value = valor
                            widget.update()
                saida.insert_pdf(doc, widgets=True, join_duplicates=False)
        return saida.tobytes(garbage=4, deflate=True)
