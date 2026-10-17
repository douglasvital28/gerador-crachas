"""Validação da planilha e geração de crachás sem arquivos temporários."""
from io import BytesIO
from pathlib import Path
import unicodedata
import pandas as pd
import pymupdf as fitz

BASE = Path(__file__).resolve().parent
MAX_PARTICIPANTES = 400
MAX_UPLOAD_BYTES = 5 * 1024 * 1024
CONVENIO = 'Possuí Plano de Saúde?'
CONTATO = 'Tipo de Contato'
OPCOES = {CONVENIO: ('Sim', 'Não'), CONTATO: ('Familiar', 'Amigo')}
COLUNAS = (
    'Nome Completo', 'Unidade', 'Poltrona', 'Nome do Plano',
    'Telefone do Plano', 'Nome Contato de Emergência',
    'Telefone Contato de Emergência',
)


def normalizar(valor):
    return ''.join(c for c in unicodedata.normalize('NFD', str(valor).strip().casefold())
                   if not unicodedata.combining(c))


def validar_opcoes(df):
    df = df.copy()
    for coluna, opcoes in OPCOES.items():
        if coluna not in df:
            df[coluna] = ''
        aceitas = {normalizar(v): v for v in ('', *opcoes)}
        valores = df[coluna].fillna('').map(normalizar)
        invalidos = ~valores.isin(aceitas)
        if invalidos.any():
            linhas = ', '.join(str(i + 2) for i in df.index[invalidos])
            raise ValueError(f'Valores inválidos em {coluna} nas linhas {linhas}. '
                             f'Use {" ou ".join(opcoes)} ou deixe em branco.')
        df[coluna] = valores.map(aceitas)
    return df


def marcar_opcao(doc, widgets, nome, valor):
    grupo = [w for w in widgets if w.field_name == nome]
    if len(grupo) != 2:
        raise ValueError(f'O modelo PDF não contém o grupo de opções {nome}.')
    # Os estados do PDF usam nomes codificados (por exemplo, N#E3o).
    estado = {'Sim': 'Sim', 'Não': 'N#E3o', 'Familiar': 'Familiar',
              'Amigo': 'Amigo', '': 'Off'}[valor]
    disponiveis = {s for w in grupo for s in w.button_states()['normal']}
    if estado not in disponiveis:
        raise ValueError(f'O modelo PDF não contém a opção {valor}.')
    for widget in grupo:
        ativo = estado if estado in widget.button_states()['normal'] else 'Off'
        doc.xref_set_key(widget.xref, 'AS', '/' + ativo)
        # O valor pertence ao campo pai, compartilhado pelas duas opções.
        doc.xref_set_key(widget.xref, 'V', 'null')
        parent = int(doc.xref_get_key(widget.xref, 'Parent')[1].split()[0])
        doc.xref_set_key(parent, 'V', '/' + estado)


def ler_planilha(conteudo):
    if len(conteudo) > MAX_UPLOAD_BYTES:
        raise ValueError('A planilha deve ter no máximo 5 MB.')
    try:
        df = pd.read_excel(BytesIO(conteudo), dtype=str, keep_default_na=False)
    except Exception as exc:
        raise ValueError('Não foi possível ler o arquivo. Envie uma planilha Excel válida (.xlsx).') from exc
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={c: CONVENIO for c in df.columns
                            if normalizar(c) == normalizar(CONVENIO)})
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
    return validar_opcoes(df).reset_index(drop=True)


def gerar_pdf(df, horario='', plataforma='', responsavel='', ddd='', telefone=''):
    if not 1 <= len(df) <= MAX_PARTICIPANTES:
        raise ValueError('Quantidade de participantes inválida.')
    df = validar_opcoes(df)
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
                        doc.xref_set_key(widget.xref, 'V', 'null')
                        doc.xref_set_key(widget.xref, 'AS', '/Off')
                        parent = int(doc.xref_get_key(widget.xref, 'Parent')[1].split()[0])
                        doc.xref_set_key(parent, 'V', '/Off')
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
                    marcar_opcao(doc, widgets, f'Possui Convenio {i+1}', pessoa[CONVENIO])
                    nome_contato = f'Tipo Contato {i+1}' if i < 2 else f'Tipo do Contato {i+1}'
                    marcar_opcao(doc, widgets, nome_contato, pessoa[CONTATO])
                saida.insert_pdf(doc, widgets=True, join_duplicates=False)
        return saida.tobytes(garbage=4, deflate=True)
