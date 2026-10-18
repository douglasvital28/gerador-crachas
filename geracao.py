"""Validação da planilha e geração de crachás sem arquivos temporários."""
from io import BytesIO
from pathlib import Path
import unicodedata
import re
from trajetos import preparar_trajetos, conflitam
from colo import eh_colo, preparar_colo
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


def separar_telefone(valor):
    """Separa números brasileiros completos; preserva números ambíguos e serviços."""
    original = str(valor).strip()
    if not original or not re.fullmatch(r'[+\d\s().-]+', original):
        return '', original
    numero = re.sub(r'\D', '', original)
    if numero.startswith(('0800', '0300', '0500', '0900')):
        return '', original
    if original.startswith('+'):
        if not numero.startswith('55') or len(numero) not in (12, 13):
            return '', original
        numero = numero[2:]
    elif numero.startswith('0055') and len(numero) in (14, 15):
        numero = numero[4:]
    elif numero.startswith('55') and len(numero) in (12, 13):
        numero = numero[2:]
    elif numero.startswith('0') and len(numero) in (11, 12):
        numero = numero[1:]
    if len(numero) not in (10, 11) or numero[0] == '0':
        return '', original
    return numero[:2], numero[2:]


def avisos_duplicidades(df):
    """Aponta possíveis repetições sem excluir ou alterar participantes."""
    avisos = []
    linhas = df.attrs.get('linhas_planilha', list(range(2, len(df) + 2)))
    for coluna in ('Nome Completo', 'Poltrona'):
        chaves = df[coluna].fillna('').map(lambda v: ' '.join(normalizar(v).split()))
        if coluna == 'Poltrona':
            chaves = chaves.where(~df.apply(eh_colo, axis=1), '')
            chaves = chaves.map(lambda v: str(int(v)) if v.isdecimal() else v)
        for chave in dict.fromkeys(chaves[chaves.ne('') & chaves.duplicated(keep=False)]):
            posicoes = [i for i, v in enumerate(chaves) if v == chave]
            if coluna == 'Poltrona' and not conflitam(df, posicoes):
                continue
            avisos.append({'Campo': coluna, 'Valor': str(df.iloc[posicoes[0]][coluna]),
                           'Linhas da planilha': ', '.join(str(linhas[i]) for i in posicoes)})
    return avisos


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


def aplicar_missao(df):
    """Acrescenta títulos a partir da coluna opcional Missão, sem duplicá-los."""
    resultado = df.copy(deep=True)
    colunas = [c for c in resultado.columns if normalizar(c) == 'missao']
    if not colunas:
        return resultado
    if len(colunas) > 1:
        raise ValueError('Existe mais de uma coluna Missão. Mantenha apenas uma.')
    def nome_com_titulo(linha):
        nome = str(linha['Nome Completo']).strip()
        missao = normalizar(linha[colunas[0]])
        prefixo = {'ministro': 'MIN.', 'ministra': 'MIN.', 'ministro(a)': 'MIN.',
                   'professor': 'PROF.', 'professora': 'PROF.', 'professor(a)': 'PROF.'}.get(missao)
        primeiro = nome.split()[0].casefold().rstrip('.') if nome else ''
        if nome and prefixo and primeiro not in {'min', 'ministro', 'ministra', 'prof', 'professor', 'professora'}:
            return prefixo + ' ' + nome
        return nome
    resultado['Nome Completo'] = resultado.apply(nome_com_titulo, axis=1)
    return resultado


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
    if 'Poltrona' not in df:
        df['Poltrona'] = ''
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
    resultado = preparar_colo(preparar_trajetos(aplicar_missao(validar_opcoes(df)))).reset_index(drop=True)
    resultado.attrs['linhas_planilha'] = [int(i) + 2 for i in df.index]
    return resultado


def abreviar_nome_pdf(nome, largura):
    """Abrevia apenas o meio do nome quando não cabe na fonte normal de 12 pt."""
    partes = str(nome).split()
    titulos = {'min', 'ministra', 'ministro', 'prof', 'professora'}
    primeiro = 0
    while primeiro < len(partes) and partes[primeiro].casefold().rstrip('.') in titulos:
        primeiro += 1
    candidatos = [i for i in range(primeiro+1, len(partes)-1)
                  if partes[i].casefold() not in {'da', 'de', 'do', 'das', 'dos', 'e'}
                  and len(partes[i].rstrip('.')) > 1]
    resultado = ' '.join(partes)
    for i in sorted(candidatos, key=lambda i: fitz.get_text_length(partes[i], fontname='helv', fontsize=12), reverse=True):
        if fitz.get_text_length(resultado, fontname='helv', fontsize=12) <= largura:
            break
        partes[i] = partes[i][0] + '.'
        resultado = ' '.join(partes)
    return resultado


def gerar_pdf(df, horario='', plataforma='', responsavel='', ddd='', telefone=''):
    if not 1 <= len(df) <= MAX_PARTICIPANTES:
        raise ValueError('Quantidade de participantes inválida.')
    df = preparar_colo(validar_opcoes(df))
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
                    ddd_plano, telefone_plano = separar_telefone(pessoa['Telefone do Plano'])
                    ddd_contato, telefone_contato = separar_telefone(pessoa['Telefone Contato de Emergência'])
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
                        f'DDD Convenio{sufixo}': ddd_plano,
                        f'TELEFONE DO CONVÊNIO{sufixo}': telefone_plano,
                        f'NOME DO CONTATO{sufixo}': pessoa['Nome Contato de Emergência'],
                        f'DDD Contato{sufixo}': ddd_contato,
                        f'TELEFONE DO CONTATO{sufixo}': telefone_contato,
                    }
                    existentes = {w.field_name for w in widgets}
                    if campos.keys() - existentes:
                        raise ValueError('O modelo PDF não contém todos os campos necessários.')
                    for widget in widgets:
                        if widget.field_name in campos:
                            valor = campos[widget.field_name]
                            if widget.field_name in (f'NOME{sufixo}', f'NOME DO CONTATO{sufixo}', f'RESPONSÁVEL DA CARAVANA{sufixo}'):
                                valor = abreviar_nome_pdf(valor, max(1, widget.rect.width - 6))
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
