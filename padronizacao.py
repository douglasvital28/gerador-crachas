"""Padronização dos dados visíveis sem alterar identificadores de inscrições."""
CAMPOS_MAIUSCULOS = (
    'Nome Completo', 'Unidade', 'Poltrona', 'Nome do Plano', 'Telefone do Plano',
    'Nome Contato de Emergência', 'Telefone Contato de Emergência',
)


def maiusculas_participantes(df):
    resultado = df.copy(deep=True)
    for campo in CAMPOS_MAIUSCULOS:
        if campo in resultado:
            resultado[campo] = resultado[campo].fillna('').map(lambda v: str(v).strip().upper())
    return resultado


def ordenar_por_missao(df):
    """Ordena linhas inteiras por missão e nome, mantendo IDs e assentos juntos."""
    import re
    import unicodedata
    def chave_texto(valor):
        texto = unicodedata.normalize('NFD', str(valor).strip().casefold())
        texto = ''.join(c for c in texto if not unicodedata.combining(c))
        return ' '.join(texto.replace('(a)', '').replace('(as)', '').split())
    grupos = (
        ('ministro', 'ministra', 'ministros', 'ministras'),
        ('professor', 'professora', 'professores', 'professoras'),
        ('coordenador', 'coordenadora', 'coordenadores', 'coordenadoras'),
        ('assistente de ministro', 'assistente de ministra', 'assistentes de ministros', 'assistente de ministros', 'assistentes de ministro'),
        ('responsavel de setor', 'responsaveis de setor', 'responsavel do setor', 'responsaveis do setor'),
        ('assistente de familia', 'assistentes de familia'),
        ('membro', 'membros'),
        ('frequentador', 'frequentadora', 'frequentadores', 'frequentadoras'),
        ('primeira vez',),
    )
    prioridades = {nome: i for i, grupo in enumerate(grupos) for nome in grupo}
    coluna = next((c for c in df if chave_texto(c) == 'missao'), None)
    def chave(i):
        nome = chave_texto(df.iloc[i]['Nome Completo'])
        missao = chave_texto(df.iloc[i][coluna]) if coluna else ''
        prioridade = prioridades.get(missao)
        if prioridade is None:
            titulo = re.match(r'^(min\.?|ministro|ministra|prof\.?|professor|professora)\s+', nome)
            prioridade = (0 if titulo.group(1).startswith('min') else 1) if titulo else len(grupos)
        nome = re.sub(r'^(?:(?:min\.?|ministro|ministra|prof\.?|professor|professora)\s+)+', '', nome)
        return prioridade, nome, i
    indices = sorted(range(len(df)), key=chave)
    resultado = df.iloc[indices].copy().reset_index(drop=True)
    resultado.attrs = df.attrs.copy()
    if 'linhas_planilha' in df.attrs:
        resultado.attrs['linhas_planilha'] = [df.attrs['linhas_planilha'][i] for i in indices]
    return resultado
