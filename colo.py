"""Crianças de colo acompanham um responsável sem ocupar assento."""
import re
import unicodedata


def chave(value):
    text=' '.join(''.join(c for c in unicodedata.normalize('NFD',str(value).casefold()) if not unicodedata.combining(c)).split())
    return re.sub(r'^(?:(?:min|ministro|ministra|prof|professor|professora)\.?\s+)+','',text)


def eh_colo(pessoa):
    return 'colo' in chave(pessoa.get('Modalidade','')) or chave(pessoa.get('Poltrona',''))=='colo'


def preparar_colo(df):
    result=df.copy(deep=True)
    for i in range(len(result)):
        if eh_colo(result.iloc[i]):result.iat[i,result.columns.get_loc('Poltrona')]='COLO'
    return result


def acompanhantes(tabela, responsavel):
    nome=chave(tabela.iloc[responsavel]['Nome Completo'])
    return [i for i in range(len(tabela)) if eh_colo(tabela.iloc[i]) and chave(tabela.iloc[i].get('Responsável acompanhante',''))==nome]
