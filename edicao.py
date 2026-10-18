"""Validação das correções feitas na tabela de participantes."""
from geracao import validar_opcoes
from trajetos import preparar_trajetos
from colo import preparar_colo


def validar_edicao(df):
    resultado = df.fillna('').map(lambda v: str(v).strip())
    linhas = df.attrs.get('linhas_planilha', list(range(2, len(df) + 2)))
    vazios = [str(linhas[i]) for i, nome in enumerate(resultado['Nome Completo']) if not nome]
    if vazios:
        raise ValueError('Preencha Nome Completo na tabela para as linhas da planilha: '
                         + ', '.join(vazios) + '.')
    resultado.attrs = df.attrs.copy()
    return preparar_colo(preparar_trajetos(validar_opcoes(resultado)))
