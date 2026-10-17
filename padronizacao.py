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
