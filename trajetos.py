"""Trajeto individual e compatibilidade de ocupação por sentido da viagem."""
import re
import unicodedata

TRAJETOS = ('IDA E VOLTA', 'SOMENTE IDA', 'SOMENTE VOLTA')


def normalizar(valor):
    return ' '.join(''.join(c for c in unicodedata.normalize('NFD', str(valor).casefold())
                            if not unicodedata.combining(c)).split())


def inferir_trajeto(observacoes):
    texto = normalizar(observacoes)
    if re.search(r'\bnao (?:so|somente|apenas)\b|\b(?:ida|vai) e (?:volta|retorno)\b', texto):
        return 'IDA E VOLTA'
    texto = re.sub(r'\b(so|somente|apenas) (?:vai|ira) (?:voltar|retornar)\b', r'\1 volta', texto)
    ida = bool(re.search(r'\b(?:so|somente|apenas) (?:de |na )?(?:vai|ira|ir|ida)\b|\b(?:vai|ira) (?:so|somente|apenas)\b', texto))
    volta = bool(re.search(r'\b(?:so|somente|apenas) (?:de |na )?(?:volta|voltara|voltar|retorno|retornar)\b|\b(?:volta|voltara) (?:so|somente|apenas)\b', texto))
    return 'SOMENTE IDA' if ida and not volta else 'SOMENTE VOLTA' if volta and not ida else 'IDA E VOLTA'


def preparar_trajetos(df):
    resultado = df.copy(deep=True)
    for destino, aliases in [('Observações', {'observacao', 'observacoes'}), ('Trajeto', {'trajeto'})]:
        colunas = [c for c in resultado if normalizar(c) in aliases]
        if len(colunas) > 1:
            raise ValueError(f'Existe mais de uma coluna {destino}. Mantenha apenas uma.')
        if colunas and colunas[0] != destino:
            resultado = resultado.rename(columns={colunas[0]: destino})
    if 'Trajeto' not in resultado:
        resultado['Trajeto'] = ''
    valores = []
    for _, pessoa in resultado.iterrows():
        valor = str(pessoa['Trajeto']).strip()
        if not valor:
            valor = inferir_trajeto(pessoa.get('Observações', ''))
        else:
            chave = normalizar(valor)
            equivalentes = {normalizar(t): t for t in TRAJETOS}
            equivalentes.update({'ida': 'SOMENTE IDA', 'volta': 'SOMENTE VOLTA'})
            if chave not in equivalentes:
                raise ValueError('Trajeto inválido: escolha Ida e volta, Somente ida ou Somente volta.')
            valor = equivalentes[chave]
        valores.append(valor)
    resultado['Trajeto'] = valores
    return resultado


def sentidos(pessoa):
    return {'IDA'} if pessoa.get('Trajeto') == 'SOMENTE IDA' else {'VOLTA'} if pessoa.get('Trajeto') == 'SOMENTE VOLTA' else {'IDA', 'VOLTA'}


def conflitam(tabela, indices):
    return any(sentidos(tabela.iloc[a]) & sentidos(tabela.iloc[b])
               for pos, a in enumerate(indices) for b in indices[pos+1:])
