"""Trajeto individual e compatibilidade de ocupação por sentido da viagem."""
import re
import unicodedata

TRAJETOS = ('IDA E VOLTA', 'SOMENTE IDA', 'SOMENTE VOLTA')


def normalizar(valor):
    return ' '.join(''.join(c for c in unicodedata.normalize('NFD', str(valor).casefold())
                            if not unicodedata.combining(c)).split())


def inferir_trajeto(observacoes):
    """Reconhece afirmações claras; dúvidas, terceiros e sinais conflitantes mantêm o padrão."""
    texto = normalizar(observacoes)
    # Não atribuir ao passageiro frases sobre familiares ou planos ainda incertos.
    ambiguo = r'\b(?:talvez|depende|provavelmente|possivelmente|se|caso|acho|pretendo|a confirmar|nao sei|nao tenho certeza)\b'
    terceiros = r'\b(?:meu|minha|meus|minhas|acompanhante|acompanhantes|ele|ela|eles|elas)\b'
    if re.search(ambiguo, texto) or re.search(terceiros, texto) or '?' in texto:
        return 'IDA E VOLTA'
    if re.search(r'\bnao (?:so|somente|apenas)\b', texto):
        return 'IDA E VOLTA'
    ida_verbo = r'(?:vou|vai|irei|ira|iremos|irao|ir)'
    volta_verbo = r'(?:volto|volta|voltarei|voltara|voltaremos|voltarao|voltar|retorno|retorna|retornarei|retornara|retornaremos|retornarao|retornar|regresso|regressarei|regressar)'
    # "Só vou retornar" refere-se à volta, e não à ida.
    texto = re.sub(r'\b(so|somente|apenas) (?:vou|vai|irei|ira) (?:voltar|retornar|regressar)\b', r'\1 volta', texto)
    exclusivo = r'(?:so|somente|apenas|unicamente|exclusivamente)'
    ida = bool(re.search(rf'\b{exclusivo} (?:de |na |a )?(?:ida|{ida_verbo})\b|\b(?:ida|{ida_verbo}) {exclusivo}\b', texto))
    volta = bool(re.search(rf'\b{exclusivo} (?:de |na |a )?(?:volta|retorno|{volta_verbo})\b|\b(?:volta|retorno|{volta_verbo}) {exclusivo}\b', texto))
    grupo = r'(?:com|na|no|pela|pelo) (?:a |o )?(?:caravana|grupo|onibus|excursao)'
    fim = r'(?=\s*(?:[.!;,]|e nao\b|$))'
    # Negação só é aceita isolada ou quando explicitamente ligada à caravana.
    nao_ida = bool(re.search(rf'\bnao (?:{ida_verbo}|vou ir)(?: {grupo}|\s*{fim})', texto))
    nao_volta = bool(re.search(rf'\bnao (?:{volta_verbo}|vou voltar|vou retornar|irei voltar|irei retornar)(?: {grupo}|\s*{fim})', texto))
    proprio = r'(?:por conta propria|por meios proprios|de (?:carro|uber|taxi|aviao|moto|trem)|(?:com|em|de|num|numa|no|na) (?:um |uma )?outr[oa] (?:onibus|caravana|grupo|excursao))'
    def afirmado(padrao):
        return any(not re.search(r'\bnao(?: \w+){0,3}\s*$', texto[max(0, m.start()-40):m.start()])
                   for m in re.finditer(padrao, texto))
    fora_ida = afirmado(rf'\b(?:{ida_verbo}|a ida sera|ida) {proprio}\b')
    fora_volta = afirmado(rf'\b(?:{volta_verbo}|a volta sera|o retorno sera|volta|retorno) {proprio}\b')
    ida = ida or nao_volta or fora_volta
    volta = volta or nao_ida or fora_ida
    # Afirmação explícita dos dois sentidos impede inferência parcial sem exclusão clara.
    if not (nao_ida or nao_volta or fora_ida or fora_volta) and re.search(r'\b(?:ida|vai|vou) e (?:volta|retorno|volto)\b', texto):
        return 'IDA E VOLTA'
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
