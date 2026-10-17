"""Leitura privada do Google Sheets e atualização sem sobrescrever correções locais."""
from hashlib import sha256
from urllib.parse import urlparse, quote
import re
from trajetos import preparar_trajetos
import pandas as pd
import requests
from geracao import COLUNAS, CONVENIO, CONTATO, MAX_PARTICIPANTES, aplicar_missao

CAMPO_ID = '_ID_Inscricao'
CAMPOS = (*COLUNAS, CONVENIO, CONTATO, 'Missão', 'Observações', 'Trajeto')


def id_planilha(link):
    texto = link.strip()
    if re.fullmatch(r'[A-Za-z0-9_-]{20,200}', texto):
        return texto
    url = urlparse(texto)
    if url.scheme != 'https' or url.hostname != 'docs.google.com' or url.username or url.password:
        raise ValueError('Informe o link https://docs.google.com/spreadsheets/d/... da planilha de respostas.')
    match = re.match(r'^/spreadsheets/d/([A-Za-z0-9_-]{20,200})(?:/|$)', url.path)
    if not match:
        raise ValueError('O link precisa ser de uma planilha Google Sheets, não do formulário.')
    return match.group(1)


def ler_api(token, caminho, params):
    # O token só é enviado para a API fixa do Google. Não seguir redirecionamentos.
    try:
        resposta = requests.get('https://sheets.googleapis.com/v4/spreadsheets/' + caminho,
                                params=params, headers={'Authorization': 'Bearer ' + token},
                                timeout=(5, 20), allow_redirects=False, stream=True)
        with resposta:
            if resposta.status_code == 401:
                raise ValueError('A autorização Google expirou. Saia e conecte sua conta novamente.')
            if resposta.status_code in (403, 404):
                raise ValueError('Não foi possível acessar a planilha. Confira a conta, o link, a permissão de leitura e se a API Google Sheets foi ativada.')
            if resposta.status_code == 429:
                raise ValueError('O Google limitou as consultas. Aguarde um pouco antes de tentar novamente.')
            if resposta.status_code != 200:
                raise ValueError('O Google não conseguiu atender à consulta. Tente novamente mais tarde.')
            partes, tamanho = [], 0
            for parte in resposta.iter_content(16384):
                tamanho += len(parte)
                if tamanho > 5 * 1024 * 1024:
                    raise ValueError('A resposta da planilha excede 5 MB. Reduza os dados para este teste.')
                partes.append(parte)
            import json
            return json.loads(b''.join(partes))
    except (requests.RequestException, UnicodeError, ValueError) as exc:
        if isinstance(exc, ValueError) and not isinstance(exc, requests.exceptions.JSONDecodeError):
            if str(exc).startswith(('A autorização', 'Não foi', 'O Google', 'A resposta')):
                raise
        raise ValueError('Não foi possível ler a resposta do Google. Confira sua conexão e tente novamente.') from None


def abas_planilha(token, planilha):
    payload = ler_api(token, id_planilha(planilha), {'fields': 'properties(title),sheets(properties(title,sheetId))'})
    return [s['properties']['title'] for s in payload.get('sheets', [])]


def respostas_planilha(token, planilha, aba):
    intervalo = "'" + aba.replace("'", "''") + "'!A1:CV402"
    payload = ler_api(token, id_planilha(planilha) + '/values/' + quote(intervalo, safe=''),
                      {'valueRenderOption': 'FORMATTED_VALUE', 'majorDimension': 'ROWS'})
    valores = payload.get('values', [])
    if len(valores) < 2:
        raise ValueError('A aba precisa ter os títulos na primeira linha e ao menos uma resposta.')
    if len(valores) - 1 > MAX_PARTICIPANTES:
        raise ValueError('A aba excede 400 respostas. Use uma aba por viagem com até 400 participantes.')
    return valores


def mapear_respostas(valores, mapeamento, coluna_id, origem):
    cabecalho, *linhas = valores
    if mapeamento.get('Nome Completo') is None:
        raise ValueError('Associe a pergunta que contém Nome Completo.')
    indices = [i for i in mapeamento.values() if i is not None]
    if any(type(i) is not int or not 0 <= i < len(cabecalho) for i in [*indices, coluna_id]):
        raise ValueError('Escolha colunas válidas para o mapeamento e identificador.')
    if len(indices) != len(set(indices)):
        raise ValueError('Cada pergunta da planilha pode preencher apenas um campo do participante.')
    registros, ids = [], set()
    for numero, linha in enumerate(linhas, 2):
        if not any(str(v).strip() for v in linha):
            continue
        def valor(indice):
            return str(linha[indice]).strip() if indice is not None and indice < len(linha) else ''
        identificador = valor(coluna_id)
        if not identificador:
            raise ValueError(f'O identificador da inscrição está vazio na linha {numero}.')
        chave = sha256((origem + '\0' + identificador).encode()).hexdigest()
        if chave in ids:
            raise ValueError('A coluna de identificação tem respostas repetidas. Escolha uma coluna única e estável por inscrição.')
        ids.add(chave)
        registro = {campo: valor(mapeamento.get(campo)) for campo in CAMPOS}
        if not registro['Nome Completo']:
            raise ValueError(f'O nome do participante está vazio na linha {numero}.')
        registro[CAMPO_ID] = chave
        registros.append(registro)
    if not registros:
        raise ValueError('A aba não contém participantes preenchidos.')
    return preparar_trajetos(aplicar_missao(pd.DataFrame(registros)))


def atualizar_respostas(anterior, trabalho, novas, coluna_selecao='Gerar crachá'):
    """Mescla por ID: atualiza origem intacta, preserva correções e linhas removidas."""
    for frame in (anterior, trabalho, novas):
        if CAMPO_ID not in frame or frame[CAMPO_ID].duplicated().any():
            raise ValueError('Identificadores ausentes ou repetidos. A atualização foi interrompida para preservar as correções.')
    antigas = anterior.set_index(CAMPO_ID).to_dict('index')
    recentes = novas.set_index(CAMPO_ID).to_dict('index')
    resultado = trabalho.copy(deep=True)
    for i in range(len(resultado)):
        chave = resultado.iloc[i][CAMPO_ID]
        if chave not in antigas or chave not in recentes:
            continue
        for campo in CAMPOS:
            if campo not in resultado:
                resultado[campo] = ''
            if str(resultado.iloc[i][campo]).strip().upper() == str(antigas[chave].get(campo, '')).strip().upper():
                resultado.iat[i, resultado.columns.get_loc(campo)] = recentes[chave].get(campo, '')
    existentes = set(resultado[CAMPO_ID])
    adicionar = novas.loc[~novas[CAMPO_ID].isin(existentes)].copy()
    if coluna_selecao in resultado:
        adicionar[coluna_selecao] = True
    resultado = pd.concat([resultado, adicionar], ignore_index=True).fillna('')
    if len(resultado) > MAX_PARTICIPANTES:
        raise ValueError('O total após a atualização excede 400 participantes. Nenhuma alteração foi aplicada.')
    removidos = sum(chave not in recentes for chave in existentes)
    return resultado, len(adicionar), removidos
