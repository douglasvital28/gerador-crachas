"""Interface de importação Google, isolada por conta e sessão."""
from hashlib import sha256
from io import BytesIO
import json
import streamlit as st
from geracao import normalizar
from edicao import validar_edicao
from recursos import exportar_planilha
from sheets_google import CAMPOS, CAMPO_ID, id_planilha, abas_planilha, respostas_planilha, mapear_respostas, atualizar_respostas


def conectar_google():
    try:
        auth = st.secrets.get('auth', {})
        configurado = all(auth.get(c) for c in ('redirect_uri', 'cookie_secret')) and bool(auth.get('google', {}).get('client_id'))
    except (FileNotFoundError, KeyError):
        configurado = False
    if not configurado:
        st.info('A conexão Google está em preparação. O administrador precisa configurar o projeto Google Cloud e os segredos do aplicativo.')
        return None
    if not st.user.get('is_logged_in', False):
        st.caption('Conecte sua conta para ler uma planilha privada. O aplicativo não altera sua planilha Google.')
        st.button('Conectar minha conta Google', key='gs_login', on_click=lambda: st.login('google'))
        return None
    dono = st.user.get('sub')
    if not dono:
        st.error('Não foi possível identificar a conta Google. Saia e conecte novamente.')
        return None
    if st.session_state.get('_gs_dono') not in (None, dono):
        # Nunca reutilizar dados da conta anterior na mesma sessão.
        for chave in list(st.session_state):
            if chave != 'origem_dados':
                del st.session_state[chave]
    st.session_state['_gs_dono'] = dono
    st.caption('Conta conectada: ' + str(st.user.get('email', 'Google')))
    def sair():
        st.session_state.clear()
        st.logout()
    st.button('Sair da conta Google', key='gs_logout', on_click=sair)
    token = st.user.tokens.get('access')
    if not token:
        st.error('A conta entrou, mas a autorização de leitura não está disponível. Confira expose_tokens e os escopos; depois saia e conecte novamente.')
        return None
    st.info('As correções ficam nesta sessão. Baixe o Excel antes de sair. O cadastro permanente de viagens ainda não está disponível.')
    link = st.text_input('Link da planilha de respostas', key='gs_link', placeholder='https://docs.google.com/spreadsheets/d/...')
    try:
        planilha = id_planilha(link) if link.strip() else None
    except ValueError as exc:
        st.error(str(exc))
        return None
    if st.button('Ler abas da planilha', key='gs_abas', disabled=not planilha):
        try:
            abas = abas_planilha(token, planilha)
            if not abas:
                raise ValueError('A planilha não contém abas disponíveis.')
            st.session_state['_gs_metadados'] = (planilha, abas)
        except ValueError as exc:
            st.error(str(exc))
            st.caption('Se houver uma importação anterior desta planilha, ela será mantida para você guardar o trabalho.')
    meta = st.session_state.get('_gs_metadados')
    if not meta or meta[0] != planilha:
        return None
    aba = st.selectbox('Aba com as respostas do Forms', meta[1], key='gs_aba_' + planilha)
    contexto = sha256((dono + '\0' + planilha + '\0' + aba).encode()).hexdigest()
    if st.button('Buscar respostas no Google', key='gs_buscar'):
        try:
            valores = respostas_planilha(token, planilha, aba)
            st.session_state['_gs_respostas'] = (contexto, valores)
        except ValueError as exc:
            st.error(str(exc))
            st.caption('As respostas anteriores desta aba foram mantidas; esta consulta não atualizou os participantes.')
    pacote = st.session_state.get('_gs_respostas')
    if not pacote or pacote[0] != contexto:
        return None
    valores = pacote[1]
    cabecalho = valores[0]
    hash_colunas = sha256(json.dumps(cabecalho, ensure_ascii=False).encode()).hexdigest()[:12]
    def rotulo(i):
        return 'Não importar este campo' if i is None else f'Coluna {i+1}: {cabecalho[i] or "sem título"}'
    st.caption('Associe cada campo à pergunta do seu Forms. Só Nome Completo é obrigatório; os outros podem ficar em branco.')
    mapeamento = {}
    for campo in CAMPOS:
        opcoes = [None, *range(len(cabecalho))]
        iguais = [i for i, titulo in enumerate(cabecalho) if normalizar(titulo) == normalizar(campo) or (campo == 'Observações' and normalizar(titulo) == 'observacao')]
        mapeamento[campo] = st.selectbox(campo, opcoes, index=opcoes.index(iguais[0]) if iguais else 0,
                                        format_func=rotulo, key=f'gs_mapa_{contexto}_{hash_colunas}_{campo}')
    st.caption('Para preservar correções na atualização, escolha uma coluna única e estável por inscrição. Pode ser o carimbo de data/hora, se não houver repetições. Não use nome ou poltrona.')
    candidatos = [i for i, v in enumerate(cabecalho) if normalizar(v) in ('carimbo de data/hora', 'timestamp')]
    coluna_id = st.selectbox('Identificador único da inscrição', list(range(len(cabecalho))),
                            index=candidatos[0] if candidatos else None, format_func=rotulo,
                            key=f'gs_id_{contexto}_{hash_colunas}')
    escolha = sha256(json.dumps([contexto, mapeamento, coluna_id], sort_keys=True).encode()).hexdigest()
    if st.button('Importar / atualizar participantes', key='gs_importar', disabled=coluna_id is None):
        try:
            namespace = contexto + ':' + str(coluna_id) + ':' + str(cabecalho[coluna_id])
            novas = validar_edicao(mapear_respostas(valores, mapeamento, coluna_id, namespace))
            origem = sha256(namespace.encode()).hexdigest()
            anterior = st.session_state.get('_gs_base')
            trabalho = st.session_state.get('_gs_trabalho')
            mesma_origem = st.session_state.get('_gs_origem') == origem
            coluna_selecao = 'Gerar crachá'
            adicionadas, removidas = len(novas), 0
            if mesma_origem and anterior is not None and trabalho is not None:
                tabela, adicionadas, removidas = atualizar_respostas(anterior, trabalho, novas)
                limpo = validar_edicao(tabela.drop(columns=[coluna_selecao], errors='ignore'))
            else:
                limpo = novas
                tabela = novas.copy()
                tabela.insert(0, coluna_selecao, True)
            dados = exportar_planilha(limpo)
            identidade = sha256(dados).hexdigest()
            st.session_state['mapa_tabela_' + identidade] = tabela.copy(deep=True)
            chave_revisao = 'mapa_revisao_' + identidade
            st.session_state[chave_revisao] = st.session_state.get(chave_revisao, 0) + 1
            st.session_state['_gs_bytes'] = dados
            st.session_state['_gs_identidade'] = identidade
            st.session_state['_gs_base'] = novas.copy(deep=True)
            st.session_state['_gs_trabalho'] = tabela.copy(deep=True)
            st.session_state['_gs_origem'] = origem
            st.session_state['_gs_escolha'] = escolha
            st.success(f'{len(tabela)} participantes prontos. {adicionadas} novas inscrições adicionadas nesta importação.')
            if removidas:
                st.warning(f'{removidas} inscrições não estão mais na origem e foram mantidas aqui para conferência. Nenhum participante foi removido automaticamente.')
        except ValueError as exc:
            st.error(str(exc))
    if st.session_state.get('_gs_escolha') != escolha:
        return None
    dados = st.session_state.get('_gs_bytes')
    return BytesIO(dados) if dados else None


def registrar_trabalho(tabela, identidade):
    if identidade == st.session_state.get('_gs_identidade'):
        st.session_state['_gs_trabalho'] = tabela.copy(deep=True)
