from colo import eh_colo, acompanhantes
"""Mapas dos leitos 43/46, misto 56 e LD semi-leito 44, conforme referências da caravana."""
import pandas as pd
from html import escape
from trajetos import sentidos, conflitam

MAPAS = {
    43: {
        'Piso superior': [(1, 2, 3), (4, 5, 'Escada'), (6, 7, 'Escada'),
                          *[(i, i+1, i+2) for i in range(8, 30, 3)]],
        'Piso inferior': [(i, i+1, i+2) for i in range(32, 42, 3)],
    },
    44: {
        'Semi-leito': [(1, 2, 4, 3), (5, 6, 'Escada', ''),
                       *[(i, i+1, i-1, i-2) for i in range(9, 38, 4)],
                       (41, 42, 39, 40), (43, 44, 'Banheiro', '')],
    },
    46: {
        'Piso superior': [(1, 2, 3), (4, 5, 6), (7, 8, 'Escada'), (9, 10, 'Sem poltrona'),
                          *[(i, i+1, i+2) for i in range(11, 33, 3)]],
        'Piso inferior': [(i, i+1, i+2) for i in range(35, 45, 3)],
    },
    56: {
        # Frente no topo: janela esquerda, corredor e janela direita.
        'Piso superior': [(1, 2, 4, 3), (5, 6, 8, 7),
                          (9, 10, 'Escada', ''), (11, 12, 'Escada', ''),
                          *[(i, i+1, i+3, i+2) for i in range(13, 42, 4)]],
        'Piso inferior': [(i, i+1, i+2) for i in range(45, 55, 3)],
    },
}


def numero_poltrona(valor):
    texto = '' if pd.isna(valor) else str(valor).strip()
    return int(texto) if texto.isascii() and texto.isdecimal() else None


def nome_curto(nome):
    partes = str(nome).split()
    titulos = {'min', 'ministra', 'ministro', 'prof', 'professora'}
    prefixos = []
    while partes and partes[0].casefold().rstrip('.') in titulos:
        prefixos.append(partes.pop(0))
    abreviado = partes if len(partes) <= 2 else [partes[0], partes[-1]]
    return ' '.join(prefixos + abreviado)


def ocupacao(tabela):
    lugares = {}
    for i, valor in enumerate(tabela['Poltrona']):
        if eh_colo(tabela.iloc[i]):continue
        numero = numero_poltrona(valor)
        if numero is not None:
            lugares.setdefault(numero, []).append(i)
    return lugares


def atribuir_poltrona(tabela, passageiro, poltrona, modelo):
    if modelo not in MAPAS or not 0 <= passageiro < len(tabela):
        raise ValueError('Passageiro ou modelo inválido.')
    if eh_colo(tabela.iloc[passageiro]):
        raise ValueError('Criança de colo não recebe poltrona própria. Escolha o responsável.')
    if poltrona is not None:
        if type(poltrona) is not int or not 1 <= poltrona <= modelo:
            raise ValueError('Poltrona fora do modelo escolhido.')
        outros = [i for i in ocupacao(tabela).get(poltrona, []) if i != passageiro and sentidos(tabela.iloc[i]) & sentidos(tabela.iloc[passageiro])]
        if outros:
            raise ValueError('A poltrona já está ocupada. Escolha um lugar livre.')
    resultado = tabela.copy(deep=True)
    resultado.iloc[passageiro, resultado.columns.get_loc('Poltrona')] = '' if poltrona is None else str(poltrona)
    return resultado


def mostrar_mapa(tabela, identidade, titulo=None):
    import streamlit as st

    def salvar(passageiro, lugar, modelo):
        try:
            st.session_state['mapa_tabela_' + identidade] = atribuir_poltrona(tabela, passageiro, lugar, modelo)
            chave = 'mapa_revisao_' + identidade
            st.session_state[chave] = st.session_state.get(chave, 0) + 1
            st.session_state.pop('mapa_erro_' + identidade, None)
        except ValueError as exc:
            st.session_state['mapa_erro_' + identidade] = str(exc)

    if titulo:
        st.subheader(titulo)
    st.caption('Escolha o passageiro e clique em um lugar livre. Para mudar de lugar, clique na nova poltrona.')
    st.caption('Organize um ônibus por planilha. Trocar o modelo não altera as poltronas já preenchidas.')
    modelos = [43, 46, 56, 44]
    modelo = st.selectbox('Modelo do ônibus', modelos,
                          format_func=lambda n: 'LD Semi-leito · 44 lugares' if n == 44 else 'Misto · 56 lugares (semi-leito + leito)' if n == 56 else f'Leito · {n} lugares',
                          index=2 if any(47 <= n <= 56 for n in ocupacao(tabela)) else (1 if any(44 <= n <= 46 for n in ocupacao(tabela)) else 0),
                          key='mapa_modelo_' + identidade)
    chave_revisar = 'mapa_importado_completo_' + identidade
    if chave_revisar not in st.session_state:
        st.session_state[chave_revisar] = all(str(v).strip() for v in tabela['Poltrona'])
    if st.session_state[chave_revisar] and all(str(v).strip() for v in tabela['Poltrona']):
        revisar = st.radio('Deseja revisar os assentos importados?', ['Não', 'Sim'],
                           horizontal=True, key='revisar_assentos_' + identidade)
        if revisar == 'Não':
            st.caption('Assentos importados mantidos. O PDF do mapa continua disponível na etapa 5.')
            return modelo
    vagas = [i for i, v in enumerate(tabela['Poltrona']) if not str(v).strip() and not eh_colo(tabela.iloc[i])]
    st.caption('🟢 Com poltrona marcada · 🟠 Sem poltrona — precisa escolher um assento')
    chave_passageiro = 'mapa_passageiro_' + identidade
    if chave_passageiro not in st.session_state:
        st.session_state[chave_passageiro] = vagas[0] if vagas else next((i for i in range(len(tabela)) if not eh_colo(tabela.iloc[i])), None)
    disponiveis = [i for i in range(len(tabela)) if not eh_colo(tabela.iloc[i])]
    if chave_passageiro in st.session_state and st.session_state[chave_passageiro] not in [None,*disponiveis]:
        del st.session_state[chave_passageiro]
    passageiro = st.selectbox('Passageiro para acomodar', disponiveis,
                             index=None,
                             format_func=lambda i: ('🟢 ' if str(tabela.iloc[i]['Poltrona']).strip() else '🟠 ') + f'{i+1}. {tabela.iloc[i]["Nome Completo"]} · {tabela.iloc[i].get("Trajeto", "IDA E VOLTA")} · ' + (f'✓ Poltrona {tabela.iloc[i]["Poltrona"]}' if str(tabela.iloc[i]['Poltrona']).strip() else '○ SEM POLTRONA'),
                             key=chave_passageiro)
    if passageiro is None:
        st.info('Selecione um passageiro para atribuir ou liberar uma poltrona. Você pode consultar o mapa abaixo.')
    st.button('Liberar poltrona deste passageiro',
              disabled=passageiro is None or not str(tabela.iloc[passageiro]['Poltrona']).strip(),
              key='liberar_' + identidade, on_click=salvar, args=(passageiro, None, modelo))
    if passageiro is not None:
        atual_passageiro = tabela.iloc[passageiro]
        st.caption(f'{atual_passageiro["Nome Completo"]} · Poltrona atual: {atual_passageiro["Poltrona"] or "a escolher"}')
    for i in range(len(tabela)):
        if eh_colo(tabela.iloc[i]):
            st.caption(str(tabela.iloc[i]['Nome Completo'])+' · COLO · Responsável: '+str(tabela.iloc[i].get('Responsável acompanhante','não informado')))
    lugares = ocupacao(tabela)
    invalidos = [(i+1, str(v)) for i, v in enumerate(tabela['Poltrona'])
                 if str(v).strip() and not eh_colo(tabela.iloc[i]) and numero_poltrona(v) not in range(1, modelo+1)]
    duplicados = [str(n) for n, indices in lugares.items() if conflitam(tabela, indices) and 1 <= n <= modelo]
    if invalidos:
        st.warning('Há poltronas fora deste mapa ou com formato não reconhecido: '
                   + '; '.join(f'participante {i}: {v}' for i, v in invalidos)
                   + '. Os valores foram preservados; confira o modelo ou corrija na tabela.')
    if duplicados:
        st.warning('Poltronas com conflito no mesmo sentido: ' + ', '.join(duplicados)
                   + '. Selecione um dos passageiros e mova-o para um lugar livre.')
    if st.session_state.get('mapa_erro_' + identidade):
        st.error(st.session_state['mapa_erro_' + identidade])
    ocupados = {sentido: sum(any(sentido in sentidos(tabela.iloc[i]) for i in indices)
                             for n, indices in lugares.items() if 1 <= n <= modelo)
                for sentido in ('IDA', 'VOLTA')}
    st.caption(f'{modelo - ocupados["IDA"]} livres na ida · {modelo - ocupados["VOLTA"]} livres na volta · '
               f'{sum(not str(v).strip() for v in tabela["Poltrona"])} passageiros sem poltrona')
    st.caption('Uma poltrona pode ser compartilhada por passageiros de sentidos diferentes. Confira o trajeto na listagem. Uma ocupada no outro sentido pode ser selecionada.')
    st.caption('Número livre: disponível · × Ocupada neste sentido · ↔ Disponível no seu sentido · ✓ Poltrona atual')
    with st.container(key='mapa_onibus', border=True):
        st.markdown('''<style>
.st-key-mapa_onibus {background:linear-gradient(180deg,#eff8f9 0%,#ffffff 70%);border-radius:32px 32px 18px 18px;max-width:560px;margin-inline:auto;}
.st-key-mapa_onibus [data-testid="stHorizontalBlock"] {flex-wrap:nowrap!important;gap:.4rem;}
.st-key-mapa_onibus [data-testid="stColumn"] {min-width:0!important;width:auto!important;flex:1 1 0!important;}
.st-key-mapa_onibus [data-testid="stColumn"]:nth-child(3) {flex:.3 1 0!important;}
.st-key-mapa_onibus button[kind="secondary"] {border:1.5px solid #399b9a;background:#f5fffc;border-radius:12px;min-height:54px;font-weight:600;color:#19605d;}
.st-key-mapa_onibus button[kind="primary"] {background:#166b70;border-color:#166b70;border-radius:12px;min-height:54px;color:white;}
.st-key-mapa_onibus button:disabled {background:#edf0f3!important;border-color:#ccd3db!important;color:#627084!important;opacity:1;}
.st-key-mapa_onibus button p {font-size:13px!important;}
.st-key-mapa_onibus .ocupante {text-align:center;font-size:12px;line-height:1.35;color:#304a54;overflow-wrap:anywhere;padding:0 2px 6px;}
</style>''', unsafe_allow_html=True)
        st.markdown('''<div style="text-align:center;padding:14px 0 20px;color:#245a65;letter-spacing:2px;font-size:11px;font-weight:700">
FRENTE DO ÔNIBUS<div style="margin:12px auto 0;width:70%;height:18px;border-radius:30px 30px 7px 7px;background:linear-gradient(100deg,#b9e1e7,#e5f2f5,#b9e1e7)"></div></div>''', unsafe_allow_html=True)
        pisos = st.tabs(['Piso superior · Semi-leito', 'Piso inferior · Leito'] if modelo == 56 else ['Semi-leito · LD 44'] if modelo == 44 else ['Piso superior', 'Piso inferior'])
        for (piso, filas), aba in zip(MAPAS[modelo].items(), pisos):
            with aba:
                duplas = modelo == 44 or (modelo == 56 and piso == 'Piso superior')
                st.caption('Semi-leito · 2 + 2 · duplas dos dois lados do corredor' if duplas else 'Leito · 2 + 1 · dupla à esquerda e individual à direita')
                if modelo in (46, 56) and piso == 'Piso inferior':
                    st.caption('Frente deste piso: toalete à esquerda · escada à direita')
                for fila in filas:
                    colunas = st.columns([1, 1, .3, 1, 1] if duplas else [1, 1, .3, 1])
                    with colunas[2]:
                        st.markdown('<div style="text-align:center;color:#c4d4d9;padding-top:12px">┆</div>', unsafe_allow_html=True)
                    for assento, coluna in zip(fila, [*colunas[:2], *colunas[3:]]):
                        with coluna:
                            if isinstance(assento, str):
                                if assento:
                                    st.caption('↗ Escada' if assento == 'Escada' else 'WC Banheiro' if assento == 'Banheiro' else '—')
                                continue
                            indices = lugares.get(assento, [])
                            atual = passageiro is not None and passageiro in indices
                            conflito = passageiro is not None and any(j != passageiro and sentidos(tabela.iloc[j]) & sentidos(tabela.iloc[passageiro]) for j in indices)
                            legenda = 'Atual' if atual else ('Disponível no seu sentido' if indices and passageiro is not None and not conflito else 'Ocupada' if indices else 'Livre')
                            nomes = ', '.join(str(tabela.iloc[i]['Nome Completo']) for i in indices)
                            rotulo = f'{assento:02d}' + (' ✓' if atual else (' ↔' if indices and passageiro is not None and not conflito else (' ×' if indices else '')))
                            st.button(rotulo,
                                      key=f'assento_{identidade}_{modelo}_{assento}', width='stretch',
                                      type='primary' if atual else 'secondary',
                                      disabled=passageiro is None or conflito,
                                      help=f'{legenda}: {nomes}' if nomes else ('Selecione um passageiro para atribuir esta poltrona' if passageiro is None else f'Atribuir poltrona {assento}'),
                                      on_click=salvar, args=(passageiro, assento, modelo))
                            if indices:
                                visiveis = '<br>'.join(escape(nome_curto(tabela.iloc[i]['Nome Completo'])) + ''.join('<br><small>COLO: '+escape(nome_curto(tabela.iloc[j]['Nome Completo']))+'</small>' for j in acompanhantes(tabela,i)) + ('<br><small>' + escape(tabela.iloc[i].get('Trajeto', 'IDA E VOLTA')) + '</small>' if tabela.iloc[i].get('Trajeto', 'IDA E VOLTA') != 'IDA E VOLTA' else '') for i in indices)
                                st.markdown(f'<div class="ocupante">{visiveis}</div>', unsafe_allow_html=True)
        st.caption('Superior: semi-leito 1–44 · Inferior: leito 45–56. Numeração conforme o G8 Misto enviado.' if modelo == 56 else 'LD Semi-leito · 44 lugares. Escada e banheiro conforme a referência enviada.' if modelo == 44 else 'Lado esquerdo: dupla · lado direito: individual. Mapas conforme os modelos enviados.')
    st.caption('As escolhas ficam nesta sessão. Baixe a planilha com as correções para guardar as poltronas.')

    return modelo
