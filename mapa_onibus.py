"""Mapas dos leitos de 43 e 46 lugares, conforme referências da caravana."""
import pandas as pd
from html import escape

MAPAS = {
    43: {
        'Piso superior': [(1, 2, 3), (4, 5, 'Escada'), (6, 7, 'Escada'),
                          *[(i, i+1, i+2) for i in range(8, 30, 3)]],
        'Piso inferior': [(i, i+1, i+2) for i in range(32, 42, 3)],
    },
    46: {
        'Piso superior': [(1, 2, 3), (4, 5, 6), (7, 8, 'Escada'), (9, 10, 'Sem poltrona'),
                          *[(i, i+1, i+2) for i in range(11, 33, 3)]],
        'Piso inferior': [(i, i+1, i+2) for i in range(35, 45, 3)],
    },
}


def numero_poltrona(valor):
    texto = '' if pd.isna(valor) else str(valor).strip()
    return int(texto) if texto.isascii() and texto.isdecimal() else None


def nome_curto(nome):
    partes = str(nome).split()
    return ' '.join(partes if len(partes) <= 2 else [partes[0], partes[-1]])


def ocupacao(tabela):
    lugares = {}
    for i, valor in enumerate(tabela['Poltrona']):
        numero = numero_poltrona(valor)
        if numero is not None:
            lugares.setdefault(numero, []).append(i)
    return lugares


def atribuir_poltrona(tabela, passageiro, poltrona, modelo):
    if modelo not in MAPAS or not 0 <= passageiro < len(tabela):
        raise ValueError('Passageiro ou modelo inválido.')
    if poltrona is not None:
        if type(poltrona) is not int or not 1 <= poltrona <= modelo:
            raise ValueError('Poltrona fora do modelo escolhido.')
        outros = [i for i in ocupacao(tabela).get(poltrona, []) if i != passageiro]
        if outros:
            raise ValueError('A poltrona já está ocupada. Escolha um lugar livre.')
    resultado = tabela.copy(deep=True)
    resultado.iloc[passageiro, resultado.columns.get_loc('Poltrona')] = '' if poltrona is None else str(poltrona)
    return resultado


def mostrar_mapa(tabela, identidade):
    import streamlit as st

    def salvar(passageiro, lugar, modelo):
        try:
            st.session_state['mapa_tabela_' + identidade] = atribuir_poltrona(tabela, passageiro, lugar, modelo)
            chave = 'mapa_revisao_' + identidade
            st.session_state[chave] = st.session_state.get(chave, 0) + 1
            st.session_state.pop('mapa_erro_' + identidade, None)
        except ValueError as exc:
            st.session_state['mapa_erro_' + identidade] = str(exc)

    st.subheader('Escolha de poltronas')
    st.caption('Escolha o passageiro e clique em um lugar livre. Para mudar de lugar, clique na nova poltrona.')
    st.caption('Organize um ônibus por planilha. Trocar o modelo não altera as poltronas já preenchidas.')
    modelo = st.selectbox('Modelo do ônibus', [43, 46], format_func=lambda n: f'Leito · {n} lugares',
                          index=1 if any(44 <= n <= 46 for n in ocupacao(tabela)) else 0,
                          key='mapa_modelo_' + identidade)
    vagas = [i for i, v in enumerate(tabela['Poltrona']) if not str(v).strip()]
    chave_passageiro = 'mapa_passageiro_' + identidade
    if chave_passageiro not in st.session_state:
        st.session_state[chave_passageiro] = vagas[0] if vagas else 0
    passageiro = st.selectbox('Passageiro para acomodar', list(range(len(tabela))),
                             index=None,
                             format_func=lambda i: f'{i+1}. {tabela.iloc[i]["Nome Completo"]}',
                             key=chave_passageiro)
    if passageiro is None:
        st.info('Selecione um passageiro para visualizar o mapa.')
        return
    atual_passageiro = tabela.iloc[passageiro]
    st.caption(f'{atual_passageiro["Nome Completo"]} · Poltrona atual: {atual_passageiro["Poltrona"] or "a escolher"}')
    lugares = ocupacao(tabela)
    invalidos = [(i+1, str(v)) for i, v in enumerate(tabela['Poltrona'])
                 if str(v).strip() and numero_poltrona(v) not in range(1, modelo+1)]
    duplicados = [str(n) for n, indices in lugares.items() if len(indices) > 1 and 1 <= n <= modelo]
    if invalidos:
        st.warning('Há poltronas fora deste mapa ou com formato não reconhecido: '
                   + '; '.join(f'participante {i}: {v}' for i, v in invalidos)
                   + '. Os valores foram preservados; confira o modelo ou corrija na tabela.')
    if duplicados:
        st.warning('Poltronas com mais de um passageiro: ' + ', '.join(duplicados)
                   + '. Selecione um dos passageiros e mova-o para um lugar livre.')
    if st.session_state.get('mapa_erro_' + identidade):
        st.error(st.session_state['mapa_erro_' + identidade])
    ocupados = sum(1 for n in lugares if 1 <= n <= modelo)
    st.caption(f'{modelo - ocupados} livres · {ocupados} ocupadas · '
               f'{sum(not str(v).strip() for v in tabela["Poltrona"])} passageiros sem poltrona')
    st.caption('Número livre: disponível · × Ocupada · ✓ Poltrona atual do passageiro')
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
        pisos = st.tabs(['Piso superior', 'Piso inferior'])
        for (piso, filas), aba in zip(MAPAS[modelo].items(), pisos):
            with aba:
                st.caption('Janela · corredor à direita da dupla · janela')
                if modelo == 46 and piso == 'Piso inferior':
                    st.caption('Frente deste piso: toalete à esquerda · escada à direita')
                for fila in filas:
                    colunas = st.columns([1, 1, .3, 1])
                    with colunas[2]:
                        st.markdown('<div style="text-align:center;color:#c4d4d9;padding-top:12px">┆</div>', unsafe_allow_html=True)
                    for assento, coluna in zip(fila, (colunas[0], colunas[1], colunas[3])):
                        with coluna:
                            if isinstance(assento, str):
                                st.caption('↗ ' + assento if assento == 'Escada' else '—')
                                continue
                            indices = lugares.get(assento, [])
                            atual = indices == [passageiro]
                            legenda = 'Atual' if atual else ('Ocupada' if indices else 'Livre')
                            nomes = ', '.join(str(tabela.iloc[i]['Nome Completo']) for i in indices)
                            rotulo = f'{assento:02d}' + (' ✓' if atual else (' ×' if indices else ''))
                            st.button(rotulo,
                                      key=f'assento_{identidade}_{modelo}_{assento}', width='stretch',
                                      type='primary' if atual else 'secondary',
                                      disabled=bool(indices) and not atual,
                                      help=f'{legenda}: {nomes}' if nomes else f'Atribuir poltrona {assento}',
                                      on_click=salvar, args=(passageiro, assento, modelo))
                            if indices:
                                visiveis = '<br>'.join(escape(nome_curto(tabela.iloc[i]['Nome Completo'])) for i in indices)
                                st.markdown(f'<div class="ocupante">{visiveis}</div>', unsafe_allow_html=True)
        st.caption('Lado esquerdo: dupla · lado direito: individual. Mapas conforme os modelos enviados.')
    st.button('Liberar poltrona deste passageiro',
              disabled=not str(tabela.iloc[passageiro]['Poltrona']).strip(),
              key='liberar_' + identidade, on_click=salvar, args=(passageiro, None, modelo))
    st.caption('As escolhas ficam nesta sessão. Baixe a planilha com as correções para guardar as poltronas.')
