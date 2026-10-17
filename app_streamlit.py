import logging
from hashlib import sha256
import streamlit as st
from geracao import BASE, MAX_PARTICIPANTES, CONVENIO, CONTATO, OPCOES, ler_planilha, gerar_pdf, avisos_duplicidades
from recursos import validar_caravana, exportar_planilha, renderizar_pagina
from edicao import validar_edicao
from padronizacao import maiusculas_participantes

st.set_page_config(page_title='Gerador de Crachás', page_icon='🪪')
st.title('🪪 Gerador de Crachás - Caravanas ao Solo Sagrado de Guarapiranga')
st.caption('Importe os participantes, confira os dados, escolha os assentos e prepare os documentos da viagem.')
with st.expander('Como usar e limites de importação'):
    st.markdown(f'''Até **{MAX_PARTICIPANTES} participantes** e **5 MB** por planilha.
Os dados dos participantes serão padronizados em **MAIÚSCULAS**.
Mantenha os telefones como texto no Excel para preservar zeros à esquerda.
Nos telefones do plano e do contato, informe o DDD junto ao número.
Convênio: SIM/NÃO. Tipo de contato: FAMILIAR/AMIGO. Campos opcionais podem ficar vazios.''')
st.subheader('1 · Importar participantes')
st.download_button('📄 Baixar Planilha Modelo', (BASE / 'modelo_planilha.xlsx').read_bytes(),
                   file_name='modelo_planilha.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                   on_click='ignore')
origem_dados = st.radio('Origem dos participantes', ['Planilha Excel', 'Google Sheets (teste)'],
                        horizontal=True, key='origem_dados')
if origem_dados == 'Google Sheets (teste)':
    from fonte_google import conectar_google
    df_file = conectar_google()
else:
    df_file = st.file_uploader('📥 Envie a planilha preenchida (Excel)', type=['xlsx'], max_upload_size=5)
conteudo = df_file.getvalue() if df_file else b''
identidade_planilha = sha256(conteudo).hexdigest()
df = None
selecionados = None
if df_file:
    try:
        df = maiusculas_participantes(ler_planilha(conteudo))
        st.subheader('2 · Conferir a listagem')
        st.caption('Corrija os dados diretamente na tabela. Desmarque Gerar crachá para quem não precisa de impressão.')
        coluna_selecao = 'Gerar crachá'
        while coluna_selecao in df.columns:
            coluna_selecao = '_' + coluna_selecao
        tabela = df.copy()
        tabela.insert(0, coluna_selecao, True)
        tabela = maiusculas_participantes(st.session_state.get('mapa_tabela_' + identidade_planilha, tabela))
        for campo in OPCOES:
            tabela[campo] = tabela[campo].fillna('').replace('', 'NÃO INFORMADO')
        revisao_mapa = st.session_state.get('mapa_revisao_' + identidade_planilha, 0)
        chave_editor = 'participantes_' + identidade_planilha + (f'_r{revisao_mapa}' if revisao_mapa else '')
        df = st.data_editor(
            tabela, hide_index=True, key=chave_editor,
            disabled=[c for c in df.columns if c not in ('Nome Completo', 'Unidade', 'Poltrona',
                      'Nome do Plano', 'Telefone do Plano', 'Nome Contato de Emergência',
                      'Telefone Contato de Emergência', CONVENIO, CONTATO)],
            column_order=[coluna_selecao, 'Nome Completo', CONVENIO, CONTATO,
                          *[c for c in df.columns if c not in ('Nome Completo', CONVENIO, CONTATO)]],
            column_config={coluna_selecao: st.column_config.CheckboxColumn('Gerar crachá', default=True),
                           '_ID_Inscricao': None,
                           **{c: st.column_config.SelectboxColumn(
                'Possui convênio?' if c == CONVENIO else 'Tipo de contato',
                options=['NÃO INFORMADO', *opcoes], required=False, format_func=lambda v: str(v or '').upper(),
            ) for c, opcoes in OPCOES.items()}},
        )
        for campo in OPCOES:
            df[campo] = df[campo].replace('NÃO INFORMADO', '')
        padronizada = maiusculas_participantes(df)
        if not df.equals(padronizada):
            st.session_state['mapa_tabela_' + identidade_planilha] = padronizada
            chave_revisao = 'mapa_revisao_' + identidade_planilha
            st.session_state[chave_revisao] = revisao_mapa + 1
            st.rerun()
        tabela_mapa = padronizada.copy(deep=True)
        if origem_dados == 'Google Sheets (teste)':
            from fonte_google import registrar_trabalho
            registrar_trabalho(tabela_mapa, identidade_planilha)
        mascara = df[coluna_selecao].fillna(False).astype(bool)
        df = validar_edicao(df.drop(columns=[coluna_selecao]))
        selecionados = df.loc[mascara].copy()
        duplicidades = avisos_duplicidades(df)
        if duplicidades:
            st.warning('Há nomes ou poltronas repetidos. Confira as linhas indicadas na planilha. '
                       'O aviso não impede a geração: nomes iguais podem ser de pessoas diferentes.')
            st.dataframe(duplicidades, hide_index=True)
        st.caption(f'{len(selecionados)} de {len(df)} participantes selecionados · {(len(selecionados) + 3) // 4} páginas de crachás')
        if selecionados.empty:
            st.info('Marque ao menos um participante na coluna Gerar crachá para gerar o PDF.')
        st.download_button('📊 Baixar planilha com minhas escolhas', exportar_planilha(df),
                           file_name='participantes_atualizados.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                           on_click='ignore')
        st.caption('Inclui todos os participantes com as correções e escolhas, inclusive os desmarcados. Os dados do ônibus são preenchidos na etapa 4 e não fazem parte desta planilha.')
    except ValueError as exc:
        df = None
        selecionados = None
        st.error(str(exc))

if df is None:
    st.session_state.pop('pdf', None)
    st.session_state.pop('lista_pdf', None)
    st.info('Importe uma planilha válida e confira os nomes para continuar às próximas etapas.')
    st.stop()

from mapa_onibus import mostrar_mapa
st.subheader('3 · Escolher assentos')
modelo_mapa = mostrar_mapa(tabela_mapa, identidade_planilha)

st.subheader('4 · Informar os dados do ônibus')
def maiuscula_campo(chave):
    st.session_state[chave] = st.session_state.get(chave, '').strip().upper()

coluna_viagem, coluna_responsavel = st.columns(2)
with coluna_viagem:
    horario = st.text_input('Horário de Retorno:', placeholder='18:00')
    plataforma = st.text_input('Plataforma:', key='plataforma', on_change=maiuscula_campo, args=('plataforma',))
with coluna_responsavel:
    responsavel = st.text_input('Responsável:', key='responsavel', on_change=maiuscula_campo, args=('responsavel',))
    ddd = st.text_input('DDD:', placeholder='11')
    telefone = st.text_input('Telefone do Responsável:', placeholder='91234-5678',
                             help='Você pode informar o número completo com DDD. Os campos são opcionais; preencha ambos ou deixe ambos vazios.')
erro_caravana = None
try:
    horario_validado, ddd_validado, telefone_validado = validar_caravana(horario, ddd, telefone)
except ValueError as exc:
    erro_caravana = str(exc)
    st.error(erro_caravana)
st.subheader('5 · Gerar documentos da viagem')
assinatura = sha256(conteudo + repr((horario, plataforma, responsavel, ddd, telefone)).encode()
                   + (df.to_json().encode() if df is not None else b'')
                   + (selecionados.to_json().encode() if selecionados is not None else b'')).hexdigest()
if st.session_state.get('assinatura') != assinatura:
    st.session_state.pop('pdf', None)
    st.session_state['assinatura'] = assinatura

aba_crachas, aba_lista, aba_mapa = st.tabs(['Crachás', 'Lista de embarque', 'Mapa do ônibus'])
with aba_mapa:
    from mapa_pdf import gerar_mapa_pdf
    st.subheader('Mapa preenchido para impressão')
    st.caption('A4, uma página por piso, com nomes completos em negrito. Inclui todos os passageiros e uma relação dos que ainda estão sem poltrona.')
    try:
        mapa_impresso = gerar_mapa_pdf(tabela_mapa, modelo_mapa or 43)
        st.download_button('Baixar PDF do mapa do ônibus', mapa_impresso,
                           file_name='mapa_onibus.pdf', mime='application/pdf', on_click='ignore')
    except ValueError as exc:
        st.warning(str(exc))
with aba_lista:
    from conferencia import gerar_lista
    st.subheader('Lista de conferência para impressão')
    st.caption('Lista em A4 com nome, unidade, poltrona e espaços para ida, retorno e observações.')
    abrangencia = st.selectbox('Participantes da lista', ['Todos os participantes', 'Apenas os selecionados'], key='abrangencia_lista')
    ordem_lista = st.selectbox('Ordenar lista por', ['Nome', 'Poltrona'], key='ordem_lista')
    data_viagem = st.date_input('Data da viagem (opcional)', value=None, format='DD/MM/YYYY')
    dados_lista = df if abrangencia == 'Todos os participantes' else selecionados
    assinatura_lista = sha256(dados_lista.to_json().encode()
                             + repr((ordem_lista, data_viagem, responsavel)).encode()).hexdigest()
    if st.session_state.get('assinatura_lista') != assinatura_lista:
        st.session_state.pop('lista_pdf', None)
        st.session_state['assinatura_lista'] = assinatura_lista
    st.caption(f'{len(dados_lista)} participantes na lista. Responsável: {responsavel.strip() or "preenchimento à mão"}.')
    if st.button('Gerar lista de conferência', key='gerar_lista', disabled=dados_lista.empty):
        st.session_state.pop('lista_pdf', None)
        try:
            st.session_state['lista_pdf'] = gerar_lista(dados_lista, ordem_lista, data_viagem, responsavel)
        except ValueError as exc:
            st.error(str(exc))
        except Exception:
            logging.exception('Falha ao gerar lista de conferência')
            st.error('Não foi possível gerar a lista. Tente novamente.')
    if 'lista_pdf' in st.session_state:
        st.download_button('📋 Baixar lista de conferência', st.session_state['lista_pdf'],
                           file_name='lista_conferencia.pdf', mime='application/pdf', on_click='ignore')
with aba_crachas:
    if st.button('⚙️ Gerar Crachás', key='gerar_crachas', disabled=selecionados is None or selecionados.empty or erro_caravana is not None):
        st.session_state.pop('pdf', None)
        try:
            with st.spinner('Gerando crachás…'):
                st.session_state['pdf'] = gerar_pdf(selecionados, horario_validado, plataforma, responsavel, ddd_validado, telefone_validado)
        except ValueError as exc:
            st.error(str(exc))
        except Exception:
            logging.exception('Falha ao gerar crachás')
            st.error('Não foi possível gerar os crachás. Tente novamente ou entre em contato com o responsável pelo aplicativo.')

    if 'pdf' in st.session_state:
        st.success('✅ Crachás gerados com sucesso! Confira o PDF antes de imprimir.')
        st.download_button('📁 Baixar PDF com os Crachás', st.session_state['pdf'],
                           file_name='crachas_caravana.pdf', mime='application/pdf', on_click='ignore')
        st.subheader('Prévia dos crachás')
        total_paginas = (len(selecionados) + 3) // 4
        pagina = st.number_input('Página da prévia', min_value=1, max_value=total_paginas,
                                 value=1, step=1, key='pagina_' + assinatura)
        try:
            st.image(renderizar_pagina(st.session_state['pdf'], pagina - 1), width='stretch',
                     caption=f'Página {pagina} de {total_paginas}', alt='Prévia da página de crachás gerada')
        except Exception:
            logging.exception('Falha ao renderizar prévia')
            st.warning('Não foi possível exibir a prévia. Baixe o PDF para conferir os crachás.')
