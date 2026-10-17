import logging
from hashlib import sha256
import streamlit as st
from geracao import BASE, MAX_PARTICIPANTES, CONVENIO, CONTATO, OPCOES, ler_planilha, gerar_pdf, avisos_duplicidades
from recursos import validar_caravana, exportar_planilha, renderizar_pagina
from edicao import validar_edicao

st.set_page_config(page_title='Gerador de Crachás', page_icon='🪪')
st.title('🪪 Gerador de Crachás - Caravanas ao Solo Sagrado de Guarapiranga')
st.download_button('📄 Baixar Planilha Modelo', (BASE / 'modelo_planilha.xlsx').read_bytes(),
                   file_name='modelo_planilha.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                   on_click='ignore')
st.markdown(f'''
### 📝 Como utilizar este gerador
1. Baixe a planilha modelo e preencha os participantes, mantendo os nomes das colunas.
2. Envie o Excel, corrija os dados na tabela e marque os participantes que receberão crachá.
3. Preencha as informações da caravana e clique em **Gerar Crachás**.
4. Confira a prévia do PDF e baixe para imprimir.
5. Baixe a planilha com suas escolhas para reutilizar em outra geração.

Até **{MAX_PARTICIPANTES} participantes** e **5 MB** por planilha. Campos opcionais podem ficar vazios.
Preencha telefones como texto no Excel para preservar zeros à esquerda.
Nos telefones do plano e do contato, informe o DDD junto ao número: ele será separado no crachá.
Números locais, serviços como 0800 e formatos não reconhecidos serão mantidos como informados.
Preencha **Possuí Plano de Saúde?** com Sim/Não e **Tipo de Contato** com Familiar/Amigo.
Você também pode selecionar essas opções na tabela abaixo. Em branco significa sem marcação.
''')
df_file = st.file_uploader('📥 Envie a planilha preenchida (Excel)', type=['xlsx'], max_upload_size=5)
horario = st.text_input('Horário de Retorno:', placeholder='18:00')
plataforma = st.text_input('Plataforma:')
responsavel = st.text_input('Responsável:')
ddd = st.text_input('DDD:', placeholder='11')
telefone = st.text_input('Telefone do Responsável:', placeholder='91234-5678',
                         help='Você pode informar o número completo com DDD. Os campos são opcionais; preencha ambos ou deixe ambos vazios.')
erro_caravana = None
try:
    horario_validado, ddd_validado, telefone_validado = validar_caravana(horario, ddd, telefone)
except ValueError as exc:
    erro_caravana = str(exc)
    st.error(erro_caravana)
conteudo = df_file.getvalue() if df_file else b''
identidade_planilha = sha256(conteudo).hexdigest()
df = None
selecionados = None
if df_file:
    try:
        df = ler_planilha(conteudo)
        st.caption('Corrija os dados diretamente na tabela. Desmarque Gerar crachá para quem não precisa de impressão.')
        coluna_selecao = 'Gerar crachá'
        while coluna_selecao in df.columns:
            coluna_selecao = '_' + coluna_selecao
        tabela = df.copy()
        tabela.insert(0, coluna_selecao, True)
        tabela = st.session_state.get('mapa_tabela_' + identidade_planilha, tabela).copy(deep=True)
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
                           **{c: st.column_config.SelectboxColumn(
                'Possui convênio?' if c == CONVENIO else 'Tipo de contato',
                options=['', *opcoes], required=False,
            ) for c, opcoes in OPCOES.items()}},
        )
        tabela_mapa = df.copy(deep=True)
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
        st.caption('Inclui todos os participantes com as correções e escolhas, inclusive os desmarcados. Os dados da caravana acima não fazem parte desta planilha.')
    except ValueError as exc:
        df = None
        selecionados = None
        st.error(str(exc))

assinatura = sha256(conteudo + repr((horario, plataforma, responsavel, ddd, telefone)).encode()
                   + (df.to_json().encode() if df is not None else b'')
                   + (selecionados.to_json().encode() if selecionados is not None else b'')).hexdigest()
if st.session_state.get('assinatura') != assinatura:
    st.session_state.pop('pdf', None)
    st.session_state['assinatura'] = assinatura

if st.button('⚙️ Gerar Crachás', disabled=selecionados is None or selecionados.empty or erro_caravana is not None):
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

# A lista pode ser impressa independentemente da geração de crachás.
if df is not None:
    from conferencia import gerar_lista
    st.subheader('Lista de conferência para impressão')
    st.caption('Lista em A4 com nome, unidade, poltrona e espaços para ida, retorno e observações.')
    abrangencia = st.selectbox('Participantes da lista', ['Todos os participantes', 'Apenas os selecionados'])
    ordem_lista = st.selectbox('Ordenar lista por', ['Nome', 'Poltrona'])
    data_viagem = st.date_input('Data da viagem (opcional)', value=None, format='DD/MM/YYYY')
    dados_lista = df if abrangencia == 'Todos os participantes' else selecionados
    assinatura_lista = sha256(dados_lista.to_json().encode()
                             + repr((ordem_lista, data_viagem, responsavel)).encode()).hexdigest()
    if st.session_state.get('assinatura_lista') != assinatura_lista:
        st.session_state.pop('lista_pdf', None)
        st.session_state['assinatura_lista'] = assinatura_lista
    st.caption(f'{len(dados_lista)} participantes na lista. Responsável: {responsavel.strip() or "preenchimento à mão"}.')
    if st.button('Gerar lista de conferência', disabled=dados_lista.empty):
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
else:
    st.session_state.pop('lista_pdf', None)
    st.session_state.pop('assinatura_lista', None)

if df is not None:
    from mapa_onibus import mostrar_mapa
    mostrar_mapa(tabela_mapa, identidade_planilha)
