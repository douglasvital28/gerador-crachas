import logging
from hashlib import sha256
import streamlit as st
from geracao import BASE, MAX_PARTICIPANTES, CONVENIO, CONTATO, OPCOES, ler_planilha, gerar_pdf

st.set_page_config(page_title='Gerador de Crachás', page_icon='🪪')
st.title('🪪 Gerador de Crachás - Caravanas ao Solo Sagrado de Guarapiranga')
st.download_button('📄 Baixar Planilha Modelo', (BASE / 'modelo_planilha.xlsx').read_bytes(),
                   file_name='modelo_planilha.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                   on_click='ignore')
st.markdown(f'''
### 📝 Como utilizar este gerador
1. Baixe a planilha modelo e preencha os participantes, mantendo os nomes das colunas.
2. Envie o Excel e confira a prévia dos dados.
3. Preencha as informações da caravana e clique em **Gerar Crachás**.
4. Baixe o PDF e confira antes de imprimir.

Até **{MAX_PARTICIPANTES} participantes** e **5 MB** por planilha. Campos opcionais podem ficar vazios.
Preencha telefones como texto no Excel para preservar zeros à esquerda.
Preencha **Possuí Plano de Saúde?** com Sim/Não e **Tipo de Contato** com Familiar/Amigo.
Você também pode selecionar essas opções na tabela abaixo. Em branco significa sem marcação.
''')
df_file = st.file_uploader('📥 Envie a planilha preenchida (Excel)', type=['xlsx'], max_upload_size=5)
horario = st.text_input('Horário de Retorno:')
plataforma = st.text_input('Plataforma:')
responsavel = st.text_input('Responsável:')
ddd = st.text_input('DDD:')
telefone = st.text_input('Telefone do Responsável:')
conteudo = df_file.getvalue() if df_file else b''
df = None
if df_file:
    try:
        df = ler_planilha(conteudo)
        st.caption(f'{len(df)} participantes · {(len(df) + 3) // 4} páginas de crachás')
        st.caption('Selecione as opções de convênio e tipo de contato para cada participante.')
        df = st.data_editor(
            df, hide_index=True, key='participantes_' + sha256(conteudo).hexdigest(),
            disabled=[c for c in df.columns if c not in OPCOES],
            column_order=['Nome Completo', CONVENIO, CONTATO,
                          *[c for c in df.columns if c not in ('Nome Completo', CONVENIO, CONTATO)]],
            column_config={c: st.column_config.SelectboxColumn(
                'Possui convênio?' if c == CONVENIO else 'Tipo de contato',
                options=['', *opcoes], required=False,
            ) for c, opcoes in OPCOES.items()},
        )
    except ValueError as exc:
        st.error(str(exc))

assinatura = sha256(conteudo + repr((horario, plataforma, responsavel, ddd, telefone)).encode()
                   + (df.to_json().encode() if df is not None else b'')).hexdigest()
if st.session_state.get('assinatura') != assinatura:
    st.session_state.pop('pdf', None)
    st.session_state['assinatura'] = assinatura

if st.button('⚙️ Gerar Crachás', disabled=df is None):
    st.session_state.pop('pdf', None)
    try:
        with st.spinner('Gerando crachás…'):
            st.session_state['pdf'] = gerar_pdf(df, horario, plataforma, responsavel, ddd, telefone)
    except ValueError as exc:
        st.error(str(exc))
    except Exception:
        logging.exception('Falha ao gerar crachás')
        st.error('Não foi possível gerar os crachás. Tente novamente ou entre em contato com o responsável pelo aplicativo.')

if 'pdf' in st.session_state:
    st.success('✅ Crachás gerados com sucesso! Confira o PDF antes de imprimir.')
    st.download_button('📁 Baixar PDF com os Crachás', st.session_state['pdf'],
                       file_name='crachas_caravana.pdf', mime='application/pdf', on_click='ignore')
