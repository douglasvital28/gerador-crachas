import logging
from hashlib import sha256
import streamlit as st
from geracao import BASE, MAX_PARTICIPANTES, CONVENIO, CONTATO, OPCOES, ler_planilha, gerar_pdf, avisos_duplicidades
from recursos import validar_caravana, exportar_planilha, renderizar_pagina
from edicao import validar_edicao
from trajetos import TRAJETOS, preparar_trajetos
from padronizacao import maiusculas_participantes, ordenar_por_missao

st.set_page_config(page_title='Gerador de Crachás', page_icon='🪪')
from paginas_publicas import mostrar_pagina_publica, links_publicos
if mostrar_pagina_publica():
    st.stop()

st.title('🪪 Gerador de Crachás - Caravanas ao Solo Sagrado de Guarapiranga')
st.caption('Importe os participantes, confira os dados, escolha os assentos e prepare os documentos da viagem.')
links_publicos()
with st.expander('Como usar o sistema · passo a passo'):
    st.markdown(f'''### 1 · Importe os participantes
**Excel:** baixe a **Planilha Modelo**, preencha e envie o arquivo `.xlsx`.
O limite é de **{MAX_PARTICIPANTES} participantes** e **5 MB** por arquivo.
A coluna opcional **Missão** acrescenta **MIN.** para Ministro(a) e **PROF.** para Professor(a), sem repetir um prefixo já presente no nome. No Google Sheets, associe também esse campo à pergunta correspondente.
A coluna de poltrona pode ficar vazia: os assentos podem ser escolhidos no sistema.
Mantenha os telefones como texto no Excel para preservar zeros à esquerda.
Nos telefones do plano e do contato de emergência, informe o DDD junto ao número.
**Telefone do plano iniciado por 0800:** informe o número completo, sem acrescentar DDD; o campo DDD ficará vazio no crachá.

**Google Sheets:** conecte sua conta Google e cole o link da **planilha de respostas** vinculada ao Forms.
Clique em **Ler abas da planilha**, selecione a aba e use **Buscar respostas no Google**.
Associe as perguntas aos campos do sistema e escolha um **identificador único da inscrição**:
use uma coluna estável e sem repetições, como o carimbo de data/hora, quando for único.
Não use nome ou poltrona como identificador. Depois clique em **Importar / atualizar participantes**.
Cada organizador usa sua própria conta, que precisa ter acesso à planilha.
Se o Google bloquear a entrada, entre em contato com o administrador para conferir as permissões da integração.
O aplicativo apenas lê a planilha; as alterações feitas aqui não são enviadas ao Google.
Para buscar novas inscrições, repita a busca e a atualização: as correções e poltronas da sessão são preservadas.
Inscrições removidas na origem permanecem disponíveis para conferência.

### 2 · Confira e corrija a listagem
A listagem é organizada por missão: Ministros, Professores, Coordenadores, Assistentes de Ministros, Responsáveis de Setor, Assistentes de Família, Membros, Frequentadores e Primeira Vez, em ordem alfabética dentro de cada categoria. Missões não reconhecidas ou vazias ficam ao final; prefixos MIN. e PROF. identificam essas categorias quando a missão não é reconhecida.
Revise os nomes, unidades, telefones, convênios, contatos e poltronas diretamente na tabela.
Os dados dos participantes são padronizados em **MAIÚSCULAS**, inclusive após correções.
Convênio aceita **SIM / NÃO**; tipo de contato aceita **FAMILIAR / AMIGO**.
Campos opcionais podem ficar vazios; **NÃO INFORMADO** significa que não haverá marcação no crachá.
Confira os avisos de nomes e poltronas repetidos antes de continuar.

A coluna **Gerar crachá** controla quem terá crachá impresso.
A lista de embarque permite escolher todos ou apenas os selecionados; o mapa impresso inclui todos.

**Trajeto:** todos começam como **IDA E VOLTA**. Observações claras como “só vai”, “não retornarei com a caravana” ou “voltarei de carro” sugerem SOMENTE IDA; “só volta”, “não irei com a caravana” ou “irei por conta própria” sugerem SOMENTE VOLTA. A leitura acontece na importação; confira e altere a coluna **Trajeto** quando necessário. Uma escolha de trajeto já preenchida tem prioridade sobre as observações. Textos ambíguos mantêm Ida e volta.
Dois passageiros podem usar a mesma poltrona se um viajar somente na ida e o outro somente na volta. Se houver conflito no mesmo sentido, corrija o trajeto ou a poltrona. O mapa impresso será separado por ida e volta quando houver passageiros de um só sentido; na Lista de Caravanistas, o sentido que não será utilizado aparece com hífen.

### 3 · Escolha ou revise os assentos
Organize **um ônibus por planilha** e confirme o modelo:

- **Leito 43:** piso superior 1–31 e inferior 32–43.
- **Leito 46:** piso superior 1–34 e inferior 35–46.
- **Misto 56:** semi-leito no superior 1–44 e leito no inferior 45–56.
- **LD Semi-leito 44:** 44 lugares, com escada e banheiro conforme o mapa do modelo.

Se todos já vierem com poltrona, responda **Deseja revisar os assentos importados?**
Escolha **Não** para manter as poltronas e ocultar a marcação, ou **Sim** para abrir o mapa.
Se houver passageiros sem assento, o mapa permanece disponível.
Trocar o modelo não apaga nem redistribui as poltronas existentes.

Na seleção de passageiros, **🟢** indica quem já tem poltrona e **🟠** quem ainda precisa de assento.
Escolha um passageiro e clique em uma poltrona livre; para mudar de lugar, clique em outra livre.
O botão **Liberar poltrona deste passageiro**, logo abaixo da seleção, remove a marcação atual.
No mapa, **×** indica ocupação e **✓** indica a poltrona do passageiro selecionado.
Os nomes aparecem abaixo dos assentos, com os prefixos preservados.

### 4 · Informe os dados do ônibus
Preencha plataforma, horário de retorno, responsável e telefone com DDD.
Use o horário no formato **HH:MM**. Para informar telefone do responsável, preencha DDD e número;
se não precisar desse contato, deixe ambos vazios.

### 5 · Gere e imprima os documentos
- **Crachás:** clique em **Gerar Crachás**, confira as páginas na prévia e baixe o PDF. Nomes longos de participantes, contatos e responsável podem ter nomes do meio abreviados para caber, preservando prefixos, primeiro nome e último sobrenome. O nome completo permanece na planilha.
- **Lista de embarque:** escolha os participantes e a ordem por missão e nome, nome ou poltrona.
  Informe a Data de Saída da Caravana, se desejar, e gere a Lista de Caravanistas, com o responsável e seu telefone no cabeçalho, para impressão em A4, com espaços de conferência de ida, retorno e observações.
- **Mapa do ônibus:** baixe o PDF preenchido em A4, com nomes completos em negrito e uma página por piso.
  Passageiros sem poltrona aparecem em uma relação adicional.
  Corrija conflitos de poltrona no mesmo sentido, números inválidos ou fora do modelo antes de imprimir o mapa.

### Guarde suas escolhas antes de sair
Use **Baixar planilha com minhas escolhas** para guardar as correções, seleções e poltronas,
inclusive dos passageiros desmarcados para crachá. Você pode importar esse Excel depois para continuar.
**As escolhas ficam nesta sessão; não há um cadastro permanente de viagens.**
Os dados do ônibus da etapa 4 não fazem parte do Excel exportado e precisam ser preenchidos novamente.
Após alterar dados, gere novamente os documentos afetados para imprimir a versão atual.''')
st.subheader('1 · Importar participantes')
st.download_button('📄 Baixar Planilha Modelo', (BASE / 'modelo_planilha.xlsx').read_bytes(),
                   file_name='modelo_planilha.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                   on_click='ignore')
origem_dados = st.radio('Origem dos participantes', ['Planilha Excel', 'Google Sheets'],
                        horizontal=True, key='origem_dados')
if origem_dados == 'Google Sheets':
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
        df = ordenar_por_missao(maiusculas_participantes(ler_planilha(conteudo)))
        st.subheader('2 · Conferir a listagem')
        st.caption('Listagem por missão e nome em ordem alfabética. Corrija os dados diretamente na tabela e desmarque Gerar crachá para quem não precisa de impressão.')
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
                      'Telefone Contato de Emergência', 'Trajeto', 'Observações', CONVENIO, CONTATO)],
            column_order=[coluna_selecao, 'Nome Completo', 'Trajeto', CONVENIO, CONTATO,
                          *[c for c in df.columns if c not in ('Nome Completo', 'Trajeto', CONVENIO, CONTATO)]],
            column_config={coluna_selecao: st.column_config.CheckboxColumn('Gerar crachá', default=True),
                           '_ID_Inscricao': None,
                           'Trajeto': st.column_config.SelectboxColumn('Trajeto', options=list(TRAJETOS), required=True),
                           **{c: st.column_config.SelectboxColumn(
                'Possui convênio?' if c == CONVENIO else 'Tipo de contato',
                options=['NÃO INFORMADO', *opcoes], required=False, format_func=lambda v: str(v or '').upper(),
            ) for c, opcoes in OPCOES.items()}},
        )
        for campo in OPCOES:
            df[campo] = df[campo].replace('NÃO INFORMADO', '')
        padronizada = ordenar_por_missao(maiusculas_participantes(preparar_trajetos(df)))
        if not df.equals(padronizada):
            st.session_state['mapa_tabela_' + identidade_planilha] = padronizada
            chave_revisao = 'mapa_revisao_' + identidade_planilha
            st.session_state[chave_revisao] = revisao_mapa + 1
            st.rerun()
        tabela_mapa = padronizada.copy(deep=True)
        if origem_dados == 'Google Sheets':
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
    st.subheader('Lista de Caravanistas para impressão')
    st.caption('Lista em A4 com nome, unidade, poltrona e espaços para ida, retorno e observações.')
    abrangencia = st.selectbox('Participantes da lista', ['Todos os participantes', 'Apenas os selecionados'], key='abrangencia_lista')
    ordem_lista = st.selectbox('Ordenar lista por', ['Missão e nome', 'Nome', 'Poltrona'], key='ordem_lista')
    data_viagem = st.date_input('Data de Saída da Caravana (opcional)', value=None, format='DD/MM/YYYY')
    dados_lista = df if abrangencia == 'Todos os participantes' else selecionados
    contato_lista = (f'({ddd_validado}) {telefone_validado}' if ddd_validado else '') if erro_caravana is None else ' '.join(p.strip() for p in (ddd, telefone) if p.strip())
    assinatura_lista = sha256(dados_lista.to_json().encode()
                             + repr((ordem_lista, data_viagem, responsavel, contato_lista)).encode()).hexdigest()
    if st.session_state.get('assinatura_lista') != assinatura_lista:
        st.session_state.pop('lista_pdf', None)
        st.session_state['assinatura_lista'] = assinatura_lista
    st.caption(f'{len(dados_lista)} participantes na lista. Responsável: {responsavel.strip() or "preenchimento à mão"}.')
    if st.button('Gerar lista de conferência', key='gerar_lista', disabled=dados_lista.empty):
        st.session_state.pop('lista_pdf', None)
        try:
            st.session_state['lista_pdf'] = gerar_lista(dados_lista, ordem_lista, data_viagem, responsavel, contato_lista)
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
