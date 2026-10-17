"""Informações públicas acessíveis antes de importar dados ou conectar o Google."""
CONTATO = 'douglas.vital.28@gmail.com'

PRIVACIDADE = '''
# Política de Privacidade

**Gerador de Crachás — Caravanas**  
Última atualização: 17 de outubro de 2026.  
Responsável pelo projeto e contato: Douglas Vital — [douglas.vital.28@gmail.com](mailto:douglas.vital.28@gmail.com).

## Finalidade
Este projeto beneficente auxilia organizadores de caravanas a conferir participantes,
organizar poltronas e trajetos e gerar crachás, mapas de ônibus e listas de embarque.

## Dados utilizados
O organizador pode fornecer nome, unidade religiosa, missão, poltrona, trajeto,
observações, convênio de saúde e seu telefone, nome e telefone do contato de emergência.
Também pode preencher os dados da viagem e o nome e telefone do responsável.
Algumas dessas informações podem revelar religião ou informações de saúde.
Forneça apenas os dados necessários e que você esteja autorizado a utilizar, inclusive de terceiros.

Os dados vêm de um Excel enviado pelo organizador ou de uma planilha Google que ele autorize consultar.
No login Google, o aplicativo utiliza dados de identificação da conta, como nome, e-mail e
identificador da conta, além de um token de acesso para realizar a leitura autorizada.

## Acesso ao Google Sheets
A autorização solicita identificação da conta e leitura das planilhas pelo escopo
`https://www.googleapis.com/auth/spreadsheets.readonly`.
Essa permissão permite ler planilhas acessíveis à conta; não fica limitada a um único arquivo pelo Google.
O aplicativo consulta o link e a aba selecionados pelo organizador para importar e atualizar participantes.
**O aplicativo não altera a planilha de origem, não solicita sua senha Google e não precisa que a planilha seja pública.**

Os dados recebidos do Google são utilizados somente nas funções de organização e geração de documentos
solicitadas pelo usuário. Não são vendidos, utilizados em publicidade nem usados para treinar modelos de inteligência artificial.
O uso e a transferência das informações recebidas das APIs Google seguem a
[Política de Dados do Usuário dos Serviços de API do Google](https://developers.google.com/terms/api-services-user-data-policy),
inclusive os requisitos de Uso Limitado.

## Sessão, arquivos e armazenamento
Os participantes, correções, poltronas e documentos gerados são processados no servidor e mantidos na sessão de trabalho.
O aplicativo não possui banco permanente de viagens nem histórico de participantes e não grava esses arquivos
em um cadastro permanente próprio. Não há armazenamento próprio permanente de refresh tokens.
O login usa cookies de autenticação gerenciados pelo Streamlit; dados de login podem permanecer entre visitas.

A infraestrutura de hospedagem pode processar informações técnicas, cookies e registros de funcionamento
conforme suas próprias políticas. O aplicativo é hospedado no Streamlit Community Cloud; o login e a leitura
usam serviços Google. Consulte também as políticas de [Streamlit/Snowflake](https://www.snowflake.com/en/legal/privacy/privacy-policy/)
e [Google](https://policies.google.com/privacy).

O organizador pode baixar os dados em Excel e os documentos em PDF. Esses arquivos e as impressões
ficam sob sua responsabilidade: armazene-os com cuidado, compartilhe apenas com pessoas autorizadas
e elimine-os quando não forem mais necessários. Não envie planilhas completas com dados sensíveis em pedidos de suporte.

## Compartilhamento e acesso
O aplicativo não publica listas de passageiros nem disponibiliza as planilhas a outros organizadores.
Os dados são processados pelos serviços de hospedagem e pelos serviços Google necessários ao funcionamento.
O compartilhamento dos arquivos baixados e das impressões é decidido pelo organizador.

## Controle, correção e revogação
Você pode corrigir os dados na listagem e encerrar a conexão pelo botão **Sair da conta Google**,
que limpa os dados de trabalho da sessão atual. Sair não remove as planilhas originais nem os arquivos já baixados.
Você também pode revogar a permissão do aplicativo nas
[conexões da Conta Google](https://myaccount.google.com/connections).

Para dúvidas sobre privacidade ou solicitações relacionadas a acesso, correção ou eliminação,
escreva para **douglas.vital.28@gmail.com**. Se você é passageiro, procure também o organizador
que recolheu suas informações: ele mantém a planilha original e os arquivos exportados.
Não conseguimos apagar, pelo aplicativo, cópias mantidas por terceiros ou no seu dispositivo.

## Atualizações
Esta política pode ser atualizada quando as funções ou o tratamento de dados mudarem.
A data acima identifica a versão publicada.
'''

TERMOS = '''
# Termos de Uso

**Gerador de Crachás — Caravanas**  
Última atualização: 17 de outubro de 2026.  
Contato: [douglas.vital.28@gmail.com](mailto:douglas.vital.28@gmail.com).

## Sobre o projeto
Este projeto beneficente oferece uma ferramenta de apoio à organização de caravanas:
importação de participantes, conferência de dados, escolha de poltronas e geração de documentos.
A referência a instituições ou destinos nos crachás descreve a finalidade da viagem e não representa,
por si só, vínculo oficial ou endosso dessas instituições ao aplicativo.

## Uso responsável
Utilize apenas informações e planilhas que você esteja autorizado a tratar.
O organizador deve informar os participantes sobre o uso dos dados, proteger informações pessoais
e compartilhar listas, contatos e documentos apenas com pessoas que precisem dessas informações.
A autorização de uma conta Google não substitui essa responsabilidade perante os passageiros.
Não use o serviço para acessar dados de outras pessoas sem autorização ou praticar atividades ilícitas.

## Confira antes de imprimir
Revise nomes, contatos, convênios, missão, trajetos, poltronas e dados da viagem antes de utilizar os documentos.
A classificação de trajeto pelas observações é uma sugestão por regras de texto e pode não interpretar
corretamente toda frase. Prefixos, abreviações e ordenação também devem ser conferidos pelo organizador.
A poltrona só pode ser compartilhada entre passageiros de sentidos diferentes, conforme as escolhas registradas.
O aplicativo auxilia a conferência; a confirmação do embarque e do transporte cabe à organização da viagem.

## Salve seu trabalho
As alterações ficam na sessão. Baixe o Excel com suas escolhas e os documentos necessários antes de sair.
O Excel exportado não inclui os dados do ônibus preenchidos na etapa 4.
A conexão Google lê a origem, mas não grava nela as correções ou poltronas escolhidas no aplicativo.
Não há garantia de recuperação de trabalho não exportado nem de disponibilidade contínua;
o serviço depende também de Streamlit, Google e conexão à internet.

## Privacidade e suporte
O tratamento dos dados é descrito na [Política de Privacidade](?pagina=privacidade).
Para suporte ou dúvidas sobre estes termos, entre em contato pelo e-mail
**douglas.vital.28@gmail.com**. Alterações nestes termos serão publicadas com atualização da data acima.
'''


def mostrar_pagina_publica():
    import streamlit as st
    pagina = st.query_params.get('pagina')
    if pagina not in ('privacidade', 'termos'):
        return False
    st.markdown('[← Voltar ao aplicativo](?)')
    st.markdown(PRIVACIDADE if pagina == 'privacidade' else TERMOS)
    st.markdown('[Política de Privacidade](?pagina=privacidade) · [Termos de Uso](?pagina=termos)')
    return True


def links_publicos():
    import streamlit as st
    st.markdown('[Política de Privacidade](?pagina=privacidade) · [Termos de Uso](?pagina=termos) · '
                '[Contato](mailto:douglas.vital.28@gmail.com)')
