# Teste privado de importação do Google Forms via Sheets

Este teste conecta cada organizador à própria conta Google e lê a planilha privada de respostas. O Excel continua disponível. A conexão real depende de configurar um projeto Google Cloud; testes locais com respostas simuladas não confirmam o login nem as permissões de uma conta real.

## 1. Criar o projeto

Abra o [Google Cloud Console](https://console.cloud.google.com/), entre com sua conta e crie um projeto chamado **Gerador de Crachás - Teste**. No projeto, abra a [biblioteca de APIs](https://console.cloud.google.com/apis/library), procure **Google Sheets API** e ative-a. Não é necessário publicar a planilha ou criar uma conta de serviço para este fluxo.

## 2. Configurar o consentimento

Em **Google Auth Platform**, configure o nome do aplicativo e seus contatos. Se a organização não permitir público interno, escolha público **External**, mantenha em **Testing** e adicione apenas os e-mails dos organizadores de teste em **Test users**.

Em **Data Access**, configure o escopo `https://www.googleapis.com/auth/spreadsheets.readonly`, além dos escopos de identidade `openid`, `email` e `profile` usados no login. Esse escopo permite leitura das planilhas acessíveis à conta; não concede edição. O aplicativo consulta apenas o link informado e a aba escolhida. A autorização não é limitada a um arquivo específico pelo Google.

Não publique o consentimento para uso geral nesta fase. A liberação pública de escopos sensíveis pode exigir verificação do Google.

## 3. Criar o cliente OAuth

Abra **Clients → Create client → Web application**. Cadastre em **Authorized redirect URIs** exatamente:

```text
https://SEU-APP-DE-TESTE.streamlit.app/oauth2callback
```

Substitua o endereço pelo aplicativo conectado à branch `melhorias-geracao`, incluindo `https://` e `/oauth2callback`. Use o mesmo endereço na configuração do Streamlit. Guarde o Client ID e o Client secret fora do GitHub. Não envie client secret ou tokens pelo chat.

Para testar localmente, registre separadamente `http://localhost:8501/oauth2callback` e use esse endereço nos segredos locais.

## 4. Configurar o Streamlit de teste

Confira que o aplicativo de teste usa a branch `melhorias-geracao` e Python 3.13. Abra **Manage app → Settings → Secrets** e copie o modelo de [.streamlit/secrets.example.toml](.streamlit/secrets.example.toml), preservando outros segredos existentes. Substitua URL, Client ID, Client secret e cookie secret. Gere o cookie secret com um gerenciador de senhas ou, em um terminal local:

```sh
python -c "import secrets; print(secrets.token_urlsafe(48))"
```

Cole o resultado somente no campo `cookie_secret` dos segredos. Não coloque valores reais no arquivo de exemplo nem no README. Salve e reinicie o aplicativo de teste para carregar a configuração.

O suporte a autenticação é instalado por `streamlit[auth]`. O Streamlit gerencia o retorno OAuth, `state`, `nonce` e a sessão de identidade; o aplicativo usa o access token apenas no servidor para as chamadas de leitura. Não há refresh token persistido: se a autorização expirar ou estiver ausente após retomar uma sessão, exporte o trabalho e entre novamente.

## 5. Fazer o teste

1. Escolha **Google Sheets (teste)** em Origem dos participantes.
2. Clique em **Conectar minha conta Google** e entre com um dos usuários cadastrados para teste.
3. Cole o link da planilha de respostas; clique em **Ler abas da planilha** e escolha a aba correta.
4. Clique em **Buscar respostas no Google**.
5. Associe cada campo à pergunta do Forms. Só Nome Completo é obrigatório; campos opcionais podem ficar sem associação. As escolhas de convênio e contato precisam usar Sim/Não e Familiar/Amigo.
6. Escolha um **identificador único e estável por inscrição**, como o carimbo de data/hora se ele for exclusivo. Não altere esse identificador depois. IDs vazios ou repetidos bloqueiam a atualização para evitar misturar pessoas.
7. Clique em **Importar / atualizar participantes**. Confira nomes, zeros dos telefones e campos opcionais.
8. Atribua uma poltrona e corrija algum campo. Adicione uma resposta fictícia ao Forms; busque as respostas novamente e clique em Importar/atualizar. A nova inscrição deve entrar e a poltrona/correção deve permanecer.
9. Baixe a planilha com as escolhas antes de sair. Teste outra conta Google: ela deve ver somente os dados da própria sessão e conseguir ler apenas arquivos autorizados àquela conta.

## Limites deste protótipo

- Sem banco permanente, histórico ou cadastro salvo de viagens. A sessão e o Excel exportado são os únicos meios atuais de guardar as correções.
- A atualização compara por ID a origem anterior, a tabela corrigida e as novas respostas. Campos corrigidos localmente são preservados; campos intactos acompanham mudanças da origem. Novas inscrições são adicionadas.
- Inscrições removidas da planilha são mantidas na sessão com aviso para conferência. IDs alterados são tratados como novas inscrições; não se faz deduplicação por nome.
- Não há escrita no Google nem sincronização entre organizadores. A origem não recebe as escolhas de poltrona.
- Até 400 respostas por aba, primeiras 100 colunas e resposta de até 5 MB. Títulos na primeira linha.
- Nenhum cache compartilhado de respostas ou tokens. Trocar de conta limpa os dados privados da sessão. O identificador interno oculto é exportado com os participantes para rastreabilidade; não é uma credencial.

## Referências

- [Login e configuração de tokens no Streamlit](https://docs.streamlit.io/develop/api-reference/user/st.login)
- [Escopos da API Google Sheets](https://developers.google.com/workspace/sheets/api/scopes)
- [Configuração OAuth para aplicativos web no Google](https://developers.google.com/identity/protocols/oauth2/web-server)

## Páginas públicas e endereço de produção

As páginas de privacidade e termos são abertas sem login e sem importar participantes. O contato público é douglas.vital.28@gmail.com. Os links aparecem também na página inicial.

Após publicar esta alteração na main e conferir os endereços em uma janela sem login, use no Branding:

- Página inicial: https://gerador-crachas-m7wgvbxff7ufbpivmfzfhy.streamlit.app/
- Política de privacidade: https://gerador-crachas-m7wgvbxff7ufbpivmfzfhy.streamlit.app/?pagina=privacidade
- Termos de uso: https://gerador-crachas-m7wgvbxff7ufbpivmfzfhy.streamlit.app/?pagina=termos

A existência dessas páginas não confirma publicação ou verificação OAuth pelo Google. Atenda às exigências exibidas pelo Console. Para uso público verificado, podem ser exigidos comprovação dos domínios e revisão dos escopos. Para um pequeno grupo conhecido pessoalmente, confira a exceção de uso pessoal nas regras atuais do Google; limites e avisos de aplicativo não verificado ainda podem se aplicar.

O endereço de retorno OAuth da produção deve ser cadastrado no cliente e repetido em redirect_uri dos Secrets do app principal:

`https://gerador-crachas-m7wgvbxff7ufbpivmfzfhy.streamlit.app/oauth2callback`

Preserve o retorno do teste enquanto ele estiver em uso. Nunca publique client_secret, cookie_secret ou tokens no repositório ou nas páginas públicas.
