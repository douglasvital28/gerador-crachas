# Gerador de Crachás

Aplicativo Streamlit para gerar crachás de caravanas ao Solo Sagrado de Guarapiranga a partir de uma planilha Excel. Cada página contém até quatro crachás.

## Executar

Use Python 3.11 ou superior:

```sh
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app_streamlit.py
```

No Streamlit Community Cloud, selecione `app_streamlit.py` como arquivo principal e Python 3.11 ou superior. Mantenha `cracha_template.pdf` e `modelo_planilha.xlsx` no repositório.

## Utilizar

1. Baixe a planilha modelo no aplicativo e mantenha os nomes das colunas.
2. Preencha os participantes. Cada linha preenchida precisa de Nome Completo; linhas totalmente vazias são ignoradas.
3. Formate telefones e poltronas como Texto no Excel antes de digitar, para preservar zeros à esquerda. Zeros já perdidos pelo Excel não podem ser recuperados pelo aplicativo.
4. Envie a planilha (até 5 MB e 400 participantes), corrija nome, unidade, poltrona, plano, contato e telefones na tabela e informe os dados da caravana. Nome Completo precisa continuar preenchido. Colunas extras do arquivo são preservadas, mas não editáveis.
5. Na coluna **Gerar crachá**, desmarque quem não precisa de impressão. Todos começam marcados. O PDF contém apenas os selecionados, na ordem da planilha, agrupados em quatro por página. Marque ao menos uma pessoa. Gere os crachás e confira a prévia, selecionando uma página por vez. Baixe o PDF para imprimir em tamanho real.
6. Use **Baixar planilha com minhas escolhas** para guardar todos os participantes com as correções e seleções de convênio e contato, inclusive os desmarcados para impressão e reutilizar o Excel em outra geração. A seleção para impressão não é salva no Excel: ao reenviar o arquivo, todos começam marcados. A exportação é uma nova tabela de participantes: não inclui os campos da caravana nem preserva outras abas, fórmulas ou a formatação do arquivo enviado.

Os campos opcionais podem ficar vazios. A fonte é reduzida para acomodar textos maiores; se ficar pequena demais para leitura, o aplicativo solicita abreviar o conteúdo. As colunas `Possuí Plano de Saúde?` (Sim/Não) e `Tipo de Contato` (Familiar/Amigo) preenchem automaticamente as marcações do PDF. Você também pode selecionar essas opções na tabela do aplicativo antes de gerar. Valores em branco deixam as opções desmarcadas; valores inválidos indicam a linha da planilha a corrigir. Maiúsculas/minúsculas e acentos são normalizados. Planilhas antigas sem essas duas colunas continuam funcionando: faça as escolhas na tabela.

Nos telefones do plano e do contato de emergência, números brasileiros completos são separados entre DDD e telefone no PDF: `(11) 91234-5678` vira DDD `11` e telefone `912345678`. Também são reconhecidos prefixos `+55`, `0055`, `55` e um zero antes do DDD, com comprimento compatível. Números locais, serviços 0800/0300/0500/0900, números de outros países e formatos com ramal são preservados sem separação. A identificação por formato não verifica se a linha telefônica existe.

O aplicativo avisa sobre nomes repetidos (ignorando acentos, caixa e espaços extras) e poltronas repetidas (tratando `001` e `1` como a mesma poltrona). O aviso mostra as linhas originais da planilha, inclusive quando há linhas vazias no meio, e não bloqueia a geração nem remove participantes. Confira a origem antes de corrigir possíveis homônimos ou números de poltrona repetidos entre ônibus.

Horário, DDD e telefone do responsável são opcionais. Se informar horário, use `HH:MM` (por exemplo `18:00`; `8:05` é normalizado para `08:05`). Para telefone, informe DDD de dois dígitos e número com 8 ou 9 dígitos, ou o número completo com DDD no próprio campo telefone. Se os dois campos contiverem DDD, eles precisam coincidir. Dados incompletos ou formatos inválidos impedem gerar o PDF até a correção. Essa validação verifica formato e consistência, não se o telefone está ativo.

A prévia renderiza somente a página selecionada em memória. Trocar de página mantém o PDF; alterar os dados remove o resultado anterior e a próxima geração começa na página 1. A prévia é uma imagem para conferência; o download continua sendo o PDF editável original. Se a prévia falhar, o download permanece disponível.

O PDF é gerado em memória, sem arquivos temporários no servidor, e fica disponível na sessão do usuário. Alterar a planilha, as correções, a seleção para impressão ou os dados da caravana remove o resultado anterior para evitar baixar um PDF desatualizado. Recarregar a página pode encerrar a sessão; baixe o arquivo para guardá-lo. O limite de participantes é uma proteção inicial, não uma garantia de capacidade sob muitos acessos simultâneos.

## Testes

```sh
pip install -r requirements-dev.txt
python -m unittest discover -s tests -v
```

Os testes verificam validação, zeros à esquerda, textos grandes, quantidades de 1, 4 e 5 participantes, campos do formulário e preservação de valores entre páginas. Também verificam todas as combinações de marcação nos quatro crachás, em múltiplas páginas, crachás vazios, validação das opções e seleção pela interface. Utilizam dados fictícios.

## Lista de conferência em papel

Após enviar e corrigir os participantes, use **Lista de conferência para impressão**. Escolha todos os participantes ou apenas os marcados na coluna Gerar crachá, ordene por nome ou poltrona e informe a data da viagem se desejar. O nome do responsável vem do campo da caravana; se estiver vazio, o PDF reserva uma linha para preenchimento à mão.

Clique em **Gerar lista de conferência** e depois **Baixar lista de conferência**. O PDF é A4, com nome, unidade, poltrona, quadrados para ida/retorno e espaço para observações. Cabeçalho, total de passageiros e títulos das colunas são repetidos nas páginas. Poltronas são ordenadas naturalmente (1, 2, 10); poltronas vazias ficam no final.

A lista usa os dados corrigidos na tabela e pode ser gerada sem gerar os crachás. Não contém telefones nem informações de convênio. Mudar os dados incluídos, responsável, ordenação ou data remove o download anterior para evitar uma lista desatualizada. A lista é destinada ao controle manual no papel: as marcações feitas na impressão não são salvas no aplicativo.

## Mapa clicável de poltronas

A etapa **3 · Escolher assentos**, depois da conferência da tabela, organiza um ônibus por planilha. Selecione Leito 43, Leito 46, Misto 56 ou LD Semi-leito 44 lugares, escolha o passageiro e clique no número de uma poltrona livre. O mapa segue as referências fornecidas: no modelo 43, superior 1–31 e inferior 32–43; no modelo 46, superior 1–34 e inferior 35–46. No modelo G8 Misto, o piso superior tem semi-leito 1–44 em disposição 2 + 2, e o inferior tem leito 45–56 em disposição 2 + 1. O LD 44 usa duplas dos dois lados do corredor, escada na frente à direita e banheiro ao fundo, conforme a referência. As escadas e as áreas sem poltrona respeitam as posições dos esquemas.

As abas alternam entre os pisos. Um número livre pode ser escolhido; `×` indica ocupação e `✓` a poltrona do passageiro atual. Lugares ocupados mostram o primeiro nome e o último sobrenome abaixo do número. O nome completo continua disponível na ajuda ao passar sobre a poltrona. Lugares de outros passageiros ficam indisponíveis. Clicar em outro lugar livre move o passageiro e libera o anterior; **Liberar poltrona deste passageiro** remove somente a atribuição dele.

Poltronas importadas são preservadas, e `001` é reconhecida como 1. Duplicidades e números fora do modelo geram avisos, sem apagar os dados. Trocar de modelo não renumera nem remove lugares. Arquivos sem a coluna Poltrona também são aceitos: o campo é criado vazio para preencher pelo mapa ou tabela.

O mapa mantém as outras correções e a seleção de impressão da tabela. Atribuições aparecem nos crachás, lista de conferência e planilha exportada; resultados anteriores são invalidados quando os dados mudam. As escolhas ficam na sessão individual e não são sincronizadas com outros usuários. Baixe a planilha com as escolhas para guardá-las.

## Google Sheets (teste)

A origem **Google Sheets (teste)** permite conectar uma conta Google e importar uma aba privada de respostas do Forms, associando perguntas aos campos do aplicativo. O login real depende da configuração descrita em [CONFIGURAR_GOOGLE.md](CONFIGURAR_GOOGLE.md). Sem ela, a tela informa que a conexão está em preparação; o upload Excel continua funcionando.

O protótipo preserva correções e poltronas na atualização por identificador único da inscrição, com isolamento por conta/planilha/aba. Não inclui armazenamento permanente de viagens: exporte o Excel antes de sair. A autorização real e a leitura de planilhas precisam ser conferidas no aplicativo de teste após cadastrar o projeto Google Cloud e seus usuários de teste.


## Jornada de uso

A tela agora segue cinco etapas numeradas: **1. Importar participantes** (Excel ou Google Sheets), **2. Conferir a listagem**, **3. Escolher assentos**, **4. Informar os dados do ônibus** e **5. Gerar documentos da viagem**. Instruções ficam recolhidas em Como usar. A etapa final abre primeiro a lista de embarque; a aba Crachás mantém geração, download e prévia.

Nomes, unidade, plano, contatos e textos de poltrona são padronizados em MAIÚSCULAS antes da conferência e após correções. Plataforma e responsável também são convertidos ao editar. Os dados padronizados são usados no mapa, PDF e planilha exportada. Zeros dos telefones e identificadores internos de inscrição são preservados. Convênio e tipo de contato aparecem em maiúsculas na interface e no Excel; internamente mantêm os valores compatíveis com a marcação do PDF. Não informado corresponde a um campo vazio.

As etapas de assentos, ônibus e documentos aparecem depois de uma importação válida. Um nome apagado precisa ser corrigido antes de prosseguir. A atualização do Google compara textos sem diferenciar maiúsculas/minúsculas para que essa padronização não seja confundida com uma correção manual.

### Revisão e impressão dos assentos

Quando a importação já traz todos os assentos preenchidos, escolha se deseja revisá-los. Ao escolher Não, o mapa interativo fica oculto e as poltronas são mantidas. Se houver passageiros sem assento, a marcação permanece visível. Confirme o modelo do ônibus mesmo ao dispensar a revisão.

Na etapa 5, a aba **Mapa do ônibus** permite baixar o mapa em PDF A4, com nomes completos em negrito e uma página por piso. Todos os passageiros entram no mapa, independentemente da seleção de crachás. Quem estiver sem poltrona aparece em uma relação adicional. Poltronas duplicadas, inválidas ou fora do modelo precisam ser corrigidas antes de imprimir.

### Ordem da listagem

A conferência, a seleção de passageiros e a exportação seguem missão e nome: Ministros, Professores, Coordenadores, Assistentes de Ministros, Assistentes de Família, Membros, Frequentadores e Primeira Vez. Dentro de cada categoria, os nomes ficam em ordem alfabética, sem considerar prefixos ou acentos na comparação. A coluna opcional Missão/Missao identifica a categoria; prefixos MIN. e PROF. são usados quando a missão não é reconhecida. Demais missões e campos vazios ficam ao final. Poltronas, identificadores e seleção de crachás acompanham cada passageiro. Na lista impressa, Missão e nome é a ordem padrão; Nome e Poltrona continuam disponíveis.

### Passageiros somente de ida ou de volta

A coluna Trajeto permite escolher IDA E VOLTA (padrão), SOMENTE IDA ou SOMENTE VOLTA. Na importação, Observações/Observacoes/Observacao pode sugerir o trajeto para expressões claras como só vai, somente ida, só volta e somente volta. Textos ambíguos mantêm o padrão e devem ser conferidos pelo organizador. Um trajeto já preenchido tem prioridade e é preservado na exportação e reimportação. No Google Sheets, associe Observações e, se houver, Trajeto às perguntas correspondentes.

Dois passageiros podem compartilhar a poltrona quando um viaja somente na ida e o outro somente na volta. O mapa mostra os dois nomes com seus trajetos e bloqueia ocupações simultâneas no mesmo sentido. Ao mudar um trajeto manualmente, confira os avisos de conflito. A Lista de Caravanistas identifica os trajetos parciais e coloca hífen no sentido que não será utilizado; o mapa em PDF separa ida e volta quando necessário.
