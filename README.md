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

## Testes automáticos no GitHub

O fluxo **Testes automáticos**, em `.github/workflows/testes.yml`, executa a suíte em Linux com Python 3.11 (mínimo documentado) e 3.13 (versão usada no Streamlit). Ele roda em propostas destinadas à `main`, alterações em `main` e `melhorias-geracao` e também pode ser iniciado manualmente pela aba **Actions** após estar na branch padrão.

Cada proposta mostra dois resultados: **Testes Python 3.11** e **Testes Python 3.13**. Antes de incorporar uma atualização, confira se ambos estão verdes. Em caso de falha, abra **Details** para consultar qual teste ou instalação falhou. Execuções anteriores da mesma referência são canceladas quando uma nova atualização chega.

O fluxo usa apenas leitura do repositório, sem segredos, e testa dados fictícios. Não publica o aplicativo. Os resultados são informativos até serem configurados como verificações obrigatórias em uma regra de proteção da `main`; este fluxo não altera essa configuração do repositório.

Referência: [documentação oficial de testes Python no GitHub Actions](https://docs.github.com/en/actions/tutorials/build-and-test-code/python).
