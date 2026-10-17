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
4. Envie a planilha (até 5 MB e 400 participantes), confira a prévia e informe os dados da caravana.
5. Gere e baixe o PDF. Confira antes de imprimir.

Os campos opcionais podem ficar vazios. A fonte é reduzida para acomodar textos maiores; se ficar pequena demais para leitura, o aplicativo solicita abreviar o conteúdo. As colunas `Possuí Plano de Saúde?` (Sim/Não) e `Tipo de Contato` (Familiar/Amigo) preenchem automaticamente as marcações do PDF. Você também pode selecionar essas opções na tabela do aplicativo antes de gerar. Valores em branco deixam as opções desmarcadas; valores inválidos indicam a linha da planilha a corrigir. Maiúsculas/minúsculas e acentos são normalizados. Planilhas antigas sem essas duas colunas continuam funcionando: faça as escolhas na tabela.

O PDF é gerado em memória, sem arquivos temporários no servidor, e fica disponível na sessão do usuário. Alterar a planilha, as seleções da tabela ou os dados da caravana remove o resultado anterior para evitar baixar um PDF desatualizado. Recarregar a página pode encerrar a sessão; baixe o arquivo para guardá-lo. O limite de participantes é uma proteção inicial, não uma garantia de capacidade sob muitos acessos simultâneos.

## Testes

```sh
pip install -r requirements-dev.txt
python -m unittest discover -s tests -v
```

Os testes verificam validação, zeros à esquerda, textos grandes, quantidades de 1, 4 e 5 participantes, campos do formulário e preservação de valores entre páginas. Também verificam todas as combinações de marcação nos quatro crachás, em múltiplas páginas, crachás vazios, validação das opções e seleção pela interface. Utilizam dados fictícios.
