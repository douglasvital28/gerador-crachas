import unittest
import pandas as pd
from padronizacao import ordenar_por_missao
from conferencia import ordenar_participantes


class OrganizacaoTest(unittest.TestCase):
    def test_hierarquia_nome_acentos_e_dados_preservados(self):
        missoes = ['Primeira Vez', 'Frequentador', 'Membro', 'Assistente de Família',
                   'Responsável de Setor', 'Assistentes de Ministros', 'Coordenador(a)', 'Professora', 'Ministro(a)']
        tabela = pd.DataFrame([{'Missão': m, 'Nome Completo': n, 'Poltrona': str(i),
                               '_ID_Inscricao': str(i), 'Gerar crachá': i % 2 == 0}
                              for i, (m, n) in enumerate((m, n) for m in missoes for n in ['ZÉLIA', 'ÁLVARO'])])
        ordenada = ordenar_por_missao(tabela)
        self.assertEqual(ordenada['Missão'].tolist(), [m for m in reversed(missoes) for _ in range(2)])
        self.assertEqual(ordenada['Nome Completo'].tolist(), ['ÁLVARO', 'ZÉLIA']*len(missoes))
        self.assertEqual(tabela.set_index('_ID_Inscricao').to_dict('index'), ordenada.set_index('_ID_Inscricao').to_dict('index'))
        self.assertEqual(ordenar_participantes(tabela, 'Missão e nome').to_dict('records'), ordenada.to_dict('records'))
        self.assertEqual(ordenar_por_missao(ordenada).to_dict('records'), ordenada.to_dict('records'))

    def test_atualiza_regra_antiga_em_memoria_sem_perder_poltronas(self):
        from io import BytesIO
        from hashlib import sha256
        from unittest.mock import patch
        from streamlit.testing.v1 import AppTest
        from geracao import BASE, COLUNAS
        from test_geracao import pessoa, excel
        import padronizacao
        linhas = [pessoa('Ana') + ['Assistente de Família'],
                  pessoa('Bruno') + ['Responsável de Setor'],
                  pessoa('Carla') + ['Assistente de Ministro']]
        for i, linha in enumerate(linhas):
            linha[2] = str(i+1)
        dados = excel(linhas, [*COLUNAS, 'Missão'])
        ident = sha256(dados).hexdigest()
        # Simula a atualização do app com a implementação anterior ainda importada.
        with patch('streamlit.file_uploader', return_value=BytesIO(dados)), \
             patch.object(padronizacao, 'VERSAO_ORDENACAO', None), \
             patch.object(padronizacao, 'ordenar_por_missao', lambda df: df):
            app = AppTest.from_file(str(BASE / 'app_streamlit.py')).run()
            self.assertFalse(app.exception)
            tabela = app.dataframe[0].value
            self.assertEqual(tabela['Nome Completo'].tolist(), ['CARLA', 'BRUNO', 'ANA'])
            self.assertEqual(tabela['Poltrona'].tolist(), ['3', '2', '1'])

    def test_modelo_com_missao_importavel(self):
        from geracao import BASE, ler_planilha
        from test_geracao import excel
        colunas = pd.read_excel(BASE / 'modelo_planilha.xlsx').columns.tolist()
        self.assertEqual(colunas.count('Missão'), 1)
        linha = {c: '' for c in colunas}
        linha.update({'Nome Completo': 'Ana Silva', 'Missão': 'Responsável de Setor'})
        tabela = ler_planilha(excel([[linha[c] for c in colunas]], colunas))
        self.assertEqual(tabela.iloc[0]['Missão'], 'Responsável de Setor')
        self.assertEqual(tabela.iloc[0]['Nome Completo'], 'Ana Silva')

    def test_responsavel_setor_variantes_e_posicao(self):
        for missao in ['Responsável de Setor', 'RESPONSAVEL DE SETOR', 'Responsáveis de Setor', 'Responsável do Setor']:
            tabela = pd.DataFrame({'Nome Completo': ['ZÉ', 'ANA', 'CARLA'],
                                   'Missao': ['Assistente de Família', missao, 'Assistente de Ministro']})
            self.assertEqual(ordenar_por_missao(tabela)['Nome Completo'].tolist(), ['CARLA', 'ANA', 'ZÉ'])

    def test_sem_missao_prefixos_e_desconhecidos(self):
        tabela = pd.DataFrame({'Nome Completo': ['ZÉLIA', 'PROF. ANA', 'MIN. ZÉ', 'MIN. ÁLVARO', 'ANA']})
        self.assertEqual(ordenar_por_missao(tabela)['Nome Completo'].tolist(), ['MIN. ÁLVARO', 'MIN. ZÉ', 'PROF. ANA', 'ANA', 'ZÉLIA'])
        tabela['Missao'] = ['', '', '', '', 'Membro']
        self.assertEqual(ordenar_por_missao(tabela)['Nome Completo'].tolist(), ['MIN. ÁLVARO', 'MIN. ZÉ', 'PROF. ANA', 'ANA', 'ZÉLIA'])
