import unittest
import pandas as pd
from padronizacao import ordenar_por_missao
from conferencia import ordenar_participantes


class OrganizacaoTest(unittest.TestCase):
    def test_hierarquia_nome_acentos_e_dados_preservados(self):
        missoes = ['Primeira Vez', 'Frequentador', 'Membro', 'Assistente de Família',
                   'Assistentes de Ministros', 'Coordenador(a)', 'Professora', 'Ministro(a)']
        tabela = pd.DataFrame([{'Missão': m, 'Nome Completo': n, 'Poltrona': str(i),
                               '_ID_Inscricao': str(i), 'Gerar crachá': i % 2 == 0}
                              for i, (m, n) in enumerate((m, n) for m in missoes for n in ['ZÉLIA', 'ÁLVARO'])])
        ordenada = ordenar_por_missao(tabela)
        self.assertEqual(ordenada['Missão'].tolist(), [m for m in reversed(missoes) for _ in range(2)])
        self.assertEqual(ordenada['Nome Completo'].tolist(), ['ÁLVARO', 'ZÉLIA']*8)
        self.assertEqual(tabela.set_index('_ID_Inscricao').to_dict('index'), ordenada.set_index('_ID_Inscricao').to_dict('index'))
        self.assertEqual(ordenar_participantes(tabela, 'Missão e nome').to_dict('records'), ordenada.to_dict('records'))
        self.assertEqual(ordenar_por_missao(ordenada).to_dict('records'), ordenada.to_dict('records'))

    def test_sem_missao_prefixos_e_desconhecidos(self):
        tabela = pd.DataFrame({'Nome Completo': ['ZÉLIA', 'PROF. ANA', 'MIN. ZÉ', 'MIN. ÁLVARO', 'ANA']})
        self.assertEqual(ordenar_por_missao(tabela)['Nome Completo'].tolist(), ['MIN. ÁLVARO', 'MIN. ZÉ', 'PROF. ANA', 'ANA', 'ZÉLIA'])
        tabela['Missao'] = ['', '', '', '', 'Membro']
        self.assertEqual(ordenar_por_missao(tabela)['Nome Completo'].tolist(), ['MIN. ÁLVARO', 'MIN. ZÉ', 'PROF. ANA', 'ANA', 'ZÉLIA'])
