from io import BytesIO
from hashlib import sha256
import unittest
from unittest.mock import patch
from pypdf import PdfReader
from streamlit.testing.v1 import AppTest
from geracao import BASE, COLUNAS, ler_planilha, avisos_duplicidades
from recursos import exportar_planilha
from mapa_onibus import atribuir_poltrona
from mapa_pdf import gerar_mapa_pdf
from conferencia import gerar_lista
from trajetos import inferir_trajeto, preparar_trajetos, TRAJETOS, sentidos
from sheets_google import mapear_respostas, atualizar_respostas
from test_geracao import pessoa, excel


class TrajetosTest(unittest.TestCase):
    def test_vocabulario_ampliado_e_ambiguidades(self):
        exemplos = {
            'SOMENTE IDA': ['Não retornarei com a Caravana', 'Não voltarei com o grupo',
                           'Não vou retornar com a caravana', 'Não retorno', 'Não voltarei.',
                           'Voltarei de carro', 'Volto por conta própria', 'Retorno por meios próprios',
                           'Volto em outro ônibus', 'Retornarei com outra caravana', 'A volta será de Uber',
                           'Irei apenas', 'Somente irei', 'Apenas ida', 'Vou na caravana, volto de carro'],
            'SOMENTE VOLTA': ['Não irei com a caravana', 'Não vou com o grupo', 'Não vou.',
                             'Vou de carro', 'Irei por conta própria', 'Ida por meios próprios',
                             'Só retornarei', 'Apenas voltarei', 'Somente vou retornar',
                             'Vou com outra caravana', 'Irei de carro e voltarei com a caravana'],
            'IDA E VOLTA': ['Talvez não retornarei com a caravana', 'Minha mãe não voltará com a caravana',
                           'Ele só volta', 'Se chover, não volto', 'Não retornarei com a caravana?',
                           'Não voltarei de carro', 'Não irei de Uber', 'Não sei se vou',
                           'Não só vou como volto', 'Somente ida e volta', 'Vou e volto',
                           'Não vou e não volto', 'Ida de carro e volta de carro', 'Retorno com a caravana'],
        }
        for esperado, textos in exemplos.items():
            for texto in textos:
                with self.subTest(texto=texto):
                    self.assertEqual(inferir_trajeto(texto), esperado)

    def test_observacoes_e_escolha_explicita(self):
        for texto, esperado in [('Só vai', 'SOMENTE IDA'), ('somente ida com o grupo', 'SOMENTE IDA'),
                                ('só volta', 'SOMENTE VOLTA'), ('apenas retorno', 'SOMENTE VOLTA'),
                                ('vai e volta', 'IDA E VOLTA'), ('somente ida e volta', 'IDA E VOLTA'), ('só vai e volta', 'IDA E VOLTA'), ('só vai voltar', 'SOMENTE VOLTA'), ('não só vai como volta', 'IDA E VOLTA'), ('só vai e só volta', 'IDA E VOLTA'), ('', 'IDA E VOLTA')]:
            self.assertEqual(inferir_trajeto(texto), esperado)
        dados = excel([pessoa('Ana') + ['Só vai'], pessoa('Bruno') + ['Somente volta']], [*COLUNAS, 'Observacoes'])
        tabela = ler_planilha(dados)
        self.assertEqual(tabela['Trajeto'].tolist(), ['SOMENTE IDA', 'SOMENTE VOLTA'])
        self.assertFalse([a for a in avisos_duplicidades(tabela) if a['Campo'] == 'Poltrona'])
        tabela['Trajeto'] = 'IDA E VOLTA'
        self.assertEqual(ler_planilha(exportar_planilha(tabela))['Trajeto'].tolist(), ['IDA E VOLTA']*2)
        self.assertTrue([a for a in avisos_duplicidades(tabela) if a['Campo'] == 'Poltrona'])

    def test_combinacoes_de_ocupacao_por_sentido(self):
        tabela = ler_planilha(excel([pessoa('Ana'), pessoa('Bruno')]))
        tabela.loc[1, 'Poltrona'] = ''
        for primeiro in TRAJETOS:
            for segundo in TRAJETOS:
                tabela['Trajeto'] = [primeiro, segundo]
                with self.subTest(primeiro=primeiro, segundo=segundo):
                    if sentidos(tabela.iloc[0]) & sentidos(tabela.iloc[1]):
                        with self.assertRaises(ValueError):
                            atribuir_poltrona(tabela, 1, 1, 43)
                    else:
                        self.assertEqual(atribuir_poltrona(tabela, 1, 1, 43).iloc[1]['Poltrona'], '1')

    def test_mapa_compartilhamento_e_conflitos(self):
        tabela = ler_planilha(excel([pessoa('Ana') + ['somente ida'], pessoa('Bruno') + ['somente volta'],
                                   pessoa('Carla') + ['']], [*COLUNAS, 'Observações']))
        tabela.loc[1:, 'Poltrona'] = ''
        compartilhado = atribuir_poltrona(tabela, 1, 1, 43)
        self.assertEqual(compartilhado['Poltrona'].tolist(), ['001', '1', ''])
        with self.assertRaises(ValueError):
            atribuir_poltrona(compartilhado, 2, 1, 43)
        alterada = compartilhado.copy()
        alterada.loc[1, 'Trajeto'] = 'IDA E VOLTA'
        with self.assertRaises(ValueError):
            gerar_mapa_pdf(alterada, 43)
        liberada = atribuir_poltrona(compartilhado, 1, None, 43)
        self.assertEqual(liberada['Poltrona'].tolist(), ['001', '', ''])
        pdf = PdfReader(BytesIO(gerar_mapa_pdf(compartilhado.iloc[:2], 43)))
        self.assertEqual(len(pdf.pages), 4)
        self.assertIn('ANA', pdf.pages[0].extract_text())
        self.assertNotIn('BRUNO', pdf.pages[0].extract_text())
        self.assertIn('BRUNO', pdf.pages[2].extract_text())
        self.assertNotIn('ANA', pdf.pages[2].extract_text())
        texto = PdfReader(BytesIO(gerar_lista(compartilhado))).pages[0].extract_text()
        self.assertIn('SOMENTE IDA', texto)
        self.assertIn('SOMENTE VOLTA', texto)

    def test_interface_edita_trajeto_e_compartilha(self):
        linhas = [pessoa('Ana') + ['só vai'], pessoa('Bruno') + ['só volta']]
        linhas[1][2] = ''
        dados = excel(linhas, [*COLUNAS, 'Observações'])
        ident = sha256(dados).hexdigest()
        with patch('streamlit.file_uploader', return_value=BytesIO(dados)):
            app = AppTest.from_file(str(BASE / 'app_streamlit.py')).run()
            self.assertFalse(app.exception)
            self.assertFalse(app.button(f'assento_{ident}_43_1').disabled)
            app.button(f'assento_{ident}_43_1').click().run()
            self.assertFalse(app.exception)
            salvo = app.session_state['mapa_tabela_' + ident]
            self.assertEqual(salvo['Poltrona'].tolist(), ['001', '1'])
            self.assertTrue(any('ANA' in m.value and 'SOMENTE IDA' in m.value for m in app.markdown))
            self.assertTrue(any('BRUNO' in m.value and 'SOMENTE VOLTA' in m.value for m in app.markdown))
            revisao = app.session_state['mapa_revisao_' + ident]
            app.session_state[f'participantes_{ident}_r{revisao}'] = {
                'edited_rows': {1: {'Trajeto': 'IDA E VOLTA'}}, 'added_rows': [], 'deleted_rows': []}
            app.run()
            self.assertFalse(app.exception)
            self.assertTrue(any('conflito no mesmo sentido' in w.value for w in app.warning))

    def test_google_inferencia_e_atualizacao_preserva_manual(self):
        mapa = {'Nome Completo': 1, 'Observações': 2}
        antes = mapear_respostas([['ID', 'Nome', 'Obs'], ['1', 'Ana', 'só vai']], mapa, 0, 'viagem')
        self.assertEqual(antes.iloc[0]['Trajeto'], 'SOMENTE IDA')
        trabalho = antes.copy()
        trabalho.loc[0, 'Trajeto'] = 'SOMENTE VOLTA'
        depois = mapear_respostas([['ID', 'Nome', 'Obs'], ['1', 'Ana Silva', '']], mapa, 0, 'viagem')
        resultado, _, _ = atualizar_respostas(antes, trabalho, depois)
        self.assertEqual(resultado.iloc[0]['Trajeto'], 'SOMENTE VOLTA')
        self.assertEqual(resultado.iloc[0]['Nome Completo'], 'Ana Silva')
