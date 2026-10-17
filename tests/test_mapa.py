from hashlib import sha256
from io import BytesIO
import unittest
from unittest.mock import patch
from pypdf import PdfReader
from streamlit.testing.v1 import AppTest
from geracao import BASE, COLUNAS, ler_planilha
from mapa_onibus import MAPAS, atribuir_poltrona, ocupacao
from test_geracao import pessoa, excel


class MapaTest(unittest.TestCase):
    def test_modelos_conforme_referencias(self):
        for modelo, superior in [(43, 31), (46, 34)]:
            pisos = MAPAS[modelo]
            numeros = [[n for fila in filas for n in fila if isinstance(n, int)] for filas in pisos.values()]
            self.assertEqual(numeros[0], list(range(1, superior+1)))
            self.assertEqual(numeros[1], list(range(superior+1, modelo+1)))
        self.assertEqual(MAPAS[43]['Piso superior'][2], (6, 7, 'Escada'))
        self.assertEqual(MAPAS[46]['Piso superior'][3], (9, 10, 'Sem poltrona'))

    def test_importacao_sem_coluna_poltrona_e_conflitos(self):
        p = pessoa()
        df = ler_planilha(excel([p[:2]+p[3:]], [c for c in COLUNAS if c != 'Poltrona']))
        self.assertEqual(df.iloc[0]['Poltrona'], '')
        df = ler_planilha(excel([pessoa('A'), pessoa('B')]))
        self.assertEqual(ocupacao(df)[1], [0, 1])
        with self.assertRaises(ValueError):
            atribuir_poltrona(df, 1, 1, 43)
        movido = atribuir_poltrona(df, 1, 43, 43)
        self.assertEqual(movido.iloc[0]['Poltrona'], '001')
        self.assertEqual(movido.iloc[1]['Poltrona'], '43')
        self.assertEqual(df.iloc[1]['Poltrona'], '001')
        self.assertEqual(atribuir_poltrona(movido, 1, None, 43).iloc[1]['Poltrona'], '')
        with self.assertRaises(ValueError):
            atribuir_poltrona(df, 1, 46, 43)

    def test_cliques_preservam_edicoes_e_atualizam_pdf(self):
        linhas = [pessoa('Ana'), pessoa('Bruno'), pessoa('Carla')]
        linhas[1][2] = linhas[2][2] = ''
        dados = excel(linhas)
        ident = sha256(dados).hexdigest()
        with patch('streamlit.file_uploader', return_value=BytesIO(dados)):
            app = AppTest.from_file(str(BASE / 'app_streamlit.py')).run()
            self.assertFalse(app.exception)
            self.assertTrue(app.button(f'assento_{ident}_43_1').disabled)
            app.session_state['participantes_' + ident] = {
                'edited_rows': {1: {'Nome Completo': 'Bruno Corrigido', 'Gerar crachá': False}},
                'added_rows': [], 'deleted_rows': []}
            app.run()
            app.button(f'assento_{ident}_43_5').click().run()
            self.assertFalse(app.exception)
            salvo = app.session_state['mapa_tabela_' + ident]
            self.assertEqual(salvo.iloc[1]['Nome Completo'], 'Bruno Corrigido')
            self.assertEqual(salvo.iloc[1]['Poltrona'], '5')
            self.assertFalse(salvo.iloc[1]['Gerar crachá'])
            self.assertEqual(app.button(f'assento_{ident}_43_5').label, '05 ✓')
            self.assertEqual(app.selectbox('mapa_passageiro_' + ident).value, 1)
            app.run()
            self.assertEqual(app.selectbox('mapa_passageiro_' + ident).value, 1)
            app.session_state['participantes_' + ident + '_r1'] = {
                'edited_rows': {0: {'Unidade': 'Unidade Corrigida'}},
                'added_rows': [], 'deleted_rows': []}
            app.run()
            app.selectbox('mapa_passageiro_' + ident).set_value(2).run()
            app.selectbox('mapa_modelo_' + ident).set_value(46).run()
            app.button(f'assento_{ident}_46_46').click().run()
            self.assertFalse(app.exception)
            app.button[0].click().run()
            fields = PdfReader(BytesIO(app.session_state['pdf'])).get_fields()
            self.assertEqual(fields['IGREJA']['/V'], 'Unidade Corrigida')
            self.assertEqual(fields['POLTRONA 2']['/V'], '46')
            app.button('liberar_' + ident).click().run()
            self.assertFalse(app.exception)
            self.assertNotIn('pdf', app.session_state)
            self.assertEqual(app.session_state['mapa_tabela_' + ident].iloc[2]['Poltrona'], '')
