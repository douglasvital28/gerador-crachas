from hashlib import sha256
from io import BytesIO
import unittest
from unittest.mock import patch
from pypdf import PdfReader
from streamlit.testing.v1 import AppTest
from geracao import BASE, CONVENIO, CONTATO
from test_geracao import excel, pessoa
from recursos import exportar_planilha
from geracao import ler_planilha


class SelecaoTest(unittest.TestCase):
    def test_correcao_selecao_e_exportacao_completa(self):
        dados = excel([pessoa(f'Pessoa {i}') for i in range(5)])
        with patch('streamlit.file_uploader', return_value=BytesIO(dados)), \
             patch('recursos.exportar_planilha', wraps=exportar_planilha) as exportar:
            app = AppTest.from_file(str(BASE / 'app_streamlit.py')).run()
            editor = 'participantes_' + sha256(dados).hexdigest()
            alteracoes = {i: {'Gerar crachá': False} for i in range(4)}
            alteracoes[4] = {'Nome Completo': 'Nome Corrigido', 'Poltrona': '009',
                             'Telefone Contato de Emergência': '(21) 91234-5678'}
            app.session_state[editor] = {'edited_rows': alteracoes, 'added_rows': [], 'deleted_rows': []}
            app.run()
            self.assertFalse(app.exception)
            planilha = ler_planilha(exportar_planilha(exportar.call_args.args[0]))
            self.assertEqual(len(planilha), 5)
            self.assertNotIn('Gerar crachá', planilha.columns)
            self.assertEqual(planilha.iloc[4]['Poltrona'], '009')
            app.button('gerar_crachas').click().run()
            self.assertFalse(app.exception)
            reader = PdfReader(BytesIO(app.session_state['pdf']))
            self.assertEqual(len(reader.pages), 1)
            fields = reader.get_fields()
            self.assertEqual(fields['NOME']['/V'], 'Nome Corrigido')
            self.assertEqual(fields['POLTRONA']['/V'], '009')
            self.assertEqual(fields['DDD Contato']['/V'], '21')
            self.assertFalse(fields['NOME 2'].get('/V'))
            alteracoes[4]['Gerar crachá'] = False
            app.session_state[editor] = {'edited_rows': alteracoes, 'added_rows': [], 'deleted_rows': []}
            app.run()
            self.assertFalse(app.exception)
            self.assertTrue(app.button('gerar_crachas').disabled)
            self.assertNotIn('pdf', app.session_state)
            self.assertTrue(app.info)

    def test_nome_vazio_na_edicao_bloqueia_geracao(self):
        dados = excel([pessoa()])
        with patch('streamlit.file_uploader', return_value=BytesIO(dados)):
            app = AppTest.from_file(str(BASE / 'app_streamlit.py')).run()
            editor = 'participantes_' + sha256(dados).hexdigest()
            app.session_state[editor] = {'edited_rows': {0: {'Nome Completo': '  '}},
                                         'added_rows': [], 'deleted_rows': []}
            app.run()
            self.assertFalse(app.exception)
            self.assertTrue(app.button('gerar_crachas').disabled)
            self.assertIn('Nome Completo', app.error[0].value)

    def test_previa_navegavel_e_validacao(self):
        linhas = [pessoa(f'Pessoa {i}') for i in range(5)]
        for i, linha in enumerate(linhas):
            linha[2] = str(i + 1)
        dados = excel(linhas)
        with patch('streamlit.file_uploader', return_value=BytesIO(dados)):
            app = AppTest.from_file(str(BASE / 'app_streamlit.py')).run()
            app.text_input[0].set_value('25:00').run()
            self.assertTrue(app.button('gerar_crachas').disabled)
            self.assertTrue(app.error)
            app.text_input[0].set_value('8:05')
            app.text_input[4].set_value('(11) 91234-5678').run()
            self.assertFalse(app.button('gerar_crachas').disabled)
            app.button('gerar_crachas').click().run()
            self.assertFalse(app.exception)
            self.assertEqual(app.number_input[0].value, 1)
            fields = PdfReader(BytesIO(app.session_state['pdf'])).get_fields()
            self.assertEqual(fields['HORÁRIO DE RETORNO']['/V'], '08:05')
            self.assertEqual(fields['DDD DO RESP']['/V'], '11')
            self.assertEqual(fields['TELEFONE DO RESPONSÁVEL']['/V'], '912345678')
            app.number_input[0].set_value(2).run()
            self.assertFalse(app.exception)
            self.assertEqual(app.number_input[0].value, 2)
            self.assertIn('pdf', app.session_state)
            self.assertFalse(app.warning)
            app.text_input[0].set_value('9:00').run()
            self.assertNotIn('pdf', app.session_state)
            app.button('gerar_crachas').click().run()
            self.assertEqual(app.number_input[0].value, 1)

    def test_duplicidades_avisam_sem_bloquear(self):
        dados = excel([pessoa(), pessoa()])
        with patch('streamlit.file_uploader', return_value=BytesIO(dados)):
            app = AppTest.from_file(str(BASE / 'app_streamlit.py')).run()
            self.assertFalse(app.exception)
            self.assertTrue(any('nomes ou poltronas repetidos' in w.value for w in app.warning))
            self.assertFalse(app.button('gerar_crachas').disabled)
            app.button('gerar_crachas').click().run()
            self.assertFalse(app.exception)
            self.assertIn('pdf', app.session_state)

    def test_editar_gerar_e_invalidar_pdf(self):
        dados = excel([pessoa()])
        with patch('streamlit.file_uploader', return_value=BytesIO(dados)):
            app = AppTest.from_file(str(BASE / 'app_streamlit.py')).run()
            self.assertFalse(app.exception)
            editor = 'participantes_' + sha256(dados).hexdigest()
            app.session_state[editor] = {
                'edited_rows': {0: {CONVENIO: 'Sim', CONTATO: 'Familiar'}},
                'added_rows': [], 'deleted_rows': [],
            }
            app.run()
            app.button('gerar_crachas').click().run()
            self.assertFalse(app.exception)
            fields = PdfReader(BytesIO(app.session_state['pdf'])).get_fields()
            self.assertEqual(fields['Possui Convenio 1']['/V'], '/Sim')
            self.assertEqual(fields['Tipo Contato 1']['/V'], '/Familiar')
            app.run()
            self.assertIn('pdf', app.session_state)
            app.session_state[editor] = {
                'edited_rows': {0: {CONVENIO: 'Não', CONTATO: 'Amigo'}},
                'added_rows': [], 'deleted_rows': [],
            }
            app.run()
            self.assertNotIn('pdf', app.session_state)
            self.assertFalse(app.exception)
