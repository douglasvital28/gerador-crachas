from hashlib import sha256
from io import BytesIO
import unittest
from unittest.mock import patch
from pypdf import PdfReader
from streamlit.testing.v1 import AppTest
from geracao import BASE, CONVENIO, CONTATO
from test_geracao import excel, pessoa


class SelecaoTest(unittest.TestCase):
    def test_duplicidades_avisam_sem_bloquear(self):
        dados = excel([pessoa(), pessoa()])
        with patch('streamlit.file_uploader', return_value=BytesIO(dados)):
            app = AppTest.from_file(str(BASE / 'app_streamlit.py')).run()
            self.assertFalse(app.exception)
            self.assertEqual(len(app.warning), 1)
            self.assertFalse(app.button[0].disabled)
            app.button[0].click().run()
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
            app.button[0].click().run()
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
