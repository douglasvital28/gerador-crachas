import unittest
from unittest.mock import patch
from streamlit.testing.v1 import AppTest
from geracao import BASE


class PaginasPublicasTest(unittest.TestCase):
    def test_documentos_publicos_sem_importacao_ou_login(self):
        for pagina, titulo in [('privacidade', 'Política de Privacidade'), ('termos', 'Termos de Uso')]:
            with self.subTest(pagina=pagina), patch('streamlit.file_uploader', side_effect=AssertionError('Página pública não deve importar dados')):
                app = AppTest.from_file(str(BASE / 'app_streamlit.py'))
                app.query_params['pagina'] = pagina
                app.run()
                self.assertFalse(app.exception)
                self.assertTrue(any('# ' + titulo in m.value for m in app.markdown))
                self.assertTrue(any('douglas.vital.28@gmail.com' in m.value for m in app.markdown))
                self.assertFalse(app.get('file_uploader'))
                self.assertFalse(app.button)

    def test_inicio_tem_links_e_fluxo_normal(self):
        app = AppTest.from_file(str(BASE / 'app_streamlit.py')).run()
        self.assertFalse(app.exception)
        self.assertTrue(any('?pagina=privacidade' in m.value and '?pagina=termos' in m.value for m in app.markdown))
        self.assertTrue(app.get('file_uploader'))
