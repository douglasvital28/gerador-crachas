"""Link do Gestor de Caravanas: importa a aba do ônibus sem mudar o fluxo manual."""
import unittest
from unittest.mock import patch
from streamlit.testing.v1 import AppTest
from geracao import BASE
from sheets_google import CAMPOS

PLANILHA = 'b' * 40
AUTH = {'auth': {'redirect_uri': 'https://example.test/oauth2callback',
                 'cookie_secret': 'apenas-teste', 'google': {'client_id': 'apenas-teste'}}}
CABECALHO = [*CAMPOS[:9], 'Missão', 'Trajeto', 'Observações', 'Modalidade', 'Responsável acompanhante', 'ID da inscrição']


class Usuario(dict):
    tokens = {'access': 'token-ficticio'}


def linha(nome, poltrona, ident):
    campos = {'Nome Completo': nome, 'Unidade': 'Central', 'Poltrona': poltrona, 'Trajeto': 'IDA E VOLTA'}
    return [*[campos.get(c, '') for c in CABECALHO[:-1]], ident]


def abrir(params, logado=True, abas=('Ônibus 1', 'Ônibus 2', 'Sem poltrona'), valores=None):
    valores = valores or [CABECALHO, linha('Ana Exemplo', '12', '101'), linha('Bruno Teste', '13', '102')]
    usuario = Usuario(is_logged_in=logado, sub='organizador-1', email='teste@example.test')
    pilha = [patch('streamlit.secrets', AUTH), patch('streamlit.user', usuario),
             patch('fonte_google.abas_planilha', return_value=list(abas)),
             patch('fonte_google.respostas_planilha', side_effect=lambda *a: valores)]
    mocks = [p.start() for p in pilha]
    app = AppTest.from_file(str(BASE / 'app_streamlit.py'))
    app.query_params.update(params)
    return app.run(timeout=30), mocks, pilha


class LinkGestorTest(unittest.TestCase):
    def tearDown(self):
        patch.stopall()

    def test_importa_aba_do_onibus_uma_vez(self):
        app, mocks, _ = abrir({'planilha': PLANILHA, 'aba': 'Ônibus 2'})
        self.assertFalse(app.exception)
        self.assertEqual(app.radio('origem_dados').value, 'Google Sheets')
        self.assertEqual(app.session_state['gs_link'], 'https://docs.google.com/spreadsheets/d/' + PLANILHA + '/edit')
        mocks[3].assert_called_once_with('token-ficticio', PLANILHA, 'Ônibus 2')
        trabalho = app.session_state['_gs_trabalho']
        self.assertEqual(trabalho['Nome Completo'].tolist(), ['ANA EXEMPLO', 'BRUNO TESTE'])
        self.assertEqual(trabalho['Poltrona'].tolist(), ['12', '13'])
        # ID da inscrição sugerido como identificador.
        identificador = next(s for s in app.selectbox if s.label == 'Identificador único da inscrição')
        self.assertEqual(identificador.value, len(CABECALHO) - 1)
        app.run()
        app.run()
        self.assertFalse(app.exception)
        self.assertEqual(mocks[2].call_count, 1)
        self.assertEqual(mocks[3].call_count, 1)
        # Depois da importação automática, os botões continuam funcionando como antes.
        app.button('gs_buscar').click().run()
        self.assertEqual(mocks[3].call_count, 2)
        app.button('gerar_crachas').click().run()
        self.assertFalse(app.exception)
        self.assertIn('pdf', app.session_state)

    def test_aba_inexistente_avisa_e_nao_importa(self):
        app, mocks, _ = abrir({'planilha': PLANILHA, 'aba': 'Ônibus 9'})
        self.assertFalse(app.exception)
        self.assertTrue(any('Ônibus 9' in w.value for w in app.warning))
        mocks[3].assert_not_called()
        self.assertNotIn('_gs_trabalho', app.session_state)

    def test_sem_login_orienta_a_voltar_ao_gestor(self):
        app, mocks, _ = abrir({'planilha': PLANILHA, 'aba': 'Ônibus 1'}, logado=False)
        self.assertFalse(app.exception)
        self.assertTrue(any('Gestor de Caravanas' in i.value for i in app.info))
        self.assertTrue(any(b.key == 'gs_login' for b in app.button))
        mocks[2].assert_not_called()

    def test_link_invalido_e_sem_link_mantem_excel(self):
        for params in [{}, {'planilha': 'https://evil.test/' + PLANILHA, 'aba': 'Ônibus 1'},
                       {'planilha': PLANILHA}, {'planilha': PLANILHA, 'aba': 'x' * 101}]:
            with self.subTest(params=params):
                app, mocks, pilha = abrir(params)
                self.assertFalse(app.exception)
                self.assertEqual(app.radio('origem_dados').value, 'Planilha Excel')
                mocks[2].assert_not_called()
                for p in pilha:
                    p.stop()

    def test_sem_link_fluxo_google_manual_inalterado(self):
        app, mocks, _ = abrir({})
        app.radio('origem_dados').set_value('Google Sheets').run()
        self.assertFalse(app.exception)
        self.assertEqual(app.text_input('gs_link').value, '')
        mocks[2].assert_not_called()
        self.assertNotIn('_gs_metadados', app.session_state)


if __name__ == '__main__':
    unittest.main()
