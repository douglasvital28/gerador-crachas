from io import BytesIO
import unittest
from unittest.mock import patch, MagicMock
import requests
from streamlit.testing.v1 import AppTest
from geracao import BASE
from sheets_google import id_planilha, mapear_respostas, atualizar_respostas, abas_planilha, respostas_planilha, CAMPO_ID, CAMPOS

PLANILHA = 'a' * 40


def importar(linhas, origem='conta-1/viagem-1'):
    return mapear_respostas([['Registro', 'Nome', 'Assento', 'Unidade'], *linhas],
                           {'Nome Completo': 1, 'Poltrona': 2, 'Unidade': 3}, 0, origem)


class GoogleTest(unittest.TestCase):
    def test_missao_google_e_atualizacao(self):
        mapa = {'Nome Completo': 1, 'Missão': 2}
        antes = mapear_respostas([['ID', 'Nome', 'Missao'], ['1', 'Ana Silva', 'Ministro(a)']], mapa, 0, 'viagem')
        self.assertEqual(antes.iloc[0]['Nome Completo'], 'MIN. Ana Silva')
        depois = mapear_respostas([['ID', 'Nome', 'Missao'], ['1', 'Ana Silva', 'Professora']], mapa, 0, 'viagem')
        resultado, _, _ = atualizar_respostas(antes, antes.copy(), depois)
        self.assertEqual(resultado.iloc[0]['Nome Completo'], 'PROF. Ana Silva')
        sem_missao = importar([['1', 'Ana Silva', '', 'Central']])
        self.assertEqual(sem_missao.iloc[0]['Nome Completo'], 'Ana Silva')

    def test_interface_importa_atualiza_e_preserva_poltrona(self):
        class Usuario(dict):
            tokens = {'access': 'token-ficticio'}
        usuario = Usuario(is_logged_in=True, sub='organizador-1', email='teste@example.test')
        def resposta(chave, nome):
            campos = {'Nome Completo': nome, 'Unidade': 'Central'}
            return [chave, *[campos.get(c, '') for c in CAMPOS]]
        valores = [['Carimbo de data/hora', *CAMPOS], resposta('r1', 'Ana')]
        auth = {'auth': {'redirect_uri': 'https://example.test/oauth2callback',
                         'cookie_secret': 'apenas-teste', 'google': {'client_id': 'apenas-teste'}}}
        with patch('streamlit.secrets', auth), patch('streamlit.user', usuario), \
             patch('fonte_google.abas_planilha', return_value=['Respostas']), \
             patch('fonte_google.respostas_planilha', side_effect=lambda *args: valores):
            app = AppTest.from_file(str(BASE / 'app_streamlit.py')).run()
            app.radio('origem_dados').set_value('Google Sheets (teste)').run()
            app.text_input('gs_link').set_value('https://docs.google.com/spreadsheets/d/' + PLANILHA + '/edit').run()
            app.button('gs_abas').click().run()
            app.button('gs_buscar').click().run()
            app.button('gs_importar').click().run()
            self.assertFalse(app.exception)
            ident = app.session_state['_gs_identidade']
            app.button(f'assento_{ident}_43_12').click().run()
            self.assertEqual(app.session_state['_gs_trabalho'].iloc[0]['Poltrona'], '12')
            valores.append(resposta('r2', 'Bruno'))
            valores[1][2] = 'Unidade Nova'
            app.button('gs_buscar').click().run()
            app.button('gs_importar').click().run()
            self.assertFalse(app.exception)
            trabalho = app.session_state['_gs_trabalho']
            self.assertEqual(len(trabalho), 2)
            self.assertEqual(trabalho.iloc[0]['Poltrona'], '12')
            self.assertEqual(trabalho.iloc[0]['Unidade'], 'UNIDADE NOVA')
            self.assertEqual(trabalho.iloc[1]['Nome Completo'], 'BRUNO')
            app.button('gerar_crachas').click().run()
            self.assertFalse(app.exception)
            self.assertIn('pdf', app.session_state)
            with patch('fonte_google.respostas_planilha', side_effect=ValueError('A autorização Google expirou.')):
                app.button('gs_buscar').click().run()
                self.assertFalse(app.exception)
                self.assertEqual(app.session_state['_gs_trabalho'].iloc[0]['Poltrona'], '12')
                self.assertIn('_gs_bytes', app.session_state)
            usuario['sub'] = 'organizador-2'
            app.run()
            self.assertFalse(app.exception)
            self.assertNotIn('_gs_bytes', app.session_state)
            self.assertNotIn('pdf', app.session_state)

    def test_links_privados_validos_e_urls_rejeitadas(self):
        self.assertEqual(id_planilha('https://docs.google.com/spreadsheets/d/' + PLANILHA + '/edit#gid=0'), PLANILHA)
        for url in ['https://evil.test/' + PLANILHA, 'http://docs.google.com/spreadsheets/d/'+PLANILHA,
                    'https://docs.google.com.evil.test/spreadsheets/d/'+PLANILHA,
                    'https://docs.google.com/forms/d/'+PLANILHA]:
            with self.subTest(url=url), self.assertRaises(ValueError):
                id_planilha(url)

    def test_mapeamento_ids_estaveis_duplicados_e_contas_isoladas(self):
        dados = [['r1', 'Ana', '01', 'Central'], ['r2', 'Bruno', '', 'Norte']]
        df = importar(dados)
        invertido = importar(dados[::-1])
        self.assertEqual(df[CAMPO_ID].tolist(), invertido[CAMPO_ID].tolist()[::-1])
        self.assertEqual(df.iloc[0]['Poltrona'], '01')
        outro = importar(dados, 'conta-2/viagem-1')
        self.assertFalse(set(df[CAMPO_ID]) & set(outro[CAMPO_ID]))
        for linhas in [[dados[0], dados[0]], [['', 'Ana', '', '']], [['r1', '', '', '']]]:
            with self.assertRaises(ValueError):
                importar(linhas)

    def test_atualizacao_preserva_poltrona_correcao_e_selecao(self):
        base = importar([['r1', 'Ana', '', 'Central'], ['r2', 'Bruno', '', 'Norte']])
        trabalho = base.copy()
        trabalho['Gerar crachá'] = [False, True]
        trabalho.loc[0, 'Poltrona'] = '12'
        trabalho.loc[0, 'Nome Completo'] = 'Ana Corrigida'
        novas = importar([['r3', 'Carla', '', 'Sul'], ['r1', 'Ana Origem', '', 'Unidade Nova']])
        final, adicionadas, removidas = atualizar_respostas(base, trabalho, novas)
        self.assertEqual((adicionadas, removidas, len(final)), (1, 1, 3))
        self.assertEqual(final.iloc[0]['Nome Completo'], 'Ana Corrigida')
        self.assertEqual(final.iloc[0]['Poltrona'], '12')
        self.assertEqual(final.iloc[0]['Unidade'], 'Unidade Nova')
        self.assertFalse(final.iloc[0]['Gerar crachá'])
        self.assertEqual(final.iloc[1]['Nome Completo'], 'Bruno')
        self.assertTrue(final.iloc[2]['Gerar crachá'])
        again, adicionadas, _ = atualizar_respostas(novas, final, novas)
        self.assertEqual(adicionadas, 0)
        self.assertEqual(again.to_dict('records'), final.to_dict('records'))

    def test_api_apenas_get_google_e_erros_sem_token(self):
        import json
        resposta = MagicMock()
        resposta.status_code = 200
        resposta.__enter__.return_value = resposta
        resposta.iter_content.return_value = [json.dumps({'sheets':[{'properties':{'title':'Respostas'}}]}).encode()]
        with patch('sheets_google.requests.get', return_value=resposta) as get:
            self.assertEqual(abas_planilha('token-ficticio', PLANILHA), ['Respostas'])
            self.assertTrue(get.call_args.args[0].startswith('https://sheets.googleapis.com/'))
            self.assertFalse(get.call_args.kwargs['allow_redirects'])
        for status in [401, 403, 404, 429, 500, 302]:
            resposta.status_code = status
            with patch('sheets_google.requests.get', return_value=resposta), self.assertRaises(ValueError) as erro:
                abas_planilha('token-ficticio', PLANILHA)
            self.assertNotIn('token-ficticio', str(erro.exception))
        with patch('sheets_google.requests.get', side_effect=requests.Timeout('token-ficticio')), self.assertRaises(ValueError) as erro:
            abas_planilha('token-ficticio', PLANILHA)
        self.assertNotIn('token-ficticio', str(erro.exception))

    def test_interface_sem_configuracao_preserva_excel(self):
        with patch('streamlit.secrets', {}):
            app = AppTest.from_file(str(BASE / 'app_streamlit.py')).run()
            app.radio('origem_dados').set_value('Google Sheets (teste)').run()
            self.assertFalse(app.exception)
            self.assertTrue(any('configurar' in i.value for i in app.info))
            self.assertFalse(any(b.key == 'gerar_crachas' for b in app.button))
            app.radio('origem_dados').set_value('Planilha Excel').run()
            self.assertFalse(app.exception)
            self.assertTrue(app.get('file_uploader'))
