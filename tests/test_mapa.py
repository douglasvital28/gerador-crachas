from hashlib import sha256
from io import BytesIO
import unittest
from unittest.mock import patch
from pypdf import PdfReader
from streamlit.testing.v1 import AppTest
from geracao import BASE, COLUNAS, ler_planilha
from mapa_onibus import MAPAS, atribuir_poltrona, ocupacao, nome_curto
from test_geracao import pessoa, excel


class MapaTest(unittest.TestCase):
    def test_revisao_opcional_e_pdf_do_mapa(self):
        from mapa_pdf import gerar_mapa_pdf
        linhas = [pessoa('Min. Ana de Souza'), pessoa('Bruno')]
        linhas[1][2] = '2'
        dados = excel(linhas)
        ident = sha256(dados).hexdigest()
        with patch('streamlit.file_uploader', return_value=BytesIO(dados)):
            app = AppTest.from_file(str(BASE / 'app_streamlit.py')).run()
            self.assertFalse(app.exception)
            self.assertFalse(any(b.key.startswith('assento_') for b in app.button))
            self.assertTrue(any(b.label == 'Baixar PDF do mapa do ônibus' for b in app.get('download_button')))
            app.radio('revisar_assentos_' + ident).set_value('Sim').run()
            self.assertTrue(app.button(f'assento_{ident}_43_1'))
            app.button('liberar_' + ident).click().run()
            self.assertFalse(app.exception)
            self.assertEqual(app.session_state['mapa_tabela_' + ident].iloc[0]['Poltrona'], '')
        tabela = ler_planilha(dados)
        for modelo in MAPAS:
            pdf = PdfReader(BytesIO(gerar_mapa_pdf(tabela, modelo)))
            self.assertEqual(len(pdf.pages), len(MAPAS[modelo]))
            texto = ''.join(p.extract_text() for p in pdf.pages)
            self.assertIn('MIN. ANA DE SOUZA', texto)
            self.assertIn('BRUNO', texto)
        tabela.iloc[1, tabela.columns.get_loc('Poltrona')] = ''
        pdf = PdfReader(BytesIO(gerar_mapa_pdf(tabela, 44)))
        self.assertEqual(len(pdf.pages), 2)
        self.assertIn('PASSAGEIROS SEM POLTRONA', pdf.pages[-1].extract_text())
        self.assertIn('BRUNO', pdf.pages[-1].extract_text())
        tabela.iloc[1, tabela.columns.get_loc('Poltrona')] = '1'
        with self.assertRaises(ValueError):
            gerar_mapa_pdf(tabela, 44)
        tabela.iloc[1, tabela.columns.get_loc('Poltrona')] = '45'
        with self.assertRaises(ValueError):
            gerar_mapa_pdf(tabela, 44)

    def test_ld_44_e_identificacao_dos_passageiros(self):
        filas = MAPAS[44]['Semi-leito']
        self.assertEqual(filas[:3], [(1, 2, 4, 3), (5, 6, 'Escada', ''), (9, 10, 8, 7)])
        self.assertEqual(filas[-2:], [(41, 42, 39, 40), (43, 44, 'Banheiro', '')])
        self.assertEqual(sorted(n for fila in filas for n in fila if isinstance(n, int)), list(range(1, 45)))
        linhas = [pessoa('Ana'), pessoa('Bruno')]
        linhas[1][2] = ''
        dados = excel(linhas)
        ident = sha256(dados).hexdigest()
        with patch('streamlit.file_uploader', return_value=BytesIO(dados)):
            app = AppTest.from_file(str(BASE / 'app_streamlit.py')).run()
            app.selectbox('mapa_modelo_' + ident).set_value(44).run()
            self.assertFalse(app.exception)
            opcoes = app.selectbox('mapa_passageiro_' + ident).options
            self.assertTrue(opcoes[0].startswith('🟢 '))
            self.assertTrue(opcoes[1].startswith('🟠 '))
            self.assertIn('✓ Poltrona 001', opcoes[0])
            self.assertIn('SEM POLTRONA', opcoes[1])
            self.assertTrue(app.button('liberar_' + ident).disabled)
            app.button(f'assento_{ident}_44_44').click().run()
            self.assertTrue(app.selectbox('mapa_passageiro_' + ident).options[1].startswith('🟢 '))
            self.assertIn('✓ Poltrona 44', app.selectbox('mapa_passageiro_' + ident).options[1])
            self.assertFalse(app.button('liberar_' + ident).disabled)
            app.button('liberar_' + ident).click().run()
            self.assertTrue(app.selectbox('mapa_passageiro_' + ident).options[1].startswith('🟠 '))
            self.assertIn('SEM POLTRONA', app.selectbox('mapa_passageiro_' + ident).options[1])
            self.assertEqual([t.label for t in app.tabs if t.label in ('Crachás', 'Lista de embarque')], ['Crachás', 'Lista de embarque'])

    def test_misto_56_conforme_imagem(self):
        superior = MAPAS[56]['Piso superior']
        inferior = MAPAS[56]['Piso inferior']
        self.assertEqual(superior[:4], [(1, 2, 4, 3), (5, 6, 8, 7),
                                       (9, 10, 'Escada', ''), (11, 12, 'Escada', '')])
        self.assertEqual(superior[-1], (41, 42, 44, 43))
        self.assertEqual(sorted(n for fila in superior for n in fila if isinstance(n, int)), list(range(1, 45)))
        self.assertEqual(inferior, [(45, 46, 47), (48, 49, 50), (51, 52, 53), (54, 55, 56)])

    def test_misto_atribuicao_nos_dois_pisos(self):
        linhas = [pessoa('Ana'), pessoa('Bruno')]
        linhas[0][2] = linhas[1][2] = ''
        dados = excel(linhas)
        ident = sha256(dados).hexdigest()
        with patch('streamlit.file_uploader', return_value=BytesIO(dados)):
            app = AppTest.from_file(str(BASE / 'app_streamlit.py')).run()
            app.selectbox('mapa_modelo_' + ident).set_value(56).run()
            app.button(f'assento_{ident}_56_44').click().run()
            self.assertFalse(app.exception)
            app.selectbox('mapa_passageiro_' + ident).set_value(1).run()
            app.button(f'assento_{ident}_56_56').click().run()
            self.assertFalse(app.exception)
            app.button('gerar_crachas').click().run()
            campos = PdfReader(BytesIO(app.session_state['pdf'])).get_fields()
            self.assertEqual(campos['POLTRONA']['/V'], '44')
            self.assertEqual(campos['POLTRONA 2']['/V'], '56')
            self.assertTrue(app.button(f'assento_{ident}_56_44').disabled)

    def test_nome_curto_preserva_titulo_sem_contar_como_nome(self):
        for nome in ['Min. Maria de Souza', 'Ministra Maria de Souza',
                     'Ministro Maria de Souza', 'Prof. Maria de Souza',
                     'Professora Maria de Souza', '  MIN.  Maria de Souza ',
                     'Min. Prof. Maria de Souza']:
            with self.subTest(nome=nome):
                prefixo = ' '.join(nome.split()[:-3])
                self.assertEqual(nome_curto(nome), prefixo + ' Maria Souza')
        self.assertEqual(nome_curto('Min. Maria'), 'Min. Maria')
        self.assertEqual(nome_curto('Min.'), 'Min.')
        self.assertEqual(nome_curto('Maria Min. Souza'), 'Maria Souza')

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
            self.assertTrue(any('class="ocupante">ANA' in m.value for m in app.markdown))
            app.session_state['participantes_' + ident] = {
                'edited_rows': {1: {'Nome Completo': 'Bruno Corrigido', 'Gerar crachá': False}},
                'added_rows': [], 'deleted_rows': []}
            app.run()
            app.button(f'assento_{ident}_43_5').click().run()
            self.assertFalse(app.exception)
            salvo = app.session_state['mapa_tabela_' + ident]
            self.assertEqual(salvo.iloc[1]['Nome Completo'], 'BRUNO CORRIGIDO')
            self.assertEqual(salvo.iloc[1]['Poltrona'], '5')
            self.assertFalse(salvo.iloc[1]['Gerar crachá'])
            self.assertEqual(app.button(f'assento_{ident}_43_5').label, '05 ✓')
            self.assertTrue(any('class="ocupante">BRUNO CORRIGIDO' in m.value for m in app.markdown))
            self.assertEqual(app.selectbox('mapa_passageiro_' + ident).value, 1)
            app.run()
            self.assertEqual(app.selectbox('mapa_passageiro_' + ident).value, 1)
            app.session_state['participantes_' + ident + '_r' + str(app.session_state['mapa_revisao_' + ident])] = {
                'edited_rows': {0: {'Unidade': 'Unidade Corrigida'}},
                'added_rows': [], 'deleted_rows': []}
            app.run()
            app.selectbox('mapa_passageiro_' + ident).set_value(2).run()
            app.selectbox('mapa_modelo_' + ident).set_value(46).run()
            app.button(f'assento_{ident}_46_46').click().run()
            self.assertFalse(app.exception)
            app.button('gerar_crachas').click().run()
            fields = PdfReader(BytesIO(app.session_state['pdf'])).get_fields()
            self.assertEqual(fields['IGREJA']['/V'], 'UNIDADE CORRIGIDA')
            self.assertEqual(fields['POLTRONA 2']['/V'], '46')
            app.button('liberar_' + ident).click().run()
            self.assertFalse(app.exception)
            self.assertNotIn('pdf', app.session_state)
            self.assertEqual(app.session_state['mapa_tabela_' + ident].iloc[2]['Poltrona'], '')
