from io import BytesIO
from datetime import date
from pathlib import Path
import unittest
from unittest.mock import patch
import pymupdf
from pypdf import PdfReader
from streamlit.testing.v1 import AppTest
from conferencia import gerar_lista, ordenar_participantes
from geracao import BASE, ler_planilha
from test_geracao import excel, pessoa
from hashlib import sha256


class ConferenciaTest(unittest.TestCase):
    def test_limite_400_e_campos_para_preenchimento_manual(self):
        df = ler_planilha(excel([pessoa(f'Passageiro {i:03d}') for i in range(1, 401)]))
        pdf = gerar_lista(df)
        texto = '\n'.join(p.extract_text() for p in PdfReader(BytesIO(pdf)).pages)
        self.assertIn('Passageiro 400', texto)
        self.assertIn('____/____/________', texto)
        self.assertIn('400', texto.split('Passageiro 400')[0][-10:])
        for i in range(1, 401):
            self.assertEqual(texto.count(f'Passageiro {i:03d}'), 1)

    def test_ordenacao_nome_e_poltrona_sem_modificar_dados(self):
        linhas = [pessoa('Zélia'), pessoa('Álvaro'), pessoa('Bruna'), pessoa('José')]
        for p, poltrona in zip(linhas, ['10', '02', '', '1']):
            p[2] = poltrona
        df = ler_planilha(excel(linhas))
        self.assertEqual(ordenar_participantes(df, 'Nome')['Nome Completo'].tolist(), ['Álvaro', 'Bruna', 'José', 'Zélia'])
        self.assertEqual(ordenar_participantes(df, 'Poltrona')['Poltrona'].tolist(), ['1', '02', '10', ''])
        self.assertEqual(df.iloc[0]['Nome Completo'], 'Zélia')
        with self.assertRaises(ValueError):
            gerar_lista(df.iloc[:0])

    def test_pdf_multipagina_e_nomes_longos(self):
        linhas = [pessoa(f'Pessoa {i:03d} da Conceição') for i in range(48)]
        linhas[0][0] = 'Maria da Conceição Aparecida dos Santos de Oliveira e Albuquerque'
        linhas[0][1] = 'Unidade Central de Atendimento de Guarapiranga'
        for i, p in enumerate(linhas):
            p[2] = str(i + 1)
        pdf = gerar_lista(ler_planilha(excel(linhas)), 'Poltrona', date(2026, 11, 1), 'José da Silva')
        reader = PdfReader(BytesIO(pdf))
        self.assertGreater(len(reader.pages), 1)
        textos = [p.extract_text() for p in reader.pages]
        for texto in textos:
            for esperado in ['Lista de conferência', '01/11/2026', '48', 'José da Silva', 'Observações', 'Retorno']:
                self.assertIn(esperado, texto)
        texto = '\n'.join(textos)
        for i in range(1, 48):
            self.assertEqual(texto.count(f'Pessoa {i:03d} da Conceição'), 1)
        self.assertIn('Albuquerque', texto)
        Path('tmp/pdfs').mkdir(parents=True, exist_ok=True)
        Path('tmp/pdfs/lista-conferencia.pdf').write_bytes(pdf)
        with pymupdf.open(stream=pdf) as doc:
            for i, p in enumerate(doc):
                self.assertAlmostEqual(p.rect.width, 595.276, places=1)
                p.get_pixmap(matrix=pymupdf.Matrix(1.2, 1.2)).save(f'tmp/pdfs/lista-{i}.png')
                # Conteúdo deve permanecer dentro das margens da página.
                for x0, y0, x1, y1, *_ in p.get_text('blocks'):
                    self.assertGreaterEqual(x0, 30)
                    self.assertLessEqual(x1, p.rect.width - 30)
                    self.assertLess(y1, p.rect.height - 10)

    def test_interface_lista_independente_e_invalidacao(self):
        dados = excel([pessoa('Zélia'), pessoa('Álvaro')])
        with patch('streamlit.file_uploader', return_value=BytesIO(dados)):
            app = AppTest.from_file(str(BASE / 'app_streamlit.py')).run()
            # A lista não depende do formato de telefone/horário usado no crachá.
            app.text_input[0].set_value('25:00').run()
            app.button[1].click().run()
            self.assertFalse(app.exception)
            self.assertIn('lista_pdf', app.session_state)
            self.assertNotIn('pdf', app.session_state)
            editor = 'participantes_' + sha256(dados).hexdigest()
            app.session_state[editor] = {'edited_rows': {0: {'Gerar crachá': False}}, 'added_rows': [], 'deleted_rows': []}
            app.selectbox[0].set_value('Apenas os selecionados').run()
            self.assertNotIn('lista_pdf', app.session_state)
            app.button[1].click().run()
            texto = PdfReader(BytesIO(app.session_state['lista_pdf'])).pages[0].extract_text()
            self.assertIn('Álvaro', texto)
            self.assertNotIn('Zélia', texto)
            app.selectbox[1].set_value('Poltrona').run()
            self.assertNotIn('lista_pdf', app.session_state)
            app.session_state[editor] = {'edited_rows': {0: {'Gerar crachá': False}, 1: {'Gerar crachá': False}}, 'added_rows': [], 'deleted_rows': []}
            app.run()
            self.assertTrue(app.button[1].disabled)
            self.assertFalse(app.exception)
