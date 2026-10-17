from io import BytesIO
import unittest
import pandas as pd
import pymupdf
from pypdf import PdfReader
from geracao import COLUNAS, ler_planilha, gerar_pdf


def excel(rows, columns=COLUNAS):
    stream = BytesIO()
    pd.DataFrame(rows, columns=columns).to_excel(stream, index=False)
    return stream.getvalue()


def pessoa(nome='Maria da Conceição Aparecida de Oliveira Santos'):
    return [nome, 'Unidade Central', '001', '', '0112345678', 'José Silva', '011999999999']


class GeracaoTest(unittest.TestCase):
    def test_validacao(self):
        for data in [b'invalido', excel([]), excel([['Fulano']], ['Outra coluna']),
                     excel([pessoa('')]), excel([pessoa()] * 401), b'0' * (5*1024*1024+1)]:
            with self.subTest(tamanho=len(data)), self.assertRaises(ValueError):
                ler_planilha(data)

    def test_linhas_vazias_e_texto(self):
        df = ler_planilha(excel([pessoa(), [''] * len(COLUNAS)]))
        self.assertEqual(len(df), 1)
        self.assertEqual(df.iloc[0]['Poltrona'], '001')
        self.assertEqual(df.iloc[0]['Telefone do Plano'], '0112345678')

    def test_pdf_1_4_5_pessoas(self):
        for count in [1, 4, 5]:
            with self.subTest(count=count):
                nomes = [f'Participante {i} José da Conceição' for i in range(count)]
                pdf = gerar_pdf(ler_planilha(excel([pessoa(n) for n in nomes])), horario='18:00')
                reader = PdfReader(BytesIO(pdf))
                self.assertEqual(len(reader.pages), (count + 3)//4)
                fields = reader.get_fields()
                valores = [f.get('/V') for f in fields.values()]
                for nome in nomes:
                    self.assertIn(nome, valores)
                with pymupdf.open(stream=pdf, filetype='pdf') as doc:
                    widgets = [w for page in doc for w in page.widgets()]
                    for w in widgets:
                        self.assertIn(w.field_name, fields)
                        if w.field_type == pymupdf.PDF_WIDGET_TYPE_TEXT:
                            self.assertEqual(w.field_value or '', fields[w.field_name].get('/V') or '')
                            self.assertIsNotNone(reader.pages[0].get('/Annots'))
                if count == 5:
                    from pathlib import Path
                    Path('tmp/pdfs/conferencia.pdf').write_bytes(pdf)
                    with pymupdf.open(stream=pdf, filetype='pdf') as doc:
                        doc[0].get_pixmap(matrix=pymupdf.Matrix(1.4, 1.4)).save('tmp/pdfs/conferencia.png')

    def test_texto_longo(self):
        with self.assertRaisesRegex(ValueError, 'Texto muito longo'):
            gerar_pdf(ler_planilha(excel([pessoa('Nome ' * 100)])))

if __name__ == '__main__':
    unittest.main()
