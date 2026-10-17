from io import BytesIO
import unittest
from openpyxl import load_workbook
from pypdf import PdfReader
from PIL import Image
from geracao import CONVENIO, CONTATO, ler_planilha, gerar_pdf
from recursos import validar_caravana, exportar_planilha, renderizar_pagina
from test_geracao import excel, pessoa


class RecursosTest(unittest.TestCase):
    def test_validacao_caravana(self):
        for entrada, esperado in [
            (('', '', ''), ('', '', '')),
            (('8:05', '11', '91234-5678'), ('08:05', '11', '912345678')),
            (('23:59', '', '+55 (21) 3234-5678'), ('23:59', '21', '32345678')),
            (('00:00', '11', '(11) 91234-5678'), ('00:00', '11', '912345678')),
        ]:
            with self.subTest(entrada=entrada):
                self.assertEqual(validar_caravana(*entrada), esperado)
        for entrada in [('24:00', '', ''), ('12:60', '', ''), ('18h', '', ''),
                        ('', '1', ''), ('', '011', ''), ('', '11', ''),
                        ('', '', '912345678'), ('', '11', '123'),
                        ('', '11', '(21) 91234-5678'), ('', '', '0800 123 4567'),
                        ('', '11', '12345678 ramal 2')]:
            with self.subTest(entrada=entrada), self.assertRaises(ValueError):
                validar_caravana(*entrada)

    def test_exportacao_reimportavel_com_escolhas_e_zeros(self):
        df = ler_planilha(excel([pessoa()]))
        df.loc[0, 'Nome Completo'] = '=Nome literal'
        df.loc[0, CONVENIO] = 'Não'
        df.loc[0, CONTATO] = 'Amigo'
        resultado = exportar_planilha(df)
        wb = load_workbook(BytesIO(resultado))
        ws = wb.active
        for row in ws:
            for cell in row:
                self.assertNotEqual(cell.data_type, 'f')
        self.assertEqual(ws.freeze_panes, 'A2')
        self.assertEqual(len(ws.data_validations.dataValidation), 2)
        restaurado = ler_planilha(resultado)
        self.assertEqual(restaurado.to_dict('records'), df.to_dict('records'))
        self.assertEqual(restaurado.iloc[0]['Poltrona'], '001')
        self.assertEqual(restaurado.iloc[0]['Telefone do Plano'], '0112345678')
        fields = PdfReader(BytesIO(gerar_pdf(restaurado))).get_fields()
        self.assertEqual(fields['Tipo Contato 1']['/V'], '/Amigo')

    def test_previa_paginada(self):
        df = ler_planilha(excel([pessoa(f'Pessoa {i}') for i in range(5)]))
        pdf = gerar_pdf(df)
        imagens = [renderizar_pagina(pdf, i) for i in (0, 1)]
        self.assertNotEqual(imagens[0], imagens[1])
        for png in imagens:
            img = Image.open(BytesIO(png))
            self.assertEqual(img.format, 'PNG')
            self.assertLessEqual(img.width, 1100)
        for pagina in (-1, 2):
            with self.assertRaises(ValueError):
                renderizar_pagina(pdf, pagina)
