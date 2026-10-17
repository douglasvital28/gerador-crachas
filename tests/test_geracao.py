from io import BytesIO
import unittest
import pandas as pd
import pymupdf
from pypdf import PdfReader
from pypdf.generic import NullObject
from geracao import COLUNAS, CONVENIO, CONTATO, ler_planilha, gerar_pdf, separar_telefone, avisos_duplicidades


def excel(rows, columns=COLUNAS):
    stream = BytesIO()
    pd.DataFrame(rows, columns=columns).to_excel(stream, index=False)
    return stream.getvalue()


def pessoa(nome='Maria da Conceição Aparecida de Oliveira Santos'):
    return [nome, 'Unidade Central', '001', '', '0112345678', 'José Silva', '011999999999']


class GeracaoTest(unittest.TestCase):
    def test_separacao_telefones(self):
        for valor, esperado in [
            ('(11) 91234-5678', ('11', '912345678')),
            ('1132345678', ('11', '32345678')),
            ('011912345678', ('11', '912345678')),
            ('+55 (21) 91234-5678', ('21', '912345678')),
            ('005521912345678', ('21', '912345678')),
            ('5521912345678', ('21', '912345678')),
            ('', ('', '')), ('91234-5678', ('', '91234-5678')),
            ('0800 123 4567', ('', '0800 123 4567')),
            ('0300 123 4567', ('', '0300 123 4567')),
            ('+1 212 555 0123', ('', '+1 212 555 0123')),
            ('11 3234-5678 ramal 2', ('', '11 3234-5678 ramal 2')),
        ]:
            with self.subTest(valor=valor):
                self.assertEqual(separar_telefone(valor), esperado)

    def test_duplicidades_e_linhas_originais(self):
        linhas = [pessoa('José  Silva'), [''] * len(COLUNAS), pessoa('JOSE SILVA'), pessoa('Outro')]
        linhas[0][2] = '001'
        linhas[2][2] = '1'
        linhas[3][2] = ''
        df = ler_planilha(excel(linhas))
        avisos = avisos_duplicidades(df)
        self.assertEqual(len(avisos), 2)
        self.assertTrue(all(a['Linhas da planilha'] == '2, 4' for a in avisos))
        self.assertEqual(len(df), 3)
        self.assertEqual(df.iloc[0]['Poltrona'], '001')
        self.assertEqual(avisos_duplicidades(ler_planilha(excel([pessoa()]))), [])

    def test_ddd_pdf_todas_posicoes(self):
        linhas = []
        for i in range(5):
            p = pessoa(f'Pessoa {i}')
            p[4], p[6] = '(21) 3234-5678', '+55 (11) 91234-5678'
            linhas.append(p)
        pdf = gerar_pdf(ler_planilha(excel(linhas)))
        with pymupdf.open(stream=pdf, filetype='pdf') as doc:
            for pagina, page in enumerate(doc):
                valores = [(w.field_name, w.field_value) for w in page.widgets()]
                for prefixo, esperado in [('DDD Convenio', '21'), ('DDD Contato', '11'),
                                         ('TELEFONE DO CONVÊNIO', '32345678'),
                                         ('TELEFONE DO CONTATO', '912345678')]:
                    campos = [(n, v) for n, v in valores if n.startswith(prefixo)]
                    self.assertEqual(len(campos), 4)
                    self.assertEqual(sum(v == esperado for _, v in campos), 4 if pagina == 0 else 1)
            from pathlib import Path
            Path('tmp/pdfs').mkdir(parents=True, exist_ok=True)
            Path('tmp/pdfs/telefones.pdf').write_bytes(pdf)
            doc[0].get_pixmap().save('tmp/pdfs/telefones.png')

    def test_opcoes_normalizadas_e_invalidas(self):
        df = ler_planilha(excel([pessoa() + [' nAO ', ' FAMILIAR ']],
                                [*COLUNAS, 'Possui Plano de Saúde?', CONTATO]))
        self.assertEqual(df.iloc[0][CONVENIO], 'Não')
        self.assertEqual(df.iloc[0][CONTATO], 'Familiar')
        for coluna, valor in [(CONVENIO, 'Talvez'), (CONTATO, 'Vizinho')]:
            with self.subTest(coluna=coluna), self.assertRaisesRegex(ValueError, 'linhas 2'):
                ler_planilha(excel([pessoa() + [valor]], [*COLUNAS, coluna]))

    def test_marcacoes_todas_posicoes_e_paginas(self):
        opcoes = [('Sim', 'Familiar'), ('Não', 'Amigo'), ('Sim', 'Amigo'),
                  ('Não', 'Familiar'), ('', ''), ('Não', 'Amigo'),
                  ('Sim', 'Familiar'), ('', ''), ('Sim', 'Amigo')]
        df = ler_planilha(excel([pessoa(f'Pessoa {i}') + list(v) for i, v in enumerate(opcoes)],
                               [*COLUNAS, CONVENIO, CONTATO]))
        pdf = gerar_pdf(df)
        reader = PdfReader(BytesIO(pdf))
        self.assertEqual(len(reader.pages), 3)
        for pagina, page in enumerate(reader.pages):
            grupos = {}
            for ref in page['/Annots']:
                widget = ref.get_object()
                if '/Parent' not in widget:
                    continue
                parent = widget['/Parent']
                if parent.get('/FT') != '/Btn':
                    continue
                grupos.setdefault(parent['/T'], []).append(widget)
                self.assertIsInstance(widget.get('/V'), (type(None), NullObject))
            for nome, widgets in grupos.items():
                posicao = int(nome.split()[2 if nome.startswith('Possui') else
                                         (3 if nome.startswith('Tipo do') else 2)]) - 1
                indice = pagina * 4 + posicao
                valores = opcoes[indice] if indice < len(opcoes) else ('', '')
                esperado = valores[0 if nome.startswith('Possui') else 1]
                if esperado == 'Não':
                    # O nome interno legado usa Latin-1; o leitor pode decodificá-lo
                    # de outra forma. Verificar a opção da direita ("não").
                    direita = max(widgets, key=lambda w: w['/Rect'][0])
                    esperado = next(str(k)[1:] for k in direita['/AP']['/N'] if k != '/Off')
                esperado = '/' + (esperado or 'Off')
                self.assertEqual(widgets[0]['/Parent'].get('/V'), esperado)
                ativos = [w for w in widgets if w.get('/AS') != '/Off']
                self.assertEqual(len(ativos), int(esperado != '/Off'))
                for w in ativos:
                    self.assertEqual(w['/AS'], esperado)
                    self.assertIn(esperado, w['/AP']['/N'])
        from pathlib import Path
        Path('tmp/pdfs').mkdir(parents=True, exist_ok=True)
        Path('tmp/pdfs/opcoes.pdf').write_bytes(pdf)
        with pymupdf.open(stream=pdf, filetype='pdf') as doc:
            for i in range(len(doc)):
                doc[i].get_pixmap().save(f'tmp/pdfs/opcoes-{i}.png')

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
                    Path('tmp/pdfs').mkdir(parents=True, exist_ok=True)
                    Path('tmp/pdfs/conferencia.pdf').write_bytes(pdf)
                    with pymupdf.open(stream=pdf, filetype='pdf') as doc:
                        doc[0].get_pixmap(matrix=pymupdf.Matrix(1.4, 1.4)).save('tmp/pdfs/conferencia.png')

    def test_texto_longo(self):
        with self.assertRaisesRegex(ValueError, 'Texto muito longo'):
            gerar_pdf(ler_planilha(excel([pessoa('Nome ' * 100)])))

if __name__ == '__main__':
    unittest.main()
