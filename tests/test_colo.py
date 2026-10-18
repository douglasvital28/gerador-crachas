from io import BytesIO
import pandas as pd
import unittest
import pymupdf
from colo import preparar_colo
from geracao import gerar_pdf,avisos_duplicidades,ler_planilha
from mapa_onibus import ocupacao,atribuir_poltrona
from mapa_pdf import gerar_mapa_pdf
from recursos import exportar_planilha
from sheets_google import mapear_respostas,CAMPOS


def people():
    return pd.DataFrame([{'Nome Completo':'MIN. ANA SILVA','Unidade':'TESTE','Poltrona':'1','Nome do Plano':'','Telefone do Plano':'','Nome Contato de Emergência':'','Telefone Contato de Emergência':'','Modalidade':'COM ASSENTO','Responsável acompanhante':''},
        {'Nome Completo':'BEBÊ SILVA','Unidade':'TESTE','Poltrona':'1','Nome do Plano':'','Telefone do Plano':'','Nome Contato de Emergência':'','Telefone Contato de Emergência':'','Modalidade':'CRIANÇA DE COLO','Responsável acompanhante':'ANA SILVA'}])


class TestColo(unittest.TestCase):
    def test_colo_is_not_a_seat_and_badge_only_changes_seat_field(self):
        df=preparar_colo(people())
        assert df.iloc[1]['Poltrona']=='COLO' and ocupacao(df)=={1:[0]}
        assert not avisos_duplicidades(df)
        with self.assertRaisesRegex(ValueError, 'colo'):atribuir_poltrona(df,1,2,46)
        pdf=gerar_pdf(df)
        with pymupdf.open(stream=pdf,filetype='pdf') as doc:
            fields={w.field_name:w.field_value for w in doc[0].widgets()}
            assert 'COLO' in fields.values()
            assert 'Responsável acompanhante' not in fields.values()
        mapped=gerar_mapa_pdf(df,46)
        with pymupdf.open(stream=mapped,filetype='pdf') as doc:
            text=''.join(p.get_text() for p in doc)
            assert 'COLO' in text and 'BEBÊ SILVA' in text
        returned=ler_planilha(exportar_planilha(df))
        assert returned.iloc[1]['Poltrona']=='COLO'
