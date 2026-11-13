"""Integração optativa: executa o Excel real em Windows, sem mocks."""
import os
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase, skipUnless
from openpyxl import load_workbook
from automacao.executar_automacao import executar_automacao
from preparar_demo import preparar_demo

@skipUnless(os.name == 'nt' and os.getenv('EXECUTAR_TESTE_EXCEL') == '1',
            'Requer Windows, Excel e EXECUTAR_TESTE_EXCEL=1')
class ExcelRealTests(TestCase):
    def test_fluxo_reexecucao_e_reducao_com_recalculo(self):
        with TemporaryDirectory() as pasta:
            raiz = preparar_demo(Path(pasta))
            caminho = raiz/'dashboard'/'dashboard_exemplo.xlsx'
            for _ in range(2):
                self.assertTrue(executar_automacao(motor='win32com',_raiz_demo=raiz))
                wb = load_workbook(caminho,data_only=True)
                self.assertEqual([wb['RESUMO'].cell(r,2).value for r in range(5,9)], [6,8,5,4])
                self.assertEqual([wb['RESUMO'].cell(r,3).value for r in range(5,9)], [6300,10800,6000,6000])
                self.assertEqual(wb['VENDAS']['E8'].value, '000001')
                for r in range(8,16):
                    self.assertEqual([wb['RECEBIMENTOS'].cell(r,c).value for c in range(16,22)], ['Reservado']*6)
                wb.close()
            origem = raiz/'entrada'/'vendas_exemplo.xlsx'
            wb = load_workbook(origem)
            wb.active.delete_rows(6,4)
            wb.save(origem)
            wb.close()
            self.assertTrue(executar_automacao(motor='win32com',_raiz_demo=raiz))
            wb = load_workbook(caminho,data_only=True)
            self.assertEqual(wb['RESUMO']['B5'].value,2)
            self.assertEqual(wb['RESUMO']['C5'].value,900)
            self.assertEqual(wb['VENDAS']['C10'].value,'TOTAL')
            wb.close()
