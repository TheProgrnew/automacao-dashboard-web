import hashlib
from pathlib import Path
import tempfile
import unittest
from openpyxl import load_workbook
from automacao.executar_automacao import executar_automacao
from automacao.config import ARQUIVOS
from preparar_demo import preparar_demo

class FluxoDemoTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory()
        self.raiz=preparar_demo(Path(self.tmp.name))
        self.dashboard=self.raiz/'dashboard'/'dashboard_exemplo.xlsx'
    def tearDown(self):
        self.tmp.cleanup()
    def executar(self,**kwargs):
        return executar_automacao(_raiz_demo=self.raiz,motor="openpyxl",**kwargs)
    def hash(self):
        return hashlib.sha256(self.dashboard.read_bytes()).hexdigest()
    def test_fluxo_completo_e_reexecucao(self):
        progresso=[]
        self.assertTrue(self.executar(progresso_callback=lambda n,t:progresso.append(n)))
        self.assertEqual(progresso[-1],100)
        wb=load_workbook(self.dashboard)
        expected={'VENDAS':6,'RECEBIMENTOS':8,'CUSTOS':5,'DESPESAS':4}
        for nome,spec in ARQUIVOS.items():
            self.assertEqual(wb[nome].cell(spec['linha'],spec['destinos'][0]).value,'DEMO-0001')
            self.assertEqual(wb[nome].cell(spec['linha']+expected[nome]-1,spec['destinos'][0]).value,f'DEMO-{expected[nome]:04d}')
        self.assertEqual(wb['VENDAS']['E8'].value,'000001')
        self.assertEqual(wb['VENDAS']['X8'].value,300)
        self.assertEqual(wb['VENDAS']['X14'].value,'=SUM(X8:X13)')
        self.assertEqual(wb['VENDAS']['X5'].value,'=SUBTOTAL(9,X8:X13)')
        for r in range(8,16):
            self.assertEqual([wb['RECEBIMENTOS'].cell(r,c).value for c in range(16,22)],['Reservado']*6)
        wb.close()
        self.assertTrue(self.executar())
        wb=load_workbook(self.dashboard)
        self.assertFalse(any(wb['VENDAS'].cell(r,3).value for r in range(15,wb['VENDAS'].max_row+1)))
        self.assertFalse(any(wb['RECEBIMENTOS'].cell(r,1).value for r in range(17,wb['RECEBIMENTOS'].max_row+1)))
        self.assertEqual(wb['RESUMO']['B5'].value,6)
        wb.close()
    def test_arquivo_ausente_nao_modifica_dashboard(self):
        original=self.hash()
        (self.raiz/'entrada'/'custos_exemplo.xlsx').unlink()
        self.assertFalse(self.executar())
        self.assertEqual(self.hash(),original)
    def test_valor_invalido_nao_modifica_dashboard(self):
        arquivo=self.raiz/'entrada'/'vendas_exemplo.xlsx'
        wb=load_workbook(arquivo);wb.active['N4']='erro';wb.save(arquivo);wb.close()
        original=self.hash()
        self.assertFalse(self.executar())
        self.assertEqual(self.hash(),original)
    def test_encolhimento_remove_registros_antigos(self):
        self.assertTrue(self.executar())
        arquivo=self.raiz/'entrada'/'vendas_exemplo.xlsx'
        wb=load_workbook(arquivo);wb.active.delete_rows(6,4);wb.save(arquivo);wb.close()
        self.assertTrue(self.executar())
        wb=load_workbook(self.dashboard)
        self.assertEqual(wb['VENDAS']['C10'].value,'TOTAL')
        self.assertFalse(any(wb['VENDAS'].cell(r,3).value for r in range(11,wb['VENDAS'].max_row+1)))
        self.assertEqual(wb['VENDAS']['X5'].value,'=SUBTOTAL(9,X8:X9)')
        wb.close()
    def test_destino_externo_bloqueado(self):
        original=self.hash()
        self.assertFalse(self.executar(caminho_dashboard=Path(self.tmp.name)/'outro.xlsx'))
        self.assertEqual(self.hash(),original)
    def test_link_simbolico_bloqueado(self):
        arquivo=self.raiz/'entrada'/'vendas_exemplo.xlsx'
        externo=self.raiz/'fora.xlsx';arquivo.rename(externo)
        try:
            arquivo.symlink_to(externo)
        except OSError:
            self.skipTest('Sistema sem permissão para symlinks')
        original=self.hash()
        self.assertFalse(self.executar())
        self.assertEqual(self.hash(),original)

    def test_modo_configurado_com_dados_ficticios(self):
        self.assertTrue(self.executar(modo_teste=False))
        wb=load_workbook(self.dashboard)
        self.assertEqual(wb['RESUMO']['B5'].value,6)
        wb.close()

    def test_planilha_real_extra_bloqueia_preparo_e_execucao(self):
        extra=self.raiz/'dashboard'/'planilha_antiga.xlsb'
        extra.write_bytes(b'arquivo de teste, sem dados reais')
        original=self.hash()
        with self.assertRaisesRegex(ValueError,'Planilhas extras'):
            preparar_demo(self.raiz,restaurar=True)
        self.assertFalse(self.executar())
        self.assertEqual(self.hash(),original)
        self.assertTrue(extra.exists())

    def test_relatorios_de_exemplo_tem_valores_coerentes(self):
        for nome,spec in ARQUIVOS.items():
            wb=load_workbook(Path(__file__).resolve().parent.parent/'dados_demo'/spec['arquivo'],data_only=True)
            ws=wb.active
            for row in ws.iter_rows(min_row=4,values_only=True):
                if row[0] is None or row[0]=='TOTAL':
                    continue
                if nome=='VENDAS':
                    self.assertEqual(row[15],row[13]-row[14])
                    self.assertEqual(row[21],row[15]-row[20])
                elif nome=='RECEBIMENTOS':
                    self.assertEqual(row[26],row[21]-row[22])
                    self.assertEqual(row[27],row[22]-row[23])
                elif nome=='CUSTOS':
                    self.assertEqual(row[14],row[12]-row[13])
                    self.assertEqual(row[15],row[11]-row[12])
                elif nome=='DESPESAS':
                    self.assertEqual(row[14],row[11]-row[13])
            wb.close()
