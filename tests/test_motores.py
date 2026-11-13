"""Seleção explícita e falhas de infraestrutura sem exigir Excel nos testes."""
import os
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from unittest.mock import MagicMock, patch
from automacao.executar_automacao import executar_automacao
from automacao.motor_win32com import executar_automacao as executar_com
from preparar_demo import preparar_demo

class SelecaoMotorTests(TestCase):
    def test_padrao_e_win32com(self):
        with patch.dict(os.environ, {}, clear=True), patch('automacao.motor_win32com.executar_automacao', return_value=True) as com:
            self.assertTrue(executar_automacao())
            com.assert_called_once()
    def test_escolha_explicita_sobrepoe_ambiente(self):
        with patch.dict(os.environ, {'AUTOMACAO_MOTOR':'win32com'}), patch('automacao.motor_openpyxl.executar_automacao', return_value=True) as local:
            self.assertTrue(executar_automacao(motor='openpyxl'))
            local.assert_called_once()
    def test_escolha_pelo_env(self):
        with patch.dict(os.environ, {'AUTOMACAO_MOTOR':'openpyxl'}), patch('automacao.motor_openpyxl.executar_automacao', return_value=True) as local:
            self.assertTrue(executar_automacao())
            local.assert_called_once()
    def test_motor_invalido_nao_executa(self):
        with patch('automacao.motor_openpyxl.executar_automacao') as local, patch('automacao.motor_win32com.executar_automacao') as com:
            self.assertFalse(executar_automacao(motor='inexistente'))
            local.assert_not_called()
            com.assert_not_called()
    def test_falha_com_nao_aciona_alternativa(self):
        with patch('automacao.motor_win32com.executar_automacao', return_value=False), patch('automacao.motor_openpyxl.executar_automacao') as local:
            self.assertFalse(executar_automacao(motor='win32com'))
            local.assert_not_called()

class RecursosComTests(TestCase):
    def setUp(self):
        self.tmp = TemporaryDirectory()
        self.raiz = preparar_demo(Path(self.tmp.name))
        self.dashboard = self.raiz/'dashboard'/'dashboard_exemplo.xlsx'
        self.original = self.dashboard.read_bytes()
        self.pycom, self.cliente = MagicMock(), MagicMock()
    def tearDown(self):
        self.tmp.cleanup()
    def verificar_original(self):
        self.assertEqual(self.dashboard.read_bytes(), self.original)
        self.assertEqual(list(self.dashboard.parent.glob('demo_excel_*')), [])
    def test_falha_dispatch_libera_com(self):
        self.cliente.DispatchEx.side_effect = RuntimeError('Excel não instalado')
        with patch('automacao.motor_win32com.carregar_com', return_value=(self.pycom,self.cliente)):
            self.assertFalse(executar_com(_raiz_demo=self.raiz))
        self.pycom.CoInitialize.assert_called_once()
        self.pycom.CoUninitialize.assert_called_once()
        self.verificar_original()
    def test_falha_open_fecha_excel_e_libera_com(self):
        excel = self.cliente.DispatchEx.return_value
        excel.Workbooks.Open.side_effect = RuntimeError('Falha ao abrir')
        with patch('automacao.motor_win32com.carregar_com', return_value=(self.pycom,self.cliente)):
            self.assertFalse(executar_com(_raiz_demo=self.raiz))
        excel.Quit.assert_called_once()
        self.pycom.CoUninitialize.assert_called_once()
        self.verificar_original()
    def test_workbook_somente_leitura_fecha_sem_salvar(self):
        excel = self.cliente.DispatchEx.return_value
        wb = excel.Workbooks.Open.return_value
        wb.ReadOnly = True
        with patch('automacao.motor_win32com.carregar_com', return_value=(self.pycom,self.cliente)):
            self.assertFalse(executar_com(_raiz_demo=self.raiz))
        wb.Close.assert_called_once_with(SaveChanges=False)
        wb.Save.assert_not_called()
        excel.Quit.assert_called_once()
        self.verificar_original()
    def test_falha_close_ainda_tenta_quit(self):
        excel = self.cliente.DispatchEx.return_value
        wb = excel.Workbooks.Open.return_value
        wb.ReadOnly = True
        wb.Close.side_effect = RuntimeError('Falha ao fechar')
        with patch('automacao.motor_win32com.carregar_com', return_value=(self.pycom,self.cliente)):
            self.assertFalse(executar_com(_raiz_demo=self.raiz))
        excel.Quit.assert_called_once()
        self.pycom.CoUninitialize.assert_called_once()
        self.verificar_original()
    def test_com_do_chamador_nao_e_desinicializado(self):
        self.cliente.DispatchEx.side_effect = RuntimeError('Excel não instalado')
        with patch('automacao.motor_win32com.carregar_com', return_value=(self.pycom,self.cliente)):
            self.assertFalse(executar_com(_raiz_demo=self.raiz,com_inicializado=True))
        self.pycom.CoInitialize.assert_not_called()
        self.pycom.CoUninitialize.assert_not_called()
        self.verificar_original()
