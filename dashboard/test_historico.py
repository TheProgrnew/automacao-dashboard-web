from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from django.contrib.auth import get_user_model
from django.test import TestCase, override_settings
from django.urls import reverse
from dashboard.models import ConfiguracaoAutomacao, HistoricoExecucao
from dashboard.bloqueio_execucao import bloquear_execucao, ExecucaoOcupada

@override_settings(AUTOMACAO_MODO_TESTE=False)
class TesteHistoricoExecucao(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user(username="operador_historico", password="teste", is_staff=True)
        self.client.force_login(self.user)
        self.tmp = TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        base = Path(self.tmp.name)
        dashboard = base / "dashboard.xlsx"
        dashboard.touch()
        entrada = base / "entrada"
        saida = base / "saida"
        entrada.mkdir()
        saida.mkdir()
        ConfiguracaoAutomacao.objects.create(caminho_dashboard=str(dashboard), pasta_entrada=str(entrada), pasta_saida=str(saida))

    @patch("dashboard.views.executar_automacao", return_value=True)
    def test_sucesso_registra_historico(self, automacao):
        resposta = self.client.post(reverse("atualizar_dashboard"))
        self.assertEqual(resposta.status_code, 200)
        registro = HistoricoExecucao.objects.get()
        self.assertEqual(registro.status, "sucesso")
        self.assertEqual(registro.usuario, self.user)
        self.assertIsNotNone(registro.terminado_em)
        self.assertContains(self.client.get(reverse("inicio")), "Histórico das últimas execuções")

    @patch("dashboard.views.executar_automacao", return_value=False)
    def test_falha_retorno_false_registra(self, automacao):
        self.assertEqual(self.client.post(reverse("atualizar_dashboard")).status_code, 500)
        self.assertEqual(HistoricoExecucao.objects.get().status, "falha")

    @patch("dashboard.views.executar_automacao", side_effect=RuntimeError("erro simulado"))
    def test_excecao_registra_falha(self, automacao):
        self.assertEqual(self.client.post(reverse("atualizar_dashboard")).status_code, 500)
        self.assertEqual(HistoricoExecucao.objects.get().status, "falha")
        self.assertIsNotNone(HistoricoExecucao.objects.get().terminado_em)

    @patch("dashboard.views.bloquear_execucao", side_effect=ExecucaoOcupada())
    @patch("dashboard.views.executar_automacao")
    def test_concorrencia_bloqueada(self, automacao, trava):
        self.assertEqual(self.client.post(reverse("atualizar_dashboard")).status_code, 409)
        automacao.assert_not_called()
        self.assertFalse(HistoricoExecucao.objects.exists())

    def test_bloqueio_real_impede_segunda_entrada(self):
        with bloquear_execucao():
            with self.assertRaises(ExecucaoOcupada):
                with bloquear_execucao():
                    pass
