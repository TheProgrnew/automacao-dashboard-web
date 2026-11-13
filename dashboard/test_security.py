from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from django.contrib.auth import get_user_model
from django.test import TestCase, override_settings
from django.urls import reverse

class TesteSeguranca(TestCase):
    def setUp(self):
        User = get_user_model()
        self.comum = User.objects.create_user(username="comum", password="teste-forte-123")
        self.operador = User.objects.create_user(username="operador", password="teste-forte-123", is_staff=True)

    def test_anonimo_nao_acessa_pagina(self):
        self.assertEqual(self.client.get(reverse("inicio")).status_code, 302)

    def test_anonimo_nao_executa(self):
        self.assertEqual(self.client.post(reverse("atualizar_dashboard")).status_code, 302)

    def test_usuario_comum_nao_executa(self):
        self.client.force_login(self.comum)
        self.assertEqual(self.client.post(reverse("atualizar_dashboard")).status_code, 302)

    def test_usuario_comum_nao_configura(self):
        self.client.force_login(self.comum)
        self.assertEqual(self.client.post(reverse("configurar_automacao")).status_code, 302)

    @override_settings(AUTOMACAO_MODO_TESTE=True)
    @patch("dashboard.views.executar_automacao")
    def test_dashboard_binario_nao_dispara_automacao(self, executar):
        self.client.force_login(self.operador)
        with TemporaryDirectory() as tmp:
            base = Path(tmp)
            (base / "dashboard").mkdir()
            (base / "entrada").mkdir()
            (base / "saida").mkdir()
            (base / "dashboard" / "demonstrativo.xlsb").touch()
            with override_settings(AUTOMACAO_PASTA_TESTE=base):
                resposta = self.client.post(reverse("atualizar_dashboard"))
            self.assertEqual(resposta.status_code, 400)
            executar.assert_not_called()
