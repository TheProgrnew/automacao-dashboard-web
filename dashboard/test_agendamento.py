from unittest.mock import patch
from django.contrib.auth import get_user_model
from django.test import TestCase, override_settings
from django.urls import reverse
from django.core.management import call_command
from dashboard.models import ConfiguracaoAgendamento, HistoricoExecucao


class TestesAgendamento(TestCase):
    def setUp(self):
        User = get_user_model()
        self.admin = User.objects.create_user(username="admin_agenda", password="teste", is_staff=True)
        self.comum = User.objects.create_user(username="comum_agenda", password="teste")

    def test_desativado_por_padrao(self):
        self.assertFalse(ConfiguracaoAgendamento.objects.exists())

    def test_usuario_comum_nao_configura(self):
        self.client.force_login(self.comum)
        self.assertNotContains(self.client.get(reverse("inicio")), "Agendamento automático (administrador)")
        self.client.post(reverse("configurar_agendamento"), {"ativo": "on", "horario": "08:00", "dias_semana": ["0"]})
        self.assertFalse(ConfiguracaoAgendamento.objects.exists())

    def test_admin_salva_desativado(self):
        self.client.force_login(self.admin)
        response = self.client.post(reverse("configurar_agendamento"), {"horario": "09:30", "dias_semana": ["0", "4"]})
        self.assertEqual(response.status_code, 302)
        conf = ConfiguracaoAgendamento.objects.get(pk=1)
        self.assertFalse(conf.ativo)
        self.assertEqual(conf.dias_semana, [0, 4])

    def test_admin_nao_ativa_sem_dias(self):
        self.client.force_login(self.admin)
        response = self.client.post(reverse("configurar_agendamento"), {"ativo": "on", "horario": "08:00"})
        self.assertEqual(response.status_code, 400)

    def test_get_nao_altera_configuracao(self):
        self.client.force_login(self.admin)
        self.assertEqual(self.client.get(reverse("configurar_agendamento")).status_code, 405)

    @patch("dashboard.management.commands.executar_agendamento.bloquear_execucao")
    @patch("dashboard.management.commands.executar_agendamento.ConfiguracaoAutomacao.objects.filter")
    def test_sem_agendamento_nao_executa(self, config_filter, trava):
        call_command("executar_agendamento")
        trava.assert_not_called()
        config_filter.assert_not_called()
        self.assertFalse(HistoricoExecucao.objects.exists())
