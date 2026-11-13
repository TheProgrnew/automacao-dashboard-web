from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse
from dashboard.models import ConfiguracaoAgendamento, HistoricoExecucao

class PermissoesEProgresso(TestCase):
    def setUp(self):
        U = get_user_model()
        self.admin = U.objects.create_user(username="gestor", password="senha", is_staff=True)
        self.comum = U.objects.create_user(username="leitor", password="senha")

    def test_agendamento_oculto_sem_liberacao(self):
        self.client.force_login(self.comum)
        self.assertNotContains(self.client.get(reverse("inicio")), "Agendamento autorizado")

    def test_apenas_admin_libera(self):
        self.client.force_login(self.comum)
        self.client.post(reverse("liberar_agendamento"), {"liberar":"sim"})
        self.assertFalse(ConfiguracaoAgendamento.objects.exists())
        self.client.force_login(self.admin)
        self.client.post(reverse("liberar_agendamento"), {"liberar":"sim"})
        self.assertTrue(ConfiguracaoAgendamento.objects.get(pk=1).liberado_admin)
        self.client.force_login(self.comum)
        self.assertContains(self.client.get(reverse("inicio")), "Agendamento autorizado")

    def test_revogar_desativa_execucao(self):
        ConfiguracaoAgendamento.objects.create(pk=1, liberado_admin=True, ativo=True, dias_semana=[0])
        self.client.force_login(self.admin)
        self.client.post(reverse("liberar_agendamento"), {"liberar":"nao"})
        config = ConfiguracaoAgendamento.objects.get(pk=1)
        self.assertFalse(config.ativo)
        self.assertFalse(config.liberado_admin)

    def test_ativacao_sem_liberacao_negada(self):
        self.client.force_login(self.admin)
        response = self.client.post(reverse("configurar_agendamento"), {"ativo":"on", "horario":"08:00", "dias_semana":["0"]})
        self.assertEqual(response.status_code, 403)

    def test_progresso_exige_autorizacao(self):
        registro = HistoricoExecucao.objects.create(usuario=self.admin, progresso=42, etapa_atual="Lendo planilhas")
        self.client.force_login(self.comum)
        self.assertEqual(
            self.client.get(
                reverse("consultar_progresso", args=[registro.pk])
            ).status_code,
            403
        )
        self.client.force_login(self.admin)
        data = self.client.get(
            reverse("consultar_progresso", args=[registro.pk])
        ).json()
        self.assertEqual(data["progresso"], 42)
