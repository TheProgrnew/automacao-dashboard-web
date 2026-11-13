from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse


class CadastroAprovacaoTests(TestCase):
    def setUp(self):
        User = get_user_model()
        self.admin = User.objects.create_user(username="gestor_aprovacao", password="SenhaForte123!", is_staff=True)
        self.comum = User.objects.create_user(username="leitor_aprovacao", password="SenhaForte123!")
        self.dados = {"username": "novo_pendente", "email": "novo@example.com", "first_name": "Novo", "last_name": "Usuario", "password1": "SenhaForte123!", "password2": "SenhaForte123!"}

    def test_link_cadastro_login(self):
        resposta = self.client.get(reverse("login"))
        self.assertContains(resposta, reverse("solicitar_cadastro"))

    def test_cadastro_publico_inativo(self):
        resposta = self.client.post(reverse("solicitar_cadastro"), self.dados)
        self.assertEqual(resposta.status_code, 200)
        usuario = get_user_model().objects.get(username="novo_pendente")
        self.assertFalse(usuario.is_active)
        self.assertFalse(usuario.is_staff)
        self.assertFalse(usuario.is_superuser)
        self.assertFalse(self.client.login(username="novo_pendente", password="SenhaForte123!"))

    def test_comum_nao_aprova(self):
        pendente = get_user_model().objects.create_user(username="pendente", password="SenhaForte123!", is_active=False)
        self.client.force_login(self.comum)
        resposta = self.client.post(reverse("aprovar_cadastros"), {"usuario_id": pendente.pk, "acao": "aprovar"})
        self.assertEqual(resposta.status_code, 302)
        pendente.refresh_from_db()
        self.assertFalse(pendente.is_active)

    def test_admin_aprova(self):
        pendente = get_user_model().objects.create_user(username="pendente", password="SenhaForte123!", is_active=False)
        self.client.force_login(self.admin)
        resposta = self.client.post(reverse("aprovar_cadastros"), {"usuario_id": pendente.pk, "acao": "aprovar"})
        self.assertRedirects(resposta, reverse("aprovar_cadastros"))
        pendente.refresh_from_db()
        self.assertTrue(pendente.is_active)
        self.assertFalse(pendente.is_staff)

    def test_admin_rejeita(self):
        pendente = get_user_model().objects.create_user(username="pendente", password="SenhaForte123!", is_active=False)
        self.client.force_login(self.admin)
        self.client.post(reverse("aprovar_cadastros"), {"usuario_id": pendente.pk, "acao": "rejeitar"})
        self.assertFalse(get_user_model().objects.filter(pk=pendente.pk).exists())

    def test_nao_aprova_conta_staff_inativa(self):
        pendente = get_user_model().objects.create_user(username="staff_pendente", password="SenhaForte123!", is_active=False, is_staff=True)
        self.client.force_login(self.admin)
        self.client.post(reverse("aprovar_cadastros"), {"usuario_id": pendente.pk, "acao": "aprovar"})
        pendente.refresh_from_db()
        self.assertFalse(pendente.is_active)

    def test_anonimo_nao_ve_aprovacoes(self):
        self.assertEqual(self.client.get(reverse("aprovar_cadastros")).status_code, 302)
