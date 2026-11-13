"""Testes do fluxo de recuperação de senha (sem envio real de e-mail)."""
from django.contrib.auth import get_user_model
from django.core import mail
from django.test import TestCase, override_settings
from django.urls import reverse
import re


@override_settings(EMAIL_BACKEND="django.core.mail.backends.locmem.EmailBackend")
class RecuperacaoSenhaTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="usuario_recuperacao", email="usuario@example.com", password="SenhaAntiga123!"
        )

    def test_link_disponivel_na_pagina_login(self):
        resposta = self.client.get(reverse("login"))
        self.assertContains(resposta, reverse("recuperar_senha"))

    def test_solicitacao_envia_link_para_usuario_ativo(self):
        resposta = self.client.post(reverse("recuperar_senha"), {"email": "usuario@example.com"})
        self.assertRedirects(resposta, reverse("recuperacao_enviada"))
        self.assertEqual(len(mail.outbox), 1)
        self.assertIn("/senha/redefinir/", mail.outbox[0].body)

    def test_email_desconhecido_nao_revela_conta(self):
        resposta = self.client.post(reverse("recuperar_senha"), {"email": "naoexiste@example.com"})
        self.assertRedirects(resposta, reverse("recuperacao_enviada"))
        self.assertEqual(len(mail.outbox), 0)

    def test_usuario_inativo_nao_recebe_link(self):
        self.user.is_active = False
        self.user.save()
        self.client.post(reverse("recuperar_senha"), {"email": "usuario@example.com"})
        self.assertEqual(len(mail.outbox), 0)

    def test_redefinicao_com_link_valido(self):
        self.client.post(reverse("recuperar_senha"), {"email": "usuario@example.com"})
        link = re.search(r"https?://[^\s]+/senha/redefinir/[^\s]+", mail.outbox[0].body).group(0)
        from urllib.parse import urlsplit
        url = urlsplit(link).path
        resposta = self.client.get(url)
        self.assertEqual(resposta.status_code, 302)  # Django remove token da URL visível.
        resposta = self.client.post(resposta.url, {
            "new_password1": "NovaSenhaSegura456!", "new_password2": "NovaSenhaSegura456!"
        })
        self.assertRedirects(resposta, reverse("redefinicao_concluida"))
        self.user.refresh_from_db()
        self.assertTrue(self.user.check_password("NovaSenhaSegura456!"))
        self.assertFalse(self.user.check_password("SenhaAntiga123!"))
