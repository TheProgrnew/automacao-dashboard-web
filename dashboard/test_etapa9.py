"""Testes isolados das funcionalidades web; não executam o Excel."""
from pathlib import Path
from tempfile import TemporaryDirectory
from django.contrib.auth import get_user_model
from django.test import TestCase, override_settings
from django.urls import reverse
from dashboard.models import ConfiguracaoAutomacao


class Etapa9WebTests(TestCase):
    def setUp(self):
        User = get_user_model()
        self.staff = User.objects.create_user(username="gestor", password="SenhaForte123!", is_staff=True)
        self.comum = User.objects.create_user(username="leitor", password="SenhaForte123!")

    def test_login_anonimo_redireciona(self):
        resposta = self.client.get(reverse("inicio"))
        self.assertEqual(resposta.status_code, 302)

    def test_usuario_comum_nao_cadastra_outros(self):
        self.client.force_login(self.comum)
        self.assertEqual(self.client.get(reverse("cadastrar_usuario")).status_code, 302)

    def test_staff_acessa_cadastro(self):
        self.client.force_login(self.staff)
        self.assertEqual(self.client.get(reverse("cadastrar_usuario")).status_code, 200)

    def test_anonimo_nao_baixa(self):
        self.assertEqual(self.client.get(reverse("baixar_dashboard")).status_code, 302)

    def test_download_sem_configuracao_retorna_404(self):
        self.client.force_login(self.comum)
        with override_settings(AUTOMACAO_MODO_TESTE=False):
            self.assertEqual(self.client.get(reverse("baixar_dashboard")).status_code, 404)

    def test_download_do_arquivo_configurado(self):
        from preparar_demo import preparar_demo
        self.client.force_login(self.comum)
        with TemporaryDirectory() as pasta:
            base=preparar_demo(Path(pasta))
            arquivo=base/'dashboard'/'dashboard_exemplo.xlsx'
            ConfiguracaoAutomacao.objects.create(pk=1,caminho_dashboard=str(arquivo),pasta_entrada=str(base/'entrada'),pasta_saida=str(base/'saida'))
            with override_settings(AUTOMACAO_MODO_TESTE=False,AUTOMACAO_PASTA_CONFIGURADA=base):
                resposta=self.client.get(reverse('baixar_dashboard'))
                self.assertEqual(resposta.status_code,200)
                self.assertEqual(b''.join(resposta.streaming_content),arquivo.read_bytes())
                resposta.close()

    def test_download_nao_aceita_post(self):
        self.client.force_login(self.comum)
        self.assertEqual(self.client.post(reverse("baixar_dashboard")).status_code, 405)
