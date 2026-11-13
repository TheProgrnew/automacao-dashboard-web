import os
from io import StringIO
from unittest.mock import patch
from django.contrib.auth import get_user_model
from django.core.management import call_command, CommandError
from django.test import TestCase

class AdminEnvTests(TestCase):
    def test_cria_admin_sem_exibir_senha(self):
        senha='SenhaExclusivaDeTeste123!';out=StringIO()
        with patch.dict(os.environ,{'DASHBOARD_ADMIN_USUARIO':'admin_de_teste','DASHBOARD_ADMIN_EMAIL':'teste@example.invalid','DASHBOARD_ADMIN_SENHA':senha}):
            call_command('criar_admin_env',stdout=out)
        user=get_user_model().objects.get(username='admin_de_teste')
        self.assertTrue(user.is_superuser);self.assertTrue(user.check_password(senha))
        self.assertNotIn(senha,out.getvalue())
    def test_nao_cria_sem_senha(self):
        with patch.dict(os.environ,{'DASHBOARD_ADMIN_USUARIO':'admin','DASHBOARD_ADMIN_EMAIL':'teste@example.invalid','DASHBOARD_ADMIN_SENHA':''}):
            with self.assertRaises(CommandError):call_command('criar_admin_env')
        self.assertFalse(get_user_model().objects.exists())
    def test_nao_altera_usuario_existente(self):
        user=get_user_model().objects.create_user(username='existente',password='SenhaAntigaTeste123!')
        with patch.dict(os.environ,{'DASHBOARD_ADMIN_USUARIO':'existente','DASHBOARD_ADMIN_EMAIL':'teste@example.invalid','DASHBOARD_ADMIN_SENHA':'NovaSenhaTeste123!'}):
            with self.assertRaises(CommandError):call_command('criar_admin_env')
        user.refresh_from_db();self.assertTrue(user.check_password('SenhaAntigaTeste123!'))
    def test_recusa_senha_fraca(self):
        for senha in ['123', 'Abcd123!xyz']:
            with self.subTest(tamanho=len(senha)), patch.dict(os.environ,{'DASHBOARD_ADMIN_USUARIO':'novo','DASHBOARD_ADMIN_EMAIL':'teste@example.invalid','DASHBOARD_ADMIN_SENHA':senha}):
                with self.assertRaises(CommandError):call_command('criar_admin_env')
        self.assertFalse(get_user_model().objects.exists())
