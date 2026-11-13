"""Cria explicitamente o primeiro administrador com valores da .env."""
import os
from django.contrib.auth import get_user_model
from django.contrib.auth.password_validation import validate_password
from django.core.exceptions import ValidationError
from django.core.management.base import BaseCommand, CommandError
from django.db import IntegrityError, transaction

class Command(BaseCommand):
    help = 'Cria um administrador usando DASHBOARD_ADMIN_USUARIO, DASHBOARD_ADMIN_EMAIL e DASHBOARD_ADMIN_SENHA; não altera usuários existentes.'

    def handle(self, *args, **options):
        nome = os.getenv('DASHBOARD_ADMIN_USUARIO', '').strip()
        email = os.getenv('DASHBOARD_ADMIN_EMAIL', '').strip()
        senha = os.getenv('DASHBOARD_ADMIN_SENHA', '')
        if not nome or not email or not senha:
            raise CommandError('Preencha DASHBOARD_ADMIN_USUARIO, DASHBOARD_ADMIN_EMAIL e DASHBOARD_ADMIN_SENHA na sua .env. Não use a senha do Gmail.')
        if len(senha) < 12:
            raise CommandError('Use uma senha exclusiva do Dashboard com pelo menos 12 caracteres.')
        User = get_user_model()
        usuario = User(username=nome, email=email)
        try:
            validate_password(senha, usuario)
            usuario.full_clean(exclude=['password'])
        except ValidationError:
            raise CommandError('Confira o usuário, o e-mail e a senha do Dashboard. Use senha exclusiva com pelo menos 12 caracteres, variando letras, números e símbolos; o nome pode já estar cadastrado.') from None
        try:
            with transaction.atomic():
                if User.objects.filter(username=nome).exists():
                    raise CommandError('Este usuário já existe. Nenhuma senha foi alterada. Use python manage.py changepassword NOME_DO_USUARIO para redefinir.')
                User.objects.create_superuser(username=nome, email=email, password=senha)
        except IntegrityError:
            raise CommandError('Não foi possível criar o usuário; confira se ele já existe.') from None
        self.stdout.write(self.style.SUCCESS('Administrador criado. A senha não será exibida. Entre no Dashboard com o usuário definido na .env.'))
