from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion

class Migration(migrations.Migration):
    dependencies = [("dashboard", "0002_rename_caminho_dasgboard_configuracaoautomacao_caminho_dashboard_and_more"), migrations.swappable_dependency(settings.AUTH_USER_MODEL)]
    operations = [migrations.CreateModel(name="HistoricoExecucao", fields=[
        ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
        ("iniciado_em", models.DateTimeField(auto_now_add=True)),
        ("terminado_em", models.DateTimeField(blank=True, null=True)),
        ("status", models.CharField(choices=[("executando", "Executando"), ("sucesso", "Sucesso"), ("falha", "Falha")], default="executando", max_length=12)),
        ("modo_teste", models.BooleanField(default=False)),
        ("detalhe", models.CharField(blank=True, max_length=240)),
        ("usuario", models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, to=settings.AUTH_USER_MODEL)),
    ], options={"ordering": ["-iniciado_em"]})]
