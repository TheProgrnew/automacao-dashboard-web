from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):
    dependencies = [("dashboard", "0003_historicoexecucao")]
    operations = [
        migrations.CreateModel(name="ConfiguracaoAgendamento", fields=[
            ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
            ("ativo", models.BooleanField(default=False)),
            ("horario", models.TimeField(default="08:00")),
            ("dias_semana", models.JSONField(default=list)),
            ("atualizado_em", models.DateTimeField(auto_now=True)),
        ]),
        migrations.AddField(model_name="historicoexecucao", name="origem", field=models.CharField(choices=[("manual", "Manual"), ("agendada", "Agendada")], default="manual", max_length=12)),
    ]
