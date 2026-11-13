from django.db import migrations, models

class Migration(migrations.Migration):
    dependencies = [("dashboard", "0004_agendamento")]
    operations = [
        migrations.AddField(model_name="configuracaoagendamento", name="liberado_admin", field=models.BooleanField(default=False)),
        migrations.AddField(model_name="historicoexecucao", name="progresso", field=models.PositiveSmallIntegerField(default=0)),
        migrations.AddField(model_name="historicoexecucao", name="etapa_atual", field=models.CharField(max_length=160, blank=True)),
    ]
