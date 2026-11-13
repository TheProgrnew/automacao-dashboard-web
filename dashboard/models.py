from django.db import models

class ConfiguracaoAutomacao(models.Model):
    caminho_dashboard = models.CharField(max_length = 500)
    pasta_entrada = models.CharField(max_length=500)
    pasta_saida = models.CharField(max_length=500)
    
    atualizado_em = models.DateTimeField(auto_now = True)
    
    def __str__(self):
        return "Configuração da automação do Dashboard"

class HistoricoExecucao(models.Model):
    STATUS = [("executando", "Executando"), ("sucesso", "Sucesso"), ("falha", "Falha")]
    usuario = models.ForeignKey("auth.User", on_delete=models.SET_NULL, null=True, blank=True)
    iniciado_em = models.DateTimeField(auto_now_add=True)
    terminado_em = models.DateTimeField(null=True, blank=True)
    status = models.CharField(max_length=12, choices=STATUS, default="executando")
    modo_teste = models.BooleanField(default=False)
    origem = models.CharField(max_length=12, choices=[("manual", "Manual"), ("agendada", "Agendada")], default="manual")
    detalhe = models.CharField(max_length=240, blank=True)
    progresso = models.PositiveSmallIntegerField(default=0)
    etapa_atual = models.CharField(max_length=160, blank=True)

    class Meta:
        ordering = ["-iniciado_em"]

    def __str__(self):
        return f"Execução {self.pk} - {self.status}"


class ConfiguracaoAgendamento(models.Model):
    """Configuração única; desativada até decisão explícita do administrador."""
    ativo = models.BooleanField(default=False)
    horario = models.TimeField(default="08:00")
    # weekday() do Python: 0=segunda, 6=domingo.
    dias_semana = models.JSONField(default=list)
    atualizado_em = models.DateTimeField(auto_now=True)
    liberado_admin = models.BooleanField(default=False)

    def __str__(self):
        return "Agendamento automático"
