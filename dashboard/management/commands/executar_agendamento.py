"""Executar via Agendador de Tarefas em intervalos definidos pelo administrador.

Não instala tarefas nem ativa o agendamento. Verifica configurações no banco
antes de iniciar e só roda no minuto configurado, uma vez por data.
"""
from pathlib import Path
import logging

from django.conf import settings
from django.core.management.base import BaseCommand
from django.utils import timezone
from dashboard.bloqueio_execucao import bloquear_execucao, ExecucaoOcupada
from dashboard.models import ConfiguracaoAgendamento, ConfiguracaoAutomacao, HistoricoExecucao

logger = logging.getLogger(__name__)


class Command(BaseCommand):
    help = "Verifica e executa agendamento ativo (sem instalar tarefa no Windows)."

    def handle(self, *args, **options):
        agora = timezone.localtime()
        config = ConfiguracaoAgendamento.objects.filter(pk=1, ativo=True, liberado_admin=True).first()
        if not config or agora.weekday() not in config.dias_semana:
            self.stdout.write("Agendamento desativado ou dia não selecionado.")
            return
        if (agora.hour, agora.minute) != (config.horario.hour, config.horario.minute):
            self.stdout.write("Fora do horário configurado.")
            return
        if settings.AUTOMACAO_MODO_TESTE:
            self.stderr.write("Execução agendada bloqueada: modo de teste ativo.")
            return
        try:
            with bloquear_execucao():
                # Revalidar dentro da trava para evitar repetição no mesmo dia.
                if HistoricoExecucao.objects.filter(origem="agendada", iniciado_em__date=agora.date()).exists():
                    self.stdout.write("Execução agendada já registrada hoje.")
                    return
                caminhos = ConfiguracaoAutomacao.objects.filter(pk=1).first()
                if not caminhos:
                    self.stderr.write("Configuração do Dashboard ausente.")
                    return
                dashboard = Path(caminhos.caminho_dashboard)
                entrada = Path(caminhos.pasta_entrada)
                saida = Path(caminhos.pasta_saida)
                if not (dashboard.is_file() and dashboard.suffix.lower() == ".xlsx" and entrada.is_dir() and saida.is_dir()):
                    self.stderr.write("Caminhos de automação inválidos.")
                    return
                registro = HistoricoExecucao.objects.create(origem="agendada", usuario=None, modo_teste=False)
                try:
                    def progresso(valor, descricao):
                        HistoricoExecucao.objects.filter(pk=registro.pk).update(progresso=valor, etapa_atual=descricao)
                    import sys
                    pasta_automacao = str(Path(settings.BASE_DIR) / "automacao")
                    if pasta_automacao not in sys.path:
                        sys.path.insert(0, pasta_automacao)
                    from automacao.executar_automacao import executar_automacao
                    sucesso = executar_automacao(caminho_dashboard=str(dashboard), pasta_entrada=str(entrada), pasta_saida=str(saida), modo_teste=False, progresso_callback=progresso)
                    registro.status = "sucesso" if sucesso else "falha"
                    registro.detalhe = "Agendamento concluído" if sucesso else "Agendamento informou falha"
                except Exception:
                    logger.exception("Falha na execução agendada")
                    registro.status = "falha"
                    registro.detalhe = "Erro na execução agendada; consulte os logs"
                finally:
                    if registro.status == "sucesso":
                        registro.progresso = 100
                        registro.etapa_atual = "Concluído"
                    registro.terminado_em = timezone.now()
                    registro.save(update_fields=["status", "detalhe", "terminado_em", "progresso", "etapa_atual"])
        except ExecucaoOcupada:
            self.stdout.write("Outra execução já está em andamento.")
