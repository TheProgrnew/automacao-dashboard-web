"""Worker isolado para execução manual com progresso, adequado ao motor local de planilhas."""
import logging
from pathlib import Path
from django.conf import settings
from django.core.management.base import BaseCommand
from django.utils import timezone
from dashboard.models import HistoricoExecucao, ConfiguracaoAutomacao
from dashboard.bloqueio_execucao import bloquear_execucao, ExecucaoOcupada

logger = logging.getLogger(__name__)

class Command(BaseCommand):
    help = "Executa uma atualização previamente autorizada pela interface."

    def add_arguments(self, parser):
        parser.add_argument("registro_id", type=int)

    def handle(self, *args, **options):
        registro = HistoricoExecucao.objects.get(pk=options["registro_id"])
        if registro.status != "executando" or registro.origem != "manual":
            return
        try:
            with bloquear_execucao():
                if registro.modo_teste:
                    pasta = Path(settings.AUTOMACAO_PASTA_TESTE)
                    arquivos = list((pasta / "dashboard").glob("*.xlsx"))
                    if len(arquivos) != 1:
                        raise ValueError("Dashboard de teste ausente ou ambíguo")
                    dashboard, entrada, saida = arquivos[0], pasta / "entrada", pasta / "saida"
                else:
                    config = ConfiguracaoAutomacao.objects.filter(pk=1).first()
                    if config is None:
                        raise ValueError("Configuração de produção ausente")
                    dashboard, entrada, saida = Path(config.caminho_dashboard), Path(config.pasta_entrada), Path(config.pasta_saida)
                if not dashboard.is_file() or dashboard.suffix.lower() != ".xlsx" or not entrada.is_dir() or not saida.is_dir():
                    raise ValueError("Caminhos de automação inválidos")
                def progresso(valor, descricao):
                    HistoricoExecucao.objects.filter(pk=registro.pk).update(progresso=valor, etapa_atual=descricao)
                from automacao.executar_automacao import executar_automacao
                sucesso = executar_automacao(caminho_dashboard=str(dashboard), pasta_entrada=str(entrada),
                    pasta_saida=str(saida), modo_teste=registro.modo_teste, progresso_callback=progresso)
                registro.status = "sucesso" if sucesso else "falha"
                registro.detalhe = "Concluída" if sucesso else "A automação informou falha"
                if sucesso:
                    registro.progresso = 100
                    registro.etapa_atual = "Concluído"
        except Exception as erro:
            logger.exception("Falha na execução de atualização")
            registro.status = "falha"
            registro.detalhe = "Falha; consulte os logs"
            registro.etapa_atual = "Falha na execução"

            print(
                f"ERRO REAL DA AUTOMAÇÃO: {type(erro).__name__}: {erro}",
                flush=True
            )
        finally:
            registro.terminado_em = timezone.now()
            registro.save(update_fields=["status", "detalhe", "etapa_atual", "progresso", "terminado_em"])
