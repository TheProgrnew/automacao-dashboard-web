import os
import logging
from pathlib import Path

from django.shortcuts import render, redirect
from django.urls import reverse
from django.contrib.auth.decorators import login_required, user_passes_test
from django.contrib import messages
from django.http import HttpResponse, FileResponse, Http404
from django.contrib.auth import get_user_model
from django.contrib.auth.forms import UserCreationForm
from django import forms
from django.views.decorators.http import require_GET, require_POST

from .models import ConfiguracaoAutomacao, HistoricoExecucao, ConfiguracaoAgendamento
from .bloqueio_execucao import bloquear_execucao, ExecucaoOcupada
from django.utils import timezone

import sys

from django.conf import settings


PASTA_AUTOMACAO = settings.BASE_DIR / "automacao"

if str(PASTA_AUTOMACAO) not in sys.path:
    sys.path.insert(0, str(PASTA_AUTOMACAO))
    

from automacao.executar_automacao import executar_automacao



logger = logging.getLogger(__name__)

@login_required
def inicio(request):
    configuracao = ConfiguracaoAutomacao.objects.filter(
        pk = 1
    ).first()
    
    contexto = {
                'configuracao': configuracao,
                'historico': HistoricoExecucao.objects.select_related('usuario')[:15],
                'modo_teste': settings.AUTOMACAO_MODO_TESTE,
                'motor_demo': os.getenv('AUTOMACAO_MOTOR', 'win32com'),
                'origem_demo': os.getenv('AUTOMACAO_ORIGEM', 'local'),
                'caminho_demo': str(settings.AUTOMACAO_PASTA_TESTE / 'dashboard' / 'dashboard_exemplo.xlsx') if settings.AUTOMACAO_MODO_TESTE else (configuracao.caminho_dashboard if configuracao else ''),
                'pasta_configurada': str(settings.AUTOMACAO_PASTA_CONFIGURADA),
                'agendamento': ConfiguracaoAgendamento.objects.filter(pk=1).first() if request.user.is_staff else None,
                'agendamento_liberado': ConfiguracaoAgendamento.objects.filter(pk=1, liberado_admin=True).exists(),
                'agendamento_publico': ConfiguracaoAgendamento.objects.filter(pk=1, liberado_admin=True).first(),
                'dias_agendamento': list(ConfiguracaoAgendamento.objects.filter(pk=1).values_list('dias_semana', flat=True).first() or []) if request.user.is_staff else []
            }
    
    return render(
        request, 'index.html', contexto
        
    )

@login_required
@user_passes_test(lambda user: user.is_staff, login_url="inicio")
def configurar_automacao(request):
    if request.method != 'POST':
        return HttpResponse(
            "Método não permitido", status = 405
        )
        
    dashboard = request.POST.get('caminho_dashboard', '').strip()
    entrada = request.POST.get('pasta_entrada','').strip()
    saida = request.POST.get('pasta_saida', '').strip()
    
    if not dashboard or not entrada or not saida:
        return HttpResponse(
            "Erro: preencha os três campos.", status = 400
        )
    
    dashboard = Path(dashboard)
    entrada = Path(entrada)
    saida = Path(saida)
    
    if not dashboard.is_file():
        return HttpResponse(
        "Erro: o arquivo do Dashboard não foi encontrado.",
        status=400
    )

    if dashboard.suffix.lower() != '.xlsx':
        return HttpResponse(
            "Erro: o arquivo deve ter extensão .xlsx.",
            status=400
        )

    if not entrada.is_dir():
        return HttpResponse(
            "Erro: a pasta de entrada não foi encontrada.",
            status=400
        )

    if not saida.is_dir():
        return HttpResponse(
            "Erro: a pasta de saída não foi encontrada.",
            status=400
        )
    
    
    from automacao.motor_openpyxl import validar_caminhos
    raiz = settings.AUTOMACAO_PASTA_TESTE if settings.AUTOMACAO_MODO_TESTE else settings.AUTOMACAO_PASTA_CONFIGURADA
    try:
        validar_caminhos(dashboard, entrada, saida, raiz)
    except ValueError as erro:
        return HttpResponse(str(erro), status=400)

    ConfiguracaoAutomacao.objects.update_or_create(
        pk = 1,
        defaults= {
            'caminho_dashboard': str(dashboard.resolve()),
            'pasta_entrada': str(entrada.resolve()),
            'pasta_saida': str(saida.resolve())
        }
        
    )
    
    messages.success(
    request,
    'Configuração salva com sucesso!'
)

    return redirect('inicio')   



@login_required
@user_passes_test(lambda user: user.is_staff, login_url="inicio")
def atualizar_dashboard(request):
    if request.method != 'POST':
        return HttpResponse(
            "Método não permitido.",
            status=405
        )

    configuracao = None

    if not settings.AUTOMACAO_MODO_TESTE:
        configuracao = ConfiguracaoAutomacao.objects.filter(
            pk=1
        ).first()

        if configuracao is None:
            return HttpResponse(
                "Erro: nenhuma configuração foi salva.",
                status=400
            )

   
    if settings.AUTOMACAO_MODO_TESTE:
        pasta_teste = Path(settings.AUTOMACAO_PASTA_TESTE)

        arquivos_dashboard = list(
            (pasta_teste / "dashboard").glob("*.xlsx")
        )

        if len(arquivos_dashboard) != 1:
            return HttpResponse(
                "Erro: a pasta de testes deve conter exatamente um Dashboard .xlsx.",
                status=400
            )

        dashboard = arquivos_dashboard[0]
        entrada = pasta_teste / "entrada"
        saida = pasta_teste / "saida"
        # Não permitir que links simbólicos redirecionem a execução para fora.
        base_real = pasta_teste.resolve()
        for item in (dashboard, entrada, saida):
            if item.is_symlink() or not item.resolve().is_relative_to(base_real):
                return HttpResponse("Caminho de teste não permitido.", status=400)

    else:
        dashboard = Path(configuracao.caminho_dashboard)
        entrada = Path(configuracao.pasta_entrada)
        saida = Path(configuracao.pasta_saida)

    if not dashboard.is_file():
        return HttpResponse(
            "Erro: o arquivo do Dashboard não foi encontrado.",
            status=400
        )

    if dashboard.suffix.lower() != '.xlsx':
        return HttpResponse(
            "Erro: o arquivo do Dashboard deve ter extensão .xlsx.",
            status=400
        )

    if not entrada.is_dir():
        return HttpResponse(
            "Erro: a pasta de entrada não foi encontrada.",
            status=400
        )

    if not saida.is_dir():
        return HttpResponse(
            "Erro: a pasta de saída não foi encontrada.",
            status=400
        )

    
    # A trava abrange a chamada inteira: nunca iniciar duas atualizações no mesmo servidor.
    try:
        with bloquear_execucao():
            registro = HistoricoExecucao.objects.create(
                usuario=request.user, modo_teste=settings.AUTOMACAO_MODO_TESTE, origem="manual"
            )
            try:
                sucesso = executar_automacao(
                    caminho_dashboard=str(dashboard),
                    pasta_entrada=str(entrada),
                    pasta_saida=str(saida),
                    modo_teste=settings.AUTOMACAO_MODO_TESTE
                )
                registro.status = "sucesso" if sucesso else "falha"
                registro.detalhe = "Concluída" if sucesso else "A automação informou falha"
                return HttpResponse(
                    "Dashboard atualizado com sucesso!" if sucesso else
                    "A automação terminou, mas informou uma falha.",
                    status=200 if sucesso else 500
                )
            except Exception:
                logger.exception("Erro inesperado ao executar a automação.")
                registro.status = "falha"
                registro.detalhe = "Erro inesperado; consulte o log do servidor"
                return HttpResponse("Erro inesperado ao executar a automação.", status=500)
            finally:
                registro.terminado_em = timezone.now()
                registro.save(update_fields=["status", "detalhe", "terminado_em"])
    except ExecucaoOcupada:
        return HttpResponse("Já existe uma atualização em andamento.", status=409)


class CadastroUsuarioForm(UserCreationForm):
    email = forms.EmailField(required=True, label="E-mail")
    first_name = forms.CharField(max_length=150, required=False, label="Nome")
    last_name = forms.CharField(max_length=150, required=False, label="Sobrenome")

    class Meta(UserCreationForm.Meta):
        model = get_user_model()
        fields = ("username", "email", "first_name", "last_name", "password1", "password2")

    def clean_email(self):
        email = self.cleaned_data["email"].strip().lower()
        if get_user_model().objects.filter(email__iexact=email).exists():
            raise forms.ValidationError("Este e-mail já está cadastrado.")
        return email


def solicitar_cadastro(request):
    """Cadastro público: nenhuma conta nova recebe acesso antes da aprovação."""
    if request.user.is_authenticated:
        return redirect("inicio")
    form = CadastroUsuarioForm(request.POST or None)
    if request.method == "POST" and form.is_valid():
        usuario = form.save(commit=False)
        usuario.is_active = False
        usuario.is_staff = False
        usuario.is_superuser = False
        usuario.save()
        return render(request, "cadastro_solicitado.html")
    return render(request, "solicitar_cadastro.html", {"form": form})


@login_required
@user_passes_test(lambda user: user.is_staff, login_url="inicio")
def aprovar_cadastros(request):
    """Apenas administradores podem liberar contas pendentes, via POST e CSRF."""
    User = get_user_model()
    if request.method == "POST":
        try:
            usuario_id = int(request.POST.get("usuario_id", ""))
        except (TypeError, ValueError):
            messages.error(request, "Solicitação inválida.")
            return redirect("aprovar_cadastros")
        # A consulta exclui usuários administrativos e contas já aprovadas.
        pendente = User.objects.filter(pk=usuario_id, is_active=False,
                                       is_staff=False, is_superuser=False).first()
        if pendente is None:
            messages.error(request, "Cadastro pendente não encontrado.")
        elif request.POST.get("acao") == "aprovar":
            pendente.is_active = True
            pendente.save(update_fields=["is_active"])
            messages.success(request, "Cadastro aprovado.")
        elif request.POST.get("acao") == "rejeitar":
            pendente.delete()
            messages.success(request, "Solicitação rejeitada e excluída.")
        else:
            messages.error(request, "Ação inválida.")
        return redirect("aprovar_cadastros")
    pendentes = User.objects.filter(is_active=False, is_staff=False,
                                   is_superuser=False).order_by("date_joined", "pk")
    return render(request, "aprovar_cadastros.html", {"pendentes": pendentes})


@login_required
@user_passes_test(lambda user: user.is_staff, login_url="inicio")
def cadastrar_usuario(request):
    form = CadastroUsuarioForm(request.POST or None)
    if request.method == "POST" and form.is_valid():
        usuario = form.save(commit=False)
        usuario.is_staff = False
        usuario.is_superuser = False
        usuario.save()
        messages.success(request, "Usuário cadastrado com sucesso.")
        return redirect("cadastrar_usuario")
    return render(request, "cadastro_usuario.html", {"form": form})


@login_required
@require_GET
def baixar_dashboard(request):
    # Usuários comuns podem baixar; apenas administradores configuram ou executam.
    # Nunca aceitar um caminho de arquivo vindo de GET/POST.
    if settings.AUTOMACAO_MODO_TESTE:
        pasta = Path(settings.AUTOMACAO_PASTA_TESTE).resolve()
        arquivos = list((pasta / "dashboard").glob("*.xlsx"))
        if len(arquivos) != 1:
            raise Http404("Dashboard de teste não disponível.")
        arquivo = arquivos[0]
        if arquivo.is_symlink() or not arquivo.resolve().is_relative_to(pasta):
            raise Http404("Caminho de teste não permitido.")
    else:
        configuracao = ConfiguracaoAutomacao.objects.filter(pk=1).first()
        if not configuracao:
            raise Http404("Dashboard não configurado.")
        arquivo = Path(configuracao.caminho_dashboard)

    if not arquivo.is_file() or arquivo.suffix.lower() != ".xlsx":
        raise Http404("Dashboard não encontrado.")
    from automacao.motor_openpyxl import validar_caminhos
    raiz = Path(settings.AUTOMACAO_PASTA_TESTE if settings.AUTOMACAO_MODO_TESTE else settings.AUTOMACAO_PASTA_CONFIGURADA)
    try:
        validar_caminhos(arquivo, raiz / 'entrada', raiz / 'saida', raiz)
    except ValueError as erro:
        raise Http404(str(erro))
    # Evita download enquanto a automação está escrevendo o arquivo.
    try:
        with bloquear_execucao():
            # FileResponse transmite o arquivo; a trava precisa durar até o fim
            # da transmissão, então usamos uma cópia temporária independente.
            import tempfile
            import shutil
            temp = tempfile.TemporaryFile(mode="w+b")
            try:
                with arquivo.open("rb") as origem:
                    shutil.copyfileobj(origem, temp)
                temp.seek(0)
            except Exception:
                temp.close()
                raise
    except ExecucaoOcupada:
        return HttpResponse("Uma atualização está em andamento. Tente novamente após a conclusão.", status=409)
    return FileResponse(temp, as_attachment=True, filename=arquivo.name,
                        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@login_required
@user_passes_test(lambda user: user.is_staff, login_url="inicio")
@require_POST
def configurar_agendamento(request):
    """Apenas grava preferências; não cria nem altera tarefas do Windows."""
    from datetime import time
    ativo = request.POST.get("ativo") == "on"
    dias = request.POST.getlist("dias_semana")
    permitidos = {str(i) for i in range(7)}
    if len(dias) != len(set(dias)) or not set(dias).issubset(permitidos):
        return HttpResponse("Dias da semana inválidos.", status=400)
    horario_texto = request.POST.get("horario", "08:00")
    try:
        hora, minuto = (int(x) for x in horario_texto.split(":"))
        horario = time(hora, minuto)
    except (ValueError, TypeError):
        return HttpResponse("Horário inválido.", status=400)
    if ativo and not dias:
        return HttpResponse("Selecione ao menos um dia para ativar.", status=400)
    config, _ = ConfiguracaoAgendamento.objects.get_or_create(pk=1)
    if ativo and not config.liberado_admin:
        return HttpResponse("O administrador precisa liberar o recurso antes de ativá-lo.", status=403)
    config.ativo = ativo
    config.horario = horario
    config.dias_semana = sorted(int(d) for d in dias)
    config.save()
    messages.success(request, "Preferências de agendamento salvas. A tarefa do Windows deve ser configurada separadamente.")
    return redirect("inicio")


@login_required
@user_passes_test(lambda user: user.is_staff, login_url="inicio")
@require_POST
def liberar_agendamento(request):
    """Permissão global independente do acionamento automático."""
    config, _ = ConfiguracaoAgendamento.objects.get_or_create(pk=1)
    config.liberado_admin = request.POST.get("liberar") == "sim"
    if not config.liberado_admin:
        config.ativo = False
    config.save(update_fields=["liberado_admin", "ativo", "atualizado_em"])
    messages.success(request, "Permissão de agendamento atualizada.")
    return redirect("inicio")


@login_required
@require_GET
def consultar_progresso(request, execucao_id):
    from django.http import JsonResponse
    from django.shortcuts import get_object_or_404
    registro = get_object_or_404(HistoricoExecucao, pk=execucao_id)
    if not request.user.is_staff and registro.usuario_id != request.user.pk:
        return HttpResponse("Acesso negado.", status=403)
    return JsonResponse({"id": registro.pk, "status": registro.status,
                         "progresso": registro.progresso, "etapa": registro.etapa_atual})


@login_required
@user_passes_test(lambda user: user.is_staff, login_url="inicio")
@require_POST
def iniciar_atualizacao_assincrona(request):
    """Retorna imediatamente; subprocesso separado inicializa seu próprio motor local de planilhas."""
    import subprocess
    import os
    from django.http import JsonResponse
    if HistoricoExecucao.objects.filter(status="executando").exists():
        return JsonResponse({"erro": "Já existe uma atualização registrada em andamento."}, status=409)
    registro = HistoricoExecucao.objects.create(usuario=request.user,
                modo_teste=settings.AUTOMACAO_MODO_TESTE, origem="manual",
                progresso=0, etapa_atual="Aguardando início")
    comando = [sys.executable, str(settings.BASE_DIR / "manage.py"), "executar_atualizacao", str(registro.pk)]
    try:
        
        caminho_log = settings.BASE_DIR / "diagnostico_atualizacao.log"

        with open(caminho_log, "a", encoding="utf-8") as arquivo_log:
            kwargs = {
                "cwd": str(settings.BASE_DIR),
                "stdin": subprocess.DEVNULL,
                "stdout": arquivo_log,
                "stderr": subprocess.STDOUT,
                "close_fds": True,
            }

            if os.name == "nt":
                kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW

            subprocess.Popen(comando, **kwargs)
        
        
    except OSError:
        logger.exception("Falha ao iniciar processo de atualização")
        registro.status = "falha"
        registro.etapa_atual = "Falha ao iniciar processo"
        registro.terminado_em = timezone.now()
        registro.save(update_fields=["status", "etapa_atual", "terminado_em"])
        return JsonResponse({"erro": "Não foi possível iniciar a atualização."}, status=500)
    return JsonResponse({"id": registro.pk, "url": reverse("consultar_progresso", args=[registro.pk])}, status=202)
