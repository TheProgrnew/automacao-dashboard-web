"""
URL configuration for projeto_web project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/5.2/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.contrib.auth import views as auth_views
from django.urls import path, reverse_lazy
from dashboard import views




urlpatterns = [
    path('admin/', admin.site.urls),
    path('senha/alterar/', auth_views.PasswordChangeView.as_view(template_name='alterar_senha.html', success_url=reverse_lazy('senha_alterada')), name='alterar_senha'),
    path('senha/alterada/', auth_views.PasswordChangeDoneView.as_view(template_name='senha_alterada.html'), name='senha_alterada'),
    path('usuarios/novo/', views.cadastrar_usuario, name='cadastrar_usuario'),
    path('cadastro/solicitar/', views.solicitar_cadastro, name='solicitar_cadastro'),
    path('usuarios/aprovacoes/', views.aprovar_cadastros, name='aprovar_cadastros'),
    path('dashboard/baixar/', views.baixar_dashboard, name='baixar_dashboard'),
    path('senha/recuperar/', auth_views.PasswordResetView.as_view(
        template_name='recuperar_senha.html',
        email_template_name='emails/recuperacao_senha.txt',
        subject_template_name='emails/assunto_recuperacao.txt',
        success_url=reverse_lazy('recuperacao_enviada'),
    ), name='recuperar_senha'),
    path('senha/recuperacao-enviada/', auth_views.PasswordResetDoneView.as_view(
        template_name='recuperacao_enviada.html',
    ), name='recuperacao_enviada'),
    path('senha/redefinir/<uidb64>/<token>/', auth_views.PasswordResetConfirmView.as_view(
        template_name='redefinir_senha.html',
        success_url=reverse_lazy('redefinicao_concluida'),
    ), name='password_reset_confirm'),
    path('senha/redefinida/', auth_views.PasswordResetCompleteView.as_view(
        template_name='redefinicao_concluida.html',
    ), name='redefinicao_concluida'),
    path('login/', auth_views.LoginView.as_view(template_name='login.html'), name='login'),
    path('logout/', auth_views.LogoutView.as_view(), name='logout'),
    path('', views.inicio, name='inicio'),
    path('atualizar/', views.atualizar_dashboard, name='atualizar_dashboard'),
    path('configurar/', views.configurar_automacao , name='configurar_automacao'),
    path('agendamento/liberar/', views.liberar_agendamento, name='liberar_agendamento'),
    path('atualizar/iniciar/', views.iniciar_atualizacao_assincrona, name='iniciar_atualizacao_assincrona'),
    path('atualizar/progresso/<int:execucao_id>/', views.consultar_progresso, name='consultar_progresso'),
    path('agendamento/configurar/', views.configurar_agendamento, name='configurar_agendamento'),
    
]
