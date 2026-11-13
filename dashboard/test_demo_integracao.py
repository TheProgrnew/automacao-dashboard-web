"""Integração do worker e do download com planilhas reais fictícias."""
import os
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from django.contrib.auth import get_user_model
from django.core.management import call_command
from django.test import TestCase, override_settings
from django.urls import reverse
from openpyxl import load_workbook
from preparar_demo import preparar_demo
from dashboard.models import HistoricoExecucao, ConfiguracaoAutomacao

class IntegracaoDemoTests(TestCase):
    def test_worker_grava_planilha_historico_progresso_e_download(self):
        user=get_user_model().objects.create_user(username='demo_integracao',password='senha-de-teste',is_staff=True)
        with TemporaryDirectory() as pasta:
            raiz=preparar_demo(Path(pasta))
            with override_settings(AUTOMACAO_PASTA_TESTE=raiz,AUTOMACAO_MODO_TESTE=True):
                registro=HistoricoExecucao.objects.create(usuario=user,modo_teste=True,origem='manual')
                with patch('automacao.executar_automacao.PASTA_DEMO',raiz), patch.dict(os.environ, {'AUTOMACAO_MOTOR': 'openpyxl'}):
                    call_command('executar_atualizacao',registro.pk)
                registro.refresh_from_db()
                self.assertEqual(registro.status,'sucesso')
                self.assertEqual(registro.progresso,100)
                self.assertIsNotNone(registro.terminado_em)
                wb=load_workbook(raiz/'dashboard'/'dashboard_exemplo.xlsx')
                self.assertEqual(wb['RESUMO']['B6'].value,8)
                wb.close()
                self.client.force_login(user)
                response=self.client.get(reverse('inicio'))
                self.assertContains(response,'dados inteiramente fictícios')
                response=self.client.get(reverse('baixar_dashboard'))
                self.assertEqual(response.status_code,200)
                self.assertEqual(response['Content-Type'],'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                conteudo=b''.join(response.streaming_content)
                self.assertEqual(conteudo[:2],b'PK')
                response.close()

    def test_modo_configurado_executa_e_baixa_ficticio(self):
        user=get_user_model().objects.create_user(username='demo_configurada',password='senha-de-teste',is_staff=True)
        with TemporaryDirectory() as pasta:
            raiz=preparar_demo(Path(pasta))
            ConfiguracaoAutomacao.objects.create(pk=1,caminho_dashboard=str(raiz/'dashboard'/'dashboard_exemplo.xlsx'),pasta_entrada=str(raiz/'entrada'),pasta_saida=str(raiz/'saida'))
            with override_settings(AUTOMACAO_MODO_TESTE=False,AUTOMACAO_PASTA_CONFIGURADA=raiz):
                self.client.force_login(user)
                with patch('automacao.executar_automacao.PASTA_CONFIGURADA',raiz), patch.dict(os.environ, {'AUTOMACAO_MOTOR':'openpyxl'}):
                    response=self.client.post(reverse('atualizar_dashboard'))
                self.assertEqual(response.status_code,200)
                self.assertEqual(HistoricoExecucao.objects.get().status,'sucesso')
                response=self.client.get(reverse('baixar_dashboard'))
                self.assertEqual(response.status_code,200)
                response.close()
