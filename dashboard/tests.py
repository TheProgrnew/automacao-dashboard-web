from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from django.test import TestCase, override_settings
from django.contrib.auth import get_user_model
from django.urls import reverse

from dashboard.models import ConfiguracaoAutomacao


@override_settings(AUTOMACAO_MODO_TESTE=False)
class TesteAtualizarDashboard(TestCase):
    def setUp(self):
        user = get_user_model().objects.create_user(username="operador", password="teste-forte-123", is_staff=True)
        self.client.force_login(user)


    def test_atualizar_exige_metodo_post(self):
        resposta = self.client.get(
            reverse('atualizar_dashboard')
        )

        self.assertEqual(
            resposta.status_code,
            405
        )

    def test_atualizar_sem_configuracao(self):
        resposta = self.client.post(
            reverse('atualizar_dashboard')
        )

        self.assertContains(
            resposta,
            "nenhuma configuração foi salva",
            status_code=400
        )


    @patch('dashboard.views.executar_automacao')
    def test_atualizar_com_dashboard_inexistente(
        self,
        automacao_mock
    ):
        ConfiguracaoAutomacao.objects.create(
            caminho_dashboard=r'C:\Teste\inexistente.xlsx',
            pasta_entrada=r'C:\Teste\Entrada',
            pasta_saida=r'C:\Teste\Saida'
        )

        resposta = self.client.post(
            reverse('atualizar_dashboard')
        )

        self.assertEqual(
            resposta.status_code,
            400
        )

        automacao_mock.assert_not_called()


    @patch(
        'dashboard.views.executar_automacao',
        return_value=True
    )
    def test_atualizar_chama_automacao_com_caminhos(
        self,
        automacao_mock
    ):
        with TemporaryDirectory() as diretorio:
            base = Path(diretorio)

            dashboard = base / 'dashboard.xlsx'
            dashboard.touch()

            entrada = base / 'Entrada'
            entrada.mkdir()

            saida = base / 'Saida'
            saida.mkdir()

            configuracao = ConfiguracaoAutomacao.objects.create(
                caminho_dashboard=str(dashboard),
                pasta_entrada=str(entrada),
                pasta_saida=str(saida)
            )

            resposta = self.client.post(
                reverse('atualizar_dashboard')
            )

            self.assertEqual(
                resposta.status_code,
                200
            )

            automacao_mock.assert_called_once_with(
                caminho_dashboard=configuracao.caminho_dashboard,
                pasta_entrada=configuracao.pasta_entrada,
                pasta_saida=configuracao.pasta_saida,
                modo_teste = False
            )
    @patch('dashboard.views.executar_automacao')
    def test_atualizar_com_extensao_invalida(
        self,
        automacao_mock
    ):
        with TemporaryDirectory() as diretorio:
            base = Path(diretorio)

            dashboard = base / 'dashboard.csv'
            dashboard.touch()

            entrada = base / 'Entrada'
            entrada.mkdir()

            saida = base / 'Saida'
            saida.mkdir()

            ConfiguracaoAutomacao.objects.create(
                caminho_dashboard=str(dashboard),
                pasta_entrada=str(entrada),
                pasta_saida=str(saida)
            )

            resposta = self.client.post(
                reverse('atualizar_dashboard')
            )

            self.assertEqual(
                resposta.status_code,
                400
            )

            automacao_mock.assert_not_called()

    @patch('dashboard.views.executar_automacao')
    def test_atualizar_com_pasta_entrada_inexistente(
        self,
        automacao_mock
    ):
        with TemporaryDirectory() as diretorio:
            base = Path(diretorio)

            dashboard = base / 'dashboard.xlsx'
            dashboard.touch()

            entrada = base / 'Entrada'

            saida = base / 'Saida'
            saida.mkdir()

            ConfiguracaoAutomacao.objects.create(
                caminho_dashboard=str(dashboard),
                pasta_entrada=str(entrada),
                pasta_saida=str(saida)
            )

            resposta = self.client.post(
                reverse('atualizar_dashboard')
            )

            self.assertEqual(
                resposta.status_code,
                400
            )

            automacao_mock.assert_not_called()
            
    @patch('dashboard.views.executar_automacao')
    def test_atualizar_com_pasta_saida_inexistente(
        self,
        automacao_mock
    ):
        with TemporaryDirectory() as diretorio:
            base = Path(diretorio)

            dashboard = base / 'dashboard.xlsx'
            dashboard.touch()

            entrada = base / 'Entrada'
            entrada.mkdir()

            saida = base / 'Saida'

            ConfiguracaoAutomacao.objects.create(
                caminho_dashboard=str(dashboard),
                pasta_entrada=str(entrada),
                pasta_saida=str(saida)
            )

            resposta = self.client.post(
                reverse('atualizar_dashboard')
            )

            self.assertEqual(
                resposta.status_code,
                400
            )

            automacao_mock.assert_not_called()

    @patch(
        'dashboard.views.executar_automacao',
        return_value=False
    )
    def test_atualizar_quando_automacao_retorna_false(
        self,
        automacao_mock
    ):
        with TemporaryDirectory() as diretorio:
            base = Path(diretorio)

            dashboard = base / 'dashboard.xlsx'
            dashboard.touch()

            entrada = base / 'Entrada'
            entrada.mkdir()

            saida = base / 'Saida'
            saida.mkdir()

            ConfiguracaoAutomacao.objects.create(
                caminho_dashboard=str(dashboard),
                pasta_entrada=str(entrada),
                pasta_saida=str(saida)
            )

            resposta = self.client.post(
                reverse('atualizar_dashboard')
            )

            self.assertEqual(
                resposta.status_code,
                500
            )

            automacao_mock.assert_called_once()

    @patch(
        'dashboard.views.executar_automacao',
        side_effect=RuntimeError("Falha simulada no Excel")
    )
    def test_atualizar_quando_automacao_lanca_excecao(
        self,
        automacao_mock
    ):
        with TemporaryDirectory() as diretorio:
            base = Path(diretorio)

            dashboard = base / 'dashboard.xlsx'
            dashboard.touch()

            entrada = base / 'Entrada'
            entrada.mkdir()

            saida = base / 'Saida'
            saida.mkdir()

            ConfiguracaoAutomacao.objects.create(
                caminho_dashboard=str(dashboard),
                pasta_entrada=str(entrada),
                pasta_saida=str(saida)
            )

            resposta = self.client.post(
                reverse('atualizar_dashboard')
            )

            self.assertEqual(
                resposta.status_code,
                500
            )

            self.assertContains(
                resposta,
                "Erro inesperado ao executar a automação.",
                status_code=500
            )

            automacao_mock.assert_called_once()
            
class TesteConfigurarAutomacao(TestCase):
    def setUp(self):
        user = get_user_model().objects.create_user(username="operador", password="teste-forte-123", is_staff=True)
        self.client.force_login(user)


    def test_configurar_exige_metodo_post(self):
        resposta = self.client.get(
            reverse('configurar_automacao')
        )

        self.assertEqual(
            resposta.status_code,
            405
        )
        
    def test_configurar_com_campos_vazios(self):
        resposta = self.client.post(
            reverse('configurar_automacao'),
            data={
                'caminho_dashboard': '',
                'pasta_entrada': '',
                'pasta_saida': ''
            }
        )

        self.assertEqual(
            resposta.status_code,
            400
        )

        self.assertFalse(
            ConfiguracaoAutomacao.objects.exists()
        )

    def test_configurar_com_dashboard_inexistente(self):
        with TemporaryDirectory() as diretorio:
            base = Path(diretorio)

            dashboard = base / 'dashboard_inexistente.xlsx'

            entrada = base / 'Entrada'
            entrada.mkdir()

            saida = base / 'Saida'
            saida.mkdir()

            resposta = self.client.post(
                reverse('configurar_automacao'),
                data={
                    'caminho_dashboard': str(dashboard),
                    'pasta_entrada': str(entrada),
                    'pasta_saida': str(saida)
                }
            )

            self.assertEqual(
                resposta.status_code,
                400
            )

            self.assertFalse(
                ConfiguracaoAutomacao.objects.exists()
            )
            
    def test_configurar_com_extensao_invalida(self):
        with TemporaryDirectory() as diretorio:
            base = Path(diretorio)

            dashboard = base / 'dashboard.csv'
            dashboard.touch()

            entrada = base / 'Entrada'
            entrada.mkdir()

            saida = base / 'Saida'
            saida.mkdir()

            resposta = self.client.post(
                reverse('configurar_automacao'),
                data={
                    'caminho_dashboard': str(dashboard),
                    'pasta_entrada': str(entrada),
                    'pasta_saida': str(saida)
                }
            )

            self.assertEqual(
                resposta.status_code,
                400
            )

            self.assertFalse(
                ConfiguracaoAutomacao.objects.exists()
            )
            
    def test_configurar_com_pasta_entrada_inexistente(self):
        with TemporaryDirectory() as diretorio:
            base = Path(diretorio)

            dashboard = base / 'dashboard.xlsx'
            dashboard.touch()

            entrada = base / 'Entrada'

            saida = base / 'Saida'
            saida.mkdir()

            resposta = self.client.post(
                reverse('configurar_automacao'),
                data={
                    'caminho_dashboard': str(dashboard),
                    'pasta_entrada': str(entrada),
                    'pasta_saida': str(saida)
                }
            )

            self.assertEqual(
                resposta.status_code,
                400
            )

            self.assertFalse(
                ConfiguracaoAutomacao.objects.exists()
            )

    def test_configurar_com_pasta_saida_inexistente(self):
        with TemporaryDirectory() as diretorio:
            base = Path(diretorio)

            dashboard = base / 'dashboard.xlsx'
            dashboard.touch()

            entrada = base / 'Entrada'
            entrada.mkdir()

            saida = base / 'Saida'

            resposta = self.client.post(
                reverse('configurar_automacao'),
                data={
                    'caminho_dashboard': str(dashboard),
                    'pasta_entrada': str(entrada),
                    'pasta_saida': str(saida)
                }
            )

            self.assertEqual(
                resposta.status_code,
                400
            )

            self.assertFalse(
                ConfiguracaoAutomacao.objects.exists()
            )

    def test_configurar_salva_caminhos_validos(self):
        from preparar_demo import preparar_demo
        with TemporaryDirectory() as diretorio:
            base=preparar_demo(Path(diretorio))
            dashboard=base/'dashboard'/'dashboard_exemplo.xlsx'
            entrada,saida=base/'entrada',base/'saida'
            with override_settings(AUTOMACAO_MODO_TESTE=False,AUTOMACAO_PASTA_CONFIGURADA=base):
                resposta=self.client.post(reverse('configurar_automacao'),data={'caminho_dashboard':str(dashboard),'pasta_entrada':str(entrada),'pasta_saida':str(saida)})
            self.assertEqual(resposta.status_code,302)
            config=ConfiguracaoAutomacao.objects.get(pk=1)
            self.assertEqual(config.caminho_dashboard,str(dashboard.resolve()))
            self.assertEqual(config.pasta_entrada,str(entrada.resolve()))
            self.assertEqual(config.pasta_saida,str(saida.resolve()))

    def test_configurar_atualiza_configuracao_existente(self):
        from preparar_demo import preparar_demo
        with TemporaryDirectory() as diretorio:
            base=preparar_demo(Path(diretorio))
            dashboard=base/'dashboard'/'dashboard_exemplo.xlsx'
            entrada,saida=base/'entrada',base/'saida'
            ConfiguracaoAutomacao.objects.create(pk=1,caminho_dashboard='configuracao-antiga',pasta_entrada='entrada-antiga',pasta_saida='saida-antiga')
            with override_settings(AUTOMACAO_MODO_TESTE=False,AUTOMACAO_PASTA_CONFIGURADA=base):
                resposta=self.client.post(reverse('configurar_automacao'),data={'caminho_dashboard':str(dashboard),'pasta_entrada':str(entrada),'pasta_saida':str(saida)})
            self.assertEqual(resposta.status_code,302)
            self.assertEqual(ConfiguracaoAutomacao.objects.count(),1)
            self.assertEqual(ConfiguracaoAutomacao.objects.get(pk=1).pasta_saida,str(saida.resolve()))

class TestePaginaInicial(TestCase):
    def setUp(self):
        user = get_user_model().objects.create_user(username="operador", password="teste-forte-123", is_staff=True)
        self.client.force_login(user)


    def test_inicio_carrega_configuracao_salva(self):
        configuracao = ConfiguracaoAutomacao.objects.create(
            pk=1,
            caminho_dashboard=r'C:\Teste\dashboard.xlsx',
            pasta_entrada=r'C:\Teste\Entrada',
            pasta_saida=r'C:\Teste\Saida'
        )

        resposta = self.client.get('/')

        self.assertEqual(
            resposta.status_code,
            200
        )

        self.assertTemplateUsed(
            resposta,
            'index.html'
        )

        self.assertEqual(
            resposta.context['configuracao'],
            configuracao
        )
        
        
@override_settings(AUTOMACAO_MODO_TESTE=True)
class TesteModoTesteAutomacao(TestCase):
    def setUp(self):
        user = get_user_model().objects.create_user(username="operador", password="teste-forte-123", is_staff=True)
        self.client.force_login(user)


    @patch(
        'dashboard.views.executar_automacao',
        return_value=True
    )
    def test_modo_teste_utiliza_caminhos_locais(
        self,
        automacao_mock
    ):
        with TemporaryDirectory() as diretorio:
            base = Path(diretorio)

            dashboard_dir = base / "dashboard"
            dashboard_dir.mkdir()

            dashboard = dashboard_dir / "dashboard.xlsx"
            dashboard.touch()

            entrada = base / "entrada"
            entrada.mkdir()

            saida = base / "saida"
            saida.mkdir()

            with override_settings(AUTOMACAO_PASTA_TESTE=base):
                resposta = self.client.post(
                    reverse('atualizar_dashboard')
                )

            self.assertEqual(resposta.status_code, 200)

            automacao_mock.assert_called_once_with(
                caminho_dashboard=str(dashboard),
                pasta_entrada=str(entrada),
                pasta_saida=str(saida),
                modo_teste=True
            )