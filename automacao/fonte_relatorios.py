"""Origem local (padrão) ou leitura opcional do Gmail pessoal por IMAP/TLS.

Nenhum e-mail é enviado, apagado ou marcado como lido. Apenas os quatro nomes
conhecidos são aceitos. As credenciais vêm do ambiente/.env e nunca são logadas.
"""
import email
from email import policy
from email.utils import getaddresses
import imaplib
import logging
import os
from pathlib import Path
import re
import shutil
import ssl
import tempfile
import zipfile
from io import BytesIO
from automacao.config import ARQUIVOS

LOGGER = logging.getLogger('automacao')
MAX_EMAIL = 15 * 1024 * 1024
MAX_ANEXO = 5 * 1024 * 1024
MAX_XLSX_DESCOMPRIMIDO = 30 * 1024 * 1024

class ErroEmail(ValueError):
    """Mensagem curta que pode ser exibida sem expor a resposta de autenticação."""

class FonteRelatorios:
    def __init__(self, entrada, saida):
        self.entrada, self.saida = Path(entrada), Path(saida)
        self.origem = os.getenv('AUTOMACAO_ORIGEM', 'local')
        self.mail = None
        self.selecionado = False
        self.usuario = ''
        self.remetente = ''
        self.assunto = ''
        if self.origem not in ('local', 'gmail'):
            raise ErroEmail('AUTOMACAO_ORIGEM deve ser local ou gmail.')

    def conectar(self):
        if self.origem == 'local':
            return
        self.usuario = os.getenv('EMAIL_USUARIO', '').strip()
        senha = ''.join(os.getenv('EMAIL_SENHA_APP', '').split())
        self.remetente = os.getenv('EMAIL_REMETENTE', '').strip() or self.usuario
        self.assunto = os.getenv('EMAIL_ASSUNTO', 'PORTFOLIO_DASHBOARD_DEMO').strip()
        host = os.getenv('EMAIL_IMAP_HOST', 'imap.gmail.com').strip()
        porta = os.getenv('EMAIL_IMAP_PORT', '993').strip()
        if host != 'imap.gmail.com' or porta != '993':
            raise ErroEmail('Este adaptador usa Gmail: EMAIL_IMAP_HOST=imap.gmail.com e EMAIL_IMAP_PORT=993.')
        if not self.usuario.casefold().endswith('@gmail.com') or len(senha) != 16:
            raise ErroEmail('Preencha EMAIL_USUARIO com seu Gmail e EMAIL_SENHA_APP com a senha de app de 16 caracteres, não a senha normal.')
        if not self.remetente or not self.assunto or any(c in self.assunto+self.remetente for c in '\r\n'):
            raise ErroEmail('Confira EMAIL_ASSUNTO e EMAIL_REMETENTE na .env.')
        try:
            self.mail = imaplib.IMAP4_SSL(host, int(porta), ssl_context=ssl.create_default_context(), timeout=30)
            status, _ = self.mail.login(self.usuario, senha)
            if status != 'OK':
                raise imaplib.IMAP4.error('login recusado')
            status, _ = self.mail.select('INBOX', readonly=True)
            if status != 'OK':
                raise imaplib.IMAP4.error('caixa indisponível')
            self.selecionado = True
        except imaplib.IMAP4.error:
            raise ErroEmail('Gmail recusou o acesso. Confira a conta, a verificação em duas etapas e a senha de app.') from None
        except (OSError, ssl.SSLError):
            raise ErroEmail('Não foi possível conectar ao Gmail por TLS na porta 993. Confira a internet e o acesso de rede.') from None

    @staticmethod
    def _validar_xlsx(conteudo):
        try:
            with zipfile.ZipFile(BytesIO(conteudo)) as pacote:
                nomes = set(pacote.namelist())
                if '[Content_Types].xml' not in nomes or 'xl/workbook.xml' not in nomes:
                    raise ErroEmail('Um anexo não é uma planilha XLSX válida.')
                if sum(i.file_size for i in pacote.infolist()) > MAX_XLSX_DESCOMPRIMIDO:
                    raise ErroEmail('Planilha grande demais para a demonstração.')
                if any('vbaProject' in n or n.startswith('xl/externalLinks/') or n == 'xl/connections.xml' for n in nomes):
                    raise ErroEmail('Use apenas as planilhas fictícias sem macros, conexões ou vínculos externos.')
        except zipfile.BadZipFile:
            raise ErroEmail('Um anexo não é uma planilha XLSX válida.') from None

    def _buscar_anexos(self):
        criterio = '"' + self.assunto.replace('\\', '\\\\').replace('"', '\\"') + '"'
        try:
            # UID maior representa mensagem recebida mais recentemente nesta caixa.
            status, resposta = self.mail.uid('search', None, 'SUBJECT', criterio)
            if status != 'OK':
                raise ErroEmail('Não foi possível pesquisar as mensagens no Gmail.')
            ids = sorted((resposta[0] or b'').split(), key=int, reverse=True)[:50]
            for uid in ids:
                status, tamanho = self.mail.uid('fetch', uid, '(RFC822.SIZE)')
                texto = b' '.join(p for p in tamanho if isinstance(p, bytes))
                match = re.search(rb'RFC822\.SIZE\s+(\d+)', texto)
                if status != 'OK' or not match:
                    raise ErroEmail('Não foi possível conferir o tamanho de uma mensagem.')
                if int(match.group(1)) > MAX_EMAIL:
                    raise ErroEmail('Mensagem grande demais. Envie somente os quatro relatórios fictícios.')
                status, partes = self.mail.uid('fetch', uid, '(BODY.PEEK[])')
                bruto = next((p[1] for p in partes if isinstance(p, tuple) and isinstance(p[1], bytes)), None)
                if status != 'OK' or bruto is None or len(bruto) > MAX_EMAIL:
                    raise ErroEmail('Não foi possível ler a mensagem de demonstração.')
                mensagem = email.message_from_bytes(bruto, policy=policy.default)
                autores = [endereco.casefold() for _, endereco in getaddresses(mensagem.get_all('From', []))]
                if str(mensagem.get('Subject', '')).strip() != self.assunto or self.remetente.casefold() not in autores:
                    continue
                esperados = {s['arquivo'] for s in ARQUIVOS.values()}
                anexos = {}
                for parte in mensagem.walk():
                    nome = parte.get_filename()
                    if nome not in esperados:
                        continue
                    if nome in anexos:
                        raise ErroEmail('Há anexos duplicados. Envie cada relatório uma única vez.')
                    conteudo = parte.get_payload(decode=True)
                    if not conteudo or len(conteudo) > MAX_ANEXO:
                        raise ErroEmail('Anexo vazio ou grande demais. Use os arquivos fictícios fornecidos.')
                    self._validar_xlsx(conteudo)
                    anexos[nome] = conteudo
                faltantes = sorted(esperados - anexos.keys())
                if faltantes:
                    raise ErroEmail('O e-mail mais recente está incompleto. Anexe os quatro arquivos; faltam: '+', '.join(faltantes))
                return anexos
        except imaplib.IMAP4.error:
            raise ErroEmail('A leitura IMAP falhou. Tente novamente após conferir sua conexão e a senha de app.') from None
        except OSError:
            raise ErroEmail('A conexão caiu durante a leitura do Gmail. Tente novamente.') from None
        raise ErroEmail('Nenhum e-mail correspondente nas últimas 50 mensagens encontradas. Confira assunto, remetente e Caixa de entrada.')

    def preparar(self):
        if self.origem == 'gmail':
            anexos = self._buscar_anexos()
            # Só instala arquivos quando o e-mail tiver o conjunto completo.
            with tempfile.TemporaryDirectory(prefix='anexos_', dir=self.entrada) as pasta:
                pasta = Path(pasta)
                for nome, conteudo in anexos.items():
                    (pasta/nome).write_bytes(conteudo)
                for nome in anexos:
                    os.replace(pasta/nome, self.entrada/nome)
            LOGGER.info('Quatro relatórios recebidos pelo Gmail; nenhum conteúdo de mensagem foi registrado no log.')
        for spec in ARQUIVOS.values():
            shutil.copyfile(self.entrada/spec['arquivo'], self.saida/spec['arquivo'])

    def encerrar(self):
        if self.mail is not None:
            try:
                self.mail.logout()
            except Exception:
                LOGGER.warning('Não foi possível confirmar o encerramento da sessão IMAP.')
            finally:
                self.mail = None
                self.selecionado = False
