"""Testes offline: nenhuma conta real é acessada e nenhum e-mail é enviado."""
from email.message import EmailMessage
import hashlib
import imaplib
import os
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from unittest.mock import patch
from automacao.config import BASE_DIR, ARQUIVOS
from automacao.fonte_relatorios import FonteRelatorios, ErroEmail
from automacao.executar_automacao import executar_automacao
from preparar_demo import preparar_demo

ENV = {'AUTOMACAO_ORIGEM':'gmail','EMAIL_USUARIO':'pessoa.ficticia@gmail.com',
       'EMAIL_SENHA_APP':'abcd efgh ijkl mnop','EMAIL_ASSUNTO':'PORTFOLIO_DASHBOARD_DEMO'}

def mensagem(*, incompleta=False, duplicada=False, remetente='pessoa.ficticia@gmail.com', assunto='PORTFOLIO_DASHBOARD_DEMO'):
    msg=EmailMessage();msg['From']=remetente;msg['To']='pessoa.ficticia@gmail.com';msg['Subject']=assunto
    msg.set_content('Mensagem fictícia usada apenas em teste offline.')
    specs=list(ARQUIVOS.values())
    if incompleta:specs=specs[:-1]
    if duplicada:specs=specs+[specs[0]]
    for spec in specs:
        msg.add_attachment((BASE_DIR/'dados_demo'/spec['arquivo']).read_bytes(),maintype='application',subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',filename=spec['arquivo'])
    return msg.as_bytes()

class FakeIMAP:
    def __init__(self, mensagens):
        self.mensagens=mensagens;self.comandos=[];self.saiu=False
    def login(self, usuario, senha):
        self.usuario=usuario;self.senha=senha
        return 'OK',[]
    def select(self, pasta, readonly=False):
        self.pasta=pasta;self.readonly=readonly
        return 'OK',[b'1']
    def uid(self, comando, *args):
        self.comandos.append((comando,args))
        if comando=='search':return 'OK',[b' '.join(self.mensagens)]
        uid,opcao=args;conteudo=self.mensagens[uid]
        if opcao=='(RFC822.SIZE)':return 'OK',[b'1 (UID '+uid+b' RFC822.SIZE '+str(len(conteudo)).encode()+b')']
        if opcao=='(BODY.PEEK[])':return 'OK',[(b'1 (BODY[]',conteudo),b')']
        raise AssertionError('Operação IMAP inesperada')
    def logout(self):self.saiu=True

class GmailTests(TestCase):
    def setUp(self):
        self.tmp=TemporaryDirectory();self.raiz=preparar_demo(Path(self.tmp.name));self.fonte=None
    def tearDown(self):
        if self.fonte:self.fonte.encerrar()
        self.tmp.cleanup()
    def conectar(self, mail):
        self.fonte=FonteRelatorios(self.raiz/'entrada',self.raiz/'saida')
        with patch('automacao.fonte_relatorios.imaplib.IMAP4_SSL',return_value=mail):self.fonte.conectar()
        return self.fonte
    def test_local_nao_conecta(self):
        with patch.dict(os.environ,{'AUTOMACAO_ORIGEM':'local'},clear=True),patch('automacao.fonte_relatorios.imaplib.IMAP4_SSL') as factory:
            fonte=FonteRelatorios(self.raiz/'entrada',self.raiz/'saida');fonte.conectar();fonte.preparar();fonte.encerrar()
            factory.assert_not_called()
    def test_quatro_anexos_readonly_e_sem_espacos_na_senha(self):
        mail=FakeIMAP({b'1':mensagem()})
        with patch.dict(os.environ,ENV,clear=True):
            fonte=self.conectar(mail);fonte.preparar();fonte.encerrar()
        self.assertTrue(mail.readonly);self.assertEqual(mail.pasta,'INBOX');self.assertTrue(mail.saiu)
        self.assertEqual(mail.senha,'abcdefghijklmnop')
        self.assertTrue(any('(BODY.PEEK[])' in args for _,args in mail.comandos))
        self.assertTrue(all(comando in ('search','fetch') for comando,_ in mail.comandos))
        for spec in ARQUIVOS.values():
            self.assertEqual((self.raiz/'saida'/spec['arquivo']).read_bytes(),(BASE_DIR/'dados_demo'/spec['arquivo']).read_bytes())
    def test_mais_recente_incompleto_nao_recua_para_antigo(self):
        mail=FakeIMAP({b'1':mensagem(),b'2':mensagem(incompleta=True)})
        with patch.dict(os.environ,ENV,clear=True):
            fonte=self.conectar(mail)
            with self.assertRaisesRegex(ErroEmail,'incompleto'):fonte.preparar()
        self.assertFalse(any(args[0]==b'1' for comando,args in mail.comandos if comando=='fetch'))
    def test_assunto_deve_ser_exato(self):
        mail=FakeIMAP({b'1':mensagem(assunto='RE: PORTFOLIO_DASHBOARD_DEMO')})
        with patch.dict(os.environ,ENV,clear=True):
            fonte=self.conectar(mail)
            with self.assertRaisesRegex(ErroEmail,'Nenhum e-mail'):fonte.preparar()
    def test_remetente_diferente_e_ignorado(self):
        mail=FakeIMAP({b'1':mensagem(remetente='outra.pessoa@gmail.com')})
        with patch.dict(os.environ,ENV,clear=True):
            fonte=self.conectar(mail)
            with self.assertRaisesRegex(ErroEmail,'Nenhum e-mail'):fonte.preparar()
    def test_anexo_duplicado_interrompe(self):
        mail=FakeIMAP({b'1':mensagem(duplicada=True)})
        with patch.dict(os.environ,ENV,clear=True):
            fonte=self.conectar(mail)
            with self.assertRaisesRegex(ErroEmail,'duplicados'):fonte.preparar()
    def test_xlsx_invalido_e_recusado(self):
        with self.assertRaisesRegex(ErroEmail,'XLSX válida'):FonteRelatorios._validar_xlsx(b'nao-e-planilha')
    def test_senha_ausente_falha_antes_de_conectar(self):
        with patch.dict(os.environ,{'AUTOMACAO_ORIGEM':'gmail','EMAIL_USUARIO':'pessoa@gmail.com'},clear=True),patch('automacao.fonte_relatorios.imaplib.IMAP4_SSL') as factory:
            fonte=FonteRelatorios(self.raiz/'entrada',self.raiz/'saida')
            with self.assertRaisesRegex(ErroEmail,'senha de app'):fonte.conectar()
            factory.assert_not_called()
    def test_erro_login_nao_expoe_resposta_do_servidor(self):
        mail=FakeIMAP({});segredo='resposta-com-segredo-nao-deve-aparecer'
        with patch.dict(os.environ,ENV,clear=True),patch.object(mail,'login',side_effect=imaplib.IMAP4.error(segredo)):
            with self.assertRaises(ErroEmail) as erro:self.conectar(mail)
        self.assertNotIn(segredo,str(erro.exception))
    def test_integracao_com_motor_portavel(self):
        mail=FakeIMAP({b'3':mensagem()})
        with patch.dict(os.environ,ENV,clear=True),patch('automacao.fonte_relatorios.imaplib.IMAP4_SSL',return_value=mail):
            self.assertTrue(executar_automacao(motor='openpyxl',_raiz_demo=self.raiz))
        self.assertTrue(mail.saiu)
    def test_email_incompleto_nao_altera_dashboard(self):
        dashboard=self.raiz/'dashboard'/'dashboard_exemplo.xlsx';antes=hashlib.sha256(dashboard.read_bytes()).digest()
        mail=FakeIMAP({b'3':mensagem(incompleta=True)})
        with patch.dict(os.environ,ENV,clear=True),patch('automacao.fonte_relatorios.imaplib.IMAP4_SSL',return_value=mail):
            self.assertFalse(executar_automacao(motor='openpyxl',_raiz_demo=self.raiz))
        self.assertEqual(hashlib.sha256(dashboard.read_bytes()).digest(),antes)
        self.assertTrue(mail.saiu)
