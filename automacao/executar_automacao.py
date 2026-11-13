"""Entrada única: win32com é o principal; openpyxl é uma alternativa explícita."""
import argparse
import logging
import os
from pathlib import Path
import sys

if __package__ in (None, ''):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv
from automacao.config import BASE_DIR, PASTA_DEMO, PASTA_CONFIGURADA

load_dotenv(BASE_DIR / '.env', override=False)
MOTORES = ('win32com', 'openpyxl')

def executar_automacao(caminho_dashboard=None, pasta_entrada=None, pasta_saida=None,
                      modo_teste=None, com_inicializado=False, progresso_callback=None,
                      *, motor=None, _raiz_demo=None):
    """Mantém o contrato das views. O motor nunca é trocado silenciosamente."""
    if modo_teste is None:
        modo_teste = os.getenv('AUTOMACAO_MODO_TESTE', '1') == '1'
    escolhido = motor if motor is not None else os.getenv('AUTOMACAO_MOTOR', 'win32com')
    if escolhido not in MOTORES:
        logging.getLogger('automacao').error('AUTOMACAO_MOTOR inválido: %r', escolhido)
        return False
    if escolhido == 'win32com':
        from automacao.motor_win32com import executar_automacao as executar
    else:
        from automacao.motor_openpyxl import executar_automacao as executar
    return executar(caminho_dashboard, pasta_entrada, pasta_saida,
                    modo_teste=modo_teste, com_inicializado=com_inicializado,
                    progresso_callback=progresso_callback,
                    _raiz_demo=Path(_raiz_demo) if _raiz_demo is not None else (PASTA_DEMO if modo_teste else PASTA_CONFIGURADA))

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--motor', choices=MOTORES, help='Sobrepõe AUTOMACAO_MOTOR apenas nesta execução.')
    args = parser.parse_args()
    from dashboard.bloqueio_execucao import bloquear_execucao
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'projeto_web.settings')
    try:
        with bloquear_execucao():
            from django.conf import settings
            resultado = executar_automacao(motor=args.motor, modo_teste=settings.AUTOMACAO_MODO_TESTE)
    except Exception as erro:
        print(f'Não foi possível iniciar: {erro}', file=sys.stderr)
        resultado = False
    return 0 if resultado else 1

if __name__ == '__main__':
    raise SystemExit(main())
