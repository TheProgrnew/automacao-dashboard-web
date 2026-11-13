"""Copia somente os cinco modelos fictícios para o ambiente de execução."""
import argparse
from pathlib import Path
import shutil
from automacao.config import BASE_DIR, PASTA_DEMO, PASTA_CONFIGURADA, ARQUIVOS

def preparar_demo(destino=PASTA_DEMO, restaurar=False):
    destino = Path(destino)
    from automacao.seguranca_demo import verificar_planilhas
    verificar_planilhas(destino)
    for pasta in ['entrada', 'saida', 'dashboard']:
        (destino / pasta).mkdir(parents=True, exist_ok=True)
    modelos = BASE_DIR / 'dados_demo'
    for spec in ARQUIVOS.values():
        nome = spec['arquivo']
        alvo = destino / 'entrada' / nome
        if restaurar or not alvo.exists():
            shutil.copyfile(modelos / nome, alvo)
    alvo = destino / 'dashboard' / 'dashboard_exemplo.xlsx'
    if restaurar or not alvo.exists():
        shutil.copyfile(modelos / alvo.name, alvo)
    if restaurar:
        for spec in ARQUIVOS.values():
            (destino / 'saida' / spec['arquivo']).unlink(missing_ok=True)
    return destino

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--restaurar', action='store_true', help='Substitui somente os arquivos conhecidos da demonstração por seus modelos originais.')
    parser.add_argument('--ambiente', choices=['teste', 'configurado'], default='teste', help='Teste: TESTE_AUTOMACAO. Configurado: DEMO_CONFIGURADA.')
    args = parser.parse_args()
    destino = PASTA_DEMO if args.ambiente == 'teste' else PASTA_CONFIGURADA
    try:
        print(f'Ambiente preparado: {preparar_demo(destino, restaurar=args.restaurar)}')
    except ValueError as erro:
        parser.exit(1, f'{erro}\n')
