"""Impede misturar planilhas antigas/reais com os arquivos da demonstração."""
from pathlib import Path
from automacao.config import ARQUIVOS

def verificar_planilhas(raiz, temporario=None):
    raiz = Path(raiz)
    nomes = {spec['arquivo'] for spec in ARQUIVOS.values()}
    permitidos = {'entrada': nomes, 'saida': nomes, 'dashboard': {'dashboard_exemplo.xlsx'}}
    temp = Path(temporario).absolute() if temporario else None
    extras = []
    for pasta, esperados in permitidos.items():
        diretorio = raiz / pasta
        if diretorio.is_symlink():
            raise ValueError(f'A pasta {pasta} não pode ser um link simbólico.')
        if not diretorio.exists():
            continue
        for arquivo in diretorio.iterdir():
            if temp is not None and arquivo.absolute() == temp:
                continue
            if arquivo.suffix.lower() in {'.xlsx', '.xlsb', '.xlsm', '.xls'} and arquivo.name not in esperados:
                extras.append(str(arquivo.relative_to(raiz)))
    if extras:
        raise ValueError('Planilhas extras no ambiente demonstrativo: ' + ', '.join(sorted(extras)) +
                         '. Mova-as para outra pasta e use apenas os modelos fictícios. Nenhum arquivo extra foi apagado.')
