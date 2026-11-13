"""Executor público: 17 etapas com origem local ou Gmail opcional.

A biblioteca openpyxl implementa a escrita da aplicação Python. As planilhas
fictícias distribuídas em dados_demo são os modelos imutáveis desta versão.
"""
import argparse
from copy import copy
import logging
import math
import os
from pathlib import Path
import shutil
import sys
import tempfile

# Permite executar tanto como módulo quanto pelo caminho do arquivo.
if __package__ in (None, ''):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from automacao.config import BASE_DIR, PASTA_DEMO, PASTA_CONFIGURADA, ARQUIVOS

from automacao.fonte_relatorios import FonteRelatorios

LOGGER = logging.getLogger('automacao')

def configurar_log():
    pasta = BASE_DIR / 'logs'
    pasta.mkdir(exist_ok=True)
    LOGGER.setLevel(logging.INFO)
    if not LOGGER.handlers:
        formato = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
        for handler in [logging.StreamHandler(), logging.FileHandler(pasta/'automacao.log', encoding='utf-8')]:
            handler.setFormatter(formato)
            LOGGER.addHandler(handler)

def validar_caminhos(dashboard, entrada, saida, raiz, temporario=None):
    """O motor só modifica a árvore demonstrativa, inclusive via interface."""
    raiz = Path(raiz).absolute()
    if raiz.is_symlink() or raiz.resolve() != raiz:
        raise ValueError('Raiz de demonstração não pode ser um link simbólico.')
    esperados = [raiz/'dashboard'/'dashboard_exemplo.xlsx', raiz/'entrada', raiz/'saida']
    recebidos = [Path(dashboard).absolute(), Path(entrada).absolute(), Path(saida).absolute()]
    for atual, esperado in zip(recebidos, esperados):
        if atual != esperado or atual.is_symlink() or atual.resolve() != esperado:
            raise ValueError('Caminho fora do ambiente demonstrativo permitido.')
    if not recebidos[0].is_file() or not all(p.is_dir() for p in recebidos[1:]):
        raise ValueError('Execute python preparar_demo.py antes da automação.')
    for spec in ARQUIVOS.values():
        for pasta in recebidos[1:]:
            f = pasta / spec['arquivo']
            if f.is_symlink() or f.resolve().parent != pasta:
                raise ValueError('Arquivo de demonstração não pode redirecionar a outro caminho.')
    from automacao.seguranca_demo import verificar_planilhas
    verificar_planilhas(raiz, temporario=temporario)
    return recebidos

def preparar_dataframe(df, spec):
    if df.shape[1] != spec['largura']:
        raise ValueError(f"Esperadas {spec['largura']} colunas; recebidas {df.shape[1]}.")
    df = df.dropna(how='all')
    total = df.iloc[:,0].astype(str).str.strip().str.upper().eq('TOTAL')
    df = df.loc[~total].iloc[:,spec['selecionar']].copy().reset_index(drop=True)
    if df.empty:
        raise ValueError('Relatório sem registros.')
    # Identificadores fictícios são texto, preservando zeros à esquerda.
    if df.iloc[:,0].isna().any():
        raise ValueError('Identificador ausente.')
    df.iloc[:,0] = df.iloc[:,0].astype(str).str.strip()
    if df.iloc[:,0].duplicated().any() or df.iloc[:,0].eq('').any():
        raise ValueError('Identificadores vazios ou duplicados.')
    for indice in spec['valores']:
        valores = pd.to_numeric(df.iloc[:,indice], errors='raise')
        if valores.isna().any() or not all(math.isfinite(float(v)) for v in valores):
            raise ValueError('Valor monetário ausente ou inválido.')
        df[df.columns[indice]] = valores
    return df.astype(object).where(pd.notna(df), None)

def executar_automacao(caminho_dashboard=None, pasta_entrada=None, pasta_saida=None,
                      modo_teste=True, com_inicializado=False, progresso_callback=None, *, _raiz_demo=None):
    """Retorna bool para compatibilidade com as views e o worker Django.

    _raiz_demo permite isolamento em testes. A CLI e as views usam a raiz fixa.
    A origem local/Gmail é compartilhada; este motor não usa COM.
    """
    configurar_log()
    raiz = Path(_raiz_demo) if _raiz_demo is not None else (PASTA_DEMO if modo_teste else PASTA_CONFIGURADA)
    dashboard = Path(caminho_dashboard) if caminho_dashboard else raiz/'dashboard'/'dashboard_exemplo.xlsx'
    entrada = Path(pasta_entrada) if pasta_entrada else raiz/'entrada'
    saida = Path(pasta_saida) if pasta_saida else raiz/'saida'
    wb = None
    fonte = None
    temporario = None
    sucesso = False
    def etapa(n, texto):
        LOGGER.info('[ETAPA %02d] %s', n, texto)
        if progresso_callback:
            try:
                progresso_callback(min(99, int(n/17*99)), texto)
            except Exception:
                LOGGER.exception('Falha ao registrar progresso')
    try:
        etapa(1, 'Preparar origem dos relatórios (local ou Gmail)')
        dashboard, entrada, saida = validar_caminhos(dashboard, entrada, saida, raiz)
        fonte = FonteRelatorios(entrada, saida)
        fonte.conectar()
        etapa(2, 'Obter os quatro relatórios da origem selecionada e preparar a saída')
        fonte.preparar()
        etapa(3, 'Ler os quatro relatórios com pandas')
        brutos = {nome: pd.read_excel(saida/spec['arquivo'], skiprows=spec['pular'], header=None, dtype=object)
                  for nome,spec in ARQUIVOS.items()}
        etapa(4, 'Preparar DataFrames, selecionar colunas e remover TOTAL')
        dados = {nome: preparar_dataframe(brutos[nome],spec) for nome,spec in ARQUIVOS.items()}
        etapa(5, 'Validar quantidades, identificadores e valores')
        for nome,spec in ARQUIVOS.items():
            if len(dados[nome].columns) != len(spec['destinos']):
                raise ValueError(f'Contrato de colunas inválido: {nome}')
        etapa(6, 'Abrir Dashboard fictício com openpyxl')
        wb = load_workbook(dashboard)
        if any(nome not in wb.sheetnames for nome in ARQUIVOS):
            raise ValueError('Dashboard não contém todas as abas esperadas.')
        etapa(7, 'Localizar TOTAL e a última linha de cada aba')
        atuais = {}
        for nome,spec in ARQUIVOS.items():
            ws = wb[nome];inicio = spec['linha'];coluna = spec['destinos'][0]
            totais = [r for r in range(inicio,ws.max_row+1) if ws.cell(r,coluna).value == 'TOTAL']
            if spec['total']:
                if len(totais) != 1:
                    raise ValueError(f'Esperada uma linha TOTAL em {nome}.')
                atuais[nome] = totais[0]-inicio
            else:
                rows = [r for r in range(inicio,ws.max_row+1) if ws.cell(r,coluna).value is not None]
                atuais[nome] = max(rows,default=inicio-1)-inicio+1
        etapa(8, 'Calcular diferenças de quantidade e posições finais')
        for nome in ARQUIVOS:
            LOGGER.info('%s: %s antigos, %s novos, diferença %+d',nome,atuais[nome],len(dados[nome]),len(dados[nome])-atuais[nome])
        etapa(9, 'Ajustar a área de registros')
        for nome,spec in ARQUIVOS.items():
            ws=wb[nome];inicio=spec['linha'];antigos=atuais[nome];novos=len(dados[nome]);dif=novos-antigos
            if dif>0:
                ws.insert_rows(inicio+antigos,dif)
            elif dif<0:
                ws.delete_rows(inicio+novos,-dif)
            # Reaplicar o estilo da primeira linha aos registros adicionados.
            for r in range(inicio,inicio+novos):
                for c in range(1,max(spec['destinos'])+1):
                    ws.cell(r,c)._style=copy(ws.cell(inicio,c)._style)
                if nome=='RECEBIMENTOS':
                    for c in range(16,22):
                        cel=ws.cell(r,c)
                        if cel.value is None:
                            origem=ws.cell(inicio,c)
                            cel.value=Translator(origem.value,origin=origem.coordinate).translate_formula(cel.coordinate) if origem.data_type=='f' else origem.value
        etapa(10, 'Limpar os dados anteriores nas colunas de importação')
        for nome,spec in ARQUIVOS.items():
            for r in range(spec['linha'],spec['linha']+len(dados[nome])):
                for c in spec['destinos']:
                    wb[nome].cell(r,c).value=None
        etapa(11, 'Escrever os registros e preservar P:U de RECEBIMENTOS')
        for nome,spec in ARQUIVOS.items():
            for r,valores in enumerate(dados[nome].itertuples(index=False,name=None),spec['linha']):
                for c,valor in zip(spec['destinos'],valores):
                    cell=wb[nome].cell(r,c);cell.value=valor
                    if isinstance(valor,str):
                        cell.data_type='s' # Entradas são dados, nunca fórmulas executáveis.
        etapa(12, 'Reconstruir as linhas TOTAL')
        for nome,spec in ARQUIVOS.items():
            ws=wb[nome];final=spec['linha']+len(dados[nome])-1
            if spec['total']:
                for c in range(1,max(spec['destinos'])+1):
                    ws.cell(final+1,c).value=None
                ws.cell(final+1,spec['destinos'][0],'TOTAL')
                for indice in spec['valores']:
                    c=spec['destinos'][indice];letra=get_column_letter(c)
                    ws.cell(final+1,c,f'=SUM({letra}{spec["linha"]}:{letra}{final})')
                for cell in ws[final+1]:
                    cell.font=Font(name='Arial',bold=True,color='FFFFFF')
                    cell.fill=PatternFill('solid',fgColor='243B53')
        etapa(13, 'Atualizar fórmulas SUBTOTAL e quadro de resumo')
        for n,(nome,spec) in enumerate(ARQUIVOS.items(),5):
            ws=wb[nome];final=spec['linha']+len(dados[nome])-1
            for indice in spec['valores']:
                c=spec['destinos'][indice];letra=get_column_letter(c)
                ws.cell(5,c,f'=SUBTOTAL(9,{letra}{spec["linha"]}:{letra}{final})')
            wb['RESUMO'].cell(n,2,len(dados[nome]))
            ultimo=get_column_letter(spec['destinos'][spec['valores'][-1]])
            wb['RESUMO'].cell(n,3,f"='{nome}'!{ultimo}5")
        wb.calculation=CalcProperties(calcId=191029,fullCalcOnLoad=True)
        LOGGER.info('[AUDITORIA] Conferir registros e somas antes de salvar')
        for nome,spec in ARQUIVOS.items():
            ws=wb[nome]
            for indice in spec['valores']:
                esperado=sum(float(v) for v in dados[nome].iloc[:,indice])
                obtido=sum(float(ws.cell(r,spec['destinos'][indice]).value) for r in range(spec['linha'],spec['linha']+len(dados[nome])))
                if not math.isclose(esperado,obtido,rel_tol=1e-9,abs_tol=0.005):
                    raise ValueError(f'Divergência de valores em {nome}')
            for r,valores in enumerate(dados[nome].itertuples(index=False,name=None),spec['linha']):
                if any(ws.cell(r,c).value != valor for c,valor in zip(spec['destinos'],valores)):
                    raise ValueError(f'Divergência de registros em {nome}')
        etapa(14, 'Salvar Dashboard por substituição atômica')
        validar_caminhos(dashboard,entrada,saida,raiz)
        fd,temporario=tempfile.mkstemp(suffix='.xlsx',dir=dashboard.parent)
        os.close(fd)
        wb.save(temporario)
        os.replace(temporario,dashboard)
        temporario=None
        sucesso=True
    except Exception:
        LOGGER.exception('Execução interrompida; confira os arquivos e o log.')
    finally:
        etapa(15, 'Fechar o workbook')
        if wb is not None:
            wb.close()
        if temporario:
            Path(temporario).unlink(missing_ok=True)
        etapa(16, 'Liberar o motor local (Excel COM não utilizado)')
        etapa(17, 'Encerrar a origem dos relatórios e a sessão IMAP, se utilizada')
        if fonte is not None:
            fonte.encerrar()
    if sucesso and progresso_callback:
        try:
            progresso_callback(100,'Concluído')
        except Exception:
            LOGGER.exception('Falha ao registrar conclusão')
    LOGGER.info('[FINAL] %s','Sucesso' if sucesso else 'Falha')
    return sucesso

