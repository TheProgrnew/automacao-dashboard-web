"""Motor principal da demonstração: controla uma instância própria do Excel.

Importações de pywin32 são tardias para permitir usar openpyxl em outros SOs.
O Dashboard é atualizado numa cópia temporária. O original só é substituído
após salvar e encerrar o Excel com sucesso. A origem local/Gmail é compartilhada com o motor alternativo.
"""
import logging
import math
import os
from pathlib import Path
import shutil
import tempfile

import pandas as pd
from openpyxl.utils import get_column_letter
from automacao.config import BASE_DIR, PASTA_DEMO, PASTA_CONFIGURADA, ARQUIVOS
from automacao.motor_openpyxl import configurar_log, preparar_dataframe, validar_caminhos

from automacao.fonte_relatorios import FonteRelatorios

LOGGER = logging.getLogger('automacao')
XL_UP = -4162
XL_PASTE_FORMATS = -4122
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105

def carregar_com():
    if os.name != 'nt':
        raise RuntimeError('win32com exige Windows e Excel instalado. Para a alternativa, use --motor openpyxl.')
    try:
        import pythoncom
        import win32com.client
    except ImportError as erro:
        raise RuntimeError('Instale pywin32 no Python da sua .venv: python -m pip install -r requirements.txt') from erro
    return pythoncom, win32com.client

def intervalo(ws, linha, primeira_coluna, ultima_linha, ultima_coluna):
    return ws.Range(ws.Cells(linha, primeira_coluna), ws.Cells(ultima_linha, ultima_coluna))

def escrever_bloco(ws, linha, colunas, linhas):
    """Escreve somente colunas contíguas; LIQ/PAGTO chama em dois blocos."""
    if not linhas:
        return
    fim = linha + len(linhas) - 1
    for i, coluna in enumerate(colunas):
        if all(isinstance(valores[i], str) for valores in linhas):
            intervalo(ws, linha, coluna, fim, coluna).NumberFormat = '@'
    # Força texto literal para valores iniciados como expressão do Excel.
    dados = tuple(tuple("'" + v if isinstance(v, str) and v.startswith(('=', '+', '-', '@')) else v
                        for v in valores) for valores in linhas)
    intervalo(ws, linha, colunas[0], fim, colunas[-1]).Value2 = dados

def executar_automacao(caminho_dashboard=None, pasta_entrada=None, pasta_saida=None,
                      modo_teste=True, com_inicializado=False, progresso_callback=None,
                      *, _raiz_demo=None):
    configurar_log()
    raiz = Path(_raiz_demo) if _raiz_demo is not None else (PASTA_DEMO if modo_teste else PASTA_CONFIGURADA)
    dashboard = Path(caminho_dashboard) if caminho_dashboard else raiz/'dashboard'/'dashboard_exemplo.xlsx'
    entrada = Path(pasta_entrada) if pasta_entrada else raiz/'entrada'
    saida = Path(pasta_saida) if pasta_saida else raiz/'saida'
    pythoncom = None
    com_proprio = False
    excel = None
    wb = None
    fonte = None
    ws = resumo = faixa = modelo = None
    temporario = None
    salvo = False
    encerramento_ok = True

    def etapa(numero, texto):
        LOGGER.info('[ETAPA %02d] %s', numero, texto)
        if progresso_callback:
            try:
                progresso_callback(min(99, int(numero / 17 * 99)), texto)
            except Exception:
                LOGGER.exception('Falha ao registrar progresso')
    try:
        # Falha antes de copiar relatórios se o sistema não suportar COM.
        pythoncom, cliente = carregar_com()
        etapa(1, 'Preparar origem dos relatórios (local ou Gmail)')
        dashboard, entrada, saida = validar_caminhos(dashboard, entrada, saida, raiz, temporario=temporario)
        fonte = FonteRelatorios(entrada, saida)
        fonte.conectar()
        etapa(2, 'Obter os quatro relatórios da origem selecionada e preparar a saída')
        fonte.preparar()
        etapa(3, 'Ler os quatro relatórios com pandas')
        brutos = {nome: pd.read_excel(saida/spec['arquivo'], skiprows=spec['pular'], header=None, dtype=object)
                  for nome, spec in ARQUIVOS.items()}
        etapa(4, 'Preparar DataFrames, selecionar colunas e remover TOTAL')
        dados = {nome: preparar_dataframe(brutos[nome], spec) for nome, spec in ARQUIVOS.items()}
        etapa(5, 'Validar quantidades, identificadores e valores')
        for nome, spec in ARQUIVOS.items():
            if len(dados[nome].columns) != len(spec['destinos']):
                raise ValueError(f'Contrato de colunas inválido: {nome}')
        etapa(6, 'Abrir cópia do Dashboard no Excel com win32com')
        fd, caminho = tempfile.mkstemp(prefix='demo_excel_', suffix='.xlsx', dir=dashboard.parent)
        os.close(fd)
        temporario = Path(caminho)
        shutil.copyfile(dashboard, temporario)
        if not com_inicializado:
            pythoncom.CoInitialize()
            com_proprio = True
        excel = cliente.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        excel.AutomationSecurity = 3
        wb = excel.Workbooks.Open(str(temporario.resolve()), UpdateLinks=0, ReadOnly=False,
                                  IgnoreReadOnlyRecommended=True, AddToMru=False)
        if wb.ReadOnly:
            raise ValueError('Excel abriu o arquivo em modo somente leitura.')
        if Path(wb.FullName).resolve() != temporario.resolve():
            raise ValueError('Excel abriu um arquivo diferente da cópia de demonstração.')
        excel.Calculation = XL_CALCULATION_MANUAL
        nomes = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
        if any(nome not in nomes for nome in [*ARQUIVOS, 'RESUMO']):
            raise ValueError('Dashboard não contém todas as abas esperadas.')
        etapa(7, 'Localizar TOTAL e a última linha de cada aba')
        atuais = {}
        for nome, spec in ARQUIVOS.items():
            ws = wb.Worksheets(nome)
            inicio, coluna = spec['linha'], spec['destinos'][0]
            ultima = ws.Cells(ws.Rows.Count, coluna).End(XL_UP).Row
            totais = [r for r in range(inicio, ultima + 1) if ws.Cells(r, coluna).Value2 == 'TOTAL']
            if spec['total']:
                if len(totais) != 1:
                    raise ValueError(f'Esperada uma linha TOTAL em {nome}.')
                atuais[nome] = totais[0] - inicio
            else:
                atuais[nome] = max(0, ultima - inicio + 1)
        etapa(8, 'Calcular diferenças de quantidade e posições finais')
        for nome in ARQUIVOS:
            LOGGER.info('%s: %s antigos, %s novos, diferença %+d', nome, atuais[nome],
                        len(dados[nome]), len(dados[nome]) - atuais[nome])
        etapa(9, 'Ajustar linhas pelo Excel e copiar a formatação')
        for nome, spec in ARQUIVOS.items():
            ws = wb.Worksheets(nome)
            inicio, antigos, novos = spec['linha'], atuais[nome], len(dados[nome])
            diferenca = novos - antigos
            if diferenca > 0:
                ws.Rows(f'{inicio+antigos}:{inicio+novos-1}').Insert()
            elif diferenca < 0:
                ws.Rows(f'{inicio+novos}:{inicio+antigos-1}').Delete()
            if novos > 1:
                modelo = intervalo(ws, inicio, 1, inicio, max(spec['destinos']))
                modelo.Copy()
                intervalo(ws, inicio+1, 1, inicio+novos-1, max(spec['destinos'])).PasteSpecial(Paste=XL_PASTE_FORMATS)
            if nome == 'RECEBIMENTOS' and diferenca > 0:
                # Copia somente as células reservadas das novas linhas.
                # O Excel ajusta referências relativas de fórmulas na cópia.
                modelo = intervalo(ws, inicio, 16, inicio, 21)
                for r in range(inicio+antigos, inicio+novos):
                    modelo.Copy(Destination=intervalo(ws, r, 16, r, 21))
        excel.CutCopyMode = False
        etapa(10, 'Limpar somente as colunas de importação')
        for nome, spec in ARQUIVOS.items():
            ws = wb.Worksheets(nome)
            final = spec['linha'] + len(dados[nome]) - 1
            blocos = [(1,15),(22,28)] if nome == 'RECEBIMENTOS' else [(spec['destinos'][0],spec['destinos'][-1])]
            for c1, c2 in blocos:
                intervalo(ws, spec['linha'], c1, final, c2).ClearContents()
        etapa(11, 'Escrever blocos de dados pelo Excel, preservando P:U')
        for nome, spec in ARQUIVOS.items():
            ws = wb.Worksheets(nome)
            linhas = list(dados[nome].itertuples(index=False, name=None))
            if nome == 'RECEBIMENTOS':
                escrever_bloco(ws, spec['linha'], list(range(1,16)), [r[:15] for r in linhas])
                escrever_bloco(ws, spec['linha'], list(range(22,29)), [r[15:] for r in linhas])
            else:
                escrever_bloco(ws, spec['linha'], spec['destinos'], linhas)
            for indice in spec['valores']:
                coluna = spec['destinos'][indice]
                intervalo(ws, spec['linha'], coluna, spec['linha']+len(linhas)-1, coluna).NumberFormat = '"R$" #,##0.00'
        etapa(12, 'Reconstruir as linhas TOTAL no Excel')
        for nome, spec in ARQUIVOS.items():
            if not spec['total']:
                continue
            ws = wb.Worksheets(nome)
            final = spec['linha'] + len(dados[nome]) - 1
            faixa = intervalo(ws, final+1, 1, final+1, max(spec['destinos']))
            faixa.ClearContents()
            ws.Cells(final+1, spec['destinos'][0]).Value2 = 'TOTAL'
            for indice in spec['valores']:
                coluna = spec['destinos'][indice]
                letra = get_column_letter(coluna)
                ws.Cells(final+1, coluna).Formula = f'=SUM({letra}{spec["linha"]}:{letra}{final})'
            faixa.Font.Bold = True
            faixa.Font.Color = 16777215
            faixa.Interior.Color = 5454628  # RGB(36,59,83)
        etapa(13, 'Atualizar fórmulas e recalcular com o Excel')
        resumo = wb.Worksheets('RESUMO')
        for n, (nome, spec) in enumerate(ARQUIVOS.items(), 5):
            ws = wb.Worksheets(nome)
            final = spec['linha'] + len(dados[nome]) - 1
            for indice in spec['valores']:
                coluna = spec['destinos'][indice]
                letra = get_column_letter(coluna)
                ws.Cells(5, coluna).Formula = f'=SUBTOTAL(9,{letra}{spec["linha"]}:{letra}{final})'
            resumo.Cells(n, 2).Value2 = len(dados[nome])
            letra = get_column_letter(spec['destinos'][spec['valores'][-1]])
            resumo.Cells(n, 3).Formula = f"='{nome}'!{letra}5"
        excel.Calculation = XL_CALCULATION_AUTOMATIC
        excel.CalculateFullRebuild()
        LOGGER.info('[AUDITORIA] Conferir dados escritos e resultados calculados pelo Excel')
        for n, (nome, spec) in enumerate(ARQUIVOS.items(), 5):
            ws = wb.Worksheets(nome)
            for r, valores in enumerate(dados[nome].itertuples(index=False, name=None), spec['linha']):
                if any(ws.Cells(r,c).Value2 != valor for c,valor in zip(spec['destinos'],valores)):
                    raise ValueError(f'Divergência de registros em {nome}')
            for indice in spec['valores']:
                esperado = sum(float(v) for v in dados[nome].iloc[:,indice])
                obtido = float(ws.Cells(5, spec['destinos'][indice]).Value2)
                if not math.isclose(esperado, obtido, rel_tol=1e-9, abs_tol=0.005):
                    raise ValueError(f'Divergência de subtotal calculado em {nome}')
                if spec['total']:
                    valor_total = float(ws.Cells(spec['linha']+len(dados[nome]), spec['destinos'][indice]).Value2)
                    if not math.isclose(esperado, valor_total, rel_tol=1e-9, abs_tol=0.005):
                        raise ValueError(f'Divergência de TOTAL calculado em {nome}')
            esperado = sum(float(v) for v in dados[nome].iloc[:,spec['valores'][-1]])
            if not math.isclose(float(resumo.Cells(n,3).Value2), esperado, rel_tol=1e-9, abs_tol=0.005):
                raise ValueError(f'Divergência do resumo em {nome}')
        etapa(14, 'Salvar a cópia temporária pelo Excel')
        validar_caminhos(dashboard, entrada, saida, raiz, temporario=temporario)
        wb.Save()
        salvo = True
    except Exception:
        LOGGER.exception('Execução win32com interrompida; Dashboard original preservado.')
    finally:
        ws = resumo = faixa = modelo = None
        etapa(15, 'Fechar o workbook do Excel')
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                encerramento_ok = False
                LOGGER.exception('Falha ao fechar o workbook')
            finally:
                wb = None
        etapa(16, 'Encerrar a instância própria do Excel e liberar COM')
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                encerramento_ok = False
                LOGGER.exception('Falha ao encerrar a instância do Excel')
            finally:
                excel = None
        if pythoncom is not None and com_proprio:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                encerramento_ok = False
                LOGGER.exception('Falha ao liberar COM')
        etapa(17, 'Encerrar a origem dos relatórios e a sessão IMAP, se utilizada')
        if fonte is not None:
            fonte.encerrar()
    sucesso = False
    try:
        if salvo and encerramento_ok:
            validar_caminhos(dashboard, entrada, saida, raiz, temporario=temporario)
            os.replace(temporario, dashboard)
            temporario = None
            sucesso = True
    except Exception:
        LOGGER.exception('Falha ao substituir Dashboard; cópia anterior preservada.')
    finally:
        if temporario is not None:
            try:
                temporario.unlink(missing_ok=True)
            except OSError:
                LOGGER.exception('Não foi possível remover cópia temporária: %s', temporario)
    if sucesso and progresso_callback:
        try:
            progresso_callback(100, 'Concluído')
        except Exception:
            LOGGER.exception('Falha ao registrar conclusão')
    LOGGER.info('[FINAL win32com] %s', 'Sucesso' if sucesso else 'Falha')
    return sucesso
