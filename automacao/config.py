"""Contrato dos quatro relatórios fictícios; índices de colunas são baseados em zero."""
from pathlib import Path
BASE_DIR = Path(__file__).resolve().parent.parent
PASTA_DEMO = BASE_DIR / 'TESTE_AUTOMACAO'
PASTA_CONFIGURADA = BASE_DIR / 'DEMO_CONFIGURADA'
ARQUIVOS = {
    'VENDAS': dict(arquivo='vendas_exemplo.xlsx', pular=3, largura=22,
                    selecionar=list(range(22)), linha=8, destinos=list(range(3,25)), valores=list(range(13,22)), total=True),
    'RECEBIMENTOS': dict(arquivo='recebimentos_exemplo.xlsx', pular=3, largura=28,
                        selecionar=list(range(15))+list(range(21,28)), linha=8,
                        destinos=list(range(1,16))+list(range(22,29)), valores=list(range(15,22)), total=True),
    'CUSTOS': dict(arquivo='custos_exemplo.xlsx', pular=3, largura=16,
                  selecionar=list(range(16)), linha=8, destinos=list(range(3,19)), valores=list(range(11,16)), total=False),
    'DESPESAS': dict(arquivo='despesas_exemplo.xlsx', pular=3, largura=15,
                selecionar=list(range(12))+[13,14], linha=8, destinos=list(range(1,15)), valores=[11,12,13], total=False),
}
