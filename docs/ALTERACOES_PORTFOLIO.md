# Notas da versão demonstrativa

## Revisão de 13/11/2026

O projeto utiliza um cenário didático de automação de relatórios para finalidades específicas. Os exemplos apresentam vendas, recebimentos, custos e despesas de uma empresa fictícia, com nomes, códigos e valores inventados.

## Arquivos e contratos

- `vendas_exemplo.xlsx`: seis pedidos de venda fictícios.
- `recebimentos_exemplo.xlsx`: oito registros de recebimentos fictícios.
- `custos_exemplo.xlsx`: cinco registros de custos fictícios.
- `despesas_exemplo.xlsx`: quatro registros de despesas fictícias.
- `dashboard_exemplo.xlsx`: painel inicial com RESUMO, VENDAS, RECEBIMENTOS, CUSTOS e DESPESAS.

As fontes têm cabeçalho na linha 3 e registros a partir da linha 4. As quatro áreas do Dashboard começam na linha 8. As colunas P:U de RECEBIMENTOS são reservadas a controles internos do exemplo. Os motores, a origem Gmail, os testes e a documentação usam os novos nomes e contratos.

## Execução e configuração

O motor principal é `win32com`, para Windows com Excel desktop. `openpyxl` é a alternativa explícita sem Excel. A origem dos relatórios pode ser local ou Gmail pessoal; a integração Gmail usa senha de app e apenas os quatro anexos esperados.

`AUTOMACAO_MODO_TESTE=1` seleciona TESTE_AUTOMACAO. O valor `0` seleciona DEMO_CONFIGURADA e permite salvar os caminhos pela interface. Ambos os ambientes trabalham com modelos de exemplo.

Extraia em uma pasta nova para evitar misturar relatórios de versões diferentes. O preparo recusa planilhas extras e não as apaga automaticamente. As instruções de instalação, servidores, e-mail, logs e Git estão no README.

## Documentação e limites

A documentação técnica descreve a arquitetura, os contratos, as rotas, os dados, os testes e a operação da distribuição. A origem Gmail é testada com servidor simulado durante a preparação; a conta real deve ser validada por quem executar. A integração nativa com Excel também exige conferência no Windows.

A licença continua sem definição. Esta revisão não adiciona permissões de uso nem publica um repositório remoto.

Resultado da revisão: 96 testes encontrados, 95 aprovados e 1 teste nativo de Excel ignorado neste ambiente Linux. As fórmulas do resumo foram conferidas após uma execução completa do motor portável.
