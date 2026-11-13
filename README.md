# Automação de Dashboard com Python e Django

**Sistema web para automatização, consolidação e atualização de relatórios Excel.**

Projeto de portfólio desenvolvido com Python e Django para demonstrar a automação de processos administrativos, o tratamento de dados e a atualização de dashboards em planilhas eletrônicas. A aplicação oferece uma interface web autenticada para iniciar atualizações, acompanhar o progresso, consultar o histórico e baixar o Dashboard atualizado.

## Principais funcionalidades

- **Automação de relatórios:** leitura, preparação, validação e consolidação de dados com pandas.
- **Integração com Excel:** motor principal `win32com` (Windows com Excel instalado) e alternativa explícita `openpyxl`.
- **Interface web:** execução, acompanhamento de progresso, histórico e download por meio do Django.
- **Controle de acesso:** autenticação, cadastro, aprovação de usuários e permissões para execução.
- **Agendamento:** preferências de execução programada, sujeitas às condições de configuração e ao modo de operação; salvar preferências não instala uma tarefa no Windows.
- **Origem dos relatórios:** arquivos locais por padrão ou leitura opcional de anexos de Gmail pessoal por IMAP/TLS.
- **Demonstração segura:** quatro relatórios e um Dashboard `.xlsx` com dados inteiramente fictícios.

## Tecnologias utilizadas

| Tecnologia | Utilização |
| --- | --- |
| Python | Linguagem principal e execução da automação |
| Django 5.2 | Interface web, autenticação e histórico |
| pandas | Leitura, tratamento e validação dos relatórios |
| pywin32 (`win32com`) | Atualização e recálculo pelo Microsoft Excel no Windows |
| openpyxl | Alternativa de atualização de planilhas sem Excel instalado |
| SQLite | Banco de dados local da aplicação |
| HTML e CSS | Interface do usuário |
| Waitress | Servidor WSGI para demonstração local |
| Git e GitHub | Versionamento e disponibilização do código |

## Sobre esta versão de demonstração

Este repositório apresenta um cenário didático de uma **empresa fictícia**, com quatro relatórios: **vendas, recebimentos, custos e despesas**. Os arquivos estão na pasta [`dados_demo/`](dados_demo/) e não representam operações reais.

A execução padrão utiliza arquivos locais. A integração com Gmail é opcional e exige configuração própria. O Dashboard demonstrativo é `.xlsx` e não inclui macros nem processamento de `.xlsb`.

A aplicação foi preparada para execução local. Publicar o código no GitHub não hospeda o sistema; uma implantação em outro ambiente exige configuração e validação específicas.

## Por onde começar

Para executar o projeto, consulte as instruções abaixo sobre preparação do ambiente, configuração da `.env`, criação do usuário e inicialização do Django. Para conhecer os componentes internos, consulte a [documentação técnica](docs/DOCUMENTACAO_TECNICA.md).

## NAVEGAÇÃO RÁPIDA

- [Documentação técnica do projeto](docs/DOCUMENTACAO_TECNICA.md)

- [Extrair o ZIP e encontrar a pasta correta](#comece-aqui-extração-e-pasta-correta)
- [Ligar, parar e reiniciar o servidor](#servidores-ligar-parar-e-reiniciar)
- [Entender o Waitress](#o-que-é-waitress-e-como-ligá-lo)
- [Encontrar e ler os logs](#logs-onde-encontrar-e-como-ler)
- [Configurar Gmail e acesso ao Dashboard](#51-guia-completo-gmail-senha-de-app-e-acesso-ao-dashboard)
- [Enviar os quatro relatórios fictícios](#etapa-g-enviar-os-quatro-arquivos-fictícios-para-sua-conta)
- [Ligar e desligar o modo teste](#71-ligar-e-desligar-o-modo-teste-em-settingspy)

> **LEITURA DOS TÍTULOS:** no VS Code, abra o README e pressione `Ctrl+Shift+V` para ver a pré-visualização formatada. No GitHub ou GitLab, os títulos e tabelas são renderizados automaticamente.

# CONTEXTO E FINALIDADES DO SISTEMA

Este projeto apresenta um sistema de automação para finalidades específicas de tratamento e consolidação de relatórios. Ele recebe arquivos, prepara e valida dados, atualiza um Dashboard Excel e oferece uma interface web para acompanhar o processo.

O cenário didático utiliza uma empresa fictícia e quatro relatórios genéricos: **vendas, recebimentos, custos e despesas**. Nomes de clientes, fornecedores, equipes, identificadores e valores são inventados exclusivamente para os exemplos. Os arquivos não representam operações de uma entidade real.

## Servidor local e possível uso organizacional

A demonstração roda em um **servidor local**, no computador de quem realiza o teste. O mesmo tipo de solução pode ser adaptado a um servidor organizacional, conforme os requisitos de acesso, infraestrutura e operação de cada projeto. Essa adaptação exige validação própria e não é configurada automaticamente por este pacote.

## Antes de executar: escolher a origem dos relatórios

| Forma de testar | Configuração na `.env` | O que precisa existir antes da execução |
| --- | --- | --- |
| Arquivos locais | `AUTOMACAO_ORIGEM=local` | Ambiente preparado com `preparar_demo.py` e quatro relatórios na pasta `entrada` |
| Recebimento real pelo Gmail pessoal | `AUTOMACAO_ORIGEM=gmail` | Conta e senha de app configuradas e **um e-mail já recebido com os quatro relatórios fictícios anexados** |

**No teste pelo Gmail, ter os arquivos apenas no computador não é suficiente. Você precisa enviá-los por e-mail antes de executar a atualização.** O sistema lê a mensagem recebida; ele não envia os relatórios automaticamente. Se a mensagem estiver ausente ou incompleta, a automação interrompe a execução e não utiliza silenciosamente os arquivos locais como substitutos.

Os quatro relatórios já acompanham este ZIP na pasta **`dados_demo`**. Use os modelos que já acompanham o projeto. Consulte a [Etapa G: localizar e enviar os arquivos fictícios](#etapa-g-enviar-os-quatro-arquivos-fictícios-para-sua-conta) para o passo a passo.

# COMECE AQUI: EXTRAÇÃO E PASTA CORRETA

> **ANTES DE EXECUTAR:** extraia o projeto do ZIP. Trabalhe na pasta que contém `manage.py`, `requirements.txt`, `preparar_demo.py` e `README.md`.

## Extrair o ZIP no Windows

1. Na pasta Downloads, localize `automacao-dashboard-portfolio.zip`.
2. Clique com o botão direito no ZIP e escolha **Extrair Tudo**.
3. Escolha uma pasta de destino e conclua a extração. Se quiser um caminho mais curto, use `C:\Projetos`.
4. Abra a pasta extraída. Se existir outra pasta chamada `automacao-dashboard-portfolio` dentro dela, abra também essa pasta interna.
5. Pare de entrar em subpastas quando encontrar `manage.py`, `requirements.txt`, `preparar_demo.py` e `README.md` juntos. Essa é a **raiz do projeto**.
6. No VS Code, use **Arquivo → Abrir Pasta** e selecione essa raiz. Depois, abra **Terminal → Novo Terminal**.

Não execute os scripts pela visualização do arquivo ZIP no Explorador. Também não escolha a pasta externa que apenas contém outra pasta com o projeto. A estrutura externa pode variar conforme a forma de extração; a presença de `manage.py` é o critério para identificar a pasta correta.

Você pode mover a pasta interna completa para um lugar como `C:\Projetos\automacao-dashboard-portfolio`. Não mova apenas os arquivos Python e não apague pastas para tentar “corrigir” o caminho. Se já criou a `.venv` antes de mover o projeto, recrie o ambiente virtual no destino em vez de reutilizar uma `.venv` movida.

## Conferir onde o terminal está aberto

No PowerShell:

```powershell
Get-Location
Get-ChildItem -Name
```

O primeiro comando mostra a pasta atual; o segundo lista os arquivos dela. Se a lista mostrar somente outra pasta com o mesmo nome do projeto, entre nela:

```powershell
cd .\automacao-dashboard-portfolio
```

Se o projeto estiver em outro local, informe seu caminho. O exemplo abaixo deve ser adaptado à sua pasta real:

```powershell
cd "C:\Projetos\automacao-dashboard-portfolio"
```

> **DIRETÓRIO CORRETO:** execute os comandos deste README na raiz do projeto. Os exemplos `python manage.py ...` e `python preparar_demo.py` dependem de você estar nessa pasta. O prefixo `PS C:\...>` exibido pelo terminal não faz parte do comando e não deve ser copiado.

# SERVIDORES: LIGAR, PARAR E REINICIAR

## O que significa ligar o servidor

O servidor mantém a aplicação disponível para o navegador. O terminal fica ocupado enquanto ele funciona. Abrir uma aba no navegador não inicia o servidor; fechar a aba também não o desliga.

Existem duas opções neste guia. Escolha **uma por vez** para acompanhar a demonstração sem confundir processos e portas:

| Opção | Para que serve aqui | Endereço usado neste guia | Como parar |
| --- | --- | --- | --- |
| `runserver` | Desenvolvimento e testes iniciais com o servidor do Django | `http://127.0.0.1:8000/` | `Ctrl+C` no terminal que o iniciou |
| Waitress | Servir a mesma aplicação Django por WSGI, em uma execução local semelhante à organização de um servidor de aplicação | `http://127.0.0.1:8080/` | `Ctrl+C` no terminal que o iniciou |

Para começar os estudos, use `runserver`. Ele é destinado a desenvolvimento, não a produção. O Waitress é um servidor WSGI: recebe requisições HTTP do navegador e as entrega à aplicação Python/Django. Ele não substitui o Django, não é banco de dados e não faz a automação das planilhas por conta própria. [Documentação de implantação do Django](https://docs.djangoproject.com/en/5.2/howto/deployment/) e [documentação do Waitress](https://docs.pylonsproject.org/projects/waitress/en/latest/).

Usar Waitress no próprio computador continua sendo uma demonstração local. A implantação em outro ambiente depende de preparação específica; executar esse comando não publica o projeto na internet nem comprova sua implantação na organização.

## Antes de ligar qualquer opção

- Confirme a raiz do projeto e a instalação de `requirements.txt` na `.venv`.
- Prepare os arquivos fictícios e a `.env` conforme as etapas do guia.
- Crie o banco e o usuário do Dashboard antes de tentar entrar.
- Confira se não existe outro servidor desta demonstração aberto em outro terminal.
- Se estiver trocando de servidor, aguarde qualquer atualização terminar e desligue o servidor anterior primeiro.

> **SERVIDOR DESLIGADO DURANTE ALTERAÇÕES:** antes de editar `.env` ou `settings.py`, instalar dependências, executar migrações ou restaurar os arquivos fictícios, aguarde a automação concluir e pare o servidor com `Ctrl+C`. Depois da alteração, inicie novamente a opção escolhida.

## Ligar com runserver

Na raiz do projeto, usando o Python da `.venv` sem precisar ativá-la:

> **ANTES DE LIGAR:** execute na raiz do projeto e confirme que não há outro servidor desta demonstração aberto. Para parar, aguarde a atualização terminar e pressione `Ctrl+C` no mesmo terminal.

```powershell
.\.venv\Scripts\python.exe manage.py runserver 127.0.0.1:8000
```

Aguarde a mensagem de inicialização e abra <http://127.0.0.1:8000/>. Mantenha esse terminal aberto. Se houver erro de inicialização, o servidor não ficou disponível só porque você executou o comando; leia a mensagem apresentada.

## O que é Waitress e como ligá-lo

Waitress é uma biblioteca/servidor para aplicações Python que usam a interface WSGI. Neste projeto, o objeto servido é `projeto_web.wsgi:application`. Ele já está listado no `requirements.txt`.

A demonstração usa `127.0.0.1`, para acesso no próprio computador, e a porta `8080`. O exemplo usa uma thread de atendimento para simplificar a execução local. O processamento assíncrono das planilhas continua sendo realizado pelo worker separado que já existe na aplicação. [Opções oficiais do Waitress](https://docs.pylonsproject.org/projects/waitress/en/latest/runner.html).

1. Pare o `runserver`, se estiver aberto, e aguarde o terminal voltar ao prompt.
2. Na raiz do projeto, prepare os arquivos estáticos. Isso reúne CSS e outros recursos em `staticfiles`, para que o middleware WhiteNoise já configurado possa servi-los:

```powershell
.\.venv\Scripts\python.exe manage.py collectstatic --noinput
```

3. Inicie o Waitress:

> **ANTES DE LIGAR:** execute na raiz do projeto e confirme que não há outro servidor desta demonstração aberto. Para parar, aguarde a atualização terminar e pressione `Ctrl+C` no mesmo terminal.

```powershell
.\.venv\Scripts\python.exe -m waitress --listen=127.0.0.1:8080 --threads=1 projeto_web.wsgi:application
```

4. Aguarde a inicialização e abra <http://127.0.0.1:8080/>.
5. Mantenha esse terminal aberto enquanto usa a aplicação.

O `collectstatic` não liga o servidor e normalmente termina devolvendo o prompt. Já o comando do Waitress mantém o processo em execução. Diferentemente do fluxo de desenvolvimento com recarga automática, não conte com o Waitress para aplicar alterações nos arquivos sem reiniciá-lo.

## Parar o servidor corretamente

1. Aguarde a atualização das planilhas terminar. Fechar o servidor web não é uma forma segura de cancelar uma atualização: o worker da automação pode continuar em outro processo.
2. Clique no terminal que está executando o `runserver` ou o Waitress.
3. Pressione `Ctrl+C`.
4. Aguarde o retorno do prompt, por exemplo `PS C:\Projetos\automacao-dashboard-portfolio>`.

O retorno do prompt após o encerramento indica que aquele comando terminou. O comando `cls` apenas limpa a tela; ele não desliga o servidor. O botão **Sair** da página encerra a sessão do usuário, mas também não desliga o servidor.

Se iniciou duas instâncias por engano, cada terminal corresponde a um processo. Pare somente os processos deste projeto que você iniciou e identificou. Não encerre todos os processos Python do computador indiscriminadamente.

## Reiniciar depois de alterar a configuração

1. Aguarde a automação terminar e pare o servidor.
2. Faça a alteração e salve o arquivo.
3. Execute novamente o comando do servidor escolhido.
4. Abra ou atualize a página no endereço correspondente à porta usada.

> **REINÍCIO NECESSÁRIO:** alterações em Gmail/senha de app, origem dos relatórios, motor, modo teste e demais valores da `.env` exigem reiniciar o servidor para que seus novos processos recebam a configuração atualizada.

Se o terminal disser que a porta já está em uso, verifique os terminais abertos. Não tente resolver iniciando mais cópias do mesmo servidor. A porta `8000` pertence ao exemplo de `runserver`; a `8080`, ao exemplo de Waitress.

## Preciso de servidor para executar o script diretamente?

Não. O comando `automacao/executar_automacao.py` pode funcionar pelo terminal sem `runserver` e sem Waitress. Para um primeiro teste direto, mantenha o servidor parado, execute o script e aguarde a conclusão. Para usar a página e o histórico web, ligue uma das opções de servidor.

# LOGS: ONDE ENCONTRAR E COMO LER

## O que é um log

Um log é um registro de acontecimentos do programa: quando começou uma etapa, quantos registros foram encontrados, se houve erro e se a atualização terminou. Ele ajuda a identificar em que ponto o processo parou. O log não é a planilha final nem o histórico de usuários do banco.

## Quais registros este projeto gera

Os caminhos abaixo são relativos à **raiz do projeto**, a pasta de `manage.py`:

| Registro | Onde encontrar | Quando aparece | Para que olhar |
| --- | --- | --- | --- |
| Log principal da automação | `logs/automacao.log` | Após o executor configurar o log, seja pelo terminal ou pela interface | Etapas 01 a 17, contagens, auditoria e resultado final |
| Diagnóstico do worker da interface | `diagnostico_atualizacao.log` | Quando o botão da página inicia a atualização assíncrona | Saída e erros do processo separado, inclusive falhas antes de começar o motor |
| Saída do servidor web | Terminal que executa `runserver` ou Waitress | Enquanto o servidor estiver ligado | Inicialização, erros do servidor e mensagens exibidas durante a execução |
| Histórico da aplicação | Tabela de histórico na página, gravada no SQLite local | Execuções iniciadas pelo fluxo web | Usuário, início, fim, estado e progresso |

A saída do servidor web não é automaticamente gravada em um arquivo de log por esses comandos. O histórico da página também não substitui os arquivos de diagnóstico. Na execução direta do script, o log principal é gerado, mas não é criado um registro no histórico web.

Se o projeto estiver em `C:\Projetos\automacao-dashboard-portfolio`, o log principal ficará em:

```text
C:\Projetos\automacao-dashboard-portfolio\logs\automacao.log
```

Os logs podem não existir antes da primeira execução. Isso é normal. Uma falha antes de carregar o motor pode aparecer somente no terminal ou no diagnóstico do worker.

## Abrir no VS Code ou no Explorador

No painel **Explorador** do VS Code, abra a pasta `logs` e selecione `automacao.log`. O arquivo `diagnostico_atualizacao.log` fica diretamente na raiz do projeto. Você também pode abrir os mesmos arquivos pelo Explorador do Windows, usando um editor de texto.

Os registros são acrescentados ao final dos arquivos. Se o editor não atualizar a tela imediatamente, reabra o arquivo ou use os comandos de leitura abaixo.

## Ler as últimas linhas pelo PowerShell

Abra **um segundo terminal**, também na raiz do projeto. Não tente digitar esses comandos no terminal ocupado pelo servidor:

```powershell
Get-Content .\logs\automacao.log -Tail 50
```

Para o diagnóstico do worker:

```powershell
Get-Content .\diagnostico_atualizacao.log -Tail 50
```

Para acompanhar novas linhas do log principal enquanto a atualização ocorre:

```powershell
Get-Content .\logs\automacao.log -Tail 30 -Wait
```

> **PARAR A LEITURA NÃO É PARAR O SERVIDOR:** no segundo terminal, `Ctrl+C` encerra apenas o acompanhamento de `Get-Content -Wait`. Para desligar o servidor, use `Ctrl+C` no terminal em que ele foi iniciado, depois de aguardar a automação terminar.

## Entender as mensagens

| Texto no registro | Significado |
| --- | --- |
| `INFO` | Informação de andamento ou resultado |
| `[ETAPA 03]` | O programa chegou à leitura dos relatórios |
| `[AUDITORIA]` | Conferência dos dados antes de salvar |
| `[FINAL] Sucesso` ou `[FINAL win32com] Sucesso` | O motor informou conclusão com sucesso |
| `WARNING` | Aviso que precisa ser lido; não significa necessariamente que toda a atualização falhou |
| `ERROR`, `Traceback`, `[FINAL] Falha` | Houve falha; leia o contexto para localizar a causa |

Procure o horário da execução que está investigando. Uma falha antiga não significa que a execução mais recente falhou. Como os testes automatizados exercitam erros de propósito, eles também podem deixar mensagens de falha nesses arquivos; confira o contexto e o resultado final dos testes.

Ao pedir ajuda, envie somente o trecho da execução com o erro e as linhas anteriores que ajudam a entendê-lo. Confira e oculte dados pessoais e caminhos privados antes de compartilhar. Nunca envie sua `.env` ou senhas junto com o log.

## Quando limpar os logs

Para o estudo, normalmente basta ler as últimas linhas. Se quiser arquivar ou limpar os registros, aguarde a automação terminar, pare o servidor e feche o acompanhamento de logs. Guarde uma cópia se ainda precisa investigar um erro. Limpar um log não corrige a causa da falha.


# 1. O QUE A DEMONSTRAÇÃO FAZ

- Recebe quatro planilhas fictícias em uma pasta local.
- Copia os relatórios para a pasta de processamento.
- Lê e prepara os dados, removendo as linhas TOTAL dos relatórios.
- Valida colunas, identificadores e valores monetários.
- Ajusta a quantidade de linhas do Dashboard.
- Atualiza registros, totais, fórmulas e um resumo por relatório.
- Preserva as colunas P:U de RECEBIMENTOS.
- Audita os valores escritos antes de salvar.
- Exibe o progresso e registra o histórico quando executada pela interface.

O mesmo executor pode ser iniciado diretamente pelo terminal. Nesse caso, o log é gravado, mas o histórico do Django não é atualizado.

# 2. TECNOLOGIAS E DIFERENÇAS DA APLICAÇÃO ORIGINAL

| Parte | Versão de portfólio |
| --- | --- |
| Interface | Django 5.2, HTML e CSS |
| Preparação dos dados | pandas |
| Leitura e escrita do Dashboard | Principal: Excel COM (`win32com`). Alternativa: `openpyxl`. Arquivo fictício `.xlsx` |
| Banco local | SQLite, criado pelas migrações |
| Configuração | python-dotenv e `.env` |
| E-mail | Padrão local; Gmail pessoal opcional por IMAP/TLS e senha de app |
| Excel instalado | Necessário para `win32com`; dispensado na alternativa `openpyxl` |
| Execução | Terminal ou interface local autenticada |

O projeto fornece um Dashboard de exemplo `.xlsx`, compatível com as duas opções de motor. Não inclui macros nem processamento de `.xlsb`. A escolha de controlar o Excel com `win32com` não exige trocar o formato dos arquivos. Não altere apenas a extensão de uma planilha para tentar convertê-la.

A alternativa `openpyxl` facilita a execução sem Excel. Ela precisa ser escolhida explicitamente: o programa não troca de motor automaticamente quando o Excel está ausente.

A parte portável e os testes automatizados foram validados em Python 3.12.14, Django 5.2.19, pandas 2.2.3 e openpyxl 3.1.5. A integração real com Excel ainda precisa ser validada no Windows. O `requirements.txt` permite atualizações dentro das faixas indicadas. Outros ambientes, incluindo Python 3.14, precisam executar os testes após instalar suas dependências.

# 3. ORGANIZAÇÃO DAS PASTAS

| Caminho | Responsabilidade |
| --- | --- |
| `automacao/executar_automacao.py` | Entrada única: escolhe o motor e preserva o contrato das views |
| `automacao/motor_win32com.py` | Motor principal: abre, atualiza, recalcula e fecha o Excel |
| `automacao/motor_openpyxl.py` | Motor alternativo: modifica o arquivo sem abrir o Excel |
| `automacao/fonte_relatorios.py` | Origem local ou Gmail; download dos quatro anexos conhecidos |
| `automacao/verificar_gmail.py` | Testa o acesso de leitura sem atualizar o Dashboard |
| `automacao/config.py` | Define os nomes, colunas e posições dos relatórios |
| `dashboard/` | Views, modelos, permissões e interface Django |
| `projeto_web/` | Configuração e rotas da aplicação |
| `dados_demo/` | Quatro relatórios fictícios e o modelo inicial do Dashboard |
| `preparar_demo.py` | Copia os modelos para o ambiente de execução |
| `tests/test_demo.py` | Testa o fluxo real de planilhas e casos de falha |
| `.env.example` | Exemplo público com configurações fictícias |
| `.gitignore` | Evita versionar configurações locais e resultados |
| `TESTE_AUTOMACAO/` | Criada localmente; contém arquivos de trabalho |
| `logs/automacao.log` | Log gerado pelo executor |

# 4. PREPARAR O AMBIENTE NO WINDOWS

> **ANTES DE ALTERAR:** aguarde a automação terminar e pare `runserver` ou Waitress com `Ctrl+C`. Depois das alterações, reinicie o servidor escolhido.

Extraia esta versão em uma pasta nova, por exemplo:

`C:\Projetos\automacao-dashboard-portfolio`

Extraia em uma pasta nova. Evite misturar arquivos de instalações diferentes.

Abra a pasta pelo VS Code: **Arquivo → Abrir Pasta**. O terminal integrado pode usar PowerShell ou CMD; os comandos dependem do shell selecionado.

## PowerShell

Execute um comando por vez na raiz, onde está `manage.py`:

```powershell
cd C:\Projetos\automacao-dashboard-portfolio
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
Copy-Item .env.example .env
.\.venv\Scripts\python.exe preparar_demo.py
```

**Execute uma linha por vez, pressionando Enter ao final de cada comando.** O nome do script é exatamente `preparar_demo.py`.

Para preparar os arquivos no PowerShell sem ativar o ambiente, copie apenas esta linha:

```powershell
.\.venv\Scripts\python.exe .\preparar_demo.py
```

Se desejar limpar a tela, execute `cls` em uma linha separada. Não acrescente `cls` ao nome do script: `preparar_demo.pycls` não existe.

Se aparecer `[Errno 2] No such file or directory`, confira se você abriu a pasta que contém o arquivo:

```powershell
Get-ChildItem -Name
```

A lista deve conter `preparar_demo.py`, `manage.py` e `requirements.txt`. Se houver apenas outra pasta chamada `automacao-dashboard-portfolio`, entre nessa subpasta antes de continuar.

Usar o caminho do Python da `.venv` funciona sem alterar a política de execução do PowerShell. Se preferir ativar o ambiente e sua política permitir:

```powershell
.\.venv\Scripts\Activate.ps1
```

Com o ambiente ativado, os exemplos abaixo podem ser executados com `python`.

## CMD

```bat
cd /d C:\Projetos\automacao-dashboard-portfolio
py -m venv .venv
.venv\Scripts\activate.bat
python -m pip install -r requirements.txt
copy .env.example .env
python preparar_demo.py
```

## Linux ou macOS

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
cp .env.example .env
python preparar_demo.py
```

Se `py` não existir no Windows e `python --version` funcionar, substitua `py` por `python` no comando de criação do ambiente.

## O que é a `.venv`?

É uma pasta com um ambiente Python separado para o projeto. As bibliotecas instaladas ali ficam associadas à aplicação. A pasta não deve ser enviada para o Git.

# 5. ENTENDER O `.env.example`

> **ANTES DE ALTERAR:** aguarde a automação terminar e pare `runserver` ou Waitress com `Ctrl+C`. Depois das alterações, reinicie o servidor escolhido.

`.env.example` documenta as variáveis necessárias. `.env` é sua cópia local, carregada pelo Django.

| Variável | Uso |
| --- | --- |
| `DJANGO_SECRET_KEY` | Chave local do Django; o valor fornecido é um marcador de desenvolvimento |
| `DJANGO_DEBUG` | `1` mostra detalhes de erros durante desenvolvimento |
| `DJANGO_ALLOWED_HOSTS` | Hosts aceitos no acesso local |
| `DJANGO_TIME_ZONE` | Fuso usado pela aplicação |
| `AUTOMACAO_MODO_TESTE` | `1`: ambiente de teste. `0`: ambiente configurado com dados fictícios |
| `AUTOMACAO_MOTOR` | `win32com` por padrão; `openpyxl` para executar sem Excel |
| `DJANGO_DEPLOYMENT_MODE` | Deve permanecer `0` na demonstração local |
| `DJANGO_EMAIL_BACKEND` | Exibe mensagens de recuperação de senha no terminal |
| `AUTOMACAO_ORIGEM` | `local` dispensa e-mail; `gmail` habilita a leitura da conta pessoal configurada |
| `EMAIL_IMAP_HOST`, `EMAIL_IMAP_PORT` | Servidor `imap.gmail.com`, porta `993` com TLS |
| `EMAIL_USUARIO`, `EMAIL_SENHA_APP` | Gmail completo e senha de app; usados somente na origem `gmail` |
| `EMAIL_REMETENTE`, `EMAIL_ASSUNTO` | Filtros para localizar a mensagem com os quatro anexos |
| `DASHBOARD_ADMIN_USUARIO`, `DASHBOARD_ADMIN_EMAIL`, `DASHBOARD_ADMIN_SENHA` | Criação explícita do administrador local com `criar_admin_env` |

Na origem `local`, não é necessário cadastrar e-mail. Para a origem `gmail`, siga o guia detalhado abaixo antes de executar. Os caminhos da demonstração são definidos relativamente à raiz do projeto, sem caminhos pessoais no `.env`.

A `.env.example` pode ser publicada. A `.env` é ignorada pelo Git. O marcador de chave e o modo DEBUG destinam-se ao desenvolvimento local; esta entrega não configura hospedagem pública da aplicação.

# 5.1. GUIA COMPLETO: GMAIL, SENHA DE APP E ACESSO AO DASHBOARD

Este guia começa na preparação da conta e termina com o Dashboard atualizado. Você só precisa de Gmail se quiser demonstrar a leitura dos anexos pelo e-mail. Para executar somente com os arquivos locais, mantenha `AUTOMACAO_ORIGEM=local` e pule a preparação do Gmail.

## Antes de começar: são duas contas diferentes

| Acesso | Para que serve | Usuário | Senha |
| --- | --- | --- | --- |
| Gmail no navegador | Abrir sua caixa de e-mail e gerenciar a conta Google | Seu endereço Gmail | Sua senha normal do Google, usada somente no site Google |
| Programa lendo Gmail | Baixar os quatro relatórios enviados por você | Mesmo endereço Gmail | Senha de app gerada pelo Google |
| Dashboard local | Entrar na página Django e clicar em Atualizar | Nome escolhido para o Dashboard, por exemplo `admin_demo` | Outra senha, criada por você exclusivamente para o Dashboard |

A senha de app não é uma chave de API nem o código temporário recebido no celular. O programa não pede sua senha normal do Google. A senha do Dashboard também não deve ser a senha do Gmail.

## Etapa A: conferir os programas e a pasta do projeto

1. No Windows, tenha o Python instalado e disponível no terminal. Verifique:

```powershell
py --version
```

2. Se o Python não estiver instalado, use o [site oficial do Python](https://www.python.org/downloads/windows/). Na instalação tradicional, marque a opção para adicionar o Python ao PATH quando ela estiver disponível. Feche e abra o terminal depois da instalação.
3. Abra a pasta extraída no VS Code ou abra o PowerShell nessa pasta. O comando abaixo deve listar `manage.py`, `requirements.txt` e `preparar_demo.py`:

```powershell
Get-ChildItem -Name
```

4. Execute cada linha abaixo separadamente. Os exemplos usam o Python da `.venv`, por isso não exigem ativar scripts no PowerShell:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Se a `.venv` já existir, use-a e execute apenas a instalação das dependências. Para este projeto, `imaplib`, `email`, `ssl` e `tempfile` já vêm com o Python; não é necessário instalá-los com pip.

5. Se você ainda não tem uma `.env`, copie o modelo:

```powershell
Copy-Item .env.example .env
```

Se a `.env` já existir, abra e edite a que existe. Não a substitua, pois ela pode conter sua configuração. Use o VS Code para abrir o arquivo chamado exatamente `.env`, na mesma pasta de `manage.py`. Ele não deve se chamar `.env.txt`.

## Etapa B: ter uma conta Gmail pessoal

1. Abra [Gmail](https://mail.google.com/) no navegador.
2. Se já possui Gmail, entre na sua conta pessoal. Se não possui, escolha **Criar conta**, selecione a opção de uso pessoal e conclua o cadastro apresentado pelo Google.
3. Anote seu endereço completo, por exemplo `seunome@gmail.com`.
4. Confirme que consegue abrir a Caixa de entrada normalmente.

O adaptador fornecido foi feito para Gmail pessoal (`@gmail.com`). Contas de trabalho, Outlook, Yahoo e outros provedores não têm necessariamente o mesmo método de autenticação e não estão implementados neste guia. Se a sua conta não permite senha de app, use o modo local; não desative proteções da conta para contornar a restrição.

## Etapa C: primeiro passo no Google — ativar a verificação em duas etapas

1. Abra [sua Conta do Google](https://myaccount.google.com/).
2. Confira a foto/inicial e o endereço da conta para não alterar outra conta aberta no navegador.
3. Entre em **Segurança e login** ou **Segurança**.
4. Em **Como você faz login no Google**, abra **Verificação em duas etapas**.
5. Siga o procedimento do Google, confirme sua identidade e configure a segunda etapa oferecida.
6. Ao terminar, confira se ela aparece como ativada.

Essa segunda etapa protege o acesso à conta. O Google pode solicitar uma confirmação no celular ou outro método disponível. Os nomes dos menus podem variar. [Instruções oficiais da verificação em duas etapas](https://support.google.com/accounts/answer/185839?hl=pt-BR).

## Etapa D: segundo passo no Google — gerar a senha de app

1. Com a mesma conta aberta e a verificação em duas etapas ativada, acesse [Senhas de app](https://myaccount.google.com/apppasswords).
2. Se solicitado, confirme sua identidade no Google.
3. Dê ao aplicativo um nome que você reconheça, como `Dashboard Portfolio`.
4. Escolha a opção de criar/gerar a senha.
5. Copie a senha de app exibida, de 16 caracteres. Ela pode aparecer separada em grupos por espaços.
6. Guarde-a somente na sua `.env`, no campo `EMAIL_SENHA_APP`. Não a envie em mensagens, prints, GitHub ou GitLab.

A senha normalmente é exibida uma única vez. Caso perca o valor, gere outra. Se a opção não aparecer, pode haver restrição da conta, Proteção Avançada ou configuração apenas com chaves de segurança. O Google recomenda “Fazer login com o Google” quando disponível; esta implementação educativa usa senha de app e não implementa OAuth. [Orientação oficial sobre senhas de app](https://support.google.com/accounts/answer/185833?hl=pt-BR).

Não procure por “permitir aplicativos menos seguros”. Para Gmail pessoal, o acesso IMAP já fica habilitado; não é necessário encontrar um botão “Ativar IMAP”. [Explicação oficial do Gmail](https://support.google.com/mail/answer/7126229?hl=pt-BR).

## Etapa E: preencher o acesso do programa ao Gmail na .env

Abra sua `.env` e localize estas variáveis. Os textos em maiúsculas abaixo são instruções para substituir; não são credenciais utilizáveis:

```dotenv
AUTOMACAO_MODO_TESTE=1
AUTOMACAO_MOTOR=win32com
AUTOMACAO_ORIGEM=gmail
EMAIL_IMAP_HOST=imap.gmail.com
EMAIL_IMAP_PORT=993
EMAIL_USUARIO=SEU_ENDERECO_COMPLETO@gmail.com
EMAIL_SENHA_APP=COLE_AQUI_A_SENHA_DE_APP
EMAIL_REMETENTE=SEU_ENDERECO_COMPLETO@gmail.com
EMAIL_ASSUNTO=PORTFOLIO_DASHBOARD_DEMO
```

Preencha assim:

| Campo | O que colocar |
| --- | --- |
| `AUTOMACAO_MODO_TESTE` | Comece com `1`, para usar `TESTE_AUTOMACAO` |
| `AUTOMACAO_MOTOR` | `win32com` se tem Windows e Excel desktop; `openpyxl` se vai executar sem Excel |
| `AUTOMACAO_ORIGEM` | `gmail` para baixar os anexos; `local` para usar apenas arquivos da pasta |
| `EMAIL_IMAP_HOST` | Mantenha `imap.gmail.com` |
| `EMAIL_IMAP_PORT` | Mantenha `993`; a conexão usa TLS |
| `EMAIL_USUARIO` | Endereço completo da conta Gmail que vai receber os relatórios |
| `EMAIL_SENHA_APP` | Senha de app de 16 caracteres; pode colar sem os espaços de apresentação |
| `EMAIL_REMETENTE` | Conta que enviará os relatórios. Para enviar a si mesmo, repita seu Gmail; em branco, o programa usa `EMAIL_USUARIO` |
| `EMAIL_ASSUNTO` | Mantenha `PORTFOLIO_DASHBOARD_DEMO` para acompanhar este guia |

Não deixe espaços dentro dos nomes das variáveis. Mantenha uma variável por linha e salve o arquivo com `Ctrl+S`. O programa remove os espaços de apresentação da senha de app; não remove caracteres da senha do Dashboard.

Se você usa `openpyxl`, altere apenas a linha do motor para:

```dotenv
AUTOMACAO_MOTOR=openpyxl
```

O servidor de recebimento Gmail é `imap.gmail.com:993`. SMTP é outro protocolo, usado para enviar e-mail; ele não é necessário para baixar os relatórios neste projeto. [Servidores oficiais do Gmail](https://developers.google.com/workspace/gmail/imap/imap-smtp?hl=en).

## Etapa F: testar a conexão antes de mexer no Dashboard

Pare o servidor Django, se estiver aberto, com `Ctrl+C`. No terminal, na pasta do projeto:

```powershell
.\.venv\Scripts\python.exe -m automacao.verificar_gmail
```

Resultado esperado:

```text
Conexão com Gmail e acesso somente de leitura à Caixa de entrada: OK.
```

Esse comando autentica, abre a Caixa de entrada em modo de leitura e encerra a sessão. Ele não envia mensagens, não baixa anexos, não altera o Dashboard e não exibe a senha.

Se falhar, confira se a origem está como `gmail`, se usou a conta correta e se colou a senha de app, em vez da senha normal. Não publique o conteúdo da `.env` para pedir ajuda; informe somente a mensagem de erro, ocultando seus dados pessoais.

## Etapa G: enviar os quatro arquivos fictícios para sua conta

**Esta etapa é obrigatória se `AUTOMACAO_ORIGEM=gmail`.** Ela prepara uma leitura real de e-mail usando somente anexos fictícios. Com `AUTOMACAO_ORIGEM=local`, não é necessário enviar e-mail.

### Onde encontrar os arquivos de teste

1. Localize a pasta onde você extraiu `automacao-dashboard-portfolio.zip`.
2. Abra a pasta que contém `manage.py`, `README.md` e `preparar_demo.py`.
3. Dentro dela, abra **`dados_demo`**. Essa pasta já é fornecida no pacote; os relatórios não dependem de download em outro site.
4. Você verá cinco planilhas: quatro relatórios de origem e um Dashboard de destino. Envie apenas os quatro relatórios listados na tabela a seguir.

Por exemplo, se extraiu o projeto em `C:\Projetos\automacao-dashboard-portfolio`, os arquivos ficam em:

```text
C:\Projetos\automacao-dashboard-portfolio\dados_demo
```

No seu computador, o início do caminho pode ser diferente, como a pasta Downloads. No VS Code, a mesma pasta aparece no painel **Explorador**, à esquerda. Para conferir pelo PowerShell, na raiz do projeto:

```powershell
Get-ChildItem .\dados_demo\*.xlsx
```

Esse comando lista os cinco arquivos; ele não envia e-mails. O `preparar_demo.py` copia os modelos para as pastas de trabalho, mas também não realiza o envio. O envio é feito por você no Gmail, conforme os passos abaixo.

### Como encaminhar os arquivos para o teste de recebimento

1. Execute o preparo local:

```powershell
.\.venv\Scripts\python.exe preparar_demo.py
```

2. No navegador, abra o Gmail da conta que você definiu como remetente.
3. Clique em **Escrever**.
4. Em **Para**, informe o Gmail definido em `EMAIL_USUARIO`. Você pode enviar para a própria conta.
5. Em **Assunto**, digite exatamente:

```text
PORTFOLIO_DASHBOARD_DEMO
```

6. Clique no ícone de anexar arquivo e selecione **os quatro relatórios da pasta `dados_demo`**:

| Arquivo a anexar | Conteúdo |
| --- | --- |
| [vendas_exemplo.xlsx](dados_demo/vendas_exemplo.xlsx) | Vendas de exemplo |
| [recebimentos_exemplo.xlsx](dados_demo/recebimentos_exemplo.xlsx) | Recebimentos de exemplo |
| [custos_exemplo.xlsx](dados_demo/custos_exemplo.xlsx) | Custos de exemplo |
| [despesas_exemplo.xlsx](dados_demo/despesas_exemplo.xlsx) | Despesas de exemplo |

7. Não anexe o arquivo `dashboard_exemplo.xlsx`. Ele é o destino local da atualização, não um relatório de origem.
8. Envie os quatro arquivos como anexos reais, sem ZIP e sem substituir por links do Google Drive.
9. Aguarde o envio terminar e confirme que a mensagem chegou à **Caixa de entrada** da conta destinatária. Se caiu em Spam, mova-a para a Caixa de entrada após conferir que é sua mensagem.

O leitor procura o assunto exato e o remetente configurado, nos 50 resultados mais recentes encontrados na Caixa de entrada. Usa a mensagem correspondente recebida mais recentemente, pela ordem dos UIDs IMAP. Se essa mensagem estiver incompleta, a atualização é interrompida; o programa não reutiliza silenciosamente um e-mail antigo.

Envie os quatro relatórios juntos, preservando os nomes. Há limite de 5 MiB por anexo e de 15 MiB por mensagem recebida. Os modelos fornecidos são pequenos. Somente os quatro nomes esperados são gravados. Use exclusivamente os dados fictícios desta distribuição.

## Etapa H: executar a atualização pelo terminal

Antes de iniciar com `AUTOMACAO_ORIGEM=gmail`, confira:

- A mensagem já chegou à Caixa de entrada da conta indicada em `EMAIL_USUARIO`.
- O remetente corresponde a `EMAIL_REMETENTE` (ou ao próprio usuário, se o campo estiver vazio).
- O assunto é exatamente `PORTFOLIO_DASHBOARD_DEMO`, sem adicionar `RE:` ou `ENC:`.
- Os quatro `.xlsx` estão juntos na mensagem, com os nomes originais, como anexos e não como links.
- A `.env` está salva e a verificação de conexão da Etapa F terminou com sucesso.

**Se você ainda não enviou os arquivos, volte à Etapa G antes de executar.** O modo Gmail autentica e busca os anexos de verdade na sua conta pessoal; os dados continuam fictícios. Essa execução usa sua conta pessoal e os relatórios da empresa fictícia; não equivale à validação de um ambiente organizacional.

O valor de `AUTOMACAO_MODO_TESTE` escolhe as pastas; ele não substitui o envio do e-mail. É possível fazer o recebimento real pelo Gmail mantendo `AUTOMACAO_MODO_TESTE=1`, desde que `AUTOMACAO_ORIGEM=gmail`.

Com Windows e Excel instalado:

```powershell
.\.venv\Scripts\python.exe automacao\executar_automacao.py --motor win32com
```

Ou, sem Excel:

```powershell
.\.venv\Scripts\python.exe automacao\executar_automacao.py --motor openpyxl
```

O programa conecta ao Gmail, lê os anexos, coloca os arquivos em `TESTE_AUTOMACAO/entrada`, prepara as cópias de processamento em `saida`, atualiza o Dashboard e encerra os recursos. A leitura não marca as mensagens como lidas nem apaga e-mails. O programa não envia nenhum e-mail.

Confira `logs/automacao.log` e `TESTE_AUTOMACAO/dashboard/dashboard_exemplo.xlsx`. No `openpyxl`, abra a planilha em um aplicativo de planilhas para recalcular as fórmulas. No `win32com`, o recálculo é solicitado ao Excel.

## Etapa I: criar o login local do Dashboard pela .env

Além das configurações do Gmail, preencha estas três variáveis na sua `.env`:

```dotenv
DASHBOARD_ADMIN_USUARIO=admin_demo
DASHBOARD_ADMIN_EMAIL=SEU_EMAIL_PARA_CADASTRO
DASHBOARD_ADMIN_SENHA=CRIE_UMA_SENHA_EXCLUSIVA_PARA_O_DASHBOARD
```

Escolha uma senha exclusiva com no mínimo 12 caracteres. Combine letras, números e símbolos e evite informações óbvias do usuário. Se a senha contiver `#` ou espaços, coloque o valor inteiro entre aspas duplas na `.env`. Não use a senha normal do Google nem a senha de app nesse campo.

O e-mail de cadastro do Dashboard pode ser o mesmo endereço Gmail, mas a senha deve ser diferente. O login na tela será `DASHBOARD_ADMIN_USUARIO`, por exemplo `admin_demo`, e não automaticamente o e-mail.

Salve a `.env` e execute:

```powershell
.\.venv\Scripts\python.exe manage.py migrate
.\.venv\Scripts\python.exe manage.py criar_admin_env
```

O primeiro comando cria/atualiza as tabelas do banco local. O segundo valida os dados e cria um administrador com as credenciais informadas. Ele não exibe a senha e não altera usuários já existentes. A senha é armazenada pelo Django como hash no banco; a cópia que você escreveu na `.env` continua sendo texto e deve ser protegida.

Se preferir não colocar a senha do Dashboard na `.env`, deixe os três campos de administrador sem uso e crie a conta interativamente:

```powershell
.\.venv\Scripts\python.exe manage.py createsuperuser
```

O terminal pedirá usuário, e-mail e senha. Ao digitar a senha, pode não aparecer nenhum caractere; isso é normal. Escolha apenas uma das formas de criar o mesmo usuário.

## Etapa J: entrar e atualizar pelo navegador

1. Confira se `AUTOMACAO_ORIGEM=gmail` e se `AUTOMACAO_MOTOR` tem a opção desejada.
2. Inicie o servidor:

> **ANTES DE LIGAR:** execute na raiz do projeto e confirme que não há outro servidor desta demonstração aberto. Para parar, aguarde a atualização terminar e pressione `Ctrl+C` no mesmo terminal.

```powershell
.\.venv\Scripts\python.exe manage.py runserver
```

3. Abra <http://127.0.0.1:8000/> no navegador.
4. Entre com o nome de usuário e a senha **do Dashboard**.
5. Confira na página o motor e a origem selecionados.
6. Clique em **Atualizar Dashboard** e acompanhe o progresso.
7. Aguarde a conclusão e clique em **Baixar Dashboard atualizado**.
8. Confira a planilha baixada. A execução pela página também aparece no histórico.
9. Para sair da conta, clique em **Sair**. Para desligar o servidor, aguarde a atualização terminar e pressione `Ctrl+C` no terminal.

Depois de mudar a `.env`, pare e reinicie o servidor. Uma variável já definida no terminal tem prioridade sobre a `.env`; editar o arquivo sozinho não atualiza um processo já iniciado.

## Etapa K: voltar à demonstração sem e-mail

Na `.env`, altere:

```dotenv
AUTOMACAO_ORIGEM=local
```

Reinicie o servidor, se necessário. O programa passa a usar os quatro arquivos da pasta `entrada`, sem conectar ao Gmail. `AUTOMACAO_MODO_TESTE` escolhe as pastas e `AUTOMACAO_MOTOR` escolhe como atualizar a planilha; são três escolhas independentes.

## Etapa L: trocar senhas e resolver problemas

| Situação | O que fazer |
| --- | --- |
| “Senhas de app” não aparece | Confira a verificação em duas etapas e as restrições explicadas pelo Google; se indisponível, use a origem local |
| Gmail recusou o acesso | Confira endereço completo e senha de app. Não use a senha normal nem o código de confirmação do celular |
| Senha normal do Google foi alterada | As senhas de app podem ser revogadas. Gere uma nova e atualize somente sua `.env` |
| Nenhum e-mail correspondente | Confira remetente, assunto exato e se a mensagem está na Caixa de entrada |
| E-mail mais recente incompleto | Envie uma nova mensagem com os quatro arquivos corretos juntos |
| Planilha renomeada ou ZIP | Reenvie os quatro `.xlsx` com os nomes originais |
| Usuário do Dashboard já existe | Entre com a conta existente; o comando de criação não troca senhas |
| Alterei a senha na `.env`, mas o login não mudou | A `.env` só é lida pelo comando de criação. Para usuário existente, use `changepassword` |
| `no such table` | Execute `manage.py migrate` antes de criar usuário ou iniciar a aplicação |

Para trocar a senha de um usuário já criado, substitua `admin_demo` pelo nome escolhido:

```powershell
.\.venv\Scripts\python.exe manage.py changepassword admin_demo
```

Digite a nova senha duas vezes. Se mantiver uma senha de criação na `.env`, atualize-a ou remova-a para não guardar uma informação antiga. Depois de criar o administrador, o login consulta o banco, não esses campos da `.env`.

Para revogar o acesso do programa ao Gmail, exclua a senha de app correspondente em [Senhas de app](https://myaccount.google.com/apppasswords) e retire seu valor da `.env`. Não é necessário desligar a verificação em duas etapas.

## Etapa M: enviar o código ao portfólio sem enviar senhas

O arquivo público é `.env.example`, com campos vazios e valores fictícios. Suas credenciais ficam somente na `.env` local. Antes de preparar o commit, confira:

```powershell
git check-ignore .env
git status --short
```

O primeiro comando deve mostrar `.env`. O segundo não deve listar esse arquivo para envio. Não coloque senhas em código, README, prints ou URLs. Se algum segredo já tiver sido publicado, retire o acesso revogando-o/trocando-o; apenas apagar o arquivo de um novo commit não remove o segredo do histórico anterior.

**Conferência final:** Gmail com verificação em duas etapas, senha de app na `.env`, teste de conexão aprovado no seu computador, e-mail recebido com quatro anexos fictícios, Dashboard preparado, administrador criado, login funcionando e atualização concluída. Sem Gmail, use `local` e faça a mesma conferência a partir do preparo do Dashboard.

**Limites dos testes desta entrega:** o comportamento IMAP é exercitado com mensagens fictícias e servidor simulado nos testes. Nenhuma conta Gmail real foi acessada na preparação. A autenticação real depende da sua conta, da disponibilidade de senhas de app e da sua conexão. O Excel COM real também precisa ser validado no seu Windows.


# 6. ESCOLHER O MOTOR E EXECUTAR O ARQUIVO SOLICITADO

> **EXECUÇÃO DIRETA:** use a raiz do projeto e mantenha o servidor parado neste primeiro teste. O script não precisa de `runserver` nem de Waitress.

## Principal: Windows com Excel instalado

O `requirements.txt` instala `pywin32` somente no Windows. O Excel desktop deve estar instalado, ativado e abrir normalmente na sua conta Windows. Use esta versão em uma sessão local de desenvolvimento.

O `.env.example` já define:

```dotenv
AUTOMACAO_MOTOR=win32com
```

Execute o comando abaixo, que escolhe o motor principal explicitamente:

```bash
python automacao/executar_automacao.py --motor win32com
```

O programa cria uma instância própria do Excel com `DispatchEx`, abre uma cópia temporária do Dashboard, atualiza os dados e chama `CalculateFullRebuild` para recalcular as fórmulas. Após salvar e encerrar o Excel, substitui o Dashboard demonstrativo. As outras janelas do Excel do usuário não são utilizadas.

Se Windows, pywin32 ou Excel não estiverem disponíveis, a execução falha com uma mensagem no log. Não há mudança automática para `openpyxl`.

## Alternativa: executar sem Excel

```bash
python automacao/executar_automacao.py --motor openpyxl
```

Para que a interface Django também use essa alternativa, altere sua `.env` e reinicie o servidor:

```dotenv
AUTOMACAO_MOTOR=openpyxl
```

A opção `--motor` vale somente para aquela execução no terminal. A interface e o worker usam `AUTOMACAO_MOTOR` do ambiente ou da `.env`. Variáveis já existentes no ambiente têm prioridade sobre a `.env`.

## Comandos gerais


Com a `.venv` ativada e os modelos preparados:

```bash
python automacao/executar_automacao.py
```

Também é possível executar como módulo:

```bash
python -m automacao.executar_automacao
```

No PowerShell sem ativação:

```powershell
.\.venv\Scripts\python.exe automacao\executar_automacao.py
```

`executar_automacao.py` é um script Python, executado pelo interpretador. Ele não é um arquivo `.exe`.

Ao concluir, consulte:

- `TESTE_AUTOMACAO/dashboard/dashboard_exemplo.xlsx`: Dashboard atualizado.
- `TESTE_AUTOMACAO/saida/`: cópias dos quatro relatórios processados.
- `logs/automacao.log`: etapas, contagens e eventuais erros.

Sem `--motor`, a escolha vem de `AUTOMACAO_MOTOR`; na ausência da variável, o padrão é `win32com`.

O processo retorna código `0` em caso de sucesso e `1` em caso de falha. Para consultar o código imediatamente após executar, use `$LASTEXITCODE` no PowerShell ou `echo %ERRORLEVEL%` no CMD.

# 7. EXECUTAR PELA INTERFACE DJANGO

Com a `.venv` ativada:

> **ANTES DE LIGAR:** execute na raiz do projeto e confirme que não há outro servidor desta demonstração aberto. Para parar, aguarde a atualização terminar e pressione `Ctrl+C` no mesmo terminal.

```bash
python manage.py migrate
python manage.py createsuperuser
python manage.py runserver
```

Crie seu próprio usuário e senha no terminal; nenhum usuário pronto é distribuído no projeto.

Abra <http://127.0.0.1:8000/>, faça login com o superusuário e clique em **Atualizar Dashboard**. A interface acompanha um processo separado e grava o histórico. Ao concluir, clique em **Baixar Dashboard atualizado**.

A página mostra o motor selecionado. Antes de iniciar o servidor, confira `AUTOMACAO_MOTOR` na `.env`.

Com modo teste ligado, as pastas são fixas e não é necessário salvar caminhos. Com teste desligado, prepare `DEMO_CONFIGURADA` e salve os caminhos pela interface, conforme a seção abaixo.

Para parar o servidor, pressione `Ctrl+C` no terminal.

O servidor `runserver` serve para desenvolvimento local. Publicar o código em um repositório não hospeda automaticamente esta aplicação.

## Usuários e agendamento

O código preserva cadastro, aprovação, recuperação de senha, permissões e histórico do projeto. Usuários comuns podem consultar a interface e baixar o arquivo; a execução é restrita a administradores.

As telas de preferências de agendamento foram mantidas. Com modo teste ligado, a execução agendada é bloqueada. Com teste desligado, o comando de agendamento avalia as preferências e os caminhos configurados; salvar preferências não instala tarefas no Windows. Para começar, use a execução manual.

# 7.1. LIGAR E DESLIGAR O MODO TESTE EM SETTINGS.PY

> **ANTES DE ALTERAR:** aguarde a automação terminar e pare `runserver` ou Waitress com `Ctrl+C`. Depois das alterações, reinicie o servidor escolhido.

O arquivo `projeto_web/settings.py` possui esta linha:

```python
AUTOMACAO_MODO_TESTE = os.environ.get("AUTOMACAO_MODO_TESTE", "1") == "1"
```

Ela lê uma variável de ambiente. O texto `"1"` vira `True` (ligado); `"0"` vira `False` (desligado). Na ausência da variável, o padrão é `"1"`. A `.env` é carregada antes dessa leitura. Variáveis já definidas no terminal têm prioridade sobre a `.env`.

| Estado | Valor na `.env` | Resultado em settings | Pastas utilizadas |
| --- | --- | --- | --- |
| Teste ligado | `AUTOMACAO_MODO_TESTE=1` | `True` | `TESTE_AUTOMACAO` automaticamente |
| Teste desligado | `AUTOMACAO_MODO_TESTE=0` | `False` | Caminhos salvos na interface, dentro de `DEMO_CONFIGURADA` |

**Desligar o teste não ativa a leitura de e-mail por si só. `AUTOMACAO_ORIGEM=gmail` é a escolha separada que habilita o Gmail pessoal. Os dois estados de teste usam os mesmos modelos genéricos de exemplo.** `DJANGO_DEBUG` controla a exibição de erros do Django e é independente de `AUTOMACAO_MODO_TESTE`.

## Ligar o teste

1. Pare o servidor com `Ctrl+C` e aguarde qualquer atualização terminar.
2. Abra a `.env` e defina `AUTOMACAO_MODO_TESTE=1`.
3. Execute um comando por linha:

> **ANTES DE LIGAR:** execute na raiz do projeto e confirme que não há outro servidor desta demonstração aberto. Para parar, aguarde a atualização terminar e pressione `Ctrl+C` no mesmo terminal.

```powershell
python preparar_demo.py
python manage.py runserver
```

A interface utiliza `TESTE_AUTOMACAO/dashboard/dashboard_exemplo.xlsx`, `TESTE_AUTOMACAO/entrada` e `TESTE_AUTOMACAO/saida`. As configurações salvas no banco não são usadas nesse estado.

## Desligar o teste e usar o ambiente configurado

1. Pare o servidor e aguarde a atualização terminar.
2. Abra a `.env` e defina `AUTOMACAO_MODO_TESTE=0`.
3. Prepare outra cópia fictícia e reinicie:

> **ANTES DE LIGAR:** execute na raiz do projeto e confirme que não há outro servidor desta demonstração aberto. Para parar, aguarde a atualização terminar e pressione `Ctrl+C` no mesmo terminal.

```powershell
python preparar_demo.py --ambiente configurado
python manage.py runserver
```

4. Entre como administrador. O formulário de caminhos ficará disponível. Salve os três caminhos absolutos que correspondem à pasta atual do projeto:

| Campo | Caminho de exemplo |
| --- | --- |
| Dashboard | `C:\Projetos\automacao-dashboard-portfolio\DEMO_CONFIGURADA\dashboard\dashboard_exemplo.xlsx` |
| Entrada | `C:\Projetos\automacao-dashboard-portfolio\DEMO_CONFIGURADA\entrada` |
| Saída | `C:\Projetos\automacao-dashboard-portfolio\DEMO_CONFIGURADA\saida` |

Substitua `C:\Projetos\automacao-dashboard-portfolio` pela pasta real onde você extraiu este pacote. Clique em **Salvar caminhos** e depois em **Atualizar Dashboard**. Essa validação também se aplica ao download.

Na execução direta pelo terminal, a seleção vem de `settings.py` e os caminhos padrão são os da pasta correspondente ao modo. A CLI não lê os caminhos salvos no banco.

## Se preferir editar o booleano diretamente em settings.py

É possível substituir a linha de leitura por uma das opções abaixo, usando apenas uma por vez:

```python
AUTOMACAO_MODO_TESTE = True  # Teste ligado.
```

ou:

```python
AUTOMACAO_MODO_TESTE = False  # Teste desligado; caminhos fictícios configurados.
```

Nesse caso, a variável da `.env` deixa de controlar essa opção. Para voltar ao controle pela `.env`, restaure a linha original com `os.environ.get`. A CLI e a interface usam o valor de `settings.py`. Reinicie o servidor após qualquer alteração.

## Se a mudança na .env não surtir efeito

Confira se o PowerShell já possui uma variável com o mesmo nome:

```powershell
Get-Item Env:AUTOMACAO_MODO_TESTE -ErrorAction SilentlyContinue
```

Para remover somente essa variável da sessão atual e voltar a ler a `.env`:

```powershell
Remove-Item Env:AUTOMACAO_MODO_TESTE -ErrorAction SilentlyContinue
```

Depois reinicie o servidor. A escolha do motor (`AUTOMACAO_MOTOR`) é outra configuração: teste ligado ou desligado pode usar `win32com` ou `openpyxl`.

# 7.2. DASHBOARD FICTÍCIO E ARQUIVOS ANTIGOS

O arquivo de execução fornecido por este pacote é **`dados_demo/dashboard_exemplo.xlsx`**. Ele contém as abas RESUMO, VENDAS, RECEBIMENTOS, CUSTOS e DESPESAS, com registros, fornecedores, códigos e valores sintéticos. As áreas foram organizadas para ilustrar a importação de relatórios genéricos, com clientes, fornecedores e valores inventados.

Use apenas o Dashboard e os relatórios de exemplo fornecidos. Arquivos de outros projetos não devem ser colocados nas pastas desta demonstração; renomear um arquivo não muda seu conteúdo.

Se você já extraiu outro projeto nessa pasta, prefira extrair o novo ZIP em **uma pasta nova**. O preparo agora identifica planilhas extras em `entrada`, `saida` e `dashboard` e interrompe a operação. Mova os arquivos antigos para fora do ambiente demonstrativo; o programa não os apaga.

Para substituir os cinco arquivos de trabalho conhecidos pelos modelos fictícios atualizados, execute:

```powershell
python preparar_demo.py --restaurar
```

Para restaurar somente o ambiente configurado:

```powershell
python preparar_demo.py --ambiente configurado --restaurar
```

Esses comandos não apagam o banco de usuários nem o histórico. Eles não removem planilhas extras: se o preparo apontar um arquivo antigo, mova-o antes e execute o comando novamente.

# 8. QUATRO RELATÓRIOS E UM DASHBOARD

Os quatro arquivos de origem ficam em `dados_demo/`. O quinto arquivo é o Dashboard central que recebe os dados.

| Origem | Registros fictícios | Colunas na origem | Linhas iniciais ignoradas | Destino no Dashboard |
| --- | ---: | ---: | ---: | --- |
| [vendas_exemplo.xlsx](dados_demo/vendas_exemplo.xlsx) | 6 | 22 | 3 | VENDAS, C8:X13 |
| [recebimentos_exemplo.xlsx](dados_demo/recebimentos_exemplo.xlsx) | 8 | 28 | 3 | RECEBIMENTOS, A8:O15 e V8:AB15 |
| [custos_exemplo.xlsx](dados_demo/custos_exemplo.xlsx) | 5 | 16 | 3 | CUSTOS, C8:R12 |
| [despesas_exemplo.xlsx](dados_demo/despesas_exemplo.xlsx) | 4 | 15 | 3 | DESPESAS, A8:N11 |

O modelo de Dashboard começa com dois registros sintéticos em cada aba, com identificadores `BASE-DEMO-0001` e `BASE-DEMO-0002` e valores de exemplo. A primeira execução demonstra o crescimento da área de dados. As próximas execuções substituem os mesmos dados sem duplicação.

Regras principais:

- VENDAS importa as 22 colunas; os identificadores permanecem texto, inclusive zeros à esquerda.
- RECEBIMENTOS importa A:O e V:AB. P:U ficam reservadas no destino. O modelo utiliza o marcador `Reservado` nessas células.
- CUSTOS importa 16 colunas e não possui TOTAL no corpo do destino.
- DESPESAS descarta a coluna M da origem e importa A:L e N:O para A:N do destino.
- A origem tem uma linha TOTAL com valores numéricos fictícios. Ela é removida antes da importação; os totais de destino são reconstruídos.
- Na linha 5 de cada aba, as colunas monetárias recebem `SUBTOTAL`. As linhas TOTAL de VENDAS e RECEBIMENTOS recebem `SUM`.
- RESUMO mostra quantidades de registros e o saldo de cada relatório. Os saldos são medidas de relatórios distintos; não devem ser somados entre si.

Depois da primeira atualização, os saldos esperados são: VENDAS, R$ 6.300,00; RECEBIMENTOS, R$ 10.800,00; CUSTOS, R$ 6.000,00; DESPESAS, R$ 6.000,00.

**Com `win32com`**, o Excel recalcula e o motor confere os subtotais, totais e saldos calculados antes de salvar. **Com `openpyxl`**, as fórmulas são gravadas, mas não recalculadas: abra o Dashboard em um aplicativo de planilhas com cálculo automático para atualizar os resultados exibidos. A auditoria dessa alternativa confere os dados escritos diretamente.

# 9. AS 17 ETAPAS

O fluxo possui 17 etapas, com o encerramento de cada recurso identificado explicitamente:

| Etapa | Procedimento na demonstração |
| ---: | --- |
| 01 | Validar o ambiente; preparar origem local ou autenticar no Gmail pessoal |
| 02 | Obter os quatro anexos da origem escolhida e preparar a pasta de processamento |
| 03 | Ler os relatórios com pandas |
| 04 | Preparar DataFrames, selecionar colunas e remover TOTAL |
| 05 | Validar quantidades, identificadores e valores |
| 06 | Principal: abrir cópia no Excel com `win32com`. Alternativa: abrir com `openpyxl` |
| 07 | Localizar TOTAL ou a última linha de cada aba |
| 08 | Calcular quantidades e diferenças de linhas |
| 09 | Ajustar a área de registros |
| 10 | Limpar as colunas de dados importados |
| 11 | Escrever registros preservando P:U em RECEBIMENTOS |
| 12 | Reconstruir linhas TOTAL |
| 13 | Atualizar fórmulas e resumo; no principal, recalcular pelo Excel. Auditar antes de salvar |
| 14 | Salvar cópia temporária; no principal, a substituição é confirmada após fechar o Excel |
| 15 | Fechar o workbook |
| 16 | Principal: encerrar a instância própria do Excel e liberar COM. Alternativa: liberar motor local |
| 17 | Encerrar a origem local ou a sessão IMAP, se utilizada |

Com origem `local`, não há conexão de e-mail. Com origem `gmail`, as etapas 01/02/17 autenticam, baixam os anexos e encerram a sessão de leitura. No motor principal, as operações de Excel e COM são reais. Na alternativa `openpyxl`, nenhum Excel é iniciado.

# 10. REPETIR E RESTAURAR A DEMONSTRAÇÃO

> **ANTES DE ALTERAR:** aguarde a automação terminar e pare `runserver` ou Waitress com `Ctrl+C`. Depois das alterações, reinicie o servidor escolhido.

Execute novamente o executor para repetir a atualização. Na etapa 02, os relatórios de `entrada` substituem as cópias em `saida`. Para experimentar mudanças, edite os arquivos de `TESTE_AUTOMACAO/entrada`, mantendo o contrato de colunas.

Para restaurar os quatro relatórios e o Dashboard aos modelos iniciais:

```bash
python preparar_demo.py --restaurar
```

Esse comando substitui os cinco arquivos conhecidos do ambiente de demonstração e remove suas quatro cópias na saída. Ele não apaga usuários, histórico ou outros arquivos. Sem `--restaurar`, o preparo mantém arquivos já existentes.

# 11. VERIFICAÇÕES

```bash
python manage.py check
python manage.py test -v 1
```

Também é possível executar somente os testes do motor, sem criar o banco do Django:

```bash
python -m unittest discover -s tests -v
```

Os testes do motor cobrem o fluxo real, repetição, redução de registros, preservação de P:U, identificadores textuais e falhas de validação. Os testes Django cobrem autenticação, permissões, cadastro, histórico, recuperação de senha e integração.

## Teste real do Excel no Windows

A suíte comum utiliza o motor `openpyxl` para a integração de planilhas e mocks para verificar seleção e limpeza de recursos COM. Ela não comprova que o Excel instalado funciona. Há um teste separado, optativo, que abre o Excel real e confere recálculo, repetição, redução de registros e preservação de P:U.

No PowerShell, com a `.venv` ativada:

```powershell
$env:EXECUTAR_TESTE_EXCEL = "1"
python -m unittest tests.test_win32com_real -v
Remove-Item Env:EXECUTAR_TESTE_EXCEL
```

No CMD:

```bat
set EXECUTAR_TESTE_EXCEL=1
python -m unittest tests.test_win32com_real -v
set EXECUTAR_TESTE_EXCEL=
```

Esse teste trabalha em uma pasta temporária com cópias fictícias. Sem Windows e sem a variável de autorização, ele aparece como `skipped`. Com a variável ativa no Windows, a ausência de Excel/pywin32 causa falha, indicando que o ambiente ainda precisa ser preparado.

**Limite desta entrega:** o Excel COM real não foi executado no ambiente Linux usado na preparação. Execute esse teste no seu Windows antes de apresentar o motor principal como validado no seu computador.

Os testes com erro proposital registram mensagens de falha no log. O resultado final `OK` do executor de testes informa se os comportamentos esperados passaram.

# 12. PUBLICAR NA SUA PRÓPRIA CONTA GIT

Git é a ferramenta de versionamento. GitHub e GitLab hospedam os repositórios. O projeto pode ser enviado para qualquer um dos dois.

Use esta pasta nova, sem copiar a pasta `.git` do projeto anterior. Assim, o portfólio começa com um histórico independente.

## Criar um repositório novo

Entre na sua conta pessoal no GitHub ou GitLab. Crie um repositório vazio chamado, por exemplo, `automacao-dashboard-portfolio`. Não inicialize o repositório remoto com README ou outros arquivos, pois esta pasta já os possui. Copie a URL HTTPS retornada pelo serviço.

## Configurar e revisar a cópia local

Na raiz desta versão:

```bash
git init -b main
git config --local user.name "SEU NOME"
git config --local user.email "SEU EMAIL DE COMMITS"
git config --local credential.useHttpPath true
git status --short
git check-ignore .env db.sqlite3 TESTE_AUTOMACAO/dashboard/dashboard_exemplo.xlsx
```

Substitua os marcadores de nome e e-mail pelos dados da sua conta pessoal. Você pode usar o e-mail privado de commits disponibilizado pelo serviço. `--local` configura apenas este repositório.

`user.name` e `user.email` definem a autoria dos commits. Eles não fazem login nem trocam a conta autenticada. O envio precisa autenticar na conta pessoal escolhida. `credential.useHttpPath` ajuda o gerenciador de credenciais a distinguir repositórios, mas não garante a escolha da conta por si só.

## Fazer o primeiro commit

```bash
git add .
git diff --cached --stat
git diff --cached --name-only
git commit -m "Adiciona demonstracao de automacao com dados ficticios"
```

Na lista de arquivos preparados, devem aparecer o código, a documentação, `.env.example` e os cinco modelos em `dados_demo/`. `.env`, banco SQLite, logs, ambiente virtual e arquivos de trabalho não devem aparecer.

## Enviar para a conta pessoal

No GitHub, a forma da URL é:

```bash
git remote add origin https://github.com/SEU-USUARIO/automacao-dashboard-portfolio.git
git remote -v
git push -u origin main
```

No GitLab, use a URL exata fornecida pelo seu repositório pessoal, substituindo a URL acima. Confira se a URL corresponde ao repositório escolhido.

Ao solicitar autenticação, selecione sua conta pessoal. Não coloque senha ou token dentro da URL nem nos arquivos do projeto. Caso o Windows reutilize outra conta, resolva a autenticação antes de tentar novo envio; não apague credenciais de outras contas indiscriminadamente.

Se a cópia já tiver um `origin`, confira `git remote -v` antes de alterar. Esta entrega não traz remotos configurados.

## Usar Source Control do VS Code

1. Abra somente a pasta desta versão do portfólio.
2. Se necessário, selecione **Initialize Repository / Inicializar Repositório**.
3. Configure a identidade local pelos comandos anteriores.
4. Em **Source Control**, revise a lista e clique no `+` para preparar as alterações.
5. Digite a mensagem do commit e confirme o commit.
6. Configure `origin` com a URL pessoal e confira `git remote -v` no terminal.
7. Faça **Push / Sincronizar**, conferindo a conta solicitada na autenticação.

O botão **Publish to GitHub** pode utilizar a conta conectada ao VS Code. Confira a conta antes de publicar.

# 13. ERROS COMUNS

| Mensagem ou situação | Como resolver |
| --- | --- |
| `No module named django` ou outra biblioteca | Use o Python da `.venv` e instale `requirements.txt` |
| Planilhas extras no ambiente | Mova planilhas antigas/reais para outra pasta e prepare novamente; prefira uma extração limpa |
| Arquivos da demonstração ausentes | Execute `python preparar_demo.py` |
| `can't open file ... preparar_demo.pycls` | Use `python .\preparar_demo.py`; execute `cls` separadamente |
| `no such table` | Execute `python manage.py migrate` |
| Sem permissão para atualizar | Entre com o superusuário criado por você |
| `PermissionError` ao salvar | Feche o Dashboard no Excel e tente novamente |
| `win32com exige Windows` | Execute no Windows com Excel ou escolha `--motor openpyxl` explicitamente |
| Excel não abre / erro COM | Confira a instalação e ativação do Excel desktop e execute o teste real no Windows |
| Motor inválido | Use exatamente `win32com` ou `openpyxl` em `AUTOMACAO_MOTOR` |
| Fórmulas ainda não mostram valores | Abra a planilha com cálculo automático habilitado |
| Relatório com quantidade incorreta de colunas | Restaure os modelos ou confira `automacao/config.py` |
| Tela informa execução em andamento após interrupção forçada | Pare os processos e consulte o registro em `/admin/`; uma interrupção forçada pode deixar histórico pendente |
| `remote origin already exists` | Confira `git remote -v`; não sobrescreva o remoto sem conferir |
| `403` no push | Confira a conta autenticada, a URL e o acesso ao repositório pessoal |

# 14. COMO ESTUDAR O PROJETO

Comece por `preparar_demo.py`: acompanhe como `Path` e `shutil.copyfile` localizam e copiam os modelos. Depois leia `automacao/config.py`, relacionando as posições definidas com as planilhas. Leia o pequeno roteador `executar_automacao.py` e veja como ele seleciona um dos dois motores.

Em seguida, leia `preparar_dataframe` em `motor_openpyxl.py` (preparação compartilhada pelos dois motores) e acompanhe a transformação dos dados. Só depois percorra as etapas de abertura, escrita, auditoria e salvamento. Por último, siga o caminho `urls.py → views.py → executar_automacao` para entender a integração com Django.

Sugestão de exercício: após preparar a demonstração, altere um valor monetário em um relatório da pasta `entrada`, execute novamente e confira a célula correspondente e o subtotal no Dashboard.

# DOCUMENTAÇÃO TÉCNICA E LICENCIAMENTO

A [documentação técnica do projeto](docs/DOCUMENTACAO_TECNICA.md) reúne arquitetura, responsabilidades dos módulos, requisitos, contratos das planilhas, banco, rotas, processamento, configuração, logs, testes, limitações e implantação pretendida. O README permanece como guia operacional para iniciantes.

**Licença ainda não definida:** esta distribuição não recebeu um arquivo `LICENSE`. A publicação em repositório público não torna o software automaticamente open source; essa classificação depende também dos termos de licenciamento. A escolha de licença será tratada separadamente. Esta atualização de nomes e dados não adiciona uma licença nem altera permissões de uso.

Referências: [OSI](https://opensource.org/osd) e [GitHub: licenciamento](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/licensing-a-repository).

# REFERÊNCIAS

- [Waitress: execução e opções](https://docs.pylonsproject.org/projects/waitress/en/latest/runner.html)
- [Django: implantação e servidores](https://docs.djangoproject.com/en/5.2/howto/deployment/)
- [Django: primeiro projeto](https://docs.djangoproject.com/en/5.2/intro/tutorial01/)
- [GitHub: adicionar código local](https://docs.github.com/en/migrations/importing-source-code/using-the-command-line-to-import-source-code/adding-locally-hosted-code-to-github)
- [Git: configuração](https://git-scm.com/docs/git-config)
- [pandas: leitura de Excel](https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html)
- [Excel: CalculateFullRebuild](https://learn.microsoft.com/pt-br/office/vba/api/excel.application.calculatefullrebuild)
- [pywin32: projeto oficial](https://github.com/mhammond/pywin32)
- [openpyxl: uso básico](https://openpyxl.readthedocs.io/en/stable/tutorial.html)
