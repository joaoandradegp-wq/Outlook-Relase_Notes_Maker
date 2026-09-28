<h1 align="center">📑 Outlook Release Notes Maker</h1>

<p align="center">
Gerador automático de <b>Release Notes / GMUD</b> a partir de e-mails do Outlook, transformando solicitações de aprovação e registros de origem em uma apresentação PowerPoint pronta para utilização.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Status-Em%20Evolução-brightgreen">
  <img src="https://img.shields.io/badge/Linguagem-Node.js-green">
  <img src="https://img.shields.io/badge/Automação-Outlook-blue">
  <img src="https://img.shields.io/badge/Output-PowerPoint-orange">
  <img src="https://img.shields.io/badge/Python-PPTX-yellow">
  <img src="https://img.shields.io/badge/Language-PT--BR-lightgrey">
</p>

---

<h2>📌 Sobre</h2>

<p>
O <b>GMUD Release Notes Maker</b> é uma ferramenta desenvolvida para automatizar a criação de apresentações de <b>GMUD</b> a partir das informações disponíveis no <b>Outlook Web</b>.
</p>

<p>
O sistema acessa a caixa de e-mail através do navegador, localiza as mensagens relacionadas à aprovação de negócios GMUD, identifica o período desejado, extrai as demandas presentes no campo <b>Registro de origem</b> e captura os respectivos pedidos e respostas de aprovação.
</p>

<p>
Após a coleta das informações, os dados são enviados para um script Python responsável por preencher um template <b>PowerPoint (.pptx)</b>, mantendo a identidade visual definida no modelo.
</p>

<p align="center">
  <b>Outlook → E-mails GMUD → Registro de origem → Demandas → Aprovações → PowerPoint</b>
</p>

---

<h2>🔄 Funcionamento</h2>

<pre>
┌──────────────┐     ┌───────────────┐     ┌─────────────────┐     ┌───────────────┐
│ Outlook Web  │ ──► │   index.js    │ ──► │ E-mails GMUD    │ ──► │ Registro de   │
│              │     │   Playwright  │     │ e Aprovações    │     │ origem        │
└──────────────┘     └───────────────┘     └─────────────────┘     └───────┬───────┘
                                                                           │
                                                                           ▼
                                                                  ┌─────────────────┐
                                                                  │   Demandas +    │
                                                                  │   Prints        │
                                                                  └────────┬────────┘
                                                                           │
                                                                           ▼
                                                                  ┌─────────────────┐
                                                                  │ gerar_pptx.py   │
                                                                  │ Python / PPTX   │
                                                                  └────────┬────────┘
                                                                           │
                                                                           ▼
                                                                  ┌─────────────────┐
                                                                  │   GMUD-AAAA-MM  │
                                                                  │      .pptx      │
                                                                  └─────────────────┘
</pre>

<p>
O fluxo completo é iniciado pelo arquivo <b>gerar_releasenotes.bat</b>. O usuário informa o mês desejado e o script executa o <b>index.js</b>.
</p>

<p>
O Node.js abre o Outlook Web através do Playwright em modo visual, permitindo que o usuário faça login manualmente quando necessário. Depois da autenticação, a automação executa a pesquisa, processa os e-mails encontrados e chama o Python para gerar a apresentação.
</p>

---

<h2>🚀 Principais funcionalidades</h2>

<ul>
  <li>📧 Pesquisa automática de e-mails no Outlook Web</li>
  <li>📅 Seleção do mês de referência da GMUD</li>
  <li>🔎 Busca por assunto e remetente configurados</li>
  <li>📋 Leitura automática do campo <b>Registro de origem</b></li>
  <li>🧩 Extração das demandas / User Stories</li>
  <li>🔄 Suporte a conteúdo HTML e fallback para texto puro</li>
  <li>📨 Abertura automática do fluxo de <b>Encaminhar</b></li>
  <li>📜 Expansão de conteúdo anterior ou colapsado do e-mail</li>
  <li>🔎 Busca da tabela de implantação GMUD</li>
  <li>📸 Captura de screenshots dos e-mails processados</li>
  <li>📤 Busca do pedido original de aprovação em <b>Itens Enviados</b></li>
  <li>🖼️ Captura dos prints do pedido e da resposta de aprovação</li>
  <li>📊 Geração automática de PowerPoint</li>
  <li>📑 Criação dinâmica de slides de demandas</li>
  <li>📝 Normalização automática dos títulos das User Stories</li>
  <li>🖼️ Inserção automática das imagens de aprovação no template</li>
  <li>🚫 Remoção de slides sem conteúdo</li>
  <li>🧪 Geração de arquivos de diagnóstico quando a extração falha</li>
  <li>💾 Utilização de sessão persistente do navegador</li>
  <li>⚙️ Configuração centralizada no código</li>
</ul>

---

<h2>📂 Estrutura do projeto</h2>

<pre>
GMUD-Release_Notes/
│
├── gerar_releasenotes.bat
├── index.js
├── gerar_pptx.py
│
├── templates/
│   └── release_para_dev.pptx
│
├── .owa-session/
│   └── sessão persistente do Outlook
│
└── output/
    ├── GMUD-AAAA-MM.pptx
    ├── config-pptx.json
    ├── resposta-0.png
    ├── pedido-0.png
    ├── debug-rascunho-0.png
    └── debug-rascunho-0.txt
</pre>

<p>
O arquivo <b>index.js</b> concentra a automação do Outlook, extração das informações e orquestração do processo.
</p>

<p>
O arquivo <b>gerar_pptx.py</b> é responsável exclusivamente pela montagem da apresentação a partir do template.
</p>

<p>
A pasta <b>templates</b> contém o modelo visual utilizado na geração do PowerPoint, enquanto a pasta <b>output</b> recebe os arquivos produzidos durante a execução.
</p>

---

<h2>🧩 Componentes</h2>

<h3>📧 index.js</h3>

<p>
É o componente principal da aplicação e responsável por controlar todo o fluxo de automação.
</p>

<ul>
  <li>🌐 Inicialização do navegador com Playwright</li>
  <li>🔐 Acesso ao Outlook Web</li>
  <li>🔎 Pesquisa dos e-mails</li>
  <li>📅 Definição do período de análise</li>
  <li>📋 Extração das demandas</li>
  <li>📸 Captura dos screenshots</li>
  <li>📤 Busca dos pedidos de aprovação</li>
  <li>📦 Montagem dos dados para o PowerPoint</li>
  <li>🐍 Execução do script Python</li>
</ul>

<p>
O navegador é iniciado com uma sessão persistente em <b>.owa-session</b>, permitindo preservar o estado do navegador entre execuções.
</p>

<h3>🤖 Playwright</h3>

<p>
O Playwright é utilizado para controlar o Outlook Web através do navegador.
</p>

<p>
A automação utiliza seletores compatíveis com interfaces em português e inglês, permitindo localizar elementos como:
</p>

<ul>
  <li>🔎 Pesquisa / Search</li>
  <li>📤 Encaminhar / Forward</li>
  <li>📨 Enviar / Send</li>
  <li>🗑️ Descartar / Discard</li>
  <li>❌ Fechar / Close</li>
  <li>✅ OK / Yes</li>
  <li>📁 Itens Enviados / Sent Items</li>
</ul>

<p>
Essa estratégia reduz a dependência de uma única tradução da interface do Outlook.
</p>

<h3>🐍 gerar_pptx.py</h3>

<p>
Responsável pela geração da apresentação PowerPoint utilizando <b>python-pptx</b> e <b>Pillow</b>.
</p>

<p>
O script recebe um arquivo JSON intermediário contendo o template, o nome do arquivo de saída, o período, as demandas e as aprovações.
</p>

---

<h2>📅 Seleção do período</h2>

<p>
O arquivo <b>gerar_releasenotes.bat</b> solicita ao usuário o número do mês que será utilizado na geração do relatório.
</p>

<pre>
==========================================
  GMUD - Gerar PPTX
==========================================

Digite o numero do mes (1 a 12).
Ex: 9 = Setembro.

Mes:
</pre>

<p>
O valor informado precisa estar entre <b>1 e 12</b>. Caso contrário, o script solicita novamente.
</p>

<p>
O mês é enviado para o Node.js através da variável de ambiente <b>GMUD_MES</b>.
</p>

<h3>📆 Definição do ano</h3>

<p>
Quando o ano não é informado através da variável <b>GMUD_ANO</b>, o sistema determina automaticamente o ano de referência.
</p>

<ul>
  <li>📅 Mês já ocorrido → ano atual</li>
  <li>🔄 Mês ainda não ocorrido → ano anterior</li>
</ul>

<p>
Também é possível definir explicitamente o ano utilizando:
</p>

<pre>
set GMUD_MES=9
set GMUD_ANO=2026
</pre>

---

<h2>🔎 Pesquisa dos e-mails</h2>

<p>
Após o login no Outlook, o sistema monta automaticamente uma consulta utilizando:
</p>

<ul>
  <li>📌 Assunto contendo <b>de acordo de negocios gmud</b></li>
  <li>👤 Remetente configurado</li>
  <li>📅 Data inicial do mês selecionado</li>
  <li>📅 Data final do mês selecionado</li>
</ul>

<p>
A consulta considera tanto mensagens recebidas quanto mensagens enviadas dentro do período.
</p>

<pre>
subject:"de acordo de negocios gmud"
from:aprovador@suaempresa.com.br
(received:DD/MM/AAAA..DD/MM/AAAA
 OR sent:DD/MM/AAAA..DD/MM/AAAA)
</pre>

<p>
A quantidade de e-mails encontrados é exibida no console antes do processamento individual.
</p>

---

<h2>📋 Extração das demandas</h2>

<p>
Para cada e-mail encontrado, o sistema procura o conteúdo associado ao campo <b>Registro de origem</b>.
</p>

<p>
A primeira tentativa é realizada diretamente na estrutura HTML das tabelas do e-mail.
</p>

<p>
Quando o conteúdo HTML não é encontrado, o sistema utiliza um <b>fallback</b> que procura o texto diretamente no corpo da mensagem.
</p>

<h3>🔄 Estratégia de fallback</h3>

<pre>
E-mail
  │
  ├──► Tabela HTML
  │       │
  │       └──► Registro de origem
  │
  └──► Texto puro
          │
          └──► Registro de origem
</pre>

<p>
As linhas de demanda são identificadas através de um padrão que aceita:
</p>

<ul>
  <li>🔢 Número com pelo menos quatro dígitos</li>
  <li>📝 Prefixo opcional <b>User Story</b></li>
  <li>🔤 Separador <b>:</b> ou <b>-</b></li>
</ul>

<p>
Exemplos reconhecidos:
</p>

<pre>
12345: Atualização do processo
User Story 12345: Nova funcionalidade
12345 - Correção de integração
</pre>

---

<h2>📤 Processamento através de Encaminhar</h2>

<p>
Quando o e-mail é aberto, o sistema utiliza o botão <b>Encaminhar</b> para acessar uma representação completa do conteúdo da conversa.
</p>

<p>
O fluxo possui diversas tentativas para localizar e executar o botão:
</p>

<ul>
  <li>🔘 Menu item em português</li>
  <li>🔘 Botão em português</li>
  <li>🔘 Menu item em inglês</li>
  <li>🔘 Botão em inglês</li>
  <li>🔘 Busca por texto</li>
  <li>🖱️ Clique forçado</li>
  <li>🖱️ Clique nativo via DOM</li>
</ul>

<p>
Depois que o rascunho é aberto, o sistema aguarda o carregamento da tabela e a estabilização do conteúdo antes de realizar a extração.
</p>

---

<h2>🔄 Conteúdo colapsado</h2>

<p>
Conversas antigas do Outlook podem apresentar partes do histórico ocultas atrás de controles como <b>...</b> ou opções de exibição de conteúdo anterior.
</p>

<p>
O sistema tenta expandir automaticamente esses conteúdos antes de procurar o <b>Registro de origem</b>.
</p>

<ul>
  <li>🔎 Procura o controle de conteúdo anterior</li>
  <li>🖱️ Executa o clique</li>
  <li>⏳ Aguarda o corpo estabilizar</li>
  <li>🔁 Repete o processo até o limite configurado</li>
</ul>

<p>
O número máximo de rodadas de expansão é controlado por:
</p>

<pre>
rodadasExpansao: 3
</pre>

---

<h2>⏳ Controle de carregamento</h2>

<p>
Como o Outlook Web carrega conteúdo dinamicamente, o sistema possui tempos específicos para cada etapa.
</p>

<ul>
  <li><b>botaoEncaminhar</b>: tempo máximo para encontrar o botão Encaminhar</li>
  <li><b>rascunhoAbrir</b>: tempo máximo para o rascunho aparecer</li>
  <li><b>tabelaAparecer</b>: tempo máximo para a tabela carregar</li>
  <li><b>corpoEstavelPor</b>: tempo que o conteúdo precisa permanecer estável</li>
  <li><b>corpoEstabilizar</b>: tempo máximo para estabilização</li>
  <li><b>registroOrigem</b>: tempo máximo para encontrar o Registro de origem</li>
  <li><b>rodadasExpansao</b>: quantidade máxima de tentativas de expansão</li>
</ul>

<p>
Esses valores podem ser ajustados diretamente no bloco <b>CONFIG.tempos</b> do <b>index.js</b>.
</p>

---

<h2>📸 Screenshots</h2>

<p>
Durante o processamento, o sistema salva screenshots dos e-mails para utilização no relatório e também para diagnóstico.
</p>

<h3>📨 Resposta</h3>

<pre>
output/resposta-0.png
</pre>

<p>
Representa o e-mail processado e é utilizado como base para a etapa de aprovação.
</p>

<h3>📤 Pedido de aprovação</h3>

<pre>
output/pedido-0.png
</pre>

<p>
O sistema acessa a pasta <b>Itens Enviados</b>, procura uma mensagem com o mesmo assunto do e-mail processado e captura o pedido original.
</p>

---

<h2>🧪 Diagnóstico</h2>

<p>
Quando o sistema não consegue localizar o <b>Registro de origem</b>, são gerados automaticamente arquivos de diagnóstico.
</p>

<pre>
output/
├── debug-rascunho-0.png
└── debug-rascunho-0.txt
</pre>

<p>
O arquivo PNG registra o estado visual da página.
</p>

<p>
O arquivo TXT contém o texto encontrado nos diferentes contextos da página, facilitando a análise de problemas de carregamento ou mudanças na estrutura do Outlook.
</p>

<p>
Essa camada de diagnóstico é especialmente útil quando o Outlook altera a estrutura HTML da interface ou quando o conteúdo da conversa demora mais que o tempo configurado para aparecer.
</p>

---

<h2>📨 Busca do pedido de aprovação</h2>

<p>
Depois que as demandas são extraídas, o sistema acessa <b>Itens Enviados</b> para localizar o pedido de aprovação correspondente.
</p>

<p>
A pesquisa utiliza o mesmo assunto identificado no e-mail processado.
</p>

<pre>
subject:"Assunto do e-mail"
</pre>

<p>
Quando a mensagem é localizada, o sistema abre o primeiro resultado e salva:
</p>

<pre>
output/pedido-0.png
</pre>

<p>
O slide de aprovação somente recebe conteúdo quando existem os dois arquivos:
</p>

<ul>
  <li>📤 Print do pedido</li>
  <li>📨 Print da resposta</li>
</ul>

---

<h2>📊 Geração do PowerPoint</h2>

<p>
Após o processamento dos e-mails, o Node.js consolida as informações em um arquivo JSON intermediário.
</p>

<pre>
output/config-pptx.json
</pre>

<p>
O arquivo contém:
</p>

<ul>
  <li><b>template</b>: caminho do template PPTX</li>
  <li><b>output</b>: caminho do arquivo final</li>
  <li><b>mesAno</b>: período exibido na apresentação</li>
  <li><b>demandas</b>: lista de demandas extraídas</li>
  <li><b>aprovacoes</b>: pares de imagens de pedido e resposta</li>
</ul>

<p>
O Node.js executa então o script Python passando esse JSON como argumento.
</p>

<pre>
python gerar_pptx.py output/config-pptx.json
</pre>

---

<h2>🎨 Template PowerPoint</h2>

<p>
A apresentação é construída a partir de um template existente, configurado em:
</p>

<pre>
templates/release_para_dev.pptx
</pre>

<p>
O template precisa possuir quatro slides-base:
</p>

<ul>
  <li>📌 Slide 1: capa</li>
  <li>📋 Slide 2: lista de demandas</li>
  <li>📝 Slide 3
