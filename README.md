<h1 align="center">📧 Outlook Release Notes Maker</h1>

<p align="center">
Automação para consulta de aprovações GMUD no Outlook Web, extração das demandas e geração automática de apresentações PowerPoint a partir de um modelo corporativo.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Status-Em%20Evolução-brightgreen">
  <img src="https://img.shields.io/badge/Linguagem-Node.js-green">
  <img src="https://img.shields.io/badge/Automação-Playwright-blue">
  <img src="https://img.shields.io/badge/Output-PowerPoint-orange">
  <img src="https://img.shields.io/badge/Python-python-yellow">
  <img src="https://img.shields.io/badge/Language-PT--BR-lightgrey">
</p>

---

<h2>📌 Sobre</h2>

<p>
O <b>Outlook GMUD PPTX</b> é uma ferramenta desenvolvida para automatizar a consolidação de aprovações de negócios GMUD recebidas através do <b>Outlook Web</b>.
</p>

<p>
O sistema realiza a busca de e-mails com o assunto <b>"De acordo de negócios GMUD"</b>, identifica as conversas relacionadas à aprovadora configurada e extrai automaticamente as informações de <b>"Registro de origem"</b> presentes na tabela <b>"UAT - Implantação GMUD"</b>.
</p>

<p>
Além das demandas, o sistema captura os registros visuais do <b>pedido de aprovação</b> e da respectiva <b>resposta</b>. Essas informações são utilizadas para gerar uma apresentação <b>PowerPoint (.pptx)</b> baseada em um modelo corporativo previamente definido.
</p>

<p align="center">
  <b>Outlook Web → Playwright → E-mails → Demandas + Aprovações → Python → PowerPoint</b>
</p>

---

<h2>🔄 Funcionamento</h2>

<pre>
┌────────────────┐     ┌──────────────┐     ┌────────────────┐     ┌──────────────────┐
│  Outlook Web   │ ──► │  index.js    │ ──► │ Busca / Filtro │ ──► │ E-mails GMUD     │
└────────────────┘     └──────────────┘     └────────────────┘     └──────────────────┘
                                                                      │
                                                                      ▼
                                                          ┌──────────────────────┐
                                                          │ Registro de origem   │
                                                          │ Pedido + Resposta    │
                                                          └──────────────────────┘
                                                                      │
                                                                      ▼
                                                          ┌──────────────────────┐
                                                          │   gerar_pptx.py      │
                                                          └──────────────────────┘
                                                                      │
                                                                      ▼
                                                          ┌──────────────────────┐
                                                          │   GMUD-AAAA-MM.pptx  │
                                                          └──────────────────────┘
</pre>

<p>
O fluxo principal é executado pelo <b>index.js</b>. O navegador é aberto através do Playwright, os e-mails são pesquisados no Outlook Web e cada conversa válida é processada individualmente.
</p>

<p>
Após a extração das informações, o Node.js prepara os dados e chama o <b>gerar_pptx.py</b>, que utiliza o modelo corporativo para montar a apresentação final.
</p>

---

<h2>🚀 Principais funcionalidades</h2>

<ul>
  <li>📧 Busca automática de e-mails de aprovação GMUD</li>
  <li>🔎 Pesquisa utilizando assunto, participantes e período</li>
  <li>📅 Processamento por mês, com possibilidade de definir o período manualmente</li>
  <li>👤 Validação da participação da aprovadora na conversa</li>
  <li>📋 Extração automática do campo <b>Registro de origem</b></li>
  <li>🧩 Suporte à tabela <b>UAT - Implantação GMUD</b></li>
  <li>🔄 Processamento de conversas longas e históricos encaminhados</li>
  <li>📜 Expansão automática de conteúdo colapsado no Outlook</li>
  <li>🛟 Fallback para extração através de texto puro quando a tabela HTML não é encontrada</li>
  <li>📸 Captura das telas de pedido e resposta da aprovação</li>
  <li>✂️ Recorte automático da região relevante dos prints</li>
  <li>📊 Geração automática de apresentação PowerPoint</li>
  <li>📑 Uma página de lista com todas as demandas do período</li>
  <li>📋 Uma página individual para cada demanda</li>
  <li>✅ Uma página de aprovações para cada e-mail processado</li>
  <li>🎨 Preservação da formatação original do modelo corporativo</li>
  <li>🔄 Duplicação automática dos slides do modelo</li>
  <li>🔤 Normalização automática dos títulos das demandas</li>
  <li>🧪 Arquivos de diagnóstico para casos em que a extração não seja concluída</li>
</ul>

---

<h2>📂 Estrutura do projeto</h2>

<pre>
outlook-gmud-pptx/
│
├── index.js
├── gerar_pptx.py
│
├── templates/
│   └── release_para_dev.pptx
│
├── output/
│   ├── GMUD-AAAA-MM.pptx
│   ├── pedido-0.png
│   ├── resposta-0.png
│   ├── config-pptx.json
│   └── debug-rascunho-*.png / *.txt
│
└── .owa-session/
    └── sessão persistente do Outlook
</pre>

---

<h2>🧩 Componentes</h2>

<h3>📧 index.js</h3>

<p>
É o componente principal da aplicação e responsável por controlar todo o fluxo de automação.
</p>

<ul>
  <li>🌐 Abre o Outlook Web através do Playwright</li>
  <li>🔐 Mantém uma sessão persistente do navegador</li>
  <li>🔎 Executa a busca dos e-mails</li>
  <li>📅 Define o período de processamento</li>
  <li>📧 Processa cada conversa encontrada</li>
  <li>📋 Extrai as demandas</li>
  <li>👤 Valida a participação da aprovadora</li>
  <li>📸 Captura os prints de aprovação</li>
  <li>🐍 Executa o script Python para geração do PowerPoint</li>
</ul>

<h3>🐍 gerar_pptx.py</h3>

<p>
Responsável pela montagem efetiva da apresentação PowerPoint.
</p>

<p>
O script utiliza o arquivo <b>release_para_dev.pptx</b> como modelo, evitando recriar o design da apresentação através de código. Os slides existentes são duplicados e apenas os textos e imagens necessários são substituídos, preservando a estrutura visual original do modelo.
</p>

<ul>
  <li>📊 Atualiza o mês/ano da capa</li>
  <li>📋 Preenche a lista de demandas</li>
  <li>📑 Duplica o slide de demanda conforme a quantidade encontrada</li>
  <li>📸 Duplica os slides de aprovação conforme os e-mails processados</li>
  <li>🖼️ Substitui as imagens mantendo proporção e posicionamento</li>
  <li>🔤 Normaliza títulos de User Stories</li>
  <li>🗑️ Remove slides-modelo que não sejam necessários</li>
</ul>

<p>
A estratégia de duplicação permite preservar as características visuais do modelo, incluindo formatação, imagens, fontes, fundos e elementos gráficos.
</p>

---

<h2>📊 Estrutura da apresentação</h2>

<p>
O PowerPoint é construído utilizando quatro modelos principais de slide:
</p>

<ul>
  <li><b>Slide 1 — Capa:</b> atualiza somente o mês e ano do relatório</li>
  <li><b>Slide 2 — Demandas:</b> apresenta a lista de todas as demandas encontradas no período</li>
  <li><b>Slide 3 — Demanda:</b> gera uma página individual para cada demanda</li>
  <li><b>Slide 4 — Aprovações:</b> gera uma página para cada e-mail, apresentando o pedido e a resposta</li>
</ul>

<p align="center">
  <b>Capa → Lista de Demandas → Detalhamento → Aprovações</b>
</p>

---

<h2>🔎 Busca dos e-mails</h2>

<p>
A busca é realizada diretamente no Outlook Web utilizando os recursos de pesquisa da própria plataforma.
</p>

<p>
O sistema combina o assunto do e-mail, o participante configurado e o intervalo de datas para reduzir os resultados antes do processamento individual de cada conversa.
</p>

<p>
O assunto utilizado como referência é:
</p>

<pre>
De acordo de negócios GMUD
</pre>

<p>
Além da query do Outlook, o assunto é validado novamente no conteúdo apresentado na lista de resultados antes que o e-mail seja processado.
</p>

---

<h2>📅 Período de processamento</h2>

<p>
Por padrão, o sistema trabalha com o <b>mês anterior ao mês de execução</b>.
</p>

<p>
Por exemplo, uma execução realizada em outubro processará os e-mails de setembro.
</p>

<p>
Também é possível definir manualmente o mês através da variável de ambiente <b>GMUD_MES</b>. O ano pode ser definido através de <b>GMUD_ANO</b>.
</p>

<pre>
GMUD_MES=8
GMUD_ANO=2026
</pre>

<p>
Essa configuração permite executar novamente relatórios de períodos anteriores sem alterar o código-fonte.
</p>

---

<h2>📋 Extração das demandas</h2>

<p>
Após abrir cada conversa válida, o sistema utiliza a funcionalidade <b>Encaminhar</b> do Outlook para obter o histórico da conversa no corpo do rascunho sem enviar nenhuma mensagem.
</p>

<p>
A partir desse conteúdo, procura a tabela <b>UAT - Implantação GMUD</b> e localiza o campo <b>Registro de origem</b>.
</p>

<p>
Quando a tabela HTML está disponível, o conteúdo da célula é convertido em uma lista de demandas. Caso a tabela tenha sido achatada em texto simples pelo Outlook, o sistema utiliza um mecanismo de fallback para localizar o mesmo campo diretamente no texto.
</p>

---

<h2>👤 Validação da aprovação</h2>

<p>
O sistema também verifica se a aprovadora configurada realmente participou da conversa.
</p>

<p>
Essa validação evita incluir no relatório conversas nas quais o usuário apenas estava copiado em um processo de aprovação pertencente a outra área ou responsável.
</p>

<p>
A regra pode ser controlada através da configuração:
</p>

<pre>
remetenteAprovador
</pre>

<p>
Por padrão, o sistema procura pelo nome ou endereço de e-mail configurado para a aprovadora.
</p>

---

<h2>📸 Captura das aprovações</h2>

<p>
Para cada conversa processada, o sistema captura dois registros visuais:
</p>

<ul>
  <li>📤 <b>Pedido:</b> e-mail original enviado para aprovação</li>
  <li>📥 <b>Resposta:</b> resposta contendo o acordo da aprovadora</li>
</ul>

<p>
O pedido é localizado posteriormente dentro da pasta <b>Itens Enviados</b>, utilizando o mesmo assunto da conversa processada.
</p>

<p>
Os prints são capturados em viewport fixo de <b>1280 × 720</b>. O tratamento visual final é realizado pelo <b>gerar_pptx.py</b>, que recorta a região correspondente ao painel de leitura e preserva a proporção da imagem.
</p>

---

<h2>🔐 Autenticação</h2>

<p>
A aplicação não utiliza integração direta através de Azure AD/Entra ID. O acesso ao Outlook Web é realizado através de <b>automação de navegador</b>.
</p>

<p>
Na primeira execução, o Chromium é aberto para que o usuário realize o login manualmente, incluindo MFA quando necessário.
</p>

<p>
A sessão é armazenada localmente em:
</p>

<pre>
.owa-session/
</pre>

<p>
Nas execuções seguintes, essa sessão persistente normalmente permite acessar o Outlook sem repetir o processo completo de autenticação.
</p>

<p>
A pasta <b>.owa-session</b> contém informações de sessão e não deve ser versionada no Git.
</p>

---

<h2>⚙️ Configuração</h2>

<p>
As principais configurações ficam concentradas no objeto <b>CONFIG</b> do <b>index.js</b>.
</p>

<ul>
  <li><b>outlookUrl</b> — endereço do Outlook Web</li>
  <li><b>destinatario</b> — destinatário utilizado na busca</li>
  <li><b>assuntoBusca</b> — assunto utilizado na pesquisa</li>
  <li><b>assuntoRegex</b> — validação adicional do assunto</li>
  <li><b>tituloTabelaRegex</b> — identificação da tabela GMUD</li>
  <li><b>labelRegistroOrigem</b> — identificação do campo de origem</li>
  <li><b>remetenteAprovador</b> — validação da participação da aprovadora</li>
  <li><b>outputDir</b> — diretório dos arquivos gerados</li>
  <li><b>templatePptx</b> — caminho do modelo PowerPoint</li>
  <li><b>gerarPptxScript</b> — caminho do script Python</li>
  <li><b>pythonBin</b> — interpretador Python utilizado para gerar o PPTX</li>
</ul>

---

<h2>🚀 Execução</h2>

<p>
Instale as dependências do projeto e execute o arquivo principal:
</p>

<pre>
node index.js
</pre>

<p>
O navegador será aberto automaticamente.
</p>

<p>
Na primeira execução:
</p>

<pre>
1. Chromium é aberto
2. Realize o login no Outlook Web
3. Conclua o MFA, se solicitado
4. Aguarde o carregamento da caixa de entrada
5. Pressione ENTER no terminal
6. O processamento será iniciado
</pre>

<p>
Ao final, o PowerPoint será salvo na pasta <b>output/</b>.
</p>

---

<h2>🐍 Dependências Python</h2>

<p>
O <b>gerar_pptx.py</b> utiliza Python para manipular o arquivo PowerPoint e as imagens capturadas.
</p>

<p>
As principais bibliotecas utilizadas são:
</p>

<ul>
  <li><b>python-pptx</b> — manipulação e geração do PowerPoint</li>
  <li><b>Pillow</b> — leitura, recorte e dimensionamento das imagens</li>
</ul>

<p>
O interpretador Python utilizado pelo Node pode ser configurado em:
</p>

<pre>
CONFIG.pythonBin
</pre>

<p>
No Windows, o projeto pode utilizar um caminho absoluto para garantir que o Python correto seja chamado, evitando conflitos quando existem múltiplas instalações de Python no sistema.
</p>

---

<h2>📁 Arquivos gerados</h2>

<p>
Durante a execução, os arquivos são armazenados na pasta <b>output/</b>.
</p>

<ul>
  <li><b>GMUD-AAAA-MM.pptx</b> — apresentação final</li>
  <li><b>pedido-N.png</b> — captura do pedido de aprovação</li>
  <li><b>resposta-N.png</b> — captura da resposta da aprovação</li>
  <li><b>config-pptx.json</b> — dados utilizados pelo gerador Python</li>
  <li><b>debug-rascunho-N.png</b> — captura para diagnóstico de falhas</li>
  <li><b>debug-rascunho-N.txt</b> — conteúdo textual utilizado no diagnóstico</li>
</ul>

---

<h2>🛟 Tratamento de conversas longas</h2>

<p>
Conversas com muitos encaminhamentos e respostas podem possuir partes do histórico ocultas ou carregadas de forma assíncrona pelo Outlook Web.
</p>

<p>
Para lidar com esse cenário, o sistema possui mecanismos específicos para:
</p>

<ul>
  <li>🔄 Aguardar o carregamento do corpo da mensagem</li>
  <li>⏳ Detectar quando o conteúdo terminou de estabilizar</li>
  <li>📜 Rolar o conteúdo para forçar o carregamento</li>
  <li>🔽 Expandir conteúdos anteriormente colapsados</li>
  <li>🔎 Procurar especificamente pelo campo <b>Registro de origem</b></li>
  <li>🛟 Utilizar extração por texto como fallback</li>
</ul>

<p>
Caso a informação ainda não seja localizada, o sistema salva arquivos de diagnóstico na pasta <b>output/</b> para facilitar a análise do conteúdo retornado pelo Outlook.
</p>

---

<h2>🎨 Modelo PowerPoint</h2>

<p>
O arquivo:
</p>

<pre>
templates/release_para_dev.pptx
</pre>

<p>
é utilizado como base visual da apresentação.
</p>

<p>
O projeto não recria o layout do PowerPoint através de código. Em vez disso, utiliza o próprio modelo corporativo como estrutura e duplica os slides necessários, alterando somente os dados variáveis.
</p>

<p>
Essa abordagem permite manter a identidade visual original mesmo quando a quantidade de demandas e aprovações varia entre os meses.
</p>

---

<h2>⚠️ Limitações conhecidas</h2>

<ul>
  <li>🌐 Os seletores utilizados pelo Playwright dependem da estrutura atual do Outlook Web</li>
  <li>🖥️ A captura das aprovações depende do viewport fixo de 1280 × 720</li>
  <li>✂️ Alterações no layout do Outlook podem exigir recalibração da região de recorte dos prints</li>
  <li>🔐 A autenticação depende de uma sessão válida do navegador</li>
  <li>🐍 O caminho do Python pode precisar ser ajustado em outras máquinas</li>
  <li>📄 O modelo <b>release_para_dev.pptx</b> precisa estar disponível no diretório configurado</li>
</ul>

---

<h2>🛠 Tecnologias</h2>

<ul>
  <li>Node.js</li>
  <li>JavaScript</li>
  <li>Playwright</li>
  <li>Python</li>
  <li>python-pptx</li>
  <li>Pillow</li>
  <li>Outlook Web</li>
  <li>PowerPoint / PPTX</li>
  <li>HTML / DOM</li>
  <li>Regex</li>
</ul>

---

<h2>🔮 Próximos passos</h2>

<ul>
  <li>⚙️ Tornar as configurações externas ao código</li>
  <li>📅 Facilitar a seleção de períodos diretamente na execução</li>
  <li>🧪 Ampliar os mecanismos de diagnóstico</li>
  <li>🌐 Tornar os seletores do Outlook mais resilientes a alterações de interface</li>
  <li>📊 Expandir as informações apresentadas no relatório</li>
  <li>⏰ Automatizar a execução periódica do processo</li>
</ul>

---

<h2>📊 Resultado</h2>

<p align="center">
  <b>Outlook Web</b>
  <br>
  ↓
  <br>
  <b>E-mails GMUD</b>
  <br>
  ↓
  <br>
  <b>Registro de origem</b>
  <br>
  ↓
  <br>
  <b>Demandas + Aprovações</b>
  <br>
  ↓
  <br>
  <b>Template Corporativo</b>
  <br>
  ↓
  <br>
  <b>PowerPoint</b>
</p>

---

<p align="center">
<b>Outlook GMUD PPTX</b> automatiza a consolidação das aprovações de negócios GMUD, transformando informações dispersas em e-mails do Outlook em uma apresentação PowerPoint estruturada e pronta para utilização.
</p>
