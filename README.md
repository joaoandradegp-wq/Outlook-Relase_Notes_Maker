<h1 align="center">📧 Outlook Release Notes Maker</h1>

<p align="center">
Automação para consulta de aprovações GMUD no Outlook Web, extração das demandas e geração automática de apresentações PowerPoint a partir de um modelo corporativo.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Status-Em%20Evolução-brightgreen">
  <img src="https://img.shields.io/badge/Linguagem-Node.js-green">
  <img src="https://img.shields.io/badge/Automação-Playwright-blue">
  <img src="https://img.shields.io/badge/Output-PowerPoint-orange">
  <img src="https://img.shields.io/badge/Language-PT--BR-lightgrey">
</p>

---

<h2>📌 Sobre</h2>

<p>
O <b>Release Notes Maker</b> é uma ferramenta desenvolvida para automatizar a consolidação de aprovações de negócios GMUD recebidas através do <b>Outlook Web</b>.
</p>

<p>
O sistema realiza a busca de e-mails com o assunto <b>"De acordo de negócios GMUD"</b>, enviados para o destinatário configurado, e extrai automaticamente as informações de <b>"Registro de origem"</b> presentes na tabela <b>"UAT - Implantação GMUD"</b>.
</p>

<p>
As informações extraídas são utilizadas para gerar uma apresentação <b>PowerPoint (.pptx)</b> estruturada de acordo com o modelo corporativo previamente definido.
</p>

<p>
Como não existe acesso para registrar uma aplicação no <b>Azure AD/Entra ID</b>, a autenticação é realizada através de <b>automação de navegador</b>. O login é feito manualmente na primeira execução e a sessão é armazenada localmente para as execuções seguintes.
</p>

<p align="center">
  <b>Outlook Web → Playwright → E-mails GMUD → Registro de origem → Python → PowerPoint</b>
</p>

---

<h2>🔄 Funcionamento</h2>

<pre>
┌────────────────┐     ┌──────────────┐     ┌────────────────┐     ┌──────────────────┐
│  Outlook Web   │ ──► │   index.js   │ ──► │ Busca / Filtro │ ──► │ E-mails GMUD     │
└────────────────┘     └──────────────┘     └────────────────┘     └──────────────────┘
                                                                      │
                                                                      ▼
                                                          ┌──────────────────────┐
                                                          │ Registro de origem   │
                                                          │ UAT - Implantação    │
                                                          │ GMUD                 │
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
O fluxo principal é executado pelo <b>index.js</b>. O navegador é aberto através do Playwright, os e-mails são pesquisados no Outlook Web e cada e-mail encontrado é processado individualmente.
</p>

<p>
Após a extração das informações, o Node.js prepara os dados e chama o <b>gerar_pptx.py</b>, responsável pela geração da apresentação PowerPoint.
</p>

---

<h2>🚀 Principais funcionalidades</h2>

<ul>
  <li>📧 Busca automática de e-mails de aprovação GMUD</li>
  <li>🔎 Pesquisa por assunto e destinatário configurado</li>
  <li>📅 Processamento automático do mês anterior ao mês de execução</li>
  <li>📋 Extração automática do campo <b>Registro de origem</b></li>
  <li>🧩 Suporte à tabela <b>UAT - Implantação GMUD</b></li>
  <li>🔄 Processamento do histórico das conversas</li>
  <li>📜 Tratamento de conteúdo carregado pelo Outlook Web</li>
  <li>🛟 Fallback para extração através de texto quando necessário</li>
  <li>📊 Geração automática de apresentação PowerPoint</li>
  <li>📑 Organização das demandas encontradas no período</li>
  <li>🎨 Utilização de modelo corporativo como base visual</li>
  <li>🔄 Duplicação dos slides conforme a quantidade de informações processadas</li>
  <li>🧪 Suporte a diagnóstico de problemas relacionados aos seletores do Outlook</li>
  <li>🔐 Sessão persistente do navegador para evitar novo login a cada execução</li>
</ul>

---

<h2>📂 Estrutura do projeto</h2>

<pre>
outlook-gmud-pptx/
│
├── index.js
├── gerar_pptx.py
├── package.json
│
├── templates/
│   └── release_para_dev.pptx
│
├── output/
│   ├── GMUD-AAAA-MM.pptx
│   ├── pedido-N.png
│   ├── resposta-N.png
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
É o componente principal da aplicação e responsável por controlar o fluxo de automação.
</p>

<ul>
  <li>🌐 Abre o Outlook Web através do Playwright</li>
  <li>🔐 Mantém uma sessão persistente do navegador</li>
  <li>🔎 Executa a busca dos e-mails</li>
  <li>📅 Define o período de processamento</li>
  <li>📧 Processa os e-mails encontrados</li>
  <li>📋 Extrai os registros de origem</li>
  <li>🐍 Executa o script Python para geração do PowerPoint</li>
</ul>

<h3>🐍 gerar_pptx.py</h3>

<p>
Responsável pela geração da apresentação PowerPoint a partir das informações extraídas pelo Node.js.
</p>

<p>
O script utiliza o arquivo <b>release_para_dev.pptx</b> como modelo, mantendo o layout corporativo existente. A apresentação é preenchida a partir da estrutura visual do modelo, evitando a necessidade de recriar o design através de código.
</p>

<ul>
  <li>📊 Atualiza as informações do período</li>
  <li>📋 Preenche as informações das demandas</li>
  <li>📑 Organiza as demandas nos slides</li>
  <li>🖼️ Manipula as imagens capturadas durante o processamento</li>
  <li>🔤 Ajusta os textos conforme os dados extraídos</li>
  <li>🗑️ Remove elementos do modelo que não sejam necessários</li>
</ul>

<p>
A utilização do modelo PowerPoint permite preservar a identidade visual original da apresentação, incluindo formatação, fontes, imagens, fundos e elementos gráficos.
</p>

---

<h2>📊 Estrutura da apresentação</h2>

<p>
O PowerPoint é construído utilizando o modelo corporativo definido em:
</p>

<pre>
templates/release_para_dev.pptx
</pre>

<p>
A estrutura da apresentação é baseada nos modelos existentes no arquivo e recebe os dados extraídos dos e-mails processados.
</p>

<p align="center">
  <b>Modelo Corporativo → Demandas GMUD → Informações Extraídas → PowerPoint</b>
</p>

---

<h2>🔎 Busca dos e-mails</h2>

<p>
A busca é realizada diretamente no Outlook Web utilizando os recursos de pesquisa da própria plataforma.
</p>

<p>
O sistema procura por e-mails cujo assunto contenha:
</p>

<pre>
De acordo de negócios GMUD
</pre>

<p>
O destinatário utilizado na pesquisa é definido através da configuração <b>destinatario</b> no arquivo <b>index.js</b>.
</p>

<p>
Dessa forma, o endereço utilizado pela automação não precisa ser exposto na documentação ou diretamente no código de processamento.
</p>

<p>
Após a busca inicial, o sistema processa os resultados individualmente para extrair as informações necessárias.
</p>

---

<h2>📅 Período de processamento</h2>

<p>
Por padrão, o sistema processa o <b>mês anterior ao mês de execução</b>.
</p>

<p>
Por exemplo, uma execução realizada em outubro processará os e-mails correspondentes ao período de setembro, considerando do primeiro até o último dia do mês.
</p>

<p>
O nome do arquivo gerado segue o padrão:
</p>

<pre>
GMUD-AAAA-MM.pptx
</pre>

<p>
Dessa forma, o relatório gerado fica identificado automaticamente pelo período processado.
</p>

---

<h2>📋 Extração das demandas</h2>

<p>
Após localizar os e-mails, o sistema procura pela tabela:
</p>

<pre>
UAT - Implantação GMUD
</pre>

<p>
Dentro da tabela, é localizada a linha correspondente ao campo:
</p>

<pre>
Registro de origem
</pre>

<p>
O conteúdo da coluna correspondente é extraído e convertido em uma lista de demandas. Os registros podem estar separados por bullets ou quebras de linha.
</p>

<p>
Quando a estrutura HTML da tabela não está disponível ou o Outlook apresenta o conteúdo de maneira diferente, o sistema utiliza mecanismos alternativos de extração para localizar as informações no conteúdo textual da mensagem.
</p>

---

<h2>🔐 Autenticação</h2>

<p>
A aplicação não utiliza integração direta através de Azure AD/Entra ID. O acesso ao Outlook Web é realizado através de <b>automação de navegador</b> utilizando Playwright.
</p>

<p>
Na primeira execução, o Chromium é aberto para que o usuário realize o login manualmente, incluindo usuário, senha e MFA quando necessário.
</p>

<p>
Após o login, o usuário deve aguardar o carregamento da caixa de entrada e pressionar <b>ENTER</b> no terminal para iniciar o processamento.
</p>

<p>
A sessão é armazenada localmente em:
</p>

<pre>
.owa-session/
</pre>

<p>
Nas execuções seguintes, a sessão salva normalmente permite acessar o Outlook sem repetir o processo completo de autenticação.
</p>

<p>
Caso o token expire por política do tenant, será necessário realizar o login novamente.
</p>

<p>
A pasta <b>.owa-session</b> contém informações de sessão e não deve ser versionada no Git.
</p>

---

<h2>⚙️ Configuração</h2>

<p>
As principais configurações relacionadas ao comportamento da automação ficam concentradas no objeto <b>CONFIG</b> do arquivo <b>index.js</b>.
</p>

<ul>
  <li><b>outlookUrl</b> - endereço utilizado para acesso ao Outlook Web</li>
  <li><b>destinatario</b> - endereço utilizado como filtro da pesquisa</li>
  <li><b>assuntoBusca</b> - assunto utilizado na pesquisa dos e-mails</li>
  <li><b>assuntoRegex</b> - validação adicional do assunto</li>
  <li><b>tituloTabelaRegex</b> - identificação da tabela GMUD</li>
  <li><b>labelRegistroOrigem</b> - identificação do campo de origem</li>
  <li><b>outputDir</b> - diretório dos arquivos gerados</li>
  <li><b>templatePptx</b> - caminho do modelo PowerPoint</li>
  <li><b>gerarPptxScript</b> - caminho do script Python</li>
  <li><b>pythonBin</b> - interpretador Python utilizado para gerar o PPTX</li>
</ul>

---

<h2>🚀 Execução</h2>

<p>
Instale as dependências do projeto:
</p>

<pre>
npm install
npx playwright install chromium
</pre>

<p>
Depois execute:
</p>

<pre>
npm start
</pre>

<p>
O navegador será aberto automaticamente.
</p>

---

<h2>📖 Como utilizar</h2>

<p>
O <b>Release Notes Maker</b> foi desenvolvido para gerar automaticamente o relatório mensal de aprovações de negócios GMUD a partir dos e-mails disponíveis no Outlook Web.
</p>

<h3>1. Instale as dependências</h3>

<p>
Na primeira utilização, abra o terminal na pasta do projeto e execute:
</p>

<pre>
npm install
npx playwright install chromium
</pre>

<h3>2. Inicie a aplicação</h3>

<p>
Execute:
</p>

<pre>
npm start
</pre>

<h3>3. Faça o login no Outlook</h3>

<p>
Na primeira execução, uma janela do Chromium será aberta automaticamente com o Outlook Web.
</p>

<p>
Realize o login utilizando suas credenciais corporativas e conclua o MFA, caso seja solicitado.
</p>

<h3>4. Inicie o processamento</h3>

<p>
Após o carregamento da caixa de entrada, volte ao terminal e pressione <b>ENTER</b>.
</p>

<p>
A aplicação iniciará automaticamente a busca pelos e-mails que atendem aos critérios configurados:
</p>

<ul>
  <li>Assunto configurado em <b>assuntoBusca</b></li>
  <li>Destinatário configurado em <b>destinatario</b></li>
  <li>Período correspondente ao mês anterior à execução</li>
</ul>

<h3>5. Aguarde o processamento</h3>

<p>
O sistema abrirá e processará os e-mails encontrados, identificará a tabela <b>UAT - Implantação GMUD</b> e extrairá os registros do campo <b>Registro de origem</b>.
</p>

<p>
As informações serão então encaminhadas para o processo de geração do PowerPoint.
</p>

<h3>6. Consulte o PowerPoint gerado</h3>

<p>
Ao final da execução, o relatório estará disponível na pasta:
</p>

<pre>
output/
</pre>

<p>
O arquivo seguirá o padrão:
</p>

<pre>
GMUD-AAAA-MM.pptx
</pre>

<p>
A apresentação utilizará o modelo corporativo definido em <b>templates/release_para_dev.pptx</b>.
</p>

<h3>7. Execuções posteriores</h3>

<p>
Após o primeiro login, a sessão do Outlook é armazenada na pasta:
</p>

<pre>
.owa-session/
</pre>

<p>
Nas próximas execuções, normalmente não será necessário realizar o login novamente. Basta executar:
</p>

<pre>
npm start
</pre>

<p>
Caso a sessão tenha expirado por política do tenant, o Outlook solicitará um novo login. Realize a autenticação novamente e continue o processo normalmente.
</p>

<h3>📌 Resumo rápido</h3>

<pre>
npm install
npx playwright install chromium
npm start
        ↓
Login no Outlook
        ↓
Pressione ENTER
        ↓
Busca dos e-mails GMUD
        ↓
Extração do Registro de origem
        ↓
Geração do PowerPoint
        ↓
output/GMUD-AAAA-MM.pptx
</pre>

---

<h2>🐍 Dependências Python</h2>

<p>
O <b>gerar_pptx.py</b> utiliza Python para manipular o arquivo PowerPoint e as imagens utilizadas na apresentação.
</p>

<p>
As principais bibliotecas utilizadas são:
</p>

<ul>
  <li><b>python-pptx</b> - manipulação e geração do PowerPoint</li>
  <li><b>Pillow</b> - leitura, recorte e dimensionamento das imagens</li>
</ul>

<p>
O ambiente Python utilizado pelo projeto precisa possuir as bibliotecas necessárias para executar o script de geração do PowerPoint.
</p>

---

<h2>📁 Arquivos gerados</h2>

<p>
Durante a execução, os arquivos são armazenados na pasta <b>output/</b>.
</p>

<ul>
  <li><b>GMUD-AAAA-MM.pptx</b> - apresentação final</li>
  <li><b>pedido-N.png</b> - captura do pedido de aprovação</li>
  <li><b>resposta-N.png</b> - captura da resposta da aprovação</li>
  <li><b>config-pptx.json</b> - dados utilizados pelo gerador Python</li>
  <li><b>debug-rascunho-N.png</b> - captura utilizada para diagnóstico</li>
  <li><b>debug-rascunho-N.txt</b> - conteúdo textual utilizado para diagnóstico</li>
</ul>

---

<h2>🛟 Tratamento de conteúdo do Outlook</h2>

<p>
O Outlook Web pode apresentar diferenças na estrutura do conteúdo dependendo do idioma, versão da interface ou configuração do tenant.
</p>

<p>
O sistema possui mecanismos para localizar os elementos necessários mesmo quando pequenas alterações ocorrem na interface.
</p>

<p>
Os pontos mais sensíveis estão relacionados aos seletores utilizados pelo Playwright:
</p>

<ul>
  <li>🔎 <b>Caixa de busca</b> - seletor baseado em <b>aria-label</b>, como "Search" ou "Pesquisar"</li>
  <li>📧 <b>Lista de e-mails</b> - utilização de elementos com <b>[role="option"]</b></li>
  <li>📄 <b>Corpo do e-mail</b> - utilização de iframe, como <b>iframe[title="Message Body"]</b></li>
</ul>

<p>
Caso o script não encontre algum elemento, a execução pode ser acompanhada visualmente porque o navegador é iniciado com <b>headless: false</b>.
</p>

<p>
Nesses casos, o DevTools do navegador pode ser utilizado para verificar a estrutura atual dos elementos e ajustar os seletores marcados no código com:
</p>

<pre>
AJUSTAR SE PRECISAR
</pre>

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
O projeto utiliza o próprio modelo corporativo como estrutura da apresentação, evitando recriar o layout através de código.
</p>

<p>
Os dados extraídos dos e-mails são inseridos no modelo de acordo com a estrutura definida no PowerPoint.
</p>

<p>
Essa abordagem permite manter a identidade visual original da apresentação, incluindo fontes, imagens, fundos, posicionamento e elementos gráficos.
</p>

---

<h2>⚠️ Limitações conhecidas</h2>

<ul>
  <li>🌐 Os seletores utilizados pelo Playwright dependem da estrutura atual do Outlook Web</li>
  <li>🖥️ Alterações na interface do Outlook podem exigir ajustes nos seletores</li>
  <li>🔐 A autenticação depende de uma sessão válida do navegador</li>
  <li>⏳ O token da sessão pode expirar de acordo com as políticas do tenant</li>
  <li>📄 O modelo <b>release_para_dev.pptx</b> precisa estar disponível no diretório configurado</li>
  <li>🐍 O ambiente Python precisa possuir as bibliotecas necessárias para geração do PPTX</li>
  <li>📧 Alterações na estrutura dos e-mails ou da tabela <b>UAT - Implantação GMUD</b> podem afetar a extração</li>
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
  <b>Demandas GMUD</b>
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
<b>Release Notes Maker</b> automatiza a consolidação das aprovações de negócios GMUD, transformando informações presentes nos e-mails do Outlook em uma apresentação PowerPoint estruturada e pronta para utilização.
</p>
