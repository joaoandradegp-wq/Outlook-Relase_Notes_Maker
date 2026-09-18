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
O sistema realiza a busca de e-mails com o assunto <b>"De acordo de negócios GMUD"</b>, enviados para o endereço configurado, e extrai automaticamente as informações de <b>"Registro de origem"</b> presentes na tabela <b>"UAT - Implantação GMUD"</b>.
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
  <li>🔎 Pesquisa por assunto e destinatário</li>
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
Além do assunto, a busca considera o destinatário configurado:
</p>

<pre>
jessica.fachina@unidas.com.br
</pre>

<p>
O período de busca considera o mês anterior ao mês em que o script é executado, abrangendo desde o primeiro até o último dia do mês anterior.
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
Por exemplo, uma execução realizada em outubro processará os e-mails correspondentes ao período de setembro, considerando do primeiro ao último dia do mês.
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
As configurações relacionadas ao comportamento da automação ficam definidas no código do projeto.
</p>

<p>
Entre os principais parâmetros estão:
</p>

<ul>
  <li><b>Outlook Web</b> - endereço utilizado para acesso à plataforma</li>
  <li><b>Destinatário</b> - endereço utilizado como filtro da pesquisa</li>
  <li><b>Assunto</b> - texto utilizado para localizar os e-mails GMUD</li>
  <li><b>Período</b> - intervalo correspondente ao mês anterior</li>
  <li><b>Tabela GMUD</b> - identificação da tabela <b>UAT - Implantação GMUD</b></li>
  <li><b>Registro de origem</b> - campo utilizado para extração das demandas</li>
  <li><b>Output</b> - diretório utilizado para armazenar os arquivos gerados</li>
  <li><b>Template PowerPoint</b> - modelo utilizado para geração da apresentação</li>
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

<h3>1. Instale as depen
