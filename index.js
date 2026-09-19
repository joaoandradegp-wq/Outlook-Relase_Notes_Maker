/**
 * outlook-gmud-pptx
 * -----------------
 * Le e-mails do Outlook Web da Unidas com o assunto "De acordo de negocios GMUD",
 * enviados para jessica.fachina@unidas.com.br, dentro do mes vigente.
 * Extrai a lista de "Registro de origem" da tabela "UAT - Implantacao GMUD"
 * e gera um PPTX com uma demanda (ou lista de demandas) por slide.
 *
 * IMPORTANTE (leia antes de rodar):
 * - Na primeira execucao, uma janela do Chromium vai abrir. Faca login manualmente
 *   (incluindo o MFA). A sessao fica salva em .owa-session/, entao nas proximas
 *   execucoes normalmente nao vai pedir login de novo.
 * - Os seletores de tela (aria-label, role, etc) do Outlook Web podem variar
 *   um pouco dependendo da versao/tenant. Se algo nao funcionar, rode com
 *   headless: false (ja e o padrao aqui) e ajuste os seletores marcados com
 *   "AJUSTAR SE PRECISAR" observando o DevTools do navegador.
 */

const { chromium } = require('playwright');
const { spawnSync } = require('child_process');
const path = require('path');
const fs = require('fs');
const readline = require('readline');

// ------------------------- CONFIGURAÇÃO -------------------------

const CONFIG = {
  userDataDir: path.join(__dirname, '.owa-session'),
  outlookUrl: 'https://outlook.office.com/mail/',
  destinatario: 'jessica.fachina@unidas.com.br',
  assuntoBusca: 'de acordo de negocios gmud', // usado na query de busca (sem acento p/ compatibilidade)
  assuntoRegex: /de\s*acordo\s*de\s*neg[oó]cios\s*gmud/i, // validação extra do assunto real
  tituloTabelaRegex: /UAT\s*-\s*Implanta[cç][aã]o\s*GMUD/i,
  labelRegistroOrigem: /registro\s*de\s*origem/i,
  // Usado só no fallback por texto puro (ver buscarRegistroOrigemViaTexto): reconhece uma
  // linha de demanda tipo "711716 - OPORTUNIDADE / LEAD | ..." ou "User Story 716019: ...".
  regexLinhaDemanda: /^(user\s*story\s*)?\d{4,}/i,
  outputDir: path.join(__dirname, 'output'),
  // TEMPOS DE ESPERA (ms) do fluxo "Encaminhar". Se algum e-mail ainda vier sem as demandas,
  // aumente estes valores (principalmente registroOrigem e corpoEstabilizar).
  tempos: {
    botaoEncaminhar: 10000, // espera o botão "Encaminhar" aparecer depois de abrir o e-mail
    rascunhoAbrir: 8000, // espera o rascunho abrir (botão "Enviar") depois de cada tentativa de clique
    tabelaAparecer: 10000, // espera qualquer tabela aparecer no corpo do rascunho
    corpoEstavelPor: 1500, // corpo precisa ficar sem mudar por esse tempo pra considerar "carregado"
    corpoEstabilizar: 15000, // limite máximo esperando o corpo parar de mudar
    registroOrigem: 15000, // espera a linha "Registro de origem" aparecer (1ª rodada; as seguintes usam metade)
    rodadasExpansao: 3, // quantas vezes repete "rolar + expandir '...' + esperar" antes de desistir
  },
  // Modelo de PPTX corporativo (release_para_dev.pptx) e script Python que sabe montar
  // o deck final a partir dele — ver gerar_pptx.py, que faz a montagem de verdade
  // duplicando os slides do modelo (preserva 100% da formatação original).
  templatePptx: path.join(__dirname, 'templates', 'release_para_dev.pptx'),
  gerarPptxScript: path.join(__dirname, 'gerar_pptx.py'),
  // IMPORTANTE: usamos o caminho COMPLETO em vez de só "python" porque nessa máquina há
  // mais de um Python no PATH (ex: o do MSYS2/mingw64) e "python" pode resolver pro
  // interpretador errado — um que não tem python-pptx instalado (nem pip). O pip do
  // terminal aponta pro Python 3.13 abaixo, então é nele que instalamos python-pptx e é
  // ele que precisa ser chamado aqui. Se você reinstalar o Python ou usar outra máquina,
  // ajuste este caminho (rode `pip show python-pptx` no terminal certo pra confirmar onde
  // ele está instalado, ou `py -0p` pra listar todos os Pythons registrados no Windows).
  pythonBin:
    process.platform === 'win32'
      ? 'C:\\Users\\joaoa\\AppData\\Local\\Programs\\Python\\Python313\\python.exe'
      : 'python3',
};

const MESES_PT = [
  'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro',
];

function formatarMesAno(data) {
  // Formato usado no modelo do PPTX: "Agosto/26" (mês por extenso + ano com 2 dígitos)
  const mes = MESES_PT[data.getMonth()];
  const anoCurto = String(data.getFullYear()).slice(-2);
  return `${mes}/${anoCurto}`;
}

// ------------------------- UTILITÁRIOS -------------------------

function getIntervaloMesAnterior() {
  const agora = new Date();

  // Mês escolhido manualmente (ex: pelo iniciar.bat, que define GMUD_MES=1..12).
  // O ano é o atual, exceto se o mês escolhido ainda não chegou este ano (aí usa o ano
  // passado). Dá pra forçar o ano com GMUD_ANO.
  const mesEscolhido = parseInt(process.env.GMUD_MES, 10);
  if (mesEscolhido >= 1 && mesEscolhido <= 12) {
    let ano = parseInt(process.env.GMUD_ANO, 10);
    if (!ano) ano = mesEscolhido > agora.getMonth() + 1 ? agora.getFullYear() - 1 : agora.getFullYear();
    return { inicio: new Date(ano, mesEscolhido - 1, 1), fim: new Date(ano, mesEscolhido, 0) };
  }

  // Padrão (sem GMUD_MES): mês anterior ao mês vigente na data de execução do script.
  // Ex: rodando em outubro/2026, filtra e-mails de 01/09/2026 a 30/09/2026.
  const inicio = new Date(agora.getFullYear(), agora.getMonth() - 1, 1);
  const fim = new Date(agora.getFullYear(), agora.getMonth(), 0); // último dia do mês anterior
  return { inicio, fim };
}

function formatarDataOWA(data) {
  // IMPORTANTE: o Outlook Web espera o formato dd/mm/yyyy (dia primeiro), não mm/dd/yyyy.
  // Bug identificado pelo usuário: a versão anterior gerava mm/dd/yyyy, o que fazia a
  // busca por data (received:/sent:) se comportar de forma incoerente/errada.
  const dd = String(data.getDate()).padStart(2, '0');
  const mm = String(data.getMonth() + 1).padStart(2, '0');
  const yyyy = data.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function aguardarEnter(mensagem) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(mensagem, () => {
      rl.close();
      resolve();
    });
  });
}

// Converte o HTML da célula (com <li>, <br>, etc) em uma lista de linhas de texto limpo
function extrairLinhasDeHtml(html) {
  const texto = html
    .replace(/<li[^>]*>/gi, '\n• ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&');

  return texto
    .split('\n')
    .map((l) => l.replace(/^•\s*/, '').trim())
    .filter(Boolean);
}

// ------------------------- BUSCA NO OUTLOOK WEB -------------------------

async function abrirOutlookELogar(context) {
  const page = context.pages()[0] || (await context.newPage());
  await page.goto(CONFIG.outlookUrl, { waitUntil: 'domcontentloaded' });

  await aguardarEnter(
    '\n>> Se for necessário, faça login manualmente (usuário/senha/MFA) na janela aberta.\n' +
      '>> Quando a caixa de entrada estiver carregada, volte aqui e pressione ENTER para continuar...\n'
  );

  return page;
}

// AJUSTAR SE PRECISAR: seletores ampliados (cobre "Pesquisar"/"Search"/role=searchbox).
// Reaproveitada tanto pela busca inicial (buscarEmails) quanto pela busca do "pedido de
// aprovação" dentro de "Itens Enviados" (capturarPrintPedidoAprovacao).
function localizarCaixaBusca(page) {
  return page.locator(
    [
      'input[aria-label*="pesquis" i]',
      'input[aria-label*="search" i]',
      'input[placeholder*="pesquis" i]',
      'input[placeholder*="search" i]',
      'input[role="searchbox"]',
      '[role="searchbox"]',
    ].join(', ')
  );
}

async function buscarEmails(page) {
  const intervaloMes = getIntervaloMesAnterior();
  const dataInicio = formatarDataOWA(intervaloMes.inicio);
  const dataFim = formatarDataOWA(intervaloMes.fim);

  // Query: só conversas em que a JESSICA ENVIOU uma mensagem (o "DE ACORDO" dela):
  // subject:"..." from:jessica (received:X..Y OR sent:X..Y)
  // Antes era "participants:", que também trazia "de acordo" de OUTRAS BOs em que o usuário
  // só era copiado (participants casa com Para/De/Cc/Cco de qualquer mensagem). Com "from:"
  // esses ficam de fora já na busca. Datas no formato dd/mm/yyyy (o Outlook Web espera dia
  // primeiro — ver formatarDataOWA).
  const query = `subject:"${CONFIG.assuntoBusca}" from:${CONFIG.destinatario} (received:${dataInicio}..${dataFim} OR sent:${dataInicio}..${dataFim})`;

  const caixaBusca = localizarCaixaBusca(page);

  await caixaBusca.first().waitFor({ state: 'visible', timeout: 30000 });
  await caixaBusca.first().click();

  // Limpa o campo (seleciona tudo + apaga) antes de digitar, caso já tenha algo
  await page.keyboard.press('Control+A');
  await page.keyboard.press('Backspace');

  // IMPORTANTE: usamos digitação simulada (não fill()) porque o Outlook Web precisa
  // dos eventos de teclado reais para "ouvir" a digitação e habilitar a busca.
  await caixaBusca.first().pressSequentially(query, { delay: 30 });
  await page.waitForTimeout(500);
  await page.keyboard.press('Enter');

  // Espera os resultados carregarem
  await page.waitForTimeout(4000);

  // AJUSTAR SE PRECISAR: seletor da lista de e-mails na área de resultados
  const itens = page.locator('[role="option"]');
  let total = await itens.count();

  // Fallback: se o Enter não disparou a busca (0 resultados), tenta clicar no
  // ícone de lupa/buscar, caso exista, e espera de novo.
  if (total === 0) {
    const botaoBuscar = page.locator(
      'button[aria-label*="pesquis" i], button[aria-label*="search" i], button[title*="pesquis" i], button[title*="search" i]'
    );
    if (await botaoBuscar.count()) {
      await botaoBuscar.first().click();
      await page.waitForTimeout(4000);
      total = await itens.count();
    }
  }

  console.log(`\nEncontrados ${total} e-mail(s) com os critérios de busca.`);

  return { total, intervaloMes };
}

// Extrai a linha de assunto do texto da lista (linha que bate com o padrão do assunto).
// Evita depender de abrir o e-mail e procurar um role="heading" — a página tem vários
// (inclusive o do painel de navegação lateral), e não dá pra garantir qual é o certo.
function extrairAssuntoDoTextoDaLista(textoLinha) {
  const linhas = textoLinha.split('\n').map((l) => l.trim()).filter(Boolean);
  return linhas.find((l) => CONFIG.assuntoRegex.test(l)) || null;
}

// Procura, em todos os "contextos" passados (iframes e/ou a própria página), uma linha de
// tabela cujo rótulo bata com CONFIG.labelRegistroOrigem, e retorna o HTML da célula
// seguinte (o conteúdo real). Retorna também os rótulos encontrados (pra diagnóstico) e o
// total de linhas vistas.
// IMPORTANTE: recebe "contextos" (não só "frames") porque o corpo do e-mail LIDO fica
// dentro de um iframe, mas o corpo do RASCUNHO DE ENCAMINHAMENTO é renderizado direto na
// página principal (sem iframe) — por isso não dá pra sempre assumir que a tabela está
// num iframe; é preciso poder incluir a página principal na busca também.
async function buscarTabelaEmContextos(contextos) {
  let htmlConteudo = '';
  const labelsEncontrados = [];
  let totalLinhasVistas = 0;

  for (const ctx of contextos) {
    let linhas = [];
    try {
      linhas = await ctx.locator('table tr').all();
    } catch {
      continue; // iframe pode não estar mais anexado (navegação/re-render); ignora e segue
    }
    totalLinhasVistas += linhas.length;

    for (const linha of linhas) {
      const celulas = linha.locator('td, th');
      const qtdCelulas = await celulas.count().catch(() => 0);
      if (qtdCelulas >= 2) {
        const label = (await celulas.nth(0).innerText().catch(() => '')).trim();
        if (!label) continue;
        labelsEncontrados.push(label);
        if (CONFIG.labelRegistroOrigem.test(label)) {
          htmlConteudo = await celulas.nth(1).innerHTML();
          break;
        }
      }
    }
    if (htmlConteudo) break;
  }

  return { htmlConteudo, labelsEncontrados, totalLinhasVistas };
}

// FALLBACK: quando a tabela "UAT - Implantação GMUD" não é achada como <table> de verdade
// (buscarTabelaEmContextos retorna vazio), tenta achar o rótulo "Registro de origem" direto
// no TEXTO puro (innerText) do body de cada contexto e ler as linhas seguintes como demandas.
// Por que isso acontece: em e-mails encaminhados/respondidos várias vezes (threads longas,
// como o de 11/08 — resposta que cita um encaminhamento que cita o pedido original), o
// Outlook Web às vezes "achata" a tabela HTML original em texto simples nas camadas mais
// antigas da citação, mesmo com o conteúdo certinho visualmente. O texto continua com a
// mesma estrutura (rótulo em uma linha, valor nas linhas seguintes até a linha em branco),
// só não é mais um <table> real — por isso buscarTabelaEmContextos (que só olha table/tr)
// não encontra nada, mas dá pra recuperar a informação lendo o texto puro.
async function buscarRegistroOrigemViaTexto(contextos) {
  for (const ctx of contextos) {
    let texto = '';
    try {
      texto = await ctx.locator('body').innerText();
    } catch {
      continue; // contexto pode não estar mais anexado
    }

    const linhas = texto.split('\n').map((l) => l.trim());
    // Só considera a linha do rótulo se ela for CURTA (só o rótulo, tipo célula de tabela
    // achatada) — evita casar com um trecho de demanda que por acaso contenha as palavras
    // "registro de origem" no meio de uma frase longa.
    const idxRotulo = linhas.findIndex(
      (l) => CONFIG.labelRegistroOrigem.test(l) && l.length <= 40
    );
    if (idxRotulo === -1) continue;

    const demandas = [];
    for (let i = idxRotulo + 1; i < linhas.length; i++) {
      const linha = linhas[i];
      if (!linha) break; // linha em branco = fim da lista
      if (!CONFIG.regexLinhaDemanda.test(linha)) break; // não parece mais uma demanda
      demandas.push(linha);
    }

    if (demandas.length > 0) return demandas;
  }
  return [];
}

// As mesmas estratégias de nome (pt-BR/en-US) usadas tanto pra achar o botão "Encaminhar"
// quanto, depois, pra achar "Enviar"/"Fechar"/"Descartar" dentro do rascunho.
// IMPORTANTE: os regex de fechar/descartar são ANCORADOS (^...$) de propósito — um regex
// solto como /fechar|close/i também casa com o botão fixo "Fechar pesquisa" da barra de
// busca (sempre presente na tela) e, ao ser clicado, fecha a busca inteira em vez do
// rascunho, bagunçando a lista de resultados pros itens seguintes.
// DESCOBERTA (via DevTools, outerHTML real do botão "Encaminhar"): apesar de ser uma tag
// <button>, o elemento tem role="menuitem" explícito (não "button"!) — por isso toda busca
// por getByRole('button', ...) sempre dava 0 resultados. O barramento de ações do Outlook
// Web (Responder/Responder a todos/Encaminhar) é implementado como um menu (role="menu"),
// com cada ação sendo um role="menuitem". Por isso incluímos os dois roles nas estratégias.
function construirEstrategiasBotao(nomesPt, nomesEn) {
  const alternativasPt = nomesPt.join('|');
  const alternativasEn = nomesEn.join('|');
  const regexPt = new RegExp(`^\\s*(${alternativasPt})\\s*$`, 'i');
  const regexEn = new RegExp(`^\\s*(${alternativasEn})\\s*$`, 'i');
  return [
    { desc: `role=menuitem name=/${alternativasPt}/i (exato)`, build: (ctx) => ctx.getByRole('menuitem', { name: regexPt }) },
    { desc: `role=button name=/${alternativasPt}/i (exato)`, build: (ctx) => ctx.getByRole('button', { name: regexPt }) },
    { desc: `role=menuitem name=/${alternativasEn}/i (exato)`, build: (ctx) => ctx.getByRole('menuitem', { name: regexEn }) },
    { desc: `role=button name=/${alternativasEn}/i (exato)`, build: (ctx) => ctx.getByRole('button', { name: regexEn }) },
    { desc: `texto exato "${nomesPt[0]}"`, build: (ctx) => ctx.getByText(regexPt, { exact: true }) },
  ];
}

// Procura um elemento clicável em vários "contextos" (a página principal + cada iframe),
// tentando várias estratégias de nome em cascata. Retorna o primeiro achado, já indicando
// em qual contexto (útil pra log/diagnóstico). Frames vêm ANTES da página principal na
// ordem de busca: o botão "Encaminhar" de verdade normalmente está dentro do iframe do
// painel de leitura, então priorizamos os iframes pra evitar casar com elementos soltos
// da página principal (foi isso que aconteceu com a busca por aria-label antes).
async function localizarEmContextos(contexts, estrategias) {
  for (const ctx of contexts) {
    for (const { desc, build } of estrategias) {
      const alvo = build(ctx).first();
      const qtd = await alvo.count().catch(() => 0);
      if (qtd > 0) {
        return { alvo, desc };
      }
    }
  }
  return null;
}

const ESTRATEGIAS_ENCAMINHAR = construirEstrategiasBotao(['encaminhar'], ['forward']);
const ESTRATEGIAS_ENVIAR = construirEstrategiasBotao(['enviar'], ['send']);
const ESTRATEGIAS_DESCARTAR = construirEstrategiasBotao(['descartar', 'excluir rascunho'], ['discard', 'delete draft']);
const ESTRATEGIAS_FECHAR = construirEstrategiasBotao(['fechar'], ['close']);
// Botão de confirmação da caixa de diálogo "Descartar mensagem" ("Tem certeza que deseja
// descartar esse rascunho?") que o Outlook Web abre depois de clicar em "Descartar".
const ESTRATEGIAS_CONFIRMAR = construirEstrategiasBotao(['ok', 'sim'], ['ok', 'yes']);

// Depois de clicar em "Descartar", o Outlook Web abre uma caixa de diálogo de confirmação
// ("Descartar mensagem" / "Tem certeza que deseja descartar esse rascunho?") com os botões
// OK/Cancelar. Sem confirmar isso, o rascunho nunca é descartado de fato e o fluxo trava
// esperando uma interação que nunca vem. Essa função procura o botão OK por um tempo curto
// e clica se aparecer; se não aparecer (o Descartar já resolveu tudo sem diálogo), segue
// em frente normalmente.
async function confirmarDescarteSeAparecer(page) {
  for (let tentativa = 0; tentativa < 5; tentativa++) {
    const botaoOk = await localizarEmContextos(todosOsContextos(page), ESTRATEGIAS_CONFIRMAR);
    if (botaoOk) {
      await botaoOk.alvo.click({ timeout: 3000 }).catch(() => {});
      await page.waitForTimeout(500);
      return true;
    }
    await page.waitForTimeout(300);
  }
  return false;
}

function todosOsContextos(page) {
  const frames = page.frames().filter((f) => f !== page.mainFrame());
  return [...frames, page];
}

// Clique "nativo": chama elemento.click() direto via DOM (page.evaluate), sem passar pela
// simulação de mouse do Playwright. Usado como último recurso quando o clique simulado
// (normal ou forçado) não tem efeito nenhum — o que costuma acontecer quando o handler da
// biblioteca de UI depende só do evento 'click' do DOM e não de toda a sequência de eventos
// de ponteiro que o Playwright dispara (mousemove/mousedown/mouseup/click).
async function cliqueNativo(alvo) {
  const elementHandle = await alvo.elementHandle().catch(() => null);
  if (!elementHandle) return false;
  try {
    await elementHandle.evaluate((el) => el.click());
    return true;
  } catch {
    return false;
  } finally {
    await elementHandle.dispose().catch(() => {});
  }
}

// Botão "..." que o Outlook Web usa pra esconder o conteúdo citado/anterior de uma
// mensagem (padrão de "Show trimmed content"). O "Encaminhar" só garante que a mensagem
// MAIS RECENTE da conversa vem expandida por padrão — em threads mais longas, mensagens
// anteriores (que podem ser justamente onde está a tabela "Registro de origem") continuam
// colapsadas atrás desse botão MESMO dentro do rascunho de encaminhamento. É por isso que
// e-mails com conversa curta funcionam de primeira e os com conversa longa não.
const ESTRATEGIAS_MOSTRAR_ANTERIOR = [
  { desc: 'texto "..." (três pontos, conteúdo citado colapsado)', build: (ctx) => ctx.getByText(/^\s*\.{3}\s*$/) },
  {
    desc: 'aria-label "mostrar/show" + "anterior/trimmed/previous"',
    build: (ctx) =>
      ctx.locator(
        '[aria-label*="mostrar" i][aria-label*="anterior" i], [aria-label*="show" i][aria-label*="trimmed" i], [aria-label*="show" i][aria-label*="previous" i], [aria-label*="mostrar" i][aria-label*="conteúdo" i]'
      ),
  },
];

// Clica repetidamente em qualquer botão "..." de conteúdo colapsado que aparecer, em
// qualquer contexto, até não achar mais nenhum (ou até o limite de cliques). Precisa
// re-localizar o alvo a cada iteração (em vez de pegar todos de uma vez com .all()) porque
// cada clique costuma FAZER O BOTÃO DESAPARECER e re-renderizar o DOM ao redor, invalidando
// qualquer referência antiga.
async function expandirTodoConteudoColapsado(page, maxCliques = 30) {
  let cliques = 0;
  for (let tentativa = 0; tentativa < maxCliques; tentativa++) {
    let achou = null;
    for (const ctx of todosOsContextos(page)) {
      for (const { build } of ESTRATEGIAS_MOSTRAR_ANTERIOR) {
        const alvo = build(ctx).first();
        if (await alvo.count().catch(() => 0)) {
          achou = alvo;
          break;
        }
      }
      if (achou) break;
    }
    if (!achou) break;

    let clicou = await achou
      .click({ timeout: 1500 })
      .then(() => true)
      .catch(() => false);
    if (!clicou) clicou = await cliqueNativo(achou);
    if (!clicou) break;

    cliques += 1;
    // Em vez de esperar um tempo fixo curto (400ms), espera o corpo parar de mudar: em
    // conversas longas o conteúdo expandido demora pra ser renderizado, e o próximo "..."
    // (mensagem mais antiga) só aparece depois disso.
    await esperarCorpoEstabilizar(page, { estavelPorMs: 800, timeoutMs: 5000 });
  }
  return cliques > 0;
}

// Espera, com polling, alguma linha de tabela (table tr) aparecer em qualquer contexto
// (iframes ou página principal). Necessário porque o corpo do rascunho de encaminhamento
// (com o histórico da conversa e a tabela "UAT - Implantação GMUD") é injetado de forma
// assíncrona DEPOIS que o botão "Enviar" já está visível — então checar só uma vez logo
// após abrir o rascunho é uma race condition clássica, principalmente em conversas longas
// ou em máquinas/redes mais lentas.
async function esperarTabelaCarregar(page, timeoutMs = 10000, intervaloMs = 400) {
  const inicio = Date.now();
  while (Date.now() - inicio < timeoutMs) {
    for (const ctx of todosOsContextos(page)) {
      const qtd = await ctx.locator('table tr').count().catch(() => 0);
      if (qtd > 0) return true;
    }
    await page.waitForTimeout(intervaloMs);
  }
  return false;
}

// Espera o botão "Encaminhar" aparecer. Em e-mails com muitas mensagens na conversa o painel
// de leitura demora mais pra renderizar, e procurar o botão uma única vez logo depois de
// abrir o e-mail falhava ("não achei o botão Encaminhar").
async function esperarBotaoEncaminhar(page, timeoutMs = CONFIG.tempos.botaoEncaminhar) {
  const inicio = Date.now();
  while (Date.now() - inicio < timeoutMs) {
    const encontrado = await localizarEmContextos(todosOsContextos(page), ESTRATEGIAS_ENCAMINHAR);
    if (encontrado) return encontrado;
    await page.waitForTimeout(500);
  }
  return null;
}

// "Assinatura" do estado atual do corpo: tamanho total do texto + quantidade de linhas de
// tabela, somando todos os contextos. Se isso para de mudar, o corpo terminou de carregar.
async function medirCorpo(page) {
  let texto = 0;
  let linhas = 0;
  for (const ctx of todosOsContextos(page)) {
    try {
      texto += (await ctx.locator('body').innerText({ timeout: 2000 })).length;
      linhas += await ctx.locator('table tr').count();
    } catch {
      // contexto pode ter sido desanexado; ignora
    }
  }
  return `${texto}|${linhas}`;
}

// Espera o corpo do e-mail/rascunho ficar ESTÁVEL (sem mudar por estavelPorMs). É isso que
// resolve o caso de e-mails com muitas mensagens: antes o script só esperava aparecer
// QUALQUER "table tr" (que pode ser uma tabela de assinatura) e já lia o DOM, mesmo com o
// resto do histórico ainda sendo injetado.
async function esperarCorpoEstabilizar(
  page,
  { estavelPorMs = CONFIG.tempos.corpoEstavelPor, timeoutMs = CONFIG.tempos.corpoEstabilizar } = {}
) {
  const inicio = Date.now();
  let ultima = await medirCorpo(page);
  let desde = Date.now();
  while (Date.now() - inicio < timeoutMs) {
    await page.waitForTimeout(300);
    const atual = await medirCorpo(page);
    if (atual !== ultima) {
      ultima = atual;
      desde = Date.now();
    } else if (Date.now() - desde >= estavelPorMs) {
      return true;
    }
  }
  return false;
}

// Espera especificamente a célula com o rótulo "Registro de origem" aparecer (em vez de
// qualquer tabela). Retorna true se apareceu dentro do tempo.
async function esperarRegistroOrigem(page, timeoutMs = CONFIG.tempos.registroOrigem) {
  const inicio = Date.now();
  while (Date.now() - inicio < timeoutMs) {
    for (const ctx of todosOsContextos(page)) {
      const qtd = await ctx
        .locator('table tr td, table tr th')
        .filter({ hasText: CONFIG.labelRegistroOrigem })
        .count()
        .catch(() => 0);
      if (qtd > 0) return true;
    }
    await page.waitForTimeout(500);
  }
  return false;
}

// Rola tudo que for "rolável" até o final e volta ao topo, pra forçar o Outlook a renderizar
// trechos longos do corpo que ficam fora da área visível em conversas grandes.
async function rolarCorpoParaCarregar(page) {
  for (const destino of ['fim', 'topo']) {
    for (const ctx of todosOsContextos(page)) {
      await ctx
        .evaluate((d) => {
          const els = [document.scrollingElement, ...document.querySelectorAll('*')];
          for (const el of els) {
            if (el && el.scrollHeight > el.clientHeight + 50) el.scrollTop = d === 'fim' ? el.scrollHeight : 0;
          }
        }, destino)
        .catch(() => {});
    }
    await page.waitForTimeout(500);
  }
}

// Método principal de extração: usa o "Encaminhar", que faz o Outlook Web pré-popular
// o corpo do e-mail com TODO o histórico da conversa já expandido (é assim que
// encaminhamento sempre funcionou, pra quem recebe ver a conversa inteira sem precisar
// clicar em nada). Lemos a tabela dali e descartamos o rascunho sem enviar.
// Substitui a tentativa anterior de clicar em "..."/aria-expanded na tela (frágil e
// causava efeito colateral clicando em elementos errados da interface, como o dropdown
// de pastas "Todas as pastas ▾").
async function buscarTabelaViaEncaminhar(page, indice) {
  const encontrado = await esperarBotaoEncaminhar(page);
  if (!encontrado) {
    console.log(`  - Item ${indice}: não achei o botão "Encaminhar".`);
    return null;
  }

  await encontrado.alvo.scrollIntoViewIfNeeded({ timeout: 3000 }).catch(() => {});

  // Confere se o rascunho abriu (procurando o botão "Enviar" em qualquer contexto),
  // com um pequeno polling em vez de checar só uma vez logo após o clique.
  async function esperarRascunhoAbrir() {
    const inicio = Date.now();
    while (Date.now() - inicio < CONFIG.tempos.rascunhoAbrir) {
      const botaoEnviar = await localizarEmContextos(todosOsContextos(page), ESTRATEGIAS_ENVIAR);
      if (botaoEnviar) return true;
      await page.waitForTimeout(400);
    }
    return false;
  }

  // Tenta em cascata: clique normal -> clique forçado (ignora checagens de "está clicável"
  // do Playwright) -> clique nativo via DOM (bypassa toda a simulação de mouse, útil quando
  // o handler da lib de UI só escuta o evento 'click' bruto). Para na primeira tentativa que
  // efetivamente abrir o rascunho (confirmado pelo botão "Enviar").
  await encontrado.alvo.click({ timeout: 5000 }).catch(() => {});
  let rascunhoAbriu = await esperarRascunhoAbrir();

  if (!rascunhoAbriu) {
    await encontrado.alvo.click({ timeout: 5000, force: true }).catch(() => {});
    rascunhoAbriu = await esperarRascunhoAbrir();
  }

  if (!rascunhoAbriu) {
    await cliqueNativo(encontrado.alvo);
    rascunhoAbriu = await esperarRascunhoAbrir();
  }

  if (!rascunhoAbriu) {
    console.log(`  - Item ${indice}: cliquei em "Encaminhar" mas o rascunho não abriu.`);
    return null;
  }

  // O botão "Enviar" aparecer só confirma que o CHROME do rascunho renderizou — não que o
  // CORPO do e-mail (histórico da conversa com a tabela) já foi injetado no DOM. Em
  // conversas longas isso é assíncrono e pode demorar mais que o resto do fluxo, causando
  // "labelsEncontrados: []" mesmo com a tabela presente no e-mail original. Por isso
  // esperamos explicitamente por alguma linha de tabela aparecer em qualquer contexto
  // antes de tentar extrair, em vez de ler o DOM imediatamente.
  await esperarTabelaCarregar(page, CONFIG.tempos.tabelaAparecer);
  await esperarCorpoEstabilizar(page);

  // Em threads longas, mensagens mais antigas da conversa (onde a tabela pode estar) vêm
  // colapsadas atrás de um "..." mesmo dentro do rascunho de encaminhamento, e o corpo é
  // injetado aos poucos. Repete até N rodadas de: rolar -> expandir "..." -> esperar o corpo
  // estabilizar -> esperar o rótulo "Registro de origem" -> tentar ler a tabela.
  let resultado = { htmlConteudo: '', labelsEncontrados: [], totalLinhasVistas: 0 };
  const maxRodadas = CONFIG.tempos.rodadasExpansao;
  for (let rodada = 1; rodada <= maxRodadas; rodada++) {
    await rolarCorpoParaCarregar(page);
    await expandirTodoConteudoColapsado(page);
    await esperarCorpoEstabilizar(page);
    await esperarRegistroOrigem(page, rodada === 1 ? CONFIG.tempos.registroOrigem : CONFIG.tempos.registroOrigem / 2);
    resultado = await buscarTabelaEmContextos(todosOsContextos(page));
    if (resultado.htmlConteudo) break;
    if (rodada < maxRodadas) {
      console.log(`  - Item ${indice}: rodada ${rodada}/${maxRodadas} sem achar "Registro de origem", esperando mais...`);
    }
  }

  // Se ainda não achou como <table>, tenta o fallback por texto puro (ver
  // buscarRegistroOrigemViaTexto) antes de desistir de vez — cobre o caso de threads muito
  // aninhadas onde a tabela original chega "achatada" em texto simples nessa camada.
  let demandasFallback = [];
  if (!resultado.htmlConteudo) {
    demandasFallback = await buscarRegistroOrigemViaTexto(todosOsContextos(page));
  }

  // DIAGNÓSTICO: se mesmo assim não achou nada (nem tabela, nem fallback de texto), tira um
  // print do rascunho de encaminhamento ANTES de descartar (pra dar pra ver visualmente o
  // que estava no corpo nesse momento) e salva o texto puro de cada contexto, pra comparar
  // com o item 0 que funcionou e descobrir se é falta de tempo de carregamento, estrutura de
  // HTML diferente, ou o e-mail realmente não ter a tabela.
  if (!resultado.htmlConteudo && demandasFallback.length === 0) {
    if (!fs.existsSync(CONFIG.outputDir)) fs.mkdirSync(CONFIG.outputDir, { recursive: true });
    const printDebug = path.join(CONFIG.outputDir, `debug-rascunho-${indice}.png`);
    await page.screenshot({ path: printDebug, fullPage: true }).catch(() => {});

    const textoDebug = path.join(CONFIG.outputDir, `debug-rascunho-${indice}.txt`);
    let dump = '';
    for (const ctx of todosOsContextos(page)) {
      const texto = await ctx.locator('body').innerText().catch(() => '(sem acesso ao body desse contexto)');
      dump += `----- contexto -----\n${texto}\n\n`;
    }
    fs.writeFileSync(textoDebug, dump, 'utf-8');

    console.log(
      `  - Item ${indice}: total de linhas de tabela vistas: ${resultado.totalLinhasVistas}. Print e texto de diagnóstico salvos em output/debug-rascunho-${indice}.png/.txt`
    );
  }

  // Descarta o rascunho de encaminhamento sem enviar, pra não deixar lixo na caixa.
  // Tenta "Descartar" direto primeiro; se não achar, tenta "Fechar" (nome ANCORADO — ver
  // construirEstrategiasBotao acima — pra nunca casar com "Fechar pesquisa").
  const botaoDescartar = await localizarEmContextos(todosOsContextos(page), ESTRATEGIAS_DESCARTAR);
  if (botaoDescartar) {
    await botaoDescartar.alvo.click({ timeout: 3000 }).catch(() => {});
    await page.waitForTimeout(500);
    await confirmarDescarteSeAparecer(page);
  } else {
    const botaoFechar = await localizarEmContextos(todosOsContextos(page), ESTRATEGIAS_FECHAR);
    if (botaoFechar) {
      await botaoFechar.alvo.click({ timeout: 3000 }).catch(() => {});
      await page.waitForTimeout(500);
      const botaoDescartar2 = await localizarEmContextos(todosOsContextos(page), ESTRATEGIAS_DESCARTAR);
      if (botaoDescartar2) {
        await botaoDescartar2.alvo.click({ timeout: 3000 }).catch(() => {});
        await page.waitForTimeout(500);
        await confirmarDescarteSeAparecer(page);
      }
    } else {
      console.log(`  - Item ${indice}: não consegui descartar o rascunho de encaminhamento (pode ficar salvo em Rascunhos).`);
    }
  }

  return { ...resultado, demandasFallback };
}

async function processarEmail(page, indice, intervaloMes) {
  const itens = page.locator('[role="option"]');

  // Não filtramos mais por data nem por remetente aqui: a busca do Outlook Web
  // (subject + from:jessica + received/sent no intervalo) já garante isso.
  // Filtrar de novo lendo o texto da linha era redundante e frágil (ex: itens recentes
  // aparecem sem o ano na data, e o "remetente" visível na prévia nem sempre é a Jessica,
  // já que ela é a destinatária e pode não ser quem enviou a última mensagem da conversa).
  const textoLinha = (await itens.nth(indice).innerText().catch(() => '')) || '';

  // FILTRO: assunto, lido direto da lista, como validação extra (a query já filtra por
  // subject, mas a busca do Outlook às vezes é mais "solta" do que um match exato).
  const assuntoTexto = extrairAssuntoDoTextoDaLista(textoLinha);
  if (!assuntoTexto) {
    console.log(`  - Item ${indice}: assunto não bate na lista, pulando.`);
    return null;
  }

  console.log(`  - Item ${indice}: assunto OK, abrindo... [lista: ${textoLinha.replace(/\s+/g, ' ').slice(0, 140)}]`);

  // AJUSTAR SE PRECISAR: reabre a lista a cada iteração para evitar referências obsoletas
  await itens.nth(indice).click();
  await page.waitForTimeout(1000);

  // Print de tela cheia AQUI, antes de qualquer clique — nesse momento a mensagem mais
  // recente da conversa ("De acordo" da Jessica) já vem expandida por padrão no painel de
  // leitura, com remetente/data/hora visíveis. O recorte da região certa (só o painel de
  // leitura, sem o menu lateral) é feito depois, no gerar_pptx.py.
  if (!fs.existsSync(CONFIG.outputDir)) fs.mkdirSync(CONFIG.outputDir, { recursive: true });
  const printResposta = path.join(CONFIG.outputDir, `resposta-${indice}.png`);
  await page.screenshot({ path: printResposta }).catch(() => {});

  // Vamos direto pro "Encaminhar" como método principal (ver buscarTabelaViaEncaminhar):
  // não tentamos mais clicar em nada na tela (nem [aria-expanded="false"] nem "..."),
  // porque essa lógica era genérica/frágil demais e acabava clicando em elementos
  // errados da interface (ex: abrindo o dropdown "Todas as pastas ▾").
  const resultado = await buscarTabelaViaEncaminhar(page, indice);
  const htmlConteudo = resultado ? resultado.htmlConteudo : '';
  const labelsEncontrados = resultado ? resultado.labelsEncontrados : [];
  const demandasFallback = resultado ? resultado.demandasFallback || [] : [];

  let demandas;
  if (htmlConteudo) {
    demandas = extrairLinhasDeHtml(htmlConteudo);
  } else if (demandasFallback.length > 0) {
    console.log(`  - Item ${indice}: tabela HTML não encontrada, mas consegui recuperar via texto puro (fallback).`);
    demandas = demandasFallback;
  } else {
    console.log(
      `  - Item ${indice}: não achei o rótulo "Registro de origem". Rótulos encontrados: [${labelsEncontrados.join(' | ')}].`
    );
    return null;
  }

  console.log(`  - Item ${indice}: ${demandas.length} demanda(s) extraída(s).`);
  return { assunto: assuntoTexto, demandas, printResposta };
}

// Busca, dentro de "Itens Enviados", o e-mail original (o pedido de aprovação que o
// próprio usuário mandou) com o MESMO assunto exato do e-mail já processado, e tira um
// print de tela cheia dele — mostrando remetente/destinatário/data/hora, igual ao print
// da "resposta" que já capturamos em processarEmail. Roda DEPOIS do loop principal (não
// no meio dele), pra não precisar voltar pra busca original no meio do processamento.
async function capturarPrintPedidoAprovacao(page, assuntoTexto, indice) {
  // AJUSTAR SE PRECISAR: nome da pasta pode variar (ex: "Sent Items" em inglês)
  const linkItensEnviados = page.getByText(/^\s*(itens enviados|sent items)\s*$/i).first();
  if (!(await linkItensEnviados.count().catch(() => 0))) {
    console.log(`  - Item ${indice}: não achei a pasta "Itens Enviados" pra buscar o pedido de aprovação.`);
    return null;
  }
  await linkItensEnviados.click();
  await page.waitForTimeout(2000);

  const caixaBusca = localizarCaixaBusca(page);
  if (!(await caixaBusca.first().count().catch(() => 0))) {
    console.log(`  - Item ${indice}: não achei a caixa de busca dentro de "Itens Enviados".`);
    return null;
  }
  await caixaBusca.first().click();
  await page.keyboard.press('Control+A');
  await page.keyboard.press('Backspace');
  // Busca pelo assunto exato (sem os filtros de participants/data — aqui já estamos
  // dentro da pasta certa e o assunto sozinho já deve ser específico o suficiente).
  await caixaBusca.first().pressSequentially(`subject:"${assuntoTexto}"`, { delay: 20 });
  await page.waitForTimeout(500);
  await page.keyboard.press('Enter');
  await page.waitForTimeout(3000);

  const itens = page.locator('[role="option"]');
  const total = await itens.count().catch(() => 0);
  if (total === 0) {
    console.log(`  - Item ${indice}: não achei o e-mail de pedido em "Itens Enviados" (assunto: "${assuntoTexto}").`);
    return null;
  }

  await itens.first().click();
  await page.waitForTimeout(1500);

  const printPedido = path.join(CONFIG.outputDir, `pedido-${indice}.png`);
  await page.screenshot({ path: printPedido }).catch(() => {});

  return printPedido;
}

// ------------------------- GERAÇÃO DO PPTX (via Python + modelo corporativo) -------------------------

// A geração de verdade é feita pelo gerar_pptx.py, que abre o modelo release_para_dev.pptx
// e duplica os slides dele (preservando 100% da formatação/fundo/logo da Livre) — muito
// mais fiel do que recriar o design do zero aqui em JS. O Node só monta os dados e chama
// o script; ver gerar_pptx.py para a lógica de montagem do deck em si.
async function gerarPptxViaPython(resultados, intervaloMes) {
  if (!fs.existsSync(CONFIG.templatePptx)) {
    throw new Error(
      `Modelo de PPTX não encontrado em "${CONFIG.templatePptx}". Coloque o arquivo ` +
        `release_para_dev.pptx nessa pasta (ou ajuste CONFIG.templatePptx) antes de rodar.`
    );
  }

  const demandas = [];
  const aprovacoes = [];
  for (const r of resultados) {
    demandas.push(...r.demandas);
    if (r.printPedido && r.printResposta) {
      aprovacoes.push({ pedido: r.printPedido, resposta: r.printResposta });
    } else {
      console.log(`  - Aviso: e-mail "${r.assunto}" sem os dois prints de aprovação — não entra no slide de Aprovações.`);
    }
  }

  const nomeArquivo = path.join(
    CONFIG.outputDir,
    `GMUD-${intervaloMes.inicio.getFullYear()}-${String(intervaloMes.inicio.getMonth() + 1).padStart(2, '0')}.pptx`
  );

  const config = {
    template: CONFIG.templatePptx,
    output: nomeArquivo,
    mesAno: formatarMesAno(intervaloMes.inicio),
    demandas,
    aprovacoes,
  };

  const caminhoConfig = path.join(CONFIG.outputDir, 'config-pptx.json');
  fs.writeFileSync(caminhoConfig, JSON.stringify(config, null, 2), 'utf-8');

  const resultado = spawnSync(CONFIG.pythonBin, [CONFIG.gerarPptxScript, caminhoConfig], {
    encoding: 'utf-8',
  });

  if (resultado.error) {
    throw new Error(`Falha ao rodar o Python (${CONFIG.pythonBin} não encontrado?): ${resultado.error.message}`);
  }
  if (resultado.status !== 0) {
    throw new Error(`gerar_pptx.py terminou com erro:\n${resultado.stderr || resultado.stdout}`);
  }

  console.log(`\nPPTX gerado em: ${nomeArquivo}`);
  return nomeArquivo;
}

// ------------------------- FLUXO PRINCIPAL -------------------------

async function main() {
  const context = await chromium.launchPersistentContext(CONFIG.userDataDir, {
    headless: false, // precisa ser visível pelo menos na 1ª vez (login/MFA)
    // IMPORTANTE: viewport fixo — os prints capturados aqui são recortados por coordenada
    // fixa depois (ver recortar_print_leitura em gerar_pptx.py). Se o viewport variar de
    // máquina pra máquina, os recortes saem errados.
    viewport: { width: 1280, height: 720 },
  });

  try {
    const page = await abrirOutlookELogar(context);
    const { total, intervaloMes } = await buscarEmails(page);

    const resultados = [];
    for (let i = 0; i < total; i++) {
      const resultado = await processarEmail(page, i, intervaloMes);
      if (resultado) resultados.push(resultado);
    }

    // Captura dos prints de "pedido de aprovação" DEPOIS do loop principal (não no meio
    // dele) — evita ter que voltar pra busca original a cada iteração, já que buscar em
    // "Itens Enviados" navega a página inteira pra outro lugar.
    for (let i = 0; i < resultados.length; i++) {
      resultados[i].printPedido = await capturarPrintPedidoAprovacao(page, resultados[i].assunto, i);
    }

    if (resultados.length === 0) {
      console.log('\nNenhum e-mail válido processado. PPTX não foi gerado.');
    } else {
      await gerarPptxViaPython(resultados, intervaloMes);
    }

    await aguardarEnter(
      '\n>> Terminado. Dá uma olhada no PPTX (se gerado) na pasta output/.\n' +
        '>> Pressione ENTER quando quiser fechar o navegador...\n'
    );
  } finally {
    await context.close();
  }
}

main().catch((err) => {
  console.error('Erro ao executar o script:', err);
  process.exit(1);
});
