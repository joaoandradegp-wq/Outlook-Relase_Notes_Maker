const { chromium } = require('playwright');
const { spawnSync } = require('child_process');
const path = require('path');
const fs = require('fs');
const readline = require('readline');

const CONFIG = {
  userDataDir: path.join(__dirname, '.owa-session'),
  outlookUrl: 'https://outlook.office.com/mail/',
  destinatario: 'jessica.fachina@unidas.com.br',
  assuntoBusca: 'de acordo de negocios gmud',
  assuntoRegex: /de\s*acordo\s*de\s*neg[oó]cios\s*gmud/i,
  tituloTabelaRegex: /UAT\s*-\s*Implanta[cç][aã]o\s*GMUD/i,
  labelRegistroOrigem: /registro\s*de\s*origem/i,
  regexLinhaDemanda: /^(user\s*story\s*)?\d{4,}/i,
  outputDir: path.join(__dirname, 'output'),
  // TEMPOS DE ESPERA (ms) do fluxo "Encaminhar". Se algum e-mail ainda vier sem as demandas,
  // aumente estes valores (principalmente registroOrigem e corpoEstabilizar).
  remetenteAprovador: /jessica\s*fachina|jessica\.fachina@unidas\.com\.br/i,

  tempos: {
    botaoEncaminhar: 10000,
    rascunhoAbrir: 8000,
    tabelaAparecer: 10000,
    corpoEstavelPor: 1500,
    corpoEstabilizar: 15000,
    registroOrigem: 15000,
    rodadasExpansao: 3,
  },
  templatePptx: path.join(__dirname, 'templates', 'release_para_dev.pptx'),
  gerarPptxScript: path.join(__dirname, 'gerar_pptx.py'),
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
  const mes = MESES_PT[data.getMonth()];
  const anoCurto = String(data.getFullYear()).slice(-2);
  return `${mes}/${anoCurto}`;
}

function getIntervaloMesAnterior() {
  const agora = new Date();

  const mesEscolhido = parseInt(process.env.GMUD_MES, 10);
  if (mesEscolhido >= 1 && mesEscolhido <= 12) {
    let ano = parseInt(process.env.GMUD_ANO, 10);
    if (!ano) ano = mesEscolhido > agora.getMonth() + 1 ? agora.getFullYear() - 1 : agora.getFullYear();
    return { inicio: new Date(ano, mesEscolhido - 1, 1), fim: new Date(ano, mesEscolhido, 0) };
  }

  const inicio = new Date(agora.getFullYear(), agora.getMonth() - 1, 1);
  const fim = new Date(agora.getFullYear(), agora.getMonth(), 0);
  return { inicio, fim };
}

function formatarDataOWA(data) {
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

async function abrirOutlookELogar(context) {
  const page = context.pages()[0] || (await context.newPage());
  await page.goto(CONFIG.outlookUrl, { waitUntil: 'domcontentloaded' });

  await aguardarEnter(
    '\n>> Se for necessário, faça login manualmente (usuário/senha/MFA) na janela aberta.\n' +
      '>> Quando a caixa de entrada estiver carregada, volte aqui e pressione ENTER para continuar...\n'
  );

  return page;
}

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

  await page.keyboard.press('Control+A');
  await page.keyboard.press('Backspace');

  await caixaBusca.first().pressSequentially(query, { delay: 30 });
  await page.waitForTimeout(500);
  await page.keyboard.press('Enter');

  await page.waitForTimeout(4000);

  const itens = page.locator('[role="option"]');
  let total = await itens.count();

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

function extrairAssuntoDoTextoDaLista(textoLinha) {
  const linhas = textoLinha.split('\n').map((l) => l.trim()).filter(Boolean);
  return linhas.find((l) => CONFIG.assuntoRegex.test(l)) || null;
}

async function buscarTabelaEmContextos(contextos) {
  let htmlConteudo = '';
  const labelsEncontrados = [];
  let totalLinhasVistas = 0;

  for (const ctx of contextos) {
    let linhas = [];
    try {
      linhas = await ctx.locator('table tr').all();
    } catch {
      continue;
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

async function buscarRegistroOrigemViaTexto(contextos) {
  for (const ctx of contextos) {
    let texto = '';
    try {
      texto = await ctx.locator('body').innerText();
    } catch {
      continue;
    }

    const linhas = texto.split('\n').map((l) => l.trim());
    const idxRotulo = linhas.findIndex(
      (l) => CONFIG.labelRegistroOrigem.test(l) && l.length <= 40
    );
    if (idxRotulo === -1) continue;

    const demandas = [];
    for (let i = idxRotulo + 1; i < linhas.length; i++) {
      const linha = linhas[i];
      if (!linha) break;
      if (!CONFIG.regexLinhaDemanda.test(linha)) break;
      demandas.push(linha);
    }

    if (demandas.length > 0) return demandas;
  }
  return [];
}

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
const ESTRATEGIAS_CONFIRMAR = construirEstrategiasBotao(['ok', 'sim'], ['ok', 'yes']);

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
    await esperarCorpoEstabilizar(page, { estavelPorMs: 800, timeoutMs: 5000 });
  }
  return cliques > 0;
}

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

async function esperarBotaoEncaminhar(page, timeoutMs = CONFIG.tempos.botaoEncaminhar) {
  const inicio = Date.now();
  while (Date.now() - inicio < timeoutMs) {
    const encontrado = await localizarEmContextos(todosOsContextos(page), ESTRATEGIAS_ENCAMINHAR);
    if (encontrado) return encontrado;
    await page.waitForTimeout(500);
  }
  return null;
}

async function medirCorpo(page) {
  let texto = 0;
  let linhas = 0;
  for (const ctx of todosOsContextos(page)) {
    try {
      texto += (await ctx.locator('body').innerText({ timeout: 2000 })).length;
      linhas += await ctx.locator('table tr').count();
    } catch {
    }
  }
  return `${texto}|${linhas}`;
}

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

async function lerRemetentesDoRascunho(contextos) {
  const remetentes = [];
  for (const ctx of contextos) {
    let texto = '';
    try {
      texto = await ctx.locator('body').innerText({ timeout: 3000 });
    } catch {
      continue;
    }
    const linhas = texto.split('\n').map((l) => l.trim());
    for (let i = 0; i < linhas.length; i++) {
      const m = linhas[i].match(/^(?:De|From)\s*:\s*(.*)$/i);
      if (!m) continue;
      let valor = m[1].trim();
      if (!valor) valor = (linhas.slice(i + 1).find(Boolean) || '').trim();
      if (valor && !remetentes.includes(valor)) remetentes.push(valor);
    }
  }
  return remetentes;
}

async function buscarTabelaViaEncaminhar(page, indice) {
  const encontrado = await esperarBotaoEncaminhar(page);
  if (!encontrado) {
    console.log(`  - Item ${indice}: não achei o botão "Encaminhar".`);
    return null;
  }

  await encontrado.alvo.scrollIntoViewIfNeeded({ timeout: 3000 }).catch(() => {});

  async function esperarRascunhoAbrir() {
    const inicio = Date.now();
    while (Date.now() - inicio < CONFIG.tempos.rascunhoAbrir) {
      const botaoEnviar = await localizarEmContextos(todosOsContextos(page), ESTRATEGIAS_ENVIAR);
      if (botaoEnviar) return true;
      await page.waitForTimeout(400);
    }
    return false;
  }

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

  await esperarTabelaCarregar(page, CONFIG.tempos.tabelaAparecer);
  await esperarCorpoEstabilizar(page);

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

  let demandasFallback = [];
  if (!resultado.htmlConteudo) {
    demandasFallback = await buscarRegistroOrigemViaTexto(todosOsContextos(page));
  }

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
  const remetentes = await lerRemetentesDoRascunho(todosOsContextos(page));

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

  const assuntoTexto = extrairAssuntoDoTextoDaLista(textoLinha);
  if (!assuntoTexto) {
    console.log(`  - Item ${indice}: assunto não bate na lista, pulando.`);
    return null;
  }

  console.log(`  - Item ${indice}: assunto OK, abrindo... [lista: ${textoLinha.replace(/\s+/g, ' ').slice(0, 140)}]`);

  await itens.nth(indice).click();
  await page.waitForTimeout(1000);

  if (!fs.existsSync(CONFIG.outputDir)) fs.mkdirSync(CONFIG.outputDir, { recursive: true });
  const printResposta = path.join(CONFIG.outputDir, `resposta-${indice}.png`);
  await page.screenshot({ path: printResposta }).catch(() => {});

  const resultado = await buscarTabelaViaEncaminhar(page, indice);
  if (resultado && CONFIG.remetenteAprovador) {
    const remetentes = resultado.remetentes || [];
    if (remetentes.length === 0) {
      console.log(`  - Item ${indice}: aviso — não consegui ler os remetentes do histórico; não dá pra confirmar que é da Jessica, seguindo mesmo assim.`);
    } else if (!remetentes.some((r) => CONFIG.remetenteAprovador.test(r))) {
      console.log(`  - Item ${indice}: IGNORADO — a Jessica não enviou nenhuma mensagem nessa conversa (remetentes vistos: ${remetentes.join(' | ')}).`);
      return null;
    }
  }

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

async function capturarPrintPedidoAprovacao(page, assuntoTexto, indice) {
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

async function main() {
  const context = await chromium.launchPersistentContext(CONFIG.userDataDir, {
    headless: false,
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
