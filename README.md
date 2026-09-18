# outlook-gmud-pptx

Script Node.js que abre o Outlook Web, busca e-mails de "De acordo de negócios GMUD"
enviados para `jessica.fachina@unidas.com.br` dentro do mês vigente, extrai a lista de
"Registro de origem" da tabela "UAT - Implantação GMUD" e gera um PPTX (uma slide por
e-mail encontrado).

Como não há acesso para registrar um app no Azure AD/Entra ID, a autenticação é feita
via **automação de navegador** (login manual na primeira execução, sessão salva localmente
para as próximas).

## Instalação

```bash
npm install
npx playwright install chromium
```

## Uso

```bash
npm start
```

Na primeira execução:
1. Uma janela do Chromium vai abrir no Outlook Web.
2. Faça login manualmente (usuário, senha e MFA).
3. Quando a caixa de entrada carregar, volte ao terminal e pressione ENTER.

Nas próximas execuções, a sessão salva em `.owa-session/` normalmente evita pedir login
de novo (a menos que o token expire por política do tenant — nesse caso, basta logar de
novo quando pedido).

O PPTX gerado fica em `output/GMUD-AAAA-MM.pptx`.

## Se algo não funcionar (seletores de tela)

O Outlook Web pode variar pequenos detalhes de tela (idioma, versão, tenant). Os pontos
mais sensíveis a ajustar, marcados no código com `AJUSTAR SE PRECISAR`, são:

- **Caixa de busca**: seletor por `aria-label` ("Search"/"Pesquisar").
- **Lista de e-mails**: seletor `[role="option"]`.
- **Iframe do corpo do e-mail**: `iframe[title="Message Body"]` (ou variações).

Se o script não encontrar elementos, rode e observe a janela do navegador: com
`headless: false` (padrão já configurado) dá pra ver exatamente onde ele trava, abrir o
DevTools (F12) e conferir o seletor certo para seu tenant.

## Regras implementadas

- **Filtro de e-mails**: assunto contendo "De acordo de negócios GMUD" (case-insensitive),
  destinatário `jessica.fachina@unidas.com.br`, dentro do **mês anterior** ao mês em que o
  script é executado (dia 1 ao último dia corrido do mês passado).
- **Extração**: dentro da tabela "UAT - Implantação GMUD" (ou variações), localiza a linha
  cujo rótulo (coluna esquerda) seja "Registro de origem" e extrai o conteúdo da coluna
  direita, separando por bullet ou quebra de linha.
- **PPTX**: um slide por e-mail processado, com o assunto do e-mail e a lista de demandas
  em bullets.
