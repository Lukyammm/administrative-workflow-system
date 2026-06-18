# Relatório de Revisão de Código — Erros Graves

> Atualizado em 2026-06-18. Esta revisão analisou o backend Apps Script
> (`CODE.gs`) e o front-end (`index.html`, `script.html`). As correções
> descritas abaixo já foram aplicadas nesta branch.

## Visão geral

O Cosign gerencia armários hospitalares e manipula dados sensíveis (pacientes,
prontuários, acompanhantes, termos de responsabilidade). A revisão focou em
**erros graves** de segurança, autorização e integridade. O problema central
era que **a autorização era controlada pelo cliente**: o `usuarioId` usado para
decidir permissões vinha cru dos parâmetros da requisição, e a maioria das ações
não fazia nenhuma verificação de sessão.

## Problemas críticos (corrigidos)

### 1. Autorização controlada pelo cliente → token de sessão server-side
**Antes:** `definirContextoUsuario` lia `usuarioId` dos parâmetros e
`validarPermissaoAdmin` só checava se esse id era admin na planilha. Qualquer
requisição com `usuarioId=1` ao endpoint público `/exec` assumia o admin. Além
disso, leituras de PII (`getHistorico`, `getTermo`, `getMovimentacoes`, etc.) não
tinham verificação alguma.

**Correção:**
- Login (`autenticarUsuario`) agora **emite um token** de sessão. As sessões são
  persistidas server-side numa aba `Sessões` (token armazenado **hasheado**,
  com `usuarioId`, `perfil`, flag de troca de senha e expiração de 8h).
- `resolverSessao(token)` valida a sessão e retorna o `usuarioId`/`perfil`
  **da planilha**, ignorando qualquer `usuarioId` enviado pelo cliente.
- **Gate central** em `handlePost`: ações são classificadas em públicas
  (`autenticarUsuario`, `verificarInicializacao`, `inicializarPlanilha`),
  autenticadas (exigem token válido) e administrativas (exigem perfil `admin`).
  Requisições sem sessão válida recebem acesso negado antes de qualquer execução.
- `validarPermissaoAdmin` passou a confiar somente no id resolvido pela sessão.
- Logout/expiração: ação `encerrarSessao` remove a sessão server-side.
- Front-end: o token é guardado e enviado em toda chamada (`callGoogleScript`);
  respostas de sessão inválida forçam novo login.

### 2. Senha admin padrão `admin123` → troca obrigatória no primeiro acesso
**Antes:** o seed criava o admin com a senha conhecida `admin123` sem obrigar a
troca.

**Correção:** nova coluna `Precisa Trocar Senha` na aba `Usuários`. O admin
semeado (e novos usuários criados pelo admin) são marcados para troca obrigatória;
o gate bloqueia todas as ações (exceto `trocarSenha`/`encerrarSessao`) até a troca.
Nova ação `trocarSenha` (valida senha atual, exige mínimo de 6 caracteres) e modal
forçado no front-end.

## Problemas altos/médios (corrigidos)

### 3. Vazamento de exceções ao cliente
**Antes:** ~40 retornos `error: error.toString()` expunham stack/internos.
**Correção:** os retornos ao cliente passaram a usar mensagem genérica
("Erro ao processar a solicitação."); o detalhe continua registrado server-side
via `registrarLog('ERRO', ...)`.

### 4. Clickjacking (`XFrameOptionsMode.ALLOWALL`)
**Antes:** `doGet` permitia embed em qualquer iframe.
**Correção:** trocado para `XFrameOptionsMode.DEFAULT`.

### 5. `innerHTML` com conteúdo de célula (XSS armazenado)
**Antes:** a visão mobile de tabelas fazia `valor.innerHTML = celula.innerHTML`,
reinterpretando HTML que pode derivar de dados do usuário.
**Correção:** os nós DOM já renderizados são **clonados** (`cloneNode`), sem
re-parsear HTML; fallback para `textContent`.

### 6. Comparação frouxa de IDs (`==`)
**Antes:** `linha[0] == dadosTermo.armarioId` (número vs string).
**Correção:** comparação normalizada com `String(...).trim() === ...`.

## Observações e limitações

- **Endpoint público (`SCRIPT_URL`)**: num web app Apps Script publicado, a URL
  `/exec` é pública por design. A segurança real agora está no token de sessão e
  no gate server-side, não no segredo da URL.
- **Hashing de senha**: usa SHA-256 + salt por usuário (`calcularHashSenha`).
  Para um hardening futuro, considerar key-stretching; o Apps Script não oferece
  bcrypt/Argon2 nativamente.
- **Sessões em planilha**: cada requisição autenticada lê a aba `Sessões`
  (renovação de TTL é preguiçosa para reduzir escritas). Adequado à escala do
  sistema; para volumes maiores, avaliar `CacheService`/`PropertiesService`.

## Como testar

1. Rodar `inicializarPlanilha` (cria a aba `Sessões` e a coluna
   `Precisa Trocar Senha`).
2. **Bypass**: chamar o `/exec` com `action=getHistorico`/`getUsuarios` sem token
   (ou com `usuarioId=1`) → deve retornar acesso negado.
3. **Fluxo feliz**: login na UI → recebe token; admin semeado é forçado a trocar
   `admin123` no primeiro acesso; navegação normal após a troca.
4. **Expiração/logout**: após `encerrarSessao` ou TTL, ações voltam a exigir login.
5. Conferir que a página não pode ser embutida em iframe externo.
