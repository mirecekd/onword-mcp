# Office Add-in jako alternativa/doplněk k onword-mcp

Rešerše: 2026-07-28

Otázka, která rešerši vyvolala: *"Dalo by se onword-mcp přepsat jako add-in do Office?"*

Krátká odpověď: **Ne jako přepis. Add-in nemůže být MCP server.** Existují tři reálné
architektury (A/B/D níže), které se liší tím, kdo řídí agenta a co za to uživatel zaplatí
při instalaci. Společným základem všech tří je jedna a tatáž vrstva: Office.js implementace
současných operací (`ops.js`).

## Obsah

- [1. Zásadní omezení: add-in nemůže být MCP server](#1-zásadní-omezení-add-in-nemůže-být-mcp-server)
- [2. Capability matrix: Office.js vs. současný COM backend](#2-capability-matrix-officejs-vs-současný-com-backend)
- [3. Blocker: Chrome 142+ Local Network Access](#3-blocker-chrome-142-local-network-access)
- [4. Průzkum existujících projektů](#4-průzkum-existujících-projektů)
- [5. Pi agent: browser-capable, ale bez MCP](#5-pi-agent-browser-capable-ale-bez-mcp)
- [6. CLI agenti v prohlížeči: nelze](#6-cli-agenti-v-prohlížeči-nelze)
- [7. ACP - Agent Client Protocol](#7-acp---agent-client-protocol)
- [8. Zamítnuté alternativy](#8-zamítnuté-alternativy)
- [9. Varianty A / B / D](#9-varianty-a--b--d)
- [10. Doporučený postup](#10-doporučený-postup)
- [11. Body k doověření před implementací](#11-body-k-doověření-před-implementací)

## 1. Zásadní omezení: add-in nemůže být MCP server

Taskpane Office add-inu je webová stránka běžící ve WebView2 (Windows), WKWebView (macOS)
nebo v iframe s HTML5 `sandbox` (Word na webu). Webová platforma nemá `bind()`/`listen()` -
nelze otevřít poslouchající TCP socket.

Jediná specifikace, která by to umožnila, je WICG Direct Sockets (`TCPServerSocket`), a její
explainer uvádí, že API bude dostupné **pouze pro Isolated Web Apps**, gated přes
`direct-sockets` Permissions-Policy. Office add-in není IWA.

Důsledek: **MCP klient (Cline, Claude Desktop) se nikdy nemůže připojit přímo do add-inu.**
Add-in umí pouze outbound spojení:

- `fetch` / XHR
- WebSocket klient
- SSE

Z toho plyne jediná možná topologie pro MCP scénář - add-in je *executor*, MCP server je
samostatný lokální proces, který drží WebSocket a posílá do taskpane příkazy:

```text
Cline/Claude  --MCP(stdio|http)-->  lokální bridge proces (Python/Node)
                                          ^
                                          | WebSocket
                                          | (bridge = WS server, taskpane = WS klient)
                                          v
                                   Word taskpane add-in (Office.js) --> otevřený dokument
```

Ani `Office.context.ui.openBrowserWindow` ani `displayDialogAsync` na tomto nic nemění.

## 2. Capability matrix: Office.js vs. současný COM backend

### Requirement sets - orientace

| Sada | Platformy |
|---|---|
| `WordApi 1.1` - `1.9` | cross-platform; 1.9 = web + Windows M365 2411+ + Mac 16.91 + iPad |
| `WordApiDesktop 1.1` - `1.5` | **pouze Windows/Mac (část iPad) - nikdy Word na webu** |
| `WordApiOnline 1.1` | aktuálně prázdná (žádná API) |
| `WordApi BETA` | preview CDN, ne produkce |

### Mapování současných toolů

| Tool (COM) | Office.js ekvivalent | Req. set |
|---|---|---|
| `get_document_info` | `document.properties`, `body.paragraphs` | 1.1 |
| `get_document_outline` | `body.paragraphs` + `style`, `listItemOrNullObject.level`, `text` | 1.1 / 1.3 |
| `read_paragraphs` | `paragraph.text` | 1.1 |
| `find_text` | `body.search(q, {...})` | 1.1 |
| `get_page_paragraphs` | `Word.Page`, `pane.pages`, `page.index`, `page.getRange()` | **WordApiDesktop 1.2** |
| `get_selection` | `document.getSelection()` | 1.1 |
| `replace_paragraph` | `paragraph.getRange().insertText(t, "Replace")` | 1.1 |
| `insert_after_paragraph` / `insert_before_paragraph` | `paragraph.insertParagraph(t, "After"/"Before")` | 1.1 |
| `append_to_document` | `body.insertParagraph(t, "End")` | 1.1 |
| `delete_paragraphs` | `paragraph.delete()` | 1.1 |
| `replace_text_in_paragraph` | `paragraph.search()` + `range.insertText(..., "Replace")` | 1.1 |
| `insert_at_selection` | `document.getSelection().insertText(t, "Replace")` | 1.1 |
| `set_track_changes` | `document.changeTrackingMode` | 1.4 |
| `save_document` | `document.save()` | 1.1 |
| `get_paragraph_formatting` | `paragraph.style`, `alignment`, `leftIndent`, `font` | 1.1 |
| `list_document_styles` | `document.getStyles()` | 1.5 |
| `set_paragraph_style` | `paragraph.style` / `styleBuiltIn` | 1.1 |
| `list_indent` / `list_outdent` / `set_list_level` | `paragraph.listItem.level`, `list.setLevelIndents` | 1.3 |
| `convert_to_list` / `remove_list` | `paragraph.startNewList()`, `attachToList()`, `detachFromList()` | 1.3 |
| `set_paragraph_indent` | `paragraph.leftIndent`, `firstLineIndent` | 1.1 |
| `set_paragraph_alignment` | `paragraph.alignment` | 1.1 |
| `format_text` | `range.font.bold` / `italic` / `size` / `name` | 1.1 |

Pokrytí je tedy prakticky úplné.

### Co se Office.js přechodem získá

- Funguje i na **macOS, iPadu a ve Wordu na webu** (kromě `Word.Page`), ne jen Windows.
- Odpadá COM: `CoInitialize`, `GetActiveObject`, cross-thread marshalling.
- Odpadá problém s AV blokací `uvx` trampolíny (viz Troubleshooting v README).
- Track changes a co-authoring přes Microsoftem podporovanou cestu, ne "cizí proces
  šťourající ve Wordu".
- `context.sync()` batchuje více operací do jednoho round-tripu - na velkých dokumentech
  potenciálně rychlejší než COM po jedné operaci.

### Co se ztratí nebo zkomplikuje

- **Uživatel musí mít otevřený taskpane.** Zavře panel, backend je mrtvý. COM funguje vždy,
  když Word běží.
- Není ekvivalent `ScreenUpdating = False` (částečně kompenzuje batching).
- Čísla stránek jen na desktopu (`WordApiDesktop 1.2`), na webu nutná graceful degradace.
- **Indenty v bodech, ne v centimetrech** - současné API má `left_cm` / `first_line_cm`,
  bude potřeba přepočet (1 cm = 28.3465 bodu).
- Odstavce v tabulkách nejsou v `body.paragraphs` u req. setů 1.1/1.2, až od **1.3**.
- Deployment je výrazně těžší - viz sekce 3 a 9.

### Poznámky k výkonu

- `body.paragraphs.load()` na dokumentu s 1000+ odstavci je drahé; načítat jen potřebné
  properties a stránkovat stejně jako dnes.
- Pozor na `context.trackedObjects` - untracked objekty přežijí jen do nejbližšího
  `context.sync()`.

## 3. Blocker: Chrome 142+ Local Network Access

**Chrome 142+ zavádí Local Network Access (LNA):** stránka z veřejné origin (např. GitHub
Pages) nesmí bez explicitního povolení kontaktovat `localhost` - ani přes `fetch`, ani přes
WebSocket.

To ruší kombinaci, která by jinak byla ideální: *cert-free hosting taskpane na GitHub Pages*
plus *`wss://localhost` bridge na lokální MCP server*. Platí to shodně pro PTY terminál i pro
ACP klienta.

| | čistě GH Pages | GH Pages + localhost | lokální hosting taskpane |
|---|---|---|---|
| Bez dev certu | ano | ano | **ne** |
| Bez běžícího procesu | ano | **ne** | ne |
| Dosáhne na lokální agenta/MCP | **ne** | LNA prompt (fragilní) | ano |
| Dosáhne na Word (Office.js) | ano | ano | ano |

Volba je tedy binární: **buď bezbolestná instalace, nebo lokální most. Ne oboje.**

## 4. Průzkum existujících projektů

### hewliyang/office-agents - nejlepší blueprint pro variantu A

Monorepo tří samostatných add-inů (Word, Excel, PowerPoint), každý s vlastním manifestem a
portem. Klíčová část je `packages/bridge` (5 souborů, ~2100 LOC: `protocol.ts`, `server.ts`,
`client.ts`, `http-client.ts`, `cli.ts`).

Architektura:

- Node proces = **HTTPS + WebSocket server na `localhost:4017`**, TLS z
  `~/.office-addin-dev-certs/`.
- Taskpane (běží na `https://localhost:3002`) je **WS klient** - vytočí
  `wss://localhost:4017/ws` a pošle `hello` se snapshotem session včetně **seznamu svých
  toolů**.
- Poté se role obrací: **Node posílá `invoke`, taskpane odpovídá `response`** (reverse RPC,
  prohlížeč je RPC target).
- Externí volající (jejich CLI; u nás by to byl MCP server) mluví na tentýž port přes plain
  **HTTPS REST API**.

Konstanty z `protocol.ts`:

```ts
export const BRIDGE_PROTOCOL_VERSION = 1;
export const DEFAULT_BRIDGE_HOST = "localhost";
export const DEFAULT_BRIDGE_PORT = 4017;
export const DEFAULT_BRIDGE_WS_PATH = "/ws";
```

Hodnocení: architektura je 1:1 to, co varianta A potřebuje, a je hotová a funkční. Slabina
jsou Word tooly - jen 4 read tooly plus `execute_office_js` escape hatch, žádné chirurgické
operace nad paragraph indexy. **MCP zde není vůbec** (0 výskytů `mcp` /
`modelcontextprotocol` / `streamable`).

Verdikt: **portovat protokol, tooly napsat vlastní.**

### damianofalcioni/pi-for-word - nejlepší hosting a shell

Pozor: default branch je **`master`**, ne `main` (raw URL s `/main/` vrací 404).

- Word taskpane add-in, klasický **XML manifest** (`OfficeApp` `xsi:type="TaskPaneApp"` +
  `VersionOverridesV1_0`), ne unified/JSON manifest.
- Host pouze Word (`<Hosts><Host Name="Document"/>`), `DesktopFormFactor`,
  `Permissions: ReadWriteDocument`.
- Dva manifesty: `manifest.xml` (localhost:3000) a `manifest.production.xml` (GitHub Pages).
- Vevnitř běží celý **Pi agent** (`@mariozechner/pi-agent-core` + `pi-web-ui` `ChatPanel`)
  ve verzi 0.73.1 - tzn. plnohodnotný chat klient uvnitř Wordu.
- Licence **MIT**.

**Hosting je největší nález celé rešerše:** statický build na GitHub Pages, **žádný
localhost, žádné dev certy, žádný běžící proces**. Instalace = uživatel nasdílí manifest
ukazující na GH Pages.

Slabiny:

- 4 Word tooly a **nejsou chirurgické** - vkládají přes Markdown→HTML a pozici hledají
  textovými "anchory". Žádné paragraph indexy.
- **Past:** JS REPL tool na Office.js vůbec nedosáhne - běží v sandboxovaném iframe. Úvaha
  "vezmu REPL a mám okamžitou paritu" tedy nefunguje.
- API key je BYO, uložený v taskpane. U Anthropicu z prohlížeče je potřeba
  `anthropic-dangerous-direct-browser-access` header.

Verdikt: **fork shellu a hosting patternu ano, tooly napsat celé znovu.**

### patniko/github-copilot-office - slepá ulička

Electron tray app + Express 5 + WebSocket. Na první pohled stejný pattern jako
office-agents, ale:

- Express server je **hloupá roura** - žádný tool registry, žádné timeouty, žádný reconnect.
- **Žádný externí vstupní bod.** Směr RPC je
  `Copilot CLI child process → stdio → Express/WS → taskpane`.
- Port by znamenal reimplementovat v Pythonu *druhou* polovinu protokolu, tedy LSP-framed
  JSON-RPC klienta Copilot CLI - a ten kód v repu není, je to closed-source
  `@github/copilot` npm balík.

Hosty: Word, OneNote, PowerPoint, Excel (Excel podle issue #8 nestabilní). Druhý contributor
je SteveSandersonMS (Blazor lead), takže nejde o úplný toy.

Verdikt: **použitelné jsou jen manifest/sideload triky.**

### lancedesk/ms-office-ai-helper - zahodit

Paster výstupu z ChatGPT do dokumentu plus `eval()` escape hatch. Žádný transport, žádný
tool protokol.

### dsbissett/office-addin-mcp - zajímavá, ale fragilní cesta

Go binary, MCP přes stdio. Připojuje se **zvenčí na WebView2 přes Chrome DevTools Protocol**
a injektuje Office.js. Vyžaduje spouštět Office s
`WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9222`.

Funguje bez psaní vlastního add-inu, ale je to hack závislý na spuštění Office se speciální
env proměnnou.

### GongRzhe/Office-Word-MCP-Server - jiný problém

python-docx, ~2.1k stars. Pracuje nad **zavřenými soubory**. Nedokáže se dotknout otevřeného
dokumentu, tedy neřeší use case onword-mcp.

## 5. Pi agent: browser-capable, ale bez MCP

Repo `earendil-works/pi` (dříve `badlogic/pi-mono`), autor Mario Zechner. Balíčky
`@mariozechner/pi-agent-core`, `pi-ai`, `pi-web-ui`; `pi-for-word` používá 0.73.1, upstream
je na 0.82.x.

Je to **skutečný agent, ne wrapper**: multi-turn tool-calling loop, streaming,
sessions/resume, custom system prompt, context compaction, model picker, cost/token display,
attachments, artifacts panel. `pi-web-ui` je browser-native (ESM, žádné node builtiny),
takže v taskpane skutečně běží.

**Ale Pi nemá MCP, a to záměrně.** Z README `@mariozechner/pi-coding-agent`:

> **No MCP.** Build CLI tools with READMEs (see Skills), or build an extension that adds MCP
> support.

A z `docs/usage.md`:

> "It intentionally does not include built-in MCP, sub-agents, permission popups, plan mode,
> to-dos, or background bash."

Ověřeno grepem přes rozbalené tarbally:

| Balíček | Výskyty `mcp` / `modelcontextprotocol` |
|---|---|
| `@mariozechner/pi-agent-core@0.73.1` | 0 (všech 22 souborů) |
| `@mariozechner/pi-web-ui@0.73.1` | 0 (všech 371 souborů) |
| `@earendil-works/pi-agent-core@0.82.1` | 0 |
| `@mariozechner/pi-ai@0.73.1` | 1, nesouvisející - OAuth scope string `user:mcp_servers` |
| `damianofalcioni/pi-for-word` src | 0 |

Autor to zdůvodnil blogpostem "What if you don't need MCP?".

Důsledek: **self-contained varianta neznamená "MCP místo Cline", znamená "žádné MCP vůbec".**
Tooly by se registrovaly přímo jako Pi tool definitions v JS a současný MCP server by v tom
scénáři nebyl použit k ničemu.

## 6. CLI agenti v prohlížeči: nelze

Ověřeno u Claude Code, Cline, OpenAI Codex CLI, gemini-cli, opencode, aider, goose,
agent-zero. Ani jeden nelze spustit v taskpane. Není to opomenutí web buildu, ale
fundamentální závislost na `node:fs`, `node:child_process`, PTY a filesystem workspace.

`@anthropic-ai/claude-code` (v2.1.220) dnes už ani není Node balík: `main: null`,
`browser: null`, `dependencies: {}`, `bin: {"claude": "bin/claude.exe"}`. Je to **thin
resolver pro prebuilt nativní binárky** přes optionalDependencies
(`@anthropic-ai/claude-code-{linux,darwin,win32}-{x64,arm64}`), každá ~256-275 MB.

Jedna výjimka, která ale problém neřeší: **`@anthropic-ai/claude-agent-sdk` má skutečný
browser build** (`/browser` export). Je to však remote-transport klient, ne lokální agent, a
**explicitně odmítá všechny MCP servery kromě in-process SDK ones** - onword-mcp by k němu
nepřipojil.

## 7. ACP - Agent Client Protocol

`agentclientprotocol.com`, od Zed Industries. **LSP pro agenty**: standardizované JSON-RPC
rozhraní, kterým jedno UI řídí libovolného agenta.

- Claude Code přes adapter `zed-industries/claude-code-acp`
- gemini-cli má ACP nativně
- Codex, opencode, goose - adaptéry existují
- Poskytuje streaming zpráv, **tool call approval gating**, session management
- V roce 2026 tento prostor podle rešerše jednoznačně vyhrál

Topologie pro Word:

```text
Word taskpane (ACP klient, vlastní UI)
      | wss://localhost
lokální helper proces
      | ACP (JSON-RPC over stdio)
libovolný agent: claude-code / gemini-cli / codex / opencode
      | MCP
onword-mcp  --COM nebo bridge-->  Word
```

Add-in je pak agent-agnostický a uživatel si vybere, čím to řídí.

Střízlivá poznámka: v tomto obrázku je onword-mcp stále potřeba, ale agent běží lokálně a na
Word by dosáhl i přes COM. ACP UI je tedy hlavně **pohodlnější vstupní bod** (chat ve Wordu
místo přepínání do terminálu), ne nová schopnost.

Upozornění: npm balík `@zed-industries/agent-client-protocol` je podle rešerše
stale/přejmenovaný - před implementací dohledat aktuální.

### Alternativa: PTY terminál v taskpane

Standardní stack: `@xterm/xterm` 6.0.0 + `@xterm/addon-fit` 0.11.0 + `node-pty` 1.1.0 +
`ws` 8.21.1. Hotové projekty: ttyd, gotty, wetty, code-server.

Nevýhody proti ACP:

- `@xterm/addon-attach` (0.12.0) **neřeší resize**, wire formát si každý projekt vymýšlí sám.
- `node-pty` je nativní modul, na Windows jen ConPTY (Win10 1809+/build 18309), potřebuje
  prebuilds nebo C++/Python toolchain.
- Terminál v úzkém taskpane je špatné UX a nelze z něj strukturovaně čerpat tool calls.

Doporučení rešerše: **ACP, ne PTY terminál.**

## 8. Zamítnuté alternativy

| Přístup | Verdikt |
|---|---|
| Microsoft Graph API | Nelze chirurgicky editovat. Pouze download/upload celého souboru. |
| Office Scripts | **Pouze Excel**, pro Word neexistuje. |
| VBA + lokální HTTP listener | Funguje, ale makra, bezpečnostní nastavení, křehké. |
| VSTO / COM add-in (C#) hostující MCP endpoint | Technicky možné (běží ve procesu Wordu), ale ClickOnce deployment, signing, admin práva. Žádný známý MCP-in-VSTO projekt. |
| python-docx / OOXML | Pouze zavřené soubory. Neřeší use case. |

Celkový verdikt na state of the art: **současný pywin32 COM přístup je stále jediný
mechanismus, který dává externímu samostatnému MCP serveru chirurgický, živý a
co-authoring-safe přístup k již otevřenému dokumentu.** Nic ho nevytlačilo. Office.js je
lepší editační plocha, ale za cenu toho, že musí běžet uvnitř hosta.

## 9. Varianty A / B / D

### A - Bridge (MCP)

Současný onword-mcp plus `wss://localhost` plus add-in jako executor.

- Řídí: Cline / Claude Desktop / Claude Code
- Agent vidí Word **i celý zbytek světa** (git, filesystem, další MCP servery)
- Cena: lokální proces, dev cert, řešení LNA (lokálním hostingem taskpane)
- Blueprint: `hewliyang/office-agents` `packages/bridge`

### B - Self-contained (Pi nebo vlastní agent loop v taskpane)

- Řídí: chat panel uvnitř Wordu
- GitHub Pages, cert-free, server-free, BYO API key
- **Žádný lokální proces, tedy ani MCP, ani CLI agent**
- Agent vidí **jen dokument** - žádný git, žádný filesystem, žádné jiné MCP
- Funguje na Windows, macOS, webu, iPadu
- Blueprint: `damianofalcioni/pi-for-word`

### D - ACP UI

- Řídí: chat panel ve Wordu, který ovládá libovolného lokálního CLI agenta
- Nejlepší UX z A
- **Největší kus práce**, a Word tooly do agenta stejně dorazí přes MCP, takže varianta A
  musí existovat první

### Srovnání

| | A: Bridge | B: Self-contained | D: ACP UI |
|---|---|---|---|
| MCP | ano | **ne, nikdy** | ano (přes A) |
| Instalace u klienta | proces + wss + cert | **jen manifest → GH Pages** | proces + adapter |
| Kdo platí LLM | existující předplacený klient | BYO API key v taskpane | existující klient |
| Kontext agenta | plný (git, fs, další MCP) | **izolovaný, jen Word** | plný |
| Funguje na Macu/webu | ano | ano | ano |
| Pracnost | střední | střední | vysoká |

Klíčový rozdíl, který stojí za nejvíc pozornosti: ve variantě B se ztrácí to, že agent vidí i
zbytek světa. Dnes lze zadat "vezmi data z toho CSV a přepiš podle nich tabulku v kapitole 3"
nebo "porovnej to s tím, co je v gitu". Pi v taskpane vidí jen ten dokument. To je funkčně
výrazně méně, i když UX je pohodlnější.

## 10. Doporučený postup

Podstatné zjištění: **`ops.js` - Office.js implementace současných operací - je společný
základ pro všechny tři varianty.** Jednou ho volá WS dispatcher (A), podruhé Pi tool registry
(B), potřetí totéž co A (D). To je ta skutečná práce a nezávisí na volbě varianty.

1. **Tento dokument** - zafixovat zjištění, aby se neztratila.
2. **`addin/ops.js` + debug taskpane** - operace v Office.js, ověřitelné klikáním v živém
   Wordu, bez agenta a bez bridge. Nejrizikovější technická část (`Word.Page`, `listItem`,
   `changeTrackingMode`, přepočet cm↔body) a společná všem variantám.
3. **Rozhodnout A / B / D** podle výsledků kroku 2, ne podle spekulace.
4. Implementovat zvolenou variantu. COM backend zachovat jako fallback pro Windows.

Doporučená struktura, pokud se půjde do varianty A:

```text
src/onword_mcp/
  backends/__init__.py      # výběr/autodetekce backendu
  backends/com.py           # = dnešní word.py, beze změny
  backends/bridge.py        # WSS server, hello/invoke/response, timeouty
addin/
  manifest.xml              # dev (localhost)
  manifest.gh-pages.xml     # prod (GitHub Pages)
  taskpane.html             # bez Reactu a bez webpacku, Office.js z CDN
  taskpane.js               # WS klient + dispatcher
  ops.js                    # Office.js implementace operací
```

Přepínač `--backend com|bridge` s autodetekcí: je-li připojený taskpane, použít ho, jinak COM.

## 11. Body k doověření před implementací

- **Přesné JSON tvary zpráv** bridge protokolu (`hello` / `invoke` / `response`, korelace id,
  error envelope, timeouty, heartbeat, reconnect) - v této rešerši jsou zachyceny jen
  konstanty a směr toku. Před portem do Pythonu znovu vytáhnout z
  `hewliyang/office-agents` `packages/bridge/src/protocol.ts`.
- **Chování LNA v WebView2 Office hosta** - je Office WebView2 verzí, na kterou se LNA
  vztahuje? Ověřit empiricky, ne z dokumentace.
- **Dev cert na firemním Windows stroji** - `office-addin-dev-certs` instaluje cert do user
  store; ověřit, zda to projde bez admin práv a bez konfliktu s firemní politikou (analogie
  problému s AV blokací `uvx` trampolíny, viz README).
- **Aktuální jméno npm balíku pro ACP** (`@zed-industries/agent-client-protocol` je
  stale/přejmenovaný).
- **`Word.Page` na Word na webu** - potvrdit graceful degradaci a rozhodnout, co vrátit
  místo čísel stránek.
- **Sideload bez admin práv** - shared folder katalog vs. registry vs. centralizovaný deploy
  v M365 admin center.

## Zdroje

- Word JS API requirement sets:
  <https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets>
- `Word.Page` (WordApiDesktop 1.2):
  <https://learn.microsoft.com/en-us/javascript/api/word/word.page>
- Word JS API reference: <https://learn.microsoft.com/en-us/javascript/api/word>
- <https://github.com/hewliyang/office-agents>
- <https://github.com/damianofalcioni/pi-for-word> (branch `master`)
- <https://github.com/patniko/github-copilot-office>
- <https://github.com/lancedesk/ms-office-ai-helper>
- <https://github.com/dsbissett/office-addin-mcp>
- <https://github.com/GongRzhe/Office-Word-MCP-Server>
- <https://github.com/earendil-works/pi>
- ACP: <https://agentclientprotocol.com>
- <https://github.com/zed-industries/claude-code-acp>
