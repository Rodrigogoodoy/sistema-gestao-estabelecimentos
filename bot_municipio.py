import time
import unicodedata
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeout


def _sem_acento(texto: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


TIMEOUT = 25_000
TIMEOUT_FICHA = 35_000

# Compound key: municipality_code(6 digits) + CNES(7 digits) — extracted from href
FICHA_URL_BASE = (
    "https://cnes.datasus.gov.br/pages/estabelecimentos/ficha/index.jsp?coUnidade={}"
)

_UF_CODIGOS = {
    "AC": "12", "AL": "27", "AP": "16", "AM": "13",
    "BA": "29", "CE": "23", "DF": "53", "ES": "32",
    "GO": "52", "MA": "21", "MT": "51", "MS": "50",
    "MG": "31", "PA": "15", "PB": "25", "PR": "41",
    "PE": "26", "PI": "22", "RJ": "33", "RN": "24",
    "RS": "43", "RO": "11", "RR": "14", "SC": "42",
    "SP": "35", "SE": "28", "TO": "17",
}

# Coleta todos os campos da ficha completa (INPUT:text).
# - SELECT: usa .text (rótulo visível), corrige "Natureza Jurídica(Grupo)"
# - Campos duplicados: mantém a PRIMEIRA ocorrência (a certa)
_JS_COLETAR_FICHA = (
    "() => {"
    "  const result = {};"
    # Campos que devem ser capturados mesmo vazios
    "  const SEMPRE_CAPTURAR = ['Data Desativação', 'Motivo Desativação'];"
    # Campos que usam ÚLTIMA ocorrência (a da seção detalhada, mais fiel ao DataSUS)
    "  const ULTIMA_OCORRENCIA = ['Natureza Jurídica(Grupo)'];"
    "  const labels = [...document.querySelectorAll('label')];"
    "  for (const label of labels) {"
    "    const chave = (label.innerText || '').trim().replace(/:$/, '').trim();"
    "    if (!chave || chave.startsWith('{{') || chave.length > 120) continue;"
    # Pula campos já vistos, EXCETO os que devem usar a última ocorrência
    "    if (result[chave] !== undefined && !ULTIMA_OCORRENCIA.includes(chave)) continue;"
    "    let input = null;"
    "    const forAttr = label.getAttribute('for');"
    "    if (forAttr) input = document.getElementById(forAttr);"
    "    if (!input) {"
    "      let el = label.nextElementSibling;"
    "      while (el && !input) {"
    "        if (['INPUT','TEXTAREA','SELECT'].includes(el.tagName)) input = el;"
    "        else if (el.querySelector) input = el.querySelector('input,textarea,select');"
    "        el = el.nextElementSibling;"
    "      }"
    "    }"
    "    if (!input) { const p = label.parentElement; if (p) input = p.querySelector('input,textarea,select'); }"
    "    if (!input) { const p = label.parentElement; if (p && p.parentElement) input = p.parentElement.querySelector('input,textarea,select'); }"
    "    if (!input) continue;"
    "    const tipo = (input.type || '').toLowerCase();"
    "    if (tipo === 'checkbox' || tipo === 'radio') continue;"
    "    let valor;"
    "    if (input.tagName === 'SELECT') {"
    # Lê o texto visível da opção selecionada; se Angular não populou ainda, retorna ''
    "      const opt = input.options[input.selectedIndex];"
    "      valor = (opt && opt.text && !opt.text.startsWith('{{')) ? opt.text.trim() : '';"
    "    } else {"
    "      valor = (input.value || input.innerText || '').trim().replace(/\\n+/g, ' ');"
    "    }"
    # Campos obrigatórios: salva mesmo vazio
    "    if (SEMPRE_CAPTURAR.includes(chave)) {"
    "      if (result[chave] === undefined) result[chave] = valor || '';"
    "    } else if (valor && !valor.startsWith('{{') && valor !== 'on' && valor !== 'off') {"
    "      result[chave] = valor;"
    "    }"
    "  }"
    # Garante que Data/Motivo Desativação sempre existem
    "  if (!('Data Desativação' in result)) result['Data Desativação'] = '';"
    "  if (!('Motivo Desativação' in result)) result['Motivo Desativação'] = '';"
    "  return result;"
    "}"
)

_JS_CONTAR_TOTAL = (
    "() => {"
    "  const text = document.body.innerText;"
    "  const m = text.match(/Total[:\\s]+(\\d+)/i)"
    "    || text.match(/(\\d+)\\s+(?:registro|estabelecimento|resultado)/i)"
    "    || text.match(/de\\s+(\\d+)/i);"
    "  return m ? parseInt(m[1]) : null;"
    "}"
)

# Extrai o coUnidade composto (13 dígitos) do href de cada linha.
# Formato: {cod_municipio_6dig}{cnes_7dig} — extraído de a[href*="coUnidade="]
_JS_COLETAR_CNES_PAGINA = (
    "() => {"
    "  const nums = [];"
    "  const rows = document.querySelectorAll('table tbody tr');"
    "  for (const row of rows) {"
    "    let found = false;"
    "    for (const a of row.querySelectorAll('a[href]')) {"
    "      const m = (a.href||'').match(/coUnidade=(\\d+)/i);"
    "      if (m) { nums.push(m[1]); found=true; break; }"
    "    }"
    "    if (found) continue;"
    # Fallback: qualquer 7 dígitos em célula (CNES puro)
    "    for (const td of row.querySelectorAll('td')) {"
    "      const t = (td.innerText||'').trim();"
    "      if (/^\\d{7}$/.test(t)) { nums.push(t); break; }"
    "    }"
    "  }"
    "  return nums;"
    "}"
)


def _clicar_atende_sus_sim(page: Page):
    candidatos = [
        "button:has-text('Sim')",
        "a:has-text('Sim')",
        "label:has-text('Sim')",
        "span:has-text('Sim')",
    ]
    for sel in candidatos:
        loc = page.locator(sel)
        for i in range(loc.count()):
            el = loc.nth(i)
            txt = (el.inner_text() or "").strip()
            if txt == "Sim" and el.is_visible():
                el.click()
                time.sleep(0.3)
                print("  Filtro 'Atende SUS: Sim' ativado")
                return
    print("  Aviso: botão 'Atende SUS: Sim' não encontrado — pesquisando com 'Todos'")


def _selecionar_uf(page: Page, uf: str):
    codigo = _UF_CODIGOS.get(uf.upper())
    if not codigo:
        raise Exception(f"UF desconhecida: {uf}")
    campo = page.locator("select[ng-model='Estado']")
    campo.wait_for(state="visible", timeout=TIMEOUT)
    campo.select_option(value=codigo)


def _aguardar_municipio_populado(page: Page):
    page.wait_for_function(
        "(function(){ var s = document.querySelector(\"select[ng-model='Municipio']\"); return s && s.options.length > 1; })()",
        timeout=TIMEOUT,
    )


def _selecionar_municipio(page: Page, nome: str, cod_ibge: str):
    """
    O site usa código de 6 dígitos para município (ex: 354990),
    mas cod_ibge tem 7 dígitos (ex: 3549904). Por isso tenta por nome primeiro.
    """
    campo = page.locator("select[ng-model='Municipio']")
    campo.wait_for(state="visible", timeout=TIMEOUT)

    nome_normalizado = _sem_acento(nome.upper())
    options = campo.locator("option").all()
    for opt in options:
        val = (opt.get_attribute("value") or "").strip()
        txt = _sem_acento((opt.inner_text() or "").upper().strip())
        if nome_normalizado in txt or (len(cod_ibge) == 7 and cod_ibge[:6] == val):
            campo.select_option(value=val)
            return

    # último recurso: tenta o próprio cod_ibge
    try:
        campo.select_option(value=cod_ibge)
        return
    except Exception:
        pass

    sample = [(o.get_attribute("value"), o.inner_text().strip()) for o in options[:10]]
    raise Exception(
        f"Município não encontrado. nome={nome} cod_ibge={cod_ibge}. "
        f"Primeiras options: {sample}"
    )


def _set_registros_por_pagina(page: Page, qtd: str = "30"):
    campo = page.locator("select[ng-model='registrosPorPagina']")
    if campo.count() > 0:
        try:
            campo.first.select_option(value=qtd)
        except Exception:
            pass


def _clicar_pesquisar(page: Page):
    botao = page.locator("button:has-text('Pesquisar')").first
    botao.wait_for(state="visible", timeout=TIMEOUT)
    botao.click()


def _aguardar_resultados(page: Page):
    page.wait_for_timeout(1000)
    page.wait_for_load_state("networkidle", timeout=TIMEOUT)
    page.wait_for_timeout(600)


def _tem_resultados(page: Page) -> bool:
    return page.locator("table tbody tr").count() > 0


def _ir_proxima_pagina(page: Page) -> bool:
    candidatos = [
        "ul.pagination li:not(.disabled):last-child a",
        "ul.pagination li.next:not(.disabled) a",
        "[uib-pagination] li:not(.disabled):last-child a",
        "button:has-text('Próximo')",
        "a:has-text('Próximo')",
        "a:has-text('›')",
        "a:has-text('»')",
        "a[ng-click*='next']",
        "li[ng-class*='next']:not([class*='disabled']) a",
    ]
    for sel in candidatos:
        loc = page.locator(sel)
        if loc.count() > 0:
            el = loc.first
            if el.is_visible() and el.is_enabled():
                el.click()
                _aguardar_resultados(page)
                return True
    return False


def _goto_com_retry(page: Page, url: str, tentativas: int = 3):
    for i in range(tentativas):
        try:
            page.goto(url, timeout=40_000)
            page.wait_for_load_state("networkidle", timeout=40_000)
            return
        except Exception as e:
            if i < tentativas - 1:
                print(f"  Timeout ao carregar página, tentativa {i+2}/{tentativas}...")
                time.sleep(3)
            else:
                raise e


def _aguardar_ficha_carregada(page: Page):
    """Aguarda o AngularJS popular inputs e SELECTs da ficha."""
    try:
        page.wait_for_function(
            "(function(){"
            "  var inp = document.querySelector('input[ng-model]');"
            "  if (!inp || !inp.value || inp.value.startsWith('{{')) return false;"
            # Aguarda também os SELECTs terem opções carregadas
            "  var selects = document.querySelectorAll('select[ng-model]');"
            "  for (var s of selects) { if (s.options.length <= 1) return false; }"
            "  return true;"
            "})()",
            timeout=15_000,
        )
    except Exception:
        pass
    page.wait_for_timeout(600)


def _executar_pesquisa(page: Page, uf: str, municipio: str, cod_ibge: str, url_base: str):
    _goto_com_retry(page, url_base)
    _selecionar_uf(page, uf)
    time.sleep(0.8)
    _aguardar_municipio_populado(page)
    _selecionar_municipio(page, municipio, cod_ibge)
    time.sleep(0.3)
    _set_registros_por_pagina(page, "30")
    time.sleep(0.3)
    _clicar_atende_sus_sim(page)
    time.sleep(0.3)
    _clicar_pesquisar(page)
    _aguardar_resultados(page)


def processar_municipio(
    page: Page,
    uf: str,
    municipio: str,
    cod_ibge: str,
    url_base: str,
    delay_entre_fichas: float = 0.3,
) -> list:
    """
    Passo 1 — percorre resultados (paginação) e coleta os coUnidade compostos.
    Passo 2 — navega até ficha completa de cada um:
               CNES, CNPJ, Natureza Jurídica(Grupo) correto,
               Data Desativação, Motivo Desativação, e mais.
    """
    # ── Passo 1: coletar todos os coUnidade ───────────────────────────────────
    _executar_pesquisa(page, uf, municipio, cod_ibge, url_base)

    if not _tem_resultados(page):
        print(f"  Nenhum estabelecimento encontrado para {municipio}/{uf}")
        return []

    total_estimado = page.evaluate(_JS_CONTAR_TOTAL)
    if total_estimado:
        print(f"  Total indicado pelo site: {total_estimado}")

    todos_ids: list = []
    vistos_ids: set = set()
    pagina = 1

    while True:
        nums = page.evaluate(_JS_COLETAR_CNES_PAGINA)
        novos = [c for c in nums if c and c not in vistos_ids]
        todos_ids.extend(novos)
        vistos_ids.update(novos)
        print(f"  Página {pagina}: {len(novos)} IDs (total: {len(todos_ids)})")

        if nums and not novos:
            print("  Sem IDs novos, encerrando paginação.")
            break
        if not _ir_proxima_pagina(page):
            break
        pagina += 1
        if not _tem_resultados(page):
            break

    if not todos_ids:
        print(f"  Nenhum ID coletado para {municipio}/{uf}")
        return []

    print(f"  Coletando fichas completas: {len(todos_ids)} estabelecimentos...")

    # ── Passo 2: visitar cada ficha ───────────────────────────────────────────
    todos: list = []
    vistos_chave: set = set()

    for i, uid in enumerate(todos_ids):
        url_ficha = FICHA_URL_BASE.format(uid)
        try:
            _goto_com_retry(page, url_ficha, tentativas=2)
            _aguardar_ficha_carregada(page)

            dados = page.evaluate(_JS_COLETAR_FICHA)
            dados = {k: (v if not str(v).startswith("{{") else "") for k, v in dados.items()}
            dados["UF"] = uf
            dados["Município"] = municipio
            dados["Cod IBGE"] = cod_ibge

            # CNPJ = apenas CNPJ Próprio do estabelecimento
            # CNPJ Mantenedora vai em coluna separada e NÃO substitui o CNPJ próprio
            def _limpar_cnpj(v):
                v = str(v or "").strip()
                return "" if v in ("---", "-", "nan", "None") else v

            cnpj = _limpar_cnpj(dados.get("CNPJ", ""))
            if not cnpj:
                cnpj = _limpar_cnpj(dados.get("CNPJ Próprio", ""))
            dados["CNPJ"] = cnpj

            # Extrai CNES puro dos últimos 7 dígitos do ID composto se não encontrado
            if "CNES" not in dados and len(uid) >= 7:
                dados["CNES"] = uid[-7:]

            chave = dados.get("CNES") or dados.get("Nome") or dados.get("Nome Fantasia") or ""
            if chave and chave not in vistos_chave:
                vistos_chave.add(chave)
                todos.append(dados)
                nome = dados.get("Nome") or dados.get("Nome Fantasia") or f"CNES {uid[-7:]}"
                print(f"    [{i + 1}/{len(todos_ids)}] OK: {nome[:60]}")

            time.sleep(delay_entre_fichas)

        except PlaywrightTimeout:
            print(f"    [{i + 1}/{len(todos_ids)}] Timeout: ID {uid}")
        except Exception as e:
            print(f"    [{i + 1}/{len(todos_ids)}] Erro ID {uid}: {e}")

    return todos
