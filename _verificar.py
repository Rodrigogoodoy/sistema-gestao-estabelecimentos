"""
Verifica no DataSUS quantos estabelecimentos cada município tem
e compara com o que foi coletado.
"""
import time, unicodedata
from playwright.sync_api import sync_playwright

MUNICIPIOS = [
    {"uf": "PR", "municipio": "Lapa",                "cod_ibge": "4113205", "coletado": 34},
    {"uf": "PR", "municipio": "Contenda",             "cod_ibge": "4106209", "coletado": 17},
    {"uf": "PR", "municipio": "Campo Largo",          "cod_ibge": "4104204", "coletado": 67},
    {"uf": "PR", "municipio": "Rio Negro",            "cod_ibge": "4122305", "coletado": 34},
    {"uf": "PR", "municipio": "São Mateus do Sul",    "cod_ibge": "4125605", "coletado": 45},
    {"uf": "PR", "municipio": "Três Barras do Paraná","cod_ibge": "4127858", "coletado": 20},
    {"uf": "SC", "municipio": "Rio Negrinho",         "cod_ibge": "4215000", "coletado": 52},
    {"uf": "SC", "municipio": "São Bento do Sul",     "cod_ibge": "4215802", "coletado": 89},
    {"uf": "SC", "municipio": "Fraiburgo",            "cod_ibge": "4205506", "coletado": 75},
    {"uf": "SC", "municipio": "Videira",              "cod_ibge": "4219309", "coletado": 106},
    {"uf": "SC", "municipio": "Caçador",              "cod_ibge": "4203006", "coletado": 86},
    {"uf": "SC", "municipio": "Monte Carlo",          "cod_ibge": "4211058", "coletado": 19},
    {"uf": "SC", "municipio": "Tangará",              "cod_ibge": "4217907", "coletado": 15},
    {"uf": "SC", "municipio": "Curitibanos",          "cod_ibge": "4204806", "coletado": 70},
]

_UF_CODIGOS = {
    "PR": "41", "SC": "42",
}

URL = "https://cnes.datasus.gov.br/pages/estabelecimentos/consulta.jsp"
TIMEOUT = 25_000

def sem_acento(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").upper()

def verificar(page, item):
    uf, mun, cod, coletado = item["uf"], item["municipio"], item["cod_ibge"], item["coletado"]
    try:
        page.goto(URL, timeout=40_000)
        page.wait_for_load_state("networkidle", timeout=40_000)

        # Seleciona UF
        campo_uf = page.locator("select[ng-model='Estado']")
        campo_uf.wait_for(state="visible", timeout=TIMEOUT)
        campo_uf.select_option(value=_UF_CODIGOS[uf])
        time.sleep(1)

        # Aguarda municípios carregarem
        page.wait_for_function(
            "(function(){ var s = document.querySelector(\"select[ng-model='Municipio']\"); return s && s.options.length > 1; })()",
            timeout=TIMEOUT,
        )

        # Seleciona município por nome
        campo_mun = page.locator("select[ng-model='Municipio']")
        nome_norm = sem_acento(mun)
        options = campo_mun.locator("option").all()
        selecionado = False
        for opt in options:
            val = (opt.get_attribute("value") or "").strip()
            txt = sem_acento((opt.inner_text() or "").strip())
            if nome_norm in txt or cod[:6] == val:
                campo_mun.select_option(value=val)
                selecionado = True
                break

        if not selecionado:
            return {"status": "MUNICIPIO_NAO_ENCONTRADO", "site": None, "coletado": coletado}

        time.sleep(0.3)

        # Clica "Atende SUS: Sim"
        for sel in ["button:has-text('Sim')", "a:has-text('Sim')", "label:has-text('Sim')"]:
            loc = page.locator(sel)
            for i in range(loc.count()):
                el = loc.nth(i)
                if (el.inner_text() or "").strip() == "Sim" and el.is_visible():
                    el.click()
                    time.sleep(0.3)
                    break

        # Pesquisa
        page.locator("button:has-text('Pesquisar')").first.click()
        time.sleep(1)
        page.wait_for_load_state("networkidle", timeout=TIMEOUT)
        time.sleep(1)

        # Tenta várias formas de ler o total do site
        total_site = page.evaluate(
            "() => {"
            "  const t = document.body.innerText;"
            "  const padroes = ["
            "    /Total[:\\s]+(\\d+)/i,"
            "    /de\\s+(\\d+)\\s+registro/i,"
            "    /de\\s+(\\d+)\\s+resultado/i,"
            "    /(\\d+)\\s+registro/i,"
            "    /(\\d+)\\s+resultado/i,"
            "    /Exibindo.*de\\s+(\\d+)/i,"
            "    /\\d+\\s*-\\s*\\d+\\s+de\\s+(\\d+)/i,"
            "    /total.*?:\\s*(\\d+)/i"
            "  ];"
            "  for (const p of padroes) { const m = t.match(p); if (m) return parseInt(m[1]); }"
            # Tenta pegar do paginador: último número de página × 30
            "  const pags = [...document.querySelectorAll('ul.pagination li a')]"
            "    .map(a => parseInt(a.innerText)).filter(n => !isNaN(n));"
            "  if (pags.length) return Math.max(...pags);"
            "  return null;"
            "}"
        )

        linhas = page.locator("table tbody tr").count()
        return {"status": "OK", "site": total_site, "linhas_pag1": linhas, "coletado": coletado}

    except Exception as e:
        return {"status": f"ERRO: {e}", "site": None, "coletado": coletado}


print(f"\n{'='*70}")
print(f"  VERIFICAÇÃO DataSUS — {len(MUNICIPIOS)} municípios")
print(f"{'='*70}")
print(f"{'Município':<28} {'UF':<4} {'Site':>6} {'Coletado':>9} {'Status'}")
print(f"{'-'*70}")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()

    for item in MUNICIPIOS:
        res = verificar(page, item)
        mun_label = item["municipio"][:27]
        site_str  = str(res["site"]) if res["site"] else "?"
        col_str   = str(res["coletado"])

        if res["site"] and res["coletado"] < res["site"] * 0.9:
            flag = "⚠ FALTANDO"
        elif res["status"] != "OK":
            flag = res["status"]
        else:
            flag = "OK"

        print(f"{mun_label:<28} {item['uf']:<4} {site_str:>6} {col_str:>9}   {flag}")

    browser.close()

print(f"{'='*70}\n")
