"""Teste rápido: verifica 3 CNPJs com filtro Atende SUS."""
import time
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout
from config import TARGET_URL, HEADLESS

TIMEOUT = 20_000
_JS_PEGAR_HREF = (
    "() => {"
    "  const links = [...document.querySelectorAll('a[href*=\"coUnidade=\"]')]"
    "    .filter(a => { const m = (a.getAttribute('href')||'').match(/coUnidade=(\\d+)/); return m && parseInt(m[1]) > 0; });"
    "  return links.length > 0 ? links[0].href : null;"
    "}"
)

def buscar(page, cnpj):
    page.goto(TARGET_URL, timeout=TIMEOUT)
    page.wait_for_load_state("networkidle", timeout=TIMEOUT)
    loc = page.locator("button:has-text('Sim')")
    for i in range(loc.count()):
        el = loc.nth(i)
        if el.is_visible() and el.inner_text().strip() == "Sim":
            el.click(); time.sleep(0.3); break
    for sel in ["input[ng-model*='cnpj']","input[ng-model*='Cnpj']","input[placeholder*='CNPJ']"]:
        loc2 = page.locator(sel)
        if loc2.count() > 0:
            loc2.first.fill(cnpj); time.sleep(0.3); loc2.first.press("Tab"); time.sleep(0.3); break
    page.locator("button:has-text('Pesquisar')").first.click()
    page.wait_for_timeout(800)
    page.wait_for_load_state("networkidle", timeout=TIMEOUT)
    page.wait_for_timeout(400)
    return page.evaluate(_JS_PEGAR_HREF) is not None

df = pd.read_excel("ESTABELECIMENTOS_CNES.xlsx", skiprows=4, dtype=str, nrows=3)
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    for _, row in df.iterrows():
        cnpj = str(row["CNPJ"]).replace(".","").replace("/","").replace("-","")
        nome = str(row.get("Nome Fantasia", row.get("Razao Social","")))[:45]
        try:
            resultado = buscar(page, cnpj)
            print(f"{'SIM' if resultado else 'NAO'} | {nome}")
        except Exception as e:
            print(f"ERRO | {nome} | {e}")
    browser.close()
