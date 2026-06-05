"""
Diagnóstico: abre a página do CNES e imprime as options de cada select visível.
Rode uma vez para descobrir os valores exatos dos dropdowns.
"""
from playwright.sync_api import sync_playwright

URL = "https://cnes.datasus.gov.br/pages/estabelecimentos/consulta.jsp"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    page.goto(URL)
    page.wait_for_load_state("networkidle", timeout=25_000)
    page.wait_for_timeout(2000)

    selects = page.locator("select")
    print(f"\n{selects.count()} selects encontrados na página\n")

    for i in range(selects.count()):
        s = selects.nth(i)
        ng = s.get_attribute("ng-model") or "(sem ng-model)"
        sid = s.get_attribute("id") or "(sem id)"
        print(f"--- Select {i+1}: ng-model={ng}  id={sid} ---")
        options = s.locator("option").all()
        for opt in options[:20]:
            val = opt.get_attribute("value") or ""
            txt = opt.inner_text().strip()
            print(f"  value={val!r:30s}  label={txt!r}")
        if len(options) > 20:
            print(f"  ... (+{len(options)-20} mais)")
        print()

    browser.close()
