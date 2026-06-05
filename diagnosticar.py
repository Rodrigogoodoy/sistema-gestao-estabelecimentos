"""
Script de diagnóstico — busca por nome para achar um CNPJ real no CNES,
depois testa o fluxo completo de coleta da ficha.
"""
from playwright.sync_api import sync_playwright
from config import TARGET_URL

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, slow_mo=300)
    page = browser.new_page()

    print("1. Abrindo o site e buscando por 'EINSTEIN' para achar um CNPJ real...")
    page.goto(TARGET_URL)
    page.wait_for_load_state("networkidle")

    campo = page.locator("input[placeholder*='CNPJ']").first
    campo.fill("EINSTEIN")
    page.locator("button:has-text('Pesquisar')").first.click()

    page.wait_for_timeout(800)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)
    page.screenshot(path="debug_busca_nome.png")

    # Pega CNPJs dos resultados
    dados_resultados = page.evaluate("""
        () => {
            const linhas = [...document.querySelectorAll('table tbody tr')];
            return linhas.slice(0, 5).map(tr => {
                const cells = [...tr.querySelectorAll('td')];
                return cells.map(c => c.innerText.trim()).join(' | ');
            });
        }
    """)
    print(f"\n2. Primeiros resultados da tabela:")
    for linha in dados_resultados:
        print(f"   {linha}")

    # Pega o primeiro link de ficha com coUnidade válido
    href = page.evaluate("""
        () => {
            const links = [...document.querySelectorAll('a[href*="coUnidade="]')]
                .filter(a => {
                    const m = (a.getAttribute('href')||'').match(/coUnidade=(\\d+)/);
                    return m && parseInt(m[1]) > 0;
                });
            return links.length > 0 ? links[0].href : null;
        }
    """)
    print(f"\n3. Primeiro link de ficha encontrado: {href}")

    if href:
        print("\n4. Acessando a ficha...")
        page.goto(href)
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)
        page.screenshot(path="debug_ficha.png")

        # Coleta dados da tabela
        dados = {}
        linhas = page.locator("table tr")
        for i in range(min(linhas.count(), 50)):
            celulas = linhas.nth(i).locator("td, th")
            if celulas.count() >= 2:
                chave = celulas.nth(0).inner_text().strip().rstrip(":")
                valor = celulas.nth(1).inner_text().strip()
                if chave and len(chave) < 100 and not chave.startswith("{{"):
                    dados[chave] = valor

        print(f"\n5. Dados coletados da ficha ({len(dados)} campos):")
        for k, v in list(dados.items())[:20]:
            print(f"   {k}: {v[:80]}")

        body = page.evaluate("() => document.body.innerText")
        print(f"\n6. Texto da ficha (primeiros 1000 chars):")
        print("-" * 50)
        print(body[:1000])
    else:
        print("Nenhum resultado encontrado para 'EINSTEIN'.")
        body = page.evaluate("() => document.body.innerText")
        print(body[:500])

    print("\nFechando em 10 segundos...")
    page.wait_for_timeout(10000)
    browser.close()
    print("Concluído.")
