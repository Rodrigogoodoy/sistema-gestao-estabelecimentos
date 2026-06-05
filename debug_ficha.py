"""
Script rápido para inspecionar a estrutura da ficha do CNES.
Abre o site, faz a pesquisa, captura o primeiro href e inspeciona o HTML da ficha.
"""
from playwright.sync_api import sync_playwright
import time

URL = "https://cnes.datasus.gov.br/pages/estabelecimentos/consulta.jsp"

# JS: captura o primeiro href de qualquer <a> na tabela
JS_HREFS = """
() => {
  const rows = document.querySelectorAll('table tbody tr');
  const result = [];
  for (const row of rows) {
    for (const a of row.querySelectorAll('a')) {
      result.push({href: a.href, text: a.innerText.trim(), onclick: a.getAttribute('ng-click') || ''});
    }
    // also capture buttons
    for (const b of row.querySelectorAll('button')) {
      result.push({href: '', text: b.innerText.trim(), onclick: b.getAttribute('ng-click') || ''});
    }
    break; // só primeira linha
  }
  return result;
}
"""

# JS: captura todos os labels + o que vem depois (input, span, div)
JS_INSPECT = """
() => {
  const items = [];
  const labels = document.querySelectorAll('label');
  for (const label of labels) {
    const chave = (label.innerText || '').trim().replace(/:$/, '').trim();
    if (!chave || chave.startsWith('{{') || chave.length > 120) continue;
    const info = {label: chave, input: null, span: null, siblings: []};

    // Check input
    const forAttr = label.getAttribute('for');
    let inp = forAttr ? document.getElementById(forAttr) : null;
    if (!inp) {
      let el = label.nextElementSibling;
      for (let j = 0; j < 5 && el; j++) {
        if (['INPUT','TEXTAREA','SELECT'].includes(el.tagName)) { inp = el; break; }
        const sub = el.querySelector && el.querySelector('input,textarea,select');
        if (sub) { inp = sub; break; }
        el = el.nextElementSibling;
      }
    }
    if (inp) {
      if (inp.tagName === 'SELECT') {
        info.input = {tag: 'SELECT', value: inp.value, text: inp.selectedIndex >= 0 ? inp.options[inp.selectedIndex].text : ''};
      } else {
        info.input = {tag: inp.tagName, value: inp.value, type: inp.type};
      }
    }

    // Check sibling spans/divs
    let el = label.nextElementSibling;
    for (let j = 0; j < 5 && el; j++) {
      const txt = (el.innerText || '').trim();
      if (txt) info.siblings.push({tag: el.tagName, text: txt.substring(0, 60)});
      el = el.nextElementSibling;
    }

    items.push(info);
    if (items.length >= 30) break;
  }
  return items;
}
"""

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)  # visível para debug
    page = browser.new_page()

    print("Abrindo site CNES...")
    page.goto(URL, timeout=40_000)
    page.wait_for_load_state("networkidle", timeout=40_000)

    # Seleciona SP
    page.locator("select[ng-model='Estado']").select_option(value="35")
    time.sleep(0.8)
    page.wait_for_function(
        "(function(){ var s = document.querySelector(\"select[ng-model='Municipio']\"); return s && s.options.length > 1; })()",
        timeout=25_000
    )
    # Lista options do município para debug
    options = page.evaluate("""
    () => {
      const s = document.querySelector("select[ng-model='Municipio']");
      if (!s) return [];
      return [...s.options].slice(0, 20).map(o => ({value: o.value, text: o.text}));
    }
    """)
    print("Primeiras options de Município:", options)
    # Seleciona São José dos Campos pelo nome
    campo = page.locator("select[ng-model='Municipio']")
    opts = campo.locator("option").all()
    for opt in opts:
        txt = (opt.inner_text() or "").strip().upper()
        if "SAO JOSE DOS CAMPOS" in txt or "SÃO JOSÉ DOS CAMPOS" in txt or "JOSE DOS CAMPOS" in txt:
            val = opt.get_attribute("value") or ""
            campo.select_option(value=val)
            print(f"Selecionou: value={val} text={txt}")
            break
    time.sleep(0.3)
    # Clica Sim (Atende SUS)
    for el in page.locator("button").all():
        if (el.inner_text() or "").strip() == "Sim" and el.is_visible():
            el.click()
            break
    time.sleep(0.3)
    page.locator("button:has-text('Pesquisar')").first.click()
    page.wait_for_load_state("networkidle", timeout=25_000)
    page.wait_for_timeout(1000)

    print("\n=== LINKS / BOTÕES NA PRIMEIRA LINHA DA TABELA ===")
    hrefs = page.evaluate(JS_HREFS)
    for h in hrefs:
        print(f"  tag=a  href={h['href'][:120]}  text={h['text']}  ng-click={h['onclick']}")

    # Tenta clicar no primeiro link de ficha ou botão "+"
    rows = page.locator("table tbody tr")
    if rows.count() > 0:
        first_row = rows.first
        # Tenta clicar em qualquer botão com "+" ou link
        btn = first_row.locator("button")
        if btn.count() > 0:
            btn.first.click()
            print("\nClicou no primeiro botão da linha.")
            page.wait_for_timeout(2000)

            # Verifica se abriu modal
            modal = page.locator(".modal-content")
            if modal.count() > 0:
                print("=== MODAL ABERTO — inspecionando labels ===")
                items = page.evaluate("""
                () => {
                  const modal = document.querySelector('.modal-content');
                  if (!modal) return [];
                  const items = [];
                  for (const label of modal.querySelectorAll('label')) {
                    const chave = (label.innerText||'').trim().replace(/:$/,'').trim();
                    if (!chave || chave.startsWith('{{') || chave.length > 120) continue;
                    let inp = null;
                    const fa = label.getAttribute('for');
                    if (fa) inp = document.getElementById(fa);
                    if (!inp) {
                      let el = label.nextElementSibling;
                      while (el && !inp) {
                        if (['INPUT','TEXTAREA','SELECT'].includes(el.tagName)) inp = el;
                        else if (el.querySelector) inp = el.querySelector('input,textarea,select');
                        el = el.nextElementSibling;
                      }
                    }
                    if (!inp) { const p=label.parentElement; if(p) inp=p.querySelector('input,textarea,select'); }
                    let valor = '';
                    if (inp) {
                      if (inp.tagName === 'SELECT') {
                        valor = inp.selectedIndex >= 0 ? (inp.options[inp.selectedIndex].text||'').trim() : inp.value;
                      } else {
                        valor = (inp.value || inp.innerText || '').trim();
                      }
                    }
                    items.push({label: chave, valor: valor});
                  }
                  return items;
                }
                """)
                print(f"Total labels no modal: {len(items)}")
                for it in items:
                    print(f"  [{it['label']}] = '{it['valor'][:60]}'")

                # Também navega para a URL da ficha para ver estrutura
                href_ficha = hrefs[0]['href'] if hrefs else None
                if href_ficha:
                    print(f"\nNavegando para ficha: {href_ficha}")
                    page.goto(href_ficha, timeout=40_000)
                    page.wait_for_load_state("networkidle", timeout=40_000)
                    page.wait_for_timeout(2000)
                    ficha_items = page.evaluate("""
                    () => {
                      const items = [];
                      for (const label of document.querySelectorAll('label')) {
                        const chave = (label.innerText||'').trim().replace(/:$/,'').trim();
                        if (!chave || chave.startsWith('{{') || chave.length > 120) continue;
                        let inp = null;
                        const fa = label.getAttribute('for');
                        if (fa) inp = document.getElementById(fa);
                        if (!inp) {
                          let el = label.nextElementSibling;
                          while (el && !inp) {
                            if (['INPUT','TEXTAREA','SELECT'].includes(el.tagName)) inp = el;
                            else if (el.querySelector) inp = el.querySelector('input,textarea,select');
                            el = el.nextElementSibling;
                          }
                        }
                        if (!inp) { const p=label.parentElement; if(p) inp=p.querySelector('input,textarea,select'); }
                        let valor = ''; let elem_tipo = 'N/A';
                        if (inp) {
                          elem_tipo = inp.tagName + (inp.type ? ':'+inp.type : '');
                          if (inp.tagName === 'SELECT') {
                            valor = inp.selectedIndex >= 0 ? (inp.options[inp.selectedIndex].text||'').trim() : '';
                          } else {
                            valor = (inp.value || inp.innerText || '').trim();
                          }
                        } else {
                          // fallback: sibling span/div
                          let el = label.nextElementSibling;
                          for (let j=0; j<3 && el; j++) {
                            const txt = (el.innerText||'').trim();
                            if (txt && !txt.startsWith('{{')) { valor = txt; elem_tipo = 'SPAN:'+el.tagName; break; }
                            el = el.nextElementSibling;
                          }
                        }
                        items.push({label: chave, valor: valor, tipo: elem_tipo});
                      }
                      return items;
                    }
                    """)
                    print(f"Labels na FICHA PAGE ({page.url[:80]}):")
                    for it in ficha_items[:40]:
                        print(f"  [{it['label']}] ({it['tipo']}) = '{it['valor'][:60]}'")

            else:
                print("=== SEM MODAL — verificando página atual ===")
                print("URL atual:", page.url)
                page.wait_for_timeout(2000)
                items = page.evaluate(JS_INSPECT)
                print(f"Labels encontrados na página: {len(items)}")
                for it in items[:20]:
                    print(f"  [{it['label']}] input={it.get('input')} siblings={it['siblings'][:3]}")

    browser.close()
