import time
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeout
from config import TARGET_URL, DELAY_BETWEEN_ENTRIES, DELAY_ON_ERROR, MAX_RETRIES
from logger import log_erro, log_sucesso

TIMEOUT_PADRAO = 20_000

# JS que retorna a URL da primeira ficha encontrada (ou null)
_JS_PEGAR_HREF_FICHA = (
    "() => {"
    "  const links = [...document.querySelectorAll('a[href*=\"coUnidade=\"]')]"
    "    .filter(a => { const m = (a.getAttribute('href')||'').match(/coUnidade=(\\d+)/); return m && parseInt(m[1]) > 0; });"
    "  return links.length > 0 ? links[0].href : null;"
    "}"
)


def _aguardar_pagina(page: Page):
    page.wait_for_load_state("networkidle", timeout=TIMEOUT_PADRAO)


def _preencher_e_pesquisar(page: Page, cnpj: str):
    """Preenche o campo CNPJ e dispara a pesquisa."""
    seletores = [
        "input[ng-model*='cnpj']",
        "input[ng-model*='Cnpj']",
        "input[ng-model*='CNPJ']",
        "input[ng-model*='nuCnpj']",
        "input[placeholder*='CNPJ']",
        "input[placeholder*='cnpj']",
        "input[name*='cnpj']",
        "input[name*='CNPJ']",
        "#cnpj",
    ]

    campo = None
    for sel in seletores:
        loc = page.locator(sel)
        if loc.count() > 0:
            campo = loc.first
            break

    if campo is None:
        inputs = page.locator("input:visible")
        info = [
            f"ng-model={inputs.nth(i).get_attribute('ng-model')} "
            f"placeholder={inputs.nth(i).get_attribute('placeholder')} "
            f"id={inputs.nth(i).get_attribute('id')}"
            for i in range(min(inputs.count(), 10))
        ]
        raise Exception(f"Campo CNPJ não encontrado. Inputs visíveis: {info}")

    # fill() já limpa o campo antes de digitar
    campo.fill(cnpj)
    time.sleep(0.3)

    # Se o campo rejeitou sem formatação, tenta com pontos/barras
    valor_atual = campo.input_value()
    if cnpj not in valor_atual.replace(".", "").replace("/", "").replace("-", ""):
        formatado = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:14]}"
        campo.fill(formatado)
        time.sleep(0.3)

    campo.press("Tab")      # dispara validação AngularJS
    time.sleep(0.4)         # deixa Angular processar o ng-model

    botao = page.locator("button:has-text('Pesquisar')").first
    botao.wait_for(state="visible", timeout=TIMEOUT_PADRAO)
    botao.click()


def _aguardar_e_verificar_resultado(page: Page) -> bool:
    """
    Aguarda o Angular fazer a requisição ao servidor e renderizar os resultados.
    Retorna True se encontrou dados, False se não encontrou.
    """
    page.wait_for_timeout(600)
    page.wait_for_load_state("networkidle", timeout=TIMEOUT_PADRAO)
    page.wait_for_timeout(400)

    href = page.evaluate(_JS_PEGAR_HREF_FICHA)
    return href is not None


def _ir_para_ficha(page: Page):
    """Navega diretamente para a URL da ficha (evita problemas de visibilidade do elemento)."""
    href = page.evaluate(_JS_PEGAR_HREF_FICHA)
    if not href:
        raise Exception("Nenhum link de ficha com coUnidade válido encontrado")
    page.goto(href)


_JS_COLETAR_FICHA = """
() => {
    const result = {};

    // Pares label -> input de texto (exclui checkboxes e radios)
    const labels = [...document.querySelectorAll('label')];
    for (const label of labels) {
        const chave = (label.innerText || '').trim().replace(/:$/, '');
        if (!chave || chave.startsWith('{{') || chave.length > 120) continue;

        let input = null;
        const forAttr = label.getAttribute('for');
        if (forAttr) input = document.getElementById(forAttr);

        if (!input) {
            let el = label.nextElementSibling;
            while (el && !input) {
                if (['INPUT','TEXTAREA','SELECT'].includes(el.tagName)) input = el;
                else if (el.querySelector) input = el.querySelector('input,textarea,select');
                el = el.nextElementSibling;
            }
        }
        if (!input) {
            const p = label.parentElement;
            if (p) input = p.querySelector('input,textarea,select');
        }
        if (!input) {
            const p = label.parentElement;
            if (p && p.parentElement) input = p.parentElement.querySelector('input,textarea,select');
        }

        if (!input) continue;

        const tipo = (input.type || '').toLowerCase();
        // Captura especificamente o checkbox "Atende SUS"
        if (tipo === 'checkbox') {
            if (chave.includes('Atende SUS') || chave.includes('Atende Sus')) {
                result[chave] = input.checked ? 'Sim' : 'Não';
            }
            continue;
        }
        if (tipo === 'radio') continue;

        const valor = (input.value || input.innerText || '').trim().replace(/\\n+/g, ' ');
        const camposObrigatorios = ['Data Desativação', 'Motivo Desativação'];
        if (camposObrigatorios.includes(chave)) {
            result[chave] = valor || '';  // sempre captura, mesmo vazio
        } else if (valor && !valor.startsWith('{{') && valor !== 'on' && valor !== 'off') {
            result[chave] = valor;
        }
    }

    // Garante que campos de desativação sempre existem no resultado
    if (!('Data Desativação' in result)) result['Data Desativação'] = '';
    if (!('Motivo Desativação' in result)) result['Motivo Desativação'] = '';

    // Fallback: tabelas HTML
    if (Object.keys(result).length === 0) {
        const rows = [...document.querySelectorAll('table tr')];
        for (const row of rows) {
            const cells = [...row.querySelectorAll('td,th')];
            if (cells.length >= 2) {
                const chave = (cells[0].innerText||'').trim().replace(/:$/,'');
                const valor = (cells[1].innerText||'').trim().replace(/\\n+/g, ' ');
                if (chave && valor && !chave.startsWith('{{') && chave.length < 120)
                    result[chave] = valor;
            }
        }
    }

    return result;
}
"""


def _coletar_dados_ficha(page: Page) -> dict:
    """Coleta todos os dados visíveis da ficha do estabelecimento."""
    _aguardar_pagina(page)
    dados = page.evaluate(_JS_COLETAR_FICHA)
    return {k: v for k, v in dados.items() if v and not v.startswith("{{")}


def processar_cnpj(page: Page, cnpj: str, precisa_goto: bool = True) -> tuple:
    """
    Pesquisa um CNPJ no CNES.
    Retorna (resultado_dict, precisa_goto_proximo).
    Se precisa_goto=False, já estamos na página de busca e evitamos o reload.
    """
    resultado = {"CNPJ Pesquisado": cnpj, "ENCONTRADO": "NÃO"}

    for tentativa in range(1, MAX_RETRIES + 1):
        try:
            if precisa_goto or tentativa > 1:
                page.goto(TARGET_URL)
                _aguardar_pagina(page)

            _preencher_e_pesquisar(page, cnpj)
            encontrou = _aguardar_e_verificar_resultado(page)

            if not encontrou:
                return resultado, False  # ainda na página de busca, próximo não precisa goto

            _ir_para_ficha(page)
            dados = _coletar_dados_ficha(page)

            resultado["ENCONTRADO"] = "SIM"
            resultado.update(dados)
            resultado["CNPJ Pesquisado"] = cnpj

            log_sucesso(cnpj)
            time.sleep(DELAY_BETWEEN_ENTRIES)
            return resultado, True  # navegou para ficha, próximo precisa goto

        except PlaywrightTimeout as e:
            log_erro(cnpj, f"Timeout (tentativa {tentativa}): {e}")
        except Exception as e:
            log_erro(cnpj, f"Erro (tentativa {tentativa}): {e}")

        if tentativa < MAX_RETRIES:
            time.sleep(DELAY_ON_ERROR)

    resultado["ENCONTRADO"] = "ERRO"
    log_erro(cnpj, "Máximo de tentativas atingido sem sucesso")
    return resultado, True  # após erros, recarrega por segurança
