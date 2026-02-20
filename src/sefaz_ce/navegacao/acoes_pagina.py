"""
Funções atômicas de interação com o browser (clicar, aguardar, captura de tela em erro).
"""

import asyncio
from pathlib import Path

from playwright.async_api import Page

from sefaz_ce import configuracoes
from sefaz_ce.logger import configurar_logger_da_aplicacao

logger = configurar_logger_da_aplicacao(__name__)


async def aguardar_pagina_carregar(pagina: Page) -> None:
    """Aguarda a página atingir estado networkidle (rede ociosa)."""
    logger.debug("Aguardando estado networkidle da página.")
    await pagina.wait_for_load_state("networkidle", timeout=configuracoes.TIMEOUT_PAGINA_CARREGAR_MS)
    logger.debug("Página em estado networkidle.")


async def clicar_em_link_por_texto(pagina: Page, texto_do_link: str) -> None:
    """
    Clica em um link cujo texto visível corresponde exatamente ao informado.
    Aguarda o elemento estar visível antes de clicar.
    """
    logger.debug("Procurando link com texto: %s", texto_do_link)
    locator = pagina.get_by_role("link", name=texto_do_link)
    await locator.wait_for(state="visible", timeout=configuracoes.TIMEOUT_AGUARDAR_ELEMENTO_MS)
    await locator.click()
    logger.info("Clicado no link: %s", texto_do_link)


async def clicar_em_elemento_por_texto(pagina: Page, texto_visivel: str) -> None:
    """
    Clica em qualquer elemento (botão, link, etc.) cujo texto visível corresponda ao informado.
    Aguarda o elemento estar visível antes de clicar.
    """
    logger.debug("Procurando elemento com texto: %s", texto_visivel)
    locator = pagina.get_by_text(texto_visivel, exact=True)
    await locator.wait_for(state="visible", timeout=configuracoes.TIMEOUT_AGUARDAR_ELEMENTO_MS)
    await locator.click()
    logger.info("Clicado no elemento: %s", texto_visivel)


async def tirar_captura_de_tela_em_erro(pagina: Page, nome_arquivo: str) -> Path:
    """
    Salva um screenshot da página na pasta de capturas de erro.
    Cria a pasta se não existir. Retorna o caminho do arquivo salvo.
    """
    pasta = configuracoes.PASTA_CAPTURAS_ERROS_ABSOLUTA
    pasta.mkdir(parents=True, exist_ok=True)
    caminho = pasta / nome_arquivo
    await pagina.screenshot(path=str(caminho))
    logger.warning("Captura de tela salva em erro: %s", caminho)
    return caminho
