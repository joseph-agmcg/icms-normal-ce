"""
Ponto de entrada principal da aplicação.
Executa o fluxo completo de automação no portal DAE da SEFAZ-CE.
"""

import asyncio
import sys

from sefaz_ce.automacao_sefaz_ce import AutomacaoConsultaDAE
from sefaz_ce.logger import configurar_logger_da_aplicacao

logger = configurar_logger_da_aplicacao(__name__)


async def _rodar_automacao(headless: bool = False) -> None:
    """Executa a automação e trata falhas críticas."""
    automacao = AutomacaoConsultaDAE(headless=headless)
    try:
        await automacao.executar_fluxo_completo()
    except Exception:
        logger.critical("Falha crítica na execução da automação.", exc_info=True)
        raise


def main() -> None:
    """Entrada principal: configura asyncio e dispara a automação."""
    logger.info("Iniciando automação SEFAZ-CE DAE.")
    headless = "--headless" in sys.argv
    try:
        asyncio.run(_rodar_automacao(headless=headless))
    except KeyboardInterrupt:
        logger.warning("Execução interrompida pelo usuário.")
        sys.exit(130)
    except Exception:
        sys.exit(1)
    logger.info("Encerramento normal.")


if __name__ == "__main__":
    main()
