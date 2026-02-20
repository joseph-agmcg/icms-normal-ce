"""
Configuração centralizada do logging.
Saída: terminal (INFO+) e arquivo em logs/ (DEBUG+) com timestamp no nome.
"""

import logging
import sys
from datetime import datetime
from pathlib import Path


def configurar_logger_da_aplicacao(nome_do_modulo: str) -> logging.Logger:
    """
    Retorna um logger configurado para o módulo.
    Escreve no terminal em INFO+ e em arquivo .log em logs/ em DEBUG+.
    """
    logger = logging.getLogger(nome_do_modulo)

    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)
    logger.propagate = False

    pasta_logs = Path(__file__).resolve().parent.parent.parent / "logs"
    pasta_logs.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_log = pasta_logs / f"sefaz_ce_{timestamp}.log"

    formato_detalhado = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    formato_console = logging.Formatter("%(levelname)-8s | %(name)s | %(message)s")

    handler_arquivo = logging.FileHandler(arquivo_log, encoding="utf-8")
    handler_arquivo.setLevel(logging.DEBUG)
    handler_arquivo.setFormatter(formato_detalhado)
    logger.addHandler(handler_arquivo)

    handler_console = logging.StreamHandler(sys.stdout)
    handler_console.setLevel(logging.INFO)
    handler_console.setFormatter(formato_console)
    logger.addHandler(handler_console)

    logger.debug("Logger configurado: arquivo=%s", arquivo_log)
    return logger
