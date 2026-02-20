"""
Resolução de captcha de imagem via Anti-Captcha (anti-captcha.com).
Usado no formulário DAE da SEFAZ-CE: imagem #imgCaptcha → texto no campo #strCAPTCHA.
"""

from __future__ import annotations

from anticaptchaofficial.imagecaptcha import imagecaptcha

from sefaz_ce import configuracoes
from sefaz_ce.logger import configurar_logger_da_aplicacao

logger = configurar_logger_da_aplicacao(__name__)


def resolver_imagem(caminho_imagem: str) -> str | None:
    """
    Envia a imagem do captcha para o Anti-Captcha e retorna o texto reconhecido.

    Args:
        caminho_imagem: Caminho do arquivo da imagem (ex.: screenshot do #imgCaptcha).

    Returns:
        Texto do captcha (até 6 caracteres) ou None em caso de erro.
    """
    api_key = configuracoes.ANTI_CAPTCHA_API_KEY

    solver = imagecaptcha()
    solver.set_verbose(0)
    solver.set_key(api_key)

    texto = solver.solve_and_return_solution(caminho_imagem)
    if texto is None or texto == 0 or (isinstance(texto, str) and not texto.strip()):
        logger.warning("Anti-Captcha não retornou solução: %s", getattr(solver, "error_code", "unknown"))
        return None

    resultado = texto.strip() if isinstance(texto, str) else str(texto).strip()
    logger.debug("Captcha resolvido: %s", resultado)
    return resultado
