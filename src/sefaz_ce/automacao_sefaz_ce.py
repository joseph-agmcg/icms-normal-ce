"""
Classe principal que orquestra o fluxo de navegação no portal DAE da SEFAZ-CE.
Para cada IE: preenche campo IE → Avançar → seleciona receita 1015 → Preencher DAE
→ preenche formulário (período, valor principal, data de pagamento). Não envia o formulário.
"""

import asyncio
import tempfile
from datetime import datetime
from pathlib import Path

from playwright.async_api import async_playwright, Browser, Page, BrowserContext

from sefaz_ce import configuracoes
from sefaz_ce.logger import configurar_logger_da_aplicacao
from sefaz_ce.navegacao.acoes_pagina import (
    aguardar_pagina_carregar,
    tirar_captura_de_tela_em_erro,
)
from sefaz_ce.resolver_captcha import resolver_imagem

logger = configurar_logger_da_aplicacao(__name__)

# Timeout para detectar se a IE foi encontrada (página com cmbReceita)
TIMEOUT_IE_ENCONTRADA_MS = 12_000


def _valor_total_invalido(valor: object) -> bool:
    """Retorna True se o valor TOTAL não deve ser processado (pular a IE)."""
    if valor is None:
        return True
    if isinstance(valor, (int, float)) and valor == 0:
        return True
    if isinstance(valor, str):
        s = valor.strip().lower()
        if not s or s == "null":
            return True
    return False


class AutomacaoConsultaDAE:
    """
    Automação por IE: acessa default.asp, preenche IE, Avançar, seleciona receita,
    Preencher DAE, preenche o formulário (sem enviar). Uma execução por vez com
    timer entre IEs.
    """

    def __init__(self, headless: bool = False) -> None:
        self._headless = headless
        self._playwright = None
        self._browser: Browser | None = None
        self._context: BrowserContext | None = None
        self._pagina: Page | None = None

    async def _iniciar_browser(self) -> None:
        """Inicia o Playwright e abre um navegador com uma página."""
        logger.info("Iniciando navegador.")
        self._playwright = await async_playwright().start()
        self._browser = await self._playwright.chromium.launch(headless=self._headless)
        self._context = await self._browser.new_context(
            viewport={"width": 1280, "height": 720},
            ignore_https_errors=True,
        )
        self._pagina = await self._context.new_page()
        self._pagina.set_default_timeout(configuracoes.TIMEOUT_AGUARDAR_ELEMENTO_MS)
        logger.debug("Navegador e página prontos.")

    async def _encerrar_browser(self) -> None:
        """Fecha o navegador e encerra o Playwright."""
        if self._context:
            await self._context.close()
        if self._browser:
            await self._browser.close()
        if self._playwright:
            await self._playwright.stop()
        logger.info("Navegador encerrado.")

    def _nome_captura_erro(self, etapa: str, ie: str = "") -> str:
        """Gera nome de arquivo para captura de tela em caso de erro."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        sufixo = f"_{ie}" if ie else ""
        return f"erro_{etapa}{sufixo}_{timestamp}.png"

    async def _acessar_pagina_inicial(self) -> None:
        """Acessa a URL do DAE (default.asp)."""
        logger.info("Acessando %s", configuracoes.URL_PORTAL_DAE_SEFAZ_CE)
        await self._pagina.goto(configuracoes.URL_PORTAL_DAE_SEFAZ_CE)
        await aguardar_pagina_carregar(self._pagina)

    async def _preencher_ie_e_avancar(self, ie_digitos: str) -> bool:
        """
        Preenche o campo IE e clica em Avançar.
        Retorna True se a próxima página tiver o select de receita (#cmbReceita); False se der erro (IE não encontrada).
        """
        try:
            campo = self._pagina.locator(configuracoes.SELETOR_CAMPO_IE)
            await campo.wait_for(state="visible", timeout=configuracoes.TIMEOUT_AGUARDAR_ELEMENTO_MS)
            await campo.fill("")
            await campo.fill(ie_digitos)
            logger.debug("Campo IE preenchido com %s", ie_digitos)

            botao = self._pagina.locator(configuracoes.SELETOR_BOTAO_AVANCAR)
            await botao.wait_for(state="visible", timeout=configuracoes.TIMEOUT_AGUARDAR_ELEMENTO_MS)
            await botao.click()
            logger.debug("Clicado em Avançar.")

            # Aguarda aparecer o select de receita (sucesso) ou timeout (IE não encontrada)
            try:
                await self._pagina.locator(configuracoes.SELETOR_SELECT_RECEITA).wait_for(
                    state="visible", timeout=TIMEOUT_IE_ENCONTRADA_MS
                )
                return True
            except Exception:
                logger.warning("Select de receita não apareceu no tempo esperado (IE possivelmente não encontrada).")
                return False
        except Exception as e:
            logger.exception("Erro ao preencher IE e avançar: %s", e)
            await tirar_captura_de_tela_em_erro(
                self._pagina, self._nome_captura_erro("avancar_ie", ie_digitos)
            )
            return False

    async def _selecionar_receita_e_preencher_dae(self) -> None:
        """Seleciona a opção 1015 no #cmbReceita e clica em Preencher DAE."""
        select = self._pagina.locator(configuracoes.SELETOR_SELECT_RECEITA)
        await select.select_option(value=configuracoes.VALOR_OPCAO_RECEITA_ICMS_MENSAL)
        logger.debug("Receita 1015 selecionada.")

        botao = self._pagina.locator(configuracoes.SELETOR_BOTAO_PREENCHER_DAE)
        await botao.click()
        await aguardar_pagina_carregar(self._pagina)
        logger.debug("Clicado em Preencher DAE.")

    async def _preencher_formulario_dae(
        self,
        mes_ref: int,
        ano_ref: int,
        valor_principal: float | None,
    ) -> None:
        """
        Preenche período de referência, data de pagamento (dia 20 do mês posterior) e valor principal (coluna TOTAL).
        Não clica em Cadastrar (não envia).
        """
        # Período de referência: mês e ano
        mes_str = f"{mes_ref:02d}"
        ano_str = str(ano_ref)
        await self._pagina.locator(configuracoes.SELETOR_MES_REFERENCIA).fill(mes_str)
        await self._pagina.locator(configuracoes.SELETOR_ANO_REFERENCIA).fill(ano_str)

        # Data de pagamento: dia 20 do mês posterior ao período de referência
        # Ex.: período 01/2026 → pagamento em 20/02/2026
        if mes_ref == 12:
            mes_pag, ano_pag = 1, ano_ref + 1
        else:
            mes_pag, ano_pag = mes_ref + 1, ano_ref
        await self._pagina.locator(configuracoes.SELETOR_DIA_PAGAMENTO).fill("20")
        await self._pagina.locator(configuracoes.SELETOR_MES_PAGAMENTO).fill(f"{mes_pag:02d}")
        await self._pagina.locator(configuracoes.SELETOR_ANO_PAGAMENTO).fill(str(ano_pag))

        # Valor principal (formato ex.: 1277.39)
        if valor_principal is not None:
            valor_str = f"{valor_principal:.2f}".replace(",", ".")
            await self._pagina.locator(configuracoes.SELETOR_VALOR_PRINCIPAL).fill(valor_str)
            logger.debug("Formulário preenchido: período %s/%s, valor %s", mes_str, ano_str, valor_str)
        else:
            logger.warning("Valor TOTAL não informado para esta IE; campo Valor Principal deixado em branco.")

    async def _resolver_e_preencher_captcha(self) -> bool:
        """
        Aguarda a imagem do captcha, envia ao Anti-Captcha, preenche o campo strCAPTCHA.
        Retorna True se o captcha foi resolvido e preenchido; False em caso de erro ou API key vazia.
        """
        try:
            img = self._pagina.locator(configuracoes.SELETOR_IMAGEM_CAPTCHA)
            await img.wait_for(state="visible", timeout=configuracoes.TIMEOUT_AGUARDAR_ELEMENTO_MS)
        except Exception as e:
            logger.warning("Captcha não encontrado na página (pode não estar presente): %s", e)
            return True  # sem captcha na tela, segue fluxo

        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
            path_captcha = Path(f.name)
        try:
            await img.screenshot(path=str(path_captcha))
            texto = await asyncio.to_thread(resolver_imagem, str(path_captcha))
            if not texto:
                logger.error("Não foi possível resolver o captcha (Anti-Captcha sem retorno ou API key vazia).")
                return False
            campo = self._pagina.locator(configuracoes.SELETOR_INPUT_CAPTCHA)
            await campo.fill("")
            await campo.fill(texto[:6])  # maxlength=6
            logger.debug("Captcha preenchido com sucesso.")
            return True
        except Exception as e:
            logger.exception("Erro ao resolver/preencher captcha: %s", e)
            return False
        finally:
            path_captcha.unlink(missing_ok=True)

    async def executar_fluxo_por_ie(
        self,
        lista_dados: list[dict[str, object]],
    ) -> tuple[list[str], list[tuple[str, str]]]:
        """
        Para cada item em lista_dados (ie, ie_digitos, valor_normal, mes_ref, ano_ref):
        acessa a página, preenche IE, avança, seleciona receita, preenche DAE, preenche o formulário (sem enviar).
        Aguarda INTERVALO_ENTRE_EXECUCOES_MS entre uma IE e a próxima.
        Retorna (lista de IEs com sucesso, lista de (IE, motivo) com erro para revisão).
        """
        ies_sucesso: list[str] = []
        ies_erro: list[tuple[str, str]] = []  # (ie, motivo simplificado)
        total = len(lista_dados)

        try:
            await self._iniciar_browser()
            await self._acessar_pagina_inicial()

            for indice, item in enumerate(lista_dados):
                ie = str(item.get("ie", ""))
                ie_digitos = str(item.get("ie_digitos", ""))
                valor_normal = item.get("valor_normal")
                try:
                    mes_ref = int(item.get("mes_ref"))
                    ano_ref = int(item.get("ano_ref"))
                except (TypeError, ValueError):
                    ies_erro.append((ie or "(vazio)", "Período (mês/ano) ausente nos dados da planilha"))
                    continue

                if not ie or not ie_digitos:
                    ies_erro.append((ie or "(vazio)", "IE inválida ou vazia"))
                    continue

                # Pula IEs sem valor TOTAL (None, 0, vazio, null) — não executa e não conta como erro
                if _valor_total_invalido(valor_normal):
                    logger.info(
                        "IE %s pulada: valor TOTAL ausente, zero ou vazio (não executada).",
                        ie,
                    )
                    continue

                logger.info("Processando IE %s (%d/%d).", ie, indice + 1, total)

                # Timer entre IEs (exceto antes da primeira)
                if indice > 0:
                    ms = configuracoes.INTERVALO_ENTRE_EXECUCOES_MS
                    logger.info(
                        "Aguardando %d ms (%.1f s) antes da próxima IE.",
                        ms, ms / 1000.0,
                    )
                    await asyncio.sleep(ms / 1000.0)
                    await self._acessar_pagina_inicial()

                # Passo 1: preencher IE e Avançar
                ok = await self._preencher_ie_e_avancar(ie_digitos)
                if not ok:
                    ies_erro.append((ie, "IE não encontrada no site ao clicar em Avançar"))
                    continue

                # Passo 2: selecionar receita e Preencher DAE
                try:
                    await self._selecionar_receita_e_preencher_dae()
                except Exception as e:
                    logger.exception("Erro ao selecionar receita/preencher DAE para IE %s: %s", ie, e)
                    await tirar_captura_de_tela_em_erro(
                        self._pagina, self._nome_captura_erro("receita_dae", ie)
                    )
                    motivo = str(e).split("\n")[0].strip() if str(e) else "Falha ao selecionar receita ou Preencher DAE"
                    if len(motivo) > 80:
                        motivo = motivo[:77] + "..."
                    ies_erro.append((ie, motivo))
                    continue

                # Passo 3: aguardar formulário e preencher (sem enviar)
                try:
                    await self._pagina.locator(configuracoes.SELETOR_MES_REFERENCIA).wait_for(
                        state="visible", timeout=configuracoes.TIMEOUT_AGUARDAR_ELEMENTO_MS
                    )
                    await self._preencher_formulario_dae(mes_ref, ano_ref, valor_normal)
                except Exception as e:
                    logger.exception("Erro ao preencher formulário para IE %s: %s", ie, e)
                    await tirar_captura_de_tela_em_erro(
                        self._pagina, self._nome_captura_erro("formulario", ie)
                    )
                    motivo = str(e).split("\n")[0].strip() if str(e) else "Falha ao preencher formulário"
                    if len(motivo) > 80:
                        motivo = motivo[:77] + "..."
                    ies_erro.append((ie, motivo))
                    continue

                # Passo 4: resolver captcha e preencher campo strCAPTCHA
                try:
                    ok_captcha = await self._resolver_e_preencher_captcha()
                    if not ok_captcha:
                        ies_erro.append((ie, "Falha ao resolver ou preencher o captcha"))
                        continue
                except Exception as e:
                    logger.exception("Erro no captcha para IE %s: %s", ie, e)
                    await tirar_captura_de_tela_em_erro(
                        self._pagina, self._nome_captura_erro("captcha", ie)
                    )
                    motivo = str(e).split("\n")[0].strip() if str(e) else "Erro ao processar captcha"
                    if len(motivo) > 80:
                        motivo = motivo[:77] + "..."
                    ies_erro.append((ie, motivo))
                    continue

                ies_sucesso.append(ie)
                logger.info("IE %s concluída (formulário e captcha preenchidos, não enviado).", ie)

        finally:
            await self._encerrar_browser()

        logger.info("Fluxo finalizado: %d sucesso, %d erro.", len(ies_sucesso), len(ies_erro))
        return ies_sucesso, ies_erro

    async def executar_fluxo_completo(self) -> None:
        """
        Mantido para compatibilidade: inicia browser, acessa a página inicial e encerra.
        Para execução por IE com formulário, use executar_fluxo_por_ie(lista_dados).
        """
        try:
            await self._iniciar_browser()
            await self._acessar_pagina_inicial()
            logger.info("Fluxo completo (somente acesso) executado.")
        finally:
            await self._encerrar_browser()
