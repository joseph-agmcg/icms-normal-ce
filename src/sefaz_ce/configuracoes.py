"""
Constantes, seletores, URLs e timeouts do projeto.
Carrega variáveis de ambiente do .env — nenhum outro módulo deve usar load_dotenv ou os.environ.
"""

import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

# --- URLs ---
URL_PORTAL_DAE_SEFAZ_CE = "https://servicos.sefaz.ce.gov.br/internet/dae/aplic/default.asp"

# --- Página inicial: campo IE e botão Avançar ---
SELETOR_CAMPO_IE = "input[name=\"txtValor\"]"
SELETOR_BOTAO_AVANCAR = "input[name=\"ok\"]"

# --- Página seleção de receita ---
SELETOR_SELECT_RECEITA = "#cmbReceita"
VALOR_OPCAO_RECEITA_ICMS_MENSAL = "1015 - ICMS Regime Mensal de Apuração1"
SELETOR_BOTAO_PREENCHER_DAE = "input[name=\"ok\"]"

# --- Formulário Preencher DAE (preencher.asp) — não enviar, só preencher ---
SELETOR_MES_REFERENCIA = "#txtMesPeriodoReferencia"
SELETOR_ANO_REFERENCIA = "#txtAnoPeriodoReferencia"
SELETOR_DIA_PAGAMENTO = "input[name=\"txtDiaPagamento\"]"
SELETOR_MES_PAGAMENTO = "input[name=\"txtMesPagamento\"]"
SELETOR_ANO_PAGAMENTO = "input[name=\"txtAnoPagamento\"]"
SELETOR_VALOR_PRINCIPAL = "input[name=\"txtValorPrincipal\"]"

# --- Captcha no formulário DAE (após valor principal) ---
SELETOR_IMAGEM_CAPTCHA = "#imgCaptcha"
SELETOR_INPUT_CAPTCHA = "#strCAPTCHA"

# --- Textos dos links/botões do fluxo (usados com get_by_text no Playwright) ---
TEXTO_LINK_PORTAL_DE_SERVICOS = "Portal de Serviços"
TEXTO_LINK_SERVICOS = "Serviços"
TEXTO_LINK_EMISSAO_DE_DAE = "Emissão de DAE"
TEXTO_LINK_EMISSAO_DAE_ICMS = "Emissão DAE ICMS"

# --- Timeouts (milissegundos) ---
TIMEOUT_PAGINA_CARREGAR_MS = 30_000
TIMEOUT_AGUARDAR_ELEMENTO_MS = 15_000

# ========== CONFIGURAÇÕES DE LOTE (altere aqui) ==========
INTERVALO_ENTRE_EXECUCOES_MS = 10_000   # tempo entre cada execução (ms). Ex.: 600_000 = 10 min, 30_000 = 30 s
QUANTIDADE_POR_VEZ = 1         # quantas IEs executar por vez (1 = uma de cada vez em sequência)
# =========================================================

# --- Pastas (a partir de variáveis de ambiente) ---
PASTA_SAIDA_RESULTADOS = os.getenv("PASTA_SAIDA_RESULTADOS", "resultados")
PASTA_CAPTURAS_DE_TELA_ERROS = os.getenv("PASTA_CAPTURAS_DE_TELA_ERROS", "capturas_erros")

# Caminhos absolutos a partir da raiz do projeto
_raiz_projeto = Path(__file__).resolve().parent.parent.parent
PASTA_SAIDA_RESULTADOS_ABSOLUTA = _raiz_projeto / PASTA_SAIDA_RESULTADOS
PASTA_CAPTURAS_ERROS_ABSOLUTA = _raiz_projeto / PASTA_CAPTURAS_DE_TELA_ERROS

# --- Anti-Captcha (resolução de captcha de imagem no formulário DAE) — obrigatória ---
_key = os.getenv("ANTI_CAPTCHA_API_KEY")
if not _key or not str(_key).strip():
    raise SystemExit("ANTI_CAPTCHA_API_KEY é obrigatória. Configure no .env.")
ANTI_CAPTCHA_API_KEY = str(_key).strip()
