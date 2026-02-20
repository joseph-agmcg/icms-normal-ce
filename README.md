# ICMS Normal - Ceará

Automação com **Playwright (Python assíncrono)**

---

## Passo a passo do processo manual que a automação substitui

1. Acessar no navegador o portal DAE da SEFAZ-CE (`https://servicos.sefaz.ce.gov.br/internet/dae/aplic/default.asp`).
2. Informar a Inscrição Estadual (I.E.) no campo e clicar em **Avançar**.
3. Na tela de seleção de receita, escolher **1015 - ICMS Regime Mensal de Apuração** e clicar em **Preencher DAE**.
4. No formulário DAE, preencher período de referência (mês/ano), data de pagamento (ex.: dia 20), valor principal e, quando houver, o CAPTCHA.
5. Clicar no **botão de cadastro** para enviar o formulário.
6. Fazer o **download do PDF** da guia DAE gerada.
7. **Renomear o arquivo** no padrão: `DAE ICMS NORMAL {inscrição estadual} REF {competência}` — exemplo: `DAE ICMS NORMAL 062586416 REF 012026` (REF = mês/ano da competência, ex.: 012026 = jan/2026).
8. Repetir os passos 1 a 7 para cada I.E./filial (em geral a partir de uma planilha Excel com as I.E.s e valores).

A automação replica esse fluxo por I.E., preenchendo os campos a partir do Excel. Envio do formulário, download do PDF e renomeação automática estão nos planos futuros.

---

## Vídeo do processo manual

Quando aplicável, vídeo do processo manual para referência futura:

- **[Link do vídeo do processo manual](https://drive.google.com/file/d/1LN_sLWcaVUN4Nn3CYWAsQj0zTn1dEnVV/view?usp=sharing)**

---

## O que a automação faz

- Acessa o portal DAE da SEFAZ-CE (página inicial `default.asp`).
- Para cada I.E. da planilha: preenche I.E. e Avançar, seleciona receita **1015 (ICMS Regime Mensal)**, clica em **Preencher DAE** e preenche o formulário (período, valor, data de pagamento). Resolução do CAPTCHA via Anti-Captcha (chave no `.env`).
- Permite carregar um Excel de filiais (coluna I.E. = INSC.ESTADUAL ou sinônimos), visualizar os dados extraídos e rodar o fluxo em lote para todas as I.E.
- Em erro: registra log e salva screenshot na pasta de erros configurada.

**Planos futuros (a implementar):**

- Enviar o formulário (clique no botão de cadastro após preenchimento).
- Download automático do PDF da guia DAE gerada.
- Renomeação automática do arquivo no padrão: `DAE ICMS NORMAL {inscrição estadual} REF {competência}` (ex.: `DAE ICMS NORMAL 062586416 REF 012026`).
- Empacotar toda a automação em um executável (ex.: PyInstaller/Nuitka) para distribuição sem exigir instalação de Python.

---

## Dependências

- **Python** 3.13+
- **Pacotes Python** (na raiz do projeto):
  ```bash
  pip install -r requirements.txt
  ```
  Inclui: `playwright`, `openpyxl`, `customtkinter`, `python-dotenv`, `anticaptchaofficial`.
- **Navegador Playwright** (Chromium):
  ```bash
  python -m playwright install chromium
  ```

O projeto não é instalado como pacote (`pip install -e .`); use `PYTHONPATH=src` ou o script `run.ps1` para executar.

---

## Variáveis de ambiente necessárias

Copie `.env.example` para `.env` e preencha conforme necessário:

| Variável | Obrigatória | Descrição |
|----------|-------------|-----------|
| `ANTI_CAPTCHA_API_KEY` | **Sim** | Chave da Anti-Captcha para resolução do CAPTCHA do DAE. Obrigatória; ausência = erro fatal. |
| `PASTA_SAIDA_RESULTADOS` | Não | Pasta para resultados (padrão: `resultados`). |
| `PASTA_CAPTURAS_DE_TELA_ERROS` | Não | Pasta para screenshots em caso de erro (padrão: `capturas_erros`). |

**Não commitar o arquivo `.env`.**

---

## Como executar

### Interface gráfica (Excel + lote)

1. Defina o `PYTHONPATH` e rode a GUI:
   ```powershell
   $env:PYTHONPATH="src"; python -m sefaz_ce.gui_app
   ```
2. Selecione um arquivo `.xlsx` com coluna de Inscrição Estadual (ex.: "INSC.ESTADUAL").
3. Use **Ver dados extraídos** para conferir a extração.
4. Execute a automação para todas as I.E. da planilha.

### Só automação (linha de comando)

- Com script (recomendado):
  ```powershell
  .\run.ps1
  .\run.ps1 --headless
  ```
- Manualmente:
  ```powershell
  $env:PYTHONPATH="src"; python -m sefaz_ce.main
  $env:PYTHONPATH="src"; python -m sefaz_ce.main --headless
  ```

**Planilha Excel:** Cabeçalho com coluna de I.E. (ex.: "INSC.ESTADUAL"); dados lidos até a linha de TOTAL. I.E. com 8 dígitos é normalizada com zero à esquerda (9 dígitos no Ceará).

**Logs:** Terminal em INFO+; arquivo em `logs/` em DEBUG+.

---

## Estrutura do projeto

| Pasta/Arquivo | Função |
|---------------|--------|
| **`pyproject.toml`** | Metadados do projeto (nome, versão). |
| **`requirements.txt`** | Dependências para `pip install -r requirements.txt`. |
| **`run.ps1`** | Script para rodar a automação com `PYTHONPATH=src`. |
| **`.env`** | Variáveis sensíveis. **Não commitar.** |
| **`.env.example`** | Exemplo do `.env` sem valores reais. |
| **`logs/`** | Criada automaticamente; arquivos `.log` com timestamp. |
| **`src/sefaz_ce/main.py`** | Ponto de entrada CLI; chama `executar_fluxo_completo()`; suporta `--headless`. |
| **`src/sefaz_ce/gui_app.py`** | Interface gráfica (CustomTkinter): upload Excel, extração I.E., execução em lote. |
| **`src/sefaz_ce/automacao_sefaz_ce.py`** | Classe principal `AutomacaoConsultaDAE`; orquestra o fluxo por I.E. |
| **`src/sefaz_ce/configuracoes.py`** | Carrega `.env`; URLs, seletores, timeouts, pastas. |
| **`src/sefaz_ce/excel_filiais.py`** | Extração de dados do Excel (formato ICMS/filiais). |
| **`src/sefaz_ce/logger.py`** | Logging: terminal INFO+, arquivo DEBUG+ em `logs/`. |
| **`src/sefaz_ce/navegacao/`** | Ações atômicas na página (clicar links, aguardar, screenshot em erro). |
