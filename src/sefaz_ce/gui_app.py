"""
Interface desktop com CustomTkinter.
Upload de arquivo Excel → extração de I.E. e coluna TOTAL → detalhamento (executáveis x ignoradas)
e opção de escolher quais IEs executar. Execução em lote com resultado (sucesso/erro).
"""

import asyncio
import sys
import threading
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from sefaz_ce import configuracoes
from sefaz_ce.automacao_sefaz_ce import AutomacaoConsultaDAE, _valor_total_invalido
from sefaz_ce.excel_filiais import extrair_todos_os_dados, obter_dados_para_dae, obter_ies_dos_dados
from sefaz_ce.excel_filiais import _obter_chave_ie
from sefaz_ce.logger import configurar_logger_da_aplicacao

logger = configurar_logger_da_aplicacao(__name__)


def _ie_para_exibicao(ie: str | None) -> str:
    """Retorna a I.E. sem pontos, traços ou barras para exibição na GUI."""
    if ie is None:
        return ""
    s = str(ie).strip()
    return s.replace(".", "").replace("-", "").replace("/", "")


def _formato_intervalo_ms(ms: int) -> str:
    """Retorna o intervalo em texto legível (ex.: '10 min', '30 s')."""
    if ms >= 60_000:
        return f"{ms // 60_000} min"
    if ms >= 1_000:
        return f"{ms // 1_000} s"
    return f"{ms} ms"


def _contar_executaveis_ignoradas(lista_dados: list[dict[str, object]]) -> tuple[int, int]:
    """Retorna (qtd_executáveis, qtd_ignoradas)."""
    executaveis = sum(1 for item in lista_dados if not _valor_total_invalido(item.get("valor_normal")))
    return executaveis, len(lista_dados) - executaveis


def _executar_lote_em_background(
    lista_dados: list[dict[str, object]],
    headless: bool,
    result_callback: None = None,
) -> None:
    """
    Roda a automação para todas as IEs em sequência (quantidade e intervalo em configuracoes).
    result_callback(ies_sucesso, ies_erro) é chamado ao final na thread da GUI.
    """
    total = len(lista_dados)
    intervalo_txt = _formato_intervalo_ms(configuracoes.INTERVALO_ENTRE_EXECUCOES_MS)
    logger.info(
        "Iniciando lote: %d I.E.(s), %d por vez, %s entre cada.",
        total, configuracoes.QUANTIDADE_POR_VEZ, intervalo_txt,
    )
    print(
        f"\n[Progresso] {total} I.E.(s). {configuracoes.QUANTIDADE_POR_VEZ} por vez, "
        f"{intervalo_txt} entre cada.\n",
        flush=True,
    )

    async def _rodar() -> tuple[list[str], list[str]]:
        automacao = AutomacaoConsultaDAE(headless=headless)
        return await automacao.executar_fluxo_por_ie(lista_dados)

    ies_sucesso, ies_erro = asyncio.run(_rodar())

    print(f"\n[Progresso] Concluído: {len(ies_sucesso)} sucesso, {len(ies_erro)} erro.", flush=True)
    logger.info("Lote finalizado: %d sucesso, %d erro.", len(ies_sucesso), len(ies_erro))

    if result_callback:
        try:
            result_callback(ies_sucesso, ies_erro)
        except Exception:
            pass


def _mostrar_janela_resultado(
    parent: ctk.CTk,
    ies_sucesso: list[str],
    ies_erro: list[tuple[str, str]],
) -> None:
    """Abre janela com IEs com sucesso e IEs com erro (IE + motivo resumido)."""
    janela = ctk.CTkToplevel(parent)
    janela.title("Resultado da automação — SEFAZ-CE")
    janela.geometry("620x440")
    janela.minsize(450, 340)

    ctk.CTkLabel(
        janela,
        text="Revise as I.E. processadas:",
        font=ctk.CTkFont(size=14, weight="bold"),
    ).pack(pady=(12, 8))

    frame = ctk.CTkFrame(janela, fg_color="transparent")
    frame.pack(fill="both", expand=True, padx=16, pady=8)

    # Coluna Sucesso
    ctk.CTkLabel(frame, text="I.E. com sucesso", font=ctk.CTkFont(size=12, weight="bold"), text_color="green").grid(row=0, column=0, padx=(0, 12), pady=(0, 4), sticky="w")
    txt_sucesso = ctk.CTkTextbox(frame, width=280, height=240, font=ctk.CTkFont(family="Consolas", size=11))
    txt_sucesso.grid(row=1, column=0, padx=(0, 12), pady=(0, 12), sticky="nsew")
    txt_sucesso.insert("1.0", "\n".join(_ie_para_exibicao(ie) for ie in ies_sucesso) if ies_sucesso else "(nenhuma)")
    txt_sucesso.configure(state="disabled")

    # Coluna Erro (IE — motivo)
    ctk.CTkLabel(frame, text="I.E. com erro (revisar)", font=ctk.CTkFont(size=12, weight="bold"), text_color="orange").grid(row=0, column=1, padx=(12, 0), pady=(0, 4), sticky="w")
    txt_erro = ctk.CTkTextbox(frame, width=280, height=240, font=ctk.CTkFont(family="Consolas", size=11))
    txt_erro.grid(row=1, column=1, padx=(12, 0), pady=(0, 12), sticky="nsew")
    if ies_erro:
        linhas_erro = [f"{_ie_para_exibicao(ie)} — {motivo}" for ie, motivo in ies_erro]
        txt_erro.insert("1.0", "\n".join(linhas_erro))
    else:
        txt_erro.insert("1.0", "(nenhuma)")
    txt_erro.configure(state="disabled")

    frame.grid_columnconfigure(0, weight=1)
    frame.grid_columnconfigure(1, weight=1)

    ctk.CTkLabel(
        janela,
        text=f"Total: {len(ies_sucesso)} sucesso  |  {len(ies_erro)} com erro",
        font=ctk.CTkFont(size=12),
        text_color="gray",
    ).pack(pady=(0, 12))


class JanelaSelecaoIEs(ctk.CTkToplevel):
    """
    Janela de detalhamento: lista todas as IEs com valor TOTAL e status (executável / será ignorada).
    Usuário marca quais executar; botão 'Executar selecionados' roda só as marcadas.
    """

    def __init__(
        self,
        parent: ctk.CTk,
        lista_dados: list[dict[str, object]],
        ao_executar: "None | tuple[callable, bool]" = None,
    ) -> None:
        super().__init__(parent)
        self.title("Escolher IEs para executar — SEFAZ-CE")
        self.geometry("700x540")
        self.minsize(500, 400)

        self._lista_dados = lista_dados
        self._ao_executar_callback, self._headless = ao_executar or (None, False)
        self._vars: list[tuple[dict[str, object], ctk.BooleanVar]] = []

        qtd_exec, qtd_ign = _contar_executaveis_ignoradas(lista_dados)
        total = len(lista_dados)

        # Resumo no topo
        ctk.CTkLabel(
            self,
            text=f"Total: {total} I.E.(s)  |  Com valor TOTAL (executáveis): {qtd_exec}  |  Sem valor (ignoradas): {qtd_ign}",
            font=ctk.CTkFont(size=12),
            text_color="gray",
        ).pack(pady=(12, 4))

        ctk.CTkLabel(
            self,
            text="Marque as I.E. que deseja executar. I.E. sem valor TOTAL não podem ser marcadas.",
            font=ctk.CTkFont(size=11),
            text_color="gray",
        ).pack(pady=(0, 8))

        # Frame rolável com linhas: [checkbox] IE | Valor TOTAL | Status
        self._frame_scroll = ctk.CTkScrollableFrame(self, width=600, height=280)
        self._frame_scroll.pack(fill="both", expand=True, padx=12, pady=8)

        for item in lista_dados:
            ie = _ie_para_exibicao(item.get("ie") or item.get("ie_digitos") or "")
            val = item.get("valor_normal")
            executavel = not _valor_total_invalido(val)
            valor_txt = f"{val:.2f}" if isinstance(val, (int, float)) else (str(val) if val is not None else "—")
            status = "Executável" if executavel else "Será ignorada"

            row = ctk.CTkFrame(self._frame_scroll, fg_color="transparent")
            row.pack(fill="x", pady=2)

            var = ctk.BooleanVar(value=executavel)
            if executavel:
                cb = ctk.CTkCheckBox(row, text="", variable=var, width=28)
                cb.pack(side="left", padx=(0, 8), pady=4)
                self._vars.append((item, var))
            else:
                ctk.CTkLabel(row, text="  —  ", width=36).pack(side="left", padx=(0, 8), pady=4)

            ctk.CTkLabel(row, text=ie, font=ctk.CTkFont(family="Consolas", size=12), width=140).pack(side="left", padx=4, pady=4)
            ctk.CTkLabel(row, text=valor_txt, font=ctk.CTkFont(family="Consolas", size=12), width=100).pack(side="left", padx=4, pady=4)
            ctk.CTkLabel(
                row,
                text=status,
                font=ctk.CTkFont(size=11),
                text_color="green" if executavel else "gray",
            ).pack(side="left", padx=4, pady=4)

        # Botões: selecionar todas / desmarcar todas
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=12, pady=8)

        ctk.CTkButton(btn_frame, text="Selecionar todas executáveis", command=self._marcar_todas, width=200).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_frame, text="Desmarcar todas", command=self._desmarcar_todas, width=160).pack(side="left", padx=(0, 8))

        # Contador e botão executar selecionados
        self._label_contador = ctk.CTkLabel(
            btn_frame,
            text=f"Selecionadas: {qtd_exec}",
            font=ctk.CTkFont(size=12),
            text_color="gray",
        )
        self._label_contador.pack(side="left", padx=12)
        for item, var in self._vars:
            var.trace_add("write", lambda *_: self._atualizar_contador())

        self._btn_executar_sel = ctk.CTkButton(
            btn_frame,
            text=f"Executar selecionadas ({qtd_exec})",
            command=self._ao_executar_selecionadas,
            width=220,
            fg_color="green",
            hover_color="darkgreen",
        )
        self._btn_executar_sel.pack(side="right")

    def _marcar_todas(self) -> None:
        for _, var in self._vars:
            var.set(True)

    def _desmarcar_todas(self) -> None:
        for _, var in self._vars:
            var.set(False)

    def _atualizar_contador(self) -> None:
        n = sum(1 for _, var in self._vars if var.get())
        self._label_contador.configure(text=f"Selecionadas: {n}")
        self._btn_executar_sel.configure(text=f"Executar selecionadas ({n})")

    def _ao_executar_selecionadas(self) -> None:
        selecionados = [item for item, var in self._vars if var.get()]
        if not selecionados:
            messagebox.showwarning("Aviso", "Nenhuma I.E. selecionada. Marque ao menos uma para executar.")
            return
        self.destroy()
        if self._ao_executar_callback:
            self._ao_executar_callback(selecionados)


class AppFiliais(ctk.CTk):
    """Janela principal: upload do Excel e execução em lote."""

    def __init__(self) -> None:
        super().__init__()
        self.title("SEFAZ-CE — Automação por I.E.")
        self.geometry("520x380")
        self.minsize(400, 340)

        self._caminho_arquivo: Path | None = None
        self._lista_ies: list[str] = []
        self._dados_extraidos: list[dict[str, object]] = []
        self._lista_dados: list[dict[str, object]] = []  # ie, ie_digitos, valor_normal, mes_ref, ano_ref
        self._nomes_colunas: list[str] = []
        self._nome_para_indice: dict[str, int] = {}
        self._mes_ref: int = 1
        self._ano_ref: int = 2026
        self._em_execucao = False

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self._criar_widgets()

    def _criar_widgets(self) -> None:
        padding = {"padx": 20, "pady": 12}

        # Título
        self._titulo = ctk.CTkLabel(
            self,
            text="Upload do arquivo Excel (filiais)",
            font=ctk.CTkFont(size=16, weight="bold"),
        )
        self._titulo.pack(**padding)

        # Botão de upload
        self._btn_upload = ctk.CTkButton(
            self,
            text="Selecionar arquivo .xlsx",
            command=self._ao_selecionar_arquivo,
            width=280,
        )
        self._btn_upload.pack(**padding)

        # Label com resultado do upload (arquivo + quantidade de I.E.)
        self._label_arquivo = ctk.CTkLabel(
            self,
            text="Nenhum arquivo selecionado.",
            font=ctk.CTkFont(size=12),
            text_color="gray",
        )
        self._label_arquivo.pack(**padding)

        # Botão ver dados extraídos
        self._btn_ver_dados = ctk.CTkButton(
            self,
            text="Ver dados extraídos",
            command=self._ao_ver_dados,
            width=280,
            state="disabled",
        )
        self._btn_ver_dados.pack(**padding)

        # Botão escolher IEs e executar (detalhamento + seleção)
        self._btn_escolher = ctk.CTkButton(
            self,
            text="Escolher IEs e executar",
            command=self._ao_abrir_selecao_ies,
            width=280,
            state="disabled",
        )
        self._btn_escolher.pack(**padding)

        # Botão executar todas executáveis
        self._btn_executar = ctk.CTkButton(
            self,
            text="Executar todas executáveis",
            command=self._ao_executar_todas,
            state="disabled",
            width=280,
            fg_color="green",
            hover_color="darkgreen",
        )
        self._btn_executar.pack(**padding)

        # Status (executando / concluído)
        self._label_status = ctk.CTkLabel(
            self,
            text="",
            font=ctk.CTkFont(size=12),
            text_color="gray",
        )
        self._label_status.pack(**padding)

    def _ao_selecionar_arquivo(self) -> None:
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
        )
        if not caminho:
            return

        self._caminho_arquivo = Path(caminho)
        logger.info("Arquivo selecionado: %s", self._caminho_arquivo.resolve())
        try:
            self._dados_extraidos, self._nome_para_indice, self._mes_ref, self._ano_ref = extrair_todos_os_dados(self._caminho_arquivo)
            self._nomes_colunas = [k for k, _ in sorted(self._nome_para_indice.items(), key=lambda x: x[1])]
            self._lista_ies = obter_ies_dos_dados(self._dados_extraidos, self._nome_para_indice)
            logger.info(
                "Extração concluída: %d colunas (%s), %d linhas, %d I.E.(s): %s",
                len(self._nomes_colunas),
                self._nomes_colunas,
                len(self._dados_extraidos),
                len(self._lista_ies),
                self._lista_ies if len(self._lista_ies) <= 15 else self._lista_ies[:15] + [f"... +{len(self._lista_ies) - 15}"],
            )
        except FileNotFoundError:
            logger.error("Arquivo não encontrado: %s", caminho)
            messagebox.showerror("Erro", f"Arquivo não encontrado:\n{caminho}")
            return
        except Exception as e:
            logger.exception("Erro ao ler Excel: %s", e)
            messagebox.showerror("Erro", f"Erro ao ler o arquivo:\n{e}")
            return

        nome = self._caminho_arquivo.name
        try:
            self._lista_dados = obter_dados_para_dae(
                self._dados_extraidos,
                self._nome_para_indice,
                mes_ref=self._mes_ref,
                ano_ref=self._ano_ref,
            )
        except ValueError as e:
            messagebox.showerror("Erro", str(e))
            return

        qtd_total = len(self._lista_dados)
        qtd_exec, qtd_ign = _contar_executaveis_ignoradas(self._lista_dados)
        texto = (
            f"Arquivo: {nome}\n"
            f"Total: {qtd_total} I.E.(s) extraída(s). "
            f"Com valor TOTAL (executáveis): {qtd_exec}. Sem valor (ignoradas): {qtd_ign}."
        )
        self._label_arquivo.configure(text=texto, text_color="white")
        self._btn_executar.configure(state="normal" if qtd_exec > 0 else "disabled")
        self._btn_escolher.configure(state="normal" if qtd_total > 0 else "disabled")
        self._btn_ver_dados.configure(state="normal" if self._dados_extraidos else "disabled")
        self._label_status.configure(text="")
        print(
            f"[Upload] Arquivo: {nome} — {qtd_total} IEs | executáveis: {qtd_exec} | ignoradas: {qtd_ign}",
            flush=True,
        )

    def _ao_abrir_selecao_ies(self) -> None:
        """Abre janela de detalhamento e seleção de IEs; ao confirmar, executa só as selecionadas."""
        if not self._lista_dados:
            messagebox.showinfo("Aviso", "Nenhum dado carregado. Selecione o arquivo Excel primeiro.")
            return
        janela = JanelaSelecaoIEs(
            self,
            self._lista_dados,
            ao_executar=(self._iniciar_execucao_com_lista, False),
        )
        janela.focus()

    def _iniciar_execucao_com_lista(self, lista_selecionada: list[dict[str, object]]) -> None:
        """Chamado pela janela de seleção: executa automação apenas para a lista fornecida."""
        if self._em_execucao or not lista_selecionada:
            return
        self._em_execucao = True
        self._btn_upload.configure(state="disabled")
        self._btn_executar.configure(state="disabled")
        self._btn_escolher.configure(state="disabled")
        self._label_status.configure(
            text=f"Executando {len(lista_selecionada)} I.E.(s) selecionada(s). Veja o console.",
            text_color="orange",
        )

        def _ao_finalizar(ies_sucesso: list[str], ies_erro: list[tuple[str, str]]) -> None:
            self.after(0, lambda: self._finalizar_execucao(ies_sucesso, ies_erro))

        def _rodar() -> None:
            _executar_lote_em_background(
                lista_selecionada,
                headless=False,
                result_callback=_ao_finalizar,
            )

        thread = threading.Thread(target=_rodar, daemon=True)
        thread.start()

    def _ao_ver_dados(self) -> None:
        """Abre janela com os dados extraídos do Excel para verificação."""
        if not self._dados_extraidos or not self._nomes_colunas:
            messagebox.showinfo("Dados", "Nenhum dado extraído para exibir.")
            return
        logger.debug("Abrindo janela de visualização dos dados extraídos.")
        janela = ctk.CTkToplevel(self)
        janela.title("Dados extraídos do Excel")
        janela.geometry("900x500")
        janela.minsize(400, 300)

        # Cabeçalho
        ctk.CTkLabel(
            janela,
            text=f"Arquivo: {self._caminho_arquivo.name}  |  {len(self._dados_extraidos)} linhas  |  {len(self._nomes_colunas)} colunas",
            font=ctk.CTkFont(size=12),
        ).pack(pady=8)

        # Área de texto com dados em formato tabela
        texto = ctk.CTkTextbox(janela, font=ctk.CTkFont(family="Consolas", size=12), wrap="none")
        texto.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        # Montar tabela: linha de cabeçalho + linhas de dados
        colunas = self._nomes_colunas
        larguras = [max(len(str(col)), 4) for col in colunas]
        for i, col in enumerate(colunas):
            larguras[i] = max(
                larguras[i],
                max((len(str(linha.get(col, "") or "")) for linha in self._dados_extraidos[:100]), default=0),
            )
            larguras[i] = min(larguras[i], 30)

        def _cel(s: object, w: int) -> str:
            ss = "" if s is None else str(s).strip()
            return (ss[: w - 2] + "..") if len(ss) > w else ss.ljust(w)

        ie_key = _obter_chave_ie(self._nome_para_indice)

        linha_cab = " | ".join(_cel(col, larguras[i]) for i, col in enumerate(colunas))
        sep = "-+-".join("-" * w for w in larguras)
        linhas_txt = [linha_cab, sep]
        for linha in self._dados_extraidos:
            def _valor_celula(col: str, i: int) -> str:
                v = linha.get(col)
                if ie_key and col == ie_key and v is not None:
                    return _cel(_ie_para_exibicao(str(v)), larguras[i])
                return _cel(v, larguras[i])
            linhas_txt.append(" | ".join(_valor_celula(col, i) for i, col in enumerate(colunas)))
        texto.insert("1.0", "\n".join(linhas_txt))
        texto.configure(state="disabled")

    def _ao_executar_todas(self) -> None:
        """Executa automação para todas as IEs com valor TOTAL (executáveis)."""
        if self._em_execucao or not self._lista_dados:
            return
        # Apenas IEs executáveis (com valor TOTAL)
        lista_executaveis = [
            item for item in self._lista_dados
            if not _valor_total_invalido(item.get("valor_normal"))
        ]
        if not lista_executaveis:
            messagebox.showwarning(
                "Aviso",
                "Nenhuma I.E. com valor TOTAL para executar. Use 'Escolher IEs e executar' para ver o detalhamento.",
            )
            return

        self._em_execucao = True
        self._btn_upload.configure(state="disabled")
        self._btn_executar.configure(state="disabled")
        self._btn_escolher.configure(state="disabled")
        self._label_status.configure(
            text=f"Executando {len(lista_executaveis)} I.E.(s). Veja o console.",
            text_color="orange",
        )

        def _ao_finalizar(ies_sucesso: list[str], ies_erro: list[tuple[str, str]]) -> None:
            self.after(0, lambda: self._finalizar_execucao(ies_sucesso, ies_erro))

        def _rodar() -> None:
            _executar_lote_em_background(
                lista_executaveis,
                headless=False,
                result_callback=_ao_finalizar,
            )

        thread = threading.Thread(target=_rodar, daemon=True)
        thread.start()

    def _finalizar_execucao(self, ies_sucesso: list[str], ies_erro: list[tuple[str, str]]) -> None:
        self._em_execucao = False
        self._btn_upload.configure(state="normal")
        self._btn_executar.configure(state="normal")
        self._btn_escolher.configure(state="normal")
        total = len(ies_sucesso) + len(ies_erro)
        self._label_status.configure(
            text=f"Concluído: {len(ies_sucesso)}/{total} sucesso, {len(ies_erro)} erro.",
            text_color="orange" if ies_erro else "green",
        )
        _mostrar_janela_resultado(self, ies_sucesso, ies_erro)


def main() -> None:
    """Ponto de entrada da interface gráfica."""
    logger.info("Iniciando interface SEFAZ-CE.")
    app = AppFiliais()
    app.mainloop()
    logger.info("Interface encerrada.")


if __name__ == "__main__":
    main()
    sys.exit(0)
