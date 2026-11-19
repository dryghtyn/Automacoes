import subprocess
import importlib.util
import sys

DEPENDENCIAS = [
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
]

for nome_import, nome_pip in DEPENDENCIAS:
    if importlib.util.find_spec(nome_import) is None:
        print(f"Instalando {nome_pip}...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", nome_pip, "--quiet"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import ctypes
from pathlib import Path

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

try:
    import pandas as pd
except ImportError:
    pd = None

sys.path.insert(0, str(Path(__file__).parent))
import tema

_TEMA_ARG = ""
if "--tema" in sys.argv:
    try:
        _TEMA_ARG = sys.argv[sys.argv.index("--tema") + 1].strip(" \"")
    except (IndexError, ValueError):
        _TEMA_ARG = ""

TEMA = _TEMA_ARG if _TEMA_ARG else tema.carregar_tema()
DIR_RAIZ = Path(__file__).parent


def _cor_mutado():
    return "#9a9a9a" if TEMA != "dark" else "#8d8d8d"


def _cor_destaque():
    return "#005a9e" if TEMA != "dark" else "#3794ff"


TRATATIVAS = {
    "Duplicado": "Verificar se houve cupom cancelado reemitido, erro de digitação "
                 "ou lançamento duplicado no emissor.",
    "Salto de sequência": "Verificar cupons cancelados, estornos ou falha de numeração "
                          "no emissor que quebrou a sequência.",
}


class ComparadorNFCE:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de NFCe")
        self.root.geometry("860x620")
        self.root.minsize(720, 480)

        self.df = None
        self.colunas = []
        self.nome_origem = ""
        self.caminho_origem = ""

        self.serie_col = tk.StringVar()
        self.num_col = tk.StringVar()

        self._construir_ui()
        tema.aplicar_tema(self.root, TEMA)

    def _construir_ui(self):
        frame_top = ttk.LabelFrame(self.root, text="Planilha", padding=8)
        frame_top.pack(fill=tk.X, padx=8, pady=(8, 4))

        btn_linha = ttk.Frame(frame_top)
        btn_linha.pack(fill=tk.X)
        self.btn_carregar = ttk.Button(btn_linha, text="Carregar Planilha", command=self._carregar_planilha)
        self.btn_carregar.pack(side=tk.LEFT)
        self.lbl_planilha = ttk.Label(btn_linha, text="Nenhuma planilha carregada", foreground=_cor_mutado())
        self.lbl_planilha.pack(side=tk.LEFT, padx=10)

        cfg_linha = ttk.Frame(frame_top)
        cfg_linha.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(cfg_linha, text="Coluna Série:").pack(side=tk.LEFT)
        self.combo_serie = ttk.Combobox(cfg_linha, textvariable=self.serie_col, state="readonly", width=24)
        self.combo_serie.pack(side=tk.LEFT, padx=(4, 16))
        ttk.Label(cfg_linha, text="Coluna Nº Cupom:").pack(side=tk.LEFT)
        self.combo_num = ttk.Combobox(cfg_linha, textvariable=self.num_col, state="readonly", width=24)
        self.combo_num.pack(side=tk.LEFT, padx=(4, 0))

        self.btn_comparar = ttk.Button(frame_top, text="Comparar", command=self._comparar)
        self.btn_comparar.pack(side=tk.RIGHT)
        self.btn_comparar.state(["disabled"])

        frame_resumo = ttk.LabelFrame(self.root, text="Resumo da Comparação", padding=8)
        frame_resumo.pack(fill=tk.X, padx=8, pady=2)

        self._metricas = {}
        metricas = [
            ("series", "Séries"),
            ("registros", "Registros"),
            ("cupons", "Cupons"),
            ("ok", "OK"),
            ("duplicados", "Duplicados"),
            ("saltos", "Saltos"),
        ]
        for chave, rotulo in metricas:
            celula = ttk.Frame(frame_resumo)
            celula.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=6)
            ttk.Label(celula, text=rotulo, foreground=_cor_mutado(),
                      font=("Segoe UI", 9)).pack(anchor="center")
            self._metricas[chave] = ttk.Label(celula, text="0", font=("Segoe UI", 16, "bold"))
            self._metricas[chave].pack(anchor="center")

        frame_result = ttk.LabelFrame(self.root, text="Resultado da Comparação", padding=6)
        frame_result.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        cols = ("serie", "tipo", "valor", "linhas", "detalhe", "tratativa")
        self.tree = ttk.Treeview(frame_result, columns=cols, show="headings", height=12)
        self.tree.heading("serie", text="Série")
        self.tree.heading("tipo", text="Tipo de Falha")
        self.tree.heading("valor", text="Valor")
        self.tree.heading("linhas", text="Linha(s)")
        self.tree.heading("detalhe", text="Detalhe")
        self.tree.heading("tratativa", text="Tratativa")
        self.tree.column("serie", width=90, anchor=tk.CENTER)
        self.tree.column("tipo", width=120, anchor=tk.CENTER)
        self.tree.column("valor", width=90, anchor=tk.CENTER)
        self.tree.column("linhas", width=110, anchor=tk.CENTER)
        self.tree.column("detalhe", width=190, anchor=tk.W)
        self.tree.column("tratativa", width=360, anchor=tk.W)

        scroll_y = ttk.Scrollbar(frame_result, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        frame_status = ttk.Frame(self.root)
        frame_status.pack(fill=tk.X, padx=8, pady=(0, 2))
        self.lbl_indicador = ttk.Label(frame_status, text="●", foreground="#4caf50",
                                       font=("Segoe UI", 11))
        self.lbl_indicador.pack(side=tk.LEFT, padx=(0, 4))
        self.lbl_status = ttk.Label(frame_status, text="Aguardando planilha", foreground=_cor_mutado(),
                                    font=("Segoe UI", 9, "bold"))
        self.lbl_status.pack(side=tk.LEFT)
        self.progress = ttk.Progressbar(frame_status, mode="indeterminate", length=120)
        self.progress.pack(side=tk.LEFT, padx=12)
        self.progress.pack_forget()
        self.btn_salvar = ttk.Button(frame_status, text="Salvar Resultado + Marcar Original",
                                     command=self._salvar_resultado)
        self.btn_salvar.pack(side=tk.RIGHT)
        self.btn_salvar.state(["disabled"])

        frame_log = ttk.LabelFrame(self.root, text="Log", padding=4)
        frame_log.pack(fill=tk.BOTH, padx=8, pady=(0, 8))
        self.log = scrolledtext.ScrolledText(frame_log, height=5, state=tk.DISABLED, font=("Consolas", 9))
        self.log.pack(fill=tk.BOTH, expand=True)

    def _log(self, msg):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)

    def _carregar_planilha(self):
        caminho = filedialog.askopenfilename(title="Selecionar planilha",
                                             filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("Todos", "*.*")])
        if not caminho:
            return
        try:
            if caminho.lower().endswith(".csv"):
                self.df = pd.read_csv(caminho, encoding="utf-8-sig")
            else:
                self.df = pd.read_excel(caminho, engine="openpyxl")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler a planilha.\n{e}")
            self._log(f"Erro ao carregar: {e}")
            return
        self.caminho_origem = caminho
        self.nome_origem = os.path.basename(caminho)
        self.colunas = [str(c) for c in self.df.columns]
        self.serie_col.set("")
        self.num_col.set("")
        self.combo_serie["values"] = self.colunas
        self.combo_num["values"] = self.colunas
        self.lbl_planilha.configure(text=f"{self.nome_origem} ({len(self.df)} linhas)")
        self.btn_comparar.state(["!disabled"])
        self.btn_salvar.state(["disabled"])
        self._limpar_tree()
        self._reiniciar_resumo()
        self._set_status("Planilha carregada. Selecione as colunas e clique em Comparar.", "ok")
        self._log(f"Planilha carregada: {self.nome_origem} com {len(self.df)} linhas.")

    def _limpar_tree(self):
        self.tree.delete(*self.tree.get_children())

    def _set_status(self, texto, estado="ok"):
        cor = {"rodando": "#ff9800", "ok": "#4caf50", "erro": "#e53935"}.get(estado, "#4caf50")
        self.lbl_indicador.configure(foreground=cor)
        self.lbl_status.configure(text=texto)
        if estado == "rodando":
            self.progress.pack(side=tk.LEFT, padx=12)
            self.progress.start(12)
        else:
            self.progress.stop()
            self.progress.pack_forget()

    def _comparar(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Carregue uma planilha primeiro.")
            return
        if not self.serie_col.get() or not self.num_col.get():
            messagebox.showwarning("Aviso", "Selecione a coluna de Série e a coluna de Nº do Cupom.")
            return
        self._set_status("Comparando...", "rodando")
        self._reiniciar_resumo()
        self.btn_comparar.state(["disabled"])
        thread = threading.Thread(target=self._rodar_comparacao, daemon=True)
        thread.start()

    def _reiniciar_resumo(self):
        for chave in self._metricas:
            self._metricas[chave].configure(text="0")

    def _rodar_comparacao(self):
        try:
            resultados, marcacoes, totais, resumo, log_linhas = self._analisar()
            self._mostrar_resultados(resultados)
            self._resultado = resultados
            self._marcacoes = marcacoes
            self._totais = totais
            self._resumo = resumo
            self.root.after(0, lambda: self.btn_salvar.state(["!disabled"]))
            self.root.after(0, lambda: self.btn_comparar.state(["!disabled"]))
            total = len(resultados)
            self.root.after(0, lambda: self._atualizar_resumo(resumo))
            for linha in log_linhas:
                self.root.after(0, lambda l=linha: self._log(l))
            self.root.after(0, lambda: self._set_status(
                f"Comparação concluída: {total} falha(s) encontrada(s) em "
                f"{len(totais)} série(s).", "ok"))
            self.root.after(0, lambda: self._log(f"Comparação concluída. Falhas: {total}."))
        except Exception as e:
            self.root.after(0, lambda: self.btn_comparar.state(["!disabled"]))
            self.root.after(0, lambda: self._set_status("Erro na comparação.", "erro"))
            self.root.after(0, lambda: messagebox.showerror("Erro", f"Erro na comparação.\n{e}"))
            self.root.after(0, lambda: self._log(f"Erro: {e}"))

    def _atualizar_resumo(self, resumo):
        mapeamento = {
            "series": "series",
            "registros": "registros",
            "cupons": "cupons",
            "ok": "ok",
            "duplicados": "duplicados",
            "saltos": "saltos",
        }
        for chave_widget, chave_resumo in mapeamento.items():
            self._metricas[chave_widget].configure(text=str(resumo.get(chave_resumo, 0)))
        if resumo.get("duplicados") or resumo.get("saltos"):
            self._metricas["ok"].configure(foreground=_cor_mutado())
        else:
            self._metricas["ok"].configure(foreground=_cor_destaque())

    def _analisar(self):
        df = self.df.copy()
        serie_col = self.serie_col.get()
        num_col = self.num_col.get()

        resultados = []
        marcacoes = {}
        log_linhas = []

        for idx in df.index:
            marcacoes[idx] = []

        def falta(serie, valor, linhas, tipo, detalhe):
            resultados.append({"serie": serie, "tipo": tipo, "valor": valor,
                               "linhas": linhas, "detalhe": detalhe,
                               "tratativa": TRATATIVAS.get(tipo, "")})
            for linha in linhas:
                if tipo not in marcacoes[linha]:
                    marcacoes[linha].append(tipo)

        resumo = {"series": 0, "registros": 0, "cupons": 0,
                  "ok": 0, "duplicados": 0, "saltos": 0}

        grupos = df.groupby(serie_col, dropna=False)
        for serie, sub in grupos:
            reg = int(sub.shape[0])
            numeros = pd.to_numeric(sub[num_col], errors="coerce").dropna()
            resumo["series"] += 1
            resumo["registros"] += reg

            contagem = sub[num_col].value_counts(dropna=False)
            duplicados = contagem[contagem > 1]
            n_dup = int(len(duplicados))
            resumo["duplicados"] += n_dup
            for valor, qtd in duplicados.items():
                linhas = [int(i) for i in sub.index[sub[num_col] == valor].tolist()]
                falta(serie, valor, linhas, "Duplicado", f"{qtd} ocorrências")

            n_saltos = 0
            if not numeros.empty:
                numeros_i = numeros.astype(int)
                minimo = int(numeros_i.min())
                maximo = int(numeros_i.max())
                presentes = set(numeros_i.tolist())
                faltantes = []
                for n in range(minimo + 1, maximo):
                    if n not in presentes:
                        faltantes.append(n)
                        n_saltos += 1
                        tipo = "Salto de sequência"
                        detalhe = f"Falta o número {n} da sequência ({minimo} a {maximo} na série)"
                        resultados.append({"serie": serie, "tipo": tipo, "valor": n,
                                           "linhas": [], "detalhe": detalhe,
                                           "tratativa": TRATATIVAS.get(tipo, "")})
            resumo["saltos"] += n_saltos

            distinct = int(numeros.nunique()) if not numeros.empty else 0
            resumo["cupons"] += distinct

            parte = (f"Série {serie}: {reg} registro(s), {distinct} cupom(ns) distintos"
                     f" [{minimo} a {maximo}]" if not numeros.empty
                     else f"Série {serie}: {reg} registro(s)")
            if n_dup and n_saltos:
                parte += f" → {n_dup} duplicado(s) e {n_saltos} salto(s)"
            elif n_dup:
                parte += f" → {n_dup} duplicado(s)"
            elif n_saltos:
                parte += f" → {n_saltos} salto(s)"
            else:
                parte += " → OK"
            log_linhas.append(parte)

        resumo["ok"] = resumo["registros"] - sum(1 for ts in marcacoes.values() if ts)

        # Totais de cupons por série
        totais = []
        for serie, sub in grupos:
            numeros = pd.to_numeric(sub[num_col], errors="coerce").dropna()
            totais.append({
                "serie": serie,
                "registros": int(sub.shape[0]),
                "cupons_distintos": int(numeros.nunique()) if not numeros.empty else 0,
                "min": int(numeros.min()) if not numeros.empty else None,
                "max": int(numeros.max()) if not numeros.empty else None,
            })

        return resultados, marcacoes, totais, resumo, log_linhas

    def _mostrar_resultados(self, resultados):
        self._limpar_tree()
        for r in resultados:
            linhas_txt = ", ".join(str(l + 2) for l in r["linhas"]) if r["linhas"] else "—"
            self.tree.insert("", "end",
                             values=(str(r["serie"]), r["tipo"], str(r["valor"]),
                                     linhas_txt, r["detalhe"], r["tratativa"]))

    def _salvar_resultado(self):
        if not hasattr(self, "_resultado") or not self._resultado:
            messagebox.showwarning("Aviso", "Nenhum resultado para salvar.")
            return
        padrao = f"{Path(self.nome_origem).stem}_resultado_compare.xlsx"
        caminho = filedialog.asksaveasfilename(title="Salvar resultado",
                                               defaultextension=".xlsx",
                                               initialfile=padrao,
                                               filetypes=[("Excel", "*.xlsx")])
        if not caminho:
            return
        try:
            linhas_export = []
            for r in self._resultado:
                linhas_txt = ", ".join(str(l + 2) for l in r["linhas"]) if r["linhas"] else "—"
                linhas_export.append({
                    "Série": r["serie"],
                    "Tipo de Falha": r["tipo"],
                    "Valor": r["valor"],
                    "Linha(s) da planilha": linhas_txt,
                    "Detalhe": r["detalhe"],
                    "Tratativa": r["tratativa"],
                })
            df_resultado = pd.DataFrame(linhas_export)

            df_marcado = self.df.copy()
            rotulos_falha = {"Duplicado": "Duplicado", "Salto de sequência": "Salto"}
            col_falha = "FALHA"
            df_marcado[col_falha] = ""
            col_situacao = "SITUAÇÃO"
            df_marcado[col_situacao] = "OK"
            for idx, tipos in self._marcacoes.items():
                if tipos:
                    rotulos = ", ".join(rotulos_falha.get(t, t) for t in tipos)
                    df_marcado.at[idx, col_falha] = rotulos
                    df_marcado.at[idx, col_situacao] = rotulos

            totais_export = []
            for t in getattr(self, "_totais", []):
                faixa = f"{t['min']} a {t['max']}" if t["min"] is not None else "—"
                totais_export.append({
                    "Série": t["serie"],
                    "Total de Cupons (distintos)": t["cupons_distintos"],
                    "Total de Registros": t["registros"],
                    "Faixa de numeração": faixa,
                })

            with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
                df_resultado.to_excel(writer, sheet_name="Falhas", index=False)
                if totais_export:
                    pd.DataFrame(totais_export).to_excel(writer, sheet_name="Totais por Série", index=False)
                df_marcado.to_excel(writer, sheet_name="Planilha Marcada", index=False)

            messagebox.showinfo("Salvo", f"Resultado salvo em:\n{caminho}")
            self._log(f"Resultado salvo em {caminho}")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar.\n{e}")
            self._log(f"Erro ao salvar: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ComparadorNFCE(root)
    root.mainloop()
