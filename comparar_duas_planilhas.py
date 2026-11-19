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
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import time
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


def _cor_mutado():
    return "#9a9a9a" if TEMA != "dark" else "#8d8d8d"


def _cor_destaque():
    return "#005a9e" if TEMA != "dark" else "#3794ff"


class CompararDuasPlanilhas:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparar Duas Planilhas de NFCe")
        self.root.geometry("900x660")
        self.root.minsize(780, 520)

        self.df_a = None
        self.df_b = None
        self.nome_a = ""
        self.nome_b = ""

        self._resultado = None

        self._construir_ui()
        tema.aplicar_tema(self.root, TEMA)

    def _construir_ui(self):
        frame_modo = ttk.LabelFrame(self.root, text="Campos de Comparação (marque as exigências)", padding=8)
        frame_modo.pack(fill=tk.X, padx=8, pady=(8, 4))

        self.ck_chave = ttk.Checkbutton(frame_modo, text="Chave de Acesso",
                                        command=lambda: self._marcar_filtro("chave"))
        self.ck_chave.invoke()
        self.ck_chave.pack(side=tk.LEFT, padx=(0, 12))
        self.ck_serienum = ttk.Checkbutton(frame_modo, text="Série + Número do Cupom",
                                           command=lambda: self._marcar_filtro("serienum"))
        self.ck_serienum.pack(side=tk.LEFT)

        frame_plan = ttk.LabelFrame(self.root, text="Planilhas", padding=8)
        frame_plan.pack(fill=tk.X, padx=8, pady=2)

        self._montar_lado_planilha(frame_plan, "A")
        self._montar_lado_planilha(frame_plan, "B")

        frame_acao = ttk.Frame(self.root)
        frame_acao.pack(fill=tk.X, padx=8, pady=4)
        self.btn_comparar = ttk.Button(frame_acao, text="Comparar", command=self._comparar)
        self.btn_comparar.pack(side=tk.LEFT)
        self.lbl_info = ttk.Label(frame_acao, text="", foreground=_cor_mutado())
        self.lbl_info.pack(side=tk.LEFT, padx=12)

        frame_result = ttk.LabelFrame(self.root, text="Resultado da Comparação", padding=6)
        frame_result.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        self.notebook = ttk.Notebook(frame_result)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self._aba_resultado("soa", "Somente na Planilha A")
        self._aba_resultado("sob", "Somente na Planilha B")
        self._aba_resultado("ambas", "Existem nas Duas")
        self._aba_resultado("dup", "Duplicados")

        self.tree_soa, self.lbl_soa = self._criar_tree(self.f_soa)
        self.tree_sob, self.lbl_sob = self._criar_tree(self.f_sob)
        self.tree_ambas, self.lbl_ambas = self._criar_tree(self.f_ambas)
        self.tree_dup, self.lbl_dup = self._criar_tree(self.f_dup, com_linhas=True)

        frame_status = ttk.Frame(self.root)
        frame_status.pack(fill=tk.X, padx=8, pady=(0, 2))
        self.lbl_indicador = ttk.Label(frame_status, text="●", foreground=_cor_mutado(),
                                       font=("Segoe UI", 11))
        self.lbl_indicador.pack(side=tk.LEFT, padx=(0, 4))
        self.lbl_status = ttk.Label(frame_status, text="Carregue as duas planilhas e selecione as colunas.",
                                    foreground=_cor_mutado(), font=("Segoe UI", 9, "bold"))
        self.lbl_status.pack(side=tk.LEFT)
        self.btn_exportar = ttk.Button(frame_status, text="Exportar Resultado",
                                       command=self._exportar)
        self.btn_exportar.pack(side=tk.RIGHT)
        self.btn_exportar.state(["disabled"])

        frame_log = ttk.LabelFrame(self.root, text="Log", padding=4)
        frame_log.pack(fill=tk.BOTH, padx=8, pady=(0, 8))
        self.log = scrolledtext.ScrolledText(frame_log, height=5, state=tk.DISABLED, font=("Consolas", 9))
        self.log.pack(fill=tk.BOTH, expand=True)

    def _montar_lado_planilha(self, parent, lado):
        frame = ttk.LabelFrame(parent, text=f"Planilha {lado}", padding=6)
        frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

        linha_btn = ttk.Frame(frame)
        linha_btn.pack(fill=tk.X)
        btn = ttk.Button(linha_btn, text=f"Carregar Planilha {lado}",
                         command=lambda l=lado: self._carregar_planilha(l))
        btn.pack(side=tk.LEFT)
        lbl = ttk.Label(linha_btn, text="Nenhuma", foreground=_cor_mutado())
        lbl.pack(side=tk.LEFT, padx=8)

        cfg = ttk.Frame(frame)
        cfg.pack(fill=tk.X, pady=(6, 0))

        if lado == "A":
            self.lbl_arquivo_a = lbl
            self.cfg_chave_a = ttk.Frame(cfg)
            self.cfg_chave_a.pack(fill=tk.X)
            ttk.Label(self.cfg_chave_a, text="Coluna Chave de Acesso:").pack(side=tk.LEFT)
            self.combo_chave_a = ttk.Combobox(self.cfg_chave_a, state="readonly", width=18)
            self.combo_chave_a.pack(side=tk.LEFT, padx=(4, 0))
            self.cfg_serie_a = ttk.Frame(cfg)
            self.cfg_serie_a.pack(fill=tk.X, pady=(2, 0))
            ttk.Label(self.cfg_serie_a, text="Coluna Série:").pack(side=tk.LEFT)
            self.combo_serie_a = ttk.Combobox(self.cfg_serie_a, state="readonly", width=18)
            self.combo_serie_a.pack(side=tk.LEFT, padx=(4, 8))
            self.cfg_num_a = ttk.Frame(cfg)
            self.cfg_num_a.pack(fill=tk.X, pady=(2, 0))
            ttk.Label(self.cfg_num_a, text="Coluna Nº Cupom:").pack(side=tk.LEFT)
            self.combo_num_a = ttk.Combobox(self.cfg_num_a, state="readonly", width=18)
            self.combo_num_a.pack(side=tk.LEFT, padx=(4, 0))
        else:
            self.lbl_arquivo_b = lbl
            self.cfg_chave_b = ttk.Frame(cfg)
            self.cfg_chave_b.pack(fill=tk.X)
            ttk.Label(self.cfg_chave_b, text="Coluna Chave de Acesso:").pack(side=tk.LEFT)
            self.combo_chave_b = ttk.Combobox(self.cfg_chave_b, state="readonly", width=18)
            self.combo_chave_b.pack(side=tk.LEFT, padx=(4, 0))
            self.cfg_serie_b = ttk.Frame(cfg)
            self.cfg_serie_b.pack(fill=tk.X, pady=(2, 0))
            ttk.Label(self.cfg_serie_b, text="Coluna Série:").pack(side=tk.LEFT)
            self.combo_serie_b = ttk.Combobox(self.cfg_serie_b, state="readonly", width=18)
            self.combo_serie_b.pack(side=tk.LEFT, padx=(4, 8))
            self.cfg_num_b = ttk.Frame(cfg)
            self.cfg_num_b.pack(fill=tk.X, pady=(2, 0))
            ttk.Label(self.cfg_num_b, text="Coluna Nº Cupom:").pack(side=tk.LEFT)
            self.combo_num_b = ttk.Combobox(self.cfg_num_b, state="readonly", width=18)
            self.combo_num_b.pack(side=tk.LEFT, padx=(4, 0))

    def _criar_tree(self, parent, com_linhas=False):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True)
        lbl = ttk.Label(frame, text="Sem resultados.", foreground=_cor_mutado())
        lbl.pack(anchor="w", pady=(2, 4))
        if com_linhas:
            cols = ("chave", "serie", "num", "qtd", "linhas", "fonte")
        else:
            cols = ("chave", "serie", "num", "fonte")
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=12)
        tree.heading("chave", text="Chave de Acesso")
        tree.heading("serie", text="Série")
        tree.heading("num", text="Nº Cupom")
        if com_linhas:
            tree.heading("qtd", text="Ocorrências")
            tree.heading("linhas", text="Linha(s)")
        tree.heading("fonte", text="Presente em")
        tree.column("chave", width=300, anchor=tk.W)
        tree.column("serie", width=100, anchor=tk.CENTER)
        tree.column("num", width=110, anchor=tk.CENTER)
        if com_linhas:
            tree.column("qtd", width=90, anchor=tk.CENTER)
            tree.column("linhas", width=160, anchor=tk.W)
        tree.column("fonte", width=200, anchor=tk.CENTER)
        scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        return tree, lbl

    def _aba_resultado(self, nome, titulo):
        aba = ttk.Frame(self.notebook)
        self.notebook.add(aba, text=titulo)
        setattr(self, f"f_{nome}", aba)

    def _marcar_filtro(self, chave):
        pass

    def _log(self, msg):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)

    def _set_status(self, texto, cor=None):
        self.lbl_indicador.configure(foreground=cor or _cor_mutado())
        self.lbl_status.configure(text=texto)

    def _carregar_planilha(self, lado):
        caminho = filedialog.askopenfilename(title=f"Selecionar Planilha {lado}",
                                             filetypes=[("Excel", "*.xlsx *.xls"),
                                                        ("CSV", "*.csv"),
                                                        ("Todos", "*.*")])
        if not caminho:
            return
        try:
            if caminho.lower().endswith(".csv"):
                df = pd.read_csv(caminho, dtype=str)
            else:
                df = pd.read_excel(caminho, engine="openpyxl", dtype=str)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler a planilha {lado}.\n{e}")
            self._log(f"Erro ao carregar {lado}: {e}")
            return
        colunas = [str(c) for c in df.columns]
        if lado == "A":
            self.df_a = df
            self.nome_a = os.path.basename(caminho)
            self.combo_serie_a["values"] = colunas
            self.combo_num_a["values"] = colunas
            self.combo_chave_a["values"] = colunas
            self.lbl_arquivo_a.configure(text=f"{self.nome_a} ({len(df)} linhas)")
        else:
            self.df_b = df
            self.nome_b = os.path.basename(caminho)
            self.combo_serie_b["values"] = colunas
            self.combo_num_b["values"] = colunas
            self.combo_chave_b["values"] = colunas
            self.lbl_arquivo_b.configure(text=f"{self.nome_b} ({len(df)} linhas)")
        self._log(f"Planilha {lado} carregada: {os.path.basename(caminho)} com {len(df)} linhas.")
        self._verificar_prontas()

    def _verificar_prontas(self):
        if self.df_a is not None and self.df_b is not None:
            self._set_status("Linhas carregadas. Selecione as colunas e clique em Comparar.", "#ff9800")

    def _filtros_selecionados(self):
        selecionados = []
        if self.ck_chave.instate(["selected"]):
            selecionados.append("chave")
        if self.ck_serienum.instate(["selected"]):
            selecionados.append("serienum")
        return selecionados

    def _coluna_para(self, lado, filtro):
        if lado == "A":
            return {"chave": self.combo_chave_a, "serienum": self.combo_serie_a}[filtro].get()
        return {"chave": self.combo_chave_b, "serienum": self.combo_serie_b}[filtro].get()

    def _comparar(self):
        if self.df_a is None or self.df_b is None:
            messagebox.showwarning("Aviso", "Carregue as duas planilhas primeiro.")
            return
        filtros = self._filtros_selecionados()
        if not filtros:
            messagebox.showwarning("Aviso", "Marque ao menos um campo de comparação (Chave de Acesso ou Série + Nº Cupom).")
            return
        faltando = []
        for f in filtros:
            if f == "serienum":
                for label, combos in (("Série", (self.combo_serie_a, self.combo_serie_b)),
                                      ("Nº Cupom", (self.combo_num_a, self.combo_num_b))):
                    if not combos[0].get():
                        faltando.append(f"Planilha A - {label}")
                    if not combos[1].get():
                        faltando.append(f"Planilha B - {label}")
            else:
                if not self._coluna_para("A", f):
                    faltando.append("Planilha A - Chave de Acesso")
                if not self._coluna_para("B", f):
                    faltando.append("Planilha B - Chave de Acesso")
        if faltando:
            messagebox.showwarning("Aviso",
                "Selecione as colunas para os filtros marcados:\n" + "\n".join(faltando))
            return
        vazias = self._verificar_colunas_vazias(filtros)
        if vazias:
            messagebox.showwarning("Colunas sem dados",
                "As colunas a seguir não têm nenhum valor preenchido e farão todas as "
                "linhas serem ignoradas na comparação:\n\n" + "\n".join(vazias) +
                "\n\nCorrija a seleção de colunas antes de comparar.")
            return
        self.btn_comparar.state(["disabled"])
        self.btn_exportar.state(["disabled"])
        self._set_status("Comparando...", "#ff9800")
        thread = threading.Thread(target=self._rodar_comparacao, daemon=True)
        thread.start()

    def _rodar_comparacao(self):
        try:
            resultado = self._analisar()
            self._resultado = resultado
            self.root.after(0, self._exibir_resultado)
            self.root.after(0, lambda: self.btn_exportar.state(["!disabled"]))
            self.root.after(0, lambda: self._set_status("Comparação concluída.", "#4caf50"))
        except Exception as e:
            self.root.after(0, lambda: self._set_status("Erro na comparação.", "#e53935"))
            self.root.after(0, lambda: messagebox.showerror("Erro", f"Erro na comparação.\n{e}"))
            self.root.after(0, lambda: self._log(f"Erro: {e}"))
        finally:
            self.root.after(0, lambda: self.btn_comparar.state(["!disabled"]))

    def _resolver_coluna(self, df, texto):
        if not texto:
            return None
        if texto in df.columns:
            return texto
        for c in df.columns:
            if str(c) == texto:
                return c
        return texto

    def _valor_limpo(self, row, col):
        if col is None or col == "":
            return ""
        v = row.get(col, "")
        if v is None:
            return ""
        s = str(v).strip()
        if s.lower() in ("nan", "none", "nat"):
            return ""
        return s

    def _verificar_colunas_vazias(self, filtros):
        avisos = []
        for lado, df in (("A", self.df_a), ("B", self.df_b)):
            if "chave" in filtros:
                col = self._resolver_coluna(df, self._coluna_para(lado, "chave"))
                if col and not any(self._valor_limpo(r, col) for _, r in df.iterrows()):
                    avisos.append(f"Planilha {lado} - Chave de Acesso (coluna '{self._coluna_para(lado, 'chave')}')")
            if "serienum" in filtros:
                serie = self._resolver_coluna(df, self._coluna_para(lado, "serienum"))
                num = self._resolver_coluna(df, self.combo_num_a.get() if lado == "A" else self.combo_num_b.get())
                if serie and not any(self._valor_limpo(r, serie) for _, r in df.iterrows()):
                    avisos.append(f"Planilha {lado} - Série (coluna '{self._coluna_para(lado, 'serienum')}')")
                if num and not any(self._valor_limpo(r, num) for _, r in df.iterrows()):
                    avisos.append(f"Planilha {lado} - Nº Cupom (coluna '{self.combo_num_a.get() if lado == 'A' else self.combo_num_b.get()}')")
        return avisos

    def _chave_filtros(self, row, lado, filtros):
        df = self.df_a if lado == "A" else self.df_b
        partes = []
        for f in filtros:
            if f == "serienum":
                serie = self._valor_limpo(row, self._resolver_coluna(df, self._coluna_para(lado, "serienum")))
                num = self._valor_limpo(row, self._resolver_coluna(df, self.combo_num_a.get() if lado == "A" else self.combo_num_b.get()))
                if not serie or not num:
                    return None
                partes.append(serie)
                partes.append(num)
            else:
                col = self._resolver_coluna(df, self._coluna_para(lado, f))
                val = self._valor_limpo(row, col)
                if not val:
                    return None
                partes.append(val)
        return "|".join(partes)

    def _analisar(self):
        filtros = self._filtros_selecionados()

        set_a = {}
        ignoradas_a = []
        for idx, row in self.df_a.iterrows():
            k = self._chave_filtros(row, "A", filtros)
            if k:
                set_a.setdefault(k, []).append((idx, row))
            else:
                ignoradas_a.append(idx + 2)
        set_b = {}
        ignoradas_b = []
        for idx, row in self.df_b.iterrows():
            k = self._chave_filtros(row, "B", filtros)
            if k:
                set_b.setdefault(k, []).append((idx, row))
            else:
                ignoradas_b.append(idx + 2)

        soa = []
        for k in sorted(set_a):
            if k not in set_b:
                for idx, row in set_a[k]:
                    item = self._extrair(row, "A")
                    item["fonte"] = "Somente na Planilha A"
                    soa.append(item)
        sob = []
        for k in sorted(set_b):
            if k not in set_a:
                for idx, row in set_b[k]:
                    item = self._extrair(row, "B")
                    item["fonte"] = "Somente na Planilha B"
                    sob.append(item)
        ambas = []
        for k in sorted(set_a):
            if k in set_b:
                for idx, row in set_a[k]:
                    item = self._extrair(row, "A")
                    item["fonte"] = "Nas duas"
                    ambas.append(item)

        desc = "+".join(filtros)
        self._log(f"Comparação por [{desc}]: A={len(set_a)} chaves, B={len(set_b)} chaves.")
        if ignoradas_a:
            self._log(f"Aviso: {len(ignoradas_a)} linha(s) da Planilha A ignoradas (chave vazia/incompleta): {ignoradas_a[:20]}")
        if ignoradas_b:
            self._log(f"Aviso: {len(ignoradas_b)} linha(s) da Planilha B ignoradas (chave vazia/incompleta): {ignoradas_b[:20]}")

        dup = []
        for nome, chaves, lado in (("A", set_a, "A"), ("B", set_b, "B")):
            for k in sorted(chaves):
                ocorrencias = chaves[k]
                if len(ocorrencias) > 1:
                    item = self._extrair(ocorrencias[0][1], lado)
                    linhas = [str(idx + 2) for idx, _ in ocorrencias]
                    item["fonte"] = f"Duplicado na Planilha {nome}"
                    item["qtd"] = len(ocorrencias)
                    item["linhas"] = ", ".join(linhas)
                    dup.append(item)
            dups = {k: v for k, v in chaves.items() if len(v) > 1}
            if dups:
                self._log(f"Aviso: {len(dups)} chave(s) duplicada(s) na Planilha {nome}.")
        return {"soa": soa, "sob": sob, "ambas": ambas, "dup": dup}

    def _extrair(self, row, lado):
        df = self.df_a if lado == "A" else self.df_b
        num_col = self.combo_num_a.get() if lado == "A" else self.combo_num_b.get()
        return {
            "chave": self._valor_limpo(row, self._resolver_coluna(df, self._coluna_para(lado, "chave"))),
            "serie": self._valor_limpo(row, self._resolver_coluna(df, self._coluna_para(lado, "serienum"))),
            "num": self._valor_limpo(row, self._resolver_coluna(df, num_col)),
        }

    def _exibir_resultado(self):
        r = self._resultado
        for chave in ("soa", "sob", "ambas"):
            tree = {"soa": self.tree_soa, "sob": self.tree_sob,
                    "ambas": self.tree_ambas}[chave]
            lbl = {"soa": self.lbl_soa, "sob": self.lbl_sob,
                   "ambas": self.lbl_ambas}[chave]
            tree.delete(*tree.get_children())
            linhas = r[chave]
            for item in linhas:
                tree.insert("", "end",
                            values=(item["chave"], item["serie"], item["num"], item["fonte"]))
            lbl.configure(text=f"{len(linhas)} registro(s).")

        self.tree_dup.delete(*self.tree_dup.get_children())
        dups = r.get("dup", [])
        for item in dups:
            self.tree_dup.insert("", "end",
                                 values=(item["chave"], item["serie"], item["num"],
                                         item["qtd"], item["linhas"], item["fonte"]))
        self.lbl_dup.configure(text=f"{len(dups)} chave(s) duplicada(s).")

        self._log(f"Resultado: só A={len(r['soa'])}, só B={len(r['sob'])}, "
                  f"ambas={len(r['ambas'])}, duplicados={len(dups)}.")

    def _exportar(self):
        if not self._resultado:
            messagebox.showwarning("Aviso", "Nenhum resultado para exportar.")
            return
        stem_a = Path(self.nome_a).stem if self.nome_a else "A"
        stem_b = Path(self.nome_b).stem if self.nome_b else "B"
        padrao = f"comparacao_{stem_a}_vs_{stem_b}.xlsx"
        caminho = filedialog.asksaveasfilename(title="Exportar Resultado",
                                               defaultextension=".xlsx",
                                               initialfile=padrao,
                                               filetypes=[("Excel", "*.xlsx")])
        if not caminho:
            return
        try:
            df_soa = pd.DataFrame(self._resultado["soa"])
            df_sob = pd.DataFrame(self._resultado["sob"])
            df_ambas = pd.DataFrame(self._resultado["ambas"])
            df_dup = pd.DataFrame(self._resultado.get("dup", []))
            with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
                df_soa.to_excel(writer, sheet_name="Somente na Planilha A", index=False)
                df_sob.to_excel(writer, sheet_name="Somente na Planilha B", index=False)
                df_ambas.to_excel(writer, sheet_name="Existem nas Duas", index=False)
                df_dup.to_excel(writer, sheet_name="Duplicados", index=False)
            messagebox.showinfo("Salvo", f"Resultado salvo em:\n{caminho}")
            self._log(f"Resultado exportado para {caminho}")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível exportar.\n{e}")
            self._log(f"Erro ao exportar: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = CompararDuasPlanilhas(root)
    root.mainloop()
