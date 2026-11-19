import subprocess
import importlib.util
import sys

DEPENDENCIAS = [
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("pynput", "pynput"),
    ("mss", "mss"),
]

for nome_import, nome_pip in DEPENDENCIAS:
    if importlib.util.find_spec(nome_import) is None:
        print(f"Instalando {nome_pip}...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", nome_pip, "--quiet"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )

import tkinter as tk
from tkinter import ttk, messagebox
import ctypes
import json
import os
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

def _elevar_admin():
    if ctypes.windll.shell32.IsUserAnAdmin():
        return
    script = os.path.abspath(sys.argv[0])
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, f'"{script}"', None, 1
    )
    sys.exit()

_elevar_admin()


sys.path.insert(0, str(Path(__file__).parent))
import tema

DIR_AUTOMACOES = Path(__file__).parent / "automacoes"
SCRIPT_AUTOMACAO = Path(__file__).parent / "automacao.py"
PYTHON = sys.executable


def _carregar_automacoes():
    automacoes = []
    if not DIR_AUTOMACOES.exists():
        return automacoes
    for pasta in sorted(DIR_AUTOMACOES.iterdir()):
        if not pasta.is_dir():
            continue
        nome = pasta.name
        descricao = ""
        script = "automacao.py"
        args = ["--config", "{pasta}/config_passos.json"]
        meta = pasta / "automacao.json"
        if meta.exists():
            try:
                with open(meta, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                nome = dados.get("nome") or nome
                descricao = dados.get("descricao", "")
                script = dados.get("script", script)
                args = dados.get("args", args)
            except Exception:
                pass
        if not (Path(__file__).parent / script).exists():
            continue
        pasta_str = str(pasta).replace("\\", "/")
        args = [str(a).replace("{pasta}", pasta_str) for a in args]
        automacoes.append({
            "nome": nome,
            "descricao": descricao,
            "pasta": pasta,
            "script": script,
            "args": args,
        })
    return automacoes


class SeletorAutomacoes:
    def __init__(self, root):
        self.root = root
        self.root.title("Selecionar Automação")
        self.root.geometry("520x420")
        self.root.minsize(420, 320)

        frame = ttk.Frame(root, padding=(16, 12))
        frame.pack(fill=tk.BOTH, expand=True)

        cabecalho = ttk.Frame(frame)
        cabecalho.pack(fill=tk.X)

        ttk.Label(cabecalho, text="Qual automação deseja executar?",
                  font=("Segoe UI", 13, "bold")).pack(side=tk.LEFT)

        tema_atual = tema.carregar_tema()
        frame_tema = ttk.Frame(cabecalho)
        frame_tema.pack(side=tk.RIGHT)
        ttk.Label(frame_tema, text="Tema:").pack(side=tk.LEFT, padx=(0, 4))
        self.cb_tema = ttk.Combobox(frame_tema, state="readonly",
                                    values=["Claro", "Escuro"], width=8)
        self.cb_tema.set("Escuro" if tema_atual == "dark" else "Claro")
        self.cb_tema.bind("<<ComboboxSelected>>", self._trocar_tema)
        self.cb_tema.pack(side=tk.LEFT)

        botoes = ttk.Frame(frame)
        botoes.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 0))

        ttk.Button(botoes, text="Abrir Automação",
                   command=self._abrir_selecionada).pack(side=tk.RIGHT)
        ttk.Button(botoes, text="Fechar",
                   command=root.destroy).pack(side=tk.RIGHT, padx=8)

        self.lista = ttk.Treeview(frame, columns=("desc",), show="tree", height=10)
        self.lista.heading("#0", text="Automação")
        self.lista.column("#0", width=300, minwidth=180)
        self.lista.heading("desc", text="Descrição")
        self.lista.column("desc", width=180, minwidth=120)

        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.lista.yview)
        self.lista.configure(yscrollcommand=scroll.set)

        self.lista.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=12)
        scroll.pack(side=tk.LEFT, fill=tk.Y, pady=12)

        self.lista.bind("<Double-1>", lambda e: self._abrir_selecionada())
        self.lista.bind("<Return>", lambda e: self._abrir_selecionada())

        self._preencher_lista()

        if self.lista.get_children():
            self.lista.selection_set(self.lista.get_children()[0])

        self._aplicar_tema_atual()

    def _aplicar_tema_atual(self):
        tema_atual = tema.carregar_tema()
        tema.aplicar_tema(self.root, tema_atual)

    def _trocar_tema(self, event=None):
        escolha = self.cb_tema.get()
        tema_novo = "dark" if escolha == "Escuro" else "light"
        tema.salvar_tema(tema_novo)
        self._aplicar_tema_atual()

    def _preencher_lista(self):
        self.automacoes = _carregar_automacoes()
        for i, auto in enumerate(self.automacoes):
            nome = auto["nome"]
            if i == 0:
                nome += "  (Padrão)"
            self.lista.insert("", "end", iid=str(i),
                              text=nome, values=(auto["descricao"],))
        if not self.automacoes:
            self.lista.insert("", "end", text="(Nenhuma automação encontrada)",
                              values=("Crie uma pasta em automacoes/",))

    def _obter_selecionada(self):
        sel = self.lista.selection()
        if not sel:
            return None
        iid = sel[0]
        if not iid.isdigit() or int(iid) >= len(self.automacoes):
            return None
        return self.automacoes[int(iid)]

    def _abrir_selecionada(self):
        if not self.automacoes:
            messagebox.showwarning("Aviso",
                "Nenhuma automação encontrada em:\n" + str(DIR_AUTOMACOES))
            return
        sel = self._obter_selecionada()
        if sel is None:
            messagebox.showinfo("Selecione", "Selecione uma automação na lista.")
            return
        try:
            tema_novo = "dark" if self.cb_tema.get() == "Escuro" else "light"
            script_path = (Path(__file__).parent / sel["script"]).resolve()
            cmd = [PYTHON, str(script_path)] + list(sel["args"]) + ["--tema", tema_novo]
            subprocess.Popen(cmd, cwd=str(Path(__file__).parent),
                             creationflags=subprocess.CREATE_NO_WINDOW)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a automação.\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = SeletorAutomacoes(root)
    root.mainloop()
