import os
import sys
import json
import base64
import threading
from pathlib import Path
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog

BASE = Path(__file__).parent.resolve()

DEPENDENCIAS = []
# (sem dependǦncias externas; usa stdlib)

# DPI awareness
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def _carregar_config(caminho):
    p = Path(caminho)
    padrao = {"owner": "", "repo": "", "branch": "main", "token": "", "arquivos": []}
    if not p.exists():
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            with open(p, "w", encoding="utf-8") as f:
                json.dump(padrao, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        return padrao
    try:
        with open(p, "r", encoding="utf-8") as f:
            d = json.load(f)
        d.setdefault("owner", "")
        d.setdefault("repo", "")
        d.setdefault("branch", "main")
        d.setdefault("token", "")
        d.setdefault("arquivos", [])
        return d
    except Exception:
        return {"owner": "", "repo": "", "branch": "main", "token": "", "arquivos": []}


def _descobrir_py(base):
    arquivos = []
    ignorar = {"__pycache__", ".git", "node_modules", "venv", ".venv"}
    for root, dirs, files in os.walk(base):
        dirs[:] = [d for d in dirs if d not in ignorar]
        for fn in files:
            if not fn.endswith(".py"):
                continue
            full = Path(root) / fn
            try:
                rel = full.relative_to(base).as_posix()
            except ValueError:
                continue
            arquivos.append(rel)
    arquivos.sort()
    return arquivos


class EnviarGithubApp:
    def __init__(self, root, config_path):
        self.root = root
        self.root.title("Enviar arquivos para o GitHub")
        self.root.geometry("720x560")
        self.root.minsize(560, 440)
        self.config_path = Path(config_path)
        self.config = _carregar_config(self.config_path)
        self.arquivos_disponiveis = _descobrir_py(BASE)
        self.checks = {}
        self.selected_files = []

        frame = ttk.Frame(root, padding=(12, 10))
        frame.pack(fill=tk.BOTH, expand=True)

        info = ttk.LabelFrame(frame, text="Configura��ǜo (github_upload_config.json)", padding=(8, 6))
        info.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(info, text=f"Owner: {self.config.get('owner') or '(vazio)'}").pack(anchor="w")
        ttk.Label(info, text=f"Repo: {self.config.get('repo') or '(vazio)'}").pack(anchor="w")
        ttk.Label(info, text=f"Branch: {self.config.get('branch') or 'main'}").pack(anchor="w")
        ttk.Label(info, text=f"Token: {'***' if self.config.get('token') else '(vazio)'}").pack(anchor="w")
        ttk.Label(info, text=f"Arquivo de config: {self.config_path}", font=("Segoe UI", 8)).pack(anchor="w")

        if not self.config.get("owner") or not self.config.get("repo") or not self.config.get("token"):
            ttk.Label(info, text="Preencha owner/repo/token no arquivo de config antes de enviar.",
                      foreground="red").pack(anchor="w", pady=(4, 0))

        self.notebook = ttk.Notebook(frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        tab_py = ttk.Frame(self.notebook)
        tab_arquivos = ttk.Frame(self.notebook)
        self.notebook.add(tab_py, text="Arquivos .py")
        self.notebook.add(tab_arquivos, text="Enviar arquivos")

        self._construir_tab_py(tab_py)
        self._construir_tab_arquivos(tab_arquivos)

        self.log = scrolledtext.ScrolledText(frame, height=8, state="disabled", font=("Consolas", 9))
        self.log.pack(fill=tk.BOTH, expand=False)

    def _construir_tab_py(self, parent):
        lista_frame = ttk.LabelFrame(parent, text="Arquivos .py dispon��veis", padding=(8, 6))
        lista_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 8))

        canvas = tk.Canvas(lista_frame, borderwidth=0)
        scroll = ttk.Scrollbar(lista_frame, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)
        self.inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        selecionados = set(self.config.get("arquivos") or [])
        todos = not selecionados
        for rel in self.arquivos_disponiveis:
            var = tk.BooleanVar(value=(rel in selecionados) if selecionados else True)
            self.checks[rel] = var
            ttk.Checkbutton(self.inner, text=rel, variable=var).pack(anchor="w")

        botoes = ttk.Frame(parent)
        botoes.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(botoes, text="Selecionar todos", command=self._sel_todos).pack(side=tk.LEFT)
        ttk.Button(botoes, text="Limpar", command=self._limpar).pack(side=tk.LEFT, padx=6)
        ttk.Button(botoes, text="Enviar selecionados", command=self._enviar_py).pack(side=tk.RIGHT)

    def _construir_tab_arquivos(self, parent):
        frame_lista = ttk.LabelFrame(parent, text="Arquivos selecionados", padding=(8, 6))
        frame_lista.pack(fill=tk.BOTH, expand=True, pady=(8, 8))

        self.lista_arquivos = tk.Listbox(frame_lista)
        self.lista_arquivos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll = ttk.Scrollbar(frame_lista, orient="vertical", command=self.lista_arquivos.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.lista_arquivos.configure(yscrollcommand=scroll.set)

        botoes = ttk.Frame(parent)
        botoes.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(botoes, text="Selecionar arquivos", command=self._selecionar_arquivos).pack(side=tk.LEFT)
        ttk.Button(botoes, text="Limpar lista", command=self._limpar_arquivos).pack(side=tk.LEFT, padx=6)
        ttk.Button(botoes, text="Enviar arquivos", command=self._enviar_arquivos).pack(side=tk.RIGHT)

    def _sel_todos(self):
        for v in self.checks.values():
            v.set(True)

    def _limpar(self):
        for v in self.checks.values():
            v.set(False)

    def _selecionar_arquivos(self):
        caminhos = filedialog.askopenfilenames(title="Selecionar arquivos para enviar")
        if not caminhos:
            return
        for c in caminhos:
            p = Path(c)
            if p not in self.selected_files:
                self.selected_files.append(p)
                self.lista_arquivos.insert(tk.END, str(p))

    def _limpar_arquivos(self):
        self.selected_files = []
        self.lista_arquivos.delete(0, tk.END)

    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.log.configure(state="disabled")

    def _enviar_py(self):
        sel = [rel for rel, v in self.checks.items() if v.get()]
        if not sel:
            messagebox.showwarning("Aviso", "Selecione ao menos um arquivo.")
            return
        if not self.config.get("owner") or not self.config.get("repo") or not self.config.get("token"):
            messagebox.showerror("Erro", "Configure owner/repo/token no arquivo de config.")
            return
        paths = [BASE / rel for rel in sel]
        threading.Thread(target=self._upload_paths, args=(paths,), daemon=True).start()

    def _enviar_arquivos(self):
        if not self.selected_files:
            messagebox.showwarning("Aviso", "Selecione ao menos um arquivo.")
            return
        if not self.config.get("owner") or not self.config.get("repo") or not self.config.get("token"):
            messagebox.showerror("Erro", "Configure owner/repo/token no arquivo de config.")
            return
        threading.Thread(target=self._upload_paths, args=(list(self.selected_files),), daemon=True).start()

    def _upload_paths(self, selecionados):
        owner = self.config["owner"].strip()
        repo = self.config["repo"].strip()
        branch = (self.config.get("branch") or "main").strip() or "main"
        token = self.config["token"].strip()
        for item in selecionados:
            alvo = Path(item)
            self.root.after(0, self._log, f"Enviando {alvo.name}...")
            try:
                if not alvo.is_file():
                    self.root.after(0, self._log, f"  ERRO: arquivo nǜo encontrado")
                    continue
                try:
                    rel = alvo.relative_to(BASE.resolve()).as_posix()
                except ValueError:
                    rel = alvo.name
                conteudo = alvo.read_bytes()
                b64 = base64.b64encode(conteudo).decode("ascii")
                sha = None
                try:
                    gurl = f"https://api.github.com/repos/{owner}/{repo}/contents/{rel}?ref={branch}"
                    greq = urllib.request.Request(gurl)
                    greq.add_header("Authorization", f"Bearer {token}")
                    greq.add_header("Accept", "application/vnd.github+json")
                    greq.add_header("User-Agent", "painel-automacoes")
                    with urllib.request.urlopen(greq, timeout=30) as gresp:
                        gdata = json.loads(gresp.read().decode("utf-8"))
                        sha = gdata.get("sha")
                except urllib.error.HTTPError as he:
                    if he.code != 404:
                        raise
                payload = {"message": f"update {rel}", "content": b64, "branch": branch}
                if sha:
                    payload["sha"] = sha
                    acao = "atualizado"
                else:
                    acao = "criado"
                put_url = f"https://api.github.com/repos/{owner}/{repo}/contents/{rel}"
                preq = urllib.request.Request(put_url, data=json.dumps(payload).encode("utf-8"), method="PUT")
                preq.add_header("Authorization", f"Bearer {token}")
                preq.add_header("Accept", "application/vnd.github+json")
                preq.add_header("Content-Type", "application/json")
                preq.add_header("User-Agent", "painel-automacoes")
                with urllib.request.urlopen(preq, timeout=60) as presp:
                    pass
                self.root.after(0, self._log, f"  OK: {acao}")
            except Exception as e:
                self.root.after(0, self._log, f"  ERRO: {e}")
        self.root.after(0, self._log, "Conclu��do.")


def main():
    config_path = BASE / "automacoes" / "Enviar para GitHub" / "github_upload_config.json"
    # permite override via --config
    if "--config" in sys.argv:
        idx = sys.argv.index("--config")
        if idx + 1 < len(sys.argv):
            config_path = Path(sys.argv[idx + 1])
    if not config_path.is_absolute():
        config_path = BASE / config_path

    root = tk.Tk()
    app = EnviarGithubApp(root, config_path)
    root.mainloop()


if __name__ == "__main__":
    main()
