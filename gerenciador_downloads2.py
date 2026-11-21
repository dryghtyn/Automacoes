import os
import sys
import json
import time
import threading
from pathlib import Path
import urllib.parse
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext

DEPENDENCIAS = [("requests", "requests")]
for nome_import, nome_pip in DEPENDENCIAS:
    try:
        __import__(nome_import)
    except ImportError:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", nome_pip, "--quiet"])

import requests

BASE = Path(__file__).parent.resolve()

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
    padrao = {"download_dir": "", "auto_resume": True, "tasks": []}
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
        d.setdefault("download_dir", "")
        d.setdefault("auto_resume", True)
        d.setdefault("tasks", [])
        for t in d["tasks"]:
            t.setdefault("status", "paused")
            t.setdefault("downloaded", 0)
            t.setdefault("total", 0)
            t.setdefault("error", "")
        return d
    except Exception:
        return {"download_dir": "", "auto_resume": True, "tasks": []}


def _salvar_config(caminho, cfg):
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def _nome_arquivo(url):
    parsed = urllib.parse.urlparse(url)
    name = os.path.basename(parsed.path)
    if not name:
        name = "download.bin"
    # remove query
    name = name.split("?")[0]
    return name


class DownloadManagerApp:
    def __init__(self, root, config_path):
        self.root = root
        self.root.title("Gerenciador de Downloads")
        self.root.geometry("760x560")
        self.root.minsize(600, 400)
        self.config_path = Path(config_path)
        self.cfg = _carregar_config(self.config_path)
        self.lock = threading.Lock()
        self.rows = {}  # task id -> widgets
        self.threads = {}
        self.cancel_flags = {}

        frame = ttk.Frame(root, padding=(10, 8))
        frame.pack(fill=tk.BOTH, expand=True)

        # Pasta de download
        pasta_frame = ttk.Frame(frame)
        pasta_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(pasta_frame, text="Pasta de download:").pack(side=tk.LEFT)
        self.lbl_pasta = ttk.Label(pasta_frame, text=self.cfg["download_dir"] or "(não definida)", foreground="blue")
        self.lbl_pasta.pack(side=tk.LEFT, padx=6)
        ttk.Button(pasta_frame, text="Escolher pasta", command=self._escolher_pasta).pack(side=tk.RIGHT)

        # Lista de downloads
        lista_frame = ttk.LabelFrame(frame, text="Downloads", padding=(6, 4))
        lista_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        self.canvas = tk.Canvas(lista_frame, borderwidth=0)
        self.scroll = ttk.Scrollbar(lista_frame, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scroll.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Adicionar URL
        add_frame = ttk.Frame(frame)
        add_frame.pack(fill=tk.X, pady=(0, 6))
        self.entry_url = ttk.Entry(add_frame)
        self.entry_url.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(add_frame, text="Adicionar download", command=self._adicionar).pack(side=tk.RIGHT)
        # linha extra: navegação/extrair links
        nav_frame = ttk.Frame(frame)
        nav_frame.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(nav_frame, text="Abrir no navegador", command=self._abrir_navegador).pack(side=tk.LEFT)
        ttk.Button(nav_frame, text="Extrair links da página", command=self._extrair_links).pack(side=tk.LEFT, padx=6)
        ttk.Label(nav_frame, text="Cole o link copiado do navegador no campo acima.", font=("Segoe UI", 8), foreground="gray").pack(side=tk.RIGHT)

        # Controles globais
        ctrl_frame = ttk.Frame(frame)
        ctrl_frame.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(ctrl_frame, text="Continuar todos", command=self._continuar_todos).pack(side=tk.LEFT)
        ttk.Button(ctrl_frame, text="Pausar todos", command=self._pausar_todos).pack(side=tk.LEFT, padx=6)
        self.boost_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(ctrl_frame, text="Boost velocidade", variable=self.boost_var).pack(side=tk.LEFT, padx=6)
        self.tema_var = tk.StringVar(value="Claro")
        ttk.Combobox(ctrl_frame, textvariable=self.tema_var, values=["Claro", "Escuro"], state="readonly", width=9).pack(side=tk.LEFT, padx=6)
        self.tema_var.trace_add("write", lambda *a: self._aplicar_tema())
        ttk.Button(ctrl_frame, text="Atualizar do GitHub", command=self._atualizar_github).pack(side=tk.LEFT, padx=6)
        ttk.Button(ctrl_frame, text="Limpar concluídos", command=self._limpar_concluidos).pack(side=tk.RIGHT)

        self.log = scrolledtext.ScrolledText(frame, height=6, state="disabled", font=("Consolas", 9))
        self.log.pack(fill=tk.BOTH, expand=False)

        # Renderiza tasks existentes
        for t in self.cfg["tasks"]:
            self._render_task(t)
        self._aplicar_tema()
        # Auto-resume
        if self.cfg.get("auto_resume"):
            self.root.after(500, self._auto_resume)

    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.log.configure(state="disabled")

    def _escolher_pasta(self):
        d = filedialog.askdirectory(title="Escolha a pasta de download")
        if d:
            self.cfg["download_dir"] = d
            self.lbl_pasta.config(text=d)
            _salvar_config(self.config_path, self.cfg)
            self._log(f"Pasta definida: {d}")

    def _abrir_navegador(self):
        url = self.entry_url.get().strip()
        if not url:
            url = "https://www.google.com"
        import webbrowser
        try:
            webbrowser.open(url)
            self._log(f"Abrindo no navegador: {url}")
        except Exception as e:
            self._log(f"Erro ao abrir navegador: {e}")

    def _extrair_links(self):
        url = self.entry_url.get().strip()
        if not url:
            messagebox.showwarning("Aviso", "Informe a URL da página.")
            return
        try:
            r = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
            r.raise_for_status()
            html = r.text
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao buscar página: {e}")
            return
        import re
        # encontra hrefs
        hrefs = re.findall(r'href=["\']([^"\']+)["\']', html, re.IGNORECASE)
        # filtra links com extensões de arquivo ou contendo 'download'
        candidatos = []
        seen = set()
        for h in hrefs:
            full = urllib.parse.urljoin(url, h)
            if full in seen:
                continue
            seen.add(full)
            low = full.lower()
            if ("download" in low or any(low.endswith(ext) for ext in (".zip",".rar",".7z",".exe",".msi",".mp4",".mp3",".pdf",".iso",".bin",".tar",".gz",".jpg",".png"))):
                candidatos.append(full)
        if not candidatos:
            candidatos = list(seen)[:50]
        # dialog de seleção
        dlg = tk.Toplevel(self.root)
        dlg.title("Links encontrados")
        dlg.geometry("600x400")
        lb = tk.Listbox(dlg, width=90, height=20)
        lb.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        for c in candidatos:
            lb.insert(tk.END, c)
        def escolher():
            sel = lb.curselection()
            if not sel:
                return
            self.entry_url.delete(0, tk.END)
            self.entry_url.insert(0, candidatos[sel[0]])
            dlg.destroy()
        ttk.Button(dlg, text="Usar link selecionado", command=escolher).pack(pady=4)
        self._log(f"Extraídos {len(candidatos)} links de {url}")

    def _adicionar(self):
        url = self.entry_url.get().strip()
        if not url:
            messagebox.showwarning("Aviso", "Informe a URL.")
            return
        if not self.cfg["download_dir"]:
            self._escolher_pasta()
            if not self.cfg["download_dir"]:
                messagebox.showwarning("Aviso", "Escolha uma pasta de download.")
                return
        nome = _nome_arquivo(url)
        # pergunta nome
        from tkinter import simpledialog
        nome = simpledialog.askstring("Nome do arquivo", "Nome do arquivo:", initialvalue=nome, parent=self.root)
        if not nome:
            return
        dest = os.path.join(self.cfg["download_dir"], nome)
        # id
        ids = [t["id"] for t in self.cfg["tasks"]]
        nid = (max(ids) if ids else 0) + 1
        task = {"id": nid, "url": url, "filename": nome, "dest": dest, "downloaded": 0, "total": 0, "status": "paused", "error": ""}
        # se arquivo já existe parcialmente, retoma tamanho
        if os.path.exists(dest):
            task["downloaded"] = os.path.getsize(dest)
        self.cfg["tasks"].append(task)
        _salvar_config(self.config_path, self.cfg)
        self._render_task(task)
        self.entry_url.delete(0, tk.END)
        self._log(f"Adicionado: {nome}")

    def _render_task(self, task):
        tid = task["id"]
        if tid in self.rows:
            return
        row = ttk.Frame(self.inner, padding=(4, 3))
        row.pack(fill=tk.X, pady=2)
        btn_frame = ttk.Frame(row)
        btn_frame.pack(side=tk.RIGHT)
        status = ttk.Label(row, text=task["status"], width=34, anchor="w")
        status.pack(side=tk.RIGHT, padx=4)
        pb = ttk.Progressbar(row, length=180, mode="determinate")
        pb.pack(side=tk.LEFT, padx=6)
        lbl = ttk.Label(row, text=task["filename"], width=22)
        lbl.pack(side=tk.LEFT)
        btn_cont = ttk.Button(btn_frame, text="▶", width=3, command=lambda t=task: self._continuar(t))
        btn_cont.pack(side=tk.LEFT, padx=1)
        btn_pause = ttk.Button(btn_frame, text="⏸", width=3, command=lambda t=task: self._pausar(t))
        btn_pause.pack(side=tk.LEFT, padx=1)
        btn_del = ttk.Button(btn_frame, text="✕", width=3, command=lambda t=task: self._remover(t))
        btn_del.pack(side=tk.LEFT, padx=1)
        self.rows[tid] = {"frame": row, "pb": pb, "status": status, "task": task}
        self._atualizar_row(tid)

    def _atualizar_row(self, tid):
        info = self.rows.get(tid)
        if not info:
            return
        task = info["task"]
        # sync task ref from cfg
        for t in self.cfg["tasks"]:
            if t["id"] == tid:
                task = t
                info["task"] = t
                break
        pb = info["pb"]
        status = info["status"]
        total = task.get("total", 0)
        downloaded = task.get("downloaded", 0)
        if total and total > 0:
            pb["maximum"] = total
            pb["value"] = downloaded
        else:
            pb["maximum"] = 100
            pb["value"] = 0
        def _fmt(n):
            n = n or 0
            if n >= 1024*1024*1024:
                return f"{n/1024/1024/1024:.2f} GB"
            if n >= 1024*1024:
                return f"{n/1024/1024:.2f} MB"
            if n >= 1024:
                return f"{n/1024:.1f} KB"
            return f"{n} B"
        st = task.get("status", "paused")
        if task.get("error"):
            st = "erro"
        speed = task.get("speed", 0)
        speed_str = ""
        if speed and speed > 0:
            speed_str = f"  ({_fmt(speed)}/s)"
        if total and total > 0:
            status.config(text=f"{st}: {_fmt(downloaded)} / {_fmt(total)}{speed_str}")
        else:
            status.config(text=f"{st}: {_fmt(downloaded)}{speed_str}")

    def _continuar(self, task):
        tid = task["id"]
        if tid in self.threads and self.threads[tid].is_alive():
            return
        self.cancel_flags[tid] = False
        # atualiza task ref
        for t in self.cfg["tasks"]:
            if t["id"] == tid:
                t["status"] = "downloading"
                t["error"] = ""
                break
        _salvar_config(self.config_path, self.cfg)
        self._atualizar_row(tid)
        th = threading.Thread(target=self._download_worker, args=(tid,), daemon=True)
        self.threads[tid] = th
        th.start()

    def _pausar(self, task):
        tid = task["id"]
        self.cancel_flags[tid] = True
        for t in self.cfg["tasks"]:
            if t["id"] == tid:
                t["status"] = "paused"
                break
        _salvar_config(self.config_path, self.cfg)
        self._atualizar_row(tid)
        self._log(f"Pausado: {task['filename']}")

    def _remover(self, task):
        tid = task["id"]
        self.cancel_flags[tid] = True
        self.cfg["tasks"] = [t for t in self.cfg["tasks"] if t["id"] != tid]
        _salvar_config(self.config_path, self.cfg)
        info = self.rows.pop(tid, None)
        if info:
            info["frame"].destroy()
        self._log(f"Removido: {task['filename']}")

    def _download_worker(self, tid):
        # encontra task
        task = None
        for t in self.cfg["tasks"]:
            if t["id"] == tid:
                task = t
                break
        if not task:
            return
        url = task["url"]
        dest = task["dest"]
        downloaded = task.get("downloaded", 0)
        # verifica arquivo existente
        if os.path.exists(dest):
            fs = os.path.getsize(dest)
            if fs != downloaded:
                downloaded = fs
                task["downloaded"] = downloaded
        headers = {}
        if downloaded > 0:
            headers["Range"] = f"bytes={downloaded}-"
        mode = "ab" if downloaded > 0 else "wb"
        if downloaded == 0:
            # truncate
            try:
                if os.path.exists(dest):
                    os.remove(dest)
            except Exception:
                pass
        boost = getattr(self, "boost_var", None)
        boost_on = boost.get() if boost else False
        chunk_size = 1024*1024 if boost_on else 65536
        try:
            r = requests.get(url, headers=headers, stream=True, timeout=30)
            total = task.get("total", 0)
            if "Content-Range" in r.headers:
                try:
                    total = int(r.headers["Content-Range"].split("/")[-1])
                except Exception:
                    pass
            elif "Content-Length" in r.headers:
                cl = int(r.headers["Content-Length"])
                total = downloaded + cl if downloaded > 0 else cl
            task["total"] = total
            if r.status_code == 200 and downloaded > 0:
                downloaded = 0
                task["downloaded"] = 0
                mode = "wb"
                if os.path.exists(dest):
                    os.remove(dest)
                total = int(r.headers.get("Content-Length", 0))
                task["total"] = total
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            last_time = time.time()
            last_bytes = downloaded
            with open(dest, mode) as f:
                for chunk in r.iter_content(chunk_size=chunk_size):
                    if self.cancel_flags.get(tid):
                        break
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        task["downloaded"] = downloaded
                        now = time.time()
                        dt = now - last_time
                        if dt >= 0.5:
                            speed = (downloaded - last_bytes) / dt
                            task["speed"] = speed
                            last_time = now
                            last_bytes = downloaded
                        self.root.after(0, self._atualizar_row, tid)
            if self.cancel_flags.get(tid):
                task["status"] = "paused"
                self._log(f"Pausado: {task['filename']}")
            else:
                task["status"] = "done" if (total and downloaded >= total) else "done"
                self._log(f"Concluído: {task['filename']}")
        except Exception as e:
            task["status"] = "error"
            task["error"] = str(e)
            self._log(f"Erro {task['filename']}: {e}")
            # auto-retry em 5 segundos
            try:
                self.root.after(5000, lambda tid=tid: self._retry_task(tid))
            except Exception:
                pass
        finally:
            task["status"] = task.get("status", "paused")
            _salvar_config(self.config_path, self.cfg)
            self.root.after(0, self._atualizar_row, tid)

    def _retry_task(self, tid):
        for t in self.cfg["tasks"]:
            if t["id"] == tid:
                if t.get("status") != "error":
                    return
                t["status"] = "paused"
                t["error"] = ""
                _salvar_config(self.config_path, self.cfg)
                self._log(f"Retentando em 5s: {t['filename']}")
                self._continuar(t)
                return

    def _continuar_todos(self):
        for t in list(self.cfg["tasks"]):
            if t.get("status") not in ("done",):
                self._continuar(t)

    def _pausar_todos(self):
        for t in list(self.cfg["tasks"]):
            self._pausar(t)

    def _limpar_concluidos(self):
        concluidos = [t["id"] for t in self.cfg["tasks"] if t.get("status") == "done"]
        self.cfg["tasks"] = [t for t in self.cfg["tasks"] if t["id"] not in concluidos]
        for tid in concluidos:
            info = self.rows.pop(tid, None)
            if info:
                info["frame"].destroy()
        _salvar_config(self.config_path, self.cfg)
        self._log("Concluídos removidos.")

    def _aplicar_tema(self):
        estilo = ttk.Style(self.root)
        try:
            estilo.theme_use("clam")
        except Exception:
            pass
        if self.tema_var.get() == "Escuro":
            bg = "#1e1e1e"; fg = "#e6e6e6"; panel = "#2d2d30"; borda = "#3e3e42"
            campo = "#2d2d30"; campo_fg = "#e6e6e6"; sel = "#094771"
        else:
            bg = "#f0f0f0"; fg = "#1a1a1a"; panel = "#fafafa"; borda = "#c9c9c9"
            campo = "#ffffff"; campo_fg = "#1a1a1a"; sel = "#0078d4"
        estilo.configure(".", background=bg, foreground=fg)
        estilo.configure("TFrame", background=bg)
        estilo.configure("TLabel", background=bg, foreground=fg)
        estilo.configure("TLabelframe", background=bg, foreground=fg, bordercolor=borda)
        estilo.configure("TLabelframe.Label", background=bg, foreground=fg)
        estilo.configure("TButton", background=panel, foreground=fg, bordercolor=borda)
        estilo.map("TButton", background=[("active", sel)], foreground=[("active", "#ffffff")])
        estilo.configure("TEntry", fieldbackground=campo, foreground=campo_fg, bordercolor=borda, insertcolor=campo_fg)
        estilo.configure("TCheckbutton", background=bg, foreground=fg)
        estilo.configure("TCombobox", fieldbackground=campo, foreground=campo_fg, background=panel, bordercolor=borda)
        estilo.configure("TProgressbar", background=sel, troughcolor=panel)
        self.root.configure(bg=bg)
        try:
            self.canvas.configure(bg=bg)
            self.inner.configure(style="TFrame")
        except Exception:
            pass
        try:
            self.log.configure(bg=campo, fg=campo_fg, insertbackground=campo_fg)
        except Exception:
            pass
        # aplica bg a todos os frames/labels filhos
        def _rec(w):
            try:
                if isinstance(w, (tk.Frame, tk.LabelFrame, tk.Canvas)):
                    w.configure(bg=bg)
                elif isinstance(w, tk.Label):
                    w.configure(bg=bg, fg=fg)
            except Exception:
                pass
            for c in w.winfo_children():
                _rec(c)
        try:
            _rec(self.root)
        except Exception:
            pass

    def _atualizar_github(self):
        # lê github_config.json do painel (se existir) para obter owner/repo/branch
        cfg_gh = None
        try:
            gh_path = BASE / "github_config.json"
            if gh_path.exists():
                with open(gh_path, "r", encoding="utf-8") as f:
                    cfg_gh = json.load(f)
        except Exception:
            cfg_gh = None
        if not cfg_gh or not cfg_gh.get("owner") or not cfg_gh.get("repo"):
            # pede ao usuário
            from tkinter import simpledialog
            owner = simpledialog.askstring("GitHub", "Owner (usuário):", parent=self.root)
            repo = simpledialog.askstring("GitHub", "Repositório:", parent=self.root)
            branch = simpledialog.askstring("GitHub", "Branch:", initialvalue="main", parent=self.root) or "main"
            if not owner or not repo:
                self._log("Atualização cancelada: owner/repo não informados.")
                return
            cfg_gh = {"owner": owner, "repo": repo, "branch": branch, "token": ""}
        owner = cfg_gh["owner"].strip()
        repo = cfg_gh["repo"].strip()
        branch = (cfg_gh.get("branch") or "main").strip() or "main"
        token = (cfg_gh.get("token") or "").strip()
        arquivo = "gerenciador_downloads.py"
        import urllib.request
        url = f"https://raw.githubusercontent.com/{owner}/{repo}/{branch}/{arquivo}"
        try:
            req = urllib.request.Request(url)
            if token:
                req.add_header("Authorization", f"Bearer {token}")
            req.add_header("User-Agent", "gerenciador-downloads")
            with urllib.request.urlopen(req, timeout=30) as resp:
                novo = resp.read()
            # backup
            atual = BASE / arquivo
            backup = BASE / (arquivo + ".bak")
            if atual.exists():
                try:
                    backup.write_bytes(atual.read_bytes())
                except Exception:
                    pass
            atual.write_bytes(novo)
            self._log(f"Atualizado de {owner}/{repo}. Reinicie a automação para aplicar.")
            from tkinter import messagebox
            messagebox.showinfo("Atualizado", "Download da atualização concluído. Reinicie a automação para aplicar.")
        except Exception as e:
            self._log(f"Falha ao atualizar: {e}")
            from tkinter import messagebox
            messagebox.showerror("Erro", f"Falha ao baixar atualização:\n{e}")

    def _auto_resume(self):
        for t in list(self.cfg["tasks"]):
            if t.get("status") in ("downloading", "paused", "error") and t.get("downloaded", 0) < (t.get("total") or 1):
                self._continuar(t)


def main():
    config_path = BASE / "automacoes" / "Gerenciador de Downloads" / "download_config.json"
    if "--config" in sys.argv:
        idx = sys.argv.index("--config")
        if idx + 1 < len(sys.argv):
            config_path = Path(sys.argv[idx + 1])
    if not config_path.is_absolute():
        config_path = BASE / config_path

    root = tk.Tk()
    app = DownloadManagerApp(root, config_path)
    root.mainloop()


if __name__ == "__main__":
    main()
