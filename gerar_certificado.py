import subprocess
import importlib.util
import sys
import ctypes
import threading
import os
import time
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

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

_TEMA_ARG = ""
if "--tema" in sys.argv:
    try:
        _TEMA_ARG = sys.argv[sys.argv.index("--tema") + 1].strip(" \"")
    except (IndexError, ValueError):
        _TEMA_ARG = ""

sys.path.insert(0, str(Path(__file__).parent))
import tema

TEMA = _TEMA_ARG if _TEMA_ARG else tema.carregar_tema()

DIR_RAIZ = Path(__file__).parent
OPENSSL_BIN = DIR_RAIZ / "OpenSSL" / "bin"
OPENSSL_EXE = OPENSSL_BIN / "openssl.exe"
if not OPENSSL_EXE.exists():
    OPENSSL_EXE = Path("openssl")

AUTOMACAO_DIR = DIR_RAIZ / "automacoes" / "Gerar Certificado"


def _cor_mutado():
    return "#9a9a9a" if TEMA != "dark" else "#8d8d8d"


def _cor_destaque():
    return "#005a9e" if TEMA != "dark" else "#3794ff"


class GerarCertificado:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerar Certificado (PFX -> CRT/PEM/KEY)")
        self.root.geometry("760x560")
        self.root.minsize(640, 460)

        self.pfx_path = ""
        self.senha = ""
        self.senha_valida = False
        self.executando = False
        self.pasta_saida = AUTOMACAO_DIR if AUTOMACAO_DIR.exists() else DIR_RAIZ

        self._construir_ui()
        tema.aplicar_tema(self.root, TEMA)

    def _construir_ui(self):
        frame_pfx = ttk.LabelFrame(self.root, text="Certificado (.pfx)", padding=8)
        frame_pfx.pack(fill=tk.X, padx=8, pady=(8, 4))

        linha = ttk.Frame(frame_pfx)
        linha.pack(fill=tk.X)
        ttk.Button(linha, text="Selecionar .pfx", command=self._selecionar_pfx).pack(side=tk.LEFT)
        self.lbl_pfx = ttk.Label(linha, text="Nenhum arquivo selecionado", foreground=_cor_mutado())
        self.lbl_pfx.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        frame_senha = ttk.LabelFrame(self.root, text="Senha do certificado", padding=8)
        frame_senha.pack(fill=tk.X, padx=8, pady=2)

        linha2 = ttk.Frame(frame_senha)
        linha2.pack(fill=tk.X)
        ttk.Label(linha2, text="Senha:").pack(side=tk.LEFT)
        self.ent_senha = ttk.Entry(linha2, show="*", width=40)
        self.ent_senha.pack(side=tk.LEFT, padx=(4, 8), fill=tk.X, expand=True)
        self.btn_validar = ttk.Button(linha2, text="Validar Senha", command=self._validar_senha)
        self.btn_validar.pack(side=tk.LEFT)
        self.lbl_validacao = ttk.Label(frame_senha, text="", foreground=_cor_mutado())
        self.lbl_validacao.pack(anchor="w", pady=(4, 0))

        frame_saida = ttk.LabelFrame(self.root, text="Pasta de saida", padding=8)
        frame_saida.pack(fill=tk.X, padx=8, pady=2)

        linha3 = ttk.Frame(frame_saida)
        linha3.pack(fill=tk.X)
        self.ent_saida = ttk.Entry(linha3, width=60)
        self.ent_saida.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.ent_saida.insert(0, str(self.pasta_saida))
        ttk.Button(linha3, text="...", command=self._escolher_saida, width=4).pack(side=tk.LEFT, padx=(4, 0))

        frame_acoes = ttk.Frame(self.root)
        frame_acoes.pack(fill=tk.X, padx=8, pady=6)
        self.btn_gerar = ttk.Button(frame_acoes, text="Gerar Arquivos (3 comandos)", command=self._gerar)
        self.btn_gerar.pack(side=tk.LEFT)
        self.btn_gerar.state(["disabled"])
        self.btn_abrir = ttk.Button(frame_acoes, text="Abrir pasta de saida", command=self._abrir_pasta)
        self.btn_abrir.pack(side=tk.LEFT, padx=8)

        frame_status = ttk.Frame(self.root)
        frame_status.pack(fill=tk.X, padx=8, pady=(0, 2))
        self.lbl_indicador = ttk.Label(frame_status, text="?", foreground="#4caf50", font=("Segoe UI", 11))
        self.lbl_indicador.pack(side=tk.LEFT, padx=(0, 4))
        self.lbl_status = ttk.Label(frame_status, text="Aguardando selecao do certificado", foreground=_cor_mutado(), font=("Segoe UI", 9, "bold"))
        self.lbl_status.pack(side=tk.LEFT)
        self.progress = ttk.Progressbar(frame_status, mode="indeterminate", length=120)
        self.progress.pack(side=tk.LEFT, padx=12)
        self.progress.pack_forget()

        frame_log = ttk.LabelFrame(self.root, text="Log", padding=4)
        frame_log.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        self.log = scrolledtext.ScrolledText(frame_log, height=8, state=tk.DISABLED, font=("Consolas", 9))
        self.log.pack(fill=tk.BOTH, expand=True)

    def _log(self, msg):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)

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

    def _selecionar_pfx(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar certificado digital (.pfx)",
            filetypes=[("PKCS#12", "*.pfx *.p12"), ("Todos", "*.*")])
        if not caminho:
            return
        self.pfx_path = caminho
        self.senha_valida = False
        self.btn_gerar.state(["disabled"])
        self.lbl_pfx.configure(text=os.path.basename(caminho))
        self.lbl_validacao.configure(text="Senha nao validada.", foreground=_cor_mutado())
        self._log(f"Certificado selecionado: {caminho}")

    def _escolher_saida(self):
        pasta = filedialog.askdirectory(title="Selecionar pasta de saida", initialdir=self.ent_saida.get())
        if pasta:
            self.ent_saida.delete(0, tk.END)
            self.ent_saida.insert(0, pasta)

    def _validar_senha(self):
        if not self.pfx_path:
            messagebox.showwarning("Aviso", "Selecione um arquivo .pfx primeiro.")
            return
        senha = self.ent_senha.get()
        if not senha:
            messagebox.showwarning("Aviso", "Digite a senha do certificado.")
            return
        self._set_status("Validando senha...", "rodando")
        self.btn_validar.state(["disabled"])
        threading.Thread(target=self._validar_thread, args=(senha,), daemon=True).start()

    def _validar_thread(self, senha):
        try:
            proc = subprocess.run(
                [str(OPENSSL_EXE), "pkcs12", "-in", self.pfx_path,
                 "-passin", f"pass:{senha}", "-noout"],
                cwd=str(OPENSSL_BIN) if OPENSSL_BIN.exists() else None,
                capture_output=True, text=True, timeout=60)
            ok = (proc.returncode == 0)
            err = (proc.stderr or proc.stdout or "").strip()
        except Exception as e:
            ok = False
            err = str(e)
        self.root.after(0, lambda: self._validar_resultado(ok, err, senha))

    def _validar_resultado(self, ok, err, senha):
        self.btn_validar.state(["!disabled"])
        if ok:
            self.senha = senha
            self.senha_valida = True
            self.btn_gerar.state(["!disabled"])
            self.lbl_validacao.configure(text="Senha valida.", foreground="#4caf50")
            self._set_status("Senha validada com sucesso.", "ok")
            self._log("Senha validada com sucesso.")
        else:
            self.senha_valida = False
            self.btn_gerar.state(["disabled"])
            self.lbl_validacao.configure(text="Senha incorreta.", foreground="#e53935")
            self._set_status("Senha incorreta.", "erro")
            self._log("Senha incorreta.")

    def _gerar(self):
        if not self.pfx_path or not self.senha_valida:
            messagebox.showwarning("Aviso", "Selecione o .pfx e valide a senha primeiro.")
            return
        base_saida = Path(self.ent_saida.get().strip() or str(self.pasta_saida))
        nome = Path(self.pfx_path).stem
        pasta = base_saida / nome
        threading.Thread(target=self._gerar_thread, args=(pasta, nome), daemon=True).start()

    def _rodar_cmd(self, args, cwd):
        return subprocess.run(
            [str(OPENSSL_EXE)] + args,
            cwd=cwd, capture_output=True, text=True, timeout=120)

    def _gerar_thread(self, pasta, nome):
        self.root.after(0, lambda: self._set_status("Gerando arquivos...", "rodando"))
        self.root.after(0, lambda: self.btn_gerar.state(["disabled"]))
        cwd = str(OPENSSL_BIN) if OPENSSL_BIN.exists() else None
        try:
            pasta.mkdir(parents=True, exist_ok=True)
            self.root.after(0, lambda: self._log(f"Pasta de saida: {pasta}"))

            self.root.after(0, lambda: self._log("Executando comando 1 (extrair certificado)..."))
            crt = pasta / f"{nome}.crt"
            r1 = self._rodar_cmd([
                "pkcs12", "-in", self.pfx_path, "-clcerts", "-nokeys",
                "-out", str(crt), "-passin", f"pass:{self.senha}"], cwd)
            if r1.returncode != 0:
                raise RuntimeError(f"Comando 1 falhou:\n{r1.stderr or r1.stdout}")
            self.root.after(0, lambda: self._log(f"OK: {crt}"))

            self.root.after(0, lambda: self._log("Executando comando 2 (extrair chave privada)..."))
            pem = pasta / f"{nome}.pem"
            r2 = self._rodar_cmd([
                "pkcs12", "-in", self.pfx_path, "-nocerts",
                "-out", str(pem), "-passin", f"pass:{self.senha}",
                "-passout", f"pass:{self.senha}"], cwd)
            if r2.returncode != 0:
                raise RuntimeError(f"Comando 2 falhou:\n{r2.stderr or r2.stdout}")
            self.root.after(0, lambda: self._log(f"OK: {pem}"))

            self.root.after(0, lambda: self._log("Executando comando 3 (converter para PKCS8)..."))
            key = pasta / f"{nome}.key"
            r3 = self._rodar_cmd([
                "pkcs8", "-in", str(pem), "-topk8", "-nocrypt",
                "-out", str(key), "-passin", f"pass:{self.senha}"], cwd)
            if r3.returncode != 0:
                raise RuntimeError(f"Comando 3 falhou:\n{r3.stderr or r3.stdout}")
            self.root.after(0, lambda: self._log(f"OK: {key}"))

            self.root.after(0, lambda: self._set_status("Arquivos gerados com sucesso.", "ok"))
            self.root.after(0, lambda: self._log("Concluido. Todos os arquivos gerados."))
            self.root.after(0, lambda: messagebox.showinfo("Sucesso", f"Arquivos gerados em:\n{pasta}"))
        except Exception as e:
            self.root.after(0, lambda: self._set_status("Erro na geracao.", "erro"))
            self.root.after(0, lambda: self._log(f"ERRO: {e}"))
            self.root.after(0, lambda: messagebox.showerror("Erro", str(e)))
        finally:
            self.root.after(0, lambda: self.btn_gerar.state(["!disabled"]))

    def _abrir_pasta(self):
        pasta = self.ent_saida.get().strip()
        nome = Path(self.pfx_path).stem if self.pfx_path else ""
        alvo = Path(pasta) / nome if nome else Path(pasta)
        if not alvo.exists():
            alvo = Path(pasta)
        try:
            os.startfile(str(alvo))
        except Exception as e:
            messagebox.showerror("Erro", f"Nao foi possivel abrir a pasta.\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = GerarCertificado(root)
    root.mainloop()
PYEOF