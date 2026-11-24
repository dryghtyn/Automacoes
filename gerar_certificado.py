import subprocess
import importlib.util
import sys
import ctypes
import threading
import os
import re
import shutil
import time
import urllib.error
import urllib.request
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
OPENSSL_CNF = OPENSSL_BIN.parent / "share" / "openssl.cnf"
if not OPENSSL_CNF.exists():
    OPENSSL_CNF = DIR_RAIZ / "OpenSSL" / "share" / "openssl.cnf"

AUTOMACAO_DIR = DIR_RAIZ / "automacoes" / "Gerar Certificado"


def _cor_mutado():
    return "#9a9a9a" if TEMA != "dark" else "#8d8d8d"


def _cor_destaque():
    return "#005a9e" if TEMA != "dark" else "#3794ff"


def _env_openssl():
    env = os.environ.copy()
    if OPENSSL_CNF.exists():
        env["OPENSSL_CONF"] = str(OPENSSL_CNF)
    return env


def _git_bash_exe():
    candidatos = [
        Path("C:/Program Files/Git/bin/bash.exe"),
        Path("C:/Program Files (x86)/Git/bin/bash.exe"),
        Path("C:/Program Files/Git/usr/bin/bash.exe"),
        Path("C:/Program Files/Git/git-bash.exe").parent / "bin" / "bash.exe",
    ]
    try:
        git = shutil.which("git")
        if git:
            raiz = Path(git).resolve().parent.parent
            candidatos.append(raiz / "bin" / "bash.exe")
            candidatos.append(raiz / "usr" / "bin" / "bash.exe")
    except Exception:
        pass
    for cand in candidatos:
        if cand.exists():
            return cand
    return None


def _bash_escape(valor):
    return str(valor).replace("'", "'\\''")


def gerar_csr_itau(client_id, nome_posto, cidade, estado, pasta_saida, nome_arquivo=None):
    base_saida = Path(pasta_saida)
    base_saida.mkdir(parents=True, exist_ok=True)
    if not nome_arquivo:
        nome_arquivo = (nome_posto or "itau").strip().replace(" ", "_")
    pasta_posto = base_saida / (nome_posto.strip() or nome_arquivo)
    pasta_posto.mkdir(parents=True, exist_ok=True)
    csr_path = pasta_posto / f"{nome_arquivo}.csr"
    key_path = pasta_posto / f"{nome_arquivo}.key"

    bash = _git_bash_exe()
    nome_esc = _bash_escape(nome_arquivo)
    client_esc = _bash_escape(client_id)
    posto_esc = _bash_escape(nome_posto)
    cidade_esc = _bash_escape(cidade)
    estado_esc = _bash_escape(estado)
    # Remove arquivos antigos para evitar reutilizacao
    try:
        if csr_path.exists():
            csr_path.unlink()
        if key_path.exists():
            key_path.unlink()
    except Exception:
        pass

    if bash is not None:
        cmd = (
            f"openssl req -new -subj '//CN={client_esc}\\OU={posto_esc}\\L={cidade_esc}\\ST={estado_esc}\\C=BR' "
            f"-out '{nome_esc}.csr' -nodes -sha512 -newkey rsa:2048 -keyout '{nome_esc}.key'"
        )
        proc = subprocess.run(
            [str(bash), "-c", cmd],
            cwd=str(pasta_posto), capture_output=True, text=True, timeout=180)
    else:
        # Fallback para o openssl empacotado, caso o Git Bash nao seja encontrado
        subject = f"/CN={client_id}/OU={nome_posto}/L={cidade}/ST={estado}/C=BR"
        cwd = str(OPENSSL_BIN) if OPENSSL_BIN.exists() else None
        proc = subprocess.run(
            [str(OPENSSL_EXE), "req", "-new", "-subj", subject,
             "-out", str(csr_path), "-nodes", "-sha512",
             "-newkey", "rsa:2048", "-keyout", str(key_path)],
            cwd=cwd, capture_output=True, text=True, timeout=120,
            env=_env_openssl())
    if proc.returncode != 0:
        raise RuntimeError(f"Falha ao gerar CSR/KEY:\n{proc.stderr or proc.stdout}")
    if not csr_path.exists():
        raise RuntimeError("CSR nao foi gerado.")
    csr_content = csr_path.read_text(encoding="utf-8", errors="replace")
    return csr_path, key_path, csr_content


class GerarCertificado:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerar Certificado")
        self.root.geometry("760x620")
        self.root.minsize(640, 520)

        self.pfx_path = ""
        self.senha = ""
        self.senha_valida = False
        self.executando = False
        self.pasta_saida = AUTOMACAO_DIR if AUTOMACAO_DIR.exists() else DIR_RAIZ

        self.var_itau_client_id = tk.StringVar()
        self.var_itau_posto = tk.StringVar()
        self.var_itau_cidade = tk.StringVar()
        self.var_itau_estado = tk.StringVar()
        self.var_itau_saida = tk.StringVar(value=str(self.pasta_saida))
        self.var_itau_arquivo = tk.StringVar()
        self.var_itau_token = tk.StringVar()
        self._itau_csr_path = None
        self._itau_nome_arquivo = None
        self._itau_pasta_posto = None

        self._construir_ui()
        tema.aplicar_tema(self.root, TEMA)

    def _construir_ui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=(8, 4))

        tab_pfx = ttk.Frame(self.notebook)
        tab_itau = ttk.Frame(self.notebook)
        self.notebook.add(tab_pfx, text="PFX -> CRT/PEM/KEY")
        self.notebook.add(tab_itau, text="Itaú (CSR)")

        self._construir_tab_pfx(tab_pfx)
        self._construir_tab_itau(tab_itau)

        bottom = ttk.Frame(self.root)
        bottom.pack(fill=tk.BOTH, expand=False, padx=8, pady=(0, 8))

        frame_status = ttk.Frame(bottom)
        frame_status.pack(fill=tk.X, pady=(0, 2))
        self.lbl_indicador = ttk.Label(frame_status, text="?", foreground="#4caf50", font=("Segoe UI", 11))
        self.lbl_indicador.pack(side=tk.LEFT, padx=(0, 4))
        self.lbl_status = ttk.Label(frame_status, text="Aguardando selecao", foreground=_cor_mutado(), font=("Segoe UI", 9, "bold"))
        self.lbl_status.pack(side=tk.LEFT)
        self.progress = ttk.Progressbar(frame_status, mode="indeterminate", length=120)
        self.progress.pack(side=tk.LEFT, padx=12)
        self.progress.pack_forget()

        frame_log = ttk.LabelFrame(bottom, text="Log", padding=4)
        frame_log.pack(fill=tk.BOTH, expand=True)
        self.log = scrolledtext.ScrolledText(frame_log, height=6, state=tk.DISABLED, font=("Consolas", 9))
        self.log.pack(fill=tk.BOTH, expand=True)

    def _construir_tab_pfx(self, parent):
        frame_pfx = ttk.LabelFrame(parent, text="Certificado (.pfx)", padding=8)
        frame_pfx.pack(fill=tk.X, padx=8, pady=(8, 4))

        linha = ttk.Frame(frame_pfx)
        linha.pack(fill=tk.X)
        ttk.Button(linha, text="Selecionar .pfx", command=self._selecionar_pfx).pack(side=tk.LEFT)
        self.lbl_pfx = ttk.Label(linha, text="Nenhum arquivo selecionado", foreground=_cor_mutado())
        self.lbl_pfx.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        frame_senha = ttk.LabelFrame(parent, text="Senha do certificado", padding=8)
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

        frame_saida = ttk.LabelFrame(parent, text="Pasta de saida", padding=8)
        frame_saida.pack(fill=tk.X, padx=8, pady=2)

        linha3 = ttk.Frame(frame_saida)
        linha3.pack(fill=tk.X)
        self.ent_saida = ttk.Entry(linha3, width=60)
        self.ent_saida.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.ent_saida.insert(0, str(self.pasta_saida))
        ttk.Button(linha3, text="...", command=self._escolher_saida, width=4).pack(side=tk.LEFT, padx=(4, 0))

        frame_acoes = ttk.Frame(parent)
        frame_acoes.pack(fill=tk.X, padx=8, pady=6)
        self.btn_gerar = ttk.Button(frame_acoes, text="Gerar Arquivos (3 comandos)", command=self._gerar)
        self.btn_gerar.pack(side=tk.LEFT)
        self.btn_gerar.state(["disabled"])
        self.btn_abrir = ttk.Button(frame_acoes, text="Abrir pasta de saida", command=self._abrir_pasta)
        self.btn_abrir.pack(side=tk.LEFT, padx=8)

    def _construir_tab_itau(self, parent):
        frame_dados = ttk.LabelFrame(parent, text="Dados do posto (Itaú)", padding=8)
        frame_dados.pack(fill=tk.X, padx=8, pady=(8, 4))

        def linha_campo(rotulo, var, show=None):
            row = ttk.Frame(frame_dados)
            row.pack(fill=tk.X, pady=2)
            ttk.Label(row, text=rotulo, width=18).pack(side=tk.LEFT)
            ent = ttk.Entry(row, textvariable=var, show=show)
            ent.pack(side=tk.LEFT, fill=tk.X, expand=True)
            return ent

        self.ent_itau_client_id = linha_campo("Client ID:", self.var_itau_client_id)
        self.ent_itau_posto = linha_campo("Nome do posto:", self.var_itau_posto)
        self.ent_itau_cidade = linha_campo("Cidade:", self.var_itau_cidade)

        row_estado = ttk.Frame(frame_dados)
        row_estado.pack(fill=tk.X, pady=2)
        ttk.Label(row_estado, text="Estado:", width=18).pack(side=tk.LEFT)
        self.cmb_itau_estado = ttk.Combobox(
            row_estado, textvariable=self.var_itau_estado, state="readonly",
            values=["AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA",
                    "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN",
                    "RS", "RO", "RR", "SC", "SP", "SE", "TO"])
        self.cmb_itau_estado.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.var_itau_estado.set("SP")

        self.ent_itau_arquivo = linha_campo("Nome do arquivo:", self.var_itau_arquivo)
        ttk.Label(frame_dados, text="(deixe vazio para usar o nome do posto com espaços trocados por _)", foreground=_cor_mutado(), font=("Segoe UI", 8)).pack(anchor="w", pady=(0, 2))

        frame_saida = ttk.LabelFrame(parent, text="Pasta de saida", padding=8)
        frame_saida.pack(fill=tk.X, padx=8, pady=2)

        linha_saida = ttk.Frame(frame_saida)
        linha_saida.pack(fill=tk.X)
        self.ent_itau_saida = ttk.Entry(linha_saida, textvariable=self.var_itau_saida)
        self.ent_itau_saida.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(linha_saida, text="...", command=self._escolher_saida_itau, width=4).pack(side=tk.LEFT, padx=(4, 0))

        frame_acoes = ttk.Frame(parent)
        frame_acoes.pack(fill=tk.X, padx=8, pady=6)
        self.btn_gerar_itau = ttk.Button(frame_acoes, text="Gerar CSR e KEY (Itaú)", command=self._gerar_itau)
        self.btn_gerar_itau.pack(side=tk.LEFT)
        ttk.Button(frame_acoes, text="Abrir pasta de saida", command=self._abrir_pasta_itau).pack(side=tk.LEFT, padx=8)

        frame_api = ttk.LabelFrame(parent, text="Envio ao Itaú (endpoint)", padding=8)
        frame_api.pack(fill=tk.X, padx=8, pady=2)

        row_token = ttk.Frame(frame_api)
        row_token.pack(fill=tk.X, pady=2)
        ttk.Label(row_token, text="Bearer Token:", width=18).pack(side=tk.LEFT)
        self.ent_itau_token = ttk.Entry(row_token, textvariable=self.var_itau_token, show="*")
        self.ent_itau_token.pack(side=tk.LEFT, fill=tk.X, expand=True)

        frame_api_acoes = ttk.Frame(frame_api)
        frame_api_acoes.pack(fill=tk.X, pady=(4, 0))
        self.btn_enviar_itau = ttk.Button(frame_api_acoes, text="Enviar CSR ao Itaú", command=self._enviar_itau)
        self.btn_enviar_itau.pack(side=tk.LEFT)

        frame_textos = ttk.Frame(parent)
        frame_textos.pack(fill=tk.BOTH, expand=True, padx=8, pady=(4, 8))

        frame_csr = ttk.LabelFrame(frame_textos, text="CSR gerado", padding=8)
        frame_csr.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))
        self.txt_csr = scrolledtext.ScrolledText(frame_csr, height=10, font=("Consolas", 9))
        self.txt_csr.pack(fill=tk.BOTH, expand=True)

        frame_retorno = ttk.LabelFrame(frame_textos, text="Retorno do banco", padding=8)
        frame_retorno.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))
        self.txt_retorno = scrolledtext.ScrolledText(frame_retorno, height=10, font=("Consolas", 9))
        self.txt_retorno.pack(fill=tk.BOTH, expand=True)

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

    def _escolher_saida_itau(self):
        pasta = filedialog.askdirectory(title="Selecionar pasta de saida", initialdir=self.var_itau_saida.get())
        if pasta:
            self.var_itau_saida.set(pasta)

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
                capture_output=True, text=True, timeout=60, env=_env_openssl())
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
            cwd=cwd, capture_output=True, text=True, timeout=120, env=_env_openssl())

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

    def _gerar_itau(self):
        client_id = self.var_itau_client_id.get().strip()
        nome_posto = self.var_itau_posto.get().strip()
        cidade = self.var_itau_cidade.get().strip()
        estado = self.var_itau_estado.get().strip()
        if not client_id or not nome_posto or not cidade or not estado:
            messagebox.showwarning("Aviso", "Preencha Client ID, nome do posto, cidade e estado.")
            return
        base_saida = Path(self.var_itau_saida.get().strip() or str(self.pasta_saida))
        nome_arquivo = self.var_itau_arquivo.get().strip()
        self.btn_gerar_itau.state(["disabled"])
        self._set_status("Gerando CSR e KEY (Itaú)...", "rodando")
        threading.Thread(
            target=self._gerar_itau_thread,
            args=(client_id, nome_posto, cidade, estado, base_saida, nome_arquivo),
            daemon=True).start()

    def _gerar_itau_thread(self, client_id, nome_posto, cidade, estado, base_saida, nome_arquivo):
        try:
            csr_path, key_path, csr_content = gerar_csr_itau(
                client_id, nome_posto, cidade, estado, base_saida, nome_arquivo or None)
            self.root.after(0, lambda: self._itau_sucesso(csr_path, key_path, csr_content))
        except Exception as e:
            self.root.after(0, lambda: self._itau_erro(str(e)))

    def _itau_sucesso(self, csr_path, key_path, csr_content):
        self._itau_csr_path = Path(csr_path)
        self._itau_nome_arquivo = Path(csr_path).stem
        self._itau_pasta_posto = Path(csr_path).parent
        self.txt_csr.delete("1.0", tk.END)
        self.txt_csr.insert(tk.END, csr_content)
        self.txt_retorno.delete("1.0", tk.END)
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(csr_content)
        except Exception:
            pass
        self._log(f"OK: {csr_path}")
        self._log(f"OK: {key_path}")
        self._log("CSR copiado para a area de transferencia.")
        self._set_status("CSR e KEY gerados com sucesso (Itaú).", "ok")
        self.btn_gerar_itau.state(["!disabled"])

    def _itau_erro(self, erro):
        self._set_status("Erro na geracao (Itaú).", "erro")
        self._log(f"ERRO: {erro}")
        self.btn_gerar_itau.state(["!disabled"])
        messagebox.showerror("Erro", erro)

    def _enviar_itau(self):
        token = self.var_itau_token.get().strip()
        if not token:
            messagebox.showwarning("Aviso", "Informe o Bearer Token.")
            return
        if not self._itau_csr_path or not Path(self._itau_csr_path).exists():
            messagebox.showwarning("Aviso", "Gere o CSR antes de enviar.")
            return
        self.btn_enviar_itau.state(["disabled"])
        self._set_status("Enviando CSR ao Itaú...", "rodando")
        threading.Thread(target=self._enviar_itau_thread, args=(token,), daemon=True).start()

    def _enviar_itau_thread(self, token):
        try:
            csr_path = Path(self._itau_csr_path)
            pasta = Path(self._itau_pasta_posto)
            nome = self._itau_nome_arquivo
            csr_content = csr_path.read_text(encoding="utf-8", errors="replace")
            url = "https://sts.itau.com.br/seguranca/v1/certificado/solicitacao"
            req = urllib.request.Request(url, data=csr_content.encode("utf-8"), method="POST")
            req.add_header("Authorization", f"Bearer {token}")
            req.add_header("Content-Type", "text/plain")
            try:
                with urllib.request.urlopen(req, timeout=120) as resp:
                    resposta = resp.read().decode("utf-8", errors="replace")
            except urllib.error.HTTPError as e:
                corpo = e.read().decode("utf-8", errors="replace")
                raise RuntimeError(f"HTTP {e.code}: {corpo}")
            # Salva retorno completo
            (pasta / f"{nome}_retorno.txt").write_text(resposta, encoding="utf-8")
            # Extrai secret
            secret = ""
            m_secret = re.search(r"Secret:\s*(.+)", resposta)
            if m_secret:
                secret = m_secret.group(1).strip()
                (pasta / f"{nome}_secret.txt").write_text(secret, encoding="utf-8")
            # Extrai certificado
            m_cert = re.search(r"-----BEGIN CERTIFICATE-----.*?-----END CERTIFICATE-----", resposta, re.DOTALL)
            if m_cert:
                (pasta / f"{nome}.crt").write_text(m_cert.group(0).strip() + "\n", encoding="utf-8")
            self.root.after(0, lambda: self._enviar_itau_sucesso(pasta, secret, bool(m_cert), resposta))
        except Exception as e:
            self.root.after(0, lambda: self._enviar_itau_erro(str(e)))

    def _enviar_itau_sucesso(self, pasta, secret, tem_cert, resposta):
        self.txt_retorno.delete("1.0", tk.END)
        self.txt_retorno.insert(tk.END, resposta)
        self._set_status("Resposta do Itaú salva com sucesso.", "ok")
        self.btn_enviar_itau.state(["!disabled"])
        msg = f"Arquivos salvos em:\n{pasta}"
        if secret:
            msg += f"\n\nSecret: {secret}"
        if not tem_cert:
            msg += "\n\nAviso: certificado não encontrado na resposta."
        messagebox.showinfo("Sucesso", msg)
        try:
            os.startfile(str(pasta))
        except Exception:
            pass

    def _enviar_itau_erro(self, erro):
        self._set_status("Erro no envio ao Itaú.", "erro")
        self._log(f"ERRO: {erro}")
        self.btn_enviar_itau.state(["!disabled"])
        messagebox.showerror("Erro", erro)

    def _copiar_csr(self):
        content = self.txt_csr.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("Aviso", "Nenhum CSR gerado.")
            return
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            self._log("CSR copiado para a area de transferencia.")
        except Exception as e:
            messagebox.showerror("Erro", f"Nao foi possivel copiar.\n{e}")

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

    def _abrir_pasta_itau(self):
        pasta = self.var_itau_saida.get().strip()
        alvo = Path(pasta)
        if not alvo.exists():
            alvo.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(alvo))
        except Exception as e:
            messagebox.showerror("Erro", f"Nao foi possivel abrir a pasta.\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = GerarCertificado(root)
    root.mainloop()
