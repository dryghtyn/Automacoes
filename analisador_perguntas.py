import subprocess
import importlib.util
import sys
import os
import json
import base64
import datetime
import threading
import time
import ctypes
from pathlib import Path
from tkinter import ttk, filedialog, messagebox, scrolledtext
import tkinter as tk

DEPENDENCIAS = [
    ("PIL", "Pillow"),
    ("pypdf", "pypdf"),
    ("docx", "python-docx"),
    ("pandas", "pandas"),
]

for nome_import, nome_pip in DEPENDENCIAS:
    if importlib.util.find_spec(nome_import) is None:
        print(f"Instalando {nome_pip}...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", nome_pip, "--quiet"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

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


class AnalisadorPerguntas:
    def __init__(self, root):
        self.root = root
        self.root.title("Analisador de Perguntas de Cliente")
        self.root.geometry("1000x720")
        self.root.minsize(800, 560)

        self.pasta_base = Path(__file__).resolve().parent
        self.pasta_auto = self.pasta_base / "automacoes" / "Analisador de Perguntas"
        self.pasta_registros = self.pasta_auto / "registros"
        try:
            self.pasta_registros.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        self.sessao_id = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.arquivo_historico = self.pasta_registros / f"historico_{self.sessao_id}.txt"

        self.anexos = []  # lista de dict: {path, nome, tipo, texto, b64, mime}
        self.conversa = []  # lista de dict: {user, model}
        self._ia_finalizada = False
        self._ultima_pergunta = ""

        self._construir_ui()
        tema.aplicar_tema(self.root, TEMA)

    def _construir_ui(self):
        frame_topo = ttk.Frame(self.root, padding=8)
        frame_topo.pack(fill=tk.X)

        ttk.Label(frame_topo, text="Analisador de Perguntas de Cliente",
                  font=("Segoe UI", 13, "bold")).pack(side=tk.LEFT)
        self.btn_config = ttk.Button(frame_topo, text="Configurar chave",
                                     command=self._configurar_chave)
        self.btn_config.pack(side=tk.RIGHT)

        frame_principal = ttk.Frame(self.root, padding=8)
        frame_principal.pack(fill=tk.BOTH, expand=True)

        # Esquerda: pergunta + anexos
        frame_esq = ttk.LabelFrame(frame_principal, text="Pergunta / Anexos", padding=6)
        frame_esq.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))

        ttk.Label(frame_esq, text="Pergunta (copie e cole):").pack(anchor=tk.W)
        self.txt_pergunta = scrolledtext.ScrolledText(frame_esq, height=6,
                                                      font=("Segoe UI", 10))
        self.txt_pergunta.pack(fill=tk.BOTH, expand=True, pady=(2, 6))

        frame_anexos = ttk.Frame(frame_esq)
        frame_anexos.pack(fill=tk.X, pady=(0, 6))
        self.btn_anexar = ttk.Button(frame_anexos, text="Anexar arquivo(s)",
                                     command=self._anexar_arquivos)
        self.btn_anexar.pack(side=tk.LEFT)
        self.btn_limpar_anexos = ttk.Button(frame_anexos, text="Limpar anexos",
                                            command=self._limpar_anexos)
        self.btn_limpar_anexos.pack(side=tk.LEFT, padx=6)
        self.lbl_anexos = ttk.Label(frame_anexos, text="Nenhum anexo",
                                    foreground=_cor_mutado(), font=("Segoe UI", 9))
        self.lbl_anexos.pack(side=tk.LEFT, padx=6)

        self.lista_anexos = tk.Listbox(frame_esq, height=4, font=("Consolas", 9))
        self.lista_anexos.pack(fill=tk.X, pady=(0, 6))

        frame_acao = ttk.Frame(frame_esq)
        frame_acao.pack(fill=tk.X)
        self.btn_enviar = ttk.Button(frame_acao, text="Enviar pergunta",
                                     command=self._enviar)
        self.btn_enviar.pack(side=tk.LEFT)
        self.btn_limpar_conv = ttk.Button(frame_acao, text="Limpar conversa",
                                         command=self._limpar_conversa)
        self.btn_limpar_conv.pack(side=tk.LEFT, padx=6)
        self.lbl_status = ttk.Label(frame_acao, text="Pronto.",
                                    foreground=_cor_mutado(), font=("Segoe UI", 9, "bold"))
        self.lbl_status.pack(side=tk.LEFT, padx=10)

        # Direita: resposta
        frame_dir = ttk.LabelFrame(frame_principal, text="Resposta da IA", padding=6)
        frame_dir.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.txt_resposta = scrolledtext.ScrolledText(frame_dir, height=20,
                                                      font=("Consolas", 10),
                                                      state=tk.DISABLED)
        self.txt_resposta.pack(fill=tk.BOTH, expand=True)

    # ---------- chave ----------
    def _caminho_config(self):
        return self.pasta_auto / "gemini_config.json"

    def _carregar_chave(self):
        caminho = self._caminho_config()
        try:
            if caminho.exists():
                with open(caminho, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                return (dados.get("api_key") or "").strip()
        except Exception:
            return ""
        return ""

    def _salvar_chave(self, chave):
        try:
            self.pasta_auto.mkdir(parents=True, exist_ok=True)
            with open(self._caminho_config(), "w", encoding="utf-8") as f:
                json.dump({"api_key": chave.strip()}, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            self._log(f"Erro ao salvar chave: {e}")
            return False

    def _configurar_chave(self):
        janela = tk.Toplevel(self.root)
        janela.title("Configurar chave API Gemini")
        janela.geometry("460x170")
        janela.transient(self.root)
        janela.grab_set()
        frame = ttk.Frame(janela, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="Chave API do Google Gemini:").pack(anchor="w")
        var = tk.StringVar(value=self._carregar_chave())
        entry = ttk.Entry(frame, textvariable=var, width=60, show="*")
        entry.pack(fill=tk.X, pady=(4, 8))
        ttk.Label(frame, text="A chave fica salva em arquivo local.",
                 foreground=_cor_mutado(), font=("Segoe UI", 9)).pack(anchor="w")

        def salvar():
            chave = var.get().strip()
            if not chave:
                messagebox.showwarning("Aviso", "Informe a chave API.")
                return
            if self._salvar_chave(chave):
                messagebox.showinfo("Salvo", "Chave API salva.")
                janela.destroy()
            else:
                messagebox.showerror("Erro", "Não foi possível salvar a chave.")

        def cancelar():
            janela.destroy()

        fb = ttk.Frame(frame)
        fb.pack(pady=8)
        ttk.Button(fb, text="Salvar", command=salvar).pack(side=tk.LEFT, padx=4)
        ttk.Button(fb, text="Cancelar", command=cancelar).pack(side=tk.LEFT, padx=4)
        janela.wait_window(janela)

    # ---------- anexos ----------
    def _anexar_arquivos(self):
        caminhos = filedialog.askopenfilenames(
            title="Selecionar arquivos",
            filetypes=[("Todos", "*.*"),
                       ("Texto", "*.txt *.csv *.md"),
                       ("PDF", "*.pdf"),
                       ("Word", "*.docx"),
                       ("Excel", "*.xlsx *.xls"),
                       ("Imagens", "*.png *.jpg *.jpeg *.webp *.bmp *.gif"),
                       ("Áudio", "*.mp3 *.wav *.ogg *.m4a *.aac *.flac")])
        if not caminhos:
            return
        for c in caminhos:
            self._processar_anexo(c)
        self._atualizar_lista_anexos()

    def _processar_anexo(self, caminho):
        path = Path(caminho)
        nome = path.name
        ext = path.suffix.lower()
        anexo = {"path": str(path), "nome": nome, "ext": ext, "texto": "", "b64": None, "mime": None}
        try:
            if ext in (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif"):
                import base64 as _b64
                data = path.read_bytes()
                anexo["b64"] = _b64.b64encode(data).decode("ascii")
                mime = "image/png"
                if ext in (".jpg", ".jpeg"):
                    mime = "image/jpeg"
                elif ext == ".webp":
                    mime = "image/webp"
                elif ext == ".gif":
                    mime = "image/gif"
                anexo["mime"] = mime
                anexo["texto"] = f"[Imagem anexada: {nome}]"
            elif ext in (".mp3", ".wav", ".ogg", ".m4a", ".aac", ".flac"):
                import base64 as _b64
                data = path.read_bytes()
                anexo["b64"] = _b64.b64encode(data).decode("ascii")
                mime = "audio/mpeg"
                if ext == ".wav":
                    mime = "audio/wav"
                elif ext == ".ogg":
                    mime = "audio/ogg"
                elif ext == ".m4a":
                    mime = "audio/mp4"
                elif ext == ".aac":
                    mime = "audio/aac"
                elif ext == ".flac":
                    mime = "audio/flac"
                anexo["mime"] = mime
                anexo["texto"] = f"[Áudio anexado: {nome}]"
            elif ext in (".txt", ".md", ".csv"):
                anexo["texto"] = path.read_text(encoding="utf-8", errors="replace")
            elif ext == ".pdf":
                try:
                    from pypdf import PdfReader
                    reader = PdfReader(str(path))
                    partes = []
                    for pg in reader.pages:
                        try:
                            partes.append(pg.extract_text() or "")
                        except Exception:
                            pass
                    anexo["texto"] = "\n".join(partes)
                except Exception as e:
                    anexo["texto"] = f"[Erro ao ler PDF: {e}]"
            elif ext == ".docx":
                try:
                    import docx
                    doc = docx.Document(str(path))
                    anexo["texto"] = "\n".join(p.text for p in doc.paragraphs)
                except Exception as e:
                    anexo["texto"] = f"[Erro ao ler DOCX: {e}]"
            elif ext in (".xlsx", ".xls"):
                try:
                    import pandas as pd
                    df = pd.read_excel(str(path))
                    anexo["texto"] = df.to_csv(index=False)
                except Exception as e:
                    anexo["texto"] = f"[Erro ao ler Excel: {e}]"
            else:
                anexo["texto"] = f"[Tipo não suportado: {ext}]"
        except Exception as e:
            anexo["texto"] = f"[Erro ao processar {nome}: {e}]"
        # truncar
        if len(anexo["texto"]) > 15000:
            anexo["texto"] = anexo["texto"][:15000] + "\n... (truncado)"
        self.anexos.append(anexo)

    def _atualizar_lista_anexos(self):
        self.lista_anexos.delete(0, tk.END)
        for a in self.anexos:
            self.lista_anexos.insert(tk.END, a["nome"])
        if self.anexos:
            self.lbl_anexos.configure(text=f"{len(self.anexos)} anexo(s)")
        else:
            self.lbl_anexos.configure(text="Nenhum anexo")

    def _limpar_anexos(self):
        self.anexos.clear()
        self._atualizar_lista_anexos()

    def _limpar_conversa(self):
        self.conversa.clear()
        self.txt_resposta.config(state=tk.NORMAL)
        self.txt_resposta.delete("1.0", tk.END)
        self.txt_resposta.config(state=tk.DISABLED)
        self._log("Conversa limpa.")

    # ---------- envio ----------
    def _enviar(self):
        pergunta = self.txt_pergunta.get("1.0", tk.END).strip()
        if not pergunta and not self.anexos:
            messagebox.showwarning("Aviso", "Digite uma pergunta ou anexe um arquivo.")
            return
        chave = self._carregar_chave()
        if not chave:
            if messagebox.askyesno("Chave API", "Chave API do Gemini não configurada. Deseja configurar agora?"):
                self._configurar_chave()
                chave = self._carregar_chave()
            if not chave:
                return

        self._ultima_pergunta = pergunta
        self.btn_enviar.config(state=tk.DISABLED)
        self.lbl_status.configure(text="Consultando Gemini...", foreground="#ff9800")
        self._adicionar_resposta(f"Pergunta: {pergunta}\n")

        self._ia_finalizada = False

        def tarefa():
            try:
                resposta = self._chamar_gemini(chave, pergunta, list(self.anexos), list(self.conversa))
                self.root.after(0, lambda: self._finalizar(resposta, None))
            except Exception as e:
                self.root.after(0, lambda: self._finalizar(None, str(e)))

        threading.Thread(target=tarefa, daemon=True).start()

        def vigia():
            time.sleep(120)
            if not self._ia_finalizada:
                self.root.after(0, lambda: self._finalizar(
                    None, "Tempo esgotado (120s) aguardando resposta."))

        threading.Thread(target=vigia, daemon=True).start()

    def _chamar_gemini(self, chave, pergunta, anexos, conversa):
        import json as _json
        import urllib.request
        import urllib.error
        import urllib.parse

        # listar modelos
        url_lista = ("https://generativelanguage.googleapis.com/v1beta/models"
                     "?key=" + urllib.parse.quote(chave, safe=""))
        with urllib.request.urlopen(url_lista, timeout=20) as resp:
            dados_lista = _json.loads(resp.read().decode("utf-8"))

        modelos = []
        for m in dados_lista.get("models", []):
            nome = (m.get("name") or "").replace("models/", "")
            metodos = m.get("supportedGenerationMethods") or []
            if "generateContent" not in metodos:
                continue
            if "gemini" not in nome.lower():
                continue
            modelos.append(nome)
        modelos.sort(key=lambda n: (0 if "flash" in n.lower() and "lite" not in n.lower() else
                                   1 if "flash" in n.lower() else 2, n))
        if not modelos:
            raise RuntimeError("Nenhum modelo disponível para esta chave.")

        # montar contents com histórico
        contents = []
        for turn in conversa:
            contents.append({"role": "user", "parts": [{"text": turn.get("user", "")}]})
            contents.append({"role": "model", "parts": [{"text": turn.get("model", "")}]})

        # partes da pergunta atual
        parts = []
        for a in anexos:
            if a.get("b64"):
                parts.append({"inline_data": {"mime_type": a.get("mime", "image/png"),
                                             "data": a["b64"]}})
            if a.get("texto"):
                parts.append({"text": f"Arquivo: {a['nome']}\n{a['texto']}"})
        if pergunta:
            parts.append({"text": pergunta})
        if not parts:
            parts.append({"text": "Analise os anexos."})
        contents.append({"role": "user", "parts": parts})

        payload = {"contents": contents}

        ultimo_erro = None
        erros = []
        for modelo in modelos[:4]:
            self._log(f"IA: tentando {modelo}...")
            endpoint = ("https://generativelanguage.googleapis.com/v1beta/models/"
                        + urllib.parse.quote(modelo, safe="") + ":generateContent?key="
                        + urllib.parse.quote(chave, safe=""))
            try:
                req = urllib.request.Request(
                    endpoint,
                    data=_json.dumps(payload).encode("utf-8"),
                    headers={"Content-Type": "application/json"},
                    method="POST",
                )
                with urllib.request.urlopen(req, timeout=60) as resp:
                    resultado = _json.loads(resp.read().decode("utf-8"))
                for cand in resultado.get("candidates", []):
                    for part in (cand.get("content") or {}).get("parts", []):
                        if "text" in part:
                            self._log(f"IA: sucesso com {modelo}.")
                            return part["text"]
                return str(resultado)
            except urllib.error.HTTPError as e:
                corpo = ""
                try:
                    corpo = e.read().decode("utf-8", errors="replace")[:300]
                except Exception:
                    pass
                self._log(f"IA: falha em {modelo} (HTTP {e.code}): {corpo[:120]}")
                erros.append(f"{modelo}: HTTP {e.code}")
                ultimo_erro = RuntimeError(f"HTTP {e.code}: {corpo}")
                continue
            except Exception as e:
                self._log(f"IA: falha em {modelo}: {str(e)[:120]}")
                erros.append(f"{modelo}: {str(e)[:60]}")
                ultimo_erro = e
                continue

        raise RuntimeError("Nenhum modelo funcionou. " + " | ".join(erros[:6]))

    def _finalizar(self, resposta, erro):
        if getattr(self, "_ia_finalizada", False):
            return
        self._ia_finalizada = True
        self.btn_enviar.config(state=tk.NORMAL)
        if erro:
            self.lbl_status.configure(text="Erro na análise.", foreground="#e53935")
            self._adicionar_resposta(f"Erro: {erro}\n")
            self._log(f"IA: erro - {erro}")
            # restaurar pergunta
            self.txt_pergunta.delete("1.0", tk.END)
            self.txt_pergunta.insert("1.0", self._ultima_pergunta)
            self._registrar(self._ultima_pergunta, erro)
            return
        self.lbl_status.configure(text="Análise concluída.", foreground="#4caf50")
        self._adicionar_resposta(f"Resposta: {resposta}\n")
        self._log("IA: análise concluída.")
        # armazenar na conversa
        self.conversa.append({"user": self._ultima_pergunta, "model": resposta})
        # limpar pergunta
        self.txt_pergunta.delete("1.0", tk.END)
        self._registrar(self._ultima_pergunta, resposta)

    def _adicionar_resposta(self, texto):
        self.txt_resposta.config(state=tk.NORMAL)
        self.txt_resposta.insert(tk.END, texto + "\n")
        self.txt_resposta.see(tk.END)
        self.txt_resposta.config(state=tk.DISABLED)

    def _registrar(self, pergunta, resposta):
        try:
            with open(self.arquivo_historico, "a", encoding="utf-8") as f:
                f.write(f"[{datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}]\n")
                f.write(f"Pergunta: {pergunta}\n")
                f.write(f"Resposta: {resposta}\n")
                f.write("-" * 60 + "\n")
        except Exception as e:
            self._log(f"Erro ao registrar: {e}")

    def _log(self, msg):
        # simples print + status
        print(f"[{time.strftime('%H:%M:%S')}] {msg}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AnalisadorPerguntas(root)
    root.mainloop()
