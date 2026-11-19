import subprocess
import importlib.util
import sys

DEPENDENCIAS = [
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("pynput", "pynput"),
    ("serial", "pyserial"),
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
import ctypes
import os
from datetime import datetime
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
    from pynput import keyboard
    HAVE_PYNPUT = True
except ImportError:
    HAVE_PYNPUT = False

try:
    import serial
    import serial.tools.list_ports as list_ports
    HAVE_SERIAL = True
except ImportError:
    HAVE_SERIAL = False

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


class LeitorIdentifid:
    def __init__(self, root):
        self.root = root
        self.root.title("Leitor Identifid")
        self.root.geometry("680x560")
        self.root.minsize(560, 420)

        self.listener = None
        self.thread_escuta = None
        self.thread_serial = None
        self.ser = None
        self.fonte = "keyboard"
        self.porta_serial = ""
        self.arquivo = ""
        self.caminho_planilha = ""
        self.linhas = []
        self.buffer = ""
        self.gravando = False
        self.sequencia = 0

        self._construir_ui()
        tema.aplicar_tema(self.root, TEMA)

    def _construir_ui(self):
        frame_formato = ttk.LabelFrame(self.root, text="Formato de Saída", padding=8)
        frame_formato.pack(fill=tk.X, padx=8, pady=(8, 4))

        ttk.Label(frame_formato, text="Formato:").pack(side=tk.LEFT)
        self.cb_formato = ttk.Combobox(frame_formato, state="readonly", width=22,
                                        values=["Planilha (.xlsx)", "Arquivo de Texto (.txt)"])
        self.cb_formato.set("Planilha (.xlsx)")
        self.cb_formato.pack(side=tk.LEFT, padx=(4, 12))

        frame_cmd = ttk.LabelFrame(self.root, text="Controle", padding=8)
        frame_cmd.pack(fill=tk.X, padx=8, pady=2)

        self.btn_iniciar = ttk.Button(frame_cmd, text="Criar Arquivo",
                                       command=self._criar_planilha)
        self.btn_iniciar.pack(side=tk.LEFT)
        self.btn_abrir = ttk.Button(frame_cmd, text="Abrir Arquivo Existente",
                                     command=self._abrir_planilha)
        self.btn_abrir.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_parar = ttk.Button(frame_cmd, text="Parar / Fechar",
                                    command=self._parar_fechar, state=tk.DISABLED)
        self.btn_parar.pack(side=tk.LEFT, padx=8)

        self.lbl_arquivo = ttk.Label(frame_cmd, text="Nenhum arquivo",
                                      foreground=_cor_mutado())
        self.lbl_arquivo.pack(side=tk.LEFT, padx=(12, 0))

        frame_tipo = ttk.LabelFrame(self.root, text="Tipo de Leitor", padding=8)
        frame_tipo.pack(fill=tk.X, padx=8, pady=2)

        ttk.Label(frame_tipo, text="Modelo:").pack(side=tk.LEFT)
        self.cb_tipo = ttk.Combobox(frame_tipo, state="readonly", width=16,
                                    values=["Companytec", "Eztech"])
        self.cb_tipo.set("Companytec")
        self.tipo_leitor = "companytec"
        self.cb_tipo.bind("<<ComboboxSelected>>", self._mudar_tipo)
        self.cb_tipo.pack(side=tk.LEFT, padx=(4, 12))

        self.lbl_tipo_info = ttk.Label(frame_tipo, text="",
                                       foreground=_cor_mutado())
        self.lbl_tipo_info.pack(side=tk.LEFT)

        frame_fonte = ttk.LabelFrame(self.root, text="Fonte de Leitura", padding=8)
        frame_fonte.pack(fill=tk.X, padx=8, pady=2)

        ttk.Label(frame_fonte, text="Fonte:").pack(side=tk.LEFT)
        self.cb_fonte = ttk.Combobox(frame_fonte, state="readonly", width=22,
                                     values=["Teclado (USB / HID)", "Porta COM (Serial)"])
        self.cb_fonte.set("Teclado (USB / HID)")
        self.cb_fonte.bind("<<ComboboxSelected>>", self._mudar_fonte)
        self.cb_fonte.pack(side=tk.LEFT, padx=(4, 16))

        self.lbl_porta = ttk.Label(frame_fonte, text="Porta COM:")
        self.cb_porta = ttk.Combobox(frame_fonte, state="readonly", width=14)
        self.btn_detectar = ttk.Button(frame_fonte, text="Detectar",
                                       command=self._tentar_calibrar)
        self.cb_porta.pack(side=tk.LEFT, padx=(4, 4))
        self.btn_detectar.pack(side=tk.LEFT, padx=(0, 8))

        self.lbl_porta_info = ttk.Label(frame_fonte, text="", foreground=_cor_mutado())
        self.lbl_porta_info.pack(side=tk.LEFT)
        self._detectar_portas(autoselecionar=False)

        frame_table = ttk.LabelFrame(self.root, text="Leituras", padding=6)
        frame_table.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        cols = ("seq", "cartao", "datahora")
        self.tree = ttk.Treeview(frame_table, columns=cols, show="headings", height=12)
        self.tree.heading("seq", text="Seq")
        self.tree.heading("cartao", text="Cartão")
        self.tree.heading("datahora", text="Data/Hora")
        self.tree.column("seq", width=70, anchor=tk.CENTER)
        self.tree.column("cartao", width=250, anchor=tk.CENTER)
        self.tree.column("datahora", width=200, anchor=tk.CENTER)
        scroll_y = ttk.Scrollbar(frame_table, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        frame_status = ttk.Frame(self.root)
        frame_status.pack(fill=tk.X, padx=8, pady=(0, 2))
        self.lbl_indicador = ttk.Label(frame_status, text="●", foreground=_cor_mutado(),
                                       font=("Segoe UI", 11))
        self.lbl_indicador.pack(side=tk.LEFT, padx=(0, 4))
        self.lbl_status = ttk.Label(frame_status, text="Parado. Clique em 'Criar Arquivo'.",
                                     foreground=_cor_mutado(), font=("Segoe UI", 9, "bold"))
        self.lbl_status.pack(side=tk.LEFT)
        self.lbl_contador = ttk.Label(frame_status, text="0 leituras", foreground=_cor_mutado())
        self.lbl_contador.pack(side=tk.RIGHT)

        frame_log = ttk.LabelFrame(self.root, text="Log", padding=4)
        frame_log.pack(fill=tk.BOTH, padx=8, pady=(0, 8))
        self.log = scrolledtext.ScrolledText(frame_log, height=5, state=tk.DISABLED, font=("Consolas", 9))
        self.log.pack(fill=tk.BOTH, expand=True)

    def _log(self, msg):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)

    def _set_status(self, texto, cor):
        self.lbl_indicador.configure(foreground=cor)
        self.lbl_status.configure(text=texto)

    def _criar_planilha(self):
        if self.gravando:
            messagebox.showinfo("Aviso", "Já está gravando. Use 'Parar / Fechar' primeiro.")
            return
        formato = self.cb_formato.get()
        if "Texto" in formato:
            caminho = filedialog.asksaveasfilename(
                title="Onde criar o arquivo de texto",
                defaultextension=".txt",
                initialfile=f"identifid_{datetime.now():%Y%m%d_%H%M%S}.txt",
                filetypes=[("Arquivo de Texto", "*.txt"), ("Todos", "*.*")])
        else:
            caminho = filedialog.asksaveasfilename(
                title="Onde criar a planilha de leituras",
                defaultextension=".xlsx",
                initialfile=f"identifid_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                filetypes=[("Excel", "*.xlsx")])
        if not caminho:
            return
        try:
            self.linhas = []
            self._gravar_arquivo(path=caminho, relatorio=False)
            self._log(f"Arquivo criado: {caminho}")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível criar o arquivo.\n{e}")
            return
        self._iniciar_gravacao(caminho, carregar_existente=False)

    def _abrir_planilha(self):
        if self.gravando:
            messagebox.showinfo("Aviso", "Já está gravando. Use 'Parar / Fechar' primeiro.")
            return
        formato = self.cb_formato.get()
        if "Texto" in formato:
            caminho = filedialog.askopenfilename(
                title="Abrir arquivo de texto existente",
                filetypes=[("Arquivo de Texto", "*.txt"),
                           ("Excel", "*.xlsx *.xls"),
                           ("CSV", "*.csv"),
                           ("Todos", "*.*")])
        else:
            caminho = filedialog.askopenfilename(
                title="Abrir planilha existente",
                filetypes=[("Excel", "*.xlsx *.xls"),
                           ("Arquivo de Texto", "*.txt"),
                           ("CSV", "*.csv"),
                           ("Todos", "*.*")])
        if not caminho:
            return
        try:
            self.linhas = self._carregar_existentes(caminho)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler o arquivo.\n{e}")
            return
        self._log(f"Arquivo aberto: {caminho} com {len(self.linhas)} registro(s) existente(s).")
        if str(caminho).lower().endswith(".txt"):
            self.cb_formato.set("Arquivo de Texto (.txt)")
        else:
            self.cb_formato.set("Planilha (.xlsx)")
        self._iniciar_gravacao(caminho, carregar_existente=True)

    def _iniciar_gravacao(self, caminho, carregar_existente):
        self.caminho_planilha = caminho
        self.arquivo = os.path.basename(caminho)
        self.lbl_arquivo.configure(text=self.arquivo)
        self.sequencia = max((int(l["seq"]) for l in self.linhas), default=0)
        self.buffer = ""
        self._limpar_tree()
        for l in self.linhas:
            self.tree.insert("", "end",
                             values=(l["seq"], l["cartao"], l["datahora"]))
        if self.tree.get_children():
            self.tree.see(self.tree.get_children()[-1])
        self.lbl_contador.configure(text=f"{len(self.linhas)} leitura(s)")
        self.gravando = True
        self._set_status("Preparando leitura...", "#ff9800")
        self.btn_iniciar.state(["disabled"])
        self.btn_abrir.state(["disabled"])
        self.btn_parar.state(["!disabled"])
        self._iniciar_captura()
        self._log("Leitura iniciada. Aguardando leitura de cartão...")

    def _iniciar_captura(self):
        if self.fonte == "serial":
            self._iniciar_serial()
        else:
            self._iniciar_listener()

    def _carregar_existentes(self, caminho):
        if str(caminho).lower().endswith(".txt"):
            return self._carregar_existentes_txt(caminho)
        colunas = {"seq": "seq", "cartao": "cartao", "datahora": "datahora"}
        try:
            if str(caminho).lower().endswith(".csv"):
                df = pd.read_csv(caminho, encoding="utf-8-sig", dtype=str)
            else:
                df = pd.read_excel(caminho, engine="openpyxl", dtype=str)
        except Exception:
            if str(caminho).lower().endswith(".csv"):
                df = pd.read_csv(caminho, encoding="latin-1", dtype=str)
            else:
                df = pd.read_excel(caminho, dtype=str)

        nomes = {str(c).strip().lower(): str(c) for c in df.columns}
        for padrao, chave in (("seq", "seq"), ("sequencia", "seq"),
                              ("cartão", "cartao"), ("cartao", "cartao"),
                              ("id", "cartao"), ("data/hora", "datahora"),
                              ("datahora", "datahora"), ("data hora", "datahora")):
            if padrao in nomes and chave not in colunas:
                colunas[chave] = nomes[padrao]

        linhas = []
        for _, linha in df.iterrows():
            seq = linha.get(colunas.get("seq"))
            cartao = linha.get(colunas.get("cartao"))
            datahora = linha.get(colunas.get("datahora"))
            try:
                seq_i = int(float(seq)) if seq not in (None, "", float("nan")) else 0
            except Exception:
                seq_i = 0
            linhas.append({
                "seq": seq_i,
                "cartao": str(cartao).strip() if cartao is not None and str(cartao) != "nan" else "",
                "datahora": str(datahora) if datahora is not None and str(datahora) != "nan" else "",
            })
        return linhas

    def _carregar_existentes_txt(self, caminho):
        linhas = []
        with open(caminho, "r", encoding="utf-8") as f:
            for numero_linha, linha in enumerate(f, start=1):
                linha = linha.rstrip("\n\r")
                if not linha:
                    continue
                partes = linha.split("\t")
                seq = 0
                cartao = ""
                if len(partes) >= 2:
                    try:
                        seq = int(partes[0].strip())
                    except Exception:
                        seq = numero_linha
                    cartao = partes[1].strip()
                elif len(partes) == 1:
                    seq = numero_linha
                    cartao = partes[0].strip()
                linhas.append({
                    "seq": seq,
                    "cartao": cartao,
                    "datahora": "",
                })
        return linhas

    def _mudar_tipo(self, event=None):
        modelo = self.cb_tipo.get()
        self.tipo_leitor = "companytec" if modelo == "Companytec" else "eztech"
        if self.tipo_leitor == "companytec":
            self.lbl_tipo_info.configure(
                text="Grava a partir de 'B32', terminando antes do '$' (16 caracteres).",
                foreground=_cor_mutado())
        else:
            self.lbl_tipo_info.configure(text="", foreground=_cor_mutado())

    def _processar_cartao(self, raw):
        if self.tipo_leitor == "companytec":
            i = raw.find("B32")
            if i != -1:
                parte = raw[i:]
                j = parte.find("$")
                if j != -1:
                    return parte[:j]
                return parte
            return raw
        return raw

    def _mudar_fonte(self, event=None):
        self.fonte = "serial" if self.cb_fonte.get().startswith("Porta") else "keyboard"
        if self.fonte == "serial":
            if not HAVE_SERIAL:
                messagebox.showwarning("Aviso",
                    "Biblioteca pyserial não disponível. Use a fonte 'Teclado (USB / HID)'.")
                self.cb_fonte.set("Teclado (USB / HID)")
                self.fonte = "keyboard"
                return
            self._detectar_portas(autoselecionar=True)
        else:
            self.lbl_porta_info.configure(text="")
        if self.gravando:
            self._parar_captura()
            self._iniciar_captura()

    def _detectar_portas(self, autoselecionar=False):
        if not HAVE_SERIAL:
            return
        portas = []
        try:
            for p in list_ports.comports():
                desc = p.description or ""
                portas.append(f"{p.device}  ({desc})" if desc else p.device)
        except Exception as e:
            self._log(f"Erro ao detectar portas COM: {e}")
        if not portas:
            self.cb_porta["values"] = []
            self.lbl_porta_info.configure(text="Nenhuma porta COM encontrada.", foreground="#e53935")
            return
        self.cb_porta["values"] = portas
        self.lbl_porta_info.configure(text=f"{len(portas)} porta(s) encontrada(s).", foreground=_cor_mutado())
        if autoselecionar and not self.cb_porta.get():
            self.cb_porta.current(0)

    def _tentar_calibrar(self):
        if not HAVE_SERIAL:
            messagebox.showwarning("Aviso",
                "Biblioteca pyserial não disponível. Use a fonte 'Teclado (USB / HID)'.")
            return
        portas = []
        try:
            for p in list_ports.comports():
                portas.append(p.device)
        except Exception as e:
            self._log(f"Erro ao detectar portas COM: {e}")
            return
        portas = [p for p in portas if p]
        if not portas:
            self.lbl_porta_info.configure(text="Nenhuma porta COM encontrada.", foreground="#e53935")
            return
        self.lbl_porta_info.configure(text=f"{len(portas)} porta(s) encontrada(s).", foreground=_cor_mutado())
        self._calibrar_serial(portas)

    def _calibrar_serial(self, portas):
        janela = tk.Toplevel(self.root)
        janela.title("Detectar Leitor")
        janela.resizable(False, False)
        janela.transient(self.root)
        janela.grab_set()
        tema.aplicar_tema(janela, TEMA)
        model_name = self.cb_tipo.get()
        ttk.Label(janela, text=f"Modelo: {model_name}",
                  font=("Segoe UI", 10, "bold")).pack(padx=16, pady=(14, 2))
        ttk.Label(janela, text="Passe o cartão no leitor",
                  justify=tk.LEFT).pack(padx=16, pady=2)
        self._lbl_cal = ttk.Label(janela, text="Aguardando leitura...", foreground=_cor_mutado())
        self._lbl_cal.pack(padx=16, pady=4)

        estado = {"encontrada": False, "porta": None, "cartao": None}

        def cancelar():
            estado["encontrada"] = True
            janela.destroy()

        btn_parar = ttk.Button(janela, text="Cancelar", command=cancelar)
        btn_parar.pack(pady=(4, 12))

        def escutar(porta):
            try:
                s = serial.Serial(port=porta, baudrate=9600, bytesize=8, parity="N",
                                  stopbits=1, timeout=0.1)
            except Exception:
                return
            raw = b""
            try:
                while True:
                    if estado["encontrada"]:
                        break
                    dados = s.read(256)
                    if not dados:
                        continue
                    raw += dados
                    texto = "".join(c for c in
                                    raw.decode("utf-8", errors="replace") if c.isprintable())
                    if self.tipo_leitor == "companytec":
                        if "B32" in texto:
                            estado["encontrada"] = True
                            estado["porta"] = porta
                            estado["cartao"] = texto.strip()
                            break
                    elif len(texto.strip()) >= 6:
                        estado["encontrada"] = True
                        estado["porta"] = porta
                        estado["cartao"] = texto.strip()
                        break
            except Exception:
                pass
            finally:
                try:
                    s.close()
                except Exception:
                    pass

        for p in portas:
            threading.Thread(target=escutar, args=(p,), daemon=True).start()

        def verificar():
            if janela.winfo_exists():
                if estado["encontrada"]:
                    janela.destroy()
                else:
                    janela.after(80, verificar)

        janela.after(80, verificar)
        janela.wait_window()
        if estado["porta"]:
            escolha = next((v for v in self.cb_porta["values"] if v.startswith(estado["porta"])),
                           estado["porta"])
            self.cb_porta.set(escolha)
            self.fonte = "serial"
            self.cb_fonte.set("Porta COM (Serial)")
            self.lbl_porta_info.configure(
                text=f"Leitor identificado na {estado['porta']}.", foreground="#4caf50")
            self._log(f"Leitor identificado na {estado['porta']}.")
            if self.gravando:
                self._parar_listener()
                self._iniciar_serial()
                self._log("Conexão COM retomada para novas leituras.")
        else:
            self.lbl_porta_info.configure(text="Nenhum dado recebido. Tente novamente.",
                                          foreground="#e53935")

    def _iniciar_serial(self):
        escolha = self.cb_porta.get().strip()
        if not escolha:
            messagebox.showwarning("Aviso",
                "Nenhuma porta COM selecionada. Clique em 'Detectar' e escolha a porta.")
            self.gravando = False
            self.btn_iniciar.state(["!disabled"])
            self.btn_abrir.state(["!disabled"])
            self.btn_parar.state(["disabled"])
            return
        porta = escolha.split()[0]
        try:
            self.ser = serial.Serial(port=porta, baudrate=9600, bytesize=8, parity="N",
                                     stopbits=1, timeout=0.1)
        except Exception as e:
            messagebox.showerror("Erro",
                f"Não foi possível abrir a porta {porta}.\n{e}\n"
                "Verifique se o dispositivo está conectado e se a porta não está em uso.")
            self.gravando = False
            self.btn_iniciar.state(["!disabled"])
            self.btn_abrir.state(["!disabled"])
            self.btn_parar.state(["disabled"])
            return
        self.porta_serial = porta
        self._set_status(f"Gravando via {porta} (9600 baud)...", "#4caf50")
        self.thread_serial = threading.Thread(target=self._loop_serial, daemon=True)
        self.thread_serial.start()
        self._log(f"Conectado à porta {porta}.")

    def _loop_serial(self):
        buffer = b""
        ultimo_dado = time.time()
        try:
            while self.gravando and self.ser is not None and self.ser.is_open:
                dados = self.ser.read(256)
                if dados:
                    buffer += dados
                    ultimo_dado = time.time()
                    if self._tem_delimitador(buffer):
                        buffer = self._extrair_frames(buffer)
                        continue
                else:
                    # nenhum dado neste ciclo
                    if buffer and (time.time() - ultimo_dado) > 0.25:
                        self._comitar_serial(buffer)
                        buffer = b""
        except Exception as e:
            if self.gravando:
                self.root.after(0, lambda: self._log(f"Erro na leitura serial: {e}"))
        finally:
            if buffer and self.gravando:
                self._comitar_serial(buffer)
            if self.ser is not None and self.ser.is_open:
                try:
                    self.ser.close()
                except Exception:
                    pass

    def _tem_delimitador(self, buffer):
        return any(b in buffer for b in (b"\r", b"\n", b"\x03"))

    def _extrair_frames(self, buffer):
        # Separa e registra cada "frame" (cartão) delimitado por \r, \n ou ETX (\x03).
        intenso = buffer.replace(b"\x03", b"\r").replace(b"\x02", b"")
        blocos = intenso.split(b"\r")
        partes = []
        for b in blocos:
            for p in b.split(b"\n"):
                partes.append(p)
        for parte in partes:
            if parte:
                self._comitar_serial(parte)
        return b""

    def _comitar_serial(self, dados):
        texto = dados.decode("utf-8", errors="replace")
        # remove bytes de controle / não imprimíveis
        texto = "".join(c for c in texto if c.isprintable())
        texto = texto.strip()
        if texto:
            self.root.after(0, self._registrar_cartao, texto)

    def _iniciar_listener(self):
        if not HAVE_PYNPUT:
            messagebox.showerror("Erro", "Biblioteca pynput não disponível.")
            self.gravando = False
            return
        self.listener = keyboard.Listener(on_press=self._on_press)
        self.listener.daemon = True
        self.thread_escuta = threading.Thread(target=self.listener.run, daemon=True)
        self.thread_escuta.start()

    def _on_press(self, key):
        try:
            if key == keyboard.Key.enter or key == keyboard.Key.return_:
                self.root.after(0, self._finalizar_leitura)
                return
            elif key == keyboard.KeyCode.from_char("\r") or key == keyboard.KeyCode.from_char("\n"):
                self.root.after(0, self._finalizar_leitura)
                return
        except Exception:
            pass

        if isinstance(key, keyboard.KeyCode):
            char = key.char
            if char is not None:
                self.buffer += char

    def _finalizar_leitura(self):
        if not self.gravando:
            self.buffer = ""
            return
        if not self.buffer:
            return
        cartao = self.buffer.strip()
        self.buffer = ""
        if cartao:
            self._registrar_cartao(cartao)

    def _registrar_cartao(self, cartao):
        cartao = self._processar_cartao(cartao)
        if not cartao:
            return
        self.sequencia += 1
        datahora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        self.linhas.append({"seq": self.sequencia, "cartao": cartao, "datahora": datahora})
        self.tree.insert("", "end", iid=str(self.sequencia - 1),
                         values=(self.sequencia, cartao, datahora))
        self.tree.see(self.tree.get_children()[-1])
        self.lbl_contador.configure(text=f"{len(self.linhas)} leitura(s)")
        self._log(f"Cartão lido: {cartao}")
        try:
            self._gravar_arquivo(path=self.caminho_planilha, relatorio=False)
        except Exception as e:
            self._log(f"Erro ao gravar: {e}")

    def _gravar_arquivo(self, path, relatorio):
        if str(path).lower().endswith(".txt"):
            with open(path, "a", encoding="utf-8") as f:
                if not self.linhas:
                    return
                ultimo = self.linhas[-1]
                f.write(f"{ultimo['seq']}\t{ultimo['cartao']}\n")
        else:
            df = pd.DataFrame(self.linhas, columns=["seq", "cartao", "datahora"])
            if not relatorio:
                df.to_excel(path, index=False)

    def _limpar_tree(self):
        self.tree.delete(*self.tree.get_children())

    def _parar_fechar(self):
        if not self.gravando:
            return
        self._parar_captura()
        try:
            if self.caminho_planilha and self.linhas \
                    and not str(self.caminho_planilha).lower().endswith(".txt"):
                self._gravar_arquivo(path=self.caminho_planilha, relatorio=False)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar o arquivo.\n{e}")
        self.gravando = False
        self._set_status("Parado. Arquivo salvo e fechado.", _cor_mutado())
        self.btn_iniciar.state(["!disabled"])
        self.btn_abrir.state(["!disabled"])
        self.btn_parar.state(["disabled"])
        self.cb_fonte.state(["!disabled"])
        self.lbl_contador.configure(text=f"{len(self.linhas)} leitura(s) salva(s)")
        self._log(f"Gravação encerrada. {len(self.linhas)} leitura(s) registrada(s).")

    def _parar_captura(self):
        self._parar_listener()
        self._parar_serial()

    def _parar_listener(self):
        if self.listener is not None:
            try:
                self.listener.stop()
            except Exception:
                pass
            self.listener = None

    def _parar_serial(self):
        self.thread_serial = None
        if self.ser is not None:
            try:
                if self.ser.is_open:
                    self.ser.close()
            except Exception:
                pass
            self.ser = None
        self.porta_serial = ""

    def _ao_fechar(self):
        if self.gravando:
            self._parar_fechar()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = LeitorIdentifid(root)
    root.protocol("WM_DELETE_WINDOW", app._ao_fechar)
    root.mainloop()
