import subprocess
import importlib.util
import sys
import argparse

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
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import time
import json
import os
import re
import ctypes
import base64
import datetime
from pathlib import Path

# Torna o processo DPI-aware (PER_MONITOR) para coordenadas e capturas de tela corretas
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

# Eleva para administrador se necessário
def _elevar_admin(args_linha=""):
    if ctypes.windll.shell32.IsUserAnAdmin():
        return
    script = os.path.abspath(sys.argv[0])
    cmd = f'"{script}"'
    if args_linha:
        cmd += f" {args_linha}"
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, cmd, None, 1
    )
    sys.exit()

# Captura o caminho do config passado via --config (usado pelo seletor)
_CONFIG_ARG = ""
if "--config" in sys.argv:
    try:
        _CONFIG_ARG = sys.argv[sys.argv.index("--config") + 1].strip(" \"")
    except (IndexError, ValueError):
        _CONFIG_ARG = ""

# Captura o tema passado via --tema (usado pelo seletor)
_TEMA_ARG = ""
if "--tema" in sys.argv:
    try:
        _TEMA_ARG = sys.argv[sys.argv.index("--tema") + 1].strip(" \"")
    except (IndexError, ValueError):
        _TEMA_ARG = ""

_args_admin = _TEMA_ARG and f' --tema "{_TEMA_ARG}"'
_elevar_admin(f'--config "{_CONFIG_ARG}"{_args_admin}' if _CONFIG_ARG else _args_admin)

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    from pynput import mouse, keyboard
    from pynput.mouse import Button as MouseButton
    from pynput.keyboard import Key
    HAVE_PYNPUT = True
except ImportError:
    HAVE_PYNPUT = False

try:
    import mss
    import mss.tools
    HAVE_MSS = True
except ImportError:
    HAVE_MSS = False

# Mapeia cada caractere especial ASCII para (código VK, precisa_shift)
_CARACTERES_ESPECIAIS = {
    "!": (0x31, True),    # Shift+1
    '"': (0xDE, True),    # Shift+'
    "#": (0x33, True),    # Shift+3
    "$": (0x34, True),    # Shift+4
    "%": (0x35, True),    # Shift+5
    "&": (0x37, True),    # Shift+7
    "'": (0xDE, False),
    "(": (0x39, True),    # Shift+9
    ")": (0x30, True),    # Shift+0
    "*": (0x38, True),    # Shift+8
    "+": (0xBB, True),    # Shift+=
    ",": (0xBC, False),
    "-": (0xBD, False),
    ".": (0xBE, False),
    "/": (0xBF, False),
    ":": (0xBA, True),    # Shift+;
    ";": (0xBA, False),
    "<": (0xBC, True),    # Shift+,
    "=": (0xBB, False),
    ">": (0xBE, True),    # Shift+.
    "?": (0xBF, True),    # Shift+/
    "@": (0x32, True),    # Shift+2
    "[": (0xDB, False),
    "\\": (0xDC, False),
    "]": (0xDD, False),
    "^": (0x36, True),    # Shift+6
    "_": (0xBD, True),    # Shift+-
    "`": (0xC0, False),
    "{": (0xDB, True),    # Shift+[
    "|": (0xDC, True),    # Shift+\
    "}": (0xDD, True),    # Shift+]
    "~": (0xC0, True),    # Shift+`
}


ARQUIVO_CONFIG = (Path(_CONFIG_ARG) if _CONFIG_ARG else (Path(__file__).parent / "config_passos.json")).resolve()

sys.path.insert(0, str(Path(__file__).parent))
import tema

TEMA = _TEMA_ARG if _TEMA_ARG else tema.carregar_tema()


def _cor_mutado():
    return "#9a9a9a" if TEMA != "dark" else "#8d8d8d"


def _cor_destaque():
    return "#005a9e" if TEMA != "dark" else "#3794ff"


# ═══════════════════════════════════════════════════════════
#  Helpers: datas e referências de tela (Validar Tela)
# ═══════════════════════════════════════════════════════════

def _parse_data(texto, formato="%d/%m/%Y"):
    return datetime.datetime.strptime(str(texto).strip(), formato).date()


def _comparar_rgb(rgb_ref, rgb_atual, tolerancia=30, amostra=3):
    # Compara os dois buffers byte a byte (amostra=3 → todos os canais RGB).
    if not rgb_ref or not rgb_atual or len(rgb_ref) != len(rgb_atual):
        return 0.0
    bate = 0
    total = 0
    for i in range(0, len(rgb_ref), amostra):
        total += 1
        if abs(rgb_ref[i] - rgb_atual[i]) <= tolerancia:
            bate += 1
    return (bate / total) if total else 0.0


def _salvar_referencia_arquivo(regiao, caminho_json):
    # Captura a região da tela e salva como JSON bruto (comparável sem dependências)
    # e também um PNG apenas para visualização.
    if not HAVE_MSS:
        raise RuntimeError("biblioteca 'mss' não instalada")
    with mss.mss() as sct:
        rect = {"left": int(regiao["x"]), "top": int(regiao["y"]),
                "width": int(regiao["largura"]), "height": int(regiao["altura"])}
        if rect["width"] <= 0 or rect["height"] <= 0:
            raise RuntimeError("região inválida (largura/altura devem ser > 0)")
        img = sct.grab(rect)
        dados = {
            "largura": img.width,
            "altura": img.height,
            "regiao": dict(regiao),
            "data_captura": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "rgb": base64.b64encode(img.rgb).decode("ascii"),
        }
        os.makedirs(os.path.dirname(caminho_json), exist_ok=True)
        with open(caminho_json, "w", encoding="utf-8") as f:
            json.dump(dados, f)
        caminho_png = os.path.splitext(caminho_json)[0] + ".png"
        try:
            mss.tools.to_png(img.rgb, img.size, output=caminho_png)
        except Exception:
            caminho_png = ""
    return caminho_json, caminho_png


def _carregar_referencia(caminho):
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            dados = json.load(f)
        dados["rgb"] = base64.b64decode(dados["rgb"])
        if not dados.get("rgb"):
            return None
        return dados
    except Exception:
        return None


class AppAutomacao:
    def __init__(self, root):
        self.root = root
        self.root.title("Automação de Cliques com Planilha")
        self.root.geometry("900x650")
        self.root.minsize(800, 500)

        # --- estado interno ---
        self.planilha_df = None
        self.colunas = []
        self.caminho_planilha = ""
        self.passos = []           # lista de dicts: {tipo, x, y, coluna, espera, rotulo}
        self.gravando = False
        self.executando = False
        self.mouse_listener = None
        self.thread_play = None
        self.thread_gravacao = None
        self.play_parar = threading.Event()
        self.play_continuar = threading.Event()
        self.pausado = False
        self.lote_ativo = False
        self.lote_tamanho = 10
        self.lote_processados = 0
        self._linhas_restantes_exec = 0
        self._tempo_passos = 0.0
        self._inicio_linha = time.time()
        self._datas_exec = []        # lista de datetime.date (modo Datas)
        self._pos_unidade = 0        # posição atual no loop (linha ou dia)
        self._pular_unidade = False  # True -> interrompe a sequência do dia/linha atual (próxima data)
        self.coluna_lido = tk.StringVar(value="Lido")

        # --- modo de execução e período (modo Datas) ---
        self.var_modo = tk.StringVar(value="planilha")
        hoje = time.strftime("%d/%m/%Y")
        self.var_data_inicial = tk.StringVar(value=hoje)
        self.var_data_final = tk.StringVar(value=hoje)
        self.var_formato_data = tk.StringVar(value="%d/%m/%Y")

        self._construir_ui()
        self._carregar_config()
        self._iniciar_tecla_parada()
        self.root.after(250, self._tick_tempo)

    # ────────────────────────────────────────────────────────
    #  UI
    # ────────────────────────────────────────────────────────

    def _construir_ui(self):
        # --- Painel de modo de execução ---
        frame_modo = ttk.Frame(self.root)
        frame_modo.pack(fill=tk.X, padx=8, pady=(8, 0))
        ttk.Label(frame_modo, text="Modo:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        ttk.Radiobutton(frame_modo, text="Planilha", variable=self.var_modo,
                        value="planilha", command=self._trocar_modo).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Radiobutton(frame_modo, text="Datas (dia a dia)", variable=self.var_modo,
                        value="datas", command=self._trocar_modo).pack(side=tk.LEFT, padx=(8, 0))
        self.frame_modo = frame_modo

        # --- Painel superior: planilha ---
        frame_planilha = ttk.LabelFrame(self.root, text="Planilha", padding=6)
        self.frame_planilha = frame_planilha

        self.btn_carregar = ttk.Button(frame_planilha, text="Carregar Planilha", command=self._carregar_planilha)
        self.btn_carregar.pack(side=tk.LEFT, padx=(0, 6))

        self.lbl_planilha = ttk.Label(frame_planilha, text="Nenhuma planilha carregada", foreground=_cor_mutado())
        self.lbl_planilha.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Label(frame_planilha, text="Coluna 'Lido':").pack(side=tk.LEFT, padx=(10, 2))
        self.entry_col_lido = ttk.Entry(frame_planilha, textvariable=self.coluna_lido, width=10)
        self.entry_col_lido.pack(side=tk.LEFT)

        ttk.Label(frame_planilha, text="Filtrar por linha:").pack(side=tk.LEFT, padx=(10, 2))
        self.spin_linha_inicio = ttk.Spinbox(frame_planilha, from_=1, to=999999, width=6)
        self.spin_linha_inicio.set(1)
        self.spin_linha_inicio.pack(side=tk.LEFT)
        ttk.Label(frame_planilha, text="até").pack(side=tk.LEFT, padx=2)
        self.spin_linha_fim = ttk.Spinbox(frame_planilha, from_=1, to=999999, width=6)
        self.spin_linha_fim.set("")
        self.spin_linha_fim.pack(side=tk.LEFT)

        # --- Painel superior: período (modo Datas) ---
        frame_datas = ttk.LabelFrame(self.root, text="Período (modo Datas)", padding=6)
        self.frame_datas = frame_datas

        ttk.Label(frame_datas, text="Data inicial:").pack(side=tk.LEFT, padx=(0, 4))
        self.entry_data_inicial = ttk.Entry(frame_datas, textvariable=self.var_data_inicial, width=12)
        self.entry_data_inicial.pack(side=tk.LEFT)

        ttk.Label(frame_datas, text="Data final:").pack(side=tk.LEFT, padx=(12, 4))
        self.entry_data_final = ttk.Entry(frame_datas, textvariable=self.var_data_final, width=12)
        self.entry_data_final.pack(side=tk.LEFT)

        ttk.Label(frame_datas, text="Formato:").pack(side=tk.LEFT, padx=(12, 4))
        self.entry_formato_data = ttk.Entry(frame_datas, textvariable=self.var_formato_data, width=14)
        self.entry_formato_data.pack(side=tk.LEFT)

        ttk.Label(frame_datas, text="(ex: %d/%m/%Y)",
                  foreground=_cor_mutado()).pack(side=tk.LEFT, padx=(4, 0))

        ttk.Label(
            frame_datas,
            text="A automação preenche os dois campos de data com o mesmo dia e avança 1 dia por rodada até a data final.",
            foreground=_cor_mutado()
        ).pack(side=tk.LEFT, padx=(16, 0))

        frame_planilha.pack(fill=tk.X, padx=8, pady=(8, 2))
        self._trocar_modo()

        # --- Painel de passos ---
        frame_passos = ttk.LabelFrame(self.root, text="Passos Gravados", padding=6)
        frame_passos.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        colunas_passos = ("#", "Tipo", "Posição (X, Y)", "Coluna", "Espera (s)", "Tam.Mín", "Rótulo")
        self.tree_passos = ttk.Treeview(frame_passos, columns=colunas_passos, show="headings",
                                        selectmode="extended", height=8)
        largs = [40, 80, 130, 100, 80, 70, 150]
        for c, l in zip(colunas_passos, largs):
            self.tree_passos.heading(c, text=c)
            self.tree_passos.column(c, width=l, anchor=tk.CENTER if c != "Rótulo" else tk.W)

        scroll_y = ttk.Scrollbar(frame_passos, orient=tk.VERTICAL, command=self.tree_passos.yview)
        self.tree_passos.configure(yscrollcommand=scroll_y.set)
        self.tree_passos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # --- Botões dos passos ---
        frame_botoes_passos = ttk.Frame(self.root)
        frame_botoes_passos.pack(fill=tk.X, padx=8, pady=2)

        self.btn_gravar = ttk.Button(frame_botoes_passos, text="Iniciar Gravação", command=self._toggle_gravacao)
        self.btn_gravar.pack(side=tk.LEFT, padx=2)

        self.btn_adicionar = ttk.Button(frame_botoes_passos, text="Adicionar Passo", command=self._adicionar_passo)
        self.btn_adicionar.pack(side=tk.LEFT, padx=2)

        self.btn_editar = ttk.Button(frame_botoes_passos, text="Editar Passo", command=self._editar_passo)
        self.btn_editar.pack(side=tk.LEFT, padx=2)

        self.btn_remover = ttk.Button(frame_botoes_passos, text="Remover Passo(s)", command=self._remover_passos)
        self.btn_remover.pack(side=tk.LEFT, padx=2)

        self.btn_limpar = ttk.Button(frame_botoes_passos, text="Limpar Tudo", command=self._limpar_passos)
        self.btn_limpar.pack(side=tk.LEFT, padx=2)

        self.btn_subir = ttk.Button(frame_botoes_passos, text="▲", width=3, command=self._subir_passo)
        self.btn_subir.pack(side=tk.LEFT, padx=2)
        self.btn_descer = ttk.Button(frame_botoes_passos, text="▼", width=3, command=self._descer_passo)
        self.btn_descer.pack(side=tk.LEFT, padx=2)

        ttk.Separator(frame_botoes_passos, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8)

        self.btn_salvar = ttk.Button(frame_botoes_passos, text="Salvar Config", command=self._salvar_config)
        self.btn_salvar.pack(side=tk.LEFT, padx=2)
        self.btn_carregar_config = ttk.Button(frame_botoes_passos, text="Carregar Config",
                                               command=self._carregar_config_manual)
        self.btn_carregar_config.pack(side=tk.LEFT, padx=2)

        # --- Execução ---
        frame_exec = ttk.LabelFrame(self.root, text="Execução", padding=6)
        frame_exec.pack(fill=tk.X, padx=8, pady=(2, 8))

        self.btn_executar = ttk.Button(frame_exec, text="▶ Executar", command=self._executar)
        self.btn_executar.pack(side=tk.LEFT, padx=2)

        self.btn_parar = ttk.Button(frame_exec, text="■ Parar", command=self._parar_execucao, state=tk.DISABLED)
        self.btn_parar.pack(side=tk.LEFT, padx=2)

        ttk.Label(frame_exec, text="Delay entre passos (s):").pack(side=tk.LEFT, padx=(16, 2))
        self.spin_delay = ttk.Spinbox(frame_exec, from_=0.0, to=10.0, increment=0.1, width=6)
        self.spin_delay.set(0.3)
        self.spin_delay.pack(side=tk.LEFT)

        ttk.Label(frame_exec, text="Repetições:").pack(side=tk.LEFT, padx=(10, 2))
        self.spin_repetir = ttk.Spinbox(frame_exec, from_=1, to=999, width=5)
        self.spin_repetir.set(1)
        self.spin_repetir.pack(side=tk.LEFT)

        ttk.Label(frame_exec, text="Tempo restante:").pack(side=tk.LEFT, padx=(10, 2))
        self.lbl_tempo_restante = ttk.Label(frame_exec, text="—", foreground=_cor_destaque(),
                                            font=("Segoe UI", 9, "bold"))
        self.lbl_tempo_restante.pack(side=tk.LEFT)

        ttk.Separator(frame_exec, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8)

        frame_intervalo = ttk.Frame(frame_exec)
        frame_intervalo.pack(side=tk.LEFT, padx=8)
        self.var_lote = tk.BooleanVar(value=False)
        self.chk_lote = ttk.Checkbutton(frame_intervalo, text="Intervalo", variable=self.var_lote,
                                         command=self._toggle_lote)
        self.chk_lote.pack(side=tk.TOP, anchor=tk.W)

        frame_int_linha = ttk.Frame(frame_intervalo)
        frame_int_linha.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(frame_int_linha, text="Linhas:").pack(side=tk.LEFT)
        self.spin_lote = ttk.Spinbox(frame_int_linha, from_=1, to=9999, width=5, state=tk.DISABLED)
        self.spin_lote.set(10)
        self.spin_lote.pack(side=tk.LEFT, padx=(2, 0))

        self.lbl_restantes = ttk.Label(frame_intervalo, text="", foreground=_cor_destaque())
        self.lbl_restantes.pack(side=tk.TOP, anchor=tk.W)

        # --- Barra de status ---
        frame_status = ttk.Frame(self.root)
        frame_status.pack(fill=tk.X, padx=8, pady=(0, 2))
        self.lbl_status = ttk.Label(frame_status, text="⏸ Aguardando", foreground=_cor_mutado(),
                                    font=("Segoe UI", 9, "bold"))
        self.lbl_status.pack(side=tk.LEFT)
        self.lbl_progresso = ttk.Label(frame_status, text="", foreground=_cor_mutado(), font=("Segoe UI", 9))
        self.lbl_progresso.pack(side=tk.RIGHT)

        # --- Log ---
        frame_log = ttk.LabelFrame(self.root, text="Log", padding=4)
        frame_log.pack(fill=tk.BOTH, padx=8, pady=(0, 8))

        self.log = scrolledtext.ScrolledText(frame_log, height=6, state=tk.DISABLED, font=("Consolas", 9))
        self.log.pack(fill=tk.BOTH, expand=True)

        for w in (self.spin_delay, self.spin_linha_inicio, self.spin_linha_fim,
                  self.entry_data_inicial, self.entry_data_final, self.entry_formato_data):
            w.bind("<KeyRelease>", lambda e: self._atualizar_tempo_restante())
            w.bind("<<Increment>>", lambda e: self._atualizar_tempo_restante())
            w.bind("<<Decrement>>", lambda e: self._atualizar_tempo_restante())

        tema.aplicar_tema(self.root, TEMA)

    def _toggle_lote(self):
        estado = tk.NORMAL if self.var_lote.get() else tk.DISABLED
        self.spin_lote.config(state=estado)

    def _trocar_modo(self, event=None):
        modo = self.var_modo.get()
        if modo == "datas":
            self.frame_planilha.pack_forget()
            self.frame_datas.pack(fill=tk.X, padx=8, pady=(8, 2), after=self.frame_modo)
            self.root.title("Automação de Datas")
        else:
            self.frame_datas.pack_forget()
            self.frame_planilha.pack(fill=tk.X, padx=8, pady=(8, 2), after=self.frame_modo)
            self.root.title("Automação de Cliques com Planilha")
        if hasattr(self, "lbl_tempo_restante"):
            self._atualizar_tempo_restante()

    def _total_dias(self):
        try:
            formato = self.var_formato_data.get().strip() or "%d/%m/%Y"
            d1 = _parse_data(self.var_data_inicial.get(), formato)
            d2 = _parse_data(self.var_data_final.get(), formato)
        except (ValueError, TypeError):
            return 0
        if d2 < d1:
            return 0
        return (d2 - d1).days + 1

    # ────────────────────────────────────────────────────────
    #  Tempo restante estimado
    # ────────────────────────────────────────────────────────

    def _formatar_tempo(self, seg):
        seg = int(round(max(seg, 0)))
        if seg < 60:
            return f"{seg}s"
        horas, resto = divmod(seg, 3600)
        minutos, secs = divmod(resto, 60)
        if horas:
            return f"{horas}h {minutos:02d}min {secs:02d}s"
        return f"{minutos}min {secs:02d}s"

    def _tempo_estimado(self):
        try:
            delay = float(self.spin_delay.get())
        except (TypeError, ValueError):
            delay = 0.3
        if delay <= 0:
            delay = 0.0

        def _soma_passo(p):
            try:
                espera = float(p.get("espera", 0))
            except (TypeError, ValueError):
                espera = 0.0
            total = espera if espera > 0 else delay
            if p.get("tipo") == "Validar Tela":
                try:
                    total += float(p.get("timeout", 0)) or 0.0
                except (TypeError, ValueError):
                    pass
                for a in p.get("acoes", []):
                    total += _soma_passo(a)
                for a in p.get("acoes_nao_encontrado", []):
                    total += _soma_passo(a)
            return total

        return sum(_soma_passo(p) for p in self.passos)

    def _linhas_restantes(self):
        if self.planilha_df is None:
            return 0
        try:
            inicio = int(self.spin_linha_inicio.get()) - 1
        except ValueError:
            inicio = 0
        fim_str = self.spin_linha_fim.get()
        try:
            fim = int(fim_str) if fim_str.strip() else len(self.planilha_df)
        except ValueError:
            fim = len(self.planilha_df)
        inicio = max(inicio, 0)
        fim = min(fim, len(self.planilha_df))
        ultima = self._ultima_linha_processada()
        if ultima > inicio:
            inicio = ultima
        return max(fim - inicio, 0)

    def _tempo_passos_ativo(self):
        if self.executando and self._tempo_passos > 0:
            return self._tempo_passos
        return self._tempo_estimado()

    def _atualizar_tempo_restante(self):
        def _calc():
            tempo_passos = self._tempo_passos_ativo()
            if self.executando:
                rows = self._linhas_restantes_exec
                if rows > 0:
                    elapsed = time.time() - self._inicio_linha
                    total = (rows - 1) * tempo_passos + max(tempo_passos - elapsed, 0)
                else:
                    total = 0
            elif self.var_modo.get() == "datas":
                total = self._total_dias() * tempo_passos
            elif self.planilha_df is None:
                total = tempo_passos  # sem planilha: tempo do total de passos configurados
            else:
                total = self._linhas_restantes() * tempo_passos
            self.lbl_tempo_restante.configure(text=self._formatar_tempo(total))
        try:
            self.root.after(0, _calc)
        except Exception:
            pass

    def _tick_tempo(self):
        if self.executando:
            self._atualizar_tempo_restante()
        try:
            self.root.after(250, self._tick_tempo)
        except Exception:
            pass

    # ────────────────────────────────────────────────────────
    #  Log
    # ────────────────────────────────────────────────────────

    def _log(self, msg):
        self.root.after(0, self._log_inner, msg)

    def _log_inner(self, msg):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)

    # ────────────────────────────────────────────────────────
    #  Tecla de atalho F12 para parar execução
    # ────────────────────────────────────────────────────────

    def _iniciar_tecla_parada(self):
        if not HAVE_PYNPUT:
            return
        def ao_pressionar(tecla):
            try:
                if tecla == Key.esc:
                    if self.executando and not self.pausado:
                        self._log("ESC pressionado — parando execução...")
                        self._parar_execucao()
                elif tecla == Key.f1:
                    if self.pausado:
                        self._log("F1 pressionado — continuando...")
                        self.play_continuar.set()
            except Exception:
                pass
            return True
        listener = keyboard.Listener(on_press=ao_pressionar, daemon=True)
        listener.start()

    # ────────────────────────────────────────────────────────
    #  Planilha
    # ────────────────────────────────────────────────────────

    def _carregar_planilha(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar Planilha",
            filetypes=[("Planilhas", "*.xlsx *.xls *.csv"), ("Todos", "*.*")]
        )
        if not caminho:
            return
        if pd is None:
            messagebox.showerror("Erro", "Biblioteca pandas não está instalada.\nExecute: pip install pandas openpyxl")
            return
        try:
            if caminho.endswith(".csv"):
                self.planilha_df = pd.read_csv(caminho, encoding="utf-8-sig")
                self.formatos = {}
            else:
                self.planilha_df = pd.read_excel(caminho, engine="openpyxl")
                self._ler_formatos_excel(caminho)
        except Exception as e:
            messagebox.showerror("Erro ao ler planilha", str(e))
            return

        self.caminho_planilha = caminho
        if not caminho.endswith(".csv"):
            self._garantir_coluna_lido()
        self.colunas = list(self.planilha_df.columns)
        nome_arquivo = os.path.basename(caminho)
        self.lbl_planilha.configure(text=f"{nome_arquivo} ({len(self.planilha_df)} linhas)")
        self.spin_linha_fim.set(len(self.planilha_df))
        qtd_formatadas = sum(1 for v in self.formatos.values() if v)
        self._log(f"Planilha carregada: {nome_arquivo} — {len(self.planilha_df)} linhas, {len(self.colunas)} colunas{qtd_formatadas and f', {qtd_formatadas} célula(s) com formatação' or ''}")
        messagebox.showinfo("Planilha", f"Carregada com sucesso!\nLinhas: {len(self.planilha_df)}\nColunas: {', '.join(self.colunas)}")
        self._atualizar_tempo_restante()

    def _ler_formatos_excel(self, caminho):
        self.formatos = {}
        try:
            import openpyxl
            wb = openpyxl.load_workbook(caminho, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
                for cell in row:
                    if cell.value is not None and cell.number_format and cell.number_format != "General":
                        self.formatos[(cell.row, cell.column)] = cell.number_format
            wb.close()
        except Exception:
            pass

    def _valor_formatado(self, coluna, idx_linha):
        raw = self.planilha_df.iloc[idx_linha][coluna]
        if pd.isna(raw):
            return ""
        col_idx = self.colunas.index(coluna) + 1
        excel_row = idx_linha + 2
        nf = self.formatos.get((excel_row, col_idx))
        if nf and isinstance(raw, (int, float)):
            return self._aplicar_formato(raw, nf)
        return str(raw)

    def _aplicar_formato(self, valor, formato):
        if isinstance(valor, float) and valor == int(valor):
            valor = int(valor)
        if isinstance(valor, int):
            if re.match(r'^0+$', formato):
                return str(valor).zfill(len(formato))
            if re.match(r'^[#0,]+(\.[0#]+)?$', formato):
                partes = formato.split(".")
                int_part = partes[0]
                dec_part = partes[1] if len(partes) > 1 else ""
                min_digits = int_part.count("0")
                if dec_part:
                    dec_digits = len(dec_part)
                    return f"{valor:0{min_digits}.{dec_digits}f}"
                return str(valor) if min_digits == 0 else f"{valor:0{min_digits}d}"
            m = re.search(r'"([^"]*)"', formato)
            if m:
                texto = m.group(1)
                if "@" in formato:
                    before, after = formato.split("@")
                    prefix = re.search(r'"([^"]*)"', before)
                    suffix = re.search(r'"([^"]*)"', after)
                    return f"{prefix.group(1) if prefix else ''}{valor}{suffix.group(1) if suffix else ''}"
        return str(valor)

    def _garantir_coluna_lido(self):
        nome = self.coluna_lido.get().strip() or "Lido"
        if nome not in self.planilha_df.columns:
            self.planilha_df[nome] = ""
            try:
                import openpyxl
                wb = openpyxl.load_workbook(self.caminho_planilha)
                ws = wb.active
                max_col = ws.max_column
                ws.cell(row=1, column=max_col + 1, value=nome)
                wb.save(self.caminho_planilha)
                wb.close()
                self._log(f"Coluna '{nome}' criada na planilha.")
            except Exception as e:
                self._log(f"Aviso: não foi possível salvar coluna '{nome}' no arquivo: {e}")
        self.colunas = list(self.planilha_df.columns)

    def _ultima_linha_processada(self):
        nome = self.coluna_lido.get().strip() or "Lido"
        if nome not in self.planilha_df.columns:
            return -1
        for idx in range(len(self.planilha_df)):
            val = self.planilha_df.iloc[idx][nome]
            if pd.isna(val) or str(val).strip() == "":
                return idx
        return len(self.planilha_df)

    def _marcar_linha_processada(self, idx_linha):
        nome = self.coluna_lido.get().strip() or "Lido"
        if nome not in self.planilha_df.columns:
            return
        self.planilha_df.at[idx_linha, nome] = "Sim"

    def _salvar_planilha_marcacoes(self):
        if not self.caminho_planilha or self.caminho_planilha.endswith(".csv"):
            return
        for tentativa in range(3):
            try:
                import openpyxl
                wb = openpyxl.load_workbook(self.caminho_planilha)
                ws = wb.active
                nome = self.coluna_lido.get().strip() or "Lido"
                col_idx = self.colunas.index(nome) + 1
                for idx in range(len(self.planilha_df)):
                    val = self.planilha_df.iloc[idx][nome]
                    ws.cell(row=idx + 2, column=col_idx, value=val if pd.notna(val) else "")
                wb.save(self.caminho_planilha)
                wb.close()
                return
            except Exception as e:
                self._log(f"Tentativa {tentativa+1}/3: erro ao salvar marcações: {e}")
                time.sleep(1)
        self._log("ERRO: não foi possível salvar marcações na planilha.")

    # ────────────────────────────────────────────────────────
    #  Gravação de cliques
    # ────────────────────────────────────────────────────────

    def _toggle_gravacao(self):
        if self.gravando:
            self._parar_gravacao()
        else:
            self._iniciar_gravacao()

    def _iniciar_gravacao(self):
        if not HAVE_PYNPUT:
            messagebox.showerror("Erro", "pynput não instalado.\nExecute: pip install pynput")
            return
        self.gravando = True
        self.btn_gravar.configure(text="Parar Gravação")
        self._log("Gravação iniciada — clique para registrar posições. Ctrl+P para adicionar Print de tela.")
        self.thread_gravacao = threading.Thread(target=self._escutar_cliques, daemon=True)
        self.thread_gravacao.start()
        self._tecla_gravacao_listener = keyboard.Listener(on_press=self._ao_pressionar_gravacao, daemon=True)
        self._tecla_gravacao_listener.start()

    def _parar_gravacao(self):
        self.gravando = False
        self.btn_gravar.configure(text="Iniciar Gravação")
        self._log(f"Gravação encerrada. Total de passos: {len(self.passos)}")
        if self.mouse_listener:
            self.mouse_listener.stop()
            self.mouse_listener = None
        if hasattr(self, '_tecla_gravacao_listener') and self._tecla_gravacao_listener:
            self._tecla_gravacao_listener.stop()
            self._tecla_gravacao_listener = None
        self._atualizar_tree()

    def _ao_pressionar_gravacao(self, tecla):
        if not self.gravando:
            return False
        try:
            if tecla == Key.ctrl_l or tecla == Key.ctrl_r:
                return True
            if hasattr(tecla, 'char') and tecla.char == 'p':
                ctrl_pressionado = (ctypes.windll.user32.GetAsyncKeyState(0x11) & 0x8000)
                if ctrl_pressionado:
                    if not HAVE_MSS:
                        self._log("ERRO: biblioteca 'mss' não instalada para Print de tela.")
                        return True
                    self.passos.append({
                        "tipo": "Print",
                        "x": 0,
                        "y": 0,
                        "coluna": "",
                        "texto_fixo": "",
                        "espera": 0.5,
                        "rotulo": f"Print {len(self.passos) + 1}",
                        "tamanho_min": 0,
                        "selecionar_tudo": False,
                        "apagar": False,
                        "monitor": 1,
                        "pasta_print": "",
                        "coluna_nome": "",
                    })
                    self._log(f"Print registrado — total: {len(self.passos)}")
                    self._atualizar_tree()
        except Exception:
            pass
        return True

    def _escutar_cliques(self):
        # Ignora cliques dentro da área da janela do app (não vira passo)
        def _dentro_janela_app(cx, cy):
            try:
                self.root.update_idletasks()
                left = self.root.winfo_rootx()
                top = self.root.winfo_rooty()
                right = left + self.root.winfo_width()
                bottom = top + self.root.winfo_height()
                return left <= cx <= right and top <= cy <= bottom
            except Exception:
                return False

        def on_click(x, y, button, pressed):
            if pressed and button == MouseButton.left:
                pos = ctypes.wintypes.POINT()
                ctypes.windll.user32.GetCursorPos(ctypes.byref(pos))
                cx, cy = pos.x, pos.y
                if cx == 0 and cy == 0:
                    cx, cy = round(x), round(y)
                if _dentro_janela_app(cx, cy):
                    return self.gravando
                shift = ctypes.windll.user32.GetAsyncKeyState(0x10) & 0x8000
                ctrl = ctypes.windll.user32.GetAsyncKeyState(0x11) & 0x8000
                if ctrl and shift:
                    tipo = "Data"
                elif ctrl:
                    tipo = "Planilha"
                elif shift:
                    tipo = "Texto"
                else:
                    tipo = "Clique"
                extra = ""
                if tipo == "Planilha":
                    extra = " — defina a coluna no editar"
                elif tipo == "Texto":
                    extra = " — defina o texto no editar"
                elif tipo == "Data":
                    extra = " — a data do dia é digitada automaticamente (Ctrl+Shift)"
                self.passos.append({
                    "tipo": tipo,
                    "x": cx,
                    "y": cy,
                    "coluna": "",
                    "texto_fixo": "",
                    "espera": 0.5,
                    "rotulo": f"{tipo} {len(self.passos) + 1}",
                    "tamanho_min": 0,
                    "selecionar_tudo": False,   # Ctrl+A
                    "apagar": False,            # Delete
                })
                self._log(f"{tipo} registrado: ({cx}, {cy}){extra} — total: {len(self.passos)}")
                self._atualizar_tree()
            return self.gravando

        with mouse.Listener(on_click=on_click) as listener:
            self.mouse_listener = listener
            listener.join()

    # ────────────────────────────────────────────────────────
    #  Treeview
    # ────────────────────────────────────────────────────────

    def _atualizar_tree(self):
        for item in self.tree_passos.get_children():
            self.tree_passos.delete(item)
        for i, p in enumerate(self.passos, 1):
            tam = p.get("tamanho_min", 0)
            if p["tipo"] == "Print":
                monitor = p.get("monitor", 1)
                pasta = p.get("pasta_print", "")
                col_nome = p.get("coluna_nome", "")
                pos_str = f"Monitor {monitor}"
                col_str = f"Pasta: {os.path.basename(pasta)}" if pasta else ""
                tam_str = f"Col: {col_nome}" if col_nome else ""
            elif p["tipo"] == "Data":
                pos_str = f"({p['x']}, {p['y']})"
                col_str = p.get("formato_data", "%d/%m/%Y")
                tam_str = ""
            elif p["tipo"] == "Pular Data":
                pos_str = "—"
                col_str = "→ próximo dia/linha"
                tam_str = ""
            elif p["tipo"] == "Registrar Ignorado":
                arq = p.get("arquivo_ignorados", "")
                pos_str = "—"
                col_str = os.path.basename(arq) if arq else "(sem arquivo)"
                tam_str = ""
            elif p["tipo"] == "Validar Tela":
                r = p.get("regiao") or {}
                pos_str = f"({r.get('x', 0)},{r.get('y', 0)}) {r.get('largura', 0)}x{r.get('altura', 0)}"
                n_sim = len(p.get('acoes', []))
                n_nao = len(p.get('acoes_nao_encontrado', []))
                col_str = f"{n_sim} ação(ões) / {n_nao} se não achar"
                tam_str = ""
            else:
                pos_str = f"({p['x']}, {p['y']})"
                col_str = p.get("coluna", "")
                tam_str = tam if tam > 0 else ""
            self.tree_passos.insert("", tk.END, values=(
                i,
                p["tipo"],
                pos_str,
                col_str,
                p.get("espera", 0),
                tam_str,
                p.get("rotulo", "")
            ))
        self._atualizar_tempo_restante()

    # ────────────────────────────────────────────────────────
    #  Editar / Remover / Reordenar passos
    # ────────────────────────────────────────────────────────

    def _adicionar_passo(self):
        novo = {
            "tipo": "Clique",
            "x": 0,
            "y": 0,
            "coluna": "",
            "texto_fixo": "",
            "espera": 0.5,
            "rotulo": f"Clique {len(self.passos) + 1}",
            "tamanho_min": 0,
            "selecionar_tudo": False,
            "apagar": False,
            "formato_data": "%d/%m/%Y",
            "regiao": {"x": 0, "y": 0, "largura": 100, "altura": 100},
            "referencia": "",
            "limiar": 0.85,
            "tolerancia": 30,
            "timeout": 3.0,
            "acoes": [],
            "acoes_nao_encontrado": [],
            "arquivo_ignorados": "",
        }
        dialogo = JanelaEditarPasso(self.root, novo, self.colunas)
        self.root.wait_window(dialogo.janela)
        if dialogo.resultado is not None:
            self.passos.append(dialogo.resultado)
            self._atualizar_tree()
            self._log(f"Passo adicionado: {dialogo.resultado['tipo']} (índice {len(self.passos)})")

    def _editar_passo(self):
        sel = self.tree_passos.selection()
        if not sel:
            messagebox.showwarning("Editar", "Selecione um passo para editar.")
            return
        idx = self.tree_passos.index(sel[0])
        passo = self.passos[idx].copy()
        dialogo = JanelaEditarPasso(self.root, passo, self.colunas)
        self.root.wait_window(dialogo.janela)
        if dialogo.resultado is not None:
            self.passos[idx] = dialogo.resultado
            self._atualizar_tree()
            self._log(f"Passo {idx + 1} atualizado.")

    def _remover_passos(self):
        sel = self.tree_passos.selection()
        if not sel:
            return
        indices = sorted([self.tree_passos.index(s) for s in sel], reverse=True)
        for i in indices:
            del self.passos[i]
        self._atualizar_tree()
        self._log(f"{len(indices)} passo(s) removido(s).")

    def _limpar_passos(self):
        if not self.passos:
            return
        if messagebox.askyesno("Limpar", "Remover todos os passos?"):
            self.passos.clear()
            self._atualizar_tree()
            self._log("Todos os passos removidos.")

    def _subir_passo(self):
        sel = self.tree_passos.selection()
        if not sel:
            return
        idx = self.tree_passos.index(sel[0])
        if idx == 0:
            return
        self.passos[idx], self.passos[idx - 1] = self.passos[idx - 1], self.passos[idx]
        self._atualizar_tree()
        self.tree_passos.selection_set(self.tree_passos.get_children()[idx - 1])

    def _descer_passo(self):
        sel = self.tree_passos.selection()
        if not sel:
            return
        idx = self.tree_passos.index(sel[0])
        if idx >= len(self.passos) - 1:
            return
        self.passos[idx], self.passos[idx + 1] = self.passos[idx + 1], self.passos[idx]
        self._atualizar_tree()
        self.tree_passos.selection_set(self.tree_passos.get_children()[idx + 1])

    # ────────────────────────────────────────────────────────
    #  Salvar / Carregar config (passos)
    # ────────────────────────────────────────────────────────

    def _salvar_config(self):
        dados = {
            "modo": self.var_modo.get(),
            "data_inicial": self.var_data_inicial.get(),
            "data_final": self.var_data_final.get(),
            "formato_data": self.var_formato_data.get(),
            "caminho_planilha": self.caminho_planilha,
            "passos": self.passos,
            "delay": self.spin_delay.get(),
        }
        try:
            with open(ARQUIVO_CONFIG, "w", encoding="utf-8") as f:
                json.dump(dados, f, ensure_ascii=False, indent=2)
            self._log(f"Configuração salva em {ARQUIVO_CONFIG}")
            messagebox.showinfo("Salvo", f"Configuração salva em:\n{ARQUIVO_CONFIG}")
        except Exception as e:
            messagebox.showerror("Erro ao salvar", str(e))

    def _carregar_config(self):
        if ARQUIVO_CONFIG.exists():
            try:
                with open(ARQUIVO_CONFIG, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                self.passos = dados.get("passos", [])
                self.caminho_planilha = dados.get("caminho_planilha", "")
                delay = dados.get("delay", "0.3")
                self.spin_delay.set(str(delay))
                self.var_modo.set("datas" if dados.get("modo") == "datas" else "planilha")
                self.var_data_inicial.set(dados.get("data_inicial", time.strftime("%d/%m/%Y")))
                self.var_data_final.set(dados.get("data_final", time.strftime("%d/%m/%Y")))
                self.var_formato_data.set(dados.get("formato_data", "%d/%m/%Y"))
                self._trocar_modo()
                self._atualizar_tree()
                if self.caminho_planilha and os.path.exists(self.caminho_planilha):
                    # tenta recarregar automaticamente
                    try:
                        if self.caminho_planilha.endswith(".csv"):
                            self.planilha_df = pd.read_csv(self.caminho_planilha, encoding="utf-8-sig")
                            self.formatos = {}
                        else:
                            self.planilha_df = pd.read_excel(self.caminho_planilha, engine="openpyxl")
                            self._ler_formatos_excel(self.caminho_planilha)
                            self._garantir_coluna_lido()
                        self.colunas = list(self.planilha_df.columns)
                        nome = os.path.basename(self.caminho_planilha)
                        self.lbl_planilha.configure(text=f"{nome} ({len(self.planilha_df)} linhas)")
                        self.spin_linha_fim.set(len(self.planilha_df))
                        self._log(f"Planilha recarregada: {nome}")
                    except Exception:
                        pass
                self._log(f"Configuração carregada: {len(self.passos)} passo(s)")
            except Exception as e:
                self._log(f"Erro ao carregar config: {e}")

    def _carregar_config_manual(self):
        caminho = filedialog.askopenfilename(
            title="Carregar Config",
            filetypes=[("JSON", "*.json"), ("Todos", "*.*")]
        )
        if not caminho:
            return
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                dados = json.load(f)
            self.passos = dados.get("passos", [])
            self.caminho_planilha = dados.get("caminho_planilha", "")
            delay = dados.get("delay", "0.3")
            self.spin_delay.set(str(delay))
            self.var_modo.set("datas" if dados.get("modo") == "datas" else "planilha")
            self.var_data_inicial.set(dados.get("data_inicial", time.strftime("%d/%m/%Y")))
            self.var_data_final.set(dados.get("data_final", time.strftime("%d/%m/%Y")))
            self.var_formato_data.set(dados.get("formato_data", "%d/%m/%Y"))
            self._trocar_modo()
            self._atualizar_tree()
            self._log(f"Config carregada manualmente: {len(self.passos)} passo(s)")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    # ────────────────────────────────────────────────────────
    #  Execução
    # ────────────────────────────────────────────────────────

    def _executar(self):
        if not HAVE_PYNPUT:
            messagebox.showerror("Erro", "pynput não instalado.")
            return
        if not self.passos:
            messagebox.showwarning("Executar", "Nenhum passo gravado.")
            return

        modo = self.var_modo.get()
        if not HAVE_MSS and any(p.get("tipo") == "Validar Tela" for p in self.passos):
            messagebox.showerror("Erro", "Existe passo 'Validar Tela', mas a biblioteca mss não está instalada.\nExecute: pip install mss")
            return

        try:
            delay = float(self.spin_delay.get())
        except ValueError:
            delay = 0.3
        repeticoes = int(self.spin_repetir.get())

        if modo == "datas":
            try:
                formato = self.var_formato_data.get().strip() or "%d/%m/%Y"
                data_ini = _parse_data(self.var_data_inicial.get(), formato)
                data_fim = _parse_data(self.var_data_final.get(), formato)
            except ValueError:
                messagebox.showerror("Datas", f"Data inicial ou final inválida.\nUse o formato: {self.var_formato_data.get()}")
                return
            if data_fim < data_ini:
                messagebox.showwarning("Executar", "Data final é anterior à data inicial.")
                return
            total_dias = (data_fim - data_ini).days + 1
            self._datas_exec = [data_ini + datetime.timedelta(days=i) for i in range(total_dias)]
            inicio, fim = 0, total_dias
        else:
            if self.planilha_df is None:
                messagebox.showwarning("Executar", "Nenhuma planilha carregada.")
                return
            inicio = int(self.spin_linha_inicio.get()) - 1  # 0-index
            fim_str = self.spin_linha_fim.get()
            if fim_str.strip():
                fim = int(fim_str)
            else:
                fim = len(self.planilha_df)

            # Se usa coluna Lido, salta para a última não processada
            ultima = self._ultima_linha_processada()
            if ultima > inicio:
                self._log(f"Retomando da linha {ultima + 1} (já processadas até {ultima})")
                inicio = ultima

            # valida intervalo
            if inicio < 0:
                inicio = 0
            if fim > len(self.planilha_df):
                fim = len(self.planilha_df)
            if inicio >= fim:
                messagebox.showwarning("Executar", "Intervalo de linhas inválido.")
                return

        self.executando = True
        self.pausado = False
        self.play_parar.clear()
        self.play_continuar.clear()
        self.lote_processados = 0
        self.lote_ativo = self.var_lote.get()
        self.lote_tamanho = int(self.spin_lote.get()) if self.lote_ativo else 999999
        self.btn_executar.config(state=tk.DISABLED)
        self.btn_parar.config(state=tk.NORMAL)
        self.btn_gravar.config(state=tk.DISABLED)
        self.btn_adicionar.config(state=tk.DISABLED)
        total_passos = len(self.passos)
        total_unidades = fim - inicio
        self._linhas_restantes_exec = total_unidades * repeticoes
        self._tempo_passos = self._tempo_estimado()
        self._inicio_linha = time.time()
        self.lbl_status.configure(text="▶ Executando...", foreground="green")
        if modo == "datas":
            self._log(f"Iniciando execução: {total_unidades} dia(s) ({self._datas_exec[0]} → {self._datas_exec[-1]}), {repeticoes}x, delay {delay}s")
            self.lbl_progresso.configure(text=f"Data 0/{total_unidades} | Passo 0/{total_passos}")
        else:
            self.lbl_progresso.configure(text=f"Linha 0/{total_unidades} | Passo 0/{total_passos}")
            self._log(f"Iniciando execução: linhas {inicio + 1}–{fim}, {repeticoes}x, delay {delay}s")
        self._atualizar_tempo_restante()
        if self.lote_ativo:
            restantes = min(self.lote_tamanho, fim - inicio)
            self.lbl_restantes.configure(text=f"Restam: {restantes}")

        self.thread_play = threading.Thread(
            target=self._loop_execucao,
            args=(inicio, fim, delay, repeticoes),
            daemon=True
        )
        self.thread_play.start()

    def _parar_execucao(self):
        self.play_parar.set()
        self._salvar_planilha_marcacoes()
        self._log("Parando execução... (após o passo atual)")

    def _mouse_move(self, x, y):
        ctypes.windll.user32.SetCursorPos(x, y)

    def _mouse_click(self):
        ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
        time.sleep(0.02)
        ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)

    def _garantir_foco(self, x, y):
        # Ativa a janela que está sob o cursor para garantir o foco de teclado.
        # ERPs desktop precisam da janela em primeiro plano para receber teclas.
        try:
            user32 = ctypes.windll.user32
            p = ctypes.wintypes.POINT(x, y)
            hwnd = user32.WindowFromPoint(p)
            if hwnd:
                user32.SetForegroundWindow(hwnd)
                time.sleep(0.15)
                # Clique de reforço para focar o controle exato do campo
                ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
                time.sleep(0.05)
                ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)
                time.sleep(0.15)
        except Exception:
            pass

    def _send_key(self, vk, scan, up=False):
        # Usa SendInput (mais confiável que keybd_event, inclui scancode real)
        if not hasattr(self, "_INPUT"):
            ULONG_PTR = ctypes.POINTER(ctypes.c_ulong)

            class KEYBDINPUT(ctypes.Structure):
                _fields_ = [
                    ("wVk", ctypes.c_ushort),
                    ("wScan", ctypes.c_ushort),
                    ("dwFlags", ctypes.c_ulong),
                    ("time", ctypes.c_ulong),
                    ("dwExtraInfo", ULONG_PTR),
                ]

            class INPUTUNION(ctypes.Union):
                _fields_ = [("ki", KEYBDINPUT)]

            class INPUTST(ctypes.Structure):
                _fields_ = [("type", ctypes.c_ulong), ("union", INPUTUNION)]

            self._INPUT = INPUTST
            self._KEYBDINPUT = KEYBDINPUT

        flags = 0x0001  # KEYEVENTF_SCANCODE
        if up:
            flags |= 0x0002  # KEYEVENTF_KEYUP
        ki = self._KEYBDINPUT()
        ki.wVk = 0x00
        ki.wScan = scan
        ki.dwFlags = flags
        ki.time = 0
        ki.dwExtraInfo = None
        inp = self._INPUT()
        inp.type = 1  # INPUT_KEYBOARD
        inp.union.ki = ki
        ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))
        time.sleep(0.03)

    def _enviar_ctrl_a(self):
        # Ctrl+A (selecionar tudo) — tudo via SendInput/scancode para garantir combinação
        self._send_key(0x11, 0x1D, up=False)  # Ctrl down (scancode 0x1D)
        self._send_key(0x41, 0x1E, up=False)  # A down (scancode 0x1E)
        time.sleep(0.03)
        self._send_key(0x41, 0x1E, up=True)   # A up
        self._send_key(0x11, 0x1D, up=True)   # Ctrl up
        time.sleep(0.05)

    def _enviar_delete(self):
        # Delete (scancode 0x53)
        self._send_key(0x2E, 0x53, up=False)
        self._send_key(0x2E, 0x53, up=True)
        time.sleep(0.05)

    def _enviar_backspace(self):
        # Backspace (scancode 0x0E)
        self._send_key(0x08, 0x0E, up=False)
        self._send_key(0x08, 0x0E, up=True)
        time.sleep(0.05)

    def _enviar_home(self):
        # Home (scancode 0x47)
        self._send_key(0x24, 0x47, up=False)
        self._send_key(0x24, 0x47, up=True)
        time.sleep(0.05)

    def _enviar_shift_end(self):
        # Shift + End (seleciona do cursor até o fim)
        self._send_key(0x10, 0x2A, up=False)  # Shift down (scancode 0x2A)
        self._send_key(0x23, 0x4F, up=False)  # End down (scancode 0x4F)
        time.sleep(0.03)
        self._send_key(0x23, 0x4F, up=True)   # End up
        self._send_key(0x10, 0x2A, up=True)   # Shift up
        time.sleep(0.05)

    def _enviar_ctrl_home(self):
        # Ctrl + Home (vai ao início absoluto do campo)
        self._send_key(0x11, 0x1D, up=False)  # Ctrl down
        self._send_key(0x24, 0x47, up=False)  # Home down
        time.sleep(0.03)
        self._send_key(0x24, 0x47, up=True)   # Home up
        self._send_key(0x11, 0x1D, up=True)   # Ctrl up
        time.sleep(0.05)

    def _enviar_ctrl_shift_end(self):
        # Ctrl + Shift + End (seleciona do cursor até o fim absoluto do campo)
        self._send_key(0x11, 0x1D, up=False)  # Ctrl down
        self._send_key(0x10, 0x2A, up=False)  # Shift down
        self._send_key(0x23, 0x4F, up=False)  # End down
        time.sleep(0.03)
        self._send_key(0x23, 0x4F, up=True)   # End up
        self._send_key(0x10, 0x2A, up=True)   # Shift up
        self._send_key(0x11, 0x1D, up=True)   # Ctrl up
        time.sleep(0.05)

    def _aplicar_pre_digitacao(self, passo):
        # Garante que o campo fique vazio antes de digitar.
        # Estratégia para ERP desktop (controles customizados):
        #  - Ativa a janela (foco de teclado)
        #  - Várias passadas de limpeza para garantir que tudo seja apagado:
        #    Home + Shift+End (seleciona do início ao fim) + Delete
        self._garantir_foco(passo["x"], passo["y"])
        selecionar = passo.get("selecionar_tudo", False)
        apagar = passo.get("apagar", False)

        if selecionar:
            self._log("        Selecionando todo o conteúdo...")
            self._enviar_ctrl_a()
            time.sleep(0.1)

        if apagar:
            # Este ERP não processa combinações com Ctrl (Ctrl+A/Ctrl+Shift+End falham).
            # Só Shift funciona. Então: Home + Shift+End (seleciona até o fim) e apaga
            # com Delete e Backspace.
            for tentativa in range(3):
                self._enviar_home()          # vai ao início do campo
                time.sleep(0.05)
                self._enviar_shift_end()     # seleciona até o fim (Shift funciona)
                time.sleep(0.08)
                self._enviar_delete()        # tenta apagar a seleção
                time.sleep(0.05)
                self._enviar_backspace()     # alternativa de apagar
                time.sleep(0.1)
                self._log(f"        Passada de limpeza {tentativa + 1}/3")
            # Fallback: repete Backspace e Delete rapidamente (com foco, apaga o resto)
            self._log("        Limpeza por repetição (Backspace/Delete)...")
            for _ in range(40):
                if self.play_parar.is_set():
                    break
                self._enviar_backspace()
                self._enviar_delete()
            self._log("        Campo limpo")

    def _colar_texto(self, texto):
        for tentativa in range(10):
            if ctypes.windll.user32.OpenClipboard(None):
                break
            time.sleep(0.2)
        else:
            self._log(f"        ERRO clipboard: não abriu após 10 tentativas")
            self._digitar_texto(texto)
            return

        ctypes.windll.user32.EmptyClipboard()
        CF_UNICODETEXT = 13
        GMEM_MOVABLE = 2

        buffer = ctypes.windll.kernel32.GlobalAlloc(GMEM_MOVABLE, (len(texto) + 1) * 2)
        if not buffer:
            ctypes.windll.user32.CloseClipboard()
            self._digitar_texto(texto)
            return

        dados = ctypes.windll.kernel32.GlobalLock(buffer)
        if not dados:
            ctypes.windll.kernel32.GlobalFree(buffer)
            ctypes.windll.user32.CloseClipboard()
            self._digitar_texto(texto)
            return

        ctypes.memmove(dados, texto.encode("utf-16-le"), len(texto) * 2 + 2)
        ctypes.windll.kernel32.GlobalUnlock(buffer)
        ctypes.windll.user32.SetClipboardData(CF_UNICODETEXT, buffer)
        ctypes.windll.user32.CloseClipboard()
        time.sleep(0.05)

        import subprocess
        subprocess.run(
            ["powershell", "-Command",
             '[System.Windows.Forms.SendKeys]::SendWait("^{v}")'],
            capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW
        )

    def _enviar_tecla(self, user32, vk, shift):
        if shift:
            user32.keybd_event(0x10, 0, 0, 0)   # Shift down
            time.sleep(0.01)
        user32.keybd_event(vk, 0, 0, 0)
        time.sleep(0.01)
        user32.keybd_event(vk, 0, 2, 0)
        time.sleep(0.01)
        if shift:
            user32.keybd_event(0x10, 0, 2, 0)   # Shift up
            time.sleep(0.01)

    def _mapping_char(self, char):
        # Retorna (vk, shift) para caracteres ASCII mapeáveis, ou None se não suportado
        if char in _CARACTERES_ESPECIAIS:
            return _CARACTERES_ESPECIAIS[char]
        if char == " ":
            return (0x20, False)
        if char.isascii():
            if char.isalpha():
                return (ord(char.upper()), char.isupper())
            if char.isdigit():
                return (ord(char), False)
        return None

    def _colar_char_clipboard(self, char):
        # Cola um único caractere via clipboard (útil para não-ASCII: acentos, ç, etc.)
        for _ in range(5):
            if ctypes.windll.user32.OpenClipboard(None):
                break
            time.sleep(0.1)
        else:
            return False
        ctypes.windll.user32.EmptyClipboard()
        CF_UNICODETEXT = 13
        GMEM_MOVABLE = 2
        buffer = ctypes.windll.kernel32.GlobalAlloc(GMEM_MOVABLE, (len(char) + 1) * 2)
        if not buffer:
            ctypes.windll.user32.CloseClipboard()
            return False
        dados = ctypes.windll.kernel32.GlobalLock(buffer)
        if not dados:
            ctypes.windll.kernel32.GlobalFree(buffer)
            ctypes.windll.user32.CloseClipboard()
            return False
        ctypes.memmove(dados, char.encode("utf-16-le"), len(char) * 2 + 2)
        ctypes.windll.kernel32.GlobalUnlock(buffer)
        ctypes.windll.user32.SetClipboardData(CF_UNICODETEXT, buffer)
        ctypes.windll.user32.CloseClipboard()
        time.sleep(0.05)
        import subprocess
        subprocess.run(
            ["powershell", "-Command",
             '[System.Windows.Forms.SendKeys]::SendWait("^{v}")'],
            capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW
        )
        return True

    def _digitar_texto(self, texto):
        self._log(f"        Modo digitação direta ({len(texto)} caracteres)...")
        user32 = ctypes.windll.user32
        for _ in range(20):
            user32.keybd_event(0x08, 0, 0, 0)   # Backspace para limpar campo
            time.sleep(0.01)
            user32.keybd_event(0x08, 0, 2, 0)
            time.sleep(0.01)
        time.sleep(0.05)
        for char in texto:
            if self.play_parar.is_set():
                return
            par = self._mapping_char(char)
            if par is not None:
                vk, shift = par
                self._enviar_tecla(user32, vk, shift)
            else:
                # Caracteres não ASCII (acentos, ç, símbolos) — cola via clipboard
                self._colar_char_clipboard(char)
            time.sleep(0.02)

    def _loop_execucao(self, inicio, fim, delay, repeticoes):
        modo = self.var_modo.get()
        indices = list(range(inicio, fim))
        total_unidades = len(indices)
        total_passos = len(self.passos)

        for rep in range(repeticoes):
            if self.play_parar.is_set():
                break

            for n in range(total_unidades):
                if self.play_parar.is_set():
                    break

                idx_unidade = indices[n]
                if modo == "datas":
                    data = self._datas_exec[idx_unidade]
                    contexto = {"data": data, "linha": None, "idx_linha": None}
                    base = f"Dia {data.strftime('%d/%m/%Y')}"
                    rot = f"Data {n+1}/{total_unidades}"
                else:
                    id_linha = idx_unidade
                    linha = self.planilha_df.iloc[id_linha]
                    contexto = {"data": None, "linha": linha, "idx_linha": id_linha}
                    base = f"L{id_linha+1}"
                    rot = f"Linha {id_linha+1-inicio}/{total_unidades}"
                    self._marcar_linha_processada(id_linha)

                self._pos_unidade = n
                self._pular_unidade = False
                self.lote_processados += 1
                self._inicio_linha = time.time()
                self._atualizar_tempo_restante()
                if self.lote_ativo:
                    restam = self.lote_tamanho - self.lote_processados
                    self.lbl_restantes.after(0, self.lbl_restantes.configure,
                                             {"text": f"Restam: {restam}"})

                for i, passo in enumerate(self.passos):
                    if self.play_parar.is_set():
                        break

                    prog = f"{rot} | Passo {i+1}/{total_passos}"
                    self.lbl_progresso.after(0, self.lbl_progresso.configure, {"text": prog})

                    self._executar_passo(passo, contexto, f"{base} P{i+1}",
                                         delay, ultimo=(i == total_passos - 1))

                    if self._pular_unidade:
                        self._log(f"{base} Sequência ignorada — avançando para o próximo item")
                        break

                self._linhas_restantes_exec -= 1

                # Verifica se atingiu o limite do lote
                if self.lote_ativo and self.lote_processados >= self.lote_tamanho:
                    unidade_nome = "dia(s)" if modo == "datas" else "linhas"
                    self.pausado = True
                    self.play_continuar.clear()
                    self.lbl_status.after(0, self.lbl_status.configure,
                                          {"text": "⏸ Aguardando F1", "foreground": "orange"})
                    self._log(f"——— Lote de {self.lote_tamanho} {unidade_nome} concluído ———")
                    self._salvar_planilha_marcacoes()
                    self._log("Salve as alterações e pressione F1 para continuar")
                    messagebox.showinfo("Lote Concluído",
                                        f"{self.lote_tamanho} {unidade_nome} processados.\n"
                                        "Salve as alterações e pressione F1 para continuar.")
                    # Aguarda F1
                    self.play_continuar.wait()
                    self.pausado = False
                    self._salvar_planilha_marcacoes()
                    if modo == "datas":
                        self.lote_processados = 0
                        self._log("Continuando execução...")
                        self.lbl_status.after(0, self.lbl_status.configure,
                                              {"text": "▶ Executando...", "foreground": "green"})
                    else:
                        # Recarrega a planilha para refletir marcações
                        self._log("Recarregando planilha...")
                        try:
                            self.planilha_df = pd.read_excel(self.caminho_planilha, engine="openpyxl")
                            self._ler_formatos_excel(self.caminho_planilha)
                            self.colunas = list(self.planilha_df.columns)
                            nome = os.path.basename(self.caminho_planilha)
                            self.lbl_planilha.configure(text=f"{nome} ({len(self.planilha_df)} linhas)")
                            self.spin_linha_fim.set(len(self.planilha_df))
                        except Exception as e:
                            self._log(f"ERRO ao recarregar planilha: {e}")
                            self._parar_execucao()
                            break
                        # Pula para a primeira linha não lida
                        inicio = self._ultima_linha_processada()
                        fim = len(self.planilha_df)
                        total_unidades = fim - inicio
                        self.lote_processados = 0
                        self._linhas_restantes_exec = total_unidades
                        self._log(f"Planilha recarregada: {total_unidades} linha(s) restante(s)")
                        self.lbl_status.after(0, self.lbl_status.configure,
                                              {"text": "▶ Executando...", "foreground": "green"})

            if rep < repeticoes - 1 and not self.play_parar.is_set():
                time.sleep(1)

        self._salvar_planilha_marcacoes()
        self._finalizar_execucao()

    def _executar_passo(self, passo, contexto, prog, delay, ultimo=True):
        tipo = passo["tipo"]
        try:
            if tipo == "Espera":
                espera = passo.get("espera", 0.5)
                if espera > 0:
                    time.sleep(espera)
                return

            if tipo == "Pular Data":
                self._pular_unidade = True
                self._log(f"{prog} Pular Data — avançando para o próximo dia/linha")
                return

            if tipo == "Validar Tela":
                self._executar_validacao(passo, contexto, prog, delay)
            elif tipo == "Print":
                self._executar_print(passo, contexto, prog)
            elif tipo == "Registrar Ignorado":
                self._registrar_ignorado(passo, contexto, prog)
            else:
                x, y = passo["x"], passo["y"]
                self._mouse_move(x, y)
                time.sleep(0.05)

                if tipo == "Clique":
                    self._mouse_click()
                    self._log(f"{prog} Clique ({x},{y}) OK")

                elif tipo == "Planilha":
                    linha = contexto.get("linha")
                    coluna = passo.get("coluna", "")
                    self._mouse_click()
                    time.sleep(0.15)
                    if linha is not None and coluna and coluna in self.planilha_df.columns:
                        valor = str(linha[coluna]) if pd.notna(linha[coluna]) else ""
                        tamanho_min = passo.get("tamanho_min", 0)
                        if tamanho_min > 0 and len(valor) < tamanho_min:
                            valor = valor.zfill(tamanho_min)
                        self._log(f"{prog} '{coluna}'=\"{valor}\" ({x},{y})")
                        self._aplicar_pre_digitacao(passo)
                        self._colar_texto(valor)
                    else:
                        self._log(f"{prog} Planilha ({x},{y}) ERRO: coluna '{coluna}' não encontrada")

                elif tipo == "Texto":
                    self._mouse_click()
                    time.sleep(0.15)
                    valor = passo.get("texto_fixo", "")
                    self._log(f"{prog} Texto=\"{valor}\" ({x},{y})")
                    self._aplicar_pre_digitacao(passo)
                    self._colar_texto(valor)

                elif tipo == "Data":
                    data = contexto.get("data")
                    if data is None:
                        self._log(f"{prog} Data ERRO: este passo exige o modo Datas")
                    else:
                        self._mouse_click()
                        time.sleep(0.15)
                        formato = passo.get("formato_data", "%d/%m/%Y")
                        valor = data.strftime(formato)
                        self._log(f"{prog} Data=\"{valor}\" ({x},{y})")
                        self._aplicar_pre_digitacao(passo)
                        self._colar_texto(valor)

                else:
                    self._log(f"{prog} Tipo desconhecido: {tipo}")

            # Aguarda o tempo configurado no passo antes de ir ao próximo
            if not ultimo and not self._pular_unidade:
                espera_passo = 0.0
                try:
                    espera_passo = float(passo.get("espera", 0))
                except (TypeError, ValueError):
                    espera_passo = 0.0
                if espera_passo > 0:
                    time.sleep(espera_passo)
                elif delay > 0:
                    time.sleep(delay)

        except Exception as e:
            self._log(f"{prog} ERRO: {e}")
            time.sleep(1)

    def _executar_print(self, passo, contexto, prog):
        monitor_idx = passo.get("monitor", 1)
        pasta = passo.get("pasta_print", "")
        coluna_nome = passo.get("coluna_nome", "")
        linha = contexto.get("linha")
        if not HAVE_MSS:
            self._log(f"{prog} Print ERRO: mss não instalado")
            return
        if not pasta:
            self._log(f"{prog} Print ERRO: pasta não definida")
            return
        try:
            os.makedirs(pasta, exist_ok=True)
            with mss.mss() as sct:
                monitors = sct.monitors[1:]
                if monitor_idx < 1 or monitor_idx > len(monitors):
                    self._log(f"{prog} Print ERRO: monitor {monitor_idx} inválido")
                    return
                monitor = monitors[monitor_idx - 1]
                img = sct.grab(monitor)
                nome_arq = None
                if coluna_nome and linha is not None and coluna_nome in self.planilha_df.columns:
                    raw = linha[coluna_nome]
                    nome_raw = str(raw) if pd.notna(raw) else ""
                    nome_arq = re.sub(r'[\\/*?:"<>|]', '_', nome_raw).strip()
                    if not nome_arq:
                        nome_arq = None
                if nome_arq is None:
                    idx = contexto.get("idx_linha")
                    if idx is not None:
                        nome_arq = f"print_{idx + 1}"
                    else:
                        nome_arq = re.sub(r'[\\/*?:"<>|]', '_', prog).strip() or "print"
                if not nome_arq.lower().endswith('.png'):
                    nome_arq += '.png'
                caminho = os.path.join(pasta, nome_arq)
                mss.tools.to_png(img.rgb, img.size, output=caminho)
                self._log(f"{prog} Print salvo: {caminho}")
        except Exception as e:
            self._log(f"{prog} Print ERRO: {e}")

    def _registrar_ignorado(self, passo, contexto, prog):
        caminho = passo.get("arquivo_ignorados", "").strip()
        if not caminho:
            self._log(f"{prog} Registrar Ignorado ERRO: arquivo não definido")
            return
        data = contexto.get("data")
        texto = ""
        try:
            if data is not None:
                formato = passo.get("formato_data", "%d/%m/%Y") or "%d/%m/%Y"
                texto = data.strftime(formato)
            else:
                linha = contexto.get("linha")
                if linha is not None:
                    col = passo.get("coluna", "")
                    if col and self.planilha_df is not None and col in self.planilha_df.columns:
                        raw = linha[col]
                        texto = str(raw) if pd.notna(raw) else ""
                    if not texto:
                        texto = f"Linha {contexto.get('idx_linha', 0) + 1}"
                else:
                    texto = ""
            if not texto:
                self._log(f"{prog} Registrar Ignorado ERRO: não há conteúdo para registrar")
                return
            caminho = os.path.expandvars(os.path.expanduser(caminho))
            pasta = os.path.dirname(caminho)
            if pasta:
                os.makedirs(pasta, exist_ok=True)
            with open(caminho, "a", encoding="utf-8") as arquivo:
                arquivo.write(texto + "\n")
            self._log(f"{prog} Ignorado registrado \"{texto}\" -> {caminho}")
        except Exception as e:
            self._log(f"{prog} Registrar Ignorado ERRO: {e}")

    def _executar_validacao(self, passo, contexto, prog, delay):
        if not HAVE_MSS:
            self._log(f"{prog} Validar Tela ERRO: mss não instalado")
            return
        regiao = passo.get("regiao") or {}
        referencia = passo.get("referencia", "")
        acoes = passo.get("acoes", [])
        acoes_nao = passo.get("acoes_nao_encontrado", [])
        try:
            limiar = float(passo.get("limiar", 0.85))
            tolerancia = int(passo.get("tolerancia", 30))
            timeout = float(passo.get("timeout", 3.0))
        except (TypeError, ValueError):
            self._log(f"{prog} Validar Tela ERRO: limiar/tolerância/timeout inválidos")
            return
        larg = int(regiao.get("largura", 0)) if regiao else 0
        alt = int(regiao.get("altura", 0)) if regiao else 0
        if not regiao or larg <= 0 or alt <= 0:
            self._log(f"{prog} Validar Tela ERRO: defina a região no editor (botão Selecionar Região)")
            return
        if not referencia or not os.path.exists(referencia):
            self._log(f"{prog} Validar Tela ERRO: defina a referência no editor (botão Capturar)")
            return
        ref = _carregar_referencia(referencia)
        if ref is None:
            self._log(f"{prog} Validar Tela ERRO: referência inválida: {referencia}")
            return
        if ref.get("largura") != larg or ref.get("altura") != alt:
            self._log(f"{prog} Validar Tela ERRO: a região não corresponde à referência — capture novamente")
            return
        rect = {"left": int(regiao.get("x", 0)), "top": int(regiao.get("y", 0)),
                "width": larg, "height": alt}
        limite = max(timeout, 0.1)
        inicio = time.time()
        achou = False
        score = 0.0
        try:
            with mss.mss() as sct:
                while True:
                    if self.play_parar.is_set():
                        return
                    img = sct.grab(rect)
                    score = _comparar_rgb(ref.get("rgb"), img.rgb, tolerancia)
                    if score >= limiar:
                        achou = True
                        break
                    if time.time() - inicio >= limite:
                        break
                    time.sleep(0.3)
        except Exception as e:
            self._log(f"{prog} Validar Tela ERRO: {e}")
            return

        if not achou:
            self._log(f"{prog} Tela NÃO detectada em {limite:.1f}s (semelhança {score:.0%})")
            if acoes_nao:
                self._log(f"{prog} Executando {len(acoes_nao)} ação(ões) 'se NÃO aparecer'")
                for j, acao in enumerate(acoes_nao):
                    if self.play_parar.is_set():
                        break
                    self._executar_passo(acao, contexto, f"{prog} N{j+1}",
                                         delay, ultimo=(j == len(acoes_nao) - 1))
            return
        self._log(f"{prog} Tela DETECTADA (semelhança {score:.0%}) — executando {len(acoes)} ação(ões)")
        for j, acao in enumerate(acoes):
            if self.play_parar.is_set():
                break
            self._executar_passo(acao, contexto, f"{prog} A{j+1}", delay, ultimo=(j == len(acoes) - 1))

    def _finalizar_execucao(self):
        self.executando = False
        self.root.after(0, self._pos_execucao)

    def _pos_execucao(self):
        self.btn_executar.config(state=tk.NORMAL)
        self.btn_parar.config(state=tk.DISABLED)
        self.btn_gravar.config(state=tk.NORMAL)
        self.btn_adicionar.config(state=tk.NORMAL)
        if self.play_parar.is_set():
            self.lbl_status.configure(text="⏹ Interrompido", foreground="red")
            self.lbl_progresso.configure(text="")
            self._log("EXECUÇÃO INTERROMPIDA pelo usuário")
        else:
            self.lbl_status.configure(text="✔ Concluído", foreground="green")
            self.lbl_progresso.configure(text="")
            self._log("EXECUÇÃO CONCLUÍDA com sucesso!")
            messagebox.showinfo("Execução Concluída",
                                "Todas as linhas foram processadas com sucesso!")
        self._atualizar_tempo_restante()


# ═══════════════════════════════════════════════════════════
#  Janela de edição de passo
# ═══════════════════════════════════════════════════════════

class JanelaEditarPasso:
    def __init__(self, parent, passo, colunas):
        self.resultado = None
        self.colunas = colunas
        self.parent = parent

        self.janela = tk.Toplevel(parent)
        self.janela.title("Editar Passo")
        self.janela.geometry("660x480")
        self.janela.transient(parent)
        self.janela.grab_set()

        self.frame = ttk.Frame(self.janela, padding=12)
        self.frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(self.frame, text="Tipo do Passo:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.tipo_var = tk.StringVar(value=passo.get("tipo", "Clique"))
        combo_tipo = ttk.Combobox(self.frame, textvariable=self.tipo_var,
                                  values=["Clique", "Planilha", "Texto", "Data", "Espera", "Pular Data", "Registrar Ignorado", "Print", "Validar Tela"],
                                  state="readonly", width=20)
        combo_tipo.grid(row=0, column=1, sticky=tk.W, pady=4)
        combo_tipo.bind("<<ComboboxSelected>>", self._on_tipo_change)

        self.lbl_pos_x = ttk.Label(self.frame, text="Posição X:")
        self.lbl_pos_x.grid(row=1, column=0, sticky=tk.W, pady=4)
        self.x_var = tk.StringVar(value=str(passo.get("x", 0)))
        self.entry_x = ttk.Entry(self.frame, textvariable=self.x_var, width=12)
        self.entry_x.grid(row=1, column=1, sticky=tk.W, pady=4)

        self.lbl_pos_y = ttk.Label(self.frame, text="Posição Y:")
        self.lbl_pos_y.grid(row=2, column=0, sticky=tk.W, pady=4)
        self.y_var = tk.StringVar(value=str(passo.get("y", 0)))
        self.entry_y = ttk.Entry(self.frame, textvariable=self.y_var, width=12)
        self.entry_y.grid(row=2, column=1, sticky=tk.W, pady=4)

        # Widgets condicionais (criados mas ocultos até o tipo certo)
        self.lbl_coluna = ttk.Label(self.frame, text="Coluna da Planilha:")
        self.coluna_var = tk.StringVar(value=passo.get("coluna", ""))
        self.combo_coluna = ttk.Combobox(self.frame, textvariable=self.coluna_var, values=colunas, width=30)

        self.lbl_texto = ttk.Label(self.frame, text="Texto a digitar:")
        self.texto_fixo_var = tk.StringVar(value=passo.get("texto_fixo", ""))
        self.entry_texto = ttk.Entry(self.frame, textvariable=self.texto_fixo_var, width=40)

        self.lbl_formato = ttk.Label(self.frame, text="Formato da data:")
        self.formato_data_var = tk.StringVar(value=passo.get("formato_data", "%d/%m/%Y"))
        self.entry_formato = ttk.Entry(self.frame, textvariable=self.formato_data_var, width=14)

        self.lbl_tamanho = ttk.Label(self.frame, text="Mín. caracteres (zeros à esq):")
        self.tamanho_var = tk.StringVar(value=str(passo.get("tamanho_min", 0)))
        self.entry_tamanho = ttk.Entry(self.frame, textvariable=self.tamanho_var, width=10)

        self.selecionar_var = tk.BooleanVar(value=bool(passo.get("selecionar_tudo", False)))
        self.chk_selecionar = ttk.Checkbutton(self.frame, text="Selecionar tudo (Ctrl+A) antes",
                                              variable=self.selecionar_var)
        self.apagar_var = tk.BooleanVar(value=bool(passo.get("apagar", False)))
        self.chk_apagar = ttk.Checkbutton(self.frame, text="Apagar (Delete) antes",
                                          variable=self.apagar_var)

        self.lbl_monitor = ttk.Label(self.frame, text="Monitor:")
        monitor_val = passo.get("monitor", 1)
        self.monitor_var = tk.StringVar(value=str(monitor_val))
        self.combo_monitor = ttk.Combobox(self.frame, textvariable=self.monitor_var,
                                          values=[], state="readonly", width=30)
        self._listar_monitores()

        self.lbl_pasta_print = ttk.Label(self.frame, text="Pasta para salvar:")
        self.pasta_print_var = tk.StringVar(value=passo.get("pasta_print", ""))
        self.frame_pasta_print = ttk.Frame(self.frame)
        self.entry_pasta_print = ttk.Entry(self.frame_pasta_print, textvariable=self.pasta_print_var, width=35)
        self.entry_pasta_print.pack(side=tk.LEFT)
        self.btn_pasta_print = ttk.Button(self.frame_pasta_print, text="...", width=3,
                                          command=self._selecionar_pasta_print)
        self.btn_pasta_print.pack(side=tk.LEFT, padx=(4, 0))

        self.lbl_coluna_nome = ttk.Label(self.frame, text="Coluna p/ nome do arquivo:")
        self.coluna_nome_var = tk.StringVar(value=passo.get("coluna_nome", ""))
        self.combo_coluna_nome = ttk.Combobox(self.frame, textvariable=self.coluna_nome_var,
                                              values=colunas, width=30)

        self.lbl_arquivo_ignorados = ttk.Label(self.frame, text="Arquivo de ignorados:")
        self.arquivo_ignorados_var = tk.StringVar(value=passo.get("arquivo_ignorados", ""))
        self.frame_arquivo_ignorados = ttk.Frame(self.frame)
        self.entry_arquivo_ignorados = ttk.Entry(self.frame_arquivo_ignorados,
                                                 textvariable=self.arquivo_ignorados_var, width=40)
        self.entry_arquivo_ignorados.pack(side=tk.LEFT)
        self.btn_arquivo_ignorados = ttk.Button(self.frame_arquivo_ignorados, text="...", width=3,
                                                command=self._escolher_arquivo_ignorados)
        self.btn_arquivo_ignorados.pack(side=tk.LEFT, padx=(4, 0))

        # --- widgets do tipo Validar Tela ---
        self.frame_validacao = ttk.LabelFrame(self.frame, text="Validação de Tela", padding=8)
        reg = passo.get("regiao") or {"x": 0, "y": 0, "largura": 100, "altura": 100}
        self.reg_x_var = tk.StringVar(value=str(reg.get("x", 0)))
        self.reg_y_var = tk.StringVar(value=str(reg.get("y", 0)))
        self.larg_var = tk.StringVar(value=str(reg.get("largura", 100)))
        self.alt_var = tk.StringVar(value=str(reg.get("altura", 100)))

        ttk.Label(self.frame_validacao, text="Região X:").grid(row=0, column=0, sticky=tk.W, padx=(0, 2), pady=2)
        ttk.Entry(self.frame_validacao, textvariable=self.reg_x_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=2)
        ttk.Label(self.frame_validacao, text="Y:").grid(row=0, column=2, sticky=tk.W, padx=(8, 2), pady=2)
        ttk.Entry(self.frame_validacao, textvariable=self.reg_y_var, width=10).grid(row=0, column=3, sticky=tk.W, pady=2)

        ttk.Label(self.frame_validacao, text="Largura:").grid(row=1, column=0, sticky=tk.W, padx=(0, 2), pady=2)
        ttk.Entry(self.frame_validacao, textvariable=self.larg_var, width=10).grid(row=1, column=1, sticky=tk.W, pady=2)
        ttk.Label(self.frame_validacao, text="Altura:").grid(row=1, column=2, sticky=tk.W, padx=(8, 2), pady=2)
        ttk.Entry(self.frame_validacao, textvariable=self.alt_var, width=10).grid(row=1, column=3, sticky=tk.W, pady=2)

        self.referencia_var = tk.StringVar(value=passo.get("referencia", ""))
        ttk.Label(self.frame_validacao, text="Referência:").grid(row=2, column=0, sticky=tk.W, padx=(0, 2), pady=2)
        ttk.Entry(self.frame_validacao, textvariable=self.referencia_var, width=28).grid(row=2, column=1, columnspan=2, sticky=tk.W, pady=2)
        ttk.Button(self.frame_validacao, text="...", width=3, command=self._escolher_referencia).grid(row=2, column=3, sticky=tk.W, pady=2)

        ttk.Button(self.frame_validacao, text="Selecionar Região...",
                   command=self._selecionar_regiao).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(4, 2))
        ttk.Button(self.frame_validacao, text="Capturar Referência",
                   command=self._capturar_referencia).grid(row=3, column=2, columnspan=2, sticky=tk.W, pady=(4, 2))

        self.limiar_var = tk.StringVar(value=str(int(float(passo.get("limiar", 0.85)) * 100)))
        self.tol_var = tk.StringVar(value=str(passo.get("tolerancia", 30)))
        self.timeout_var = tk.StringVar(value=str(passo.get("timeout", 3.0)))
        ttk.Label(self.frame_validacao, text="Semelhança mínima (%):").grid(row=4, column=0, sticky=tk.W, padx=(0, 2), pady=2)
        ttk.Entry(self.frame_validacao, textvariable=self.limiar_var, width=10).grid(row=4, column=1, sticky=tk.W, pady=2)
        ttk.Label(self.frame_validacao, text="Tol. cor (0-255):").grid(row=4, column=2, sticky=tk.W, padx=(8, 2), pady=2)
        ttk.Entry(self.frame_validacao, textvariable=self.tol_var, width=10).grid(row=4, column=3, sticky=tk.W, pady=2)
        ttk.Label(self.frame_validacao, text="Timeout (s):").grid(row=5, column=0, sticky=tk.W, padx=(0, 2), pady=2)
        ttk.Entry(self.frame_validacao, textvariable=self.timeout_var, width=10).grid(row=5, column=1, sticky=tk.W, pady=2)

        # --- ações condicionais (sub-passos) ---
        self.acoes = list(passo.get("acoes", []))
        self.acoes_nao = list(passo.get("acoes_nao_encontrado", []))

        self.notebook_acoes = ttk.Notebook(self.frame_validacao)
        self.notebook_acoes.grid(row=6, column=0, columnspan=4, sticky=tk.EW, pady=(6, 0))

        cols_acoes = ("#", "Tipo", "Posição", "Espera (s)", "Rótulo")
        for texto_aba, chave in (("Se a tela APARECER", "sim"), ("Se a tela NÃO aparecer", "nao")):
            aba = ttk.Frame(self.notebook_acoes, padding=4)
            self.notebook_acoes.add(aba, text=texto_aba)
            tree = ttk.Treeview(aba, columns=cols_acoes, show="headings",
                                selectmode="browse", height=4)
            for c, l in zip(cols_acoes, [36, 90, 130, 70, 150]):
                tree.heading(c, text=c)
                tree.column(c, width=l, anchor=tk.CENTER if c != "Rótulo" else tk.W)
            tree.pack(fill=tk.X, pady=(0, 2))
            if chave == "sim":
                self.tree_acoes = tree
            else:
                self.tree_acoes_nao = tree
            frame_btn_acoes = ttk.Frame(aba)
            frame_btn_acoes.pack(fill=tk.X)
            ttk.Button(frame_btn_acoes, text="Adicionar...", width=12,
                       command=lambda a=chave: self._acao_menu_adicionar(a)).pack(side=tk.LEFT, padx=(0, 4))
            ttk.Button(frame_btn_acoes, text="Editar", width=8,
                       command=lambda a=chave: self._acao_editar(a)).pack(side=tk.LEFT, padx=2)
            ttk.Button(frame_btn_acoes, text="Remover", width=8,
                       command=lambda a=chave: self._acao_remover(a)).pack(side=tk.LEFT, padx=2)
            ttk.Button(frame_btn_acoes, text="▲", width=3,
                       command=lambda a=chave: self._acao_mover(-1, a)).pack(side=tk.LEFT, padx=2)
            ttk.Button(frame_btn_acoes, text="▼", width=3,
                       command=lambda a=chave: self._acao_mover(1, a)).pack(side=tk.LEFT, padx=2)
        self._atualizar_acoes_tree("sim")
        self._atualizar_acoes_tree("nao")

        for c in range(4):
            self.frame_validacao.columnconfigure(c, weight=1)

        self.lbl_hint_pular = ttk.Label(self.frame,
                                        text="Pula para o próximo dia/linha e interrompe a sequência atual.",
                                        foreground="#888888")
        self.lbl_hint_pular.grid_remove()

        ttk.Label(self.frame, text="Tempo de espera (s):").grid(row=6, column=0, sticky=tk.W, pady=4)
        self.espera_var = tk.StringVar(value=str(passo.get("espera", 0.5)))
        ttk.Entry(self.frame, textvariable=self.espera_var, width=10).grid(row=6, column=1, sticky=tk.W, pady=4)

        ttk.Label(self.frame, text="Rótulo / Descrição:").grid(row=7, column=0, sticky=tk.W, pady=4)
        self.rotulo_var = tk.StringVar(value=passo.get("rotulo", ""))
        ttk.Entry(self.frame, textvariable=self.rotulo_var, width=40).grid(row=7, column=1, sticky=tk.W, pady=4)

        frame_botoes = ttk.Frame(self.janela)
        frame_botoes.pack(fill=tk.X, padx=12, pady=(0, 12))
        ttk.Button(frame_botoes, text="Salvar", command=self._salvar).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(frame_botoes, text="Cancelar", command=self._cancelar).pack(side=tk.RIGHT, padx=(4, 0))

        self._reorganizar(None)
        self.janela.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.janela.focus_set()
        tema.aplicar_tema(self.janela, TEMA)

    def _on_tipo_change(self, event=None):
        self._reorganizar(event)

    def _listar_monitores(self):
        monitores = []
        if HAVE_MSS:
            with mss.mss() as sct:
                for i, m in enumerate(sct.monitors[1:], 1):
                    monitores.append(f"{i} - {m['width']}x{m['height']} ({m['left']},{m['top']})")
        if not monitores:
            monitores = ["1 - Monitor primário"]
        self.combo_monitor["values"] = monitores
        idx = int(self.monitor_var.get()) - 1
        if 0 <= idx < len(monitores):
            self.combo_monitor.current(idx)
        else:
            self.combo_monitor.current(0)

    def _selecionar_pasta_print(self):
        pasta = filedialog.askdirectory(title="Selecionar pasta para salvar prints")
        if pasta:
            self.pasta_print_var.set(pasta)

    def _escolher_arquivo_ignorados(self):
        caminho = filedialog.asksaveasfilename(
            title="Arquivo onde registrar os ignorados",
            defaultextension=".txt",
            filetypes=[("Texto", "*.txt"), ("Todos", "*.*")])
        if caminho:
            self.arquivo_ignorados_var.set(caminho)

    def _selecionar_regiao(self):
        JanelaSelecionarRegiao(self.parent, self._regiao_callback)

    def _regiao_callback(self, dados):
        if not dados:
            return  # cancelado
        x, y, larg, alt = dados
        self.reg_x_var.set(str(x))
        self.reg_y_var.set(str(y))
        self.larg_var.set(str(larg))
        self.alt_var.set(str(alt))

    def _escolher_referencia(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar referência (.json capturado)",
            filetypes=[("JSON", "*.json"), ("Todos", "*.*")])
        if caminho:
            self.referencia_var.set(caminho)

    def _capturar_referencia(self):
        try:
            x = int(self.reg_x_var.get())
            y = int(self.reg_y_var.get())
            larg = int(self.larg_var.get())
            alt = int(self.alt_var.get())
        except ValueError:
            messagebox.showerror("Erro", "Região inválida. X, Y, Largura e Altura devem ser números inteiros.")
            return
        if larg <= 0 or alt <= 0:
            messagebox.showerror("Erro", "Largura e altura devem ser maiores que zero.")
            return
        regiao = {"x": x, "y": y, "largura": larg, "altura": alt}
        pasta = ARQUIVO_CONFIG.parent / "referencias"
        try:
            pasta.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        caminho = str(pasta / "referencia_selecionada.json")
        try:
            caminho_json, caminho_png = _salvar_referencia_arquivo(regiao, caminho)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível capturar a referência.\n{e}")
            return
        self.referencia_var.set(caminho_json)
        msg = f"Referência capturada!\n{caminho_json}"
        if caminho_png:
            msg += f"\nPNG para conferência:\n{caminho_png}"
        messagebox.showinfo("Referência", msg)

    def _acoes_alvo(self, alvo):
        return self.acoes if alvo == "sim" else self.acoes_nao

    def _tree_acoes_alvo(self, alvo):
        return self.tree_acoes if alvo == "sim" else self.tree_acoes_nao

    def _atualizar_acoes_tree(self, alvo="sim"):
        tree = self._tree_acoes_alvo(alvo)
        lista = self._acoes_alvo(alvo)
        for item in tree.get_children():
            tree.delete(item)
        for i, a in enumerate(lista, 1):
            if a["tipo"] == "Validar Tela":
                pos = f"{len(a.get('acoes', []))}+{len(a.get('acoes_nao_encontrado', []))} ações"
            elif a["tipo"] in ("Print", "Espera", "Pular Data"):
                pos = ""
            else:
                pos = f"({a.get('x', 0)}, {a.get('y', 0)})"
            tree.insert("", tk.END, values=(
                i, a["tipo"], pos, a.get("espera", 0), a.get("rotulo", "")))

    def _acao_menu_adicionar(self, alvo="sim"):
        menu = tk.Menu(self.janela, tearoff=0)
        for t in ("Clique", "Espera", "Texto", "Data", "Pular Data", "Registrar Ignorado", "Print", "Validar Tela"):
            menu.add_command(label=t, command=lambda t=t: self._acao_adicionar(t, alvo))
        try:
            menu.tk_popup(self.janela.winfo_pointerx(), self.janela.winfo_pointery())
        finally:
            menu.grab_release()

    def _acao_adicionar(self, tipo, alvo="sim"):
        novo = {
            "tipo": tipo,
            "x": 0,
            "y": 0,
            "coluna": "",
            "texto_fixo": "",
            "espera": 0.5,
            "rotulo": f"{tipo} {len(self._acoes_alvo(alvo)) + 1}",
            "tamanho_min": 0,
            "selecionar_tudo": False,
            "apagar": False,
            "formato_data": "%d/%m/%Y",
            "regiao": {"x": 0, "y": 0, "largura": 100, "altura": 100},
            "referencia": "",
            "limiar": 0.85,
            "tolerancia": 30,
            "timeout": 3.0,
            "acoes": [],
            "acoes_nao_encontrado": [],
            "arquivo_ignorados": "",
        }
        dialogo = JanelaEditarPasso(self.janela, novo, self.colunas)
        self.janela.wait_window(dialogo.janela)
        if dialogo.resultado is not None:
            self._acoes_alvo(alvo).append(dialogo.resultado)
            self._atualizar_acoes_tree(alvo)

    def _acao_selecionada(self, alvo="sim"):
        tree = self._tree_acoes_alvo(alvo)
        sel = tree.selection()
        if not sel:
            return None
        return tree.index(sel[0])

    def _acao_editar(self, alvo="sim"):
        lista = self._acoes_alvo(alvo)
        idx = self._acao_selecionada(alvo)
        if idx is None:
            messagebox.showwarning("Ações", "Selecione uma ação para editar.")
            return
        copia = lista[idx].copy()
        dialogo = JanelaEditarPasso(self.janela, copia, self.colunas)
        self.janela.wait_window(dialogo.janela)
        if dialogo.resultado is not None:
            lista[idx] = dialogo.resultado
            self._atualizar_acoes_tree(alvo)

    def _acao_remover(self, alvo="sim"):
        lista = self._acoes_alvo(alvo)
        idx = self._acao_selecionada(alvo)
        if idx is None:
            return
        del lista[idx]
        self._atualizar_acoes_tree(alvo)

    def _acao_mover(self, dir_, alvo="sim"):
        lista = self._acoes_alvo(alvo)
        tree = self._tree_acoes_alvo(alvo)
        idx = self._acao_selecionada(alvo)
        if idx is None:
            return
        novo_idx = idx + dir_
        if novo_idx < 0 or novo_idx >= len(lista):
            return
        lista[idx], lista[novo_idx] = lista[novo_idx], lista[idx]
        self._atualizar_acoes_tree(alvo)
        items = tree.get_children()
        tree.selection_set(items[novo_idx])

    def _reorganizar(self, event=None):
        self.lbl_coluna.grid_remove()
        self.combo_coluna.grid_remove()
        self.lbl_texto.grid_remove()
        self.entry_texto.grid_remove()
        self.lbl_formato.grid_remove()
        self.entry_formato.grid_remove()
        self.lbl_tamanho.grid_remove()
        self.entry_tamanho.grid_remove()
        self.chk_selecionar.grid_remove()
        self.chk_apagar.grid_remove()
        self.lbl_monitor.grid_remove()
        self.combo_monitor.grid_remove()
        self.lbl_pasta_print.grid_remove()
        self.frame_pasta_print.grid_remove()
        self.lbl_coluna_nome.grid_remove()
        self.combo_coluna_nome.grid_remove()
        self.lbl_arquivo_ignorados.grid_remove()
        self.frame_arquivo_ignorados.grid_remove()
        self.frame_validacao.grid_remove()
        self.lbl_hint_pular.grid_remove()
        self.lbl_pos_x.grid()
        self.lbl_pos_y.grid()
        self.entry_x.grid()
        self.entry_y.grid()

        tipo = self.tipo_var.get()
        if tipo == "Validar Tela":
            self.lbl_pos_x.grid_remove()
            self.lbl_pos_y.grid_remove()
            self.entry_x.grid_remove()
            self.entry_y.grid_remove()
            self.frame_validacao.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=4)
            self.janela.geometry("700x740")
        elif tipo == "Pular Data":
            self.lbl_pos_x.grid_remove()
            self.lbl_pos_y.grid_remove()
            self.entry_x.grid_remove()
            self.entry_y.grid_remove()
            self.lbl_hint_pular.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=4)
            self.janela.geometry("660x480")
        elif tipo == "Registrar Ignorado":
            self.lbl_pos_x.grid_remove()
            self.lbl_pos_y.grid_remove()
            self.entry_x.grid_remove()
            self.entry_y.grid_remove()
            self.lbl_arquivo_ignorados.grid(row=1, column=0, sticky=tk.W, pady=4)
            self.frame_arquivo_ignorados.grid(row=1, column=1, sticky=tk.W, pady=4)
            self.lbl_coluna.grid(row=2, column=0, sticky=tk.W, pady=4)
            self.combo_coluna.grid(row=2, column=1, sticky=tk.W, pady=4)
            self.lbl_formato.grid(row=3, column=0, sticky=tk.W, pady=4)
            self.entry_formato.grid(row=3, column=1, sticky=tk.W, pady=4)
            self.janela.geometry("660x480")
        else:
            self.janela.geometry("660x480")
        if tipo == "Planilha":
            self.lbl_coluna.grid(row=3, column=0, sticky=tk.W, pady=4)
            self.combo_coluna.grid(row=3, column=1, sticky=tk.W, pady=4)
            self.lbl_tamanho.grid(row=4, column=0, sticky=tk.W, pady=4)
            self.entry_tamanho.grid(row=4, column=1, sticky=tk.W, pady=4)
            self.chk_selecionar.grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=2)
            self.chk_apagar.grid(row=5, column=1, sticky=tk.W, pady=2)
        elif tipo == "Texto":
            self.lbl_texto.grid(row=3, column=0, sticky=tk.W, pady=4)
            self.entry_texto.grid(row=3, column=1, sticky=tk.W, pady=4)
            self.chk_selecionar.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=2)
            self.chk_apagar.grid(row=4, column=1, sticky=tk.W, pady=2)
        elif tipo == "Data":
            self.lbl_formato.grid(row=3, column=0, sticky=tk.W, pady=4)
            self.entry_formato.grid(row=3, column=1, sticky=tk.W, pady=4)
            self.chk_selecionar.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=2)
            self.chk_apagar.grid(row=4, column=1, sticky=tk.W, pady=2)
        elif tipo == "Print":
            self.lbl_monitor.grid(row=3, column=0, sticky=tk.W, pady=4)
            self.combo_monitor.grid(row=3, column=1, sticky=tk.W, pady=4)
            self.lbl_pasta_print.grid(row=4, column=0, sticky=tk.W, pady=4)
            self.frame_pasta_print.grid(row=4, column=1, sticky=tk.W, pady=4)
            self.lbl_coluna_nome.grid(row=5, column=0, sticky=tk.W, pady=4)
            self.combo_coluna_nome.grid(row=5, column=1, sticky=tk.W, pady=4)
        if tipo == "Espera":
            self.espera_var.set("1.0")

    def _salvar(self):
        tipo = self.tipo_var.get()
        try:
            espera = float(self.espera_var.get())
        except ValueError:
            messagebox.showerror("Erro", "Espera deve ser um número.")
            return

        regiao = None
        referencia = ""
        limiar = 0.85
        tolerancia = 30
        timeout = 3.0
        x = y = 0

        if tipo == "Validar Tela":
            try:
                x = int(self.reg_x_var.get())
                y = int(self.reg_y_var.get())
                larg = int(self.larg_var.get())
                alt = int(self.alt_var.get())
            except ValueError:
                messagebox.showerror("Erro", "Região inválida. X, Y, Largura e Altura devem ser números inteiros.")
                return
            if larg <= 0 or alt <= 0:
                messagebox.showerror("Erro", "Largura e altura devem ser maiores que zero.")
                return
            regiao = {"x": x, "y": y, "largura": larg, "altura": alt}
            try:
                limiar = float(self.limiar_var.get()) / 100.0
                tolerancia = int(self.tol_var.get())
                timeout = float(self.timeout_var.get())
            except ValueError:
                messagebox.showerror("Erro", "Semelhança, Tolerância e Timeout devem ser números.")
                return
            if not (0.0 <= limiar <= 1.0):
                messagebox.showerror("Erro", "Semelhança mínima deve estar entre 0 e 100%.")
                return
            referencia = self.referencia_var.get().strip()
            if not referencia:
                messagebox.showerror("Erro", "Capture uma referência da região (botão Capturar Referência).")
                return
            if not os.path.exists(referencia):
                messagebox.showerror("Erro", "Arquivo de referência não encontrado:\n" + referencia)
                return
        else:
            try:
                x = int(self.x_var.get())
                y = int(self.y_var.get())
            except ValueError:
                messagebox.showerror("Erro", "X e Y devem ser números inteiros.")
                return

        try:
            tamanho_min = int(self.tamanho_var.get())
        except ValueError:
            tamanho_min = 0

        formato_data = ""
        if tipo == "Data":
            formato_data = self.formato_data_var.get().strip() or "%d/%m/%Y"
        arquivo_ignorados = ""
        if tipo == "Registrar Ignorado":
            arquivo_ignorados = self.arquivo_ignorados_var.get().strip()
            if not arquivo_ignorados:
                messagebox.showerror("Erro", "Escolha o arquivo de texto para registrar os ignorados.")
                return

        monitor_idx = 1
        pasta_print = ""
        coluna_nome = ""
        if tipo == "Print":
            sel = self.combo_monitor.current()
            monitor_idx = sel + 1 if sel >= 0 else 1
            pasta_print = self.pasta_print_var.get().strip()
            coluna_nome = self.coluna_nome_var.get().strip()
            if not pasta_print:
                messagebox.showerror("Erro", "Selecione uma pasta para salvar o print.")
                return

        self.resultado = {
            "tipo": tipo,
            "x": x,
            "y": y,
            "coluna": self.coluna_var.get() if tipo in ("Planilha", "Registrar Ignorado") else "",
            "texto_fixo": self.texto_fixo_var.get() if tipo == "Texto" else "",
            "espera": espera,
            "rotulo": self.rotulo_var.get(),
            "tamanho_min": tamanho_min if tipo == "Planilha" else 0,
            "selecionar_tudo": self.selecionar_var.get() if tipo in ("Planilha", "Texto", "Data") else False,
            "apagar": self.apagar_var.get() if tipo in ("Planilha", "Texto", "Data") else False,
            "monitor": monitor_idx,
            "pasta_print": pasta_print,
            "coluna_nome": coluna_nome,
            "formato_data": formato_data,
            "regiao": regiao,
            "referencia": referencia,
            "limiar": limiar,
            "tolerancia": tolerancia,
            "timeout": timeout,
            "acoes": list(self.acoes) if tipo == "Validar Tela" else [],
            "acoes_nao_encontrado": list(self.acoes_nao) if tipo == "Validar Tela" else [],
            "arquivo_ignorados": arquivo_ignorados,
        }
        self.janela.destroy()

    def _cancelar(self):
        self.resultado = None
        self.janela.destroy()


class JanelaSelecionarRegiao:
    """Overlay em tela cheia para o usuário arrastar e definir uma região."""
    def __init__(self, master, callback):
        self.callback = callback
        user32 = ctypes.windll.user32
        self.x0 = user32.GetSystemMetrics(76)   # SM_XVIRTUALSCREEN
        self.y0 = user32.GetSystemMetrics(77)   # SM_YVIRTUALSCREEN
        w = user32.GetSystemMetrics(78)         # SM_CXVIRTUALSCREEN
        h = user32.GetSystemMetrics(79)         # SM_CYVIRTUALSCREEN

        self.janela = tk.Toplevel(master)
        self.janela.overrideredirect(True)
        self.janela.geometry(f"{w}x{h}+{self.x0}+{self.y0}")
        self.janela.attributes("-alpha", 0.35)
        self.janela.attributes("-topmost", True)

        self.canvas = tk.Canvas(self.janela, bg="#1a1a2e", highlightthickness=0, cursor="cross")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.create_text(
            w // 2, h // 2,
            text="Arraste para selecionar a região onde a tela aparece\nEsc = cancelar",
            fill="white", font=("Segoe UI", 14, "bold"), justify=tk.CENTER)

        self.inicio_xy = None
        self.ret_id = None
        self.canvas.bind("<ButtonPress-1>", self._inicio)
        self.canvas.bind("<B1-Motion>", self._mover)
        self.canvas.bind("<ButtonRelease-1>", self._fim)
        self.janela.bind("<Escape>", lambda e: self._cancelar())
        self.janela.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.janela.grab_set()
        self.janela.focus_force()

    def _inicio(self, e):
        self.inicio_xy = (e.x, e.y)

    def _mover(self, e):
        if not self.inicio_xy:
            return
        if self.ret_id is not None:
            self.canvas.delete(self.ret_id)
        self.ret_id = self.canvas.create_rectangle(
            self.inicio_xy[0], self.inicio_xy[1], e.x, e.y,
            outline="#00ff00", width=2, fill="")

    def _fim(self, e):
        if not self.inicio_xy:
            return
        x1, y1 = self.inicio_xy
        x2, y2 = e.x, e.y
        vx = min(x1, x2) + self.x0
        vy = min(y1, y2) + self.y0
        vw = abs(x2 - x1)
        vh = abs(y2 - y1)
        try:
            self.janela.grab_release()
        except Exception:
            pass
        self.janela.destroy()
        if vw >= 5 and vh >= 5:
            self.callback((vx, vy, vw, vh))
        else:
            self.callback(None)

    def _cancelar(self):
        try:
            self.janela.grab_release()
        except Exception:
            pass
        self.janela.destroy()
        self.callback(None)


# ═══════════════════════════════════════════════════════════
#  Main
# ═══════════════════════════════════════════════════════════

if __name__ == "__main__":
    root = tk.Tk()
    app = AppAutomacao(root)
    root.mainloop()
