import subprocess
import importlib.util
import sys

DEPENDENCIAS = [
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("xlrd", "xlrd"),
    ("PIL", "Pillow"),
    ("google.genai", "google-genai"),
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
import json
import datetime
import re
import math
import unicodedata
import ctypes
from pathlib import Path
from collections import defaultdict

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

CAMPOS = (
    ("bico", "Bico"),
    ("inicial", "Encerrante Inicial"),
    ("final", "Encerrante Final"),
    ("quantidade", "Quantidade Abastecida"),
    ("preco_unitario", "Preço Unitário"),
    ("data", "Data"),
)

CAMPOS_COMPARAVEIS = (
    ("final", "Encerrante Final"),
    ("quantidade", "Quantidade Abastecida"),
)

COLUNAS_DIVERGENCIAS = (
    ("bico", "Bico", 60, tk.W),
    ("origem", "Origem", 60, tk.CENTER),
    ("ini_a", "Enc. Inicial A", 105, tk.E),
    ("ini_b", "Enc. Inicial B", 105, tk.E),
    ("fin_a", "Enc. Final A", 105, tk.E),
    ("fin_b", "Enc. Final B", 105, tk.E),
    ("qtd_a", "Qtde A", 90, tk.E),
    ("qtd_b", "Qtde B", 90, tk.E),
    ("campos", "Campos Divergentes", 210, tk.W),
    ("linha_a", "Linha A", 65, tk.CENTER),
    ("linha_b", "Linha B", 65, tk.CENTER),
)

COLUNAS_IGUAIS = (
    ("bico", "Bico", 60, tk.W),
    ("origem", "Origem", 60, tk.CENTER),
    ("ini_a", "Enc. Inicial A", 105, tk.E),
    ("ini_b", "Enc. Inicial B", 105, tk.E),
    ("fin_a", "Enc. Final A", 105, tk.E),
    ("fin_b", "Enc. Final B", 105, tk.E),
    ("qtd_a", "Qtde A", 90, tk.E),
    ("qtd_b", "Qtde B", 90, tk.E),
    ("linha_a", "Linha A", 65, tk.CENTER),
    ("linha_b", "Linha B", 65, tk.CENTER),
)

COLUNAS_UNICOS = (
    ("bico", "Bico", 75, tk.W),
    ("origem", "Origem", 75, tk.CENTER),
    ("inicial", "Enc. Inicial", 115, tk.E),
    ("final", "Enc. Final", 115, tk.E),
    ("quantidade", "Quantidade", 105, tk.E),
    ("linha", "Linha", 65, tk.CENTER),
)

EPS = 1e-9

_FORMATOS_DATA_PLANILHA = (
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%d-%m-%Y",
    "%Y/%m/%d",
    "%d.%m.%Y",
)

_FORMATOS_DATA_FILTRO = (
    "%d/%m/%Y",
    "%Y-%m-%d",
    "%d-%m-%Y",
    "%d/%m/%y",
    "%Y/%m/%d",
)


def _parse_data_planilha(valor):
    if valor is None:
        return None
    if isinstance(valor, (datetime.datetime, datetime.date)):
        try:
            return valor.date() if isinstance(valor, datetime.datetime) else valor
        except Exception:
            return None
    if pd is not None and isinstance(valor, pd.Timestamp):
        try:
            return valor.date()
        except Exception:
            return None
    texto = _texto(valor)
    if not texto:
        return None
    # Excel serial
    try:
        numero = float(texto.replace(",", "."))
        if 30000 <= numero <= 60000:
            base = datetime.date(1899, 12, 30)
            return base + datetime.timedelta(days=int(numero))
    except (ValueError, OverflowError):
        pass
    # pandas fallback (tenta dayfirst e não-dayfirst)
    if pd is not None:
        for dayfirst in (True, False):
            try:
                ts = pd.to_datetime(texto, dayfirst=dayfirst, errors="raise")
                if pd.isna(ts):
                    continue
                return ts.date() if hasattr(ts, "date") else None
            except Exception:
                continue
    for fmt in _FORMATOS_DATA_PLANILHA:
        try:
            return datetime.datetime.strptime(texto, fmt).date()
        except ValueError:
            continue
    return None


def _parse_data_filtro(texto):
    texto = _texto(texto)
    if not texto:
        return None
    for fmt in _FORMATOS_DATA_FILTRO:
        try:
            return datetime.datetime.strptime(texto, fmt).date()
        except ValueError:
            continue
    # tenta pandas
    if pd is not None:
        try:
            ts = pd.to_datetime(texto, dayfirst=True, errors="raise")
            if not pd.isna(ts):
                return ts.date() if hasattr(ts, "date") else None
        except Exception:
            pass
    return None


def _cor_mutado():
    return "#9a9a9a" if TEMA != "dark" else "#8d8d8d"


def _texto(valor):
    if valor is None:
        return ""
    texto = str(valor).strip()
    if texto.lower() in ("nan", "none", "nat"):
        return ""
    return texto


def _normalizar_texto(texto):
    texto = unicodedata.normalize("NFKD", _texto(texto).lower())
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return texto.strip()


def _parse_numero(valor):
    texto = _texto(valor)
    if not texto:
        return None
    texto = texto.replace("\u00a0", "").replace(" ", "")
    texto = re.sub(r"[^0-9,.\-+]", "", texto)
    if not texto or texto in ("-", "+", ".", ","):
        return None
    negativo = texto.startswith("-")
    texto = texto.lstrip("+-")
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif texto.count(".") > 1:
        texto = texto.replace(".", "")
    try:
        numero = float(texto)
    except ValueError:
        return None
    if not math.isfinite(numero):
        return None
    return -numero if negativo else numero


def _normalizar_bico(valor):
    texto = _texto(valor)
    if not texto:
        return ""
    texto = re.sub(r"\s+", "", texto.upper())
    if re.fullmatch(r"[+-]?\d+(?:[.,]0+)?", texto):
        try:
            return str(int(float(texto.replace(",", "."))))
        except (ValueError, OverflowError):
            pass
    return re.sub(r"(?<!\d)0+(?=\d)", "", texto)


def _ordem_bico(bico):
    return [int(parte) if parte.isdigit() else parte
            for parte in re.split(r"(\d+)", str(bico))]


def _fmt_numero(valor):
    if valor is None:
        return ""
    if abs(valor) < 5e-10:
        valor = 0.0
    texto = f"{valor:.4f}".rstrip("0").rstrip(".")
    if texto in ("", "-0"):
        texto = "0"
    return texto.replace(".", ",")


def _truncar_numero(valor, casas):
    fator = 10 ** casas
    return math.trunc(valor * fator) / fator


def _arredondar_metade_acima(valor, casas):
    fator = 10 ** casas
    if valor >= 0:
        return math.floor(valor * fator + 0.5) / fator
    return math.ceil(valor * fator - 0.5) / fator


def _dentro_tolerancia(a, b, tolerancia):
    if a is None or b is None:
        return False
    if abs(a - b) <= tolerancia + EPS:
        return True
    # Encerrantes/quantidades podem estar arredondados ou truncados em uma das
    # planilhas (ex.: 17,115 vs 17,11). Considera iguais quando um dos valores
    # é o outro arredondado/truncado para alguma precisão decimal (>= 2 casas).
    if abs(a - b) > max(tolerancia, 0.0) + 0.01 + EPS:
        return False
    for casas in range(2, 7):
        fator = 10 ** casas
        for grosseiro, fino in ((a, b), (b, a)):
            if abs(grosseiro * fator - round(grosseiro * fator)) > EPS:
                continue
            for funcao in (_truncar_numero, _arredondar_metade_acima, round):
                try:
                    if abs(funcao(fino, casas) - grosseiro) <= EPS:
                        return True
                except (TypeError, ValueError, OverflowError):
                    pass
    return False


def _detectar_colunas(colunas):
    textos = [(c, _normalizar_texto(str(c))) for c in colunas]
    usados = set()

    def escolher(predicados):
        for predicado in predicados:
            for coluna, texto in textos:
                if coluna in usados:
                    continue
                if predicado(texto):
                    usados.add(coluna)
                    return str(coluna)
        return ""

    bico = escolher([
        lambda t: "bico" in t,
        lambda t: "nozzle" in t,
        lambda t: "pistola" in t,
        lambda t: "ponto de abastecimento" in t,
    ])
    inicial = escolher([
        lambda t: "inicial" in t and ("encerrante" in t or "leitura" in t or "enc" in t),
        lambda t: "encerrante" in t and ("ini" in t or "inicio" in t),
        lambda t: t == "inicial" or t.startswith("inicial ") or t.startswith("inicio"),
    ])
    final = escolher([
        lambda t: "final" in t and ("encerrante" in t or "leitura" in t or "enc" in t),
        lambda t: "encerrante" in t and ("fin" in t or "fim" in t),
        lambda t: t == "final" or t.startswith("final ") or t.startswith("fim"),
    ])
    quantidade = escolher([
        lambda t: "quantidade" in t,
        lambda t: "abastecid" in t,
        lambda t: "qtde" in t or "qtd" in t,
        lambda t: "litro" in t,
        lambda t: "volume" in t,
    ])
    data = escolher([
        lambda t: t == "data" or t == "date",
        lambda t: "data" in t and "fiscal" not in t and "hora" not in t,
        lambda t: "date" in t,
        lambda t: "dia" in t,
    ])
    return {
        "bico": bico,
        "inicial": inicial,
        "final": final,
        "quantidade": quantidade,
        "data": data,
    }


class CompararBicos:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparar Planilhas - Bicos e Encerrantes")
        self.root.geometry("1180x800")
        self.root.minsize(1000, 640)

        self.planilhas = {}
        for lado in ("A", "B", "C"):
            self.planilhas[lado] = {
                "df": None,
                "nome": "",
                "caminho": "",
                "vars": {},
                "combos": {},
                "lbl": None,
            }

        self._resultado = None
        self._retorno_comparacao = None
        self.var_tolerancia = tk.StringVar(value="0,01")
        self.var_modo = tk.StringVar(value="ab")
        self.var_filtro_data = tk.StringVar(value="todos")
        self.var_data_dia = tk.StringVar(value="")
        self.var_data_ini = tk.StringVar(value="")
        self.var_data_fim = tk.StringVar(value="")
        self._ultimo_filtro_data = None
        self.var_pergunta_ia = tk.StringVar(value="")
        self.var_fonte_ia = tk.StringVar(value="Tela (captura)")
        self._resposta_ia = None
        self._sessao_ia_id = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self._pasta_registros_ia = (
            Path(__file__).resolve().parent / "automacoes" / "Comparador de Bicos" / "registros_ia")
        try:
            self._pasta_registros_ia.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        self._arquivo_historico_ia = (
            self._pasta_registros_ia / f"historico_ia_{self._sessao_ia_id}.txt")
        self._ultimo_contador_ia = 0
        self._ultima_pergunta_ia = ""
        self._ultima_fonte_ia = ""
        self._ultimos_dados_ia = None

        self._construir_ui()
        tema.aplicar_tema(self.root, TEMA)

    def _construir_ui(self):
        frame_config = ttk.LabelFrame(self.root, text="Configuração da Comparação", padding=6)
        frame_config.pack(fill=tk.X, padx=8, pady=(8, 2))

        frame_modo = ttk.Frame(frame_config)
        frame_modo.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(frame_modo, text="Modo:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        ttk.Radiobutton(frame_modo, text="Comparar Planilha A com Planilha B",
                        variable=self.var_modo, value="ab",
                        command=self._modo_alterado).pack(side=tk.LEFT)
        ttk.Radiobutton(frame_modo,
                        text="Combinar Planilhas A + B e comparar com Planilha C",
                        variable=self.var_modo, value="abc",
                        command=self._modo_alterado).pack(side=tk.LEFT, padx=20)
        ttk.Label(frame_modo, foreground=_cor_mutado(),
                  text="No modo A + B vs C, a origem (A ou B) aparece no resultado.").pack(
                      side=tk.LEFT, padx=10)

        frame_cfg = ttk.Frame(frame_config)
        frame_cfg.pack(fill=tk.X, pady=2)
        ttk.Label(frame_cfg, text="Variação aceita para mais ou para menos:").pack(side=tk.LEFT)
        ttk.Entry(frame_cfg, textvariable=self.var_tolerancia, width=10).pack(side=tk.LEFT, padx=6)
        ttk.Label(frame_cfg, foreground=_cor_mutado(),
                  text="Ex.: 0,01 — cobre diferenças de arredondamento").pack(side=tk.LEFT, padx=8)

        frame_filtro_data = ttk.Frame(frame_config)
        frame_filtro_data.pack(fill=tk.X, pady=2)
        linha_fd = ttk.Frame(frame_filtro_data)
        linha_fd.pack(fill=tk.X)
        ttk.Radiobutton(linha_fd, text="Todos os dias", variable=self.var_filtro_data,
                        value="todos", command=self._filtro_data_alterado).pack(side=tk.LEFT)
        ttk.Radiobutton(linha_fd, text="Dia específico", variable=self.var_filtro_data,
                        value="dia", command=self._filtro_data_alterado).pack(side=tk.LEFT, padx=12)
        ttk.Radiobutton(linha_fd, text="Intervalo de dias", variable=self.var_filtro_data,
                        value="intervalo", command=self._filtro_data_alterado).pack(side=tk.LEFT, padx=12)
        ttk.Label(linha_fd, foreground=_cor_mutado(),
                  text="Usa a coluna Data de cada planilha.").pack(side=tk.LEFT, padx=10)

        self.frame_filtro_data_campos = ttk.Frame(frame_filtro_data)
        self.frame_filtro_data_campos.pack(fill=tk.X, pady=(4, 0))

        self.frame_fd_dia = ttk.Frame(self.frame_filtro_data_campos)
        ttk.Label(self.frame_fd_dia, text="Dia:").pack(side=tk.LEFT)
        ttk.Entry(self.frame_fd_dia, textvariable=self.var_data_dia, width=12).pack(
            side=tk.LEFT, padx=6)
        ttk.Button(self.frame_fd_dia, text="📅",
                   command=lambda: self._abrir_calendario(self.var_data_dia)).pack(
                       side=tk.LEFT, padx=(0, 4))
        ttk.Label(self.frame_fd_dia, foreground=_cor_mutado(),
                  text="(ex.: 15/01/2024)").pack(side=tk.LEFT, padx=4)

        self.frame_fd_intervalo = ttk.Frame(self.frame_filtro_data_campos)
        ttk.Label(self.frame_fd_intervalo, text="De:").pack(side=tk.LEFT)
        ttk.Entry(self.frame_fd_intervalo, textvariable=self.var_data_ini, width=12).pack(
            side=tk.LEFT, padx=6)
        ttk.Button(self.frame_fd_intervalo, text="📅",
                   command=lambda: self._abrir_calendario(self.var_data_ini)).pack(
                       side=tk.LEFT, padx=(0, 6))
        ttk.Label(self.frame_fd_intervalo, text="Até:").pack(side=tk.LEFT, padx=(8, 0))
        ttk.Entry(self.frame_fd_intervalo, textvariable=self.var_data_fim, width=12).pack(
            side=tk.LEFT, padx=6)
        ttk.Button(self.frame_fd_intervalo, text="📅",
                   command=lambda: self._abrir_calendario(self.var_data_fim)).pack(
                       side=tk.LEFT, padx=(0, 6))
        ttk.Label(self.frame_fd_intervalo, foreground=_cor_mutado(),
                  text="(ex.: 01/01/2024 a 31/01/2024)").pack(side=tk.LEFT, padx=4)

        self._filtro_data_alterado()

        frame_planilhas = ttk.LabelFrame(self.root, text="Planilhas", padding=10)
        frame_planilhas.pack(fill=tk.X, padx=8, pady=2)
        container_planilhas = ttk.Frame(frame_planilhas)
        container_planilhas.pack(fill=tk.X)

        frame_plan_esq = ttk.Frame(container_planilhas)
        frame_plan_esq.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 4))
        self.notebook_planilhas = ttk.Notebook(frame_plan_esq)
        self.notebook_planilhas.pack(fill=tk.BOTH, expand=True)
        for lado in ("A", "B", "C"):
            self._montar_lado_planilha(self.notebook_planilhas, lado)

        self.frame_ia_container = ttk.Frame(container_planilhas)
        self.frame_ia_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))

        frame_acao = ttk.Frame(self.root)
        frame_acao.pack(fill=tk.X, padx=8, pady=4)
        self.btn_comparar = ttk.Button(frame_acao, text="Comparar",
                                       command=self._comparar)
        self.btn_comparar.pack(side=tk.LEFT)
        self.lbl_indicador = ttk.Label(frame_acao, text="●", foreground=_cor_mutado(),
                                       font=("Segoe UI", 11))
        self.lbl_indicador.pack(side=tk.LEFT, padx=(10, 4))
        self.lbl_status = ttk.Label(frame_acao,
                                    text="Carregue as planilhas e selecione as colunas.",
                                    foreground=_cor_mutado(), font=("Segoe UI", 9, "bold"))
        self.lbl_status.pack(side=tk.LEFT)
        self.btn_exportar = ttk.Button(frame_acao, text="Exportar Resultado",
                                       command=self._exportar)
        self.btn_exportar.pack(side=tk.RIGHT)
        self.btn_exportar.state(["disabled"])
        self.btn_sql = ttk.Button(frame_acao, text="Gerar INSERTs Postgres",
                                   command=self._gerar_inserts_postgres)
        self.btn_sql.pack(side=tk.RIGHT, padx=(0, 6))
        self.btn_ia = ttk.Button(frame_acao, text="Assistente IA",
                                 command=lambda: self.entry_pergunta_ia.focus_set())
        self.btn_ia.pack(side=tk.LEFT, padx=(10, 0))

        frame_resumo = ttk.LabelFrame(self.root, text="Resumo", padding=4)
        frame_resumo.pack(fill=tk.X, padx=8, pady=2)
        self.metricas = {}
        self.rotulos_metricas = {}
        for chave, rotulo in (
            ("linhas_esq", "Linhas A"),
            ("linhas_dir", "Linhas B"),
            ("iguais", "Iguais"),
            ("divergencias", "Divergências"),
            ("so_esq", "Só em A"),
            ("so_dir", "Só em B"),
        ):
            celula = ttk.Frame(frame_resumo)
            celula.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=4)
            label_rotulo = ttk.Label(celula, text=rotulo, foreground=_cor_mutado(),
                                     font=("Segoe UI", 8))
            label_rotulo.pack(anchor="center")
            self.rotulos_metricas[chave] = label_rotulo
            self.metricas[chave] = ttk.Label(celula, text="0",
                                             font=("Segoe UI", 12, "bold"))
            self.metricas[chave].pack(anchor="center")

        frame_resultado = ttk.LabelFrame(self.root, text="Resultado da Comparação", padding=6)
        frame_resultado.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        self.notebook = ttk.Notebook(frame_resultado)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.f_div = self._aba("Divergências")
        self.f_soa = self._aba("Somente na Planilha A")
        self.f_sob = self._aba("Somente na Planilha B")
        self.f_iguais = self._aba("Iguais")

        self.tree_div = self._criar_tree(self.f_div, COLUNAS_DIVERGENCIAS)
        self.tree_soa = self._criar_tree(self.f_soa, COLUNAS_UNICOS)
        self.tree_sob = self._criar_tree(self.f_sob, COLUNAS_UNICOS)
        self.tree_iguais = self._criar_tree(self.f_iguais, COLUNAS_IGUAIS)

        frame_ia = ttk.LabelFrame(self.frame_ia_container, text="Assistente IA", padding=6)
        frame_ia.pack(fill=tk.BOTH, expand=True)

        linha_fonte = ttk.Frame(frame_ia)
        linha_fonte.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(linha_fonte, text="Fonte de dados:").pack(side=tk.LEFT)
        self.combo_fonte_ia = ttk.Combobox(
            linha_fonte, textvariable=self.var_fonte_ia, state="readonly", width=26,
            values=("Tela (captura)", "Aba Divergências", "Aba Iguais",
                    "Somente Planilha A", "Somente Planilha B",
                    "Todos os resultados"))
        self.combo_fonte_ia.pack(side=tk.LEFT, padx=6)

        linha_ia = ttk.Frame(frame_ia)
        linha_ia.pack(fill=tk.X)
        ttk.Label(linha_ia, text="Pergunta:").pack(side=tk.LEFT)
        self.entry_pergunta_ia = ttk.Entry(linha_ia, textvariable=self.var_pergunta_ia)
        self.entry_pergunta_ia.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        self.entry_pergunta_ia.bind("<Return>", lambda e: self._analisar_ia())
        self.btn_analisar_ia = ttk.Button(linha_ia, text="Analisar",
                                          command=self._analisar_ia)
        self.btn_analisar_ia.pack(side=tk.LEFT, padx=(0, 4))
        self.btn_config_ia = ttk.Button(linha_ia, text="Configurar chave",
                                       command=self._configurar_chave_ia)
        self.btn_config_ia.pack(side=tk.LEFT)

        self.lbl_status_ia = ttk.Label(frame_ia, text="Digite a pergunta e clique em Analisar tela. "
                                                     "A captura será feita da janela da automação.",
                                      foreground=_cor_mutado(), font=("Segoe UI", 9))
        self.lbl_status_ia.pack(anchor=tk.W, pady=(4, 2))

        self.txt_resposta_ia = scrolledtext.ScrolledText(frame_ia, height=8, state=tk.DISABLED,
                                                        font=("Consolas", 9))
        self.txt_resposta_ia.pack(fill=tk.BOTH, expand=True)

        frame_log = ttk.LabelFrame(self.root, text="Log", padding=4)
        frame_log.pack(fill=tk.BOTH, padx=8, pady=(0, 8))
        self.log = scrolledtext.ScrolledText(frame_log, height=3, state=tk.DISABLED,
                                             font=("Consolas", 9))
        self.log.pack(fill=tk.BOTH, expand=True)

        self._aplicar_rotulos_modo()

    def _montar_lado_planilha(self, parent, lado):
        estado = self.planilhas[lado]
        aba = ttk.Frame(parent)
        parent.add(aba, text=f"Planilha {lado}")
        frame = ttk.Frame(aba, padding=6)
        frame.pack(fill=tk.BOTH, expand=True)

        linha = ttk.Frame(frame)
        linha.pack(fill=tk.X)
        ttk.Button(linha, text=f"Carregar Planilha {lado}",
                   command=lambda l=lado: self._carregar_planilha(l)).pack(side=tk.LEFT)
        estado["lbl"] = ttk.Label(linha, text="Nenhuma", foreground=_cor_mutado())
        estado["lbl"].pack(side=tk.LEFT, padx=8)

        for campo, rotulo in CAMPOS:
            rotulo_ui = "Qtde Abastecimento" if campo == "quantidade" else rotulo
            linha_campo = ttk.Frame(frame)
            linha_campo.pack(fill=tk.X, pady=(4, 0))
            ttk.Label(linha_campo, text=f"Coluna {rotulo_ui}:", width=26).pack(side=tk.LEFT)
            var = tk.StringVar()
            combo = ttk.Combobox(linha_campo, textvariable=var, state="readonly", width=36)
            combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))
            estado["vars"][campo] = var
            estado["combos"][campo] = combo

    def _aba(self, titulo):
        aba = ttk.Frame(self.notebook)
        self.notebook.add(aba, text=titulo)
        return aba

    def _criar_tree(self, parent, colunas):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True)
        tree = ttk.Treeview(frame, columns=[c[0] for c in colunas],
                            show="headings", height=12)
        for chave, titulo, largura, alinhamento in colunas:
            tree.heading(chave, text=titulo)
            tree.column(chave, width=largura, minwidth=50,
                        anchor=alinhamento, stretch=False)
        scroll_y = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        scroll_x = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return tree

    def _log(self, mensagem):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {mensagem}\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)

    def _caminho_config_ia(self):
        pasta = Path(__file__).resolve().parent / "automacoes" / "Comparador de Bicos"
        try:
            pasta.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        return pasta / "gemini_config.json"

    def _carregar_chave_ia(self):
        caminho = self._caminho_config_ia()
        try:
            if caminho.exists():
                with open(caminho, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                return (dados.get("api_key") or "").strip()
        except Exception:
            return ""
        return ""

    def _salvar_chave_ia(self, chave):
        caminho = self._caminho_config_ia()
        try:
            caminho.parent.mkdir(parents=True, exist_ok=True)
            with open(caminho, "w", encoding="utf-8") as f:
                json.dump({"api_key": chave.strip()}, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            self._log(f"Erro ao salvar chave da IA: {e}")
            return False

    def _configurar_chave_ia(self):
        janela = tk.Toplevel(self.root)
        janela.title("Configurar chave API Gemini")
        janela.geometry("460x170")
        janela.transient(self.root)
        janela.grab_set()
        frame = ttk.Frame(janela, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="Chave API do Google Gemini:").pack(anchor="w")
        var = tk.StringVar(value=self._carregar_chave_ia())
        entry = ttk.Entry(frame, textvariable=var, width=60, show="*")
        entry.pack(fill=tk.X, pady=(4, 8))
        lbl = ttk.Label(frame, text="A chave fica salva em arquivo local (gemini_config.json).",
                        foreground=_cor_mutado(), font=("Segoe UI", 9))
        lbl.pack(anchor="w")

        def salvar():
            chave = var.get().strip()
            if not chave:
                messagebox.showwarning("Aviso", "Informe a chave API.")
                return
            if self._salvar_chave_ia(chave):
                messagebox.showinfo("Salvo", "Chave API salva.")
                janela.destroy()
            else:
                messagebox.showerror("Erro", "Não foi possível salvar a chave.")

        def cancelar():
            janela.destroy()

        frame_btn = ttk.Frame(frame)
        frame_btn.pack(pady=8)
        ttk.Button(frame_btn, text="Salvar", command=salvar).pack(side=tk.LEFT, padx=4)
        ttk.Button(frame_btn, text="Cancelar", command=cancelar).pack(side=tk.LEFT, padx=4)
        janela.wait_window(janela)

    def _capturar_janela_ia(self):
        self.root.update_idletasks()
        left = self.root.winfo_rootx()
        top = self.root.winfo_rooty()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        if width <= 0 or height <= 0:
            raise RuntimeError("Janela inválida para captura.")
        from PIL import ImageGrab
        return ImageGrab.grab(bbox=(left, top, left + width, top + height))

    def _definir_resposta_ia(self, texto):
        self.txt_resposta_ia.config(state=tk.NORMAL)
        self.txt_resposta_ia.delete("1.0", tk.END)
        self.txt_resposta_ia.insert(tk.END, texto)
        self.txt_resposta_ia.config(state=tk.DISABLED)

    def _adicionar_resposta_ia(self, texto):
        self.txt_resposta_ia.config(state=tk.NORMAL)
        self.txt_resposta_ia.insert(tk.END, "\n" + texto + "\n")
        self.txt_resposta_ia.see(tk.END)
        self.txt_resposta_ia.config(state=tk.DISABLED)

    def _analisar_ia(self):
        pergunta = self.var_pergunta_ia.get().strip()
        if not pergunta:
            messagebox.showwarning("Aviso", "Digite a pergunta para a IA.")
            return
        chave = self._carregar_chave_ia()
        if not chave:
            if messagebox.askyesno("Chave API", "Chave API do Gemini não configurada. Deseja configurar agora?"):
                self._configurar_chave_ia()
                chave = self._carregar_chave_ia()
            if not chave:
                return

        fonte = self.var_fonte_ia.get()
        self._ultima_pergunta_ia = pergunta
        self._ultima_fonte_ia = fonte
        self._ultimo_contador_ia += 1
        self.btn_analisar_ia.state(["disabled"])
        self.lbl_status_ia.configure(text="Consultando Gemini...", foreground="#ff9800")
        self._definir_resposta_ia("Aguardando resposta...")
        self._ia_finalizada = False

        imagem = None
        dados_texto = None
        self._ultimos_dados_ia = None
        if fonte == "Tela (captura)":
            try:
                imagem = self._capturar_janela_ia()
            except Exception as e:
                self.btn_analisar_ia.state(["!disabled"])
                self.lbl_status_ia.configure(text="Erro na captura.", foreground="#e53935")
                self._definir_resposta_ia(f"Erro: {e}")
                self._log(f"IA: erro na captura - {e}")
                self._ia_finalizada = True
                return
        else:
            dados_texto = self._dados_fonte_ia(fonte)
            self._ultimos_dados_ia = dados_texto
            if not dados_texto:
                self.btn_analisar_ia.state(["!disabled"])
                self.lbl_status_ia.configure(text="Sem dados para enviar.", foreground="#e53935")
                self._definir_resposta_ia(
                    "Não há resultado carregado para esta fonte. Execute a comparação primeiro.")
                self._ia_finalizada = True
                return

        def tarefa():
            try:
                resposta = self._chamar_gemini(chave, imagem, dados_texto, pergunta)
                self.root.after(0, lambda: self._finalizar_analise_ia(resposta, None))
            except Exception as e:
                self.root.after(0, lambda: self._finalizar_analise_ia(None, str(e)))

        threading.Thread(target=tarefa, daemon=True).start()

        def vigia():
            time.sleep(90)
            if not self._ia_finalizada:
                self.root.after(0, lambda: self._finalizar_analise_ia(
                    None, "Tempo esgotado (60s) aguardando resposta do Gemini."))

        threading.Thread(target=vigia, daemon=True).start()

    def _dados_fonte_ia(self, fonte):
        resultado = getattr(self, "_resultado", None)
        if not resultado:
            return ""
        if fonte == "Aba Divergências":
            itens = resultado.get("divergencias", [])
            cabecalho = ["Bico", "Origem", "Enc.Inicial A", "Enc.Inicial B",
                         "Enc.Final A", "Enc.Final B", "Qtde A", "Qtde B",
                         "Campos Divergentes", "Linha A", "Linha B"]
            linhas = [",".join(cabecalho)]
            for it in itens:
                linhas.append(",".join(str(it.get(c, "")) for c in (
                    "bico", "origem", "inicial_a", "inicial_b", "final_a",
                    "final_b", "qtd_a", "qtd_b", "campos", "linha_a", "linha_b")))
            return "\n".join(linhas)
        if fonte == "Aba Iguais":
            itens = resultado.get("iguais", [])
            cabecalho = ["Bico", "Origem", "Enc.Inicial A", "Enc.Inicial B",
                         "Enc.Final A", "Enc.Final B", "Qtde A", "Qtde B",
                         "Linha A", "Linha B"]
            linhas = [",".join(cabecalho)]
            for it in itens:
                linhas.append(",".join(str(it.get(c, "")) for c in (
                    "bico", "origem", "inicial_a", "inicial_b", "final_a",
                    "final_b", "qtd_a", "qtd_b", "linha_a", "linha_b")))
            return "\n".join(linhas)
        if fonte == "Somente Planilha A":
            itens = resultado.get("so_esq", [])
            cabecalho = ["Bico", "Origem", "Enc.Inicial", "Enc.Final", "Quantidade", "Linha"]
            linhas = [",".join(cabecalho)]
            for it in itens:
                linhas.append(",".join(str(it.get(c, "")) for c in (
                    "bico", "origem", "inicial", "final", "quantidade", "linha")))
            return "\n".join(linhas)
        if fonte == "Somente Planilha B":
            itens = resultado.get("so_dir", [])
            cabecalho = ["Bico", "Origem", "Enc.Inicial", "Enc.Final", "Quantidade", "Linha"]
            linhas = [",".join(cabecalho)]
            for it in itens:
                linhas.append(",".join(str(it.get(c, "")) for c in (
                    "bico", "origem", "inicial", "final", "quantidade", "linha")))
            return "\n".join(linhas)
        if fonte == "Todos os resultados":
            partes = []
            for nome, chave in (("Divergências", "divergencias"),
                                ("Iguais", "iguais"),
                                ("Somente Planilha A", "so_esq"),
                                ("Somente Planilha B", "so_dir")):
                itens = resultado.get(chave, [])
                partes.append(f"### {nome} ({len(itens)} registros)")
                # reutiliza formatação simples
                if itens:
                    chaves = list(itens[0].keys())
                    partes.append(",".join(chaves))
                    for it in itens:
                        partes.append(",".join(str(it.get(c, "")) for c in chaves))
            return "\n".join(partes)
        return ""

    def _chamar_gemini(self, chave, imagem=None, dados_texto=None, pergunta=""):
        import io
        import base64
        import json as _json
        import urllib.request
        import urllib.error
        import urllib.parse

        b64 = None
        if imagem is not None:
            buffer = io.BytesIO()
            imagem.save(buffer, format="PNG")
            b64 = base64.b64encode(buffer.getvalue()).decode("ascii")

        # Lista modelos disponíveis para esta chave
        self._log("IA: listando modelos disponíveis...")
        url_lista = ("https://generativelanguage.googleapis.com/v1beta/models"
                     "?key=" + urllib.parse.quote(chave, safe=""))
        try:
            with urllib.request.urlopen(url_lista, timeout=20) as resp:
                dados_lista = _json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            corpo = ""
            try:
                corpo = e.read().decode("utf-8", errors="replace")[:300]
            except Exception:
                pass
            raise RuntimeError(f"Falha ao listar modelos (HTTP {e.code}): {corpo or e.reason}")
        except Exception as e:
            raise RuntimeError(f"Falha ao listar modelos: {e}")

        modelos = []
        for m in dados_lista.get("models", []):
            nome = (m.get("name") or "").replace("models/", "")
            metodos = m.get("supportedGenerationMethods") or []
            if "generateContent" not in metodos:
                continue
            if "gemini" not in nome.lower():
                continue
            modelos.append(nome)
        # prioriza modelos flash
        modelos.sort(key=lambda n: (0 if "flash" in n.lower() and "lite" not in n.lower() else
                                 1 if "flash" in n.lower() else 2, n))
        if not modelos:
            raise RuntimeError("Nenhum modelo com generateContent disponível para esta chave.")
        self._log("IA: modelos encontrados: " + ", ".join(modelos[:10]))

        partes = []
        if b64 is not None:
            partes.append({"inline_data": {"mime_type": "image/png", "data": b64}})
        texto_prompt = ""
        if dados_texto:
            # limita tamanho para não exceder limites de token
            if len(dados_texto) > 12000:
                dados_texto = dados_texto[:12000] + "\n... (dados truncados)"
            texto_prompt += "Dados enviados (CSV):\n" + dados_texto + "\n\n"
        texto_prompt += "Pergunta: " + (pergunta or "")
        partes.append({"text": texto_prompt})

        payload_base = {
            "contents": [{
                "parts": partes
            }]
        }

        ultimo_erro = None
        erros = []
        for modelo in modelos[:6]:
            self._log(f"IA: tentando {modelo}...")
            endpoint = ("https://generativelanguage.googleapis.com/v1beta/models/"
                        + urllib.parse.quote(modelo, safe="") + ":generateContent?key="
                        + urllib.parse.quote(chave, safe=""))
            try:
                req = urllib.request.Request(
                    endpoint,
                    data=_json.dumps(payload_base).encode("utf-8"),
                    headers={"Content-Type": "application/json"},
                    method="POST",
                )
                with urllib.request.urlopen(req, timeout=25) as resp:
                    resultado = _json.loads(resp.read().decode("utf-8"))
                for cand in resultado.get("candidates", []):
                    for part in (cand.get("content") or {}).get("parts", []):
                        if "text" in part:
                            self._log(f"IA: sucesso com {modelo}.")
                            return part["text"]
                self._log(f"IA: {modelo} respondeu sem texto.")
                return str(resultado)
            except urllib.error.HTTPError as e:
                corpo = ""
                try:
                    corpo = e.read().decode("utf-8", errors="replace")[:300]
                except Exception:
                    pass
                self._log(f"IA: falha em {modelo} (HTTP {e.code}): {corpo[:120]}")
                erros.append(f"{modelo}: HTTP {e.code} {corpo[:80]}")
                ultimo_erro = RuntimeError(f"HTTP {e.code}: {corpo}")
                continue
            except Exception as e:
                self._log(f"IA: falha em {modelo}: {str(e)[:120]}")
                erros.append(f"{modelo}: {str(e)[:80]}")
                ultimo_erro = e
                continue

        detalhe = " | ".join(erros[:6])
        raise RuntimeError("Nenhum modelo funcionou. Tentativas: " + detalhe)

    def _finalizar_analise_ia(self, resposta, erro):
        if getattr(self, "_ia_finalizada", False):
            return
        self._ia_finalizada = True
        self.btn_analisar_ia.state(["!disabled"])
        pergunta = getattr(self, "_ultima_pergunta_ia", "")
        fonte = getattr(self, "_ultima_fonte_ia", "")
        dados_texto = getattr(self, "_ultimos_dados_ia", None)
        if erro:
            self.lbl_status_ia.configure(text="Erro na análise.", foreground="#e53935")
            self._adicionar_resposta_ia(f"Pergunta: {pergunta}\nErro: {erro}")
            self._log(f"IA: erro - {erro}")
            self.var_pergunta_ia.set(pergunta)
            self._registrar_ia(pergunta, fonte, dados_texto, None, erro)
            return
        self.lbl_status_ia.configure(text="Análise concluída.", foreground="#4caf50")
        self._adicionar_resposta_ia(f"Pergunta: {pergunta}\nResposta: {resposta}")
        self._log("IA: análise concluída.")
        self.var_pergunta_ia.set("")
        self._registrar_ia(pergunta, fonte, dados_texto, resposta, None)

    def _registrar_ia(self, pergunta, fonte, dados_texto, resposta, erro):
        import csv
        pasta = getattr(self, "_pasta_registros_ia", None)
        if pasta is None:
            return
        try:
            pasta.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        texto_resposta = erro if erro else (resposta or "")
        arquivo_hist = getattr(self, "_arquivo_historico_ia", None)
        if arquivo_hist is not None:
            try:
                with open(arquivo_hist, "a", encoding="utf-8") as f:
                    f.write(f"[{datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] "
                            f"Fonte: {fonte}\n")
                    f.write(f"Pergunta: {pergunta}\n")
                    f.write(f"Resposta: {texto_resposta}\n")
                    f.write("-" * 60 + "\n")
            except Exception as e:
                self._log(f"IA: erro ao salvar histórico: {e}")
        # CSV da parte analisada
        if dados_texto:
            contador = getattr(self, "_ultimo_contador_ia", 0)
            sessao = getattr(self, "_sessao_ia_id", "sessao")
            nome = f"dados_analisados_{sessao}_{contador}.csv"
            caminho = pasta / nome
            try:
                with open(caminho, "w", encoding="utf-8", newline="") as f:
                    f.write(dados_texto)
                self._log(f"IA: dados analisados salvos em {nome}")
            except Exception as e:
                self._log(f"IA: erro ao salvar dados: {e}")

    def _set_status(self, texto, cor=None):
        self.lbl_indicador.configure(foreground=cor or _cor_mutado())
        self.lbl_status.configure(text=texto)

    def _modo_atual(self):
        return "abc" if self.var_modo.get() == "abc" else "ab"

    def _planilhas_necessarias(self):
        if self._modo_atual() == "abc":
            return ("A", "B", "C")
        return ("A", "B")

    def _aplicar_rotulos_modo(self):
        modo = self._modo_atual()
        if modo == "abc":
            rotulos = {
                "linhas_esq": "Linhas A + B",
                "linhas_dir": "Linhas C",
                "so_esq": "Só em A + B",
                "so_dir": "Só em C",
            }
            titulo_soa = "Somente em A + B"
            titulo_sob = "Somente na Planilha C"
            enc_esq, enc_dir = "A + B", "C"
            linha_esq, linha_dir = "A/B", "C"
            titulo_origem = "Origem (A/B)"
        else:
            rotulos = {
                "linhas_esq": "Linhas A",
                "linhas_dir": "Linhas B",
                "so_esq": "Só em A",
                "so_dir": "Só em B",
            }
            titulo_soa = "Somente na Planilha A"
            titulo_sob = "Somente na Planilha B"
            enc_esq, enc_dir = "A", "B"
            linha_esq, linha_dir = "A", "B"
            titulo_origem = "Origem"

        for chave, texto in rotulos.items():
            self.rotulos_metricas[chave].configure(text=texto)

        self.notebook.tab(self.f_soa, text=titulo_soa)
        self.notebook.tab(self.f_sob, text=titulo_sob)

        for tree in (self.tree_div, self.tree_iguais):
            tree.heading("origem", text=titulo_origem)
            tree.heading("ini_a", text=f"Enc. Inicial {enc_esq}")
            tree.heading("ini_b", text=f"Enc. Inicial {enc_dir}")
            tree.heading("fin_a", text=f"Enc. Final {enc_esq}")
            tree.heading("fin_b", text=f"Enc. Final {enc_dir}")
            tree.heading("qtd_a", text=f"Qtde {enc_esq}")
            tree.heading("qtd_b", text=f"Qtde {enc_dir}")
            tree.heading("linha_a", text=f"Linha {linha_esq}")
            tree.heading("linha_b", text=f"Linha {linha_dir}")

    def _atualizar_status_prontidao(self):
        faltando = [lado for lado in self._planilhas_necessarias()
                    if self.planilhas[lado]["df"] is None]
        if faltando:
            self._set_status(
                "Carregue: " + ", ".join(f"Planilha {lado}" for lado in faltando) + ".",
                "#ff9800")
        else:
            self._set_status(
                "Planilhas necessárias carregadas. Confira as colunas e clique em Comparar.",
                "#ff9800")

    def _filtro_data_alterado(self):
        for frame in (self.frame_fd_dia, self.frame_fd_intervalo):
            frame.pack_forget()
        modo = self.var_filtro_data.get()
        if modo == "dia":
            self.frame_fd_dia.pack(fill=tk.X)
        elif modo == "intervalo":
            self.frame_fd_intervalo.pack(fill=tk.X)

    def _abrir_calendario(self, var_destino):
        import calendar
        janela = tk.Toplevel(self.root)
        janela.title("Selecionar data")
        janela.transient(self.root)
        janela.grab_set()
        janela.resizable(False, False)

        try:
            data_atual = _parse_data_filtro(var_destino.get()) or datetime.date.today()
        except Exception:
            data_atual = datetime.date.today()

        estado = {"ano": data_atual.year, "mes": data_atual.month}

        frame_topo = ttk.Frame(janela, padding=6)
        frame_topo.pack(fill=tk.X)

        def atualizar_titulo():
            janela.title(f"Selecionar data - {calendar.month_name[estado['mes']]} {estado['ano']}")

        def renderizar():
            for w in frame_dias.winfo_children():
                w.destroy()
            atualizar_titulo()
            # cabeçalho dias da semana
            for i, nome in enumerate(["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]):
                ttk.Label(frame_dias, text=nome, width=4, anchor="center",
                          foreground=_cor_mutado()).grid(row=0, column=i, padx=1, pady=1)
            cal = calendar.monthcalendar(estado["ano"], estado["mes"])
            for linha, semana in enumerate(cal, start=1):
                for coluna, dia in enumerate(semana):
                    if dia == 0:
                        continue
                    btn = ttk.Button(frame_dias, text=str(dia), width=4,
                                     command=lambda d=dia: selecionar(d))
                    btn.grid(row=linha, column=coluna, padx=1, pady=1)

        def selecionar(dia):
            data = datetime.date(estado["ano"], estado["mes"], dia)
            var_destino.set(data.strftime("%d/%m/%Y"))
            janela.destroy()

        def mes_anterior():
            if estado["mes"] == 1:
                estado["mes"] = 12
                estado["ano"] -= 1
            else:
                estado["mes"] -= 1
            renderizar()

        def mes_seguinte():
            if estado["mes"] == 12:
                estado["mes"] = 1
                estado["ano"] += 1
            else:
                estado["mes"] += 1
            renderizar()

        ttk.Button(frame_topo, text="◀", command=mes_anterior, width=3).pack(side=tk.LEFT)
        ttk.Label(frame_topo, text="", width=2).pack(side=tk.LEFT)
        ttk.Button(frame_topo, text="▶", command=mes_seguinte, width=3).pack(side=tk.RIGHT)

        frame_dias = ttk.Frame(janela, padding=6)
        frame_dias.pack()
        renderizar()

        janela.wait_window(janela)

    def _modo_alterado(self):
        self._limpar_resultado()
        self._atualizar_status_prontidao()
        self._log(f"Modo alterado para: {self._rotulo_modo()}.")

    def _rotulo_modo(self):
        if self._modo_atual() == "abc":
            return "combinar A + B e comparar com C"
        return "comparar A com B"

    def _ler_planilha(self, caminho):
        if caminho.lower().endswith(".csv"):
            ultimo_erro = None
            for codificacao in ("utf-8-sig", "latin-1"):
                try:
                    return pd.read_csv(caminho, dtype=str, sep=None,
                                       engine="python", encoding=codificacao)
                except UnicodeDecodeError as erro:
                    ultimo_erro = erro
            raise ultimo_erro
        return pd.read_excel(caminho, dtype=str)

    def _carregar_planilha(self, lado):
        caminho = filedialog.askopenfilename(
            title=f"Selecionar Planilha {lado}",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm"),
                       ("CSV", "*.csv"),
                       ("Todos", "*.*")])
        if not caminho:
            return
        try:
            df = self._ler_planilha(caminho)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler a planilha {lado}.\n{e}")
            self._log(f"Erro ao carregar Planilha {lado}: {e}")
            return

        estado = self.planilhas[lado]
        estado["df"] = df
        estado["nome"] = os.path.basename(caminho)
        estado["caminho"] = caminho

        colunas = [str(c) for c in df.columns]
        detectadas = _detectar_colunas(df.columns)
        detect_pu = ""
        for coluna in colunas:
            texto = _normalizar_texto(coluna)
            if ("unitario" in texto or "preco" in texto or texto == "pu"):
                detect_pu = coluna
                break
        for campo, combo in estado["combos"].items():
            combo["values"] = colunas
            if campo == "preco_unitario":
                estado["vars"][campo].set(detect_pu)
            else:
                estado["vars"][campo].set(detectadas.get(campo, ""))
        estado["lbl"].configure(text=f"{estado['nome']} ({len(df)} linhas)")
        self._limpar_resultado()
        self._log(f"Planilha {lado} carregada: {estado['nome']} com {len(df)} linhas.")
        achadas = [f"{rotulo}='{detectadas[campo]}'" for campo, rotulo in CAMPOS
                   if detectadas.get(campo)]
        if achadas:
            self._log(f"Colunas detectadas na Planilha {lado}: " + ", ".join(achadas))

        self._atualizar_status_prontidao()

    def _limpar_resultado(self):
        self._resultado = None
        for tree, aba in (
            (self.tree_div, self.f_div),
            (self.tree_soa, self.f_soa),
            (self.tree_sob, self.f_sob),
            (self.tree_iguais, self.f_iguais),
        ):
            tree.delete(*tree.get_children())
        self.notebook.tab(self.f_div, text="Divergências")
        self.notebook.tab(self.f_iguais, text="Iguais")
        for label in self.metricas.values():
            label.configure(text="0")
        self._aplicar_rotulos_modo()
        self.btn_exportar.state(["disabled"])

    def _resolver_coluna(self, df, texto):
        if not texto:
            return None
        if texto in df.columns:
            return texto
        for coluna in df.columns:
            if str(coluna) == str(texto):
                return coluna
        return None

    def _valor_linha(self, row, coluna):
        if coluna is None:
            return ""
        try:
            return row.get(coluna, "")
        except Exception:
            return ""

    def _preparar_registros(self, df, selecao, lado, filtro_data=None):
        col_bico = self._resolver_coluna(df, selecao.get("bico"))
        col_inicial = self._resolver_coluna(df, selecao.get("inicial"))
        col_final = self._resolver_coluna(df, selecao.get("final"))
        col_quantidade = self._resolver_coluna(df, selecao.get("quantidade"))
        col_data = self._resolver_coluna(df, selecao.get("data"))
 
        registros = []
        ignoradas = []
        sem_inicial = []
        fora_periodo = 0
        exemplos_data_invalida = []
        for posicao, (_, row) in enumerate(df.iterrows()):
            linha = posicao + 2
            bico_original = _texto(self._valor_linha(row, col_bico))
            bico = _normalizar_bico(bico_original)
            if not bico:
                ignoradas.append(linha)
                continue

            data_valor = None
            if filtro_data is not None:
                if col_data is None:
                    pass
                else:
                    bruto = self._valor_linha(row, col_data)
                    data_valor = _parse_data_planilha(bruto)
                    modo_f, inicio, fim = filtro_data
                    if data_valor is None:
                        if len(exemplos_data_invalida) < 5:
                            exemplos_data_invalida.append(_texto(bruto))
                        fora_periodo += 1
                        continue
                    if not (inicio <= data_valor <= fim):
                        fora_periodo += 1
                        continue

            registro = {
                "bico": bico_original or bico,
                "bico_norm": bico,
                "inicial": _parse_numero(self._valor_linha(row, col_inicial)),
                "final": _parse_numero(self._valor_linha(row, col_final)),
                "quantidade": _parse_numero(self._valor_linha(row, col_quantidade)),
                "linha": linha,
                "origem": lado,
                "data": data_valor,
            }
            if col_inicial is not None and registro["inicial"] is None:
                sem_inicial.append(linha)
            registros.append(registro)
        return registros, ignoradas, sem_inicial, fora_periodo, exemplos_data_invalida

    def _custo_secundario(self, a, b):
        custo = 0.0
        for campo in ("final", "quantidade"):
            va = a.get(campo)
            vb = b.get(campo)
            if va is not None and vb is not None:
                custo += abs(va - vb)
        return round(custo, 10)

    def _comparar_par(self, a, b, tolerancia, campos_comparados,
                      rotulo_esq="A", rotulo_dir="B"):
        campos = []
        detalhes = []
        for campo, rotulo in CAMPOS_COMPARAVEIS:
            if campo not in campos_comparados:
                continue
            va = a.get(campo)
            vb = b.get(campo)
            if va is None and vb is None:
                continue
            if va is None or vb is None:
                campos.append(rotulo)
                lado_ausente = rotulo_esq if va is None else rotulo_dir
                detalhes.append(f"{rotulo}: ausente em {lado_ausente}")
                continue
            if not _dentro_tolerancia(va, vb, tolerancia):
                campos.append(rotulo)
                detalhes.append(
                    f"{rotulo}: {rotulo_esq}={_fmt_numero(va)} / "
                    f"{rotulo_dir}={_fmt_numero(vb)} "
                    f"(diferença {_fmt_numero(va - vb)})")

        dif_final = None
        if a.get("final") is not None and b.get("final") is not None:
            dif_final = a["final"] - b["final"]
        dif_quantidade = None
        if a.get("quantidade") is not None and b.get("quantidade") is not None:
            dif_quantidade = a["quantidade"] - b["quantidade"]

        return {
            "bico": a["bico"] or b["bico"],
            "bico_norm": a["bico_norm"],
            "origem": a.get("origem", ""),
            "origem_dir": b.get("origem", ""),
            "inicial_a": a["inicial"],
            "inicial_b": b["inicial"],
            "final_a": a["final"],
            "final_b": b["final"],
            "qtd_a": a["quantidade"],
            "qtd_b": b["quantidade"],
            "dif_final": dif_final,
            "dif_quantidade": dif_quantidade,
            "campos": "; ".join(campos),
            "detalhes": "; ".join(detalhes),
            "linha_a": a["linha"],
            "linha_b": b["linha"],
        }

    def _item_unico(self, registro):
        return {
            "bico": registro["bico"],
            "bico_norm": registro["bico_norm"],
            "origem": registro.get("origem", ""),
            "inicial": registro["inicial"],
            "final": registro["final"],
            "quantidade": registro["quantidade"],
            "linha": registro["linha"],
        }

    def _logs_preparacao(self, nome, registros, ignoradas, sem_inicial, fora_periodo=0):
        logs = [f"{nome}: {len(registros)} registro(s) válido(s)."]
        if ignoradas:
            logs.append(f"{nome}: {len(ignoradas)} linha(s) sem bico ignorada(s): "
                        f"{ignoradas[:30]}")
        if sem_inicial:
            logs.append(f"{nome}: {len(sem_inicial)} linha(s) sem encerrante inicial: "
                        f"{sem_inicial[:30]}")
        if fora_periodo:
            logs.append(f"{nome}: {fora_periodo} linha(s) fora do filtro de data.")
        return logs

    def _executar_comparacao(self, registros_esq, registros_dir, tolerancia,
                             campos_comparados, casou_por_inicial,
                             rotulo_esq="A", rotulo_dir="B"):
        grupos_esq = defaultdict(list)
        grupos_dir = defaultdict(list)
        for registro in registros_esq:
            grupos_esq[registro["bico_norm"]].append(registro)
        for registro in registros_dir:
            grupos_dir[registro["bico_norm"]].append(registro)

        divergencias = []
        iguais = []
        so_esq = []
        so_dir = []

        for bico in sorted(set(grupos_esq) | set(grupos_dir), key=_ordem_bico):
            lista_esq = grupos_esq.get(bico, [])
            lista_dir = grupos_dir.get(bico, [])
            usados_dir = set()

            if not casou_por_inicial:
                quantidade_pares = min(len(lista_esq), len(lista_dir))
                for posicao in range(quantidade_pares):
                    item = self._comparar_par(
                        lista_esq[posicao], lista_dir[posicao],
                        tolerancia, campos_comparados, rotulo_esq, rotulo_dir)
                    if item["campos"]:
                        divergencias.append(item)
                    else:
                        iguais.append(item)
                for registro in lista_esq[quantidade_pares:]:
                    so_esq.append(self._item_unico(registro))
                for registro in lista_dir[quantidade_pares:]:
                    so_dir.append(self._item_unico(registro))
                continue

            for registro_esq in lista_esq:
                melhor = None
                if registro_esq["inicial"] is not None:
                    for indice_dir, registro_dir in enumerate(lista_dir):
                        if indice_dir in usados_dir or registro_dir["inicial"] is None:
                            continue
                        diferenca = abs(registro_esq["inicial"] - registro_dir["inicial"])
                        if not _dentro_tolerancia(
                                registro_esq["inicial"], registro_dir["inicial"],
                                tolerancia):
                            continue
                        chave = (diferenca,
                                 self._custo_secundario(registro_esq, registro_dir),
                                 indice_dir)
                        if melhor is None or chave < melhor[0]:
                            melhor = (chave, indice_dir)

                if melhor is None:
                    so_esq.append(self._item_unico(registro_esq))
                    continue

                indice_dir = melhor[1]
                usados_dir.add(indice_dir)
                registro_dir = lista_dir[indice_dir]
                item = self._comparar_par(
                    registro_esq, registro_dir, tolerancia,
                    campos_comparados, rotulo_esq, rotulo_dir)
                if item["campos"]:
                    divergencias.append(item)
                else:
                    iguais.append(item)

            for indice_dir, registro_dir in enumerate(lista_dir):
                if indice_dir not in usados_dir:
                    so_dir.append(self._item_unico(registro_dir))

        divergencias.sort(key=lambda i: (_ordem_bico(i["bico_norm"]), i["linha_a"]))
        iguais.sort(key=lambda i: (_ordem_bico(i["bico_norm"]), i["linha_a"]))
        so_esq.sort(key=lambda i: (_ordem_bico(i["bico_norm"]), i["linha"]))
        so_dir.sort(key=lambda i: (_ordem_bico(i["bico_norm"]), i["linha"]))
        return divergencias, iguais, so_esq, so_dir

    def _comparar_planilhas(self, df_a, df_b, sel_a, sel_b,
                             tolerancia, campos_comparados, casou_por_inicial=None,
                             filtro_data=None):
        if casou_por_inicial is None:
            casou_por_inicial = bool(sel_a.get("inicial") and sel_b.get("inicial"))
 
        registros_a, ignoradas_a, sem_inicial_a, fora_a, ex_a = self._preparar_registros(
            df_a, sel_a, "A", filtro_data)
        registros_b, ignoradas_b, sem_inicial_b, fora_b, ex_b = self._preparar_registros(
            df_b, sel_b, "B", filtro_data)

        modo_correspondencia = ("bico + encerrante inicial" if casou_por_inicial
                                else "bico + ordem de ocorrência")
        logs = self._logs_preparacao("Planilha A", registros_a,
                                      ignoradas_a, sem_inicial_a, fora_a)
        logs += self._logs_preparacao("Planilha B", registros_b,
                                       ignoradas_b, sem_inicial_b, fora_b)
        if ex_a:
            logs.append("Planilha A: datas não reconhecidas em algumas linhas. Exemplos: " + ", ".join(repr(x) for x in ex_a))
        if ex_b:
            logs.append("Planilha B: datas não reconhecidas em algumas linhas. Exemplos: " + ", ".join(repr(x) for x in ex_b))
        if filtro_data is not None and (not registros_a or not registros_b) and (ex_a or ex_b):
            raise ValueError(
                "Nenhuma linha válida após o filtro de data. "
                "Verifique o formato da coluna Data (ex.: dd/mm/aaaa).")
        logs.append(f"Correspondência: {modo_correspondencia}.")
        logs.append("Campos comparados: " + ", ".join(
            rotulo for campo, rotulo in CAMPOS_COMPARAVEIS
            if campo in campos_comparados
        ) + f". Tolerância: {_fmt_numero(tolerancia)}.")

        divergencias, iguais, so_esq, so_dir = self._executar_comparacao(
            registros_a, registros_b, tolerancia, campos_comparados,
            casou_por_inicial, "A", "B")

        resumo = {
            "linhas_esq": len(registros_a),
            "linhas_dir": len(registros_b),
            "iguais": len(iguais),
            "divergencias": len(divergencias),
            "so_esq": len(so_esq),
            "so_dir": len(so_dir),
            "ignoradas_esq": len(ignoradas_a),
            "ignoradas_dir": len(ignoradas_b),
            "sem_inicial_esq": len(sem_inicial_a),
            "sem_inicial_dir": len(sem_inicial_b),
            "linhas_a": len(registros_a),
            "linhas_b": len(registros_b),
        }
        logs.append(
            f"Resultado: {len(iguais)} igual(is), {len(divergencias)} divergência(s), "
            f"{len(so_esq)} só em A, {len(so_dir)} só em B.")

        return {
            "divergencias": divergencias,
            "iguais": iguais,
            "so_esq": so_esq,
            "so_dir": so_dir,
            "resumo": resumo,
            "logs": logs,
            "tolerancia": tolerancia,
            "campos_comparados": list(campos_comparados),
            "modo_correspondencia": modo_correspondencia,
            "modo": "ab",
            "rotulos": {
                "esq": "Planilha A",
                "dir": "Planilha B",
                "so_esq": "Somente na Planilha A",
                "so_dir": "Somente na Planilha B",
            },
        }

    def _comparar_combinado(self, df_a, df_b, df_c, sel_a, sel_b, sel_c,
                             tolerancia, campos_comparados, casou_por_inicial=None,
                             filtro_data=None):
        if casou_por_inicial is None:
            casou_por_inicial = bool(
                sel_a.get("inicial") and sel_b.get("inicial") and sel_c.get("inicial"))
 
        registros_a, ignoradas_a, sem_inicial_a, fora_a, ex_a = self._preparar_registros(
            df_a, sel_a, "A", filtro_data)
        registros_b, ignoradas_b, sem_inicial_b, fora_b, ex_b = self._preparar_registros(
            df_b, sel_b, "B", filtro_data)
        registros_c, ignoradas_c, sem_inicial_c, fora_c, ex_c = self._preparar_registros(
            df_c, sel_c, "C", filtro_data)
        registros_esq = registros_a + registros_b

        modo_correspondencia = ("bico + encerrante inicial" if casou_por_inicial
                                else "bico + ordem de ocorrência")
        logs = self._logs_preparacao("Planilha A", registros_a,
                                      ignoradas_a, sem_inicial_a, fora_a)
        logs += self._logs_preparacao("Planilha B", registros_b,
                                       ignoradas_b, sem_inicial_b, fora_b)
        logs += self._logs_preparacao("Planilha C", registros_c,
                                       ignoradas_c, sem_inicial_c, fora_c)
        if ex_a:
            logs.append("Planilha A: datas não reconhecidas. Exemplos: " + ", ".join(repr(x) for x in ex_a))
        if ex_b:
            logs.append("Planilha B: datas não reconhecidas. Exemplos: " + ", ".join(repr(x) for x in ex_b))
        if ex_c:
            logs.append("Planilha C: datas não reconhecidas. Exemplos: " + ", ".join(repr(x) for x in ex_c))
        if filtro_data is not None and (not registros_esq or not registros_c) and (ex_a or ex_b or ex_c):
            raise ValueError(
                "Nenhuma linha válida após o filtro de data. "
                "Verifique o formato da coluna Data (ex.: dd/mm/aaaa).")
        logs.append(f"Total combinado A + B: {len(registros_esq)} registro(s) válido(s).")
        logs.append(f"Correspondência: {modo_correspondencia}.")
        logs.append("Campos comparados: " + ", ".join(
            rotulo for campo, rotulo in CAMPOS_COMPARAVEIS
            if campo in campos_comparados
        ) + f". Tolerância: {_fmt_numero(tolerancia)}.")

        divergencias, iguais, so_esq, so_dir = self._executar_comparacao(
            registros_esq, registros_c, tolerancia, campos_comparados,
            casou_por_inicial, "A + B", "C")

        resumo = {
            "linhas_esq": len(registros_esq),
            "linhas_dir": len(registros_c),
            "iguais": len(iguais),
            "divergencias": len(divergencias),
            "so_esq": len(so_esq),
            "so_dir": len(so_dir),
            "ignoradas_esq": len(ignoradas_a) + len(ignoradas_b),
            "ignoradas_dir": len(ignoradas_c),
            "sem_inicial_esq": len(sem_inicial_a) + len(sem_inicial_b),
            "sem_inicial_dir": len(sem_inicial_c),
            "linhas_a": len(registros_a),
            "linhas_b": len(registros_b),
            "linhas_c": len(registros_c),
            "ignoradas_a": len(ignoradas_a),
            "ignoradas_b": len(ignoradas_b),
            "ignoradas_c": len(ignoradas_c),
            "sem_inicial_a": len(sem_inicial_a),
            "sem_inicial_b": len(sem_inicial_b),
            "sem_inicial_c": len(sem_inicial_c),
        }
        logs.append(
            f"Resultado: {len(iguais)} igual(is), {len(divergencias)} divergência(s), "
            f"{len(so_esq)} só em A + B, {len(so_dir)} só em C.")

        return {
            "divergencias": divergencias,
            "iguais": iguais,
            "so_esq": so_esq,
            "so_dir": so_dir,
            "resumo": resumo,
            "logs": logs,
            "tolerancia": tolerancia,
            "campos_comparados": list(campos_comparados),
            "modo_correspondencia": modo_correspondencia,
            "modo": "abc",
            "rotulos": {
                "esq": "A + B",
                "dir": "Planilha C",
                "so_esq": "Somente em A + B",
                "so_dir": "Somente na Planilha C",
            },
        }

    def _validar_colunas_com_dados(self, df, selecao, campos_obrigatorios, lado):
        problemas = []
        for campo in campos_obrigatorios:
            rotulo = dict(CAMPOS).get(campo, campo)
            coluna = self._resolver_coluna(df, selecao.get(campo))
            if coluna is None:
                problemas.append(f"Planilha {lado}: coluna de {rotulo} não encontrada")
                continue
            if campo == "bico":
                tem_dado = any(_normalizar_bico(v) for v in df[coluna].tolist())
            elif campo == "data":
                tem_dado = any(_parse_data_planilha(v) is not None
                               for v in df[coluna].tolist())
            else:
                tem_dado = any(_parse_numero(v) is not None for v in df[coluna].tolist())
            if not tem_dado:
                problemas.append(f"Planilha {lado}: coluna de {rotulo} está sem dados")
        return problemas

    def _comparar(self):
        modo = self._modo_atual()
        necessarias = self._planilhas_necessarias()
        faltando = [lado for lado in necessarias
                    if self.planilhas[lado]["df"] is None]
        if faltando:
            messagebox.showwarning(
                "Aviso",
                "Carregue primeiro: " + ", ".join(
                    f"Planilha {lado}" for lado in faltando) + ".")
            return

        selecoes = {
            lado: {campo: self.planilhas[lado]["vars"][campo].get() for campo, _ in CAMPOS}
            for lado in necessarias
        }

        for lado in necessarias:
            if not selecoes[lado].get("bico"):
                messagebox.showwarning(
                    "Aviso", f"Selecione a coluna de Bico na Planilha {lado}.")
                return

        if modo == "abc":
            casou_por_inicial = bool(
                selecoes["A"].get("inicial")
                and selecoes["B"].get("inicial")
                and selecoes["C"].get("inicial"))
            if not casou_por_inicial:
                self._log("Aviso: Encerrante Inicial não selecionado nas três planilhas. "
                          "A correspondência será feita por bico + ordem de ocorrência.")
                if any(selecoes[lado].get("inicial") for lado in necessarias):
                    self._log("O Encerrante Inicial selecionado em apenas parte das "
                              "planilhas não será usado na correspondência.")
            campos_comparados = tuple(
                campo for campo, _ in CAMPOS_COMPARAVEIS
                if selecoes["A"].get(campo)
                and selecoes["B"].get(campo)
                and selecoes["C"].get(campo)
            )
            if not campos_comparados:
                messagebox.showwarning(
                    "Aviso",
                    "Selecione o Encerrante Final e/ou a Quantidade Abastecida "
                    "nas três planilhas (A, B e C).")
                return
        else:
            casou_por_inicial = bool(
                selecoes["A"].get("inicial") and selecoes["B"].get("inicial"))
            if not casou_por_inicial:
                self._log("Aviso: Encerrante Inicial não selecionado nas duas planilhas. "
                          "A correspondência será feita por bico + ordem de ocorrência.")
                if selecoes["A"].get("inicial") or selecoes["B"].get("inicial"):
                    self._log("O Encerrante Inicial selecionado em apenas uma planilha "
                              "não será usado na correspondência nem na comparação.")
            campos_comparados = tuple(
                campo for campo, _ in CAMPOS_COMPARAVEIS
                if selecoes["A"].get(campo) and selecoes["B"].get(campo)
            )
            if not campos_comparados:
                messagebox.showwarning(
                    "Aviso",
                    "Selecione o Encerrante Final e/ou a Quantidade Abastecida "
                    "nas duas planilhas.")
                return

        for campo, rotulo in CAMPOS_COMPARAVEIS:
            lados_com_campo = [lado for lado in necessarias
                               if selecoes[lado].get(campo)]
            if lados_com_campo and len(lados_com_campo) < len(necessarias):
                self._log(f"Aviso: {rotulo} não está selecionado em todas as planilhas; "
                          "esse campo não será comparado.")

        # Filtro por data
        filtro_data = None
        modo_filtro = self.var_filtro_data.get()
        if modo_filtro != "todos":
            for lado in necessarias:
                if not selecoes[lado].get("data"):
                    messagebox.showwarning(
                        "Aviso",
                        f"Selecione a coluna de Data na Planilha {lado} "
                        f"para usar o filtro por data.")
                    return
            if modo_filtro == "dia":
                dia = _parse_data_filtro(self.var_data_dia.get())
                if dia is None:
                    messagebox.showwarning(
                        "Aviso", "Informe um dia válido (ex.: 15/01/2024).")
                    return
                filtro_data = ("dia", dia, dia)
            else:
                inicio = _parse_data_filtro(self.var_data_ini.get())
                fim = _parse_data_filtro(self.var_data_fim.get())
                if inicio is None or fim is None:
                    messagebox.showwarning(
                        "Aviso", "Informe datas válidas para o intervalo "
                        "(ex.: 01/01/2024 a 31/01/2024).")
                    return
                if fim < inicio:
                    messagebox.showwarning(
                        "Aviso", "A data final é anterior à data inicial.")
                    return
                filtro_data = ("intervalo", inicio, fim)

        tolerancia = _parse_numero(self.var_tolerancia.get())
        if tolerancia is None or tolerancia < 0:
            messagebox.showwarning(
                "Aviso", "Informe uma tolerância válida (ex.: 0,01).")
            return

        obrigatorios = ("bico",) + (("inicial",) if casou_por_inicial else ()) + campos_comparados
        if filtro_data is not None:
            obrigatorios = obrigatorios + ("data",)
        problemas = []
        for lado in necessarias:
            problemas.extend(self._validar_colunas_com_dados(
                self.planilhas[lado]["df"], selecoes[lado], obrigatorios, lado))
        if problemas:
            messagebox.showwarning(
                "Colunas sem dados",
                "Corrija a seleção de colunas antes de comparar:\n\n" + "\n".join(problemas))
            return

        self._limpar_resultado()
        self.btn_comparar.state(["disabled"])
        self._set_status("Comparando...", "#ff9800")
        self._log(f"Iniciando comparação ({self._rotulo_modo()}). "
                  f"Tolerância: {_fmt_numero(tolerancia)}.")

        if modo == "abc":
            config = (
                self.planilhas["A"]["df"],
                self.planilhas["B"]["df"],
                self.planilhas["C"]["df"],
                selecoes["A"],
                selecoes["B"],
                selecoes["C"],
                tolerancia,
                campos_comparados,
                casou_por_inicial,
                filtro_data,
            )
        else:
            config = (
                self.planilhas["A"]["df"],
                self.planilhas["B"]["df"],
                selecoes["A"],
                selecoes["B"],
                tolerancia,
                campos_comparados,
                casou_por_inicial,
                filtro_data,
            )

        self._ultimo_filtro_data = filtro_data
        self._retorno_comparacao = None
        thread = threading.Thread(target=self._rodar_comparacao,
                                  args=(config, modo), daemon=True)
        thread.start()
        self.root.after(30, self._acompanhar_comparacao)

    def _rodar_comparacao(self, config, modo):
        try:
            if modo == "abc":
                resultado = self._comparar_combinado(*config)
            else:
                resultado = self._comparar_planilhas(*config)
            self._retorno_comparacao = ("ok", resultado)
        except Exception as e:
            self._retorno_comparacao = ("erro", str(e))

    def _acompanhar_comparacao(self):
        retorno = self._retorno_comparacao
        if retorno is None:
            self.root.after(30, self._acompanhar_comparacao)
            return

        self._retorno_comparacao = None
        self.btn_comparar.state(["!disabled"])
        tipo, valor = retorno
        if tipo == "erro":
            self._set_status("Erro na comparação.", "#e53935")
            messagebox.showerror("Erro", f"Erro na comparação.\n{valor}")
            self._log(f"Erro: {valor}")
            return

        resultado = valor
        self._exibir_resultado(resultado)
        for linha in resultado["logs"]:
            self._log(linha)
        if self._resultado is not None:
            self.btn_exportar.state(["!disabled"])
        resumo = resultado["resumo"]
        self._set_status(
            f"Comparação concluída: {resumo['divergencias']} divergência(s), "
            f"{resumo['iguais']} igual(is).", "#4caf50")

    def _preencher_tree(self, tree, itens, tipo):
        tree.delete(*tree.get_children())
        for item in itens:
            if tipo == "par":
                valores = (
                    item["bico"],
                    item.get("origem", ""),
                    _fmt_numero(item["inicial_a"]),
                    _fmt_numero(item["inicial_b"]),
                    _fmt_numero(item["final_a"]),
                    _fmt_numero(item["final_b"]),
                    _fmt_numero(item["qtd_a"]),
                    _fmt_numero(item["qtd_b"]),
                    item["campos"],
                    item["linha_a"],
                    item["linha_b"],
                )
            elif tipo == "igual":
                valores = (
                    item["bico"],
                    item.get("origem", ""),
                    _fmt_numero(item["inicial_a"]),
                    _fmt_numero(item["inicial_b"]),
                    _fmt_numero(item["final_a"]),
                    _fmt_numero(item["final_b"]),
                    _fmt_numero(item["qtd_a"]),
                    _fmt_numero(item["qtd_b"]),
                    item["linha_a"],
                    item["linha_b"],
                )
            else:
                valores = (
                    item["bico"],
                    item.get("origem", ""),
                    _fmt_numero(item["inicial"]),
                    _fmt_numero(item["final"]),
                    _fmt_numero(item["quantidade"]),
                    item["linha"],
                )
            tree.insert("", "end", values=valores)

    def _exibir_resultado(self, resultado=None):
        if resultado is not None:
            self._resultado = resultado
        resultado = self._resultado
        if not resultado:
            return
        if resultado.get("modo") != self._modo_atual():
            self._resultado = None
            self.btn_exportar.state(["disabled"])
            self._log("Resultado descartado: o modo de comparação foi alterado "
                      "durante o processamento.")
            return

        self._preencher_tree(self.tree_div, resultado["divergencias"], "par")
        self._preencher_tree(self.tree_soa, resultado["so_esq"], "unico")
        self._preencher_tree(self.tree_sob, resultado["so_dir"], "unico")
        self._preencher_tree(self.tree_iguais, resultado["iguais"], "igual")

        resumo = resultado["resumo"]
        self.notebook.tab(self.f_div, text=f"Divergências ({resumo['divergencias']})")
        self.notebook.tab(
            self.f_soa,
            text=f"{resultado['rotulos']['so_esq']} ({resumo['so_esq']})")
        self.notebook.tab(
            self.f_sob,
            text=f"{resultado['rotulos']['so_dir']} ({resumo['so_dir']})")
        self.notebook.tab(self.f_iguais, text=f"Iguais ({resumo['iguais']})")

        for chave, label in self.metricas.items():
            label.configure(text=str(resumo.get(chave, 0)))

    def _df_pares(self, itens, com_campos, rotulo_esq="A", rotulo_dir="B",
                  rotulo_linha_esq="A", rotulo_linha_dir="B"):
        colunas = [
            ("Bico", "bico"),
            ("Origem", "origem"),
            (f"Encerrante Inicial {rotulo_esq}", "inicial_a"),
            (f"Encerrante Inicial {rotulo_dir}", "inicial_b"),
            (f"Encerrante Final {rotulo_esq}", "final_a"),
            (f"Encerrante Final {rotulo_dir}", "final_b"),
            (f"Quantidade {rotulo_esq}", "qtd_a"),
            (f"Quantidade {rotulo_dir}", "qtd_b"),
        ]
        if com_campos:
            colunas.append(("Campos Divergentes", "campos"))
            colunas.append(("Detalhes", "detalhes"))
        colunas.extend([
            (f"Diferença Final ({rotulo_esq} - {rotulo_dir})", "dif_final"),
            (f"Diferença Quantidade ({rotulo_esq} - {rotulo_dir})", "dif_quantidade"),
            (f"Linha {rotulo_linha_esq}", "linha_a"),
            (f"Linha {rotulo_linha_dir}", "linha_b"),
        ])
        dados = [
            {rotulo: item.get(chave) for rotulo, chave in colunas}
            for item in itens
        ]
        return pd.DataFrame(dados, columns=[rotulo for rotulo, _ in colunas])

    def _df_unicos(self, itens):
        colunas = [
            ("Bico", "bico"),
            ("Origem", "origem"),
            ("Encerrante Inicial", "inicial"),
            ("Encerrante Final", "final"),
            ("Quantidade Abastecida", "quantidade"),
            ("Linha", "linha"),
        ]
        dados = [
            {rotulo: item.get(chave) for rotulo, chave in colunas}
            for item in itens
        ]
        return pd.DataFrame(dados, columns=[rotulo for rotulo, _ in colunas])

    def _exportar(self):
        if not self._resultado:
            messagebox.showwarning("Aviso", "Nenhum resultado para exportar.")
            return

        resultado = self._resultado
        modo = resultado.get("modo", "ab")
        nome_a = Path(self.planilhas["A"]["nome"]).stem or "A"
        nome_b = Path(self.planilhas["B"]["nome"]).stem or "B"

        if modo == "abc":
            nome_c = Path(self.planilhas["C"]["nome"]).stem or "C"
            arquivo_padrao = f"comparacao_bicos_{nome_a}+{nome_b}_vs_{nome_c}.xlsx"
            rotulo_esq, rotulo_dir = "A + B", "C"
            rotulo_linha_esq, rotulo_linha_dir = "A/B", "C"
            aba_so_esq, aba_so_dir = "Somente A+B", "Somente C"
            titulo_so_esq = "Somente em A + B"
            titulo_so_dir = "Somente na Planilha C"
            descricao_modo = "A + B vs C"
        else:
            arquivo_padrao = f"comparacao_bicos_{nome_a}_vs_{nome_b}.xlsx"
            rotulo_esq, rotulo_dir = "A", "B"
            rotulo_linha_esq, rotulo_linha_dir = "A", "B"
            aba_so_esq, aba_so_dir = "Somente A", "Somente B"
            titulo_so_esq = "Somente na Planilha A"
            titulo_so_dir = "Somente na Planilha B"
            descricao_modo = "A vs B"

        caminho = filedialog.asksaveasfilename(
            title="Exportar Resultado",
            defaultextension=".xlsx",
            initialfile=arquivo_padrao,
            filetypes=[("Excel", "*.xlsx")])
        if not caminho:
            return

        try:
            df_div = self._df_pares(
                resultado["divergencias"], True,
                rotulo_esq, rotulo_dir, rotulo_linha_esq, rotulo_linha_dir)
            df_iguais = self._df_pares(
                resultado["iguais"], False,
                rotulo_esq, rotulo_dir, rotulo_linha_esq, rotulo_linha_dir)
            df_so_esq = self._df_unicos(resultado["so_esq"])
            df_so_dir = self._df_unicos(resultado["so_dir"])

            rotulos = [
                rotulo for campo, rotulo in CAMPOS_COMPARAVEIS
                if campo in resultado["campos_comparados"]
            ]
            resumo = resultado["resumo"]
            linhas_resumo = [
                ("Tolerância utilizada", resultado["tolerancia"]),
                ("Modo", descricao_modo),
                ("Correspondência", resultado.get("modo_correspondencia", "")),
                ("Campos comparados", ", ".join(rotulos)),
            ]
            if modo == "abc":
                linhas_resumo.extend([
                    ("Linhas válidas na Planilha A", resumo.get("linhas_a", 0)),
                    ("Linhas válidas na Planilha B", resumo.get("linhas_b", 0)),
                    ("Linhas válidas na Planilha C", resumo.get("linhas_c", 0)),
                    ("Total combinado A + B", resumo["linhas_esq"]),
                ])
            else:
                linhas_resumo.extend([
                    ("Linhas válidas na Planilha A", resumo.get("linhas_a", 0)),
                    ("Linhas válidas na Planilha B", resumo.get("linhas_b", 0)),
                ])
            linhas_resumo.extend([
                ("Iguais", resumo["iguais"]),
                ("Divergências", resumo["divergencias"]),
                (titulo_so_esq, resumo["so_esq"]),
                (titulo_so_dir, resumo["so_dir"]),
            ])
            if modo == "abc":
                linhas_resumo.extend([
                    ("Linhas sem bico na Planilha A", resumo.get("ignoradas_a", 0)),
                    ("Linhas sem bico na Planilha B", resumo.get("ignoradas_b", 0)),
                    ("Linhas sem bico na Planilha C", resumo.get("ignoradas_c", 0)),
                    ("Linhas sem encerrante inicial na Planilha A",
                     resumo.get("sem_inicial_a", 0)),
                    ("Linhas sem encerrante inicial na Planilha B",
                     resumo.get("sem_inicial_b", 0)),
                    ("Linhas sem encerrante inicial na Planilha C",
                     resumo.get("sem_inicial_c", 0)),
                ])
            else:
                linhas_resumo.extend([
                    ("Linhas sem bico na Planilha A", resumo.get("ignoradas_a", 0)),
                    ("Linhas sem bico na Planilha B", resumo.get("ignoradas_b", 0)),
                    ("Linhas sem encerrante inicial na Planilha A",
                     resumo.get("sem_inicial_a", 0)),
                    ("Linhas sem encerrante inicial na Planilha B",
                     resumo.get("sem_inicial_b", 0)),
                ])
            df_resumo = pd.DataFrame(linhas_resumo, columns=["Métrica", "Valor"])

            with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
                df_div.to_excel(writer, sheet_name="Divergencias", index=False)
                df_so_esq.to_excel(writer, sheet_name=aba_so_esq, index=False)
                df_so_dir.to_excel(writer, sheet_name=aba_so_dir, index=False)
                df_iguais.to_excel(writer, sheet_name="Iguais", index=False)
                df_resumo.to_excel(writer, sheet_name="Resumo", index=False)
                for planilha in writer.sheets.values():
                    for coluna in planilha.columns:
                        tamanho = max(
                            (len(str(c.value)) for c in coluna if c.value is not None),
                            default=0)
                        letra = coluna[0].column_letter
                        planilha.column_dimensions[letra].width = min(
                            max(tamanho + 2, 10), 55)

            messagebox.showinfo("Salvo", f"Resultado salvo em:\n{caminho}")
            self._log(f"Resultado exportado para {caminho}")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível exportar.\n{e}")
            self._log(f"Erro ao exportar: {e}")


    def _gerar_inserts_postgres(self):
        if not getattr(self, "_resultado", None):
            try:
                messagebox.showwarning("Aviso", "Nenhum resultado para gerar INSERTs.")
            except Exception:
                pass
            return

        resultado = self._resultado
        modo = resultado.get("modo", "ab")
        # Diálogo para escolher de qual resultado gerar os INSERTs
        escolha = {"valor": "todos"}
        janela_escolha = tk.Toplevel(self.root)
        janela_escolha.title("Gerar INSERTs - Escolher Resultado")
        janela_escolha.geometry("360x260")
        janela_escolha.transient(self.root)
        janela_escolha.grab_set()
        frame_e = ttk.Frame(janela_escolha, padding=12)
        frame_e.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame_e, text="Gerar INSERTs a partir de:").pack(anchor="w", pady=(0, 8))

        opcoes = [
            ("divergencias", f"Divergências ({resultado['resumo'].get('divergencias', 0)})"),
            ("iguais", f"Iguais ({resultado['resumo'].get('iguais', 0)})"),
            ("so_esq", f"{resultado['rotulos'].get('so_esq', 'Somente A')} ({resultado['resumo'].get('so_esq', 0)})"),
            ("so_dir", f"{resultado['rotulos'].get('so_dir', 'Somente B')} ({resultado['resumo'].get('so_dir', 0)})"),
            ("todos", "Todos os registros carregados"),
        ]
        var = tk.StringVar(value="todos")
        for valor, rotulo in opcoes:
            ttk.Radiobutton(frame_e, text=rotulo, variable=var, value=valor).pack(anchor="w")

        def confirmar():
            escolha["valor"] = var.get()
            janela_escolha.destroy()

        def cancelar():
            escolha["valor"] = None
            janela_escolha.destroy()

        frame_btn = ttk.Frame(frame_e)
        frame_btn.pack(pady=10)
        ttk.Button(frame_btn, text="Gerar", command=confirmar).pack(side=tk.LEFT, padx=4)
        ttk.Button(frame_btn, text="Cancelar", command=cancelar).pack(side=tk.LEFT, padx=4)

        self.root.wait_window(janela_escolha)
        if escolha["valor"] is None:
            return

        self._gerar_inserts_postgres_filtrado(escolha["valor"])
        return

    def _gerar_inserts_postgres_filtrado(self, escolha):
        if not getattr(self, "_resultado", None):
            try:
                messagebox.showwarning("Aviso", "Nenhum resultado para gerar INSERTs.")
            except Exception:
                pass
            return

        resultado = self._resultado
        modo = resultado.get("modo", "ab")
        try:
            necessarias = self._planilhas_necessarias()
        except Exception:
            necessarias = ("A", "B") if modo == "ab" else ("A", "B", "C")

        # Monta conjunto de (origem, linha) permitidos conforme escolha
        permitidos = None  # None = todos
        if escolha != "todos":
            permitidos = set()
            if escolha == "divergencias":
                for item in resultado.get("divergencias", []):
                    permitidos.add((item.get("origem", ""), item.get("linha_a")))
                    permitidos.add((item.get("origem_dir", ""), item.get("linha_b")))
            elif escolha == "iguais":
                for item in resultado.get("iguais", []):
                    permitidos.add((item.get("origem", ""), item.get("linha_a")))
                    permitidos.add((item.get("origem_dir", ""), item.get("linha_b")))
            elif escolha == "so_esq":
                for item in resultado.get("so_esq", []):
                    permitidos.add((item.get("origem", ""), item.get("linha")))
            elif escolha == "so_dir":
                for item in resultado.get("so_dir", []):
                    permitidos.add((item.get("origem", ""), item.get("linha")))

        linhas = []
        linhas.append("-- INSERTs gerados a partir da comparação de bicos/encerrantes")
        linhas.append("-- Revise os valores antes de executar no PostgreSQL.")
        linhas.append("")

        # Colunas adicionais detectadas por nome (para INSERT completo)
        def _coluna_extra(df, *palavras):
            for coluna in df.columns:
                texto = _normalizar_texto(str(coluna))
                if any(p in texto for p in palavras):
                    return str(coluna)
            return None

        # Coleta registros das planilhas carregadas (origem dos abastecimentos)
        registros_por_origem = {}
        planilhas = getattr(self, "planilhas", {})
        for lado in necessarias:
            estado = planilhas.get(lado, {}) if isinstance(planilhas, dict) else {}
            df = estado.get("df") if isinstance(estado, dict) else None
            if df is None:
                continue
            sel = {campo: estado.get("vars", {}).get(campo, tk.StringVar()).get()
                   for campo, _ in CAMPOS}
            filtro = getattr(self, "_ultimo_filtro_data", None)
            regs, _, _, _, _ = self._preparar_registros(df, sel, lado, filtro)
            # mapeia colunas extras
            col_data = None
            sel_data = sel.get("data", "")
            if sel_data and sel_data in [str(c) for c in df.columns]:
                col_data = sel_data
            if col_data is None:
                col_data = _coluna_extra(df, "data")
            col_hora = _coluna_extra(df, "hora")
            col_total = _coluna_extra(df, "total")
            col_litros = _coluna_extra(df, "litro", "volume")
            col_pu_sel = sel.get("preco_unitario", "")
            col_pu = None
            if col_pu_sel:
                try:
                    if col_pu_sel in [str(c) for c in df.columns]:
                        col_pu = col_pu_sel
                except Exception:
                    col_pu = None
            if col_pu is None:
                col_pu = (_coluna_extra(df, "unitario")
                          or _coluna_extra(df, "pu")
                          or _coluna_extra(df, "preco"))
            col_frentista = (_coluna_extra(df, "idfrentista")
                             or _coluna_extra(df, "identifid")
                             or _coluna_extra(df, "cartao")
                             or _coluna_extra(df, "frentista")
                             or _coluna_extra(df, "funcionario")
                             or _coluna_extra(df, "operador"))
            for reg in regs:
                row = None
                try:
                    # recupera a linha original do df pela posição
                    pos = reg["linha"] - 2
                    row = df.iloc[pos]
                except Exception:
                    row = None
                reg["_df"] = df
                reg["_extras"] = {
                    "data": col_data, "hora": col_hora, "total": col_total,
                    "litros": col_litros, "pu": col_pu, "frentista": col_frentista,
                }
                reg["_row"] = row
                if permitidos is not None:
                    chave = (reg.get("origem", lado), reg.get("linha"))
                    if chave not in permitidos:
                        continue
                registros_por_origem.setdefault(reg.get("origem", lado), []).append(reg)

        # Adiciona registros dos resultados (divergências, iguais, únicos) como referência
        todos = []
        for item in resultado.get("divergencias", []):
            todos.append(item)
        for item in resultado.get("iguais", []):
            todos.append(item)
        for item in resultado.get("so_esq", []):
            todos.append(item)
        for item in resultado.get("so_dir", []):
            todos.append(item)

        def _sql_str(v):
            if v is None:
                return "NULL"
            t = str(v).strip()
            if t == "" or t.lower() in ("nan", "none"):
                return "NULL"
            return "'" + t.replace("'", "''") + "'"

        def _sql_num(v):
            n = _parse_numero(v)
            if n is None:
                return "NULL"
            return str(n).replace(",", ".")

        for origem, regs in registros_por_origem.items():
            for reg in regs:
                extras = reg.get("_extras", {})
                row = reg.get("_row")
                bico_raw = reg.get("bico") or ""
                bico3 = bico_raw.strip()
                # tenta normalizar para 3 dígitos
                try:
                    bico3 = str(int(float(bico_raw.replace(",", ".")))).zfill(3)
                except (ValueError, OverflowError):
                    bico3 = bico_raw.strip().zfill(3) if bico_raw.strip().isdigit() else bico_raw.strip()
                bico3 = bico3 if len(bico3) >= 3 else bico3.zfill(3)

                def _val(col):
                    if col is None or row is None:
                        return None
                    try:
                        return row[col]
                    except Exception:
                        return None

                data_v = _val(extras.get("data"))
                hora_v = _val(extras.get("hora"))
                total_v = _val(extras.get("total"))
                litros_v = _val(extras.get("litros"))
                if litros_v is None:
                    litros_v = reg.get("quantidade")
                pu_v = _val(extras.get("pu"))
                frentista_v = _val(extras.get("frentista"))

                ei = reg.get("inicial")
                ef = reg.get("final")
                qtd = reg.get("quantidade")
                if litros_v is None:
                    litros_v = qtd

                ei_sql = "NULL" if ei is None else str(ei)
                ef_sql = "NULL" if ef is None else str(ef)
                qtd_sql = "NULL" if qtd is None else str(qtd)
                litros_sql = _sql_num(litros_v)
                total_sql = _sql_num(total_v)
                # pu (preço unitário): preserva formatação decimal original
                if isinstance(pu_v, str):
                    t = pu_v.strip().replace(" ", "").replace("R$", "").replace("r$", "")
                    if t == "" or t.lower() in ("nan", "none", "null"):
                        pu_sql = "NULL"
                    elif "," in t and "." in t:
                        if t.rfind(",") > t.rfind("."):
                            t = t.replace(".", "").replace(",", ".")
                        else:
                            t = t.replace(",", "")
                        pu_sql = t
                    elif "," in t:
                        pu_sql = t.replace(",", ".")
                    elif "." in t:
                        pu_sql = t
                    else:
                        # valor inteiro (ex.: 7590) -> formata como 7.590
                        if t.isdigit():
                            if len(t) > 3:
                                pu_sql = t[:-3] + "." + t[-3:]
                            elif len(t) == 3:
                                pu_sql = "0." + t
                            else:
                                pu_sql = t
                        else:
                            pu_sql = t
                else:
                    # número: se for inteiro grande, divide por 1000
                    try:
                        n = float(pu_v)
                        if n == int(n) and abs(n) >= 1000:
                            pu_sql = str(n / 1000.0)
                        else:
                            pu_sql = str(n)
                    except (ValueError, TypeError):
                        pu_sql = _sql_num(pu_v)
                # data apenas (sem hora)
                if data_v is not None:
                    t = str(data_v).strip()
                    if " " in t:
                        t = t.split(" ")[0]
                    data_v = t
                data_sql = _sql_str(data_v)
                hora_sql = _sql_str(hora_v)
                frentista_sql = _sql_str(frentista_v)

                linhas.append(
                    "INSERT INTO public.abastecimentos "
                    "(bomba, \"data\", hora, litros, total, imprimiu, tipo, cupom, "
                    "encerrante, docu, string_full, pu, tempo, codcb, codpdv, codpro, "
                    "idfrentista, idcliente, filial, flag_local, autorizacao, regcbc, "
                    "idabast, md5, ei, datafiscal, flagclassifica, qtdecbc, tanque, "
                    "aba_cd_id_descontinuidade, motivo_erro, turno, integraweb, "
                    "codprooriginal, abt_vl_preco_venda_a, abt_vl_preco_venda_b, "
                    "abt_vl_preco_venda_c, abt_fl_tabela_preco, encerrante_automacao, "
                    "abt_vl_preco_cadastro, abt_ds_tempo, abt_ds_cod_virgula, "
                    "abt_ds_uuid, abt_ts_abastecimento, abt_ds_referencia_premmia, "
                    "abt_cd_status_sem_parar, abt_ds_cartao_1, abt_ds_cartao_2, "
                    "abt_ds_cartao_preco, abt_cd_veiculo, controle_origem, "
                    "abt_cd_concentrador, abt_bl_venda_cancelada, abt_ds_referencia, "
                    "abt_ds_cod_virgula_valor_unitario, abt_ds_cod_virgula_valor_total, "
                    "abt_cd_mesa, abt_ds_placa, abt_bl_integracao_xodoapp, "
                    "abt_ds_referencia_shellbox, abt_bl_sou_tag, abt_cd_leitura_placa, "
                    "abt_cd_km, abt_ds_observacao, abt_cd_caixa) "
                    f"VALUES ('{bico3}', {data_sql}, {hora_sql}, {litros_sql}, "
                    f"{total_sql}, NULL, NULL, NULL, {ef_sql}, NULL, "
                    f"'(@{str(total_sql).replace('.', '')}{str(litros_sql).replace('.', '')})', "
                    f"{pu_sql}, '00:00:00', NULL, '001', '00001', NULL, NULL, 0, "
                    "NULL, NULL, NULL, NULL, NULL, "
                    f"{ei_sql}, {data_sql}, NULL, {qtd_sql}, '001', NULL, NULL, "
                    "'1º TURNO', 'S', '00001', " + pu_sql + ", " + pu_sql + ", 0.000, 'A', "
                    f"{ef_sql}, {pu_sql}, '0000', '3A', uuid_generate_v4(), "
                    f"now(), NULL, 1, {frentista_sql}, NULL, NULL, NULL, NULL, NULL, "
                    "false, NULL, NULL, NULL, 0, NULL, false, NULL, false, 0, 0, NULL, NULL);")
        linhas.append("")
        linhas.append(f"-- Total de registros: {sum(len(v) for v in registros_por_origem.values())}")

        texto = "\n".join(linhas)

        janela = tk.Toplevel(self.root)
        janela.title("INSERTs PostgreSQL - Abastecimentos")
        janela.geometry("900x600")
        frame = ttk.Frame(janela, padding=8)
        frame.pack(fill=tk.BOTH, expand=True)
        txt = scrolledtext.ScrolledText(frame, font=("Consolas", 9))
        txt.pack(fill=tk.BOTH, expand=True)
        txt.insert(tk.END, texto)
        txt.config(state=tk.NORMAL)

        def copiar():
            janela.clipboard_clear()
            janela.clipboard_append(texto)
            messagebox.showinfo("Copiado", "INSERTs copiados para a área de transferência.")

        ttk.Button(frame, text="Copiar", command=copiar).pack(pady=6)

if __name__ == "__main__":
    root = tk.Tk()
    app = CompararBicos(root)
    root.mainloop()
