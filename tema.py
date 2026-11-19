import json
import tkinter as tk
from tkinter import ttk
from pathlib import Path

ARQUIVO_TEMA = Path(__file__).parent / "tema.json"

PALETA = {
    "light": {
        "nome": "light",
        "bg": "#f0f0f0",
        "fg": "#1a1a1a",
        "panel": "#fafafa",
        "borda": "#c9c9c9",
        "selecao": "#0078d4",
        "selecao_fg": "#ffffff",
        "campo_bg": "#ffffff",
        "campo_fg": "#1a1a1a",
        "campo_sel": "#0078d4",
        "disabled_bg": "#dedede",
        "disabled_fg": "#9a9a9a",
        "logger_bg": "#ffffff",
        "logger_fg": "#1a1a1a",
        "destaque": "#005a9e",
    },
    "dark": {
        "nome": "dark",
        "bg": "#1e1e1e",
        "fg": "#e6e6e6",
        "panel": "#252526",
        "borda": "#3e3e42",
        "selecao": "#094771",
        "selecao_fg": "#ffffff",
        "campo_bg": "#2d2d30",
        "campo_fg": "#e6e6e6",
        "campo_sel": "#094771",
        "disabled_bg": "#3a3a3a",
        "disabled_fg": "#6a6a6a",
        "logger_bg": "#1e1e1e",
        "logger_fg": "#dcdcdc",
        "destaque": "#3794ff",
    },
}


def carregar_tema():
    try:
        with open(ARQUIVO_TEMA, "r", encoding="utf-8") as f:
            dados = json.load(f)
        return dados.get("tema", "light")
    except Exception:
        return "light"


def salvar_tema(tema):
    tema = "dark" if tema == "dark" else "light"
    try:
        with open(ARQUIVO_TEMA, "w", encoding="utf-8") as f:
            json.dump({"tema": tema}, f, indent=2)
    except Exception:
        pass
    return tema


def cores(tema):
    return PALETA.get("dark" if tema == "dark" else "light")


def aplicar_tema(root, tema):
    cores_t = cores(tema)
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    style.configure(".", background=cores_t["bg"], foreground=cores_t["fg"],
                    fieldbackground=cores_t["campo_bg"])
    style.map(".", background=[("disabled", cores_t["disabled_bg"])])

    style.configure("TFrame", background=cores_t["bg"])
    style.configure("TLabel", background=cores_t["bg"], foreground=cores_t["fg"])
    style.configure("TLabelframe", background=cores_t["bg"], bordercolor=cores_t["borda"],
                    foreground=cores_t["fg"])
    style.configure("TLabelframe.Label", background=cores_t["bg"], foreground=cores_t["fg"])
    style.configure("TButton",
                    background=cores_t["panel"],
                    foreground=cores_t["fg"],
                    bordercolor=cores_t["borda"],
                    lightcolor=cores_t["panel"],
                    darkcolor=cores_t["panel"],
                    focuscolor=cores_t["selecao"])
    style.map("TButton",
              background=[("active", cores_t["selecao"]),
                          ("pressed", cores_t["selecao"]),
                          ("disabled", cores_t["disabled_bg"])],
              foreground=[("active", cores_t["selecao_fg"]),
                          ("pressed", cores_t["selecao_fg"]),
                          ("disabled", cores_t["disabled_fg"])])

    style.configure("TEntry",
                    fieldbackground=cores_t["campo_bg"],
                    foreground=cores_t["campo_fg"],
                    insertcolor=cores_t["campo_fg"],
                    bordercolor=cores_t["borda"])
    style.map("TEntry",
              fieldbackground=[("disabled", cores_t["disabled_bg"])],
              foreground=[("disabled", cores_t["disabled_fg"])])

    style.configure("TSpinbox",
                    fieldbackground=cores_t["campo_bg"],
                    foreground=cores_t["campo_fg"],
                    insertcolor=cores_t["campo_fg"],
                    bordercolor=cores_t["borda"],
                    arrowcolor=cores_t["fg"])
    style.map("TSpinbox",
              fieldbackground=[("disabled", cores_t["disabled_bg"])],
              foreground=[("disabled", cores_t["disabled_fg"])])

    style.configure("TCombobox",
                    fieldbackground=cores_t["campo_bg"],
                    foreground=cores_t["campo_fg"],
                    background=cores_t["panel"],
                    bordercolor=cores_t["borda"],
                    arrowcolor=cores_t["fg"],
                    selectbackground=cores_t["selecao"],
                    selectforeground=cores_t["selecao_fg"])
    style.map("TCombobox",
              fieldbackground=[("readonly", cores_t["campo_bg"]),
                               ("disabled", cores_t["disabled_bg"])],
              foreground=[("readonly", cores_t["campo_fg"]),
                          ("disabled", cores_t["disabled_fg"])])

    style.configure("Treeview",
                    background=cores_t["campo_bg"],
                    fieldbackground=cores_t["campo_bg"],
                    foreground=cores_t["campo_fg"],
                    bordercolor=cores_t["borda"])
    style.map("Treeview",
              background=[("selected", cores_t["selecao"])],
              foreground=[("selected", cores_t["selecao_fg"])])
    style.configure("Treeview.Heading",
                    background=cores_t["panel"],
                    foreground=cores_t["fg"],
                    bordercolor=cores_t["borda"])
    style.map("Treeview.Heading", background=[("active", cores_t["selecao"])])

    style.configure("TCheckbutton", background=cores_t["bg"], foreground=cores_t["fg"])
    style.map("TCheckbutton",
              background=[("active", cores_t["bg"])],
              foreground=[("disabled", cores_t["disabled_fg"])])

    style.configure("TSeparator", background=cores_t["borda"])

    style.configure("TNotebook", background=cores_t["bg"], bordercolor=cores_t["borda"],
                    tabmargins=(4, 4, 4, 0))
    style.configure("TNotebook.Tab",
                    background=cores_t["panel"],
                    foreground=cores_t["fg"],
                    bordercolor=cores_t["borda"],
                    lightcolor=cores_t["panel"],
                    darkcolor=cores_t["panel"],
                    padding=(12, 5))
    style.map("TNotebook.Tab",
              background=[("selected", cores_t["selecao"]),
                          ("active", cores_t["selecao"])],
              foreground=[("selected", cores_t["selecao_fg"]),
                          ("active", cores_t["selecao_fg"])])

    style.configure("Vertical.TScrollbar",
                    background=cores_t["panel"],
                    troughcolor=cores_t["bg"],
                    bordercolor=cores_t["borda"],
                    arrowcolor=cores_t["fg"])
    style.map("Vertical.TScrollbar",
              background=[("active", cores_t["selecao"])])
    style.configure("Horizontal.TScrollbar",
                    background=cores_t["panel"],
                    troughcolor=cores_t["bg"],
                    bordercolor=cores_t["borda"],
                    arrowcolor=cores_t["fg"])
    style.map("Horizontal.TScrollbar",
              background=[("active", cores_t["selecao"])])

    root.configure(bg=cores_t["bg"])

    for w in _iterar(root):
        try:
            nome = str(w.winfo_class())
        except Exception:
            nome = ""
        if nome in ("Text",):
            w.configure(bg=cores_t["logger_bg"], fg=cores_t["logger_fg"],
                        insertbackground=cores_t["campo_fg"])
        elif nome in ("Canvas",):
            w.configure(bg=cores_t["bg"])

    return cores_t


def _iterar(w):
    filhos = [w]
    for child in w.winfo_children():
        filhos.append(child)
        filhos.extend(_iterar(child))
    return filhos
