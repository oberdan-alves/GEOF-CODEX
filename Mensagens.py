"""Caixas de mensagem compartilhadas com fonte legível em todos os módulos."""
import tkinter as tk
from tkinter import simpledialog

FONTE_MENSAGEM = ("Segoe UI", 16)
FONTE_BOTAO_MENSAGEM = ("Segoe UI", 14)
BOTOES = {
    "ok": (("OK", "ok"),),
    "okcancel": (("OK", "ok"), ("Cancelar", "cancel")),
    "yesno": (("Sim", "yes"), ("Não", "no")),
    "yesnocancel": (("Sim", "yes"), ("Não", "no"), ("Cancelar", "cancel")),
    "retrycancel": (("Tentar novamente", "retry"), ("Cancelar", "cancel")),
    "abortretryignore": (("Abortar", "abort"), ("Tentar novamente", "retry"), ("Ignorar", "ignore")),
}


class _Mensagem(simpledialog.Dialog):
    def __init__(self, parent, title, message, tipo, default=None, detail=None, icon=None):
        self.message = str(message or "")
        self.detail = detail
        self.icon = icon
        self.botoes = BOTOES[tipo]
        self.default = default or self.botoes[0][1]
        self.result = "cancel" if any(v == "cancel" for _, v in self.botoes) else self.botoes[-1][1]
        super().__init__(parent, title)

    def body(self, master):
        simbolo = {"info": "ℹ", "warning": "⚠", "error": "✖", "question": "?"}.get(self.icon, "")
        if simbolo:
            tk.Label(master, text=simbolo, font=("Segoe UI", 24)).pack(side="left", padx=12, anchor="n")
        texto = self.message + ("\n\n" + str(self.detail) if self.detail else "")
        tk.Label(master, text=texto, font=FONTE_MENSAGEM, justify="left",
                 wraplength=min(720, max(240, self.winfo_screenwidth() - 160))).pack(padx=16, pady=16)

    def buttonbox(self):
        frame = tk.Frame(self)
        for label, valor in self.botoes:
            botao = tk.Button(frame, text=label, font=FONTE_BOTAO_MENSAGEM,
                              command=lambda v=valor: self.escolher(v), padx=16, pady=6)
            botao.pack(side="left", padx=6, pady=12)
            if valor == self.default:
                self.initial_focus = botao
                botao.focus_set()
        frame.pack()
        self.bind("<Return>", lambda event: self.escolher(self.default))
        self.bind("<Escape>", self.cancel)

    def escolher(self, valor):
        self.result = valor
        self.cancel()


class _MessageBox:
    def _show(self, title, message, tipo, icon, **options):
        parent = options.pop("parent", None) or tk._default_root
        temporario = parent is None
        if temporario:
            parent = tk.Tk()
            parent.withdraw()
        try:
            return _Mensagem(parent, title, message, tipo, options.get("default"),
                             options.get("detail"), icon).result
        finally:
            if temporario:
                parent.destroy()

    def showinfo(self, title=None, message=None, **options):
        return self._show(title, message, "ok", "info", **options)

    def showwarning(self, title=None, message=None, **options):
        return self._show(title, message, "ok", "warning", **options)

    def showerror(self, title=None, message=None, **options):
        return self._show(title, message, "ok", "error", **options)

    def askyesno(self, title=None, message=None, **options):
        return self._show(title, message, "yesno", "question", **options) == "yes"

    def askyesnocancel(self, title=None, message=None, **options):
        resultado = self._show(title, message, "yesnocancel", "question", **options)
        return None if resultado == "cancel" else resultado == "yes"

    def askokcancel(self, title=None, message=None, **options):
        return self._show(title, message, "okcancel", "question", **options) == "ok"

    def askretrycancel(self, title=None, message=None, **options):
        return self._show(title, message, "retrycancel", "warning", **options) == "retry"

    def askquestion(self, title=None, message=None, **options):
        return self._show(title, message, "yesno", "question", **options)


messagebox = _MessageBox()
