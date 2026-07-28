"""
=========================================================
INTERFACE DA AUTOMAÇÃO
=========================================================
"""

import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

from config import TITULO, VERSAO


class Interface:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(TITULO)
        self.root.geometry("900x650")
        self.root.resizable(False, False)
        self.parar = False

        # ==============================================
        # TÍTULO
        # ==============================================
        titulo = tk.Label(
            self.root,
            text=f"{TITULO} - v{VERSAO}",
            font=("Segoe UI", 14, "bold")
        )
        titulo.pack(pady=10)

        # ==============================================
        # STATUS
        # ==============================================
        self.lbl_status = tk.Label(
            self.root,
            text="Status: Aguardando início",
            anchor="w",
            font=("Segoe UI", 10)
        )
        self.lbl_status.pack(fill="x", padx=15)

        self.lbl_contrato = tk.Label(
            self.root,
            text="Contrato: -",
            anchor="w"
        )
        self.lbl_contrato.pack(fill="x", padx=15)

        self.lbl_item = tk.Label(
            self.root,
            text="Item: -",
            anchor="w"
        )
        self.lbl_item.pack(fill="x", padx=15)

        self.lbl_sap = tk.Label(
            self.root,
            text="Contrato SAP: -",
            anchor="w"
        )
        self.lbl_sap.pack(fill="x", padx=15)

        # ==============================================
        # PROGRESSO
        # ==============================================
        self.progress = ttk.Progressbar(
            self.root,
            orient="horizontal",
            length=850,
            mode="determinate"
        )
        self.progress.pack(pady=15)

        # ==============================================
        # LOG
        # ==============================================
        self.log = ScrolledText(
            self.root,
            width=120,
            height=22,
            font=("Consolas", 10)
        )
        self.log.pack(padx=15)

        # ==============================================
        # BOTÃO
        # ==============================================
        self.botao = tk.Button(
            self.root,
            text="Encerrar",
            width=20,
            command=self.encerrar
        )
        self.botao.pack(pady=10)

        # Intercepta o clique no 'X' da janela para encerrar corretamente
        self.root.protocol("WM_DELETE_WINDOW", self.encerrar)

    # =================================================
    # MÉTODOS DE ATUALIZAÇÃO (THREAD-SAFE)
    # =================================================
    # O uso do 'self.root.after(0, ...)' garante que a atualização
    # visual seja feita pela Thread principal do Tkinter, evitando crashes.
    
    def atualizar_status(self, texto):
        try:
            self.root.after(0, lambda: self.lbl_status.config(text=f"Status: {texto}"))
        except tk.TclError:
            pass # Ignora se a janela já foi fechada

    def atualizar_contrato(self, contrato):
        try:
            self.root.after(0, lambda: self.lbl_contrato.config(text=f"Contrato: {contrato}"))
        except tk.TclError:
            pass

    def atualizar_item(self, item):
        try:
            self.root.after(0, lambda: self.lbl_item.config(text=f"Item: {item}"))
        except tk.TclError:
            pass

    def atualizar_sap(self, contrato):
        try:
            self.root.after(0, lambda: self.lbl_sap.config(text=f"Contrato SAP: {contrato}"))
        except tk.TclError:
            pass

    def escrever(self, texto):
        def _escrever_seguro():
            try:
                self.log.insert(tk.END, texto + "\n")
                self.log.see(tk.END)
            except tk.TclError:
                pass
        self.root.after(0, _escrever_seguro)

    def progresso(self, atual, total):
        def _atualizar_progresso():
            try:
                self.progress["maximum"] = total
                self.progress["value"] = atual
            except tk.TclError:
                pass
        self.root.after(0, _atualizar_progresso)

    # =================================================
    # CONTROLES
    # =================================================
    def encerrar(self):
        """Sinaliza para a automação parar e destrói a interface."""
        self.parar = True
        self.root.destroy()

    def iniciar(self):
        """Inicia o loop principal da interface gráfica."""
        self.root.mainloop()