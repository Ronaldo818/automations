"""
=========================================================
INTERFACE DA AUTOMAÇÃO
=========================================================
"""

import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
from datetime import datetime

from config import TITULO, VERSAO


class Interface:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(TITULO)
        self.root.geometry("900x680")
        self.root.resizable(False, False)
        self.parar = False

        # Variáveis de contagem
        self.total_sucesso = 0
        self.total_erro = 0

        # Título
        tk.Label(
            self.root, text=f"{TITULO} - v{VERSAO}", font=("Segoe UI", 14, "bold")
        ).pack(pady=10)

        # ==============================================
        # DASHBOARD DE INFORMAÇÕES (FRAME SUPERIOR)
        # ==============================================
        frame_info = tk.Frame(self.root)
        frame_info.pack(fill="x", padx=15, pady=5)

        # Coluna Esquerda
        col_esq = tk.Frame(frame_info)
        col_esq.pack(side="left", fill="both", expand=True)
        
        self.lbl_status = tk.Label(col_esq, text="Status: Aguardando início", anchor="w", font=("Segoe UI", 10, "bold"), fg="blue")
        self.lbl_status.pack(fill="x")
        self.lbl_contrato = tk.Label(col_esq, text="Contrato Excel: -", anchor="w")
        self.lbl_contrato.pack(fill="x")
        self.lbl_item = tk.Label(col_esq, text="Material/Item: -", anchor="w")
        self.lbl_item.pack(fill="x")
        self.lbl_sap = tk.Label(col_esq, text="Contrato Gerado SAP: -", anchor="w")
        self.lbl_sap.pack(fill="x")

        # Coluna Direita (Contadores de Auditoria)
        col_dir = tk.Frame(frame_info)
        col_dir.pack(side="right", fill="both", expand=True)
        
        self.lbl_sucesso = tk.Label(col_dir, text="Sucessos: 0", anchor="e", font=("Segoe UI", 10, "bold"), fg="green")
        self.lbl_sucesso.pack(fill="x")
        self.lbl_erro = tk.Label(col_dir, text="Erros: 0", anchor="e", font=("Segoe UI", 10, "bold"), fg="red")
        self.lbl_erro.pack(fill="x")

        # ==============================================
        # PROGRESSO
        # ==============================================
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=850, mode="determinate")
        self.progress.pack(pady=10)

        # ==============================================
        # LOG
        # ==============================================
        self.log = ScrolledText(self.root, width=120, height=22, font=("Consolas", 9), bg="#F8F9FA")
        self.log.pack(padx=15)

        # ==============================================
        # BOTÃO
        # ==============================================
        tk.Button(self.root, text="Encerrar Segurança", width=20, command=self.encerrar, bg="#DC3545", fg="white", font=("Segoe UI", 9, "bold")).pack(pady=10)
        
        # Captura o clique no "X" da janela
        self.root.protocol("WM_DELETE_WINDOW", self.encerrar)

    # =================================================
    # MÉTODOS DE ATUALIZAÇÃO SEGURA (THREAD-SAFE)
    # =================================================
    def atualizar_status(self, texto):
        if self.parar: return
        def _update():
            try: self.lbl_status.config(text=f"Status: {texto}")
            except tk.TclError: pass
        try: self.root.after(0, _update)
        except tk.TclError: pass

    def atualizar_contrato(self, contrato):
        if self.parar: return
        def _update():
            try: self.lbl_contrato.config(text=f"Contrato Excel: {contrato}")
            except tk.TclError: pass
        try: self.root.after(0, _update)
        except tk.TclError: pass

    def atualizar_item(self, item):
        if self.parar: return
        def _update():
            try: self.lbl_item.config(text=f"Material/Item: {item}")
            except tk.TclError: pass
        try: self.root.after(0, _update)
        except tk.TclError: pass

    def atualizar_sap(self, contrato):
        if self.parar: return
        def _update():
            try: self.lbl_sap.config(text=f"Contrato Gerado SAP: {contrato}")
            except tk.TclError: pass
        try: self.root.after(0, _update)
        except tk.TclError: pass

    def registrar_sucesso(self):
        if self.parar: return
        self.total_sucesso += 1
        def _update():
            try: self.lbl_sucesso.config(text=f"Sucessos: {self.total_sucesso}")
            except tk.TclError: pass
        try: self.root.after(0, _update)
        except tk.TclError: pass

    def registrar_erro(self):
        if self.parar: return
        self.total_erro += 1
        def _update():
            try: self.lbl_erro.config(text=f"Erros: {self.total_erro}")
            except tk.TclError: pass
        try: self.root.after(0, _update)
        except tk.TclError: pass

    def escrever(self, texto):
        if self.parar: return
        agora = datetime.now().strftime("%H:%M:%S")
        texto_formatado = f"[{agora}] {texto}\n"
        def _escrever_seguro():
            try:
                self.log.insert(tk.END, texto_formatado)
                self.log.see(tk.END)
            except tk.TclError:
                pass
        try: self.root.after(0, _escrever_seguro)
        except tk.TclError: pass

    def progresso(self, atual, total):
        if self.parar: return
        def _atualizar_progresso():
            try:
                self.progress["maximum"] = total
                self.progress["value"] = atual
            except tk.TclError:
                pass
        try: self.root.after(0, _atualizar_progresso)
        except tk.TclError: pass

    # =================================================
    # CONTROLES
    # =================================================
    def encerrar(self):
        self.parar = True
        self.root.destroy()

    def iniciar(self):
        self.root.mainloop()