import os
import sys
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk

from planilhar import run_planilhar_notas
from leitor import run_leitor_pdf_xml


def load_icon(name, size=(24,24)):
    for d in ('icons','assets/icons','resources/icons'):
        p = os.path.join(d, f"{name}.png")
        if os.path.exists(p):
            img = Image.open(p).resize(size)
            return ImageTk.PhotoImage(img)
    return None


def main_menu():
    root = tk.Tk()
    root.title("Selecione um Script")
    root.geometry("500x400")
    root.configure(bg="#f5f5f5")

    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('TButton', font=('Segoe UI',11), padding=10)
    style.map('TButton', background=[('active','#3a76d8')])

    frame = ttk.Frame(root, padding=20)
    frame.pack(expand=True, fill='both')

    ttk.Label(frame, text="Selecione um Script", font=('Segoe UI',18,'bold')).pack(pady=(0,10))

    btn_p = ttk.Button(frame, text="Planilhar Notas", command=run_planilhar_notas)
    btn_l = ttk.Button(frame, text="Leitor PDF/XML", command=run_leitor_pdf_xml)
    btn_p.pack(fill='x', pady=5)
    btn_l.pack(fill='x', pady=5)

    root.mainloop()