import tkinter as tk
from tkinter import ttk
import time
import threading

def show_loading_screen(main_app_callback):
    # Cria a janela de loading
    splash = tk.Tk()
    splash.overrideredirect(True)  # Remove bordas da janela
    splash.geometry("400x200+500+300")  # Posição e tamanho da janela
    splash.configure(bg='white')

    # Texto de carregamento
    label = tk.Label(splash, text="Carregando, por favor aguarde...", font=("Helvetica", 16), bg='white')
    label.pack(pady=40)

    # Barra de progresso
    progress = ttk.Progressbar(splash, mode='indeterminate', length=300)
    progress.pack(pady=20)
    progress.start(10)  

    def load_and_open():
        time.sleep(2)  # Simula o tempo de carregamento
        splash.after(0, finish_loading)
    
    def finish_loading():
        splash.destroy()  # Fecha a janela de loading
        main_app_callback()

    threading.Thread(target=load_and_open, daemon=True).start()
    splash.mainloop()

    
            