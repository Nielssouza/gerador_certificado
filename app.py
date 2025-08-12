# cert_generator.py
import os
import sqlite3
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, filedialog
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm


# ---------- Configs ----------
DB_FILE = "certificados.db"
OUT_DIR = "certificados_pdfs"
os.makedirs(OUT_DIR, exist_ok=True)


# ---------- Banco (SQLite) ----------
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS certificado (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        evento TEXT NOT NULL,
        data_emissao TEXT NOT NULL,
        numero TEXT
    )
    """)
    conn.commit()
    conn.close()

def salvar_registro(nome, evento, data_emissao, numero):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("INSERT INTO certificado (nome, evento, data_emissao, numero) VALUES (?, ?, ?, ?)",
                (nome, evento, data_emissao, numero))
    conn.commit()
    conn.close()


# ---------- Gerar PDF com ReportLab ----------
def gerar_certificado_pdf(nome, evento, data_emissao, numero, arquivo_saida):
    # landscape A4
    c = canvas.Canvas(arquivo_saida, pagesize=landscape(A4))
    largura, altura = landscape(A4)

    # Background opcional: se tiver um PNG/JPG para usar como template, descomente e ajuste o caminho
    # bg_path = "template.png"
    # if os.path.exists(bg_path):
    #     c.drawImage(bg_path, 0, 0, width=largura, height=altura)

    # Texto do certificado (exemplo simples)
    c.setFont("Helvetica-Bold", 30)
    c.drawCentredString(largura/2, altura - 60*mm, "CERTIFICADO")

    c.setFont("Helvetica", 18)
    texto = f"Certificamos que"
    c.drawCentredString(largura/2, altura - 80*mm, texto)

    c.setFont("Helvetica-Bold", 26)
    c.drawCentredString(largura/2, altura - 100*mm, nome)

    c.setFont("Helvetica", 16)
    c.drawCentredString(largura/2, altura - 120*mm, f"participou do evento / curso:")
    c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(largura/2, altura - 135*mm, evento)

    c.setFont("Helvetica", 14)
    c.drawString(30*mm, 20*mm, f"Data de emissão: {data_emissao}")
    c.drawString(30*mm, 15*mm, f"Nº certificado: {numero}")

    # Rodapé / assinatura simulada (pode desenhar uma linha para assinatura)
    c.line(largura - 90*mm, 25*mm, largura - 30*mm, 25*mm)
    c.setFont("Helvetica", 12)
    c.drawCentredString(largura - 60*mm, 18*mm, "Assinatura / Organização")

    c.showPage()
    c.save()




# ---------- Interface Tkinter ----------
def gerar_e_salvar():
    nome = entry_nome.get().strip()
    evento = entry_evento.get().strip()
    data_emissao = entry_data.get().strip()
    numero = entry_num.get().strip()

    if not nome or not evento or not data_emissao:
        messagebox.showwarning("Atenção", "Preencha Nome, Evento e Data.")
        return

    # Se o número não for informado, gera automaticamente no formato CERT-0001
    if not numero:
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("SELECT MAX(id) FROM certificado")
        ultimo_id = cur.fetchone()[0]
        conn.close()
        novo_id = (ultimo_id or 0) + 1
        numero = f"CERT-{novo_id:04d}"

    # Validação mínima da data
    try:
        datetime.strptime(data_emissao, "%Y-%m-%d")
    except Exception:
        try:
            dt = datetime.strptime(data_emissao, "%d/%m/%Y")
            data_emissao = dt.strftime("%Y-%m-%d")
            entry_data.delete(0, tk.END)
            entry_data.insert(0, data_emissao)
        except Exception:
            if not data_emissao:
                data_emissao = datetime.today().strftime("%Y-%m-%d")
                entry_data.delete(0, tk.END)
                entry_data.insert(0, data_emissao)
            else:
                messagebox.showerror("Erro", "Formato de data inválido. Use YYYY-MM-DD ou DD/MM/YYYY.")
                return

    salvar_registro(nome, evento, data_emissao, numero)

    safe_name = nome.replace(" ", "_")
    filename = f"{numero}_{safe_name}.pdf"
    arquivo = os.path.join(OUT_DIR, filename)
    gerar_certificado_pdf(nome, evento, data_emissao, numero, arquivo)

    messagebox.showinfo("Pronto", f"Certificado gerado: {arquivo}")
    # Opcional: abrir pasta
    if messagebox.askyesno("Abrir pasta?", "Deseja abrir a pasta com os certificados?"):
        import subprocess, sys
        if sys.platform.startswith("win"):
            subprocess.Popen(f'explorer "{os.path.abspath(OUT_DIR)}"')
        elif sys.platform.startswith("darwin"):
            subprocess.Popen(["open", os.path.abspath(OUT_DIR)])
        else:
            subprocess.Popen(["xdg-open", os.path.abspath(OUT_DIR)])


# Monta a janela
init_db()
root = tk.Tk()
root.title("Gerador de Certificados - Simples")

frame = tk.Frame(root, padx=12, pady=12)
frame.pack()

tk.Label(frame, text="Nome completo:").grid(row=0, column=0, sticky="w")
entry_nome = tk.Entry(frame, width=50)
entry_nome.grid(row=0, column=1, pady=4)

tk.Label(frame, text="Evento / Curso:").grid(row=1, column=0, sticky="w")
entry_evento = tk.Entry(frame, width=50)
entry_evento.grid(row=1, column=1, pady=4)

tk.Label(frame, text="Data (YYYY-MM-DD ou DD/MM/YYYY):").grid(row=2, column=0, sticky="w")
entry_data = tk.Entry(frame, width=20)
entry_data.grid(row=2, column=1, sticky="w", pady=4)
entry_data.insert(0, datetime.today().strftime("%Y-%m-%d"))

tk.Label(frame, text="Número do certificado (opcional):").grid(row=3, column=0, sticky="w")
entry_num = tk.Entry(frame, width=30)
entry_num.grid(row=3, column=1, sticky="w", pady=4)

btn = tk.Button(frame, text="Salvar e Gerar PDF", command=gerar_e_salvar, width=20)
btn.grid(row=4, column=0, columnspan=2, pady=12)

root.mainloop()
