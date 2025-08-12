import os
import sqlite3
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, filedialog
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth
from PIL import Image, ImageTk
import pandas as pd



# ---------- Configs ----------
dados_dir = os.path.join(os.path.dirname(__file__), "dados")
os.makedirs(dados_dir, exist_ok=True)

# Arquivo do banco de dados
DB_FILE = os.path.join(dados_dir, "certificados.db")

OUT_DIR = os.path.join(dados_dir, "certificados_pdfs")
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

def draw_multiline_text_centered(c, text, center_x, start_y, max_width, line_height, font_name="Helvetica", font_size=14):
    c.setFont(font_name, font_size)
    words = text.split()
    line = ""
    y_offset = 0

    for word in words:
        test_line = line + (word + " ")
        text_width = stringWidth(test_line, font_name, font_size)
        if text_width <= max_width:
            line = test_line
        else:
            # centralizar linha atual
            line_width = stringWidth(line, font_name, font_size)
            c.drawString(center_x - line_width/2, start_y - y_offset, line.strip())
            line = word + " "
            y_offset += line_height

    if line:
        line_width = stringWidth(line, font_name, font_size)
        c.drawString(center_x - line_width/2, start_y - y_offset, line.strip())


# ---------- Gerar PDF ----------
def gerar_certificado_pdf(nome, evento, data_emissao, numero, arquivo_saida):
    # landscape A4
    c = canvas.Canvas(arquivo_saida, pagesize=landscape(A4))
    largura, altura = landscape(A4)

    # Background opcional: imagem PNG/JPG como template
    bg_path = "template.png"
    if os.path.exists(bg_path):
        c.drawImage(bg_path, 0, 0, width=largura, height=altura)

    # Título
    c.setFont("Helvetica-Bold", 30)
    c.drawCentredString(largura/2, altura - 60*mm, "CERTIFICADO")

    # Texto principal
    c.setFont("Helvetica", 18)
    c.drawCentredString(largura/2, altura - 80*mm, "Certificamos que")

    c.setFont("Helvetica-Bold", 26)
    c.drawCentredString(largura/2, altura - 100*mm, nome)

    # Texto detalhado (quebrado em 2 linhas para ficar melhor)
    c.setFont("Helvetica", 14)

    max_text_width = largura - 60*mm  # 30mm margem de cada lado
    start_y = altura - 120*mm
    line_height = 16
    center_x = largura / 2

    linha1 = f"Concluiu com sucesso o curso de {evento} realizado pela Igreja Evangélica Assembleia de Deus"
    linha2 = "Ministério Missão, sediada em Av. Sen. Canedo - Goiânia / Extensão, Sen. Canedo - GO, 75256-207."

    draw_multiline_text_centered(c, linha1, center_x, start_y, max_text_width, line_height)
    draw_multiline_text_centered(c, linha2, center_x, start_y - line_height*3, max_text_width, line_height)

    # Data e número do certificado
    c.setFont("Helvetica", 14)
    try:
        dt = datetime.strptime(data_emissao, "%Y-%m-%d")
        data_formatada = dt.strftime("%d/%m/%Y")
    except Exception:
        data_formatada = data_emissao
    c.drawString(30*mm, 20*mm, f"Data de emissão: {data_formatada}")
    c.drawString(30*mm, 15*mm, f"Nº certificado: {numero}")

    # Linha para assinatura e texto
    c.line(largura - 90*mm, 25*mm, largura - 30*mm, 25*mm)
    c.setFont("Helvetica", 12)
    c.drawCentredString(largura - 60*mm, 18*mm, "Assinatura / Organização")

    c.showPage()
    c.save()

# ---------- Exportar para Excel ----------
def gerar_em_lote():
    arquivo_excel = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx;*.xls")]
    )
     
    if not arquivo_excel:
        return
    
    try:
        df = pd.read_excel(arquivo_excel)
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível ler o arquivo: {e}")
    
    # Verifica se as colunas existem (nome, evento, data, numero opcional)
    colunas_necessarias = ['nome', 'evento', 'data_emissao']
    for col in colunas_necessarias:
        if col not in df.columns.str.lower():
            messagebox.showerror("Erro", f"Coluna '{col}' não encontrada no arquivo Excel.")
            return
    
    # Padroniza nomes das colunas para lowercase para evitar erros
    df.columns = df.columns.str.lower()

    total = len(df)
    gerados = 0

    for idx, row in df.iterrows():
        nome = str(row['nome']).strip()
        evento = str(row['evento']).strip()
        data_emissao = str(row['data_emissao']).strip()
        numero = str(row["numero"]).strip() if 'numero' in df.columns else ""

        if not nome or not evento or not data_emissao:
            continue # Pula linhas incompletas

        #Gera número automatico se estiver vazio
        if not numero:
            conn = sqlite3.connect(DB_FILE)
            cur = conn.cursor()
            cur.execute("SELECT MAX(id) FROM certificado")
            ultimo_id = cur.fetchone()[0]
            conn.close()
            novo_id = (ultimo_id or 0) + 1
            numero = f"CERT-{novo_id:04d}"

        #Salva e gera PDF
        salvar_registro(nome, evento, data_emissao, numero)

        safe_name = nome.replace(" ", "_")
        filename = f"{numero}_{safe_name}.pdf"
        arquivo = os.path.join(OUT_DIR, filename)
        gerar_certificado_pdf(nome, evento, data_emissao, numero, arquivo)

        gerados += 1

    messagebox.showinfo("Pronto", f"Certificados gerados: {gerados}")

# ---------- Interface Tkinter ----------
def gerar_e_salvar():
    nome = entry_nome.get().strip()
    evento = entry_evento.get().strip()
    data_emissao = entry_data.get().strip()
    numero = ""  # Pode ser deixado em branco para auto-gerar

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

    # Limpa os campos
    entry_nome.delete(0, tk.END)
    entry_evento.delete(0, tk.END)
    entry_data.delete(0, tk.END)

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
root.title("CertiFé")

#Frame para logo + Titulo
top_frame = tk.Frame(root, padx=12, pady=12)
top_frame.pack(pady=10)

#Carregar a logo (PNG)
logo_img = Image.open("logo.png")
logo_img = logo_img.resize((100, 100), Image.Resampling.LANCZOS)

#Converter para imagem que o tkinter entende
logo_img = ImageTk.PhotoImage(logo_img)

#label com a logo 
label_logo = tk.Label(top_frame, image=logo_img)
label_logo.pack(side="left", padx=10)
root.logo_img = logo_img

#Label com o titulo
label_title = tk.Label(top_frame, text="CertiFé", font=("Helvetica", 24, "bold"))
label_title.pack(side="left", padx=10)

frame = tk.Frame(root, padx=12, pady=12)
frame.pack()

tk.Label(frame, text="Nome completo:").grid(row=0, column=0, sticky="w")
entry_nome = tk.Entry(frame, width=50)
entry_nome.grid(row=0, column=1, pady=4)

tk.Label(frame, text="Curso:").grid(row=1, column=0, sticky="w")
entry_evento = tk.Entry(frame, width=50)
entry_evento.grid(row=1, column=1, pady=4)

tk.Label(frame, text="Data:").grid(row=2, column=0, sticky="w")
entry_data = tk.Entry(frame, width=20)
entry_data.grid(row=2, column=1, sticky="w", pady=4)
entry_data.insert(0, datetime.today().strftime("%d/%m/%Y"))

btn = tk.Button(frame, text="Salvar e Gerar PDF", command=gerar_e_salvar, width=20)
btn.grid(row=4, column=0, columnspan=2, pady=12)

btn_lote = tk.Button(frame, text="Gerar em Lote", command=gerar_em_lote, width=20)
btn_lote.grid(row=5, column=0, columnspan=2, pady=12)
root.mainloop()
