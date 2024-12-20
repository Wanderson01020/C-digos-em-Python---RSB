#!/usr/bin/env python
# coding: utf-8

# In[1]:


#!/usr/bin/env python
# coding: utf-8

# In[2]:


#!/usr/bin/env python
# coding: utf-8

import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog
import json
import os

# Configurações do servidor SMTP
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587

# Variáveis globais para armazenar o email, senha, caminho da planilha e caminho da assinatura
EMAIL_REMETENTE = ''
SENHA_REMETENTE = ''
("=", "''")
CAMINHO_ASSINATURA = ''

# Caminho do arquivo de configuração
CONFIG_FILE = 'config.json'

# Variáveis globais para armazenar os caminhos das imagens de cabeçalho e rodapé
caminho_cabecalho = None
caminho_rodape = None

# Valor fixo da tabulação
TABULACAO = 310

def formatar_corpo_html(corpo_modelo, tabulacao):
    # Substitui quebras de linha por tags HTML <br>
    corpo_modelo = corpo_modelo.replace('\n', '<br>')
    # Adiciona tabulação ao texto do corpo do e-mail
    corpo_modelo = f'<div style="max-width:50%; margin:0 auto; text-align:justify">{corpo_modelo}</div>'
    return corpo_modelo
# =========================================================================================================================
def enviar_email(destinatario, instituicao, link, corpo_modelo, assunto, tabulacao):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = destinatario
    msg['Subject'] = assunto

    # Substitui os placeholders no corpo do e-mail pelo conteúdo real
    corpo_email = corpo_modelo.replace('{instituicao}', instituicao).replace('{link}', link)

    # Corpo do e-mail com HTML
    corpo_html = f"""
    <html>
    <body>
        <div style="text-align: center;">
            <img src="cid:cabecalho" alt="Cabeçalho" style="width: 60%; height: auto;" />
        </div>
        <p>{formatar_corpo_html(corpo_email, tabulacao)}</p>
        <div style="text-align: center;">
            <img src="cid:rodape" alt="Rodapé" style="width: 60%; height: auto;" />
        </div>
        <div style="text-align: left;">
            <img src="cid:assinatura" alt="Assinatura" style="width: auto; height: auto;" />
        </div>
    </body>
    </html>
    """
    msg.attach(MIMEText(corpo_html, 'html'))
# ===============================================================================================================================
# def enviar_email(destinatario, instituicao, link, corpo_modelo, assunto, tabulacao):
#     msg = MIMEMultipart()
#     msg['From'] = EMAIL_REMETENTE
#     msg['To'] = destinatario
#     msg['Subject'] = assunto

#     # Substitui os placeholders no corpo do e-mail pelo conteúdo real
#     corpo_email = corpo_modelo.replace('{instituicao}', instituicao).replace('{link}', link)

#     # Corpo do e-mail com HTML
#     corpo_html = f"""
#     <html>
#     <body>
#         <div style="text-align: center;">
#             <img src="cid:cabecalho" alt="Cabeçalho" style="width: 60%; height: auto;" />
#         </div>
#         <p>{formatar_corpo_html(corpo_email, tabulacao)}</p>
#         <div style="text-align: center;">
#             <img src="cid:rodape" alt="Rodapé" style="width: 60%; height: auto;" />
#         </div>
#         <div style="text-align: left;">
#             <img src="cid:assinatura" alt="Assinatura" style="width: auto; height: auto;" />
#         </div>
#     </body>
#     </html>
#     """
#     msg.attach(MIMEText(corpo_html, 'html'))

#     # (restante do código)

# ===============================================================================================================================
    # Anexar imagem de cabeçalho
    if caminho_cabecalho:
        with open(caminho_cabecalho, 'rb') as img:
            img_attachment = MIMEBase('application', 'octet-stream')
            img_attachment.set_payload(img.read())
            encoders.encode_base64(img_attachment)
            img_attachment.add_header('Content-ID', '<cabecalho>')
            img_attachment.add_header('Content-Disposition', 'inline', filename="cabecalho.jpg")
            msg.attach(img_attachment)

    # Anexar imagem de rodapé
    if caminho_rodape:
        with open(caminho_rodape, 'rb') as img:
            img_attachment = MIMEBase('application', 'octet-stream')
            img_attachment.set_payload(img.read())
            encoders.encode_base64(img_attachment)
            img_attachment.add_header('Content-ID', '<rodape>')
            img_attachment.add_header('Content-Disposition', 'inline', filename="rodape.jpg")
            msg.attach(img_attachment)

    # Anexar imagem de assinatura
    if CAMINHO_ASSINATURA:
        with open(CAMINHO_ASSINATURA, 'rb') as img:
            img_attachment = MIMEBase('application', 'octet-stream')
            img_attachment.set_payload(img.read())
            encoders.encode_base64(img_attachment)
            img_attachment.add_header('Content-ID', '<assinatura>')
            img_attachment.add_header('Content-Disposition', 'inline', filename="assinatura.jpg")
            msg.attach(img_attachment)

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_REMETENTE, SENHA_REMETENTE)
        server.sendmail(EMAIL_REMETENTE, destinatario, msg.as_string())
        server.quit()
        print(f'E-mail enviado para {destinatario}')
    except Exception as e:
        print(f'Falha ao enviar e-mail para {destinatario}: {str(e)}')

def mostrar_janela_progresso(total_emails):
    janela_progresso = tk.Toplevel(root)
    janela_progresso.title("Progresso do Envio")

    tk.Label(janela_progresso, text="Enviando e-mails...").pack(padx=10, pady=5)
    progresso = tk.Label(janela_progresso, text=f"0/{total_emails} e-mails enviados")
    progresso.pack(padx=10, pady=5)
    return janela_progresso, progresso

def atualizar_progresso(progresso_label, enviado, total):
    progresso_label.config(text=f"{enviado}/{total} e-mails enviados")
    root.update()

def enviar_emails():
    corpo_modelo = text_corpo.get("1.0", tk.END).strip()  # Obtém o texto do campo de entrada
    assunto = entry_assunto.get().strip()  # Obtém o texto do campo de assunto

    if not corpo_modelo:
        messagebox.showwarning("Aviso", "O corpo do e-mail não pode estar vazio.")
        return
    if not assunto:
        messagebox.showwarning("Aviso", "O assunto do e-mail não pode estar vazio.")
        return
    
    try:
        df = pd.read_excel(CAMINHO_PLANILHA)
        total_emails = len(df)
        janela_progresso, progresso_label = mostrar_janela_progresso(total_emails)
        
        for index, row in df.iterrows():
            enviar_email(row['email'], row['instituicao'], row['link'], corpo_modelo, assunto, TABULACAO)
            atualizar_progresso(progresso_label, index + 1, total_emails)
        
        janela_progresso.destroy()
        messagebox.showinfo("Sucesso", "E-mails enviados com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao enviar e-mails: {str(e)}")

def selecionar_cabecalho():
    global caminho_cabecalho
    caminho_cabecalho = filedialog.askopenfilename(title="Selecione uma imagem de cabeçalho", filetypes=[("Imagens", "*.png;*.jpg;*.jpeg")])
    if caminho_cabecalho:
        label_cabecalho.config(text=f"Cabeçalho selecionado: {caminho_cabecalho.split('/')[-1]}")
    else:
        label_cabecalho.config(text="Nenhum cabeçalho selecionado")

def selecionar_rodape():
    global caminho_rodape
    caminho_rodape = filedialog.askopenfilename(title="Selecione uma imagem de rodapé", filetypes=[("Imagens", "*.png;*.jpg;*.jpeg")])
    if caminho_rodape:
        label_rodape.config(text=f"Rodapé selecionado: {caminho_rodape.split('/')[-1]}")
    else:
        label_rodape.config(text="Nenhum rodapé selecionado")

def selecionar_assinatura():
    global CAMINHO_ASSINATURA
    CAMINHO_ASSINATURA = filedialog.askopenfilename(title="Selecione uma imagem de assinatura", filetypes=[("Imagens", "*.png;*.jpg;*.jpeg")])
    if CAMINHO_ASSINATURA:
        label_assinatura.config(text=f"Assinatura selecionada: {CAMINHO_ASSINATURA.split('/')[-1]}")
    else:
        label_assinatura.config(text="Nenhuma assinatura selecionada")

def abrir_configuracao():
    config_janela = tk.Toplevel(root)
    config_janela.title("Configurações de E-mail")

    tk.Label(config_janela, text="Seu E-mail:").pack(padx=10, pady=5)
    email_entry = tk.Entry(config_janela, width=50)
    email_entry.pack(padx=10, pady=5)
    email_entry.insert(0, EMAIL_REMETENTE)

    tk.Label(config_janela, text="Token:").pack(padx=10, pady=5)
    token_entry = tk.Entry(config_janela, show='', width=50)  # Deixando o token visível
    token_entry.pack(padx=10, pady=5)
    token_entry.insert(0, SENHA_REMETENTE)

    def salvar_configuracao():
        global EMAIL_REMETENTE, SENHA_REMETENTE
        EMAIL_REMETENTE = email_entry.get()
        SENHA_REMETENTE = token_entry.get()
        with open(CONFIG_FILE, 'w') as f:
            json.dump({'email': EMAIL_REMETENTE, 'senha': SENHA_REMETENTE}, f)
        config_janela.destroy()

    tk.Button(config_janela, text="Salvar", command=salvar_configuracao).pack(padx=10, pady=5)

def abrir_configuracoes():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            config = json.load(f)
            global EMAIL_REMETENTE, SENHA_REMETENTE
            EMAIL_REMETENTE = config.get('email', '')
            SENHA_REMETENTE = config.get('senha', '')

# Criação da interface gráfica
root = tk.Tk()
root.title("Envio de E-mails")

# Carregar configurações
abrir_configuracoes()

# Corpo do e-mail
tk.Label(root, text="Digite o corpo do e-mail com as variáveis para mala direta {instituicao} e {link}:").pack(padx=10, pady=5)
text_corpo = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=20)
text_corpo.pack(padx=10, pady=5)

# Assunto do e-mail
tk.Label(root, text="Assunto:").pack(padx=10, pady=5)
entry_assunto = tk.Entry(root, width=80)
entry_assunto.pack(padx=10, pady=5)

# Imagem de cabeçalho
tk.Button(root, text="Selecionar Cabeçalho", command=selecionar_cabecalho).pack(padx=10, pady=5)
label_cabecalho = tk.Label(root, text="Nenhum cabeçalho selecionado")
label_cabecalho.pack(padx=10, pady=5)

# Imagem de rodapé
tk.Button(root, text="Selecionar Rodapé", command=selecionar_rodape).pack(padx=10, pady=5)
label_rodape = tk.Label(root, text="Nenhum rodapé selecionado")
label_rodape.pack(padx=10, pady=5)

# Imagem de assinatura
tk.Button(root, text="Selecionar Assinatura", command=selecionar_assinatura).pack(padx=10, pady=5)
label_assinatura = tk.Label(root, text="Nenhuma assinatura selecionada")
label_assinatura.pack(padx=10, pady=5)

# Caminho da planilha
tk.Button(root, text="Selecionar Planilha", command=lambda: globals().update({'CAMINHO_PLANILHA': filedialog.askopenfilename(title="Selecione uma planilha", filetypes=[("Arquivos Excel", "*.xlsx")])})).pack(padx=10, pady=5)

# Botão de envio
tk.Button(root, text="Enviar E-mails", command=enviar_emails).pack(padx=10, pady=10)

# Configuração do e-mail
tk.Button(root, text="Configurações de E-mail", command=abrir_configuracao).pack(padx=10, pady=10)

root.mainloop()


# In[ ]:


pip install PyInstaller


# In[ ]:




