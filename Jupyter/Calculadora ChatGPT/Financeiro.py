import pandas as pd
import smtplib
import tkinter as tk
from tkinter import filedialog, messagebox
from email.mime.text import MIMEText

class App:
    def __init__(self, master):
        # janela principal
        self.master = master
        self.master.title("Lembrete de Pagamento")

        # rótulo para selecionar planilha de Excel
        tk.Label(self.master, text="Selecione a planilha de Excel:").grid(row=0, column=0, padx=10, pady=10)

        # botão para selecionar planilha de Excel
        self.btn_selecionar = tk.Button(self.master, text="Selecionar", command=self.selecionar_planilha)
        self.btn_selecionar.grid(row=0, column=1, padx=10, pady=10)

        # rótulo para informar quantas pessoas estão devendo
        tk.Label(self.master, text="Número de pessoas devendo:").grid(row=1, column=0, padx=10, pady=10)

        # variável para exibir o número de pessoas devendo
        self.num_devendo = tk.StringVar(value="0")
        tk.Label(self.master, textvariable=self.num_devendo).grid(row=1, column=1, padx=10, pady=10)

        # rótulo para informar quantos e-mails foram enviados
        tk.Label(self.master, text="Número de e-mails enviados:").grid(row=2, column=0, padx=10, pady=10)

        # variável para exibir o número de e-mails enviados
        self.num_enviados = tk.StringVar(value="0")
        tk.Label(self.master, textvariable=self.num_enviados).grid(row=2, column=1, padx=10, pady=10)

        # botão para enviar e-mails
        self.btn_enviar = tk.Button(self.master, text="Enviar E-mails", command=self.enviar_emails, state=tk.DISABLED)
        self.btn_enviar.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

    def selecionar_planilha(self):
        # abrir janela de seleção de arquivo
        filepath = filedialog.askopenfilename(initialdir="/", title="Selecione a planilha de Excel", filetypes=[("Planilha de Excel", "*.xlsx")])

        if filepath:
            try:
                # ler planilha de excel com pandas
                self.df = pd.read_excel(filepath)

                # filtrar apenas as pessoas que estão devendo
                self.df_devedores = self.df[self.df['Valor a Pagar'] > self.df['Valor Pago']]

                # atualizar número de pessoas devendo
                self.num_devendo.set(str(len(self.df_devedores)))

                # habilitar botão para enviar e-mails
                self.btn_enviar.config(state=tk.NORMAL)

                # exibir mensagem de sucesso
                messagebox.showinfo("Sucesso", "Planilha selecionada com sucesso!")
            except:
                # exibir mensagem de erro
                messagebox.showerror("Erro", "Ocorreu um erro ao abrir a planilha de Excel. Verifique se o arquivo selecionado é uma planilha de Excel válida.")
        else:
            # exibir mensagem de cancelamento
                def enviar_emails(self):
    # pegar credenciais do usuário
    email = input("Digite seu e-mail: ")
    senha = input("Digite sua senha: ")

    # conectar ao servidor SMTP do Gmail
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        # fazer login com as credenciais do usuário
        smtp.login(email, senha)

        # percorrer cada linha da planilha de devedores
        for _, row in self.df_devedores.iterrows():
            # criar mensagem de e-mail
            mensagem = f"Olá {row['Nome']},\n\nVocê está devendo R${row['Valor a Pagar'] - row['Valor Pago']:.2f}.\n\nPor favor, faça o pagamento o mais rápido possível.\n\nObrigado."
            msg = MIMEText(mensagem)

            # adicionar informações do destinatário e remetente
            msg['Subject'] = 'Lembrete de Pagamento'
            msg['From'] = email
            msg['To'] = row['E-mail']

            # enviar e-mail
            smtp.send_message(msg)

            # atualizar número de e-mails enviados
            self.num_enviados.set(str(int(self.num_enviados.get()) + 1))

    # exibir mensagem de sucesso
    messagebox.showinfo("Sucesso", "E-mails enviados com sucesso!")
    
