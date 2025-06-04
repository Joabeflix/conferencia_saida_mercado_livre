import os
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import StringVar
from tratamento_planilha.planilha_romaneio import TratamentoPlanilhaMercadoLivre
from tratamento_planilha.gerar_relatorio import gerar_relatorio_exel
from utils.variaveis_json import *
from servico_email.servico_de_email import Email


class ConferenciaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Conferência de saída - Mercado Livre")
        self.root.geometry("1194x900")

        self.label_atual_log_y = 560
        self.label_atual_log_x = 20
        self.qtd_log = 0

        self.local_planilha_meli = None
        self.lista_confirmados_gerar_relatorio = []
        self.lista_todos_gerar_relatorio = []

        self.dados = self.carregar_dados()

        self.label_titulo = ttk.Label(root, text="⏳ PENDENTES", font=("Segoe UI", 20, "bold"), foreground="yellow")
        self.label_titulo.place(x=200, y=10)

        self.label_log = ttk.Label(root, text="✅ CONFIRMADOS", font=("Segoe UI", 20, "bold"), foreground="green")
        self.label_log.place(x=747, y=10)

        self.lista_pendentes = ttk.Text(root, width=80, height=20, font=("Segoe UI", 10))
        self.lista_pendentes.place(x=20, y=65)

        self.lista_confirmados = ttk.Text(root, width=80, height=20, font=("Segoe UI", 10))
        self.lista_confirmados.place(x=600, y=65)

        self.label_mensagem = ttk.Label(root, text="Entrada do código etiqueta", font=("Segoe UI", 10), foreground="green")
        self.label_mensagem.place(x=20, y=423)

        self.codigo_var = StringVar()
        self.entry_codigo = ttk.Entry(root, textvariable=self.codigo_var, width=40, font=("Segoe UI", 10))
        self.entry_codigo.place(x=20, y=445)
        self.entry_codigo.bind("<Return>", self.conferir_codigo)

        self.botao = ttk.Button(root, text="Conferir", bootstyle='success-outline', command=self.conferir_codigo)
        self.botao.place(x=320, y=446)
        self.gerar_relatorio = ttk.Button(root, text="Relatório Email", bootstyle='success-outline', command=self.enviar_relatorio_email)
        self.gerar_relatorio.place(x=460, y=446)

        self.label_mensagem = ttk.Label(root, text="", font=("Segoe UI", 10), foreground="blue")
        self.label_mensagem.place(x=390, y=446)

        self.label_log = ttk.Label(root, text="LOG", font=("Segoe UI", 15), foreground="pink")
        self.label_log.place(x=20, y=520)

        self.atualizar_lista()

    def carregar_dados(self):
        app = TratamentoPlanilhaMercadoLivre(
            nome_aba_cod_rastreiro=COLUNA_COD_RASTREIO,
            nome_aba_nome_cliente=COLUNA_NOME_CLIENTE)
        dicionario = app._criar_dicionario()
        
        self.local_planilha_meli = app.retornar_local_planilha()

        return dicionario
    


    def atualizar_lista(self):
        self.lista_pendentes.delete("1.0", "end")
        self.lista_confirmados.delete("1.0", "end")

        for codigo, info in self.dados.items():

            if not info["status_conferencia"]:
                text = f"{codigo} - {info['nome_cliente']}\n"
                self.lista_pendentes.insert("end", text)

            else:
                text = f"{codigo} - {info['nome_cliente']}\n"
                self.lista_confirmados.insert("end", text)


    def conferir_codigo(self, event=None):
        codigo = self.codigo_var.get().strip()
        if not codigo:
            return

        if codigo in self.dados:
            if not self.dados[codigo]["status_conferencia"]:
                self.dados[codigo]["status_conferencia"] = True
                self.mensagem(f"✅ {codigo} conferido com sucesso!", "green")
                self.lista_confirmados_gerar_relatorio.append([codigo, self.dados[codigo]["nome_cliente"]])
      
            else:
                self.mensagem(f"⚠️ {codigo} já foi conferido anteriormente.", "yellow")
          
        else:
            self.mensagem(f"❌ Código {codigo} não encontrado na lista!", "red")


        self.codigo_var.set("")
        self.atualizar_lista()

        if all(info["status_conferencia"] for info in self.dados.values()):
            self.mensagem("✅ Todos os itens foram conferidos com sucesso!", "green")

    def mensagem(self, texto, cor):
        self.label_mensagem.config(text=texto, foreground=cor)
        self.criar_label_log(texto=texto, cor=cor)


    def criar_label_log(self, texto, cor):
        x = self.label_atual_log_x
        y = self.label_atual_log_y
        y_novo = y + 20
        label_mensagem = ttk.Label(self.root, text=texto, font=("Segoe UI", 10), foreground=cor)
        label_mensagem.place(x=x, y=y)

        self.label_atual_log_y = y_novo

        self.qtd_log+=1

        if self.qtd_log == 14:
            self.label_atual_log_x += 310
            self.qtd_log = 0
            self.label_atual_log_y = 560

    def enviar_relatorio_email(self):
        gerar_relatorio_exel(rf'{self.local_planilha_meli}', nome_aba='Confirmados', dados=self.lista_confirmados_gerar_relatorio)
        email = Email(
            remetente=REMETENTE
        )
        email.definir_senha(SENHA_DE_APP_EMAIL)
        email.enviar(
            assunto="Relatório pedidos confirmados",
            mensagem=f'Segue no anexo os pedidos confirmados.',
            destinatarios=DESTINATARIOS,
            anexos=[rf'{self.local_planilha_meli}']
        )

if __name__ == "__main__":
    root = ttk.Window(themename=TEMA_PROGRAMA)
    app = ConferenciaApp(root)
    root.mainloop()

    
