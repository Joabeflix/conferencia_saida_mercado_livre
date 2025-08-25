import os
import tkinter as tk
import ttkbootstrap as ttk
from tkinter import StringVar
from datetime import datetime
from utils.variaveis_json import *
from servico_email.servico_de_email import Email
from tratamento_planilha.gerar_relatorio import gerar_relatorio_exel
from tratamento_planilha.planilha_romaneio import TratamentoPlanilhaMercadoLivre

class ConferenciaApp:
    def __init__(self, root: ttk.Window) -> None:
        self.root = root
        self.root.title("Conferência de saída - Mercado Livre")
        self.root.geometry("1194x900")

        self.label_atual_log_y = 560
        self.label_atual_log_x = 20
        self.qtd_log = 0

        self.contagem_confirmado = 1

        self.local_planilha_meli = None
        self.lista_confirmados_gerar_relatorio = []
        self.lista_todos_gerar_relatorio = []

        self.dados = self.carregar_dados()

        self.label_titulo = ttk.Label(root, text="⏳ PENDENTES", font=("Segoe UI", 20, "bold"), foreground="yellow")
        self.label_titulo.place(x=200, y=10)

        self.label_log = ttk.Label(root, text="✅ CONFIRMADOS", font=("Segoe UI", 20, "bold"), foreground="green")
        self.label_log.place(x=747, y=10)

        self.label_qtd_confirmada = ttk.Label(root, text="", font=("Segoe UI", 20), foreground="green")
        self.label_qtd_confirmada.place(x=1000, y=10)

        self.bloco_pendentes = ttk.Text(root, width=80, height=20, font=("Segoe UI", 10))
        self.bloco_pendentes.place(x=20, y=65)

        self.bloco_confirmados = ttk.Text(root, width=80, height=20, font=("Segoe UI", 10))
        self.bloco_confirmados.place(x=600, y=65)

        self.label_mensagem = ttk.Label(root, text="Entrada do código etiqueta", font=("Segoe UI", 10), foreground="green")
        self.label_mensagem.place(x=20, y=423)

        self.codigo_var = StringVar()
        self.entry_codigo = ttk.Entry(root, textvariable=self.codigo_var, width=40, font=("Segoe UI", 10))
        self.entry_codigo.place(x=20, y=445)
        self.entry_codigo.bind("<Return>", self.conferir_codigo)

        self.botao_conferir = ttk.Button(root, text="Conferir", style='success-outline', command=self.conferir_codigo)
        self.botao_conferir.place(x=320, y=446)

        self.label_observacao_email = ttk.Label(root, text="Observação Email", font=("Segoe UI", 10), foreground="green")
        self.label_observacao_email.place(x=827, y=423)

        self.entrada_observacao_email = ttk.Entry(root, width=40)
        self.entrada_observacao_email.place(x=827, y=445)
        self.gerar_relatorio = ttk.Button(root, text="Enviar Email", style='success-outline', command=self.enviar_relatorio_email)
        self.gerar_relatorio.place(x=1086, y=445)

        self.label_mensagem = ttk.Label(root, text="", font=("Segoe UI", 10), foreground="blue")
        self.label_mensagem.place(x=390, y=446)

        self.label_log = ttk.Label(root, text="LOG", font=("Segoe UI", 15), foreground="pink")
        self.label_log.place(x=20, y=520)

        self.botao_pedido_avulso = ttk.Button(root, text="Adicionar", style="success", command=self.adicionar_pedido_avulso)
        self.botao_pedido_avulso.place(x=20, y=20)

        # self.label_qtd_faltando = ttk.Label(root, text="", font=("Segoe UI", 10), foreground="blue")
        # self.label_qtd_faltando.place()

        self.atualizar_lista()

    def carregar_dados(self) -> dict:
        app = TratamentoPlanilhaMercadoLivre(
            nome_aba_cod_rastreiro=COLUNA_COD_RASTREIO,
            nome_aba_nome_cliente=COLUNA_NOME_CLIENTE)
        dicionario = app._criar_dicionario()
        
        self.local_planilha_meli = app.retornar_local_planilha()

        return dicionario
    


    def atualizar_lista(self) -> None:
        self.bloco_pendentes.delete("1.0", "end")
        self.bloco_confirmados.delete("1.0", "end")

        for codigo, info in self.dados.items():

            if not info["status_conferencia"]:
                text = f"{codigo} - {info['nome_cliente']}\n"
                self.bloco_pendentes.insert("end", text)

            else:
                text = f"{codigo} - {info['nome_cliente']}\n"
                self.bloco_confirmados.insert("end", text)


    def conferir_codigo(self, event=None) -> None:
        codigo = self.codigo_var.get().strip()
        if not codigo:
            return None

        if codigo in self.dados:
            if not self.dados[codigo]["status_conferencia"]:
                self.dados[codigo]["status_conferencia"] = True
                self.mensagem(f"✅ {codigo} Conferido", "green")
                self.lista_confirmados_gerar_relatorio.append([codigo, self.dados[codigo]["nome_cliente"]])
                self.label_qtd_confirmada.config(text=f'{self.contagem_confirmado}')
                self.contagem_confirmado+=1
      
            else:
                self.mensagem(f"⚠️ {codigo} Já conferido", "yellow")
          
        else:
            self.mensagem(f"❌ {codigo} não encontrado.", "red")


        self.codigo_var.set("")
        self.atualizar_lista()

        if all(info["status_conferencia"] for info in self.dados.values()):
            self.mensagem("✅ Todos os itens foram conferidos com sucesso!", "green")

    def adicionar_pedido_avulso(self) -> None:
    #    # dicionario[12345678910] = {'nome_cliente': nome_cliente, 'status_conferencia': False}
    #     print(self.dados)
    #     print(self.dados)
    #     self.atualizar_lista()
        def fechar(root: tk.Tk) -> None:
            root.destroy()
        def confirmar() -> None:
            if cod_rastreio.get() and nome_cliente.get():
                self.dados[cod_rastreio.get()] = {'nome_cliente': nome_cliente.get(), 'status_conferencia': False}
                self.atualizar_lista()
                fechar(tela)
            else:
                fechar(tela)

        tela = tk.Tk()
        tela.title("Ad Pedido")
        tela.geometry("250x135")
        label_1 = tk.Label(tela, text="Código de Rastreio")
        label_1.pack()
        cod_rastreio = tk.Entry(tela)
        cod_rastreio.pack()       
        label_2 = tk.Label(tela, text="Nome do cliente")
        label_2.pack()
        nome_cliente = tk.Entry(tela)
        nome_cliente.pack()
        espaco = tk.Label(tela, text="ll")
        espaco.pack()
        botao_confirmar = tk.Button(tela, text="Adicionar", command=confirmar)
        botao_confirmar.pack()
        
        tela.mainloop()

            

    def mensagem(self, texto: str, cor: str) -> None:
        self.label_mensagem.config(text=texto, foreground=cor)
        self.criar_label_log(texto=texto, cor=cor)


    def criar_label_log(self, texto: str, cor: str) -> None:
        x = self.label_atual_log_x
        y = self.label_atual_log_y
        y_novo = y + 20
        label_mensagem = ttk.Label(self.root, text=texto, font=("Segoe UI", 10), foreground=cor)
        label_mensagem.place(x=x, y=y)

        self.label_atual_log_y = y_novo

        self.qtd_log+=1

        if self.qtd_log == 14:
            self.label_atual_log_x += 227
            self.qtd_log = 0
            self.label_atual_log_y = 560


    def lista_pendentes(self) -> str | None:
        pendentes = ""
        for codigo, info in self.dados.items():
            if not info["status_conferencia"]:
                text = f"{codigo} - {info['nome_cliente']}"
                pendentes+=f"\n{text}"
        if pendentes:
            return pendentes
        
    def enviar_relatorio_email(self) -> None:
        gerar_relatorio_exel(rf'{self.local_planilha_meli}', nome_aba='Confirmados', dados=self.lista_confirmados_gerar_relatorio)
        email = Email(
            remetente=REMETENTE
        )
        email.definir_senha(SENHA_DE_APP_EMAIL)

        lista_pendentes = self.lista_pendentes()
        


        data_hora = datetime.now()

        data_hora_formatada = data_hora.strftime("%d-%m-%Y %H:%M:%S")

        observacao = self.entrada_observacao_email.get().strip()

        mensagem = f"""
Lista de pedidos confirmados no Mercado Livre.\n
DATA E HORA DO ENVIO: {data_hora_formatada}\n
{f'Pedidos não confirmados:\n{lista_pendentes}' if lista_pendentes else ''}\n
{f'Observação:\n{observacao}' if observacao  else 'Sem observações.'}

------------------------------------------------------------
USUÁRIO: {os.environ['USERNAME']}
NOME DO COMPUTADOR: {os.environ['COMPUTERNAME']}
------------------------------------------------------------
        """

        email.enviar(
            assunto="Relatório pedidos confirmados",
            mensagem=mensagem,
            destinatarios=DESTINATARIOS,
            anexos=[rf'{self.local_planilha_meli}']
        )

if __name__ == "__main__":
    root = ttk.Window(themename=TEMA_PROGRAMA)
    app = ConferenciaApp(root)
    root.mainloop()

    
