import winsound
import ttkbootstrap as ttk
from tkinter import StringVar
from utils.variaveis_json import TEMA_PROGRAMA
from funcoes_interface import adicionar_pedido_avulso, enviar_relatorio_email, gerar_dicionario_dados_vendas

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

        # Passamos o self.local para atribuir o local da planilha la na variável para usar aqui na interface
        self.dados = gerar_dicionario_dados_vendas(
            var_local_planilha=self.local_planilha_meli
        )

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
        self.gerar_relatorio = ttk.Button(root, text="Enviar Email", style='success-outline', command=lambda: enviar_relatorio_email(
            local_planilha_meli=self.local_planilha_meli,
            lista_confirmados=self.lista_confirmados_gerar_relatorio,
            lista_pendentes=self.lista_pendentes(),
            observacao_email=self.entrada_observacao_email.get().strip()
        ))

        self.gerar_relatorio.place(x=1086, y=445)

        self.label_mensagem = ttk.Label(root, text="", font=("Segoe UI", 10), foreground="blue")
        self.label_mensagem.place(x=390, y=446)

        self.label_log = ttk.Label(root, text="LOG", font=("Segoe UI", 15), foreground="pink")
        self.label_log.place(x=20, y=520)

        self.botao_pedido_avulso = ttk.Button(root, text="Adicionar", style="success", command=lambda: adicionar_pedido_avulso(
            dados=self.dados,
            funcao_atualizar_lista=self.atualizar_lista
        ))
        self.botao_pedido_avulso.place(x=20, y=20)
        self.atualizar_lista()

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

                winsound.PlaySound(rf'audios\confirmar_pedido.wav', winsound.SND_FILENAME | winsound.SND_ASYNC)
      
            else:
                self.mensagem(f"⚠️ {codigo} Já conferido", "yellow")
                winsound.PlaySound(rf'audios\pedido_repetido.wav', winsound.SND_FILENAME | winsound.SND_ASYNC)
          
        else:
            self.mensagem(f"❌ {codigo} não encontrado.", "red")
            winsound.PlaySound(rf'audios\nao_encontrado.wav', winsound.SND_FILENAME | winsound.SND_ASYNC)
        

        self.codigo_var.set("")
        self.atualizar_lista()

        if all(info["status_conferencia"] for info in self.dados.values()):
            self.mensagem("✅ Todos os itens foram conferidos com sucesso!", "green")

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
        
if __name__ == "__main__":
    root = ttk.Window(themename=TEMA_PROGRAMA)
    app = ConferenciaApp(root)
    root.mainloop()

    
