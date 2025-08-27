import os
import tkinter as tk
from datetime import datetime
from servico_email.servico_de_email import Email
from tratamento_planilha.gerar_relatorio import gerar_relatorio_exel
from tratamento_planilha.planilha_romaneio import TratamentoPlanilhaMercadoLivre
from utils.variaveis_json import SENHA_DE_APP_EMAIL, REMETENTE, DESTINATARIOS, COLUNA_COD_RASTREIO, COLUNA_NOME_CLIENTE

def adicionar_pedido_avulso(dados: dict, funcao_atualizar_lista) -> None:
#    # dicionario[12345678910] = {'nome_cliente': nome_cliente, 'status_conferencia': False}
    def fechar(root: tk.Tk) -> None:
        print('tetstetet ')
        root.destroy()
    def confirmar() -> None:
        if cod_rastreio.get() and nome_cliente.get():
            dados[cod_rastreio.get()] = {'nome_cliente': nome_cliente.get(), 'status_conferencia': False}
            funcao_atualizar_lista()
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

def enviar_relatorio_email(local_planilha_meli, lista_confirmados, lista_pendentes, observacao_email) -> None:
    gerar_relatorio_exel(rf'{local_planilha_meli}', nome_aba='Confirmados', dados=lista_confirmados)
    email = Email(
        remetente=REMETENTE
    )
    email.definir_senha(SENHA_DE_APP_EMAIL)
    
    data_hora = datetime.now()
    data_hora_formatada = data_hora.strftime("%d-%m-%Y %H:%M:%S")

    mensagem = f"""
Lista de pedidos confirmados no Mercado Livre.\n
DATA E HORA DO ENVIO: {data_hora_formatada}\n
{f'Pedidos não confirmados:\n{lista_pendentes}' if lista_pendentes else ''}\n
{f'Observação:\n{observacao_email}' if observacao_email  else 'Sem observações.'}

------------------------------------------------------------
USUÁRIO: {os.environ['USERNAME']}
NOME DO COMPUTADOR: {os.environ['COMPUTERNAME']}
------------------------------------------------------------
    """

    email.enviar(
        assunto="Relatório pedidos confirmados",
        mensagem=mensagem,
        destinatarios=DESTINATARIOS,
        anexos=[rf'{local_planilha_meli}']
    )

def gerar_dicionario_dados_vendas(var_local_planilha) -> dict:
    app = TratamentoPlanilhaMercadoLivre(
        nome_aba_cod_rastreiro=COLUNA_COD_RASTREIO,
        nome_aba_nome_cliente=COLUNA_NOME_CLIENTE)
    dicionario = app.criar_dicionario()
    var_local_planilha = app.retornar_local_planilha()
    return dicionario



