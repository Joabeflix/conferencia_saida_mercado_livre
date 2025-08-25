import pandas as pd
from utils.utils import texto_no_console
from tkinter import filedialog, Tk

class TratamentoPlanilhaMercadoLivre:
    def __init__(self, nome_aba_cod_rastreiro: str, nome_aba_nome_cliente: str):
        self.nome_aba_cod_rastreiro = nome_aba_cod_rastreiro
        self.nome_aba_nome_cliente = nome_aba_nome_cliente
        self.local_planilha: str | None = None

        try:
            self.definir_local_planilha()
            print(f'Planilha depois de definir: {self.local_planilha}')

            planilha = pd.read_excel(self.local_planilha, header=5)
            self.lista_cod_rastreio, self.lista_nome_cliente = list(planilha[self.nome_aba_cod_rastreiro]), list(planilha[self.nome_aba_nome_cliente])

        except FileNotFoundError:
            texto_no_console('Erro ao ler a planilha. O arquivo não foi encontrado. Verifique o local do arquivo.')
            return None
        


    def _criar_dicionario(self):
        
        dicionario = {}

        for cod_rastreio, nome_cliente in zip(self.lista_cod_rastreio, self.lista_nome_cliente):
            if cod_rastreio == ' ':
                continue
            dicionario[cod_rastreio[3:14]] = {'nome_cliente': nome_cliente, 'status_conferencia': False}
        return dicionario
    
    def retornar_local_planilha(self) -> str:
        return f'{self.local_planilha}'
    
    def definir_local_planilha(self):
        root = Tk()
        root.withdraw()
        arquivo_xlsx = filedialog.askopenfilename(
            title="Selecione a planilha de vendas do Mercado Livre.xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )

        if arquivo_xlsx:
            self.local_planilha = arquivo_xlsx
        else:
            raise FileNotFoundError("Nenhum arquivo .xlsx foi selecionado ou encontrado.")
