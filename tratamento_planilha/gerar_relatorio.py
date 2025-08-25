import pandas as pd
from openpyxl import load_workbook

def gerar_relatorio_exel(caminho_arquivo: str, nome_aba: str, dados: list[list]) -> None:
    """
    Abre uma planilha existente, cria uma nova aba e insere os dados.
    
    Parâmetros:
        caminho_arquivo (str): Caminho do arquivo Excel (.xlsx).
        nome_aba (str): Nome da nova aba que será criada.
        dados (list): Lista de listas com os dados. Ex: [['oi', 123], ['ola', 321]]
    """
    # Cria DataFrame com os dados
    df = pd.DataFrame(dados, columns=["Código Rastreio", "Nome Cliente"])

    # Abre a planilha existente com openpyxl (engine necessário para múltiplas abas)
    with pd.ExcelWriter(caminho_arquivo, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
        df.to_excel(writer, sheet_name=nome_aba, index=False)


