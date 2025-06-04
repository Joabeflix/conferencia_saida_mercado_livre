import json


with open(r'configuracoes\configuracoes.json', 'r', encoding='utf-8') as f:
    dados = json.load(f)

print('leitura dados json')

COLUNA_COD_RASTREIO=dados["coluna_cod_rastreio"]
COLUNA_NOME_CLIENTE=dados["coluna_nome_cliente"]
TEMA_PROGRAMA=dados["tema"]

REMETENTE=dados["remetente"]
SENHA_DE_APP_EMAIL=dados["senha_de_app_email"]
DESTINATARIOS=dados["destinatarios"]