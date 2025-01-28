import pandas as pd
import json
import os
import numpy as np

def xlsx_para_json(df, path_saida):
    

    # Cria um dicionário para armazenar os dados das abas
    dados = {}

    # Itera pelas abas do arquivo Excel
    for nome_aba, tabela in df.items():
        # Remove chaves com valores NaN
        tabela = tabela.where(pd.notna(tabela), None)
        tabela = tabela.dropna(axis=1, how='all')  # Remove colunas inteiras com NaN
        tabela_dict = tabela.to_dict(orient='records')
        tabela_dict = [{k: v for k, v in d.items() if (pd.notna(v) and v is not None)} for d in tabela_dict]

        dados["data"] = tabela_dict

    # Salva os dados no formato JSON
    os.makedirs(os.path.dirname(path_saida), exist_ok=True)
    with open(path_saida, 'w', encoding='utf-8') as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)

    print(f"Arquivo JSON salvo em: {path_saida}")

