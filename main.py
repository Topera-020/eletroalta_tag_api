import pandas as pd
import os
import json

def excel_para_json(arquivo_excel, diretorio_saida, colunas_para_manter):
    """
    Lê um arquivo Excel, converte cada aba em um JSON e salva no diretório de saída.
    
    :param arquivo_excel: Caminho do arquivo Excel a ser processado.
    :param diretorio_saida: Diretório onde os arquivos JSON serão salvos.
    :param colunas_para_manter: Dicionário com os nomes das abas como chave e uma lista de colunas para manter como valor.
    """
    # Garante que o diretório de saída exista
    os.makedirs(diretorio_saida, exist_ok=True)

    # Lê o arquivo Excel com todas as abas
    planilhas = pd.read_excel(arquivo_excel, sheet_name=None)
    
    for nome_aba, df in planilhas.items():
        print(f"Processando aba: {nome_aba}")
        
        # Verifica se há configuração de colunas para a aba
        colunas = colunas_para_manter.get(nome_aba, None)
        if colunas is None:
            print(f"Atenção: As colunas para a aba {nome_aba} não foram configuradas.")
        else:
            # Filtra apenas as colunas configuradas que existem na aba
            colunas_existentes = [col for col in colunas if col in df.columns]
            colunas_faltando = [col for col in colunas if col not in df.columns]
            
            if colunas_faltando:
                print(f"Atenção: As colunas {colunas_faltando} estão ausentes na aba {nome_aba} e serão ignoradas.")
            
            # Filtra as colunas existentes
            df = df[colunas_existentes]
            
            # Remove linhas onde 'Ativo' é False, se a coluna existir
            if 'Ativo' in df.columns:
                df = df[df['Ativo'] == True].drop(columns=['Ativo'])
            
            # Remove linhas vazias, se houver
            df = df.dropna(how="all")
            
            if not df.empty:
                # Gera o caminho para o arquivo JSON
                caminho_json = os.path.join(diretorio_saida, f"{nome_aba}.json")
                
                # Salva o DataFrame como JSON
                df.to_json(caminho_json, orient="records", indent=4, force_ascii=False)
                print(f"Aba {nome_aba} salva como JSON em {caminho_json}.")
            else:
                print(f"Aba {nome_aba} está vazia após o filtro e não será salva.")

# Exemplo de uso
if __name__ == "__main__":
    arquivo_excel = "tabelas/NCs.xlsx"  # Caminho do arquivo Excel
    diretorio_saida = "jsons"  # Diretório de saída dos JSONs
    
    # Configuração das colunas que você quer manter em cada aba
    colunas_para_manter = {

        "NCs": ["id", "Ativo", "categoriaId", "subcategoriaId", "titulo", 
                "descricao", "baseTecnica", "baseLegal", "infracao", 
                "recomendacoes", "Nota"],
        "Subcategorias": ["ID", "Subcategorias"],
        "Categorias": ["ID", "Categorias"],
    }
    
    # Executa a função
    excel_para_json(arquivo_excel, diretorio_saida, colunas_para_manter)
