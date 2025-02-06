import pandas as pd
import xml.etree.ElementTree as ET
from xml_word_tools import add_page_break, get_root, get_word_namespaces, insertParagraf, replace_document_xml


def convert_excel_to_word_xml(excel_file_path, docx_template_path, output_docx_path):
    # Definir os namespaces e criar o elemento raiz com os namespaces
    

    df = pd.read_excel(excel_file_path)
    df = df.sort_values(by=['Categoria', 'Subcategoria'])
    
    root = create_word_body_from_dataframe(df)
    
    new_document_xml = ET.tostring(root, encoding='utf-8', method='xml').decode('utf-8')
    
    replace_document_xml(docx_template_path, new_document_xml, output_docx_path)


def translate_string(input_string):
    # Função para tradução entre os estilos e os nomes das colunas
    # Na primeira parte do dicionário, a chave é o nome da coluna no DataFrame
    # Na segunda parte do dicionário, o valor é o nome do estilo no Word
    # Caso o estilo não seja encontrado, a função retorna None
    # Os estilos que não forem encontrados serão ignorados
    # Deixei comentado as colunas que não serão utilizadas
    translation_dict = {
        "ID": None,
        "Categoria": "Ttulo1",
        "Subcategoria": "Ttulo2",
        "NC": "01NC",
        "Setor": "02Setor",
        "Local": "03Local",
        "Descrição": "04Descrio",
        "Base Técnica": "05BaseTcnica",
        "Base Legal": "06BaseLegal",
        "Infração Conforme NR28": "07InfraoConforme",
        "Recomendações": "08Recomendaes",
        "Nota": "09Nota",
        #"Observação": "Observação",
    }
    if input_string not in translation_dict:
        print(f"'{input_string}' não encontrado na tradução. Retornando valor padrão.")
        return None
    return translation_dict.get(input_string)


def create_word_body_from_dataframe(df):
    namespaces = get_word_namespaces()
    
    # Criar o elemento raiz <w:document> com os namespaces
    root = get_root(namespaces)
    
    # Criar o corpo do documento
    body = ET.SubElement(root, f"{{{namespaces['w']}}}body")

    # Inicializar variáveis para armazenar os valores anteriores
    prev_titulo1 = None
    prev_titulo2 = None
    # Variável auxiliar para repetir Título2 caso Título1 não seja repetido
    # True se Título1 da linha for repetido, False caso contrário
    prev_titulo2_aux = False

    # Iterar sobre as linhas do DataFrame
    for idx, row in df.iterrows():
        
        # Resetar a variável auxiliar para cada linha nova
        prev_titulo2_aux = False
        
        # Iterar sobre as colunas do DataFrame
        for style in df.columns:
            text = row[style]
            translated_style = translate_string(style)
            
            # Se o estilo não for encontrado, pular para a próxima iteração
            if translated_style is None:
                continue

            # Condição para não repetir Título1 ou Título2
            if translated_style == "Ttulo1" and text == prev_titulo1:
                prev_titulo2_aux = True
                continue
                
            if translated_style == "Ttulo2" and text == prev_titulo2 and prev_titulo2_aux:
                continue

            # Atualizar os valores anteriores
            if translated_style == "Ttulo1":
                prev_titulo1 = text
            if translated_style == "Ttulo2":
                prev_titulo2 = text

            # Adicionar o texto ao corpo do documento
            insertParagraf(text, namespaces, body, translated_style)

        # Adicionar quebra de página após cada linha
        if len(row) > 0:
            add_page_break(body, namespaces)

    return root
