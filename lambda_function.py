import xml.etree.ElementTree as ET
import json
from xml_word_tools import add_page_break, get_root, get_word_namespaces, insertParagraf, replace_document_xml
from image_to_word import insert_images_to_word

def processar_arquivo_json(json_file_path, new_document_xml, docx_template_path):
    try:
        # Ler o conteúdo do arquivo
        with open(json_file_path, "r", encoding="utf-8") as file:
            raw_json = json.load(file)

        data = raw_json["data"]
        # Chamar a função convert_json_to_word_xml
        xml = convert_json_to_word_xml(data)
        print("XML gerado:")
        
        replace_document_xml(docx_template_path, xml, new_document_xml)
        insert_images_to_word(json_file_path, new_document_xml)
    
    except FileNotFoundError:
        print(f"Erro: O arquivo {json_file_path} não foi encontrado.")
    except json.JSONDecodeError:
        print("Erro: O arquivo JSON está mal formatado.")
    except Exception as e:
        print(f"Erro inesperado: {e}")


def convert_json_to_word_xml(data):
    # Definir os namespaces e criar o elemento raiz com os namespaces
    namespaces = get_word_namespaces()
    # Criar o elemento raiz <w:document> com os namespaces
    root = get_root(namespaces)

    print(f"Não conformidades: {len(data)}")
    
    # Ignorar a chave "ID"
    if 'ID' in data[0]:
        for item in data:
            item.pop('ID', None)

    # Criar o corpo do documento
    body = ET.SubElement(root, f"{{{namespaces['w']}}}body")
    
    prev_titulo1 = None
    prev_titulo2 = None
    prev_titulo2_aux = False
    
    # Iterar sobre os itens do JSON
    for item in data:
        prev_titulo2_aux = False
        for key, value in item.items():
            tag = translateKey(key)
            
            if tag == '':
               continue 
            if tag == "Fotos":
                #listImagesstr = ''
                for image in value:
                    #listImagesstr += f"{image} "
                    print('inserindo placeholder fotos _imagem_')
                    insertParagraf(image, namespaces, body, tag)

            else:
                if tag is None:
                    continue
                # Condição para não repetir Título1 ou Título2
                if tag == "Ttulo1" and value == prev_titulo1:
                    prev_titulo2_aux = True
                    continue
                    
                if tag == "Ttulo2" and value == prev_titulo2 and prev_titulo2_aux:
                    continue

                # Atualizar os valores anteriores
                if tag == "Ttulo1":
                    prev_titulo1 = value
                if tag == "Ttulo2":
                    prev_titulo2 = value
                    value = value
                if value:
                    insertParagraf(value, namespaces, body, tag)
        
        # Adicionar quebra de página após cada item
        add_page_break(body, namespaces)

    document_xml = ET.tostring(root, encoding='utf-8', method='xml').decode('utf-8')
    return document_xml


def translateKey(input_string):
    translation_dict = {
        #"ID": "ID",
        "categoria": "Ttulo1",
        "subcategoria": "Ttulo2",
        "titulo": "01NC",
        "setor": "02Setor",
        "local": "03Local",
        'fotos': "Fotos",
        "descricao": "04Descrio",
        "baseTecnica": "05BaseTcnica",
        "baseLegal": "06BaseLegal",
        "infracao": "07InfraoConforme",
        "recomendacao": "08Recomendaes",
        "nota": "09Nota",
        "id":""
        #"Observação": "Observação",
    }
    if input_string not in translation_dict:
        print(f"'{input_string}' não encontrado na tradução. Retornando valor padrão.")
        return ''
    return translation_dict.get(input_string)
    


           
def filtrar_imagens(json_data):
    # Lista para armazenar arrays de imagens
    lista_de_arrays = []

    for item in json_data:
        for key, value in item.items():
            if key == 'Fotos':
                print(key)
                lista_de_arrays.append(value)
    return lista_de_arrays