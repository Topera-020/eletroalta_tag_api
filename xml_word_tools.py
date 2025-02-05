import xml.etree.ElementTree as ET
import shutil
import zipfile
import os

def get_word_namespaces():
    return {
        'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
        'cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
        'cx1': 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex',
        'cx2': 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex',
        'cx3': 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex',
        'cx4': 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex',
        'cx5': 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex',
        'cx6': 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex',
        'cx7': 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex',
        'cx8': 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex',
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'aink': 'http://schemas.microsoft.com/office/drawing/2016/ink',
        'am3d': 'http://schemas.microsoft.com/office/drawing/2017/model3d',
        'o': 'urn:schemas-microsoft-com:office:office',
        'oel': 'http://schemas.microsoft.com/office/2019/extlst',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
        'v': 'urn:schemas-microsoft-com:vml',
        'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'w10': 'urn:schemas-microsoft-com:office:word',
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
        'w16cex': 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
        'w16cid': 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
        'w16': 'http://schemas.microsoft.com/office/word/2018/wordml',
        'w16du': 'http://schemas.microsoft.com/office/word/2023/wordml/word16du'
    }
    
    
def get_root(namespaces):
    return ET.Element('w:document', attrib={
        'xmlns:wpc': namespaces['wpc'],
        'xmlns:cx': namespaces['cx'],
        'xmlns:cx1': namespaces['cx1'],
        'xmlns:cx2': namespaces['cx2'],
        'xmlns:cx3': namespaces['cx3'],
        'xmlns:cx4': namespaces['cx4'],
        'xmlns:cx5': namespaces['cx5'],
        'xmlns:cx6': namespaces['cx6'],
        'xmlns:cx7': namespaces['cx7'],
        'xmlns:cx8': namespaces['cx8'],
        'xmlns:mc': namespaces['mc'],
        'xmlns:aink': namespaces['aink'],
        'xmlns:am3d': namespaces['am3d'],
        'xmlns:o': namespaces['o'],
        'xmlns:oel': namespaces['oel'],
        'xmlns:r': namespaces['r'],
        'xmlns:m': namespaces['m'],
        'xmlns:v': namespaces['v'],
        'xmlns:wp14': namespaces['wp14'],
        'xmlns:wp': namespaces['wp'],
        'xmlns:w10': namespaces['w10'],
        'xmlns:w': namespaces['w'],
        'xmlns:w14': namespaces['w14'],
        'xmlns:w15': namespaces['w15'],
        'xmlns:w16cex': namespaces['w16cex'],
        'xmlns:w16cid': namespaces['w16cid'],
        'xmlns:w16': namespaces['w16'],
        'xmlns:w16du': namespaces['w16du'],
        #'xmlns:w16sdtdh': namespaces['w16sdtdh'],
        #'xmlns:w16se': namespaces['w16se'],
        #'xmlns:wpg': namespaces['wpg'],
        #'xmlns:wpi': namespaces['wpi'],
        #'xmlns:wne': namespaces['wne'],
        #'xmlns:wps': namespaces['wps'],
        'mc:Ignorable': 'w14 w15 w16cid w16 w16cex w16du wp14' #w16se w16sdtdh
    })


def add_page_break(body, namespaces):
    p = ET.SubElement(body, f"{{{namespaces['w']}}}p")
    r = ET.SubElement(p, f"{{{namespaces['w']}}}r")
    ET.SubElement(r, f"{{{namespaces['w']}}}br", {'w:type': 'page'})
    
def insertParagraf(text, namespaces, body, style ='a'):
    p = ET.SubElement(body, f"{{{namespaces['w']}}}p")
    pPr = ET.SubElement(p, f"{{{namespaces['w']}}}pPr")
    ET.SubElement(pPr, f"{{{namespaces['w']}}}pStyle", {'w:val': style})
    r = ET.SubElement(p, f"{{{namespaces['w']}}}r")
    ET.SubElement(r, f"{{{namespaces['w']}}}t").text = str(text)
  


def replace_document_xml(docx_template_path, new_document_xml, output_docx_path):
    # Criar um diretório temporário
    temp_dir = 'temp_docx_dir'
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
    
    # Extrair o conteúdo do arquivo .docx para o diretório temporário
    with zipfile.ZipFile(docx_template_path, 'r') as docx:
        docx.extractall(temp_dir)
    
    # Substituir o arquivo document.xml no diretório temporário
    document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
    with open(document_xml_path, 'w', encoding='utf-8') as file:
        file.write(new_document_xml)
    
    # Criar um novo arquivo .docx com o conteúdo do diretório temporário
    with zipfile.ZipFile(output_docx_path, 'w', zipfile.ZIP_DEFLATED) as docx:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                # Calcula o caminho relativo para o arquivo dentro do ZIP
                relative_path = os.path.relpath(file_path, temp_dir)
                docx.write(file_path, relative_path)
    
    # Limpar o diretório temporário
    shutil.rmtree(temp_dir)
    
    print(f"O arquivo {output_docx_path} foi criado com sucesso.")

