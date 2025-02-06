import sys
from flask import Flask, request, jsonify, send_from_directory,  send_file
import json
import os

from excell_para_word import convert_excel_to_word_xml
from json_to_word import processar_arquivo_json

app = Flask(__name__)

# Definindo o diretório base
if getattr(sys, 'frozen', False):
    # Quando o código estiver rodando como .exe
    base_path = sys._MEIPASS
else:
    # Quando o código estiver rodando em ambiente de desenvolvimento
    base_path = os.path.abspath(".")

# Definir o caminho das pastas relativas ao diretório base
UPLOAD_FOLDER = os.path.join(base_path, "uploads")
OUTPUT_FOLDER = os.path.join(base_path, "outputs")

# Verificar e criar as pastas caso não existam
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)
    
docx_template_path = os.path.join(base_path, 'static', 'Banco de Dados.docx')
output_docx = 'Resultado.docx'
output_docx_path = os.path.join(OUTPUT_FOLDER, output_docx)


@app.route('/')
def index():
    return send_from_directory(os.path.join(base_path, 'static'), 'index.html')


@app.route("/conect", methods=["GET"])
def conect():
    print("conect")
    return  jsonify({"conect success": True})


@app.route("/upload_json", methods=["POST"])
def upload_json():
    print("Started")
    if not request.is_json:
        return jsonify({"error": "O conteúdo da requisição não é um JSON válido"}), 400

    json_file_path = os.path.join(UPLOAD_FOLDER, 'data.json')
    try:
        data = request.get_json()
        with open(json_file_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        
        processar_arquivo_json(json_file_path, output_docx_path, docx_template_path)

        
        # Retornar o arquivo como resposta para o Flutter
        return send_file(output_docx_path, as_attachment=True, download_name="output.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        #return jsonify({"Sucesso na criacaoo do RTI": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500







@app.route('/upload', methods=['POST'])
def upload_file():
    if 'excel_file' not in request.files:
        print("Erro: Nenhum arquivo enviado.")
        return "No file part", 400
    
    excel_file = request.files['excel_file']
    if excel_file.filename == '':
        print("Erro: Nenhum arquivo selecionado.")
        return "No selected file", 400
    
    excel_file_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
    excel_file.save(excel_file_path)
    print(f"Arquivo Excel salvo em: {excel_file_path}")
    
    try:
        # Converter Excel para XML do Word
        print("Iniciando a conversão do Excel para XML do Word...")
        convert_excel_to_word_xml(excel_file_path, docx_template_path,  output_docx_path)
        print("Conversão do Excel para XML do Word concluída.")
        os.remove(excel_file_path)
        print(f"Arquivo Word gerado com sucesso em: {output_docx_path}")
        
        return send_from_directory(OUTPUT_FOLDER, output_docx, as_attachment=True)
    
    except Exception as e:
        print(f"Erro durante o processamento do arquivo: {e}")
        return f"Erro ao processar o arquivo: {e}", 500



@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    print(f"Iniciando download do arquivo {filename}")
    return send_from_directory(OUTPUT_FOLDER, filename)




if __name__ == "__main__":
     # Exemplo de uso
    app.run(host="0.0.0.0", port=5000)
    
    
    
    # venv\Scripts\Activate.ps1
    