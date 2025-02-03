import sys
from flask import Flask, request, jsonify
import json
import os

from lambda_function import processar_arquivo_json

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


@app.route("/conect", methods=["GET"])
def conect():
    print("conect")
    return  jsonify({"conect success": True})


@app.route("/upload", methods=["POST"])
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

        #processar_arquivo_json(json_file_path, output_docx_path, docx_template_path)
        return jsonify({"Sucesso na criacaoo do RTI": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
     # Exemplo de uso
    app.run(debug=True, host="0.0.0.0", port=5000)
    