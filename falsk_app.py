from json_to_word import convert_json_to_word_xml, inserir_imagens_no_word, replace_document_xml

from flask import Flask, request, send_file, jsonify
import os

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/upload-json', methods=['POST'])
def upload_json():
    # Verificar se a requisição contém JSON
    if not request.is_json:
        return jsonify({"error": "O conteúdo da requisição não é um JSON válido"}), 400

    # Obter o conteúdo JSON da requisição
    data = request.get_json()

    
    # Exemplo de uso das informações do JSON para gerar o arquivo Word
    docx_template_path = 'tamplate.docx'
    output_docx_path = 'Resultado\\Resultado.docx'

    # Aqui você pode chamar suas funções para gerar o documento Word usando os dados do JSON
    new_document_xml = convert_json_to_word_xml(data)
    replace_document_xml(docx_template_path, new_document_xml, output_docx_path)
    # Função fictícia para inserir imagens (dependendo da lógica que você já criou)
    inserir_imagens_no_word(output_docx_path, output_docx_path, data)

    # Retornar o arquivo Word gerado
    return send_file(output_docx_path, as_attachment=True, attachment_filename='Relatorio.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


if __name__ == '__main__':
    app.run(debug=True)
