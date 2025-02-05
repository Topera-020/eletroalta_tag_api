document.getElementById('upload-form-rti').addEventListener('submit', function(event) {
    event.preventDefault();  // Evita o envio padrão do formulário

    // Obtenha o arquivo do input
    var fileInput = document.getElementById('file-rti');
    var file = fileInput.files[0];

    if (!file) {
        alert("Por favor, selecione um arquivo.");
        return;
    }

    // Verifique o tipo do arquivo
    if (!['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'].includes(file.type)) {
        alert("Por favor, envie um arquivo válido do Excel (.xlsx ou .xls).");
        return;
    }

    // Mensagem de feedback
    document.getElementById('response').innerText = 'Processando...';

    // Crie um objeto FormData para enviar o arquivo
    var formData = new FormData();
    formData.append('excel_file', file);

    // Envie o arquivo para o servidor usando Fetch API
    fetch('/upload', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (!response.ok) {
            throw new Error('Erro no servidor: ' + response.statusText);
        }
        return response.blob();
    })
    .then(blob => {
        // Crie um link temporário para o download
        var link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'Resultado.docx';  // Nome do arquivo para o download
        link.click();
        URL.revokeObjectURL(link.href);
        document.getElementById('response').innerText = 'Download concluído!';
    })
    .catch(error => {
        document.getElementById('response').innerText = 'Erro ao processar o arquivo: ' + error.message;
    });
});


document.addEventListener('DOMContentLoaded', function() {
    // Adiciona o evento de clique ao botão de download
    document.getElementById('download-rti').addEventListener('click', function(event) {
        event.preventDefault();
        // Cria um link temporário para o download do arquivo
        const link = document.createElement('a');
        link.href = 'static/Tabela Padrao.xlsx';
        link.download = 'Tabela padrao.xlsx';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });
});
