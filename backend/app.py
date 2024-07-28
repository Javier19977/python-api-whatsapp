from flask import Flask, request, jsonify, render_template
import pandas as pd
import os
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

@app.route('/')
def index():
    return """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Enviar Mensajes Personalizados</title>
    </head>
    <body>
        <h1>Enviar Mensajes Personalizados</h1>
        <form id="uploadForm">
            <label for="fileInput">Selecciona el archivo Excel:</label>
            <input type="file" id="fileInput" name="file">
            <button type="submit">Enviar</button>
        </form>
        <div id="result"></div>

        <script>
            document.getElementById('uploadForm').addEventListener('submit', async (event) => {
                event.preventDefault();
                const fileInput = document.getElementById('fileInput');
                if (!fileInput.files.length) {
                    alert('Por favor selecciona un archivo.');
                    return;
                }

                const formData = new FormData();
                formData.append('file', fileInput.files[0]);

                try {
                    const response = await fetch('/upload', {
                        method: 'POST',
                        body: formData,
                    });

                    if (!response.ok) {
                        throw new Error('Error en la solicitud: ' + response.statusText);
                    }

                    const result = await response.json();
                    document.getElementById('result').textContent = JSON.stringify(result, null, 2);
                } catch (error) {
                    document.getElementById('result').textContent = 'Error: ' + error.message;
                }
            });
        </script>
    </body>
    </html>
    """

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        print("Recibida solicitud de subida de archivo")
        
        if 'file' not in request.files:
            error_message = 'No file part in the request'
            print(error_message)
            return jsonify({'error': error_message}), 400
        
        file = request.files['file']
        if file.filename == '':
            error_message = 'No file selected'
            print(error_message)
            return jsonify({'error': error_message}), 400
        
        if file and file.filename.endswith('.xlsx'):
            print(f"Archivo recibido: {file.filename}")
            
            try:
                df = pd.read_excel(file)
                print("Archivo Excel leído exitosamente")
                
                # Aquí puedes agregar la lógica para procesar el archivo y enviar mensajes
                return jsonify({'message': 'Archivo procesado exitosamente'}), 200
            except Exception as e:
                error_message = f"Error al leer el archivo Excel: {str(e)}"
                print(error_message)
                return jsonify({'error': error_message}), 500
        else:
            error_message = 'Invalid file format'
            print(error_message)
            return jsonify({'error': error_message}), 400
    except Exception as e:
        error_message = f"Error desconocido: {str(e)}"
        print(error_message)
        return jsonify({'error': error_message}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
