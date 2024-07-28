# backend/app.py

from flask import Flask, request, jsonify
from flask_cors import CORS
import openpyxl
import requests

app = Flask(__name__)
CORS(app)

@app.route('/lead', methods=['POST'])
def send_message():
    data = request.get_json()
    message = data.get('message')
    phone = data.get('phone')

    if not message or not phone:
        return jsonify({"status": "error", "message": "Faltan parámetros 'message' o 'phone'."}), 400

    url = "http://localhost:3004/lead"
    payload = {
        "message": message,
        "phone": phone
    }
    headers = {
        "Content-Type": "application/json"
    }

    try:
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        return jsonify({"status": "success", "message": "Mensaje enviado con éxito."}), 200
    except requests.exceptions.RequestException as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"status": "error", "message": "No se ha enviado ningún archivo."}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"status": "error", "message": "El nombre del archivo está vacío."}), 400

    try:
        wb = openpyxl.load_workbook(file)
        sheet = wb.active
        messages_sent = []

        for row in sheet.iter_rows(min_row=2, values_only=True):
            phone = row[0]
            message = row[1]

            if not phone or not message:
                continue

            url = "http://localhost:3004/lead"
            payload = {
                "message": message,
                "phone": phone
            }
            headers = {
                "Content-Type": "application/json"
            }

            try:
                response = requests.post(url, json=payload, headers=headers)
                response.raise_for_status()
                messages_sent.append({"phone": phone, "status": "success"})
            except requests.exceptions.RequestException:
                messages_sent.append({"phone": phone, "status": "error"})

        return jsonify({"status": "success", "messages_sent": messages_sent}), 200

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
