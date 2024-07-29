# app.py
from flask import Flask, request, jsonify
import openpyxl
import os
from your_script import sendMessage, obtener_fecha_hora_actual

app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if not file.filename.endswith('.xlsx'):
        return jsonify({'error': 'Invalid file type. Only .xlsx files are allowed.'}), 400

    try:
        wb = openpyxl.load_workbook(file)
        sheet = wb.active
        has_invalid_data = False

        # Process the Excel file
        for idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            numero = str(row[0]).strip().replace('-', '') if row[0] else ''
            numero_ticket = str(row[1]).strip() if row[1] else ''
            nombre_tecnico = str(row[2]).strip() if row[2] else ''
            nombre_estudiante = str(row[3]).strip() if row[3] else ''
            falla_reportada = str(row[4]).strip() if row[4] else ''
            nie = str(row[5]).strip() if row[5] else ''
            serie = str(row[6]).strip() if row[6] else ''

            if not all([numero, numero_ticket, nombre_estudiante, nombre_tecnico, falla_reportada, nie, serie]):
                has_invalid_data = True
                continue

            if len(numero) == 8:
                numero = '503' + numero

            mensaje = (
               f"🛠️ *Atención de soporte* - Ticket *{numero_ticket}*\n"
                f"\n📋 Datos del ESTUDIANTE: *{nombre_estudiante}* - NIE: *{nie}*\n"
                f"\n🔧 El equipo se recibió con la siguiente falla: *{falla_reportada}*. Serie del equipo: *{serie}*.\n"
                f"\n📦 Ya está listo para ser recogido en nuestra sede de Soporte Técnico en Santa Ana.\n"
                f"\n📲 Por favor, asegúrate de traer este mensaje y tu DUI al llegar para agilizar el proceso.\n"
                f"\n🔌 Recuerda traer el cargador del equipo, a menos que ya lo hayas entregado al técnico.\n"
                f"\n⏰ *Horario de atención*: Lunes a Viernes, de 7:30 am a 12:00 pm y de 1:00 pm a 3:30 pm.\n"
                f"\n🛠️ El equipo estará disponible para ser retirado por un plazo de *15 días hábiles* a partir de esta notificación. Posteriormente, será trasladado a nuestra bodega para atender otros casos, por lo que te recomendamos recogerlo lo antes posible para evitar inconvenientes.\n"
                f"\n📍 *Dirección*: Estamos ubicados a solo una cuadra del ISSS Santa Ana, en la Avenida California, colonia El Palmar, dentro de CE INSA Industrial. [Ver ubicación](https://maps.app.goo.gl/9G4GNo8RE63VKH6r6)\n"
                f"\n🧑‍🔧 Técnico a cargo: *{nombre_tecnico}*"
            )

            sendMessage(numero, mensaje)

        return jsonify({'success': True, 'message': 'Mensajes enviados correctamente.'}), 200

    except Exception as e:
        return jsonify({'error': f'Error inesperado: {e}'}), 500

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=3001)
