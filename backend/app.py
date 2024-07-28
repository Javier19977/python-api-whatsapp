import time
from flask import Flask, request, jsonify, render_template
import pandas as pd
import requests
from flask_cors import CORS
from datetime import datetime

app = Flask(__name__)
CORS(app)

WHATSAPP_API_URL = 'https://tu-api-de-whatsapp-en-render.com/lead'  # Cambia esta URL por la correcta

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and file.filename.endswith('.xlsx'):
        try:
            df = pd.read_excel(file)
            has_invalid_data = False

            for index, row in df.iterrows():
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

                send_message(numero, mensaje)

            return jsonify({'message': 'Archivo procesado exitosamente'}), 200
        except Exception as e:
            return jsonify({'error': f'Error al procesar el archivo: {str(e)}'}), 500
    else:
        return jsonify({'error': 'Formato de archivo no válido'}), 400

@app.route('/lead', methods=['POST'])
def lead():
    try:
        data = request.json
        message = data.get('message')
        phone = data.get('phone')

        if not message or not phone:
            return jsonify({'error': 'Faltan datos'}), 400

        response = requests.post(WHATSAPP_API_URL, json={'message': message, 'phone': phone})

        if response.status_code == 200:
            return jsonify({'success': 'Mensaje enviado con éxito'}), 200
        else:
            return jsonify({'error': f'Error en la solicitud: {response.status_code}'}), response.status_code
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def send_message(para, mensaje):
    data = {
        "message": mensaje,
        "phone": para
    }
    headers = {
        'Content-Type': 'application/json'
    }

    max_intentos = 3
    intentos = 0

    while intentos < max_intentos:
        intentos += 1
        try:
            response = requests.post(WHATSAPP_API_URL, json=data, headers=headers, timeout=10)
            response.raise_for_status()

            if response.status_code == 404:
                print(f"🔍 [{obtener_fecha_hora_actual()}] El número {para} no tiene WhatsApp. Pasando al siguiente número.")
                return False
            else:
                print(f"✅ [{obtener_fecha_hora_actual()}] Mensaje enviado con éxito a {para}.")
                break

        except requests.exceptions.Timeout:
            if intentos == max_intentos:
                print(f"❌ [{obtener_fecha_hora_actual()}] El mensaje a {para} no pudo ser enviado después de {max_intentos} intentos. El número de teléfono no posee WhatsApp.")
            else:
                print(f"🔄 Reintentando... ({intentos}/{max_intentos})")
            time.sleep(5)
        except requests.exceptions.RequestException as req_err:
            print(f"🚨 [{obtener_fecha_hora_actual()}] Error de solicitud al enviar mensaje a {para}: {req_err}")
            break
        except Exception as e:
            print(f"⚠️ [{obtener_fecha_hora_actual()}] Error inesperado al enviar mensaje a {para}: {e}")
            break

    return False

def obtener_fecha_hora_actual():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
