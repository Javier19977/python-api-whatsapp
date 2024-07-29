import requests
import time
import sys
from datetime import datetime

def obtener_fecha_hora_actual():
    """Obtiene la fecha y hora actual en formato legible."""
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def sendMessage(para, mensaje):
    url = 'http://localhost:3001/lead'  # URL de tu servidor local para la API de WhatsApp
    data = {
        "message": mensaje,
        "phone": para
    }
    headers = {
        'Content-Type': 'application/json'
    }

    max_intentos = 3
    intentos = 0
    mensaje_enviado = False

    while intentos < max_intentos:
        intentos += 1
        if intentos > 1:
            sys.stdout.write(f"\r🔄 Reintentando... ({intentos}/{max_intentos})")
            sys.stdout.flush()
        else:
            print(f"📤 [{obtener_fecha_hora_actual()}] Enviando mensaje a {para}...")

        try:
            response = requests.post(url, json=data, headers=headers, timeout=10)
            response.raise_for_status()  # Lanza una excepción para errores HTTP

            if response.status_code == 404:
                sys.stdout.write(f"\r🔍 [{obtener_fecha_hora_actual()}] El número {para} no tiene WhatsApp. Pasando al siguiente número.\n")
                sys.stdout.flush()
                return False  # Indicar que no se pudo enviar el mensaje
            else:
                print(f"✅ [{obtener_fecha_hora_actual()}] Mensaje enviado con éxito a {para}.")
                mensaje_enviado = True
                break  # Salir del bucle si el mensaje fue enviado con éxito

        except requests.exceptions.Timeout:
            if intentos == max_intentos:
                # Mensaje final después de todos los intentos fallidos
                sys.stdout.write(f"\r❌ [{obtener_fecha_hora_actual()}] El mensaje a {para} no pudo ser enviado después de {max_intentos} intentos. El número de teléfono no posee WhatsApp.\n")
                sys.stdout.flush()
            else:
                sys.stdout.write(f"\r🔄 Reintentando... ({intentos}/{max_intentos})")
                sys.stdout.flush()
            time.sleep(5)  # Esperar antes de reintentar
        except requests.exceptions.RequestException as req_err:
            print(f"🚨 [{obtener_fecha_hora_actual()}] Error de solicitud al enviar mensaje a {para}: {req_err}")
            break  # No continuar intentando en caso de otros errores de solicitud
        except Exception as e:
            print(f"⚠️ [{obtener_fecha_hora_actual()}] Error inesperado al enviar mensaje a {para}: {e}")
            break  # No continuar intentando en caso de errores inesperados

    return mensaje_enviado  # Indicar si el mensaje fue enviado con éxito o no
