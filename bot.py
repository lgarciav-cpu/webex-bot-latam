from flask import Flask, request
import os
import requests
import random
from datetime import datetime
import threading
from zoneinfo import ZoneInfo

# =========================
# CONFIGURACIÓN GENERAL
# =========================
WEBEX_TOKEN = os.environ.get("WEBEX_TOKEN", "ZmY2MGJlYWYtMzgxYy00ZDljLWIyMmYtMTZkYjRlMTc2N2EzNTYxYjk5ZjgtN2Iw_PF84_1eb65fdf-9643-417f-9974-ad72cae0e10f").strip()
WEBEX_API_MESSAGES = "https://webexapis.com/v1/messages"
WEBEX_API_MEETINGS = "https://webexapis.com/v1/meetings"  # API estándar para Reuniones y Webinars

TZ = ZoneInfo("America/Lima")
app = Flask(__name__)

# Headers reutilizables para llamadas a la API de Webex
def get_webex_headers():
    if not WEBEX_TOKEN:
        raise RuntimeError("Falta WEBEX_TOKEN en las variables de entorno.")
    return {
        "Authorization": f"Bearer {WEBEX_TOKEN}",
        "Content-Type": "application/json"
    }

def send_message(room_id: str, text: str):
    """Envía un mensaje de texto plano o markdown a un room de Webex."""
    data = {"roomId": room_id, "markdown": text}
    try:
        requests.post(WEBEX_API_MESSAGES, headers=get_webex_headers(), json=data, timeout=15)
    except Exception as e:
        print(f"❌ Error al enviar mensaje: {e}")

# ==========================================
# MENÚ Y CONTROLADOR DE OPCIONES
# ==========================================
def mostrar_menu_principal(room_id):
    """Despliega el menú principal de opciones."""
    menu_texto = (
        "🤖 **¡Hola! Bienvenido al Bot de Gestión de Webex**\n\n"
        "Selecciona o escribe una de las siguientes opciones:\n\n"
        "1️⃣ **Crear Webinar / Sesión Webex**\n"
        "2️⃣ **Próximamente: Consultas / Reportes**\n"
        "3️⃣ **Próximamente: Otra función**\n\n"
        "--- \n"
        "💡 *Para ejecutar una opción, responde con el número o sigue las instrucciones del comando.*"
    )
    send_message(room_id, menu_texto)

def mostrar_instrucciones_opcion_1(room_id):
    """Instrucciones específicas para la creación de Webinars."""
    instrucciones = (
        "📌 **OPCIÓN 1: Crear Webinar**\n\n"
        "Para programar una sesión, envía el siguiente formato exacto:\n\n"
        "```text\n"
        "crear webinar;\n"
        "Título: Lanzamiento de Producto 2026;\n"
        "Fecha: 2026-08-15 15:00;\n"
        "Duración: 60;\n"
        "Panelistas: juan@empresa.com, maria@empresa.com\n"
        "```\n"
        "⚠️ *Nota: La fecha debe estar en formato `YYYY-MM-DD HH:MM` (Hora local).* "
    )
    send_message(room_id, instrucciones)

# ==========================================
# LÓGICA DE WEBINARS EN WEBEX
# ==========================================
def crear_webinar_api(titulo, fecha_inicio_dt, duracion_minutos, panelistas_emails):
    """
    Crea una sesión de tipo Webinar en Cisco Webex.
    Ajusta scheduledType a 'webinar' si la cuenta lo permite, o 'scheduledMeeting' por defecto.
    """
    payload = {
        "title": titulo,
        "start": fecha_inicio_dt.isoformat(),
        "durationMinutes": duracion_minutos,
        "scheduledType": "webinar",  # Alternativas según licenciamiento: 'scheduledMeeting'
        "invitees": [{"email": email} for email in panelistas_emails]
    }
    
    try:
        response = requests.post(WEBEX_API_MEETINGS, headers=get_webex_headers(), json=payload, timeout=20)
        if response.status_code in (200, 201):
            return True, response.json()
        else:
            print(f"❌ Error API Webex ({response.status_code}): {response.text}")
            return False, response.json().get("message", "Error desconocido al invocar API de Webex.")
    except Exception as e:
        return False, str(e)

def procesar_creacion_webinar(room_id, raw_text):
    """Parsea el texto recibido y procesa la solicitud de creación de webinar."""
    try:
        partes = [p.strip() for p in raw_text.split(";") if p.strip()]
        titulo = None
        fecha_inicio = None
        duracion = 60
        panelistas = []

        for parte in partes:
            if parte.lower().startswith("título:") or parte.lower().startswith("titulo:"):
                titulo = parte.split(":", 1)[1].strip()
            elif parte.lower().startswith("fecha:"):
                fecha_str = parte.split(":", 1)[1].strip()
                # Parsea formato AAAA-MM-DD HH:MM
                dt_naive = datetime.strptime(fecha_str, "%Y-%m-%d %H:%M")
                fecha_inicio = dt_naive.replace(tzinfo=TZ)
            elif parte.lower().startswith("duración:") or parte.lower().startswith("duracion:"):
                duracion = int(parte.split(":", 1)[1].strip())
            elif parte.lower().startswith("panelistas:"):
                raw_emails = parte.split(":", 1)[1].strip()
                panelistas = [e.strip() for e in raw_emails.split(",") if e.strip()]

        if not titulo or not fecha_inicio:
            send_message(room_id, "⚠️ **Faltan datos requeridos.** Asegúrate de incluir el Título y la Fecha.")
            return

        send_message(room_id, f"⏳ Creando webinar **'{titulo}'** en Webex...")
        
        exito, resultado = crear_webinar_api(titulo, fecha_inicio, duracion, panelistas)
        
        if exito:
            web_link = resultado.get("webLink", "No disponible")
            msg_exito = (
                f"✅ **¡Webinar Creado Exitosamente!**\n\n"
                f"• **Título:** {titulo}\n"
                f"• **Fecha/Hora:** {fecha_inicio.strftime('%Y-%m-%d %H:%M')}\n"
                f"• **Duración:** {duracion} minutos\n"
                f"• **Enlace:** {web_link}"
            )
            send_message(room_id, msg_exito)
        else:
            send_message(room_id, f"❌ **Error al crear el Webinar:** {resultado}")

    except Exception as e:
        print(f"Error procesando parser: {e}")
        send_message(room_id, "❌ Error al procesar la solicitud. Verifica el formato e intenta nuevamente.")

# ==========================================
# WEBHOOK ENDPOINT
# ==========================================
@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json or {}
    if "data" not in data:
        return "ok", 200

    msg_id = data["data"].get("id")
    if not msg_id:
        return "ok", 200

    try:
        # Obtener contenido del mensaje desde Webex
        res = requests.get(f"{WEBEX_API_MESSAGES}/{msg_id}", headers=get_webex_headers(), timeout=15)
        if res.status_code != 200:
            return "ok", 200
            
        msg = res.json()
        raw_text = (msg.get("text") or msg.get("markdown") or "").strip()
        texto_lower = raw_text.lower()
        room = msg.get("roomId")
        sender = msg.get("personEmail", "")

        # Ignorar mensajes de bots o vacíos
        if not sender or not raw_text or sender.endswith("@webex.bot"):
            return "ok", 200

        print(f"📩 [{sender}]: {raw_text}")

        # Routing por Menú u Opciones
        if texto_lower in ["hola", "menu", "menú", "opciones", "start", "ayuda"]:
            mostrar_menu_principal(room)
            
        elif texto_lower == "1":
            mostrar_instrucciones_opcion_1(room)

        elif "crear webinar" in texto_lower:
            threading.Thread(target=procesar_creacion_webinar, args=(room, raw_text)).start()

        elif texto_lower == "2":
            send_message(room, "ℹ️ La **Opción 2** aún no está configurada.")

        elif texto_lower == "3":
            send_message(room, "ℹ️ La **Opción 3** aún no está configurada.")

        else:
            send_message(room, "❓ Opción no reconocida. Escribe **menu** para ver las opciones disponibles.")

    except Exception as e:
        print("❌ Error general en Webhook:", e)

    return "ok", 200

@app.route("/ping", methods=["GET"])
def ping():
    return "Servidor activo", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))
































