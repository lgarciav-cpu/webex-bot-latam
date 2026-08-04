from flask import Flask, request
import os
import requests
import threading
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

# =========================
# CONFIGURACIÓN GENERAL
# =========================
WEBEX_TOKEN = os.environ.get("WEBEX_TOKEN", "").strip()            # Token del BOT (Hablar)
WEBEX_USER_TOKEN = os.environ.get("WEBEX_USER_TOKEN", "").strip()  # Tu Token Personal (Agendar)

WEBEX_API_MESSAGES = "https://webexapis.com/v1/messages"
WEBEX_API_MEETINGS = "https://webexapis.com/v1/meetings"

TZ = ZoneInfo("America/Lima")
app = Flask(__name__)

def get_bot_headers():
    """Headers con el token del BOT para enviar mensajes al chat."""
    if not WEBEX_TOKEN:
        raise RuntimeError("Falta WEBEX_TOKEN (Bot) en variables de entorno.")
    return {
        "Authorization": f"Bearer {WEBEX_TOKEN}",
        "Content-Type": "application/json"
    }

# 👈 AGREGA ESTA LÍNEA PARA EVITAR EL ERROR:
get_webex_headers = get_bot_headers

def get_user_headers():
    """Headers con tu TOKEN PERSONAL para agendar reuniones."""
    token = WEBEX_USER_TOKEN if WEBEX_USER_TOKEN else WEBEX_TOKEN
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

def send_message(room_id: str, text: str):
    """Usa el token del Bot para responder en el chat."""
    data = {"roomId": room_id, "markdown": text}
    try:
        requests.post(WEBEX_API_MESSAGES, headers=get_bot_headers(), json=data, timeout=15)
    except Exception as e:
        print(f"❌ Error al enviar mensaje: {e}")

def crear_webinar_api(titulo, fecha_inicio_dt, duracion_minutos, panelistas_emails):
    """
    Envía un payload estándar a Webex Meetings y captura el detalle real del error.
    """
    # 1. Convertir fecha a UTC y calcular hora de fin
    fecha_utc = fecha_inicio_dt.astimezone(timezone.utc)
    fecha_fin_utc = fecha_utc + timedelta(minutes=duracion_minutos)

    start_iso = fecha_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
    end_iso = fecha_fin_utc.strftime("%Y-%m-%dT%H:%M:%SZ")

    # 2. Payload minimalista y 100% compatible con Webex REST API
    payload = {
        "title": titulo,
        "start": start_iso,
        "end": end_iso
    }

    if panelistas_emails:
        payload["invitees"] = [{"email": email} for email in panelistas_emails if email]

    try:
        response = requests.post(WEBEX_API_MEETINGS, headers=get_user_headers(), json=payload, timeout=20)
        
        if response.status_code in (200, 201):
            return True, response.json()
        else:
            err_json = response.json()
            print(f"❌ Respuesta de Webex ({response.status_code}): {err_json}")
            
            # Extraer el detalle real dentro del arreglo 'errors'
            mensaje_error = err_json.get("message", f"Error HTTP {response.status_code}")
            if "errors" in err_json and isinstance(err_json["errors"], list) and len(err_json["errors"]) > 0:
                mensaje_error = err_json["errors"][0].get("description", mensaje_error)
                
            return False, mensaje_error

    except Exception as e:
        return False, str(e)

def procesar_creacion_webinar(room_id, raw_text):
    try:
        partes = [p.strip() for p in raw_text.split(";") if p.strip()]
        titulo = None
        fecha_inicio = None
        duracion = 60
        panelistas = []

        for parte in partes:
            p_lower = parte.lower()
            if p_lower.startswith("título:") or p_lower.startswith("titulo:"):
                titulo = parte.split(":", 1)[1].strip()
            elif p_lower.startswith("fecha:"):
                fecha_str = parte.split(":", 1)[1].strip()
                dt_naive = datetime.strptime(fecha_str, "%Y-%m-%d %H:%M")
                fecha_inicio = dt_naive.replace(tzinfo=TZ)
            elif p_lower.startswith("duración:") or p_lower.startswith("duracion:"):
                duracion = int(parte.split(":", 1)[1].strip())
            elif p_lower.startswith("panelistas:"):
                raw_emails = parte.split(":", 1)[1].strip()
                panelistas = [e.strip() for e in raw_emails.split(",") if e.strip()]

        if not titulo or not fecha_inicio:
            send_message(room_id, "⚠️ **Faltan datos requeridos.** Incluye Título y Fecha.")
            return

        send_message(room_id, f"⏳ Creando sesión **'{titulo}'**...")
        exito, resultado = crear_webinar_api(titulo, fecha_inicio, duracion, panelistas)

        if exito:
            web_link = resultado.get("webLink", "No disponible")
            msg_exito = (
                f"✅ **¡Sesión Creada Exitosamente!**\n\n"
                f"📌 **Título:** {titulo}\n"
                f"📅 **Fecha/Hora:** {fecha_inicio.strftime('%Y-%m-%d %H:%M')}\n"
                f"⏱️ **Duración:** {duracion} min\n"
                f"🔗 **Enlace:** {web_link}"
            )
            send_message(room_id, msg_exito)
        else:
            send_message(room_id, f"❌ **Error al crear:** {resultado}")

    except Exception as e:
        print(f"Error procesando parser: {e}")
        send_message(room_id, "❌ Error al procesar el texto. Verifica el formato.")

# ==========================================
# FUNCIONES DE MENÚ E INSTRUCCIONES
# ==========================================
import re
from datetime import datetime, timedelta, timezone
from flask import Flask, request, jsonify

# ==========================================
# FUNCIONES AUXILIARES DE MENÚ
# ==========================================
def mostrar_menu_principal(room_id):
    menu_texto = (
        "🤖 **¡Hola! Bienvenido al Bot de Gestión de Webex**\n\n"
        "Selecciona una de las siguientes opciones:\n\n"
        "1️⃣ **Crear Webinar / Sesión Webex**\n"
        "2️⃣ **Próximamente: Consultas**\n"
        "3️⃣ **Próximamente: Reportes**\n\n"
        "💡 *Escribe **1** o el comando para crear la sesión.*"
    )
    send_message(room_id, menu_texto)


def mostrar_instrucciones_opcion_1(room_id):
    instrucciones = (
        "📌 **OPCIÓN 1: Crear Webinar / Sesión**\n\n"
        "Envía la información con este formato exacto:\n\n"
        "```text\n"
        "crear webinar;\n"
        "Título: Lanzamiento de Producto 2026;\n"
        "Fecha: 2026-08-15 15:00;\n"
        "Duración: 60;\n"
        "Panelistas: juan@empresa.com, maria@empresa.com\n"
        "```"
    )
    send_message(room_id, instrucciones)


def extraer_datos_webinar(texto_mensaje: str):
    """Parsea el texto del usuario extraendo titulo, fecha, duracion y panelistas."""
    lineas = re.split(r'[\n;]+', texto_mensaje)
    
    datos = {}
    for linea in lineas:
        if ":" in linea:
            clave, valor = linea.split(":", 1)
            clave_clean = clave.strip().lower().replace("í", "i").replace("ó", "o")
            datos[clave_clean] = valor.strip()

    titulo = datos.get("titulo")
    fecha_raw = datos.get("fecha")
    duracion_raw = datos.get("duracion", "30")
    panelistas_raw = datos.get("panelistas", "")

    if not titulo:
        return False, "❌ Falta el campo **Título:** en el mensaje."
    if not fecha_raw:
        return False, "❌ Falta el campo **Fecha:** en el mensaje."

    try:
        fecha_dt = datetime.strptime(fecha_raw, "%Y-%m-%d %H:%M")
        fecha_dt = fecha_dt.replace(tzinfo=TZ)
    except ValueError:
        return False, f"❌ Formato de fecha inválido (`{fecha_raw}`). Usa: `YYYY-MM-DD HH:MM` (Ejemplo: `2026-08-04 15:00`)."

    try:
        duracion = int(re.sub(r'\D', '', duracion_raw))
    except ValueError:
        duracion = 30

    panelistas = [e.strip() for e in re.split(r'[,;\s]+', panelistas_raw) if "@" in e]

    return True, {
        "titulo": titulo,
        "fecha_dt": fecha_dt,
        "duracion": duracion,
        "panelistas": panelistas
    }

# ==========================================
# ENDPOINT DE WEBHOOK
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
        # 1. Obtener el contenido del mensaje desde Webex
        res = requests.get(f"{WEBEX_API_MESSAGES}/{msg_id}", headers=get_bot_headers(), timeout=15)
        if res.status_code != 200:
            return "ok", 200
            
        msg = res.json()
        raw_text = (msg.get("text") or msg.get("markdown") or "").strip()
        texto_lower = raw_text.lower()
        room = msg.get("roomId")
        sender = msg.get("personEmail", "")

        # 2. Ignorar mensajes propios o de otros bots
        if not sender or not raw_text or sender.endswith("@webex.bot"):
            return "ok", 200

        print(f"📩 [{sender}]: {raw_text}")

        # 3. Evaluar el mensaje del usuario
        if texto_lower in ["hola", "menu", "menú", "opciones", "start", "ayuda"]:
            mostrar_menu_principal(room)
            
        elif texto_lower == "1":
            mostrar_instrucciones_opcion_1(room)
            
        elif "crear webinar" in texto_lower:
            exito, resultado = extraer_datos_webinar(raw_text)
            
            if not exito:
                send_message(room, resultado)
            else:
                datos = resultado
                send_message(room, f"⌛ Creando sesión **'{datos['titulo']}'**...")
                
                creado, res_api = crear_webinar_api(
                    titulo=datos["titulo"],
                    fecha_inicio_dt=datos["fecha_dt"],
                    duracion_minutos=datos["duracion"],
                    panelistas_emails=datos["panelistas"]
                )
                
                if creado:
                    web_link = res_api.get("webLink", "No disponible")
                    send_message(room, f"✅ **¡Sesión creada con éxito!**\n\n📌 **Título:** {datos['titulo']}\n🔗 **Enlace:** {web_link}")
                else:
                    send_message(room, f"❌ **Error al crear:** {res_api}")
                    
        else:
            send_message(room, "❓ Comando no reconocido. Escribe **menu** para ver las opciones.")

    except Exception as e:
        print("❌ Error general en Webhook:", e)

    return "ok", 200

@app.route("/ping", methods=["GET", "HEAD"])
def ping():
    return "Servidor activo", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))
































