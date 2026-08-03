from flask import Flask, request
import os
import requests
import random
import openpyxl
from io import BytesIO
from datetime import datetime, timedelta, time as dtime
import threading
import time
from zoneinfo import ZoneInfo  # Python 3.9+
import glob

# --- NUEVAS LIBRERÍAS DE IA ---
import google.generativeai as genai
import PyPDF2
# =========================
# CONFIG
# =========================
WEBEX_TOKEN = os.environ.get("WEBEX_TOKEN", "").strip()
WEBEX_API = "https://webexapis.com/v1/messages"
WEBEX_WEBINAR_API = "https://webexapis.com/v1/webinars"

# Export XLSX público (o que permita descarga)
EXCEL_URL = "https://docs.google.com/spreadsheets/d/1sWFXSOY0jZ8PaSh2Lg1lnmCBGN96fLkC/export?format=xlsx"

TZ = ZoneInfo("America/Lima")

app = Flask(__name__)

GIFS_HOLA = [
    "https://media.giphy.com/media/v1.Y2lkPTc5MGI3NjExeWwzZmpsNG81ODc2YnFmM2x6ZXpmMGFsdzJ2eWo1c2owcTM5NWhrNiZlcD12MV9naWZzX3NlYXJjaCZjdD1n/ASd0Ukj0y3qMM/giphy.gif",
    "https://media.giphy.com/media/v1.Y2lkPTc5MGI3NjExNnM3Nzc4MndqZ2xvZGg0bm03MGY5bjdlOG43ejhxeTZibWozam5vYyZlcD12MV9naWZzX3NlYXJjaCZjdD1n/3o7budMRwZvNGJ3pyE/giphy.gif",
    "https://media.giphy.com/media/v1.Y2lkPTc5MGI3NjExNnM3Nzc4MndqZ2xvZGg0bm03MGY5bjdlOG43ejhxeTZibWozam5vYyZlcD12MV9naWZzX3NlYXJjaCZjdD1n/xT9IgG50Fb7Mi0prBC/giphy.gif",
    "https://media.giphy.com/media/v1.Y2lkPTc5MGI3NjExNnM3Nzc4MndqZ2xvZGg0bm03MGY5bjdlOG43ejhxeTZibWozam5vYyZlcD12MV9naWZzX3NlYXJjaCZjdD1n/vFKqnCdLPNOKc/giphy.gif",
    "https://media.giphy.com/media/v1.Y2lkPTc5MGI3NjExeWwzZmpsNG81ODc2YnFmM2x6ZXpmMGFsdzJ2eWo1c2owcTM5NWhrNiZlcD12MV9naWZzX3NlYXJjaCZjdD1n/Xev2JdopBxGj1LuGvt/giphy.gif",
]

# Para evitar enviar el mismo anuncio 60 veces
SENT_CACHE = set()  # guarda tuplas (date, hh:mm, room, msg)

# ==========================================
# CONFIGURACIÓN DE IA: LLAMA 3 (Vía Groq)
# ==========================================
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "").strip()

def consultar_ia_con_llama(pregunta):
    if not GROQ_API_KEY:
        return "IA desconectada. Falta la llave de Groq."

    print("🧠 IA: Leyendo PDFs...")
    memoria_texto = ""
    archivos_pdf = glob.glob("pdfs/*.pdf")
    for ruta in archivos_pdf:
        try:
            with open(ruta, "rb") as f:
                lector = PyPDF2.PdfReader(f)
                for pagina in lector.pages:
                    texto = pagina.extract_text()
                    if texto:
                        memoria_texto += texto + "\n"
        except Exception:
            pass

    url = "https://api.groq.com/openai/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {GROQ_API_KEY}",
        "Content-Type": "application/json"
    }
    
    payload = {
        "model": "llama-3.3-70b-versatile",
        "messages": [
            {
                "role": "system", 
                "content": f"Eres un asistente corporativo de la empresa. Usa SOLO esta información para responder, si no lo sabes, di que no está en el manual:\n\n{memoria_texto}"
            },
            {
                "role": "user", 
                "content": pregunta
            }
        ],
        "temperature": 0.3
    }

    try:
        print("🧠 IA: Consultando a Llama 3...")
        response = requests.post(url, headers=headers, json=payload, timeout=20)
        response.raise_for_status()
        data = response.json()
        print("🧠 IA: ¡Respuesta lista!")
        return data["choices"][0]["message"]["content"]
    except Exception as e:
        print(f"❌ Error de IA: {str(e)}")
        return "Lo siento, tuve un problema procesando el texto. Por favor, intenta de nuevo."

# ==========================================
# FUNCIONES DE WEBEX Y EXCEL (Tus originales)
# ==========================================
def _headers():
    if not WEBEX_TOKEN:
        raise RuntimeError("Falta WEBEX_TOKEN en variables de entorno.")
    return {
        "Authorization": f"Bearer {WEBEX_TOKEN}",
        "Content-Type": "application/json"
    }

def send_message(room_id: str, text: str):
    data = {"roomId": room_id, "text": text}
    requests.post(WEBEX_API, headers=_headers(), json=data, timeout=15)

def send_gif(room_id: str, gif_url: str):
    data = {"roomId": room_id, "files": [gif_url]}
    requests.post(WEBEX_API, headers=_headers(), json=data, timeout=15)

def leer_excel():
    try:
        r = requests.get(EXCEL_URL, timeout=30)
        r.raise_for_status()
        wb = openpyxl.load_workbook(BytesIO(r.content), data_only=True)
        ws = wb.active

        datos = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            fecha, hora, roomid, mensaje = row
            if not (fecha and hora and roomid and mensaje):
                continue
            datos.append({"Fecha": fecha, "Hora": hora, "RoomID": str(roomid).strip(), "Mensaje": str(mensaje).strip()})
        return datos
    except Exception as e:
        print("Error leyendo Excel:", e)
        return []

def normalizar_datetime(fecha_excel, hora_excel):
    try:
        if isinstance(fecha_excel, datetime):
            f = fecha_excel.date()
        else:
            f = fecha_excel

        if isinstance(hora_excel, datetime):
            t = hora_excel.time()
        elif isinstance(hora_excel, dtime):
            t = hora_excel
        elif isinstance(hora_excel, (int, float)):
            seconds = int(round(float(hora_excel) * 24 * 3600))
            hh = (seconds // 3600) % 24
            mm = (seconds % 3600) // 60
            ss = seconds % 60
            t = dtime(hh, mm, ss)
        else:
            s = str(hora_excel).lower().replace("a. m.", "am").replace("p. m.", "pm").strip()
            from dateutil import parser
            dt = parser.parse(s)
            t = dt.time()

        dt_local = datetime.combine(f, t).replace(tzinfo=TZ)
        return dt_local
    except Exception as e:
        print("Error normalizando datetime:", e, "fecha:", fecha_excel, "hora:", hora_excel)
        return None

def scheduler():
    while True:
        try:
            ahora = datetime.now(TZ)
            filas = leer_excel()

            for row in filas:
                dt_prog = normalizar_datetime(row["Fecha"], row["Hora"])
                if not dt_prog:
                    continue

                diff = (ahora - dt_prog).total_seconds()
                if 0 <= diff <= 600:
                    key = (dt_prog.date().isoformat(), dt_prog.strftime("%H:%M"), row["RoomID"], row["Mensaje"])
                    if key in SENT_CACHE:
                        continue

                    SENT_CACHE.add(key)
                    gif = random.choice(GIFS_HOLA)
                    print(f"✔ Enviando programado a {row['RoomID']} @ {dt_prog}: {row['Mensaje']}")
                    send_gif(row["RoomID"], gif)
                    send_message(row["RoomID"], row["Mensaje"])

            if len(SENT_CACHE) > 2000:
                SENT_CACHE.clear()

        except Exception as e:
            print("Error en scheduler loop:", e)

        time.sleep(30)

# ==========================================
# FUNCIONES PARA WEBEX WEBINARS
# ==========================================
def crear_webinar(titulo, fecha_inicio, duracion_minutos, panelistas_emails):
    headers = _headers()
    payload = {
        "title": titulo,
        "start": fecha_inicio.isoformat(),
        "durationMinutes": duracion_minutos,
        "panelists": [{"email": email} for email in panelistas_emails]
    }
    response = requests.post(WEBEX_WEBINAR_API, headers=headers, json=payload)
    if response.status_code in (200, 201):
        return response.json()
    else:
        print(f"Error creando webinar: {response.status_code} {response.text}")
        return None

def agregar_panelista_webinar(webinar_id, email_panelista):
    headers = _headers()
    url = f"{WEBEX_WEBINAR_API}/{webinar_id}/panelists"
    payload = {"email": email_panelista}
    response = requests.post(url, headers=headers, json=payload)
    if response.status_code in (200, 201):
        return response.json()
    else:
        print(f"Error agregando panelista: {response.status_code} {response.text}")
        return None

# ==========================================
# WEBHOOK DE WEBEX
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
        msg = requests.get(f"https://webexapis.com/v1/messages/{msg_id}", headers=_headers(), timeout=15).json()
        raw_text = (msg.get("text") or msg.get("markdown") or "").strip()
        texto = raw_text.lower()
        room = msg.get("roomId")
        sender = msg.get("personEmail", "")

        if not sender or not raw_text or sender.endswith("@webex.bot"):
            return "ok", 200

        print(f"📩 Mensaje de {sender}: '{raw_text}'")

        # Comandos para webinars
        if "crear webinar" in texto:
            # Ejemplo: extraer parámetros básicos del texto (debes adaptar el parseo real)
            # Aquí se asume que el usuario envía algo como:
            # "crear webinar Título: Mi Webinar; Fecha: 2026-08-10 15:00; Duración: 60; Panelistas: user1@example.com,user2@example.com"
            try:
                # Parseo simple (puedes mejorar con regex o NLP)
                partes = raw_text.split(";")
                titulo = None
                fecha_inicio = None
                duracion = 60
                panelistas = []

                for parte in partes:
                    if "título:" in parte.lower():
                        titulo = parte.split(":",1)[1].strip()
                    elif "fecha:" in parte.lower():
                        fecha_str = parte.split(":",1)[1].strip()
                        fecha_inicio = datetime.fromisoformat(fecha_str).replace(tzinfo=TZ)
                    elif "duración:" in parte.lower():
                        duracion = int(parte.split(":",1)[1].strip())
                    elif "panelistas:" in parte.lower():
                        panelistas = [email.strip() for email in parte.split(":",1)[1].split(",") if email.strip()]

                if not titulo or not fecha_inicio:
                    send_message(room, "Por favor, proporciona al menos el título y la fecha en formato ISO para crear el webinar.")
                    return "ok", 200

                resultado = crear_webinar(titulo, fecha_inicio, duracion, panelistas)
                if resultado:
                    send_message(room, f"Webinar '{titulo}' creado con éxito para {fecha_inicio.isoformat()}.")
                else:
                    send_message(room, "Error al crear el webinar. Por favor, intenta de nuevo.")
            except Exception as e:
                print(f"Error procesando comando crear webinar: {e}")
                send_message(room, "No pude procesar el comando para crear webinar. Asegúrate de usar el formato correcto.")
        
        elif "agregar panelista" in texto:
            # Ejemplo: "agregar panelista webinar_id email@example.com"
            try:
                tokens = raw_text.split()
                if len(tokens) >= 4:
                    webinar_id = tokens[2]
                    email_panelista = tokens[3]
                    resultado = agregar_panelista_webinar(webinar_id, email_panelista)
                    if resultado:
                        send_message(room, f"Panelista {email_panelista} agregado al webinar {webinar_id}.")
                    else:
                        send_message(room, "Error al agregar panelista. Verifica el ID del webinar y el email.")
                else:
                    send_message(room, "Formato incorrecto. Usa: agregar panelista <webinar_id> <email>")
            except Exception as e:
                print(f"Error procesando comando agregar panelista: {e}")
                send_message(room, "No pude procesar el comando para agregar panelista.")
        
        elif any(w in texto for w in ["hola", "hello", "hi", "buenas"]):
            gif = random.choice(GIFS_HOLA)
            send_gif(room, gif)
            send_message(room, "👋 ¡Hola! ¿Qué tal? Pregúntame lo que necesites.")
        elif any(w in texto for w in ["ayuda", "help"]):
            send_message(room, "Puedo enviar mensajes desde tu Excel, crear webinars en Webex y responder preguntas sobre nuestros PDFs.")
        else:
            send_message(room, "Buscando en mis archivos con Llama 3... 🦙🧠")
            def pensar_y_responder():
                respuesta_ia = consultar_ia_con_llama(raw_text)
                send_message(room, respuesta_ia)
            threading.Thread(target=pensar_y_responder).start()

    except Exception as e:
        print("❌ Error webhook:", e)

    return "ok", 200

# ==========================================
# INICIO DE SERVIDOR Y TAREAS
# ==========================================
_scheduler_started = False

def start_scheduler_once():
    global _scheduler_started
    if _scheduler_started:
        return
    _scheduler_started = True
    threading.Thread(target=scheduler, daemon=True).start()
    print("✅ Scheduler thread started")

@app.route("/ping", methods=["GET"])
def ping():
    return "ok", 200

if __name__ == "__main__":
    start_scheduler_once()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))
































