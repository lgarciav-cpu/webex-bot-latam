from flask import Flask, request
import requests
import random
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
import threading
import time
from dateutil import parser

WEBEX_BOT_TOKEN = "ZmY2MGJlYWYtMzgxYy00ZDljLWIyMmYtMTZkYjRlMTc2N2EzNTYxYjk5ZjgtN2Iw_PF84_1eb65fdf-9643-417f-9974-ad72cae0e10f"  # <-- pon tu token aquí

app = Flask(__name__)
WEBEX_API = "https://webexapis.com/v1/messages"

GIFS_HOLA = [
    "https://media.giphy.com/media/v1.Y2lkPTc5MGI3NjExeWwzZmpsNG81ODc2YnFmM2x6ZXpmMGFsdzJ2eWo1c2owcTM5NWhrNiZlcD12MV9naWZzX3NlYXJjaCZjdD1n/ASd0Ukj0y3qMM/giphy.gif",
    "https://media.giphy.com/media/v1.Y2lkPTc5MGI3NjExeWwzZmpsNG81ODc2YnFmM2x6ZXpmMGFsdzJ2eWo1c2owcTM5NWhrNiZlcD12MV9naWZzX3NlYXJjaCZjdD1n/Xev2JdopBxGj1LuGvt/giphy.gif",
    "https://media.giphy.com/media/v1.Y2lkPTc5MGI3NjExeWwzZmpsNG81ODc2YnFmM2x6ZXpmMGFsdzJ2eWo1c2owcTM5NWhrNiZlcD12MV9naWZzX3NlYXJjaCZjdD1n/a1QLZUUtCcgyA/giphy.gif",
    "https://media.giphy.com/media/v1.Y2lkPTc5MGI3NjExeWwzZmpsNG81ODc2YnFmM2x6ZXpmMGFsdzJ2eWo1c2owcTM5NWhrNiZlcD12MV9naWZzX3NlYXJjaCZjdD1n/noyBeNjH4nbtXV5ZLA/giphy.gif",
    "https://media.giphy.com/media/v1.Y2lkPTc5MGI3NjExeWwzZmpsNG81ODc2YnFmM2x6ZXpmMGFsdzJ2eWo1c2owcTM5NWhrNiZlcD12MV9naWZzX3NlYXJjaCZjdD1n/yrhhmre5fN2PtRujfo/giphy.gif",
    "https://media.giphy.com/media/v1.Y2lkPWVjZjA1ZTQ3a3lsMzAxeXpvYmg5eXhwcm9kMjdsdTNsbGpvYnN0enE0MmJjYmF4YiZlcD12MV9naWZzX3NlYXJjaCZjdD1n/3pZipqyo1sqHDfJGtz/giphy.gif",
    "https://media.giphy.com/media/v1.Y2lkPWVjZjA1ZTQ3a3lsMzAxeXpvYmg5eXhwcm9kMjdsdTNsbGpvYnN0enE0MmJjYmF4YiZlcD12MV9naWZzX3NlYXJjaCZjdD1n/FAFo1M7EC4gRZ4HETH/giphy.gif",
    "https://media.giphy.com/media/v1.Y2lkPWVjZjA1ZTQ3a3lsMzAxeXpvYmg5eXhwcm9kMjdsdTNsbGpvYnN0enE0MmJjYmF4YiZlcD12MV9naWZzX3NlYXJjaCZjdD1n/MvdaYPuKPMNZRJCl8Z/giphy.gif",
    "https://media.giphy.com/media/v1.Y2lkPWVjZjA1ZTQ3bmd0dTE3bWFpNXE3aXBzMmpyaGVoeXMxem9uM25vdWJwcW9mZnM4NyZlcD12MV9naWZzX3NlYXJjaCZjdD1n/rwsHPJ5zOkD0mVwAlC/giphy.gif",
    "https://media.giphy.com/media/v1.Y2lkPWVjZjA1ZTQ3bmd0dTE3bWFpNXE3aXBzMmpyaGVoeXMxem9uM25vdWJwcW9mZnM4NyZlcD12MV9naWZzX3NlYXJjaCZjdD1n/11Wf3llSqbkgko/giphy.gif",
    "https://media.giphy.com/media/v1.Y2lkPWVjZjA1ZTQ3YnRiamVrbjhubmhzems3NTdyd2o0bm9ybGhncWdmaDhzNWl5cGE2eSZlcD12MV9naWZzX3NlYXJjaCZjdD1n/xTk9ZY0C9ZWM2NgmCA/giphy.gif",
    "https://media.giphy.com/media/v1.Y2lkPWVjZjA1ZTQ3YnRiamVrbjhubmhzems3NTdyd2o0bm9ybGhncWdmaDhzNWl5cGE2eSZlcD12MV9naWZzX3NlYXJjaCZjdD1n/AFdcYElkoNAUE/giphy.gif",
    "https://media.giphy.com/media/v1.Y2lkPWVjZjA1ZTQ3YXdzeWJiMG4wOHZ0d3NuczdpOXk0ODM5M2lrcHp3eDQxOXo2N3NpYSZlcD12MV9naWZzX3NlYXJjaCZjdD1n/MEFVcuRIoVETUMYZEe/giphy.gif"
]

EXCEL_URL = "https://docs.google.com/spreadsheets/d/1sWFXSOY0jZ8PaSh2Lg1lnmCBGN96fLkC/export?format=xlsx"

def leer_excel():
    try:
        df = pd.read_excel(EXCEL_URL)
        print("Google Sheet leído correctamente.")
        return df
    except Exception as e:
        print("Error leyendo Google Sheet:", e)
        return pd.DataFrame()

def parse_hora_excel(valor):
    try:
        # reemplazar formatos latinos
        txt = str(valor).replace("a. m.", "AM").replace("p. m.", "PM").strip()
        dt = parser.parse(txt)
        return dt
    except:
        return None

def revisar_y_enviar_mensajes():
    try:
        df = leer_excel()

        if df.empty:
            print("Excel vacío o no accesible.")
            return

        ahora = datetime.now()
        fecha_hoy = ahora.date()

        print(f"Fecha hoy: {fecha_hoy}")
        print(f"Ahora: {ahora}")

        # Normalizar columnas
        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.date
        df["Hora_datetime"] = df["Hora"].apply(parse_hora_excel)

        filas = df[df["Fecha"] == fecha_hoy]

        print(filas[["Fecha","Hora","Hora_datetime","RoomID"]])

        for _, row in filas.iterrows():
            hora_dt = row["Hora_datetime"]

            if hora_dt is None:
                continue

            if abs((hora_dt - ahora).total_seconds()) <= 60:
                room_id = str(row["RoomID"]).strip()
                mensaje = str(row["Mensaje"]).strip()

                print(f"Enviando mensaje programado a {room_id}: {mensaje}")

                gif = buscar_gif_hola()
                send_gif(room_id, gif)
                send_message(room_id, mensaje)

    except Exception as e:
        print("Error en revisar_y_enviar_mensajes:", e)

def scheduler():
    while True:
        revisar_y_enviar_mensajes()
        time.sleep(60)  # revisa cada 1 minuto

def buscar_gif_hola():
    return random.choice(GIFS_HOLA)

def send_message(room_id, text):
    headers = {
        "Authorization": f"Bearer {WEBEX_BOT_TOKEN}",
        "Content-Type": "application/json"
    }
    data = {"roomId": room_id, "text": text}
    requests.post(WEBEX_API, headers=headers, json=data)

def send_gif(room_id, gif_url):
    headers = {
        "Authorization": f"Bearer {WEBEX_BOT_TOKEN}",
        "Content-Type": "application/json"
    }
    data = {
        "roomId": room_id,
        "files": [gif_url]
    }
    requests.post(WEBEX_API, headers=headers, json=data)

@app.route("/", methods=["POST"])
def webhook():
    data = request.json

    # Validación ultra segura
    if not data or "data" not in data:
        return "ok", 200

    message_id = data["data"].get("id")
    if not message_id:
        return "ok", 200

    # obtener detalles del mensaje
    headers = {"Authorization": f"Bearer {WEBEX_BOT_TOKEN}"}
    msg_data = requests.get(f"https://webexapis.com/v1/messages/{message_id}", headers=headers).json()

    text = msg_data.get("text")
    room_id = msg_data.get("roomId")
    sender = msg_data.get("personEmail")

    # Si no hay room o sender, no hacemos nada
    if not room_id or not sender:
        return "ok", 200

    # evitar que el bot se responda a sí mismo
    if sender.endswith("@webex.bot"):
        return "ok", 200

    # si no hay texto (GIF, sticker, archivo), responder algo opcional
    if not isinstance(text, str):
        send_message(room_id, "📎 Recibí tu mensaje, pero no tenía texto.")
        return "ok", 200
    
    text = text.lower()

    # respuestas básicas
    if "hola" in text:
        gif = buscar_gif_hola()
        print("GIF seleccionado:", gif)   # <--- DEBUG
        if gif:
            send_gif(room_id, gif)
        send_message(room_id, "👋 ¡Hola! ¿Qué tal?")
    elif "ayuda" in text:
        send_message(room_id, "Puedo saludar y pronto aprenderé recordatorios 🔔")
    else:
        send_message(room_id, "No entendí 😅. Escribe **hola** o **ayuda**.")

    return "ok", 200

if __name__ == "__main__":
    # Hilo paralelo que revisa el Excel cada minuto
    threading.Thread(target=scheduler, daemon=True).start()

    # Webhook activo
    app.run(port=5000)