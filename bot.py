from flask import Flask, request
import requests
import random
import openpyxl
from io import BytesIO
from datetime import datetime
import threading
import time
from dateutil import parser

WEBEX_TOKEN = "ZmY2MGJlYWYtMzgxYy00ZDljLWIyMmYtMTZkYjRlMTc2N2EzNTYxYjk5ZjgtN2Iw_PF84_1eb65fdf-9643-417f-9974-ad72cae0e10f"  # <-- pon tu token aquí
WEBEX_API = "https://webexapis.com/v1/messages"
EXCEL_URL = "https://docs.google.com/spreadsheets/d/1sWFXSOY0jZ8PaSh2Lg1lnmCBGN96fLkC/export?format=xlsx"

app = Flask(__name__)

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

def send_message(room_id, text):
    headers = {
        "Authorization": f"Bearer {WEBEX_TOKEN}",
        "Content-Type": "application/json"
    }
    data = {"roomId": room_id, "text": text}
    requests.post(WEBEX_API, headers=headers, json=data)


def send_gif(room_id, gif_url):
    headers = {
        "Authorization": f"Bearer {WEBEX_TOKEN}",
        "Content-Type": "application/json"
    }
    data = {"roomId": room_id, "files": [gif_url]}
    requests.post(WEBEX_API, headers=headers, json=data)

def leer_excel():
    try:
        contenido = requests.get(EXCEL_URL).content
        wb = openpyxl.load_workbook(BytesIO(contenido), data_only=True)
        ws = wb.active

        datos = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            fecha, hora, roomid, mensaje = row
            datos.append({
                "Fecha": fecha,
                "Hora": hora,
                "RoomID": roomid,
                "Mensaje": mensaje
            })
        return datos

    except Exception as e:
        print("Error leyendo Excel:", e)
        return []

def parsear_hora(h):
    try:
        txt = str(h).replace("a. m.", "AM").replace("p. m.", "PM")
        hora = parser.parse(txt)
        print("Hora leída del Excel:", hora)
        return hora
    except:
        print("Error parseando hora:", h)
        return None

def revisar_y_enviar_mensajes():
    filas = leer_excel()
    if not filas:
        return

    # Hora actual del servidor (UTC) convertida a Perú (UTC-5)
    ahora = datetime.utcnow() - timedelta(hours=5)
    fecha_hoy = ahora.date()

    print("Hora local (ajustada):", ahora)

    for row in filas:
        if not row["Fecha"]:
            continue
        
        fecha = row["Fecha"]
        hora = parsear_hora(row["Hora"])   # Esto detecta 10:36 a.m., etc.

        if hora:
            # Convertir hora del Excel también a UTC-5
            hora = hora.replace(year=ahora.year, month=ahora.month, day=ahora.day)

        room = row["RoomID"]
        mensaje = row["Mensaje"]

        # Coincidencia exacta dentro de 60 segundos
        if fecha == fecha_hoy and hora:
            diferencia = abs((hora - ahora).total_seconds())
            if diferencia <= 60:
                print(f"✔ Enviando mensaje programado a {room}: {mensaje}")
                send_gif(room)
                send_message(room, mensaje)

def scheduler():
    while True:
        filas = leer_excel()
        hoy = datetime.now().date()
        ahora = datetime.now()

        for row in filas:
            if not row["Fecha"]:
                continue

            fecha = row["Fecha"]
            hora = parsear_hora(row["Hora"])
            room = row["RoomID"]
            mensaje = row["Mensaje"]

            if fecha == hoy and hora:
                diferencia = abs((hora - ahora).total_seconds())
                if diferencia <= 60:
                    print(f" Enviando mensaje programado a {room}")
                    gif = random.choice(GIFS_HOLA)
                    send_gif(room, gif)
                    send_message(room, mensaje)

        time.sleep(60)

@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json

    if "data" not in data:
        return "ok", 200

    msg_id = data["data"]["id"]
    headers = {"Authorization": f"Bearer {WEBEX_TOKEN}"}

    msg = requests.get(f"https://webexapis.com/v1/messages/{msg_id}", headers=headers).json()

    texto = msg.get("text", "").lower()
    room = msg.get("roomId")
    sender = msg.get("personEmail")

    if sender.endswith("@webex.bot"):
        return "ok", 200

    if "hola" in texto:
        gif = random.choice(GIFS_HOLA)
        send_gif(room, gif)
        send_message(room, "👋 ¡Hola! ¿Qué tal?")
    elif "ayuda" in texto:
        send_message(room, "Puedo saludar y también enviar recordatorios desde Excel.")
    else:
        send_message(room, "No entendí 😅. Escribe *hola* o *ayuda*.")

    return "ok", 200

if __name__ == "__main__":
    threading.Thread(target=scheduler, daemon=True).start()
    app.run(port=5000)





