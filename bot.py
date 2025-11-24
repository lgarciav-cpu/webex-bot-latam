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
    headers = {"Authorization": f"Bearer {WEBEX_TOKEN}"}
    data = {"roomId": room_id, "text": text}
    requests.post(WEBEX_API, headers=headers, json=data)

def send_gif(room_id):
    gif = random.choice(GIFS_HOLA)
    headers = {"Authorization": f"Bearer {WEBEX_TOKEN}"}
    data = {"roomId": room_id, "files": [gif]}
    requests.post(WEBEX_API, headers=headers, json=data)

def leer_excel():
    try:
        df = pd.read_excel(EXCEL_URL)
        return df
    except Exception as e:
        print("Error leyendo Excel:", e)
        return None

def parse_hora(valor):
    try:
        txt = str(valor).replace("a. m.", "AM").replace("p. m.", "PM")
        return parser.parse(txt)
    except:
        return None

def ejecutar_scheduler():
    while True:
        df = leer_excel()
        if df is None:
            time.sleep(60)
            continue
        
        ahora = datetime.now()
        fecha_hoy = ahora.date()

        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.date
        df["Hora_dt"] = df["Hora"].apply(parse_hora)

        hoy = df[df["Fecha"] == fecha_hoy]

        for _, row in hoy.iterrows():
            hora_msg = row["Hora_dt"]
            if hora_msg is None:
                continue

            diferencia = abs((hora_msg - ahora).total_seconds())
            if diferencia <= 60:  # exacto a 1 minuto
                room = str(row["RoomID"])
                mensaje = str(row["Mensaje"])

                print(f"✔ Enviando mensaje programado a {room}: {mensaje}")
                send_gif(room)
                send_message(room, mensaje)

        time.sleep(60)

if __name__ == "__main__":
    ejecutar_scheduler()
