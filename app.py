"""
Slack Punti Bot
───────────────
Bot Slack per consultare punti utenti caricati via Excel.

Funzionamento:
  - Admin trascina un file .xlsx nel canale admin → bot aggiorna i dati
  - Utenti usano /punti per vedere il proprio saldo
  - /classifica mostra la Top 10
  - /stato mostra lo stato del bot

Formato Excel atteso (prima riga = header):
  | email | nome | punti |
"""

import os
import io
import pickle
import logging
import requests
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

# ═══════════════════════════════════════════════════════════════
# ⚙️  CONFIGURAZIONE E LOGGING
# ═══════════════════════════════════════════════════════════════

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('bot.log')
    ]
)
logger = logging.getLogger(__name__)

# Verifica variabili d'ambiente obbligatorie
REQUIRED_ENV = ["SLACK_BOT_TOKEN", "SLACK_APP_TOKEN", "ADMIN_CHANNEL_ID"]
for var in REQUIRED_ENV:
    if not os.environ.get(var):
        raise RuntimeError(f"❌ Variabile d'ambiente mancante: {var}")

ADMIN_CHANNEL = os.environ["ADMIN_CHANNEL_ID"]
BACKUP_FILE = "ultimo_dati.pkl"

app = App(token=os.environ["SLACK_BOT_TOKEN"])

# ═══════════════════════════════════════════════════════════════
# 💾  STATO IN MEMORIA + BACKUP SU DISCO
# ═══════════════════════════════════════════════════════════════

# {email: {"nome": str, "punti": int}}
dati_punti = {}
ultimo_aggiornamento = None
AVVIO = datetime.now()


def salva_backup():
    """Salva i dati su disco per sopravvivere ai restart"""
    try:
        with open(BACKUP_FILE, "wb") as f:
            pickle.dump({
                "dati": dati_punti,
                "aggiornato": ultimo_aggiornamento
            }, f)
        logger.info(f"💾 Backup salvato ({len(dati_punti)} utenti)")
    except Exception as e:
        logger.exception(f"❌ Errore salvataggio backup: {e}")


def carica_backup():
    """Carica i dati dal backup all'avvio del bot"""
    global dati_punti, ultimo_aggiornamento
    if not os.path.exists(BACKUP_FILE):
        logger.info("ℹ️  Nessun backup trovato, parto da zero")
        return

    try:
        with open(BACKUP_FILE, "rb") as f:
            backup = pickle.load(f)
            dati_punti = backup.get("dati", {})
            ultimo_aggiornamento = backup.get("aggiornato")
        logger.info(f"✅ Backup caricato: {len(dati_punti)} utenti")
    except Exception as e:
        logger.exception(f"❌ Errore caricamento backup: {e}")


# ═══════════════════════════════════════════════════════════════
# 📥  PARSER EXCEL
# ═══════════════════════════════════════════════════════════════

def processa_excel(file_url: str) -> dict:
    """
    Scarica un file Excel da Slack e lo converte in dizionario.

    Returns:
        dict: {email: {"nome": str, "punti": int}}
    """
    headers = {"Authorization": f"Bearer {os.environ['SLACK_BOT_TOKEN']}"}
    response = requests.get(file_url, headers=headers, timeout=30)
    response.raise_for_status()

    df = pd.read_excel(io.BytesIO(response.content))
    df.columns = df.columns.str.lower().str.strip()

    required = {'email', 'nome', 'punti'}
    if not required.issubset(df.columns):
        missing = required - set(df.columns)
        raise ValueError(
            f"Colonne mancanti: {missing}. "
            f"Il file deve avere: email, nome, punti"
        )

    df['email'] = df['email'].astype(str).str.lower().str.strip()
    df['nome'] = df['nome'].astype(str).str.strip()
    df['punti'] = pd.to_numeric(df['punti'], errors='coerce').fillna(0).astype(int)

    # Rimuovi righe con email vuota
    df = df[df['email'].str.contains('@', na=False)]

    return {
        row['email']: {"nome": row['nome'], "punti": int(row['punti'])}
        for _, row in df.iterrows()
    }


# ═══════════════════════════════════════════════════════════════
# 📤  ADMIN: UPLOAD EXCEL NEL CANALE ADMIN
# ═══════════════════════════════════════════════════════════════

@app.event("file_shared")
def gestisci_upload(event, client):
    """Rileva upload di file Excel nel canale admin"""
    global dati_punti, ultimo_aggiornamento

    try:
        file_info = client.files_info(file=event['file_id'])['file']
    except Exception as e:
        logger.error(f"Impossibile leggere info file: {e}")
        return

    # Processa solo file nel canale admin
    if ADMIN_CHANNEL not in file_info.get('channels', []):
        return

    # Solo file Excel
    if not file_info['name'].lower().endswith(('.xlsx', '.xls')):
        client.chat_postMessage(
            channel=ADMIN_CHANNEL,
            text="⚠️ Carica solo file Excel (.xlsx o .xls)"
        )
        return

    try:
        nuovi_dati = processa_excel(file_info['url_private_download'])

        if not nuovi_dati:
            raise ValueError("Il file non contiene righe valide")

        dati_punti = nuovi_dati
        ultimo_aggiornamento = datetime.now()
        salva_backup()

        client.chat_postMessage(
            channel=ADMIN_CHANNEL,
            text=(
                f"✅ *File ricevuto!*\n"
                f"• Utenti caricati: *{len(dati_punti)}*\n"
                f"• File: `{file_info['name']}`\n"
                f"• Totale punti: {sum(u['punti'] for u in dati_punti.values())}\n"
                f"\nGli utenti possono ora usare `/punti` 🎯"
            )
        )
        logger.info(f"✅ Caricati {len(dati_punti)} utenti da {file_info['name']}")

    except Exception as e:
        logger.exception("❌ Errore processamento Excel")
        client.chat_postMessage(
            channel=ADMIN_CHANNEL,
            text=(
                f"❌ *Errore nel file:*\n```{str(e)}```\n"
                f"Verifica che il file abbia le colonne: `email`, `nome`, `punti`"
            )
        )


# ═══════════════════════════════════════════════════════════════
# 👤  COMANDO: /punti
# ═══════════════════════════════════════════════════════════════

@app.command("/punti")
def mostra_punti(ack, command, client):
    ack()

    if not dati_punti:
        client.chat_postMessage(
            channel=command['user_id'],
            text="⚠️ Nessun dato disponibile. Chiedi a un admin di caricare il file."
        )
        return

    try:
        user_info = client.users_info(user=command['user_id'])
        email = user_info['user']['profile'].get('email', '').lower()

        if email in dati_punti:
            info = dati_punti[email]
            testo = f"👋 Ciao *{info['nome']}*! Hai *{info['punti']} punti* 🎯"
        else:
            testo = (
                f"⚠️ Non trovo i tuoi dati (email: `{email}`).\n"
                f"Contatta un admin per essere aggiunto al file."
            )

        client.chat_postMessage(channel=command['user_id'], text=testo)

    except Exception as e:
        logger.exception("❌ Errore in /punti")
        client.chat_postMessage(
            channel=command['user_id'],
            text=f"❌ Errore nel recupero dati. Riprova più tardi."
        )


# ═══════════════════════════════════════════════════════════════
# 🏆  COMANDO: /classifica
# ═══════════════════════════════════════════════════════════════

@app.command("/classifica")
def classifica(ack, command, client):
    ack()

    if not dati_punti:
        client.chat_postMessage(
            channel=command['user_id'],
            text="⚠️ Nessun dato disponibile."
        )
        return

    top10 = sorted(
        dati_punti.values(),
        key=lambda x: x['punti'],
        reverse=True
    )[:10]

    medaglie = ["🥇", "🥈", "🥉"] + ["🏅"] * 7
    righe = [
        f"{medaglie[i]} *{u['nome']}* — {u['punti']} pt"
        for i, u in enumerate(top10)
    ]

    client.chat_postMessage(
        channel=command['user_id'],
        text="🏆 *Top 10 Classifica*\n" + "\n".join(righe)
    )


# ═══════════════════════════════════════════════════════════════
# 📊  COMANDO: /stato (health check)
# ═══════════════════════════════════════════════════════════════

@app.command("/stato")
def stato(ack, command, client):
    ack()

    uptime = datetime.now() - AVVIO
    ore = int(uptime.total_seconds() // 3600)
    minuti = int((uptime.total_seconds() % 3600) // 60)

    ultimo = (
        ultimo_aggiornamento.strftime('%d/%m/%Y %H:%M')
        if ultimo_aggiornamento else "mai"
    )

    testo = (
        f"🟢 *Bot online*\n"
        f"• Uptime: {ore}h {minuti}m\n"
        f"• Utenti caricati: *{len(dati_punti)}*\n"
        f"• Ultimo aggiornamento dati: {ultimo}\n"
        f"• Avviato: {AVVIO.strftime('%d/%m/%Y %H:%M')}"
    )

    client.chat_postMessage(channel=command['user_id'], text=testo)


# ═══════════════════════════════════════════════════════════════
# ❓  COMANDO: /aiuto
# ═══════════════════════════════════════════════════════════════

@app.command("/aiuto")
def aiuto(ack, command, client):
    ack()
    client.chat_postMessage(
        channel=command['user_id'],
        text=(
            "*🤖 Comandi disponibili:*\n\n"
            "• `/punti` — Mostra i tuoi punti attuali\n"
            "• `/classifica` — Top 10 utenti per punti\n"
            "• `/stato` — Stato del bot e ultimo aggiornamento\n"
            "• `/aiuto` — Mostra questo messaggio\n\n"
            "_Per aggiornare i punti, un admin carica un file Excel "
            "nel canale admin._"
        )
    )


# ═══════════════════════════════════════════════════════════════
# 🚀  AVVIO
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    logger.info("🚀 Avvio Slack Punti Bot...")
    carica_backup()
    logger.info(f"📡 Admin channel: {ADMIN_CHANNEL}")
    SocketModeHandler(app, os.environ["SLACK_APP_TOKEN"]).start()