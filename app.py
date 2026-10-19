"""
Slack Punti Bot v3
──────────────────
Features:
  - /punti <alias>  → mostra saldo e salva alias per futuro uso
  - /punti          → usa l'alias salvato (dopo prima registrazione)
  - Saldo dettagliato: balance, assigned, spent
  - Classifica con dettagli
  - Backup automatico dei dati + mapping Slack→alias
"""

import os
import io
import sys
import pickle
import logging
import requests
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

# ═══════════════════════════════════════════════════════════════
# 🔧 FIX ENCODING WINDOWS
# ═══════════════════════════════════════════════════════════════
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# ═══════════════════════════════════════════════════════════════
# ⚙️  CONFIGURAZIONE E LOGGING
# ═══════════════════════════════════════════════════════════════

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('bot.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)

REQUIRED_ENV = ["SLACK_BOT_TOKEN", "SLACK_APP_TOKEN", "ADMIN_CHANNEL_ID"]
for var in REQUIRED_ENV:
    if not os.environ.get(var):
        raise RuntimeError(f"Variabile d'ambiente mancante: {var}")

ADMIN_CHANNEL = os.environ["ADMIN_CHANNEL_ID"]
DATI_BACKUP_FILE = "ultimo_dati.pkl"
ALIAS_BACKUP_FILE = "alias_mapping.pkl"

app = App(token=os.environ["SLACK_BOT_TOKEN"])

# ═══════════════════════════════════════════════════════════════
# 💾  STATO IN MEMORIA
# ═══════════════════════════════════════════════════════════════

# {alias: {"balance": int, "assigned": int, "spent": int}}
dati_punti = {}

# {slack_user_id: alias}  → mapping persistente
alias_mapping = {}

ultimo_aggiornamento = None
AVVIO = datetime.now()


def salva_dati_punti():
    try:
        with open(DATI_BACKUP_FILE, "wb") as f:
            pickle.dump({
                "dati": dati_punti,
                "aggiornato": ultimo_aggiornamento
            }, f)
        logger.info(f"Backup dati punti salvato ({len(dati_punti)} utenti)")
    except Exception as e:
        logger.exception(f"Errore salvataggio dati: {e}")


def carica_dati_punti():
    global dati_punti, ultimo_aggiornamento
    if not os.path.exists(DATI_BACKUP_FILE):
        logger.info("Nessun backup dati trovato")
        return
    try:
        with open(DATI_BACKUP_FILE, "rb") as f:
            backup = pickle.load(f)
            dati_punti = backup.get("dati", {})
            ultimo_aggiornamento = backup.get("aggiornato")
        logger.info(f"Backup dati caricato: {len(dati_punti)} utenti")
    except Exception as e:
        logger.exception(f"Errore caricamento dati: {e}")


def salva_alias_mapping():
    try:
        with open(ALIAS_BACKUP_FILE, "wb") as f:
            pickle.dump(alias_mapping, f)
    except Exception as e:
        logger.exception(f"Errore salvataggio alias mapping: {e}")


def carica_alias_mapping():
    global alias_mapping
    if not os.path.exists(ALIAS_BACKUP_FILE):
        logger.info("Nessun alias mapping salvato")
        return
    try:
        with open(ALIAS_BACKUP_FILE, "rb") as f:
            alias_mapping = pickle.load(f)
        logger.info(f"Alias mapping caricato: {len(alias_mapping)} utenti registrati")
    except Exception as e:
        logger.exception(f"Errore caricamento alias mapping: {e}")


# ═══════════════════════════════════════════════════════════════
# 📥  PARSER EXCEL (alias + balance + assigned + spent)
# ═══════════════════════════════════════════════════════════════

def processa_excel(file_url: str) -> dict:
    """
    Legge Excel ed estrae: alias, balance, assigned, spent.
    Ignora employee_id, txn_count e altre colonne.
    """
    headers = {"Authorization": f"Bearer {os.environ['SLACK_BOT_TOKEN']}"}
    response = requests.get(file_url, headers=headers, timeout=30)
    response.raise_for_status()

    df = pd.read_excel(io.BytesIO(response.content))
    df.columns = df.columns.str.lower().str.strip()

    required = {'alias', 'balance'}
    if not required.issubset(df.columns):
        missing = required - set(df.columns)
        raise ValueError(
            f"Colonne obbligatorie mancanti: {missing}. "
            f"Richieste almeno: alias, balance"
        )

    # Alias e balance obbligatori, assigned e spent opzionali
    df['alias'] = df['alias'].astype(str).str.lower().str.strip()
    df['balance'] = pd.to_numeric(df['balance'], errors='coerce').fillna(0).astype(int)

    if 'assigned' in df.columns:
        df['assigned'] = pd.to_numeric(df['assigned'], errors='coerce').fillna(0).astype(int)
    else:
        df['assigned'] = 0

    if 'spent' in df.columns:
        df['spent'] = pd.to_numeric(df['spent'], errors='coerce').fillna(0).astype(int)
    else:
        df['spent'] = 0

    # Filtra righe con alias validi
    df = df[df['alias'].notna() & (df['alias'] != '') & (df['alias'] != 'nan')]

    return {
        row['alias']: {
            "balance": int(row['balance']),
            "assigned": int(row['assigned']),
            "spent": int(row['spent'])
        }
        for _, row in df.iterrows()
    }


def processa_file_da_evento(file_info, client, channel_id):
    global dati_punti, ultimo_aggiornamento

    if not file_info['name'].lower().endswith(('.xlsx', '.xls')):
        client.chat_postMessage(
            channel=channel_id,
            text="⚠️ Carica solo file Excel (.xlsx o .xls)"
        )
        return

    try:
        nuovi_dati = processa_excel(file_info['url_private_download'])

        if not nuovi_dati:
            raise ValueError("Il file non contiene righe valide")

        dati_punti = nuovi_dati
        ultimo_aggiornamento = datetime.now()
        salva_dati_punti()

        totale_balance = sum(u['balance'] for u in dati_punti.values())
        totale_assigned = sum(u['assigned'] for u in dati_punti.values())
        totale_spent = sum(u['spent'] for u in dati_punti.values())

        client.chat_postMessage(
            channel=channel_id,
            text=(
                f"✅ *File ricevuto!*\n"
                f"• Utenti caricati: *{len(dati_punti)}*\n"
                f"• File: `{file_info['name']}`\n"
                f"• Totale balance: *{totale_balance:,}* pt\n"
                f"• Totale assegnati: *{totale_assigned:,}* pt\n"
                f"• Totale spesi: *{totale_spent:,}* pt\n"
                f"\nGli utenti possono usare `/punti <alias>` 🎯"
            )
        )
        logger.info(f"Caricati {len(dati_punti)} utenti da {file_info['name']}")

    except Exception as e:
        logger.exception("Errore processamento Excel")
        client.chat_postMessage(
            channel=channel_id,
            text=(
                f"❌ *Errore nel file:*\n```{str(e)}```\n"
                f"Verifica almeno le colonne: `alias`, `balance`"
            )
        )


# ═══════════════════════════════════════════════════════════════
# 📤  HANDLER UPLOAD
# ═══════════════════════════════════════════════════════════════

@app.event("file_shared")
def gestisci_file_shared(event, client):
    try:
        file_info = client.files_info(file=event['file_id'])['file']
    except Exception as e:
        logger.error(f"Impossibile leggere info file: {e}")
        return

    if ADMIN_CHANNEL not in file_info.get('channels', []):
        return

    processa_file_da_evento(file_info, client, ADMIN_CHANNEL)


@app.event("message")
def gestisci_message(event, client, logger):
    subtype = event.get('subtype')

    if subtype == 'file_share' and event.get('channel') == ADMIN_CHANNEL:
        files = event.get('files', [])
        for file_info in files:
            if 'url_private_download' not in file_info:
                try:
                    file_info = client.files_info(file=file_info['id'])['file']
                except Exception as e:
                    logger.error(f"Impossibile leggere info file: {e}")
                    continue
            processa_file_da_evento(file_info, client, ADMIN_CHANNEL)


# ═══════════════════════════════════════════════════════════════
# 🎨  HELPER: formatta messaggio saldo
# ═══════════════════════════════════════════════════════════════

def formatta_saldo(alias: str, info: dict) -> str:
    """Costruisce il messaggio con dettagli saldo"""
    return (
        f"👋 Ciao *{alias}*!\n\n"
        f"💰 Saldo attuale: *{info['balance']:,}* punti\n"
        f"📥 Assegnati: {info['assigned']:,} punti\n"
        f"📤 Spesi: {info['spent']:,} punti"
    )


# ═══════════════════════════════════════════════════════════════
# 👤  COMANDO: /punti [alias]
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

    slack_user_id = command['user_id']
    alias_input = command.get('text', '').strip().lower()

    # Pulizia input: rimuovi @, estrai parte prima di @amazon.com o spazi
    if alias_input:
        alias_input = alias_input.replace('@', '').split('.')[0].split(' ')[0]

    # ─── CASO 1: utente passa un alias ───
    if alias_input:
        if alias_input not in dati_punti:
            client.chat_postMessage(
                channel=slack_user_id,
                text=(
                    f"⚠️ Non trovo l'alias `{alias_input}` nel file.\n"
                    f"Verifica di aver scritto correttamente il tuo login Amazon."
                )
            )
            return

        # Salva/aggiorna il mapping Slack → alias
        alias_precedente = alias_mapping.get(slack_user_id)
        alias_mapping[slack_user_id] = alias_input
        salva_alias_mapping()

        # Messaggio contestuale
        if alias_precedente is None:
            prefisso = f"✅ Alias `{alias_input}` registrato! Lo userò la prossima volta.\n\n"
        elif alias_precedente != alias_input:
            prefisso = f"🔄 Alias aggiornato: `{alias_input}` (prima era `{alias_precedente}`)\n\n"
        else:
            prefisso = ""

        testo = prefisso + formatta_saldo(alias_input, dati_punti[alias_input])
        client.chat_postMessage(channel=slack_user_id, text=testo)
        return

    # ─── CASO 2: utente non passa alias → usa quello salvato ───
    alias_salvato = alias_mapping.get(slack_user_id)

    if alias_salvato is None:
        client.chat_postMessage(
            channel=slack_user_id,
            text=(
                "ℹ️ *Primo utilizzo?*\n"
                "Dimmi il tuo alias Amazon una volta sola:\n"
                "`/punti <tuo-alias>`\n\n"
                "_Esempio: `/punti mrossi`_\n"
                "_Dalla prossima volta ti basterà digitare `/punti`._"
            )
        )
        return

    if alias_salvato not in dati_punti:
        client.chat_postMessage(
            channel=slack_user_id,
            text=(
                f"⚠️ Il tuo alias salvato (`{alias_salvato}`) non è più nel file.\n"
                f"Riregistralo con: `/punti <nuovo-alias>`"
            )
        )
        return

    client.chat_postMessage(
        channel=slack_user_id,
        text=formatta_saldo(alias_salvato, dati_punti[alias_salvato])
    )


# ═══════════════════════════════════════════════════════════════
# 🧹  COMANDO: /dimentica  (reset alias salvato)
# ═══════════════════════════════════════════════════════════════

@app.command("/dimentica")
def dimentica_alias(ack, command, client):
    ack()
    slack_user_id = command['user_id']

    if slack_user_id in alias_mapping:
        vecchio = alias_mapping.pop(slack_user_id)
        salva_alias_mapping()
        client.chat_postMessage(
            channel=slack_user_id,
            text=f"🧹 Alias `{vecchio}` dimenticato. Usa `/punti <nuovo-alias>` per registrarne uno nuovo."
        )
    else:
        client.chat_postMessage(
            channel=slack_user_id,
            text="ℹ️ Non hai nessun alias registrato."
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
        dati_punti.items(),
        key=lambda x: x[1]['balance'],
        reverse=True
    )[:10]

    medaglie = ["🥇", "🥈", "🥉"] + ["🏅"] * 7
    righe = [
        f"{medaglie[i]} *{alias}* — {info['balance']:,} pt "
        f"_(spesi: {info['spent']:,})_"
        for i, (alias, info) in enumerate(top10)
    ]
    client.chat_postMessage(
        channel=command['user_id'],
        text="🏆 *Top 10 Classifica*\n" + "\n".join(righe)
    )


# ═══════════════════════════════════════════════════════════════
# 📊  COMANDO: /stato
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
    client.chat_postMessage(
        channel=command['user_id'],
        text=(
            f"🟢 *Bot online*\n"
            f"• Uptime: {ore}h {minuti}m\n"
            f"• Utenti nel file: *{len(dati_punti)}*\n"
            f"• Utenti Slack registrati: *{len(alias_mapping)}*\n"
            f"• Ultimo aggiornamento dati: {ultimo}\n"
            f"• Avviato: {AVVIO.strftime('%d/%m/%Y %H:%M')}"
        )
    )


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
            "• `/punti <alias>` — Mostra i tuoi punti e salva l'alias\n"
            "   _Esempio: `/punti mrossi`_\n"
            "• `/punti` — Usa l'alias già registrato\n"
            "• `/dimentica` — Rimuove il tuo alias registrato\n"
            "• `/classifica` — Top 10 utenti per punti\n"
            "• `/stato` — Stato del bot\n"
            "• `/aiuto` — Mostra questo messaggio\n\n"
            "💡 *Tip:* la prima volta usa `/punti tuoalias`, "
            "poi ti basta `/punti`."
        )
    )


# ═══════════════════════════════════════════════════════════════
# 🚀  AVVIO
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    logger.info("Avvio Slack Punti Bot v3...")
    carica_dati_punti()
    carica_alias_mapping()
    logger.info(f"Admin channel: {ADMIN_CHANNEL}")
    SocketModeHandler(app, os.environ["SLACK_APP_TOKEN"]).start()