import os
import logging
import requests
import pandas as pd
import plotly.express as px
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from dotenv import load_dotenv
from bs4 import BeautifulSoup

load_dotenv()  # Load variabel dari .env

MIRO_TOKEN = os.getenv("MIRO_TOKEN")
BOARD_ID = os.getenv("BOARD_ID")
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")

headers = {
    "Authorization": f"Bearer {MIRO_TOKEN}"
}

logging.basicConfig(level=logging.INFO)

def fetch_sticky_notes():
    url = f"https://api.miro.com/v2/boards/{BOARD_ID}/items?type=sticky_note"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()["data"]
    else:
        logging.error(f"Gagal ambil data dari Miro: {response.status_code} {response.text}")
        return []

def parse_note(note):
    try:
        content = note["data"]["content"]
        
        # Menghapus elemen HTML dengan BeautifulSoup
        clean_content = BeautifulSoup(content, "lxml").get_text()

        # Mengecek apakah format sticky note sesuai
        if "|" not in clean_content:
            logging.warning(f"Sticky note tidak sesuai format: {clean_content}")
            return None
        
        # Membagi konten berdasarkan "|"
        parts = [part.strip() for part in clean_content.split("|")]
        
        if len(parts) != 4:
            logging.warning(f"Sticky note tidak lengkap: {clean_content}")
            return None
        
        task = parts[0]
        start = parts[1]
        end = parts[2]
        person = parts[3]

        # Memastikan bahwa tanggalnya valid (contoh format: 2025-05-10)
        if len(start) == 10 and len(end) == 10:
            return {
                "Task": task,
                "Start": start,
                "End": end,
                "Person": person
            }
        else:
            logging.warning(f"Format tanggal tidak valid: {start} - {end}")
            return None
    except Exception as e:
        logging.warning(f"Gagal parsing note: {note['data']['content']} ({e})")
        return None

def generate_gantt(parsed_data, output_file="gantt_chart.png"):
    df = pd.DataFrame(parsed_data)
    fig = px.timeline(df, x_start="Start", x_end="End", y="Task", color="Person", title="Gantt Chart from Miro")
    fig.update_yaxes(autorange="reversed")
    fig.update_layout(margin=dict(l=20, r=20, t=40, b=20))
    fig.write_image(output_file)
    return output_file

async def send_gantt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user.first_name
    await context.bot.send_message(chat_id=update.effective_chat.id, text=f"Halo {user}, sedang memproses Gantt chart...")

    notes = fetch_sticky_notes()
    parsed = [parse_note(n) for n in notes if parse_note(n)]

    if not parsed:
        await update.message.reply_text("❌ Tidak ada sticky notes yang bisa diproses.")
        return

    file_path = generate_gantt(parsed)
    with open(file_path, 'rb') as photo:
        await context.bot.send_photo(chat_id=update.effective_chat.id, photo=photo, caption="📊 Gantt chart dari Miro")

async def view_sticky_notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    notes = fetch_sticky_notes()
    parsed = [parse_note(n) for n in notes if parse_note(n)]
    
    if not parsed:
        await update.message.reply_text("❌ Tidak ada sticky notes yang tersedia.")
        return

    # Format sticky notes yang berhasil diambil
    message = "📋 Daftar Sticky Notes:\n\n"
    for note in parsed:
        if 'Task' in note:
            message += f"Task: {note['Task']}\n"
            message += f"Start Date: {note['Start']}\n"
            message += f"End Date: {note['End']}\n"
            message += f"Assignee: {note['Person']}\n\n"
        else:
            message += "Sticky note tidak sesuai format atau gagal diproses.\n\n"
    
    await update.message.reply_text(message)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Selamat datang! Kirim /gantt untuk mendapatkan Gantt chart dari Miro atau /view_sticky_notes untuk melihat sticky notes.")

def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("gantt", send_gantt))
    app.add_handler(CommandHandler("view_sticky_notes", view_sticky_notes))
    app.run_polling()

if __name__ == "__main__":
    main()
