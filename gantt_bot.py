import os
import logging
import requests
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
from bs4 import BeautifulSoup
from datetime import datetime
from dotenv import load_dotenv
from openpyxl import Workbook

# ------------------ CONFIGURATION ------------------
load_dotenv()
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")

# ------------------ LOGGER SETUP ------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ------------------ FETCHING DATA ------------------
def fetch_sticky_notes(board_id, headers):
    url = f"https://api.miro.com/v2/boards/{board_id}/items?type=sticky_note"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json().get("data", [])
    logger.error(f"Gagal ambil data dari Miro: {response.status_code} {response.text}")
    return []

def parse_note(note):
    try:
        content = note["data"].get("content", "")
        color = note.get("style", {}).get("fillColor", "")
        clean = BeautifulSoup(content, "lxml").get_text()
        parts = [x.strip() for x in clean.split("|")]
        if len(parts) != 4:
            return None
        return dict(zip(["Task", "Start", "End", "Person"], parts))
    except Exception as e:
        logger.warning(f"Parse error: {e}")
        return None

# ------------------ CHART GENERATION ------------------
def generate_chart(data, image_path, excel_path, chart_type="gantt"):
    df = pd.DataFrame(data)
    df["Start"] = pd.to_datetime(df["Start"])
    df["End"] = pd.to_datetime(df["End"])
    df["Duration"] = (df["End"] - df["Start"]).dt.days

    colors = {"Critical Path": "#FF0000", "Floating Task": "#1E90FF"}
    start_min, end_max = df["Start"].min(), df["End"].max()
    total_days = (end_max - start_min).days + 1

    fig, ax = plt.subplots(figsize=(max(10, total_days * 0.3), max(6, len(df) * 0.5)))
    yticks, ylabels = [], []

    for i, row in df.iterrows():
        left = row["Start"] if chart_type == "gantt" else i
        width = row["Duration"] if chart_type == "gantt" else 0.8
        color = colors.get(row.get("Type"), "#999999")

        ax.barh(i, width, left=left, color=color, edgecolor="black")
        ax.text(row["End"] + pd.Timedelta(days=1) if chart_type == "gantt" else i + 0.4, i, row["Person"],
                va="center", fontsize=9)
        yticks.append(i)
        ylabels.append(row["Task"])

    ax.set_yticks(yticks)
    ax.set_yticklabels(ylabels)
    ax.set_xlabel("Timeline")
    ax.set_title(chart_type.upper())
    ax.grid(True, axis='x', linestyle='--', linewidth=0.5)
    ax.invert_yaxis()

    if chart_type == "gantt":
        ax.set_xlim(start_min - pd.Timedelta(days=1), end_max + pd.Timedelta(days=2))
        fig.autofmt_xdate(rotation=45)

    ax.legend(handles=[Patch(facecolor=c, edgecolor='black', label=l) for l, c in colors.items()],
              loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)

    plt.tight_layout()
    plt.savefig(image_path)

    wb = Workbook()
    ws = wb.active
    ws.title = "Gantt Data"
    ws.append(["Task", "Start", "End", "Person", "Type"])
    for row in df.itertuples(index=False):
        ws.append([row.Task, row.Start.strftime('%Y-%m-%d'), row.End.strftime('%Y-%m-%d'), row.Person, row.Type])
    wb.save(excel_path)

    return image_path, excel_path

# ------------------ TELEGRAM BOT HANDLERS ------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Halo! Kirimkan *Miro Token* kamu terlebih dahulu:", parse_mode="Markdown")

async def handle_text_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    if "miro_token" not in context.user_data:
        context.user_data["miro_token"] = text
        context.user_data["headers"] = {"Authorization": f"Bearer {text}"}
        await update.message.reply_text("✅ Miro Token disimpan.\nSekarang, kirimkan *Board ID* kamu:")
    elif "board_id" not in context.user_data:
        context.user_data["board_id"] = text
        await update.message.reply_text("✅ Board ID disimpan.\nSekarang, kamu bisa jalankan perintah /gantt.")

async def send_gantt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    headers = context.user_data.get("headers")
    board_id = context.user_data.get("board_id")
    if not headers or not board_id:
        await update.message.reply_text("❗ Kamu belum mengirimkan Miro Token atau Board ID. Gunakan /start.")
        return

    notes_raw = fetch_sticky_notes(board_id, headers)
    notes = [parse_note(n) for n in notes_raw if parse_note(n)]
    if not notes:
        await update.message.reply_text("❌ Tidak ada sticky notes yang valid ditemukan.")
        return

    context.user_data["parsed_notes"] = notes
    context.user_data["selected_tasks"] = {"Critical Path": set(), "Floating Task": set()}
    context.user_data["current_type"] = "Critical Path"

    await update.message.reply_text(
        "✅ Data berhasil dimuat.\n\nSilakan pilih task berdasarkan kategori *Critical Path* atau *Floating Task*.",
    parse_mode="Markdown"
    )

    # Panggil handle_buttons secara langsung tanpa perlu tombol
    class FakeQuery:
        def __init__(self, message):
            self.data = ""
            self.message = message
        async def answer(self, *args, **kwargs):
            pass
        async def edit_message_text(self, text, reply_markup=None, **kwargs):
            await self.message.reply_text(text, reply_markup=reply_markup, **kwargs)

    fake_update = Update(update.update_id, callback_query=FakeQuery(update.message))
    await handle_buttons(fake_update, context)

async def handle_buttons(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    notes = context.user_data.get("parsed_notes", [])
    selected = context.user_data.get("selected_tasks", {})
    tipe = context.user_data.get("current_type")

    if data in ["set_type_critical", "set_type_floating"]:
        tipe = "Critical Path" if data == "set_type_critical" else "Floating Task"
        context.user_data["current_type"] = tipe

    elif data == "reset_all":
        selected["Critical Path"].clear()
        selected["Floating Task"].clear()
        await query.message.reply_text("🔁 Semua pilihan task telah dibatalkan.")
        return

    elif data == "done_selecting":
        summary = [f"*{t}*\n" + "\n".join(f"• {notes[i]['Task']}" for i in ids) if ids else f"*{t}*\n(tidak ada)"
                   for t, ids in selected.items()]
        await query.message.reply_text("📋 *Ringkasan task yang dipilih:*\n\n" + "\n\n".join(summary), parse_mode="Markdown")
        keyboard = [[InlineKeyboardButton("✅ Generate Chart", callback_data="generate_chart")]]
        await query.message.reply_text("Lanjutkan ke pembuatan chart:", reply_markup=InlineKeyboardMarkup(keyboard))
        return

    elif data.startswith("toggle_"):
        idx = int(data.split("_")[1])
        if not tipe:
            await query.answer("Pilih tipe task dulu!", show_alert=True)
            return

        other = "Floating Task" if tipe == "Critical Path" else "Critical Path"
        if idx in selected[other]:
            await query.answer("Task sudah dipilih di tipe lain!", show_alert=True)
            return

        if idx in selected[tipe]:
            selected[tipe].remove(idx)
        else:
            selected[tipe].add(idx)

    elif data == "generate_chart":
        context.user_data["chart_type"] = "gantt"
        keyboard = [[InlineKeyboardButton("📊 Gantt Chart", callback_data="chart_gantt")],
                    [InlineKeyboardButton("📅 Timeline", callback_data="chart_timeline")]]
        await query.message.reply_text("Pilih jenis chart:", reply_markup=InlineKeyboardMarkup(keyboard))
        return

    elif data.startswith("chart_"):
        context.user_data["chart_type"] = data.split("_")[1]
        keyboard = [[InlineKeyboardButton("🖼 PNG", callback_data="format_png")],
                    [InlineKeyboardButton("📄 Excel", callback_data="format_excel")]]
        await query.message.reply_text("Pilih format output:", reply_markup=InlineKeyboardMarkup(keyboard))
        return

    elif data.startswith("format_"):
        chart_type = context.user_data.get("chart_type", "gantt")
        selected_notes = []
        for tipe, ids in selected.items():
            for idx in ids:
                note = notes[idx].copy()
                note["Type"] = tipe
                selected_notes.append(note)

        if not selected_notes:
            await query.message.reply_text("❗ Belum ada task yang dipilih.")
            return

        image, excel = generate_chart(selected_notes, f"chart_{chart_type}.png", f"chart_{chart_type}.xlsx", chart_type)
        if data.endswith("png"):
            await query.message.reply_photo(open(image, "rb"), caption=f"{chart_type.upper()} Chart (PNG)")
        if data.endswith("excel"):
            await query.message.reply_document(open(excel, "rb"), filename=excel)
        return

    # Show updated task selection
    current = context.user_data.get("current_type", "Critical Path")
    keyboard = []
    for idx, note in enumerate(notes):
        label = note["Task"]
        if idx in selected["Critical Path"]:
            label = "🔴 " + label
        elif idx in selected["Floating Task"]:
            label = "🔵 " + label

        if idx in selected["Critical Path"] and current == "Floating Task" or \
           idx in selected["Floating Task"] and current == "Critical Path":
            button = InlineKeyboardButton(f"❌ {label}", callback_data="noop")
        else:
            button = InlineKeyboardButton(label, callback_data=f"toggle_{idx}")
        keyboard.append([button])

    keyboard += [[InlineKeyboardButton("✔️ Selesai", callback_data="done_selecting"),
                  InlineKeyboardButton("🔁 Reset", callback_data="reset_all")],
                 [InlineKeyboardButton("🔴 Critical", callback_data="set_type_critical"),
                  InlineKeyboardButton("🔵 Floating", callback_data="set_type_floating")]]
    await query.edit_message_text("Pilih task yang termasuk dalam kategori: *" + current + "*", reply_markup=InlineKeyboardMarkup(keyboard))

# ------------------ MAIN FUNCTION ------------------
def main():
    application = ApplicationBuilder().token(TELEGRAM_TOKEN).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("gantt", send_gantt))
    application.add_handler(CallbackQueryHandler(handle_buttons))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text_input))

    application.run_polling()

if __name__ == "__main__":
    main()