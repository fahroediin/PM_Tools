import os
import logging
import requests
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, CallbackQueryHandler, ContextTypes
from bs4 import BeautifulSoup
from datetime import datetime
from dotenv import load_dotenv
from openpyxl import Workbook

load_dotenv()

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
        style = note.get("style", {})
        color = style.get("fillColor", "")

        clean_content = BeautifulSoup(content, "lxml").get_text()

        if "|" not in clean_content:
            return None

        parts = [part.strip() for part in clean_content.split("|")]
        if len(parts) != 4:
            return None

        task, start, end, person = parts

        return {
            "Task": task,
            "Start": start,
            "End": end,
            "Person": person,
            "Type": "Pending"
        }
    except Exception as e:
        logging.warning(f"Gagal parsing note: {e}")
        return None

def generate_excel_gantt(data, image_path="gantt_chart_styled.png", excel_path="gantt_chart_styled.xlsx"):
    df = pd.DataFrame(data)
    df["Start"] = pd.to_datetime(df["Start"])
    df["End"] = pd.to_datetime(df["End"])
    df["Duration"] = (df["End"] - df["Start"]).dt.days

    colors = {"Critical Path": "#FF0000", "Floating Task": "#1E90FF", "Pending": "#808080"}

    start_min = df["Start"].min()
    end_max = df["End"].max()
    total_days = (end_max - start_min).days + 1

    width_per_day = 0.3
    height_per_task = 0.5
    fig_width = max(10, total_days * width_per_day)
    fig_height = max(6, len(df) * height_per_task)

    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    yticks = []
    ylabels = []

    for i, row in df.iterrows():
        ax.barh(
            y=i,
            width=row["Duration"],
            left=row["Start"],
            color=colors.get(row["Type"], "gray"),
            edgecolor="black"
        )
        ax.text(
            row["End"] + pd.Timedelta(days=1),
            i,
            row["Person"],
            va="center",
            fontsize=9,
            color="black"
        )
        yticks.append(i)
        ylabels.append(row["Task"])

    ax.set_yticks(yticks)
    ax.set_yticklabels(ylabels)
    ax.set_xlabel("Timeline")
    ax.set_title("GANTT CHART")
    ax.grid(True, axis='x', which='both', linestyle='--', linewidth=0.5)
    ax.invert_yaxis()

    ax.set_xlim(start_min - pd.Timedelta(days=1), end_max + pd.Timedelta(days=2))
    ax.xaxis.set_major_locator(plt.MaxNLocator(integer=True))
    fig.autofmt_xdate(rotation=45)

    legend_elements = [Patch(facecolor=color, edgecolor='black', label=label) for label, color in colors.items()]
    ax.legend(handles=legend_elements, loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3)

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

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Halo! 👋\nGunakan perintah /gantt untuk mulai membuat Gantt Chart dari sticky notes Miro.")

async def send_gantt(update_or_query, context: ContextTypes.DEFAULT_TYPE):
    if isinstance(update_or_query, Update):
        chat = update_or_query.message
    else:
        chat = update_or_query.message

    if "parsed_notes" not in context.user_data:
        notes = fetch_sticky_notes()
        parsed = [parse_note(n) for n in notes if parse_note(n)]
        if not parsed:
            await chat.reply_text("❌ Tidak ada sticky notes yang bisa diproses.")
            return

        context.user_data["parsed_notes"] = parsed

    context.user_data["selected_tasks"] = context.user_data.get("selected_tasks", {"Critical Path": set(), "Floating Task": set()})
    context.user_data["current_type"] = context.user_data.get("current_type", None)

    keyboard = [
        [InlineKeyboardButton("🔴 Pilih Critical Path", callback_data="set_type_critical")],
        [InlineKeyboardButton("🔵 Pilih Floating Task", callback_data="set_type_floating")]
    ]

    if any(len(tasks) > 0 for tasks in context.user_data["selected_tasks"].values()):
        keyboard.append([InlineKeyboardButton("✅ Generate Gantt Chart", callback_data="generate_gantt")])

    await chat.reply_text("Pilih tipe task yang ingin diklasifikasikan:", reply_markup=InlineKeyboardMarkup(keyboard))

async def handle_buttons(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    parsed_notes = context.user_data.get("parsed_notes", [])
    selected_tasks = context.user_data.get("selected_tasks", {})

    if data == "noop":
        return

    if data == "set_type_critical":
        context.user_data["current_type"] = "Critical Path"
    elif data == "set_type_floating":
        context.user_data["current_type"] = "Floating Task"
    elif data == "done_selecting":
        await query.message.reply_text("✅ Pemilihan task selesai. Kamu bisa memilih tipe task lain atau generate Gantt Chart.")
        await send_gantt(query, context)
        return
    elif data == "generate_gantt":
        selected_notes = []
        for tipe, indices in selected_tasks.items():
            for idx in indices:
                note = parsed_notes[idx].copy()
                note["Type"] = tipe
                selected_notes.append(note)

        if not selected_notes:
            await query.message.reply_text("⚠️ Belum ada task yang dipilih.")
            return

        keyboard = [
            [InlineKeyboardButton("📊 PNG", callback_data="generate_png")],
            [InlineKeyboardButton("📊 Excel", callback_data="generate_excel")]
        ]

        await query.edit_message_text("Pilih format untuk Gantt Chart:", reply_markup=InlineKeyboardMarkup(keyboard))
        context.user_data["selected_notes"] = selected_notes
        return
    elif data == "generate_png" or data == "generate_excel":
        selected_notes = context.user_data.get("selected_notes", [])

        if not selected_notes:
            await query.message.reply_text("⚠️ Belum ada data yang dipilih.")
            return

        png_file, excel_file = generate_excel_gantt(selected_notes)
        await query.message.reply_photo(photo=open(png_file, 'rb'), caption="📊 Gantt Chart (PNG)")

        if data == "generate_excel":
            await query.message.reply_document(document=open(excel_file, 'rb'), filename="gantt_chart_styled.xlsx")

        await query.edit_message_text("Gantt Chart telah dibuat dan dikirimkan. Terima kasih!")
        return
    elif data.startswith("toggle_"):
        idx = int(data.split("_")[1])
        current_type = context.user_data["current_type"]
        if current_type is None:
            await query.edit_message_text("⚠️ Harap pilih tipe task terlebih dahulu.")
            return

        other_type = "Floating Task" if current_type == "Critical Path" else "Critical Path"
        if idx in selected_tasks[other_type]:
            await query.answer(f"❌ Task ini sudah dipilih sebagai {other_type}.", show_alert=True)
            return

        if idx in selected_tasks[current_type]:
            selected_tasks[current_type].remove(idx)
        else:
            selected_tasks[current_type].add(idx)

        context.user_data["selected_tasks"] = selected_tasks

    current_type = context.user_data.get("current_type", "")
    keyboard = []
    for idx, note in enumerate(parsed_notes):
        label = note["Task"]
        selected_in_current = idx in selected_tasks[current_type]
        selected_in_other = idx in selected_tasks["Floating Task" if current_type == "Critical Path" else "Critical Path"]

        if selected_in_other:
            label = f"❌ {label} (sudah dipilih)"
            button = InlineKeyboardButton(label, callback_data="noop")
        else:
            prefix = "✅ " if selected_in_current else ""
            button = InlineKeyboardButton(f"{prefix}{label}", callback_data=f"toggle_{idx}")

        keyboard.append([button])

    keyboard.append([InlineKeyboardButton("✔️ Selesai Memilih", callback_data="done_selecting")])

    await query.edit_message_text(
        f"Klik task untuk memilih sebagai *{current_type}*: ",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="Markdown"
    )

def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("gantt", send_gantt))
    app.add_handler(CallbackQueryHandler(handle_buttons))
    app.run_polling()

if __name__ == "__main__":
    main()
