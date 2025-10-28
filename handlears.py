from telegram import Update
from telegram.ext import CallbackContext, ConversationHandler, ContextTypes
import os 
import write_docx
from telegram import InlineKeyboardMarkup, InlineKeyboardButton

CHOOSE, FILE, DATE = range(2)

async def start(update: Update, context: CallbackContext) -> None:
    user = update.effective_user
    await update.message.reply_text(
        f"Assalomu aleykum {user.full_name}! Bakalavr diplomdan ko'chirish jarayoniga xush kelibsiz!\n"
        "Jarayonni boshlash uchun /diplom_kochirma ni bosing."
    )

async def start_diplom(update: Update, context: CallbackContext) -> int:
    keyboard = [
        [InlineKeyboardButton("👨‍🎓 BAKALAVR", callback_data="bakalavr_student:BAKALAVR")],
        [InlineKeyboardButton("🎓 MAGISTR", callback_data="magistr_student:MAGISTR")]
    ]
    reply_keyboard = InlineKeyboardMarkup(keyboard)

    await update.message.reply_text(
        text="Ko'chirma turini tanlang!",
        reply_markup=reply_keyboard
    )
    return CHOOSE

async def send_file(update: Update, context: CallbackContext) -> int:
    query = update.callback_query
    await query.answer()

    context.user_data['choose_student'] = query.data

    await query.message.reply_text(
        "Iltimos, talabalaringizning ma'lumotlarini yuboring (excel formatida)."
    )
    return FILE

async def file_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    document = update.message.document
    if document is None or not document.file_name.endswith('.xlsx'):
        await update.message.reply_text("Iltimos, Excel faylni yuboring.")
        return FILE
    file = await document.get_file()

    file_path = os.path.join('downloads', document.file_name)
    os.makedirs('downloads', exist_ok=True)
    await file.download_to_drive(file_path)
    context.user_data['file_path'] = file_path

    await update.message.reply_text("File qabul qilindi. Iltimos, qaror qabul qilingan sanani kiriting (YYYY-MM-DD formatida)")
    return DATE

async def date_handler(update: Update, context: CallbackContext) -> int:
    if not update.message.text:
        await update.message.reply_text("Iltimos, sanani matn sifatida kiriting (YYYY-MM-DD formatida).")
        return DATE
    from datetime import datetime
    try:
        date = datetime.strptime(update.message.text, "%Y-%m-%d").date()
    except ValueError:
        await update.message.reply_text("Iltimos, sanani YYYY-MM-DD formatida kiriting.")
        return DATE

    context.user_data['date'] = date
    status_message = await update.message.reply_text("🛠 Diplom ko'chirma hujjati tayyorlanmoqda...")
    file_path = context.user_data.get('file_path')
    date = context.user_data.get('date')
    choose_kochirma = context.user_data['choose_student'].split(':')[1]
    output_file = write_docx.main(file_path, date, choose_kochirma)
    await status_message.delete()
    await update.message.reply_document(document=output_file, caption="✅ Diplom ko'chirma hujjati")
