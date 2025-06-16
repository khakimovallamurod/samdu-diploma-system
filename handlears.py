from telegram import Update
from telegram.ext import CallbackContext, ConversationHandler, ContextTypes
import os 
import write_docx

FILE, DATE = range(2)

async def start(update: Update, context: CallbackContext) -> None:
    
    user = update.effective_user
    await update.message.reply_text(
        f"Assalomu aleykum {user.full_name}!. Bakalavr dimlomdan ko'chirish jarayoniga xush kelibsiz! Jarayonini boshlash uchun /diplom_kochirma ni bosing.",)
    
async def start_diplom(update: Update, context: CallbackContext) -> int:
    await update.message.reply_text(
        "Iltimos, talabalaringizning ma'lumotlarini yuboring(excel formatida).",
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

    await update.message.reply_text("Iltimos, diplomingizning raqamini (B №) kiriting. Masalan: B №1234567")
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

    output_file = write_docx.main(file_path, date)
    await status_message.delete()
    await update.message.reply_document(document=output_file, caption="✅ Diplom ko'chirma hujjati")
