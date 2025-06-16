from telegram.ext import Updater, CommandHandler, MessageHandler, filters, ConversationHandler, CallbackContext, Application
from telegram import Update
from config import get_token
import handlears

def main():
    token = get_token()
    dp = Application.builder().token(token).build()
    dp .add_handler(CommandHandler("start", handlears.start))
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("diplom_kochirma", handlears.start_diplom)],
        states={
            handlears.FILE: [MessageHandler(filters.Document.ALL, handlears.file_handler)],
            handlears.ID: [MessageHandler(filters.TEXT & ~filters.COMMAND, handlears.id_handler)],
            handlears.DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, handlears.date_handler)],
        },
        fallbacks=[],
        allow_reentry=True,
    )
    dp.add_handler(conv_handler)
    dp.run_polling()

if __name__ == "__main__":
    main()