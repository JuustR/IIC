import logging
from telegram import Update
from telegram.ext import filters, MessageHandler, ApplicationBuilder, CommandHandler, ContextTypes

TOKEN = '7351408788:AAHy0Yly70lQB_jkM1LfaQatKUXox1qtaLY'
BOT_USERNAME = '@justTasksBot'

# Create spam defence

# logging.basicConfig(
#     format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
#     level=logging.INFO
# )


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="Hello! I'm a task bot which help you to remember your plans"
    )
async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="In future"
    )
async def custom_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="Custom command"
    )

# async def echo(update: Update, context: ContextTypes.DEFAULT_TYPE):
#         await context.bot.send_message(
#             chat_id=update.effective_chat.id,
#             text=update.message.text
#         )
#
# async def caps(update: Update, context: ContextTypes.DEFAULT_TYPE):
#     text_caps = ' '.join(context.args).upper()
#     await context.bot.send_message(
#         chat_id=update.effective_chat.id,
#         text=text_caps
#     )

def handle_response(text: str):
    processed: str = text.lower()

    if 'hello' == processed: # I can add other variants
        return 'Hey there!'

    if 'how are you' == processed or 'what\'s up?' == processed:
        return 'I\'m fine'

    return 'Sorry, I dont understand what you wrote'
    # Add return 'I love smthng'

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message_type: str = update.message.chat.type
    text: str = update.message.text

    print(f'User ({update.message.chat.id}) in {message_type}: "{text}"') # log from bot

    if message_type == 'group':
        if BOT_USERNAME in text:
            new_text: str = text.replace(BOT_USERNAME, '').strip()
            response: str = handle_response(new_text)
        else:
            return
    else:
        response: str = handle_response(text)

    print('Bot:', response)
    await update.message.reply_text(response)

async def error(update: Update, context: ContextTypes.DEFAULT_TYPE):
    print(f'Update {update} caused error {context.error}')

# async def unknown(update: Update, context: ContextTypes.DEFAULT_TYPE):
#     await context.bot.send_message(
#         chat_id=update.effective_chat.id,
#         text="Sorry, I didn't understand that command."
#     )

if __name__ == '__main__':
    print('Starting Bot...')
    application = ApplicationBuilder().token(TOKEN).build()

    #Commands
    application.add_handler(CommandHandler('start', start))
    application.add_handler(CommandHandler('help', help_command))
    application.add_handler(CommandHandler('custom', custom_command))

    # echo_handler = MessageHandler(filters.TEXT & (~filters.COMMAND), echo)
    # application.add_handler(start_handler)
    # application.add_handler(echo_handler)
    #
    # caps_handler = CommandHandler('caps', caps)
    # application.add_handler(caps_handler)


    message_handler = MessageHandler(filters.TEXT, handle_message)
    application.add_handler(message_handler)

    application.add_error_handler(error)

    # Unknown always has been at the end of code
    # unknown_handler = MessageHandler(filters.COMMAND, unknown)
    # application.add_handler(unknown_handler)

    print('Polling...')
    application.run_polling(poll_interval=3)
