import pandas as pd
from typing import Final
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters

# Define your bot's token and Excel file path
BOT_TOKEN = "7909643356:AAEw_2SMUhX0eFb_zuxXETZGUijLNgMWOw8"
BOT_USERNAME: Final = 'TFC_Parade_State_Bot'
EXCEL_FILE = "13th ASCC Liaison TP 260124.xlsx"

# Load the Excel sheet
df = pd.read_excel(EXCEL_FILE, usecols="C:D")  # Load only relevant columns

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Starts the bot and asks for user input."""
    await update.message.reply_text("Welcome! Please enter your rank and name in this format:\n\nRank Name\n\nFor example: Captain JJ")

async def verify_user(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Verifies the user's rank and name against the Excel sheet."""
    user_input = update.message.text.strip()
    
    try:
        # Split input to get rank and name
        rank, name = user_input.split(maxsplit=1)
        
        # Check if the rank and name match any entry in the Excel data
        match = df[(df["Rank"] == rank) & (df["Name"] == name)]
        
        if not match.empty:
            # If there's a match, send personalized message with options
            greeting = f"Good evening {rank} {name}, would you like to update the parade state?"
            
            # Create inline keyboard with two options
            keyboard = [
                [InlineKeyboardButton("Update parade state", callback_data="update_parade")],
                [InlineKeyboardButton("Exit", callback_data="exit")]
            ]
            reply_markup = InlineKeyboardMarkup(keyboard)
            
            await update.message.reply_text(greeting, reply_markup=reply_markup)
        else:
            # No match found
            await update.message.reply_text("Sorry, rank and name not found. Please try again.")
    
    except ValueError:
        await update.message.reply_text("Invalid format. Please enter your rank and name in this format:\n\nRank Name\n\nFor example: Captain John")

async def button(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handles button presses."""
    query = update.callback_query
    await query.answer()  # Acknowledge the button press

    if query.data == "update_parade":
        await query.edit_message_text(text="Please proceed to update the parade state.")
        # Add code here if you want to handle updates
    elif query.data == "exit":
        await query.edit_message_text(text="Thank you! Exiting.")


if __name__ == "__main__":
    """Run the bot."""
    print('Starting bot...')
    app = Application.builder().token(BOT_TOKEN).build()

    # Command to start bot
    app.add_handler(CommandHandler("Start", start))

    # Message handler to verify user input
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, verify_user))

    # Callback handler for button presses
    app.add_handler(CallbackQueryHandler(button))

    # Run the bot
    print('Polling...')
    app.run_polling(poll_interval = 3)