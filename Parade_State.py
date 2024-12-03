import pandas as pd
from typing import Final
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters
from datetime import datetime

# Define your bot's token and Excel file path
BOT_TOKEN = "7909643356:AAEw_2SMUhX0eFb_zuxXETZGUijLNgMWOw8"
BOT_USERNAME: Final = 'TFC_Parade_State_Bot'
EXCEL_FILE = "sample.xlsx"



# Load the Excel sheet
df = pd.read_excel(EXCEL_FILE, usecols="C:D", header=0) # Read all rows in Column C and D
df["Rank"] = df["Rank"].astype(str).str.strip().str.lower()  # Strip whitespace and convert to lowercase
df["Name"] = df["Name"].astype(str).str.strip().str.lower()  # Strip whitespace and convert to lowercase

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    # Starts the bot and asks for user input.
    await update.message.reply_text("Welcome! Please enter your rank and name in this format:\n\nRank Name\n\nFor example: Captain JJ")

async def verify_user(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    # Verifies the user's rank and name against the Excel sheet.
    user_input = update.message.text.strip()
    
    try:
        # Split input to get rank and name
        rank, *name_parts = user_input.split()
        
        # Ensure enough entries just in case of repeated names
        if len(name_parts) < 2:
            await update.message.reply_text("Please enter at least two parts (don't have to be surname) of your name along with your rank. For example: MAJ John Doe.")
            return
        
        rank = rank.lower().strip()
        name_parts = [part.lower().strip() for part in name_parts]
        
        # Check if the rank and at least two name parts match 
        match = df[df["Rank"] == rank]

        if not match.empty:
            for _, row in match.iterrows():
                stored_name_parts = row["Name"].split()
                matches = sum(1 for part in name_parts if part in stored_name_parts)

                if matches >= 2:
                    # If there are at least two matches, send personalized message with options
                    greeting = f"Good day {rank.upper()} {row['Name'].upper()}, would you like to update parade state?"
            # Create inline keyboard with two options
            keyboard = [
                [InlineKeyboardButton("Update parade state", callback_data="update_parade")],
                [InlineKeyboardButton("Exit", callback_data="exit")]
            ]
            reply_markup = InlineKeyboardMarkup(keyboard)
            
            await update.message.reply_text(greeting, reply_markup=reply_markup)
            return
        else:
            # No match found
            await update.message.reply_text("Sorry, rank and name not found. Please try again.")
    
    except ValueError:
        await update.message.reply_text("Invalid format. Please enter your rank and name in this format:\n\nRank Name\n\nFor example: Captain John")

async def button(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handles button presses."""
    query = update.callback_query
    await query.answer()  # Acknowledge the button press
    formatted_date = datetime.now().strftime("%A, %d%m%y")
    if query.data == "update_parade":
        await query.edit_message_text(text=("Please update the parade state with the following format:\n\n1. 3WO Martin Sng (Reason, Location, Duration)\n\n" + formatted_date))       
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
