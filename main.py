"""
Discord Activity Tracker Bot
============================

An open-source Discord bot that tracks user activity in both text and voice channels
and stores data in an Excel workbook (`user_activity.xlsx`).

Features:
- Tracks message activity and voice joins
- Provides user-specific and basic server-wide activity summaries
- Built using discord.py and openpyxl

Author: Your Name
License: MIT
Repository: https://github.com/Fordywhat/discord-activity-tracker
"""

import os
from dotenv import load_dotenv
import discord
from discord.ext import commands
from discord import app_commands
import openpyxl
import datetime

# ---------------------------------------------------------------------------
# Loading Enviornment Variables
# ---------------------------------------------------------------------------
load_dotenv()

DISCORD_TOKEN = os.getenv("DISCORD_TOKEN")
GUILD_ID = int(os.getenv("GUILD_ID"))  # Convert to int for command registration


# ---------------------------------------------------------------------------
# Workbook Setup
# ---------------------------------------------------------------------------
''' Replace user-activity.xlsx in WORKBOOK_PATH to change workbook location '''
WORKBOOK_PATH = "user_activity.xlsx" 
workbook = openpyxl.load_workbook(WORKBOOK_PATH)
worksheet = workbook.active


# ---------------------------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------------------------
def get_activity(user_id, column):
    """Retrieve a specific activity cell value for a given user ID."""
    for row in worksheet.iter_rows():
        if row[0].value == user_id:
            if row[column].value is not None:
                return row[column].value
            else:
                return "N/A"
    return "N/A"


def updateSpreadsheet(user_id, column):
    """
    Update the activity spreadsheet for a user.

    Args:
        user_id (str): The user's Discord name or ID.
        column (int): The activity column index (1=message, 4=voice).
    """
    for row in worksheet.iter_rows():
        if row[0].value == user_id:
            
            # Shift "current" to "previous"
            row[column + 1].value = row[column].value
            
            # Increment total count
            if row[column + 2].value is not None:
                row[column + 2].value += 1  
            else:
                row[column + 2].value = 1  
                
            # Update current activity timestamp
            row[column].value = datetime.datetime.today().strftime('%Y-%m-%d %H:%M:%S')            
            workbook.save('user_activity.xlsx') 
            return  
            
    # If user not found, add a new row
    new_row = [user_id, None, None, 0, None, None, 0] 
    worksheet.append(new_row) 
    updateSpreadsheet(user_id, column)





# ---------------------------------------------------------------------------
# Discord Bot Setup
# ---------------------------------------------------------------------------
class Client(commands.Bot):
    """Discord bot for tracking user text and voice activity."""
    
    async def on_ready(self):
        """Triggered on bot initialization"""
        
        print(f'Logged in as {self.user}')
        print('------')
        
        '''
        If you wish to quickly sync updates to your discord, this is a template to do so
        
        try:
            #Replace "Your GUILD_ID Here" with your server's ID
            GUILD_ID = discord.Object(id=Your GUILD_ID Here)
            synced = await self.tree.sync(guild=GUILD_ID)
            print(f'Synced {len(synced)} command(s) to the guild {GUILD_ID}.')
            print('------')
        except Exception as e:
            print(f'Error syncing commands: {e}')
            print('------')
        '''

    async def on_message(self, message):
        """Triggered when a user sends a message."""
        
        # excluding bots from user-activity summary
        if message.author.bot:
            return
        
        
        print(f'Message from {message.author}: {message.content}')
        print('Updating User Activity...')

        updateSpreadsheet(message.author.name, 1)

        print('User Activity Updated.')
        print('------')


    async def on_voice_state_update(self, member, before, after):
        """Triggered when a user joins a voice channel."""
        
        if before.channel is None and after.channel is not None:
            print(f'{member} has joined voice channel: {after.channel}')
            print('Updating User Activity...')

            updateSpreadsheet(member.name, 4)

            print('User Activity Updated.')
            print('------')


# ---------------------------------------------------------------------------
# Discord Slash Commands
# ---------------------------------------------------------------------------
intents = discord.Intents.default()
intents.message_content = True
client = Client(command_prefix='!', intents=intents)

@client.tree.command(name="get-user-activity", description="Check Last User Activity", guild=discord.Object(id=GUILD_ID))
async def activity_check(interaction: discord.Interaction, user: discord.Member):
    print(f'ActivityCheck command invoked by {interaction.user} for {user}.')
    print('Checking User Activity and Displaying...')

    lastMessageDateTime = get_activity(user.name, 1)
    previousMessageDateTime = get_activity(user.name, 2)
    lastVoiceJoinDateTime = get_activity(user.name, 4)
    previousVoiceJoinDateTime = get_activity(user.name, 5)
    totalMessages = get_activity(user.name, 3)
    totalCalls = get_activity(user.name, 6)

    await interaction.response.send_message(f"""User {user.mention} was last active: \n
In Text Chat:     \t         {lastMessageDateTime}
In Voice Chat:    \t        {lastVoiceJoinDateTime}
------------------------------------------------
Before That:
In Text Chat:     \t         {previousMessageDateTime}
In Voice Chat:    \t        {previousVoiceJoinDateTime}
------------------------------------------------
Total Messages Sent:          {totalMessages}
Total Voice Calls Joined:    {totalCalls}""")

    print('User Activity Displayed.')
    print('------')

@client.tree.command(name="get-server-activity", description="Check Last Server Activity", guild=discord.Object(id=GUILD_ID))
async def server_activity_check(interaction: discord.Interaction):
    print(f'ServerActivityCheck command invoked by {interaction.user}.')
    print('Checking Server Activity and Displaying...')

    allUserVector = []
    individualUserVector = []
    mostMessagesVector = []
    mostVoiceVector = []

    for row in worksheet.iter_rows(min_row=2): # Skip header row
        user_id = row[0].value
        total_messages = row[3].value if row[3].value is not None else 0
        total_voice = row[6].value if row[6].value is not None else 0
        individualUserVector = [user_id, total_messages, total_voice]
        allUserVector.append(individualUserVector)

    mostMessagesVector = sorted(allUserVector, key=lambda x: x[1], reverse=True)[:3]
    mostVoiceVector = sorted(allUserVector, key=lambda x: x[2], reverse=True)[:3]

    finalString = "Server Activity Summary:\n\nMost Active in Text Chat:\n"
    finalString += "-------------------------------\n"
    for i, user in enumerate(mostMessagesVector, start=1):
        if user[1] != 0:
            finalString += f"{i}. {user[0]} - {user[1]} Messages Sent\n"

    finalString += "\nMost Active in Voice Chat:\n"
    finalString += "-------------------------------\n"
    for i, user in enumerate(mostVoiceVector, start=1):
        if user[2] != 0:
            finalString += f"{i}. {user[0]} - {user[2]} Calls Joined\n"

    await interaction.response.send_message(finalString)

    print('Server Activity Displayed.')
    print('------')


# ---------------------------------------------------------------------------
# Entry Point
# ---------------------------------------------------------------------------
client.run(DISCORD_TOKEN)