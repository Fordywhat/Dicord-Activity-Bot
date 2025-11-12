# Discord Activity Tracker Bot

A Docker ready open-source Discord bot that tracks user activity in text and voice channels, storing the data in an Excel spreadsheet (`user_activity.xlsx`).

## Features
- Tracks message and voice activity
- Displays user-specific and server-wide summaries
- Uses Slash Commands (`/get-user-activity` and `/get-server-activity`)
- Built with `discord.py` and `openpyxl`

## Setup (Local)

### 1. Clone the Repository
```bash
git clone https://github.com/Fordywhat/discord-activity-tracker.git
cd discord-activity-tracker
```

### 2. Edit the .env File
```
DISCORD_TOKEN=your_discord_token_here
GUILD_ID=your_guild_id_here
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Run the Bot
```bash
python main.py
```

## Setup (Docker)

### 1. Clone the Repository
```bash
git clone https://github.com/Fordywhat/discord-activity-tracker.git
cd discord-activity-tracker
```

### 2. Edit the .env File
```
DISCORD_TOKEN=your_discord_token_here
GUILD_ID=your_guild_id_here
```

### 3. Build the Docker Image
```bash
docker build -t discord-activity-tracker .
```

### 4. Run the Container
```bash
docker run --env-file .env discord-activity-tracker
```
