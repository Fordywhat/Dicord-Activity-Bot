# Use official Python image as base
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Copy requirements.txt if exists, else skip
COPY requirements.txt ./

# Install dependencies if requirements.txt exists
RUN if [ -f requirements.txt ]; then pip install --no-cache-dir -r requirements.txt; fi

# Copy all files to the container
COPY . .

# Run the bot
CMD ["python", "main.py"]