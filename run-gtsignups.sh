#!/bin/bash

# Add NVM support if you're using it
export NVM_DIR="$HOME/.nvm"
[ -s "$NVM_DIR/nvm.sh" ] && \. "$NVM_DIR/nvm.sh"

# Set the directory where your script is located
SCRIPT_DIR="/Users/nathaniel/Documents/Code/javascript/gtsignups"
LOG_DIR="$SCRIPT_DIR/logs"

# Create logs directory if it doesn't exist
mkdir -p "$LOG_DIR"

# Get current date for log file name
DATE=$(date +"%Y-%m-%d_%H-%M-%S")
LOG_FILE="$LOG_DIR/gtsignups_$DATE.log"

# Change to script directory
cd "$SCRIPT_DIR"

# Run the script and log output
echo "Starting Growth Track Signups script at $(date)" >> "$LOG_FILE" 2>&1
node index.js >> "$LOG_FILE" 2>&1
echo "Finished at $(date)" >> "$LOG_FILE" 2>&1

# Keep only last 7 days of logs
find "$LOG_DIR" -name "gtsignups_*.log" -mtime +7 -delete
