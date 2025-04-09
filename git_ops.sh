#!/bin/bash

# Load secrets from secrets.json if it exists
if [ -f "secrets.json" ]; then
    source secrets.json
fi

# Function to check if a command exists
command_exists() {
    command -v "$1" >/dev/null 2>&1
}

# Check for required commands
if ! command_exists git; then
    echo "Git is not installed. Please install Git first."
    exit 1
fi

# Function to get current timestamp
get_timestamp() {
    date +"%Y-%m-%d %H:%M:%S"
}

# Function to commit changes
commit_changes() {
    local message="$1"
    git add .
    git commit -m "$message"
}

# Function to push changes
push_changes() {
    git push origin main
}

# Main script
echo "Starting Git operations at $(get_timestamp)"

# Check if we're in a git repository
if [ ! -d ".git" ]; then
    echo "Initializing new Git repository..."
    git init
    
    # Add remote repository (you'll need to replace this with your actual GitHub repo URL)
    if [ -z "$GITHUB_REPO_URL" ]; then
        read -p "Enter your GitHub repository URL: " GITHUB_REPO_URL
    fi
    git remote add origin "$GITHUB_REPO_URL"
fi

# Check if there are any changes
if git diff --quiet && git diff --cached --quiet; then
    echo "No changes to commit."
else
    # Get commit message
    if [ -z "$1" ]; then
        read -p "Enter commit message: " commit_message
    else
        commit_message="$1"
    fi
    
    # Commit changes
    commit_changes "$commit_message"
    
    # Push changes
    push_changes
fi

echo "Git operations completed at $(get_timestamp)" 