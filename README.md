# OpenAI Agents Python

This project contains Python scripts for automating email extraction from Outlook using OpenAI agents.

## Project Structure

- `examples/email_extraction/` - Contains scripts for email extraction
  - `main.py` - Main email extraction script with agent-based architecture
  - `main_small.py` - Simplified version for testing
  - `test_main.py` - Test scripts

## Setup

1. Clone the repository:
```bash
git clone https://github.com/yourusername/openai-agents-python.git
cd openai-agents-python
```

2. Create a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Set up credentials:
```bash
cp secrets.json.template secrets.json
# Edit secrets.json with your credentials
```

## Usage

1. For basic email extraction:
```bash
python examples/email_extraction/main_small.py
```

2. For full agent-based extraction:
```bash
python examples/email_extraction/main.py
```

## Git Operations

Use the provided `git_ops.sh` script for automated Git operations:

```bash
# Make the script executable
chmod +x git_ops.sh

# Basic usage
./git_ops.sh

# With commit message
./git_ops.sh "Your commit message here"
```

## Security

- Never commit `secrets.json` to Git
- Keep your GitHub credentials secure
- Use SSH keys for GitHub authentication
