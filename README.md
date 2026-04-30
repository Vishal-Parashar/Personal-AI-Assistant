# Jarvis AI Assistant

A Python-based voice assistant that listens to your commands, talks back, checks the weather, opens applications, and uses a local Ollama AI model for natural conversation.

## Features

- **Voice Recognition**: Uses Google Speech Recognition to convert your speech to text.
- **Speech Synthesis**: Uses Windows built-in SAPI5 (via `win32com`) for highly reliable and responsive text-to-speech.
- **Interruption Support**: Press any key while the AI is speaking to immediately stop its response and make it listen to you again.
- **Local AI Brain**: Uses [Ollama](https://ollama.com/) with the `llama3.2:1b` model to answer general knowledge questions locally and privately.
- **Weather Reports**: Gets current weather data for any city using OpenWeatherMap.
- **Time Check**: Tells you the current time.
- **App Launcher**: Opens common apps and websites like Notepad, Calculator, Chrome, and YouTube.
- **ARM64 Support**: Automatically fixes speech recognition compatibility issues on modern Windows ARM64 devices (like Copilot+ PCs).

## Prerequisites

1. **Python 3.x** installed.
2. **Ollama** installed and running on your machine.
3. The `llama3.2:1b` model downloaded in Ollama:
   ```bash
   ollama run llama3.2:1b
   ```

## Installation

Install the required Python libraries using pip:

```bash
pip install SpeechRecognition requests pywin32 pyttsx3
```

*(Note: `pyttsx3` is kept as a dependency for fallback, though the script primarily uses `pywin32`'s `win32com.client` for stable speech synthesis).*

## Usage

1. Start your local Ollama server if it isn't running already.
2. Run the assistant script:
   ```bash
   python assistant.py
   ```
3. Wait for the assistant to say "Jarvis is now online."
4. Speak your commands!

### Example Commands

- *"What is the time?"*
- *"Weather in London"*
- *"Open Notepad"*
- *"Open YouTube"*
- *"What is the capital of France?"* (Handled by Ollama)
- *"Exit"* or *"Stop"* to close the assistant.

### Interrupting the Assistant

If you ask the AI a question and it starts giving a very long answer, simply **press any key** (like Spacebar or Enter) on your keyboard inside the terminal window. The AI will immediately stop talking and start listening to your next command.

## Configuration

- **Weather API**: The script currently includes a placeholder/basic OpenWeatherMap API key. If you experience rate-limits, sign up at [OpenWeatherMap](https://openweathermap.org/) to get your own API key and replace it in the `get_weather` function.
- **AI Model**: If you want to use a different Ollama model (like `llama3:8b` or `mistral`), you can change the `"model": "llama3.2:1b"` string inside the `get_ai_response` function.

## License

This is a personal project. Feel free to modify and adapt it to your needs!
