import os
import sys
import time
import datetime
import json
import base64
import requests
import subprocess
import webbrowser
import msvcrt
import speech_recognition as sr
import win32com.client
import platform
import re

try:
    import config
    OPENWEATHER_API_KEY = config.OPENWEATHER_API_KEY
except ImportError:
    OPENWEATHER_API_KEY = "YOUR_OPENWEATHERMAP_API_KEY"


# Fix for Windows ARM64 where speech_recognition fails to find the bundled FLAC utility
if platform.system() == "Windows" and platform.machine() == "ARM64":
    import speech_recognition.audio as sr_audio
    original_get_flac = sr_audio.get_flac_converter
    def get_flac_converter_arm64():
        try:
            return original_get_flac()
        except OSError:
            base_path = os.path.dirname(sr.__file__)
            fallback_path = os.path.join(base_path, "flac-win32.exe")
            if os.path.exists(fallback_path):
                return fallback_path
            raise
    sr_audio.get_flac_converter = get_flac_converter_arm64

# --- Setup & Configuration ---

OLLAMA_URL = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "llama3.2:1b"
OLLAMA_VISION_MODEL = "llava"  # Requires: ollama pull llava

# Initialize Windows SAPI5 Speech Engine
try:
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
except Exception as e:
    print(f"Error initializing TTS: {e}")
    speaker = None

# --- Core Functions ---

def speak(text, interrupt_event=None):
    """Speaks the text out loud and allows interruption via any key press or threading event."""
    print(f"Jarvis: {text}")
    
    # Initialize COM for background threads
    try:
        import pythoncom
        pythoncom.CoInitialize()
        thread_speaker = win32com.client.Dispatch("SAPI.SpVoice")
    except Exception as e:
        print(f"Error initializing thread TTS: {e}")
        thread_speaker = speaker # Fallback to global if pythoncom not available
        
    if thread_speaker:
        # Clean up markdown and special characters
        clean_text = re.sub(r'[*`#_~]', '', text)
        
        # Clear any stale keypresses in the buffer if no event is provided
        if not interrupt_event:
            while msvcrt.kbhit():
                msvcrt.getch()
            
        # Speak asynchronously (Flag = 1)
        thread_speaker.Speak(clean_text, 1)
        
        if not interrupt_event:
            print("[Press any key to interrupt...]")
            
        # Loop until it's done speaking
        while not thread_speaker.WaitUntilDone(50):
            # Check if interrupted via GUI event
            if interrupt_event and interrupt_event.is_set():
                thread_speaker.Speak("", 3)
                print("[Speech interrupted]")
                break
            # Check if a key was pressed to interrupt (Console mode)
            elif not interrupt_event and msvcrt.kbhit():
                # Consume all pressed keys
                while msvcrt.kbhit():
                    msvcrt.getch()
                
                # Purge before speak (Flag = 2) + Async (Flag = 1) = 3
                thread_speaker.Speak("", 3)
                print("[Speech interrupted]")
                break

def listen(status_callback=None):
    """Listens to the microphone and returns recognized text."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass
        
    try:
        recognizer = sr.Recognizer()
        with sr.Microphone() as source:
            if status_callback:
                status_callback("Listening...")
            else:
                print("\nListening...")
            recognizer.adjust_for_ambient_noise(source, duration=0.5)
            
            try:
                audio = recognizer.listen(source, timeout=5, phrase_time_limit=15)
                if status_callback:
                    status_callback("Recognizing...")
                else:
                    print("Recognizing...")
                text = recognizer.recognize_google(audio)
                print(f"You said: {text}")
                return text.lower()
            except sr.WaitTimeoutError:
                if status_callback: status_callback("Listening timed out.")
                return ""
            except sr.UnknownValueError:
                msg = "Could not understand the audio."
                if status_callback:
                    status_callback(msg)
                else:
                    print(msg)
                return ""
            except sr.RequestError as e:
                msg = f"API Error: {e}"
                if status_callback:
                    status_callback(msg)
                else:
                    print(msg)
                return ""
    except Exception as e:
        msg = f"Microphone error: {e}"
        print(msg)
        if status_callback:
            status_callback(msg)
        return ""

# --- Feature Functions ---

def get_weather(city):
    """Fetches weather data for a given city using OpenWeatherMap."""
    if OPENWEATHER_API_KEY == "YOUR_OPENWEATHERMAP_API_KEY" or OPENWEATHER_API_KEY == "XXXXXXXXXXXXXXXXXXXXX":
        return "Please set your OpenWeatherMap API key in the script to use weather features."
        
    url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={OPENWEATHER_API_KEY}&units=metric"
    try:
        response = requests.get(url)
        data = response.json()
        if data.get("cod") != "404" and "main" in data:
            temp = data["main"]["temp"]
            desc = data["weather"][0]["description"]
            return f"The current temperature in {city} is {temp} degrees Celsius with {desc}."
        else:
            return f"I couldn't find the weather for {city}."
    except Exception:
        return "I encountered an error while trying to fetch the weather data."

def get_ai_response(prompt):
    """Sends prompt to local Ollama server and returns the response."""
    data = {
        "model": OLLAMA_MODEL,
        "prompt": prompt,
        "stream": False
    }
    try:
        response = requests.post(OLLAMA_URL, json=data)
        if response.status_code == 200:
            return response.json().get("response", "I'm not sure what to say.")
        else:
            return "My AI brain returned an error."
    except requests.exceptions.ConnectionError:
        return "I cannot connect to my local Ollama server. Please ensure it is running."

def open_app(app_name):
    """Attempts to open common applications based on keywords."""
    app_name = app_name.lower().strip()

    # --- Browsers ---
    if "chrome" in app_name:
        subprocess.Popen("start chrome", shell=True)
        return "Opening Google Chrome."
    elif "firefox" in app_name:
        subprocess.Popen("start firefox", shell=True)
        return "Opening Firefox."
    elif "edge" in app_name:
        subprocess.Popen("start msedge", shell=True)
        return "Opening Microsoft Edge."
    elif "brave" in app_name:
        subprocess.Popen(r"start brave", shell=True)
        return "Opening Brave Browser."
    elif "opera" in app_name:
        subprocess.Popen("start opera", shell=True)
        return "Opening Opera."

    # --- Productivity & Office ---
    elif "notepad++" in app_name or "notepad plus" in app_name:
        subprocess.Popen("notepad++.exe", shell=True)
        return "Opening Notepad++."
    elif "notepad" in app_name:
        subprocess.Popen("notepad.exe")
        return "Opening Notepad."
    elif "word" in app_name:
        subprocess.Popen("start winword", shell=True)
        return "Opening Microsoft Word."
    elif "excel" in app_name:
        subprocess.Popen("start excel", shell=True)
        return "Opening Microsoft Excel."
    elif "powerpoint" in app_name:
        subprocess.Popen("start powerpnt", shell=True)
        return "Opening Microsoft PowerPoint."
    elif "outlook" in app_name:
        subprocess.Popen("start outlook", shell=True)
        return "Opening Microsoft Outlook."
    elif "onenote" in app_name:
        subprocess.Popen("start onenote", shell=True)
        return "Opening Microsoft OneNote."

    # --- System Tools ---
    elif "calculator" in app_name or "calc" in app_name:
        subprocess.Popen("calc.exe")
        return "Opening Calculator."
    elif "task manager" in app_name:
        subprocess.Popen("taskmgr.exe")
        return "Opening Task Manager."
    elif "file explorer" in app_name or "explorer" in app_name:
        subprocess.Popen("explorer.exe")
        return "Opening File Explorer."
    elif "control panel" in app_name:
        subprocess.Popen("control.exe")
        return "Opening Control Panel."
    elif "settings" in app_name:
        subprocess.Popen("start ms-settings:", shell=True)
        return "Opening Windows Settings."
    elif "registry" in app_name:
        subprocess.Popen("regedit.exe")
        return "Opening Registry Editor."
    elif "command prompt" in app_name or "cmd" in app_name:
        subprocess.Popen("cmd.exe")
        return "Opening Command Prompt."
    elif "powershell" in app_name:
        subprocess.Popen("powershell.exe")
        return "Opening PowerShell."
    elif "paint" in app_name:
        subprocess.Popen("mspaint.exe")
        return "Opening Paint."
    elif "snipping tool" in app_name or "snip" in app_name:
        subprocess.Popen("snippingtool.exe", shell=True)
        return "Opening Snipping Tool."
    elif "device manager" in app_name:
        subprocess.Popen("devmgmt.msc", shell=True)
        return "Opening Device Manager."
    elif "disk management" in app_name:
        subprocess.Popen("diskmgmt.msc", shell=True)
        return "Opening Disk Management."
    elif "event viewer" in app_name:
        subprocess.Popen("eventvwr.msc", shell=True)
        return "Opening Event Viewer."

    # --- Development Tools ---
    elif "vs code" in app_name or "vscode" in app_name or "visual studio code" in app_name:
        subprocess.Popen("code", shell=True)
        return "Opening Visual Studio Code."
    elif "visual studio" in app_name:
        subprocess.Popen("start devenv", shell=True)
        return "Opening Visual Studio."
    elif "git bash" in app_name:
        subprocess.Popen("git-bash.exe", shell=True)
        return "Opening Git Bash."
    elif "github desktop" in app_name:
        subprocess.Popen("start github", shell=True)
        return "Opening GitHub Desktop."
    elif "postman" in app_name:
        subprocess.Popen("start Postman", shell=True)
        return "Opening Postman."
    elif "android studio" in app_name:
        subprocess.Popen("start studio64", shell=True)
        return "Opening Android Studio."

    # --- Media & Entertainment ---
    elif "vlc" in app_name or "media player" in app_name:
        subprocess.Popen("vlc.exe", shell=True)
        return "Opening VLC Media Player."
    elif "spotify" in app_name:
        subprocess.Popen("start spotify", shell=True)
        return "Opening Spotify."
    elif "photos" in app_name:
        subprocess.Popen("start ms-photos:", shell=True)
        return "Opening Photos."
    elif "movies" in app_name or "windows media" in app_name:
        subprocess.Popen("start mswindowsvideo:", shell=True)
        return "Opening Movies & TV."
    elif "itunes" in app_name:
        subprocess.Popen("start itunes", shell=True)
        return "Opening iTunes."
    elif "audacity" in app_name:
        subprocess.Popen("audacity.exe", shell=True)
        return "Opening Audacity."

    # --- Communication ---
    elif "discord" in app_name:
        subprocess.Popen("start discord", shell=True)
        return "Opening Discord."
    elif "teams" in app_name:
        subprocess.Popen("start teams", shell=True)
        return "Opening Microsoft Teams."
    elif "slack" in app_name:
        subprocess.Popen("start slack", shell=True)
        return "Opening Slack."
    elif "zoom" in app_name:
        subprocess.Popen("start zoom", shell=True)
        return "Opening Zoom."
    elif "telegram" in app_name:
        subprocess.Popen("start telegram", shell=True)
        return "Opening Telegram."
    elif "whatsapp" in app_name:
        subprocess.Popen("start whatsapp", shell=True)
        return "Opening WhatsApp."
    elif "skype" in app_name:
        subprocess.Popen("start skype", shell=True)
        return "Opening Skype."

    # --- Utilities ---
    elif "7-zip" in app_name or "7zip" in app_name:
        subprocess.Popen(r"C:\Program Files\7-Zip\7zFM.exe", shell=True)
        return "Opening 7-Zip."
    elif "winrar" in app_name:
        subprocess.Popen("winrar.exe", shell=True)
        return "Opening WinRAR."

    # --- Websites ---
    elif "youtube" in app_name:
        webbrowser.open("https://www.youtube.com")
        return "Opening YouTube."
    elif "google" in app_name:
        webbrowser.open("https://www.google.com")
        return "Opening Google."
    elif "gmail" in app_name:
        webbrowser.open("https://mail.google.com")
        return "Opening Gmail."
    elif "github" in app_name:
        webbrowser.open("https://www.github.com")
        return "Opening GitHub."
    elif "netflix" in app_name:
        webbrowser.open("https://www.netflix.com")
        return "Opening Netflix."
    elif "reddit" in app_name:
        webbrowser.open("https://www.reddit.com")
        return "Opening Reddit."
    elif "twitter" in app_name or "x.com" in app_name:
        webbrowser.open("https://www.x.com")
        return "Opening X (Twitter)."
    elif "linkedin" in app_name:
        webbrowser.open("https://www.linkedin.com")
        return "Opening LinkedIn."
    elif "amazon" in app_name:
        webbrowser.open("https://www.amazon.in")
        return "Opening Amazon."
    elif "chat gpt" in app_name or "chatgpt" in app_name:
        webbrowser.open("https://chat.openai.com")
        return "Opening ChatGPT."
    elif "maps" in app_name:
        webbrowser.open("https://maps.google.com")
        return "Opening Google Maps."
    elif "translate" in app_name:
        webbrowser.open("https://translate.google.com")
        return "Opening Google Translate."

    else:
        return f"I am not configured to open '{app_name}' yet."

def summarize_pdf(filepath):
    """Extracts text from a PDF and asks the AI to summarize it."""
    try:
        import pypdf
        with open(filepath, 'rb') as file:
            reader = pypdf.PdfReader(file)
            text = ""
            for page in reader.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
                
            if not text.strip():
                return "I could not extract any text from this PDF. It might be scanned or empty."
                
            # Truncate to roughly 5000 words (approx 25000 characters) to prevent overloading local model
            max_chars = 25000
            if len(text) > max_chars:
                text = text[:max_chars] + "\n... [Text truncated due to length]"
                
            prompt = f"Please provide a comprehensive summary of the following document:\n\n{text}"
            return get_ai_response(prompt)
    except Exception as e:
        return f"An error occurred while reading the PDF: {e}"

def analyze_image(filepath, question=None):
    """Sends an image to the local Ollama LLaVA vision model for analysis."""
    supported_formats = (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp")
    ext = os.path.splitext(filepath)[1].lower()
    if ext not in supported_formats:
        return f"Unsupported image format '{ext}'. Please use: JPG, PNG, BMP, GIF, or WebP."

    try:
        with open(filepath, "rb") as img_file:
            image_b64 = base64.b64encode(img_file.read()).decode("utf-8")
    except FileNotFoundError:
        return "I could not find the image file. Please check the path and try again."
    except Exception as e:
        return f"Error reading image file: {e}"

    prompt = question if question else "Please describe this image in detail. Mention objects, colors, text, people, and any notable elements you see."

    payload = {
        "model": OLLAMA_VISION_MODEL,
        "prompt": prompt,
        "images": [image_b64],
        "stream": False
    }

    try:
        response = requests.post(OLLAMA_URL, json=payload, timeout=120)
        if response.status_code == 200:
            return response.json().get("response", "I could not generate a description for this image.")
        elif response.status_code == 404:
            return (f"The vision model '{OLLAMA_VISION_MODEL}' is not installed. "
                    f"Please run: ollama pull {OLLAMA_VISION_MODEL}")
        else:
            return f"Vision model returned an error (status {response.status_code})."
    except requests.exceptions.ConnectionError:
        return "I cannot connect to my local Ollama server. Please ensure it is running."
    except requests.exceptions.Timeout:
        return "The image analysis timed out. Try a smaller image or simpler question."

def process_command(command):
    """Processes a command string and returns a response string."""
    command = command.lower().strip()
    if not command:
        return ""
        
    if "exit" in command or "stop" in command or "goodbye" in command:
        return "Have a great day! Goodbye!"
    elif "switch to text" in command or "switch to voice" in command:
        return "I am now controlled via the GUI."
    elif "time" in command:
        now = datetime.datetime.now().strftime("%I:%M %p")
        return f"The time is {now}."
    elif "weather in" in command:
        city = command.split("weather in")[-1].strip()
        if city:
            return get_weather(city)
        return "Please specify a city for the weather."
    elif "open" in command:
        app_name = command.split("open", 1)[1].strip()
        if app_name:
            return open_app(app_name)
        return "What application would you like me to open?"
    elif any(kw in command for kw in ["analyze image", "describe image", "what's in the image", "look at image"]):
        return "Please use the image button (🖼️) in the GUI to select an image for me to analyze."
    else:
        return get_ai_response(command)

# --- Main Application Loop ---

def main():
    speak("Jarvis is now online.")
    
    print("\nHow would you like to interact with Jarvis?")
    print("1. Text Command")
    print("2. Voice Command")
    while True:
        mode = input("Enter 1 or 2: ").strip()
        if mode in ['1', '2']:
            break
        print("Invalid choice. Please enter 1 or 2.")
        
    use_voice = (mode == '2')
    
    while True:
        if use_voice:
            command = listen()
        else:
            command = input("\nYou: ").strip().lower()
            
        if not command:
            continue
            
        # 1. Exit Commands
        if "exit" in command or "stop" in command or "goodbye" in command:
            speak("have a great day! Goodbye!")
            break
            
        # 2. Mode Switching Commands
        elif "switch to text" in command:
            if not use_voice:
                speak("I am already in text mode.")
            else:
                use_voice = False
                speak("Switching to text mode.")
            continue
            
        elif "switch to voice" in command:
            if use_voice:
                speak("I am already in voice mode.")
            else:
                use_voice = True
                speak("Switching to voice mode. I am listening.")
            continue
            
        # 2. Time Command
        elif "time" in command:
            now = datetime.datetime.now().strftime("%I:%M %p")
            speak(f"The time is {now}.")
            
        # 3. Weather Command
        elif "weather in" in command:
            # Extract the city name from the command
            city = command.split("weather in")[-1].strip()
            if city:
                weather_info = get_weather(city)
                speak(weather_info)
            else:
                speak("Please specify a city for the weather.")
                
        # 4. Open App Command
        elif "open" in command:
            # Extract the app name from the command
            app_name = command.split("open", 1)[1].strip()
            if app_name:
                response = open_app(app_name)
                speak(response)
            else:
                speak("What application would you like me to open?")
                
        # 5. Default: Ask Local AI
        else:
            # We add a slight pause/acknowledgment so the user knows it's thinking
            print("Thinking...") 
            response = get_ai_response(command)
            speak(response)

if __name__ == "__main__":
    main()
