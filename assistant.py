import os
import sys
import time
import datetime
import json
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

# Initialize Windows SAPI5 Speech Engine
try:
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
except Exception as e:
    print(f"Error initializing TTS: {e}")
    speaker = None

# --- Core Functions ---

def speak(text):
    """Speaks the text out loud and allows interruption via any key press."""
    print(f"Jarvis: {text}")
    if speaker:
        # Clean up markdown and special characters
        clean_text = re.sub(r'[*`#_~]', '', text)
        
        # Speak asynchronously (Flag = 1)
        speaker.Speak(clean_text, 1)
        
        print("[Press any key to interrupt...]")
        # 2 = Speaking state
        while speaker.Status.RunningState == 2:
            # Check if a key was pressed to interrupt
            if msvcrt.kbhit():
                msvcrt.getch() # Consume the key press
                # Purge before speak (Flag = 2) with empty string stops current speech
                speaker.Speak("", 2)
                print("[Speech interrupted]")
                break
            time.sleep(0.05)

def listen():
    """Listens to the microphone and returns recognized text."""
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("\nListening...")
        recognizer.adjust_for_ambient_noise(source, duration=0.5)
        try:
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=15)
            print("Recognizing...")
            text = recognizer.recognize_google(audio)
            print(f"You said: {text}")
            return text.lower()
        except sr.WaitTimeoutError:
            return ""
        except sr.UnknownValueError:
            print("Could not understand the audio.")
            return ""
        except sr.RequestError as e:
            print(f"Could not request results from Google Speech Recognition service; {e}")
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
    if "notepad" in app_name:
        subprocess.Popen("notepad.exe")
        return "Opening Notepad."
    elif "calculator" in app_name:
        subprocess.Popen("calc.exe")
        return "Opening Calculator."
    elif "chrome" in app_name:
        subprocess.Popen("start chrome", shell=True)
        return "Opening Google Chrome."
    elif "youtube" in app_name:
        webbrowser.open("https://www.youtube.com")
        return "Opening YouTube."
    else:
        return f"I am not configured to open {app_name} yet."

# --- Main Application Loop ---

def main():
    speak("Jarvis is now online.")
    
    while True:
        command = listen()
        if not command:
            continue
            
        # 1. Exit Commands
        if "exit" in command or "stop" in command or "goodbye" in command:
            speak("Goodbye!")
            break
            
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

main()
