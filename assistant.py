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
