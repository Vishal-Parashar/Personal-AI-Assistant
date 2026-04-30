import pyttsx3
import speech_recognition as sr
import requests
import sys
import os
import platform

# Fix for Windows ARM64 (like Snapdragon PCs) where speech_recognition fails to find the bundled FLAC utility
if platform.system() == "Windows" and platform.machine() == "ARM64":
    import speech_recognition.audio as sr_audio
    original_get_flac = sr_audio.get_flac_converter
    def get_flac_converter_arm64():
        try:
            return original_get_flac()
        except OSError:
            return os.path.join(os.path.dirname(sr_audio.__file__), "flac-win32.exe")
    sr_audio.get_flac_converter = get_flac_converter_arm64

import win32com.client
import re

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def speak(text):
    print(f"AI: {text}")
    # Clean up markdown and special characters so it doesn't say "asterisk asterisk"
    clean_text = re.sub(r'[*`#_~]', '', text)
    speaker.Speak(clean_text)


def listen():
    r = sr.Recognizer()
    # r.energy_threshold = 4000  # Removing this hardcoded high threshold to make it more sensitive
    try:
        with sr.Microphone() as source:
            print("Listening (speak clearly)...")
            r.adjust_for_ambient_noise(source, duration=1)
            # Record the audio without converting format
            audio = r.listen(source, timeout=10, phrase_time_limit=15)
            print("Processing...")
    except sr.MicrophoneError as e:
        print(f"❌ Microphone error: {e}")
        return ""
    except sr.RequestError as e:
        print(f"❌ Request error: {e}")
        return ""
    except sr.WaitTimeoutError:
        print("❌ No speech detected (timeout)")
        return ""
    except Exception as e:
        print(f"❌ Microphone setup error: {e}")
        return ""
    
    try:
        # Use Google Speech Recognition API
        print("Sending to Google API...")
        text = r.recognize_google(audio, language='en-US')
        print(f"✓ You said: {text}")
        return text.lower()
    except sr.UnknownValueError:
        print("❌ Could not understand - speak louder or clearer")
        return ""
    except sr.RequestError as e:
        print(f"❌ Google API error: {e}")
        if "FLAC" in str(e):
            print("   Try installing FLAC: https://xiph.org/flac/download.html")
        return ""
    except Exception as e:
        print(f"❌ Error: {str(e)[:100]}")
        return ""


def get_ai_response(user_text):
    data = {
        "model": "llama3.2:1b",
        "prompt": user_text,
        "stream": False
    }
    response = requests.post("http://localhost:11434/api/generate", json=data)
    return response.json()['response']
# Start the assistant
speak("System online. How can I help you?")
def get_time():
    from datetime import datetime
    now = datetime.now()
    return now.strftime("The current time is %H:%M:%S.")

def get_weather(city):
    api_key = "PASTE_YOUR_API_KEY_HERE"
    # The URL where we send our request
    url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
    
    response = requests.get(url)
    
    if response.status_code == 200:
        data = response.json() # Convert the response to a Python dictionary
        temp = data['main']['temp']
        desc = data['weather'][0]['description']
        return f"The temperature in {city} is {temp} degrees Celsius with {desc}."
    else:
        return "Sorry, I couldn't find that city."

while True:
    command = listen()
    
    if not command:
        continue
    
    if "exit" in command or "stop" in command:
        speak("Shutting down. Goodbye!")
        break
    
    if "weather in" in command:
        # If user says "weather in Delhi", we extract "Delhi"
        city = command.split("in ")[-1]
        report = get_weather(city)
        speak(report)
    
    elif "time" in command:
        time_report = get_time()
        speak(time_report)
    
    else:
        # For everything else, ask your local Ollama AI
        reply = get_ai_response(command)
        speak(reply)