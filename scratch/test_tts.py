import threading
import win32com.client
import pythoncom
import time

def test_speak():
    pythoncom.CoInitialize()
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak("Hello from thread", 1)
    while not speaker.WaitUntilDone(50):
        pass
    print("Done speaking")

t = threading.Thread(target=test_speak)
t.start()
t.join()
