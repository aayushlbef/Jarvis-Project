import speech_recognition as sr
import webbrowser
import pyttsx3
import datetime
import musicLibrary
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL


engine = pyttsx3.init()

def speak(text):
    engine.say(text)
    engine.runAndWait()

def wakeUpCommand():
    while True:
        r = sr.Recognizer()
        try:
            with sr.Microphone() as source:
                print("Listening...")
                r.pause_threshold = 1
                audio = r.listen(source)

            print("Recognizing...")
            wake_up = r.recognize_google(audio)
            print(wake_up)

            if "jarvis" in wake_up.lower():
                if wake_up.lower() == "jarvis":
                    time_greeting(wake_up)
                else:
                    parts = wake_up.lower().split("jarvis", 1)
                    command = parts[1].strip()
                    if "open" in command.lower():
                        executeCommandOpen(command)
                    elif "play" in command.lower():
                        executeCommandPlay(command)
                    elif "exit" in command.lower():
                        executeCommandExit()
        except Exception as e:
            print(e)

def time_greeting(wake_up):
    hour = datetime.datetime.now().hour
    if hour >= 0 and hour < 12:
        speak("Good Morning Aayush. How may I help you today")
    elif hour >= 12 and hour< 18:
        speak("Good Afternoon Aayush. How may I help you today")
    else:
        speak("Good Evining Aayush. How may I help you today")
    listeningCommand()

def listeningCommand():
    r = sr.Recognizer()
    while True:
        with sr.Microphone() as source:
            print("Waiting for Command...")
            r.pause_threshold = 1
            audio = r.listen(source, timeout= 10)
            command = r.recognize_google(audio)
            print(command)
            if "open" in command.lower():
                executeCommandOpen(command)
            elif "play" in command.lower():
                executeCommandPlay(command)
            elif "exit" in command.lower():
                executeCommandExit()


def executeCommandOpen(command):
    try:
        if "open google" in command.lower():
            webbrowser.open("www.google.com")
        elif "open youtube" in command.lower():
            webbrowser.open("www.youtube.com")
        elif "open facebook" in command.lower():
            webbrowser.open("www.facebook.com")
        elif "open chat gpt" in command.lower():
            webbrowser.open("https://chatgpt.com/")
        elif "open gemini" in command.lower():
            webbrowser.open("https://gemini.google.com/app")
        elif "open monkey type" in command.lower():
            webbrowser.open("https://monkeytype.com/")
        elif "volume" in command.lower() or "mute" in command.lower():
            change_volume(command)
    except Exception as e:
        print(e)

def executeCommandPlay(command):
    try:
        if command.startswith("play"):
            parts = command.lower().split(" ")[1]
            link = musicLibrary.music[parts]
            webbrowser.open(link)
    except Exception as e:
        print(e)

def change_volume(command):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    current_volume = volume.GetMasterVolumeLevelScalar()
    if "increase volume" in command.lower():
        new_volume = min(current_volume + 0.1, 1.0)
        volume.SetMasterVolumeLevelScalar(new_volume, None)
        speak("Increasing volume.")
    elif "decrease volume" in command.lower():
        new_volume = max(current_volume - 0.1, 0.0)
        volume.SetMasterVolumeLevelScalar(new_volume, None)
        speak("Decreasing volume.")
    elif "mute" in command.lower():
        volume.SetMute(1, None)
        speak("Volume muted.")
    elif "unmute" in command.lower():
        volume.SetMute(0, None)
        speak("Volume unmuted.")

def executeCommandExit():
    speak("Powering Off")
    exit()
        
if __name__ == "__main__":
    speak("Initiallizing Jarvis")
    wakeUpCommand()