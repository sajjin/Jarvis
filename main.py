from __future__ import print_function

import datetime
import json
import os
import psutil
import subprocess
import threading
import time
import webbrowser
import pyautogui

import requests
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import speech_recognition as sr  # pip install speechRecognition
import keyboard
import win32con
import win32gui
from gtts import gTTS
from playsound import playsound
from wikipedia import wikipedia
import googlesearch as search
from youtubesearchpython import SearchVideos
from tkinter import *

links = []
os.chdir("C:\\Users\\sajji\\Desktop\\personal_projects\\Jarvis\\audio")

# engine = pyttsx3.init('sapi5')
# voices = engine.getProperty('voices')
# print(voices)
# for male
# engine.setProperty('voice', voices[0].id)


# for female
# engine.setProperty('voice', voices[1].id)


# Beta GUI system currently not working

# root = Tk()
#
# e = Entry(root, width=50)
# e.pack()
#
#
# def myClick(audio):
#     myLabel = Label(root, text=f"Friday: {audio}\n")
#     myLabel.pack()
#     remove_text()
#
#
# def remove_text():
#     e.delete(0, 'end')
#
#
# myButton = Button(root, text="Enter", command=myClick)
# myButton.pack()


def speak(audio):
    """
    saves the audio to a file and plays it then deletes the file
    :param audio:
    :return:
    """
    print(f"Jarvis/Friday: {audio}\n")
    tts = gTTS(text=audio, lang='en-uk')
    tts.save("speech.mp3")
    playsound("speech.mp3")
    os.remove("speech.mp3")
    # engine.say(audio)
    # engine.runAndWait()


def notification():
    playsound("Cortana Sound Effect.mp3")


def ytlink(query):
    """
    youtube search function using youtube-search
    :param query:
    :return webbrowser to searched video link
    """
    ytsearch = SearchVideos(query, offset=1,
                            mode="json", max_results=1)
    vidli = ytsearch.result()
    # parse x:
    ydata = json.loads(vidli)
    for ytv in ydata['search_result']:
        ytmli = ytv['link']

    webbrowser.open(ytmli)


# def enum_callback(hwnd, results):
#     winlist.append((hwnd, win32gui.GetWindowText(hwnd)))

def wedtn():
    """
    Get local weather data from openweathermap.org
    :return: het temp in celcius and humidity, and type of weather
    """
    BASE_URL = "https://api.openweathermap.org/data/2.5/weather?"
    CITY = "Vancouver"
    API_KEY = ""  # add api key on api.openweathermap.org
    # upadting the URL
    URL = BASE_URL + "q=" + CITY + "&appid=" + API_KEY
    # HTTP request
    response = requests.get(URL).json()
    # checking the status code of the request
    if response["cod"] != "404":
        weather = response["main"]
        temp = weather["temp"]
        temperature = int(temp - 273.15)
        hum = weather["humidity"]
        desc = response["weather"][0]["description"]
        resp_string = "The temperature in Celsius is " + str(temperature) + " With a humidity of " + str(
            hum) + " with " + str(desc)
        speak(resp_string)
        if temperature in range(-15, 0):
            speak("that's way to cold")
        if temperature in range(20, 29):
            speak("I bet you've got that ball sweat going on")
        if temperature in range(30, 39):
            speak("This is unheard of, it's almost like climate change is real")
        if temperature in range(40, 70):
            speak("KILL ME, I'M MELTING")
    else:
        # showing the error message
        speak("sorry sir. no data today")


def wishMe():
    """
    Greeting Jarvis
    :return:
    """
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        speak("Good Morning!")

    elif 12 <= hour < 18:
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")


def takeCommand():
    """
    It takes microphone input from the user and returns string output
    :return string:
    """

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        # r.adjust_for_ambient_noise(source)
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-CA')
        print(f"You said: {query}\n")

    except Exception:
        # print(e)
        print("Say that again please...")
        return "None"
    return query


def voice_controller():
    """
    looks at the input from the user and decides what to do based on the input changes if and for another command
    :return: command(s)
    """
    global v
    v = volume_down()
    query = takeCommand().lower().rstrip(" ").lstrip(" ")

    while query == "none":
        query = takeCommand().lower().rstrip(" ").lstrip(" ")

    # Logic for executing tasks based on query

    if 'and' in query:
        query2 = query.replace("and", "  ")
        query2 = ''.join(query2)
        query2 = query2.replace("also", "  ")
        query2 = ''.join(query2)
        query2 = query2.split("  ")
        ln = len(query2)
        if ln >= 2:
            query = query2[0].rstrip(" ").lstrip(" ")
            options(query)
            query3 = query2[2].rstrip(" ").lstrip(" ")
            options(query3)
        if ln == 5:
            query4 = query2[4].rstrip(" ").lstrip(" ")
            options(query4)
    else:
        options(query)

    volume_up(v)


def options(query):
    """
    decides what to do based on the input
    :param query:
    :return:
    """
    if 'wikipedia' in query:
        speak('Searching Wikipedia...')
        query = query.replace("wikipedia", "")
        results = wikipedia.summary(query, sentences=2)
        speak(results)

    elif 'open github' in query:
        webbrowser.open("https://github.com/")

    elif 'weather' in query or 'temperature' in query or 'today\'s report' in query:
        wedtn()

    elif 'the time' in query or "what's the time" in query:
        strTime = datetime.datetime.now().strftime("%I:%M")
        speak(f"the time is {strTime}")

    elif 'open code' in query:
        speak("opening your code.")
        codePath = "C:\\Users\\sajji\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"  # add your vs code path
        os.startfile(codePath)

    elif 'hide window' in query or 'hide work' in query or 'change window' in query or 'minimise window' in query:
        # close in window
        speak("ok.")
        Minimize = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(Minimize, win32con.SW_MINIMIZE)

    elif 'exit' in query or 'goodbye' in query or 'good bye' in query or 'bye' in query or "stop listening" in query:
        speak("thank you. good bye.")
        volume_up(v)
        os.system("taskkill /im cmd.exe /f")
        os.system("taskkill /im python3.9.exe /f")

    elif 'thank you' in query or 'thanks' in query:
        speak("No problem.")

    elif "hello" in query or "hello Friday" in query:
        hel = "Hello! How May i Help you.."
        speak(hel)

    elif "close chrome" in query or "close tabs" in query:
        volume_up(v)
        os.system("taskkill /im chrome.exe /f")
        speak("I am closing all tabs")

    elif "open chrome" in query or "open a new chrome window" in query:
        os.system("start chrome.exe")
        speak("Sure")

    elif "open youtube" in query:
        os.system("start msedge")
        speak("Sure")

    elif "turn off" in query or "shut down" in query:
        os.system("shutdown /s")
        speak("Shutting everything down, see you later")
        volume_up(v)
        os.system("taskkill /im cmd.exe /f")
        os.system("taskkill /im python3.9.exe /f")

    elif "restart" in query:
        os.system("shutdown /r ")
        speak("see you in a minute")
        volume_up(v)
        os.system("taskkill /im cmd.exe /f")
        os.system("taskkill /im python3.9.exe /f")

    elif 'open steam' in query:
        # added sublime exe location
        speak("opening steam ")
        codePathf = "C:\\Program Files (x86)\\Steam\\Steam.exe"
        os.startfile(codePathf)

    elif "open" in query or 'link' in query:
        keyboard.press_and_release("ctrl+w")
        if 'first' in query:
            webbrowser.open(links[0])
        elif 'second' in query:
            webbrowser.open(links[1])
        elif 'third' in query:
            webbrowser.open(links[2])
        elif 'fourth' in query:
            webbrowser.open(links[3])
        elif 'fifth' in query:
            webbrowser.open(links[4])
        elif 'sixth' in query:
            webbrowser.open(links[5])
        elif 'seventh' in query:
            webbrowser.open(links[6])
        elif 'eighth' in query:
            webbrowser.open(links[7])
        elif 'ninth' in query:
            webbrowser.open(links[8])
        elif 'tenth' in query:
            webbrowser.open(links[9])
        elif 'eleventh' in query:
            webbrowser.open(links[10])
        elif 'twelfth' in query:
            webbrowser.open(links[11])

    elif 'go to' in query or 'look up' in query or 'go-to' in query:
        query = query.split("-")
        word = " ".join(query)
        query = word.split(" ")
        word1 = " ".join(query[2:])
        query = word1.split(".")
        word = ".".join(query)
        links.clear()
        for j in search(word1):
            links.append(j)
        for i in links:
            if not i.startswith("https"):
                links.remove(i)
            elif not i.startswith('http'):
                links.remove(i)
        if "com" in query:
            webbrowser.open("https://" + word)
        else:
            webbrowser.open("https://google.com/search?q=%s" % word1)

    elif "take screenshot" in query or "take a screenshot" in query or "capture the screen" in query:
        speak("What do you want to name it?")
        name = takeCommand().lower()
        speak("Alright, taking the screenshot")
        img = pyautogui.screenshot()
        name = f"{name}.png"
        img.save("C:\\Users\\sajji\\OneDrive\\Pictures\\" + name)
        speak("The screenshot has been successfully captured")

    elif 'help' in query:
        speak("to look up something say, go to")
        speak("or say open edge")
        speak("or say play, video title to open a youtube video")
        speak("or say exit, to stop me from listening")

    elif "testing" in query or "test" in query:
        speak("I seem to be working")

    elif "youtube" in query or "play" in query:
        word = query.split(" ")

        if 'play' in word[0]:
            word.remove("play")
        if 'youtube' in word[-1]:
            word.pop(len(word) - 1)
        if 'on' in word[-1]:
            word.pop(len(word) - 1)
        query = " ".join(word)

        speak("playing " + query)
        ytlink(query)

    elif 'programs running' in query or 'running programs' in query or "programs" in query:
        cmd = 'powershell "gps | where {$_.MainWindowTitle } | select Name,Description,Id,Path'
        proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
        for line in proc.stdout:
            if not line.decode()[0].isspace():
                print(line.decode().rstrip())

    elif 'who' in query or 'what' in query or 'how' in query or 'when' in query or 'why' in query:
        webbrowser.open("https://google.com/search?q=%s" % query)

    elif 'full window' in query or 'full screen window' in query or 'full screen' in query or 'maximize window' in query or 'full-screen' in query:
        # full in window
        speak("sure.")
        time.sleep(5)
        if 'game' in query:
            keyboard.press_and_release('alt+tab')
        a = pyautogui.position()
        pyautogui.click(-255, 360)
        pyautogui.moveTo(a)
        keyboard.press_and_release('f')

    elif 'clear' in query:
        os.system('cls')

    elif 'kill' in query or 'stop' in query or 'close' in query:
        if 'kill' in query:
            query = query.split(" ")
            query.remove('kill')
            query = ' '.join(query)
        if 'stop' in query:
            query = query.split(" ")
            query.remove('stop')
            query = ' '.join(query)
        if 'close' in query:
            query = query.split(" ")
            query.remove('close')
            query = ' '.join(query)
        for proc in psutil.process_iter():
            # check whether the process name matches
            # print(proc.name())
            try:
                if any(procstr in proc.name().lower() for procstr in [query]):
                    print(f'Killing {proc.name()}')
                    volume_up(v)
                    proc.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        speak("Done")


def volume_up(v):
    """
    volume up after done with task
    :param v:
    :return:
    """
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if v is None:
            pass
        elif session.Process and session.Process.name() == "chrome.exe":
            # print("volume.GetMasterVolume(): %s" % volume.GetMasterVolume())
            volume.SetMasterVolume(v, None)


def volume_down():
    """
    volume down for everything when the user is speaking
    :return:
    """
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.Process and session.Process.name() == "chrome.exe":
            # print("volume.GetMasterVolume(): %s" % volume.GetMasterVolume())
            v = volume.GetMasterVolume()
            volume.SetMasterVolume(0.05, None)
            return v


def rec():
    """
    first recording function to record the audio it is used to find friday or jarvis if spoken
    :return:
    """
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # print("Listening for Friday...")
        # r.adjust_for_ambient_noise(source)
        audio = r.listen(source)

    try:
        # print("Recognizing...")
        query = r.recognize_google(audio, language='en-CA')

    except Exception:
        # print("say again")
        return "None"
    print("Me: " + query)
    return query


if __name__ == "__main__":
    wishMe()
    print("Listening for Jarvis/Friday...")
    while True:
        a = rec().lower()
        if "friday" in a or 'jarvis' in a:
            threading.Thread(target=notification).start()
            voice_controller()
            print("Listening for Jarvis/Friday...")
