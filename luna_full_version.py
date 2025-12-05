import speech_recognition as sr
import pyttsx3
import datetime
import wikipedia
import webbrowser
import os
import pywhatkit
import requests
import time
import socket
import threading
import pyautogui
import pygetwindow as gw
import subprocess
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import ctypes
import ast
import operator as op

# === Global Text-to-Speech Engine ===
engine = pyttsx3.init()
engine.setProperty('rate', 140)
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speak(text):
    print("Luna:", text)
    engine.say(text)
    engine.runAndWait()

# === Connectivity Check ===
def is_connected():
    try:
        socket.create_connection(("www.google.com", 80), timeout=3)
        return True
    except OSError:
        return False

# === Listen to the user ===
recognizer = sr.Recognizer()
def listen(timeout=5, phrase_time_limit=8):
    with sr.Microphone() as src:
        recognizer.adjust_for_ambient_noise(src, duration=1.2)
        try:
            print("Listening...")
            audio = recognizer.listen(src, timeout=timeout, phrase_time_limit=phrase_time_limit)
            command = recognizer.recognize_google(audio).lower()
            print("You said:", command)
            return command
        except sr.WaitTimeoutError:
            return ""
        except sr.UnknownValueError:
            return ""
        except sr.RequestError as e:
            print("Speech API error:", e)
            return ""

# === Joke Fetch ===
def get_joke_from_api():
    if not is_connected():
        speak("You're offline. I can't fetch a joke right now.")
        return
    try:
        res = requests.get("https://official-joke-api.appspot.com/random_joke", timeout=5)
        res.raise_for_status()
        joke = res.json()
        speak("Here's a joke.")
        speak(joke.get('setup', ''))
        speak(joke.get('punchline', ''))
    except Exception as ex:
        speak("Something went wrong while getting the joke.")
        print("Error:", ex)

# === Offline Responses ===
def offline_reply(prompt):
    responses = {
        "how are you": "I'm just a bunch of code, but I'm doing great!",
        "who are you": "I'm Luna, your offline voice assistant.",
        "what is your name": "My name is Luna.",
        "thank you": "You're welcome!",
        "hello": "Hi there!",
        "hi": "Hello! How can I help?"
    }
    for key in responses:
        if key in prompt:
            speak(responses[key])
            return
    speak("Sorry, I didn't understand that and I'm offline, so I can't look it up.")

# === WhatsApp Messaging ===
def send_whatsapp_message():
    speak("Opening WhatsApp.")
    open_whatsapp_store_version()
    time.sleep(5)

    speak("Please say the name of the contact.")
    contact = listen()
    if not contact:
        speak("I didn't catch the contact name.")
        return

    pyautogui.moveTo(100, 100)
    pyautogui.click()
    time.sleep(1)
    pyautogui.write(contact)
    time.sleep(2)
    pyautogui.moveTo(100, 200)
    pyautogui.click()
    time.sleep(1)

    speak("What message should I send?")
    message = listen()
    if not message:
        speak("I didn't catch the message.")
        return

    pyautogui.write(message)
    time.sleep(1)

    speak("Say 'ok' when you're ready to send the message.")
    confirmation = listen(timeout=8, phrase_time_limit=5)
    if "ok" in confirmation.lower():
        pyautogui.press('enter')
        speak("Message sent successfully.")
    else:
        speak("Message not sent.")

# === Volume Control ===
def control_volume(command):
    if "increase" in command:
        pyautogui.press("volumeup")
        speak("Volume increased.")
    elif "decrease" in command:
        pyautogui.press("volumedown")
        speak("Volume decreased.")

# === Weather Information ===
def get_weather():
    speak("Checking the weather for your location.")
    try:
        webbrowser.open("https://weather.com/weather/today")
        speak("Here is the current weather information for your area.")
    except Exception as e:
        print("Weather error:", e)
        speak("Failed to open the weather website.")

# === File Operations ===
def create_file():
    speak("What should be the name of the file?")
    filename = listen().replace(" ", "_") + ".txt"
    with open(filename, 'w') as f:
        speak("What should I write in the file?")
        content = listen()
        f.write(content)
    speak(f"File {filename} created.")

def search_file():
    speak("What file are you looking for?")
    name = listen().lower()
    for root, dirs, files in os.walk("."):
        for file in files:
            if name in file.lower():
                speak(f"Found file: {file} at {os.path.abspath(os.path.join(root, file))}")
                return
    speak("File not found.")

# === System Commands ===
def system_command(command):
    if "lock" in command:
        speak("Locking the screen.")
        ctypes.windll.user32.LockWorkStation()
    elif "shutdown" in command:
        speak("Shutting down.")
        os.system("shutdown /s /t 1")
    elif "restart" in command:
        speak("Restarting.")
        os.system("shutdown /r /t 1")

# === News ===
def get_news():
    speak("Opening latest news headlines.")
    try:
        webbrowser.open("https://news.google.com/topstories?hl=en-IN&gl=IN&ceid=IN:en")
        speak("Here are the top news stories.")
    except Exception as e:
        print("News error:", e)
        speak("Failed to open the news website.")

# === Safe Calculator ===
SAFE_OPERATORS = {
    ast.Add: op.add, ast.Sub: op.sub, ast.Mult: op.mul,
    ast.Div: op.truediv, ast.Pow: op.pow, ast.Mod: op.mod,
    ast.USub: op.neg
}

def eval_expr(expr):
    def _eval(node):
        if isinstance(node, ast.Num):
            return node.n
        if isinstance(node, ast.BinOp):
            return SAFE_OPERATORS[type(node.op)](_eval(node.left), _eval(node.right))
        if isinstance(node, ast.UnaryOp):
            return SAFE_OPERATORS[type(node.op)](_eval(node.operand))
        raise TypeError("Unsupported expression")
    return _eval(ast.parse(expr, mode='eval').body)

def calculate_expression():
    speak("What should I calculate?")
    expr = listen()
    try:
        result = eval_expr(expr)
        speak(f"The result is {result}")
    except Exception:
        speak("Sorry, I couldn't calculate that.")

# === Email Sending ===
def send_email():
    speak("Who should I send the email to?")
    to = listen()
    speak("What is the subject?")
    subject = listen()
    speak("What is the message?")
    body = listen()

    msg = MIMEMultipart()
    msg['From'] = os.environ.get("ASSISTANT_EMAIL")
    msg['To'] = to
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587, timeout=10)
        server.starttls()
        server.login(os.environ.get("ASSISTANT_EMAIL"), os.environ.get("ASSISTANT_EMAIL_PASSWORD"))
        server.send_message(msg)
        server.quit()
        speak("Email sent successfully.")
    except Exception as e:
        print("Email error:", e)
        speak("Failed to send email.")

# === Close Functions ===
def close_facebook_tab():
    try:
        fb_windows = gw.getWindowsWithTitle("Facebook")
        for win in fb_windows:
            if ("chrome" in win.title.lower() or "edge" in win.title.lower() or "firefox" in win.title.lower()):
                win.minimize()
                win.restore()
                win.activate()
                time.sleep(0.5)
                pyautogui.hotkey('ctrl', 'w')
                speak("Facebook tab closed.")
                return
        speak("No Facebook tab is currently open.")
    except Exception as e:
        print("Error closing Facebook tab:", e)
        speak("Sorry, I couldn't close Facebook right now.")

def close_gmail_tab():
    try:
        gmail_windows = gw.getWindowsWithTitle("Gmail")
        for win in gmail_windows:
            if ("chrome" in win.title.lower() or "edge" in win.title.lower() or "firefox" in win.title.lower()):
                win.minimize()
                win.restore()
                win.activate()
                time.sleep(0.5)
                pyautogui.hotkey('ctrl', 'w')
                speak("Gmail tab closed.")
                return
        speak("No Gmail tab is currently open.")
    except Exception as e:
        print("Error closing Gmail tab:", e)
        speak("Sorry, I couldn't close Gmail right now.")

def close_youtube_tab():
    try:
        youtube_windows = gw.getWindowsWithTitle("YouTube")
        for win in youtube_windows:
            if ("chrome" in win.title.lower() or "edge" in win.title.lower() or "firefox" in win.title.lower()):
                win.minimize()
                win.restore()
                win.activate()
                time.sleep(0.5)
                pyautogui.hotkey('ctrl', 'w')
                speak("YouTube tab closed.")
                return
        speak("No YouTube tab is currently open.")
    except Exception as e:
        print("Error closing YouTube tab:", e)
        speak("Sorry, I couldn't close YouTube right now.")

# === close_app ===
def close_app(command):
    apps = {
        'notepad': 'notepad.exe',
        'word': 'winword.exe',
        'vs code': 'Code.exe',
        'visual studio code': 'Code.exe',
        'calculator': 'Calculator.exe',
        'command prompt': 'cmd.exe',
        'chrome': 'chrome.exe',
        'edge': 'msedge.exe',
        'firefox': 'firefox.exe',
        'spotify': 'Spotify.exe',
        'whatsapp': 'WhatsApp.exe',
    }

    if 'youtube' in command:
        close_youtube_tab()
        return
    if 'gmail' in command:
        close_gmail_tab()
        return
    if 'facebook' in command:
        close_facebook_tab()
        return

    matched = False
    for app, process in apps.items():
        if app in command:
            speak(f"Closing {app.capitalize()}")
            os.system(f"taskkill /f /im {process}")
            matched = True
            break

    if not matched:
        speak("I didn't recognize which app to close.")

# === Launch WhatsApp (Microsoft Store Version) ===
def open_whatsapp_store_version():
    try:
        os.system('start shell:AppsFolder\\5319275A.WhatsAppDesktop_cv1g1gvanyjgm!App')
    except Exception as e:
        speak("Unable to open WhatsApp.")
        print("WhatsApp launch error:", e)

# === Command Processor ===
def handle(command):
    command = command.lower()

    if 'open youtube' in command:
        speak("Opening YouTube")
        webbrowser.open("https://www.youtube.com")

    elif 'open facebook' in command:
        speak("Opening Facebook")
        webbrowser.open("https://www.facebook.com")

    elif 'open google' in command:
        speak("What should I search?")
        query = listen()
        if query:
            speak(f"Searching Google for {query}")
            webbrowser.open(f"https://www.google.com/search?q={query}")
        else:
            speak("I didn't catch what to search.")

    elif 'open wikipedia' in command or 'tell me about' in command:
        if 'wikipedia' in command:
            speak("What should I search on Wikipedia?")
            query = listen()
        else:
            query = command.replace('tell me about', '').strip()
        if query:
            try:
                result = wikipedia.summary(query, sentences=2)
                speak("According to Wikipedia:")
                speak(result)
            except Exception as e:
                print("Wikipedia error:", e)
                speak("Sorry, I couldn't find anything.")
        else:
            speak("I didn't catch what to search.")

    elif 'play' in command:
        song = command.replace('play', '').strip()
        if song:
            speak(f"Playing {song} on YouTube")
            pywhatkit.playonyt(song)
        else:
            speak("Please tell me the song name.")

    elif 'time' in command:
        time_now = datetime.datetime.now().strftime('%I:%M %p')
        speak(f"The time is {time_now}")

    elif 'date' in command:
        date_today = datetime.datetime.now().strftime('%A, %B %d, %Y')
        speak(f"Today is {date_today}")

    elif 'word' in command or 'microsoft word' in command:
        speak("Opening Microsoft Word")
        os.system("start winword.exe")

    elif "open whatsapp" in command:
        speak("Opening WhatsApp")
        open_whatsapp_store_version()

    elif 'open vs code' in command or 'visual studio code' in command:
        speak("Opening Visual Studio Code")
        os.system("code")

    elif 'open notepad' in command:
        speak("Opening Notepad")
        threading.Thread(target=lambda: os.system("notepad.exe")).start()

    elif 'open calculator' in command:
        speak("Opening Calculator")
        threading.Thread(target=lambda: os.system("calc.exe")).start()

    elif 'open command prompt' in command or 'open cmd' in command:
        speak("Opening Command Prompt")
        os.system("start cmd")

    elif 'joke' in command:
        get_joke_from_api()

    elif 'volume' in command:
        control_volume(command)

    elif 'weather' in command:
        get_weather()

    elif 'create file' in command:
        create_file()

    elif 'find file' in command or 'search file' in command:
        search_file()

    elif 'shutdown' in command or 'restart' in command or 'lock' in command:
        system_command(command)

    elif 'news' in command:
        get_news()

    elif 'calculate' in command:
        calculate_expression()

    elif 'send email' in command:
        send_email()

    elif any(phrase in command for phrase in ["send whatsapp message", "send a whatsapp message", "send message on whatsapp"]):
        send_whatsapp_message()

    elif 'help' in command or 'what can you do' in command:
        speak("I can open websites and apps, tell jokes, control system, fetch weather and news, perform calculations, and send emails or WhatsApp messages.")

    elif 'goodbye' in command or 'exit' in command or 'bye' in command:
        speak("Goodbye! Have a great day.")
        exit()

    elif 'close' in command:
        close_app(command)

    else:
        offline_reply(command)

# === Main Loop ===
def main():
    hour = datetime.datetime.now().hour
    greet = "Good morning!" if hour < 12 else "Good afternoon!" if hour < 18 else "Good evening!"
    speak(f"{greet} Luna is online. Say 'Luna' to wake me up.")

    wake_words = ['luna', 'hey luna', 'ok luna']


    while True:
        try:
            command = listen(timeout=5, phrase_time_limit=5)
            if any(word in command for word in wake_words):
                speak("Yes, I'm listening.")
                user_command = listen(timeout=6, phrase_time_limit=10)
                if user_command:
                    handle(user_command)
                else:
                    speak("I didn't catch that. Please try again.")
                speak("Say 'Luna' to wake me up.")
            else:
                time.sleep(1)
        except KeyboardInterrupt:
            speak("Shutting down. Goodbye!")
            break

if __name__ == "__main__":
    main()
