from __future__ import print_function

import os
import random
import webbrowser
import pyjokes
import pywhatkit
import requests
import speech_recognition as sr
import pyttsx3
import datetime
import os.path
import pytz
import wikipedia
from bs4 import BeautifulSoup
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from tkinter import *
from tkinter import ttk, LabelFrame

SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

days = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']
months = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october',
          'november',
          'december']
day_e = ['st', 'th', 'rd', 'nd']
Intro = ["hello", "hi", 'hai', "good morning", "good afternoon", 'good evening', "hey"]
CALENDER = ["what do i have on", "am i busy", "do i have plans"]
OPERATION = ['plus', '+', 'minus', "-", "*", 'times', 'x', 'divide by', "/", 'modulus']
HRU = ["Hey, I'm Good !", "I'm good", "I'm good, what about you?", "I'm fine, hope you're also fine",
       "Good, how about you?", "Doing fine, and you?", "I'm doing great", "I'm doing Well"]
WRU = ["I'm your Personal Assistant", "You know me right! If not then I'm Violet your Personal Assistant",
       "You developed me, so you must know who I am",
       "Did I forget to introduce myself? I'm Violet your Personal Assistant"]
TY = ["Thank You", "Thank you so much", "Why are you saying thank you?", "My Pleasure", "You're welcome", "Welcome"]
funny = ["Good to know, that I'm funny - Haha !", "You think I'm funny", "Ya, I'm so funny",
         "I'm funny and can also make you laugh, Just ask me to tell a Joke"]


def print_violet(text):
    vocal.set(text)


def speak(text):
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[2].id)
    engine.setProperty("rate", 160)
    engine.say(text)
    engine.runAndWait()


def get_audio():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        said = ""
        try:
            said = r.recognize_google(audio)
        except:
            speak("I did not understand sir")
    return said


def authenticate():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('calendar', 'v3', credentials=creds)
    return service


def get_event(day, service):
    date = datetime.datetime.combine(day, datetime.datetime.min.time())
    end_day = datetime.datetime.combine(day, datetime.datetime.max.time())
    utc = pytz.UTC
    date = date.astimezone(utc)
    end_day = end_day.astimezone(utc)
    events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(),
                                          timeMax=end_day.isoformat(),
                                          singleEvents=True,
                                          orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        speak('No upcoming events found.')
    else:
        speak(f"you have {len(events)} events on this day")

    # Prints the start and name of the next 10 events
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        start_time = start.split("T")[1].replace("Z", "")
        h_up = start_time.split(":")
        a = int(h_up[0]) + 6
        b = int(h_up[1]) + 30
        if a >= 24:
            a = a - 24
        if b >= 60:
            b = b - 60
        h_up[0] = str(a)
        h_up[1] = str(b)
        start_time = ""
        for i in h_up:
            start_time = start_time + i + ":"
        h_up = list(start_time)
        h_up.pop()
        start_time = ""
        for i in h_up:
            start_time += i
        if int(start_time.split(':')[0]) < 12:
            start_time += "am"
        else:
            start_time = str(int(start_time.split(':')[0]) - 12)
            start_time += "pm"
        # start_time+=datetime.timedelta(hours=6,minutes=30)
        # print(start_time)
        speak(event["summary"] + " at " + start_time)


def get_date(text):
    text = text.lower()
    today = datetime.date.today()
    if text.count("today") > 0:
        return today
    day = -1
    day_week = -1
    month = -1
    year = today.year
    for word in text.split():
        if word in months:
            month = months.index(word) + 1
        elif word in days:
            day_week = days.index(word)
        elif word.isdigit():
            day = int(word)
        else:
            for i in day_e:
                found = word.find(i)
                if found > 0:
                    try:
                        day = int(word[:found])

                    except:
                        pass
    if month < today.month and month != -1:
        year = year + 1
    if day < today.day and month == -1 and day != -1:
        month = month + 1
    if month == -1 and day == -1 and day_week != -1:
        current_day_week = today.weekday()
        dif = day_week - current_day_week
        if dif < 0:
            dif += 7
            if text.count("next") >= 1:
                dif += 7
        return today + datetime.timedelta(dif)
    if month == -1 and day == -1:
        return None
    return datetime.date(month=month, day=day, year=year)


def note(text):
    if "excel" in text.lower():
        os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE")
    elif "powerpoint" in text.lower():
        os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE")
    elif "onenote" in text.lower():
        os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\ONENOTE.EXE")
    elif "word" in text.lower():
        os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE")
    elif "vs code" in text.lower():
        os.startfile("E:\\Microsoft VS Code\\Code.exe")
    elif "browser" in text.lower():
        os.startfile("C:\\Users\\faisa\\AppData\\Local\\BraveSoftware\\Brave-Browser\\Application\\brave.exe")
    elif "code blocks" in text.lower():
        os.startfile("E:\\CodeBlocks\\codeblocks.exe")
    elif "python ide" in text.lower():
        os.startfile("E:\\PyCharm Community Edition 2022.1.3\\bin\\pycharm64.exe")
    elif "idea" in text.lower():
        os.startfile("E:\\IntelliJ IDEA Community Edition 2022.2.1\\bin\\idea64.exe")


def greet(text):
    responce = ["Hello!", "Hello Faiz,Good to see you again!", "Hi there, how can I help?", "Hi",
                "hello sir,How are you", "hai faisal,Whats up"]
    for k in Intro:
        if k in text.lower():
            if text.lower() in Intro[0:3]:
                r = random.choice(responce)
                speak(r)
            elif text.lower() in Intro[2:]:
                if "good morning" in text.lower():
                    speak("good morning sir")
                if "good afternoon" in text.lower():
                    speak("good afternoon sir")
                if "good evening" in text.lower():
                    speak("good evening sir")


def calculate(text):
    try:
        if "plus" in text.lower():
            txt = text.replace("plus", "+")
            speak("Your answer is "+eval(txt))
        elif "minus" in text.lower():
            txt = text.replace("minus", "-")
            speak("Your answer is "+eval(txt))
        elif "times" in text.lower():
            txt = text.replace("times", "*")
            speak("Your answer is "+eval(txt))
        elif "x" in text.lower():
            txt = text.replace("x", "*")
            speak("Your answer is "+eval(txt))
        elif "modulus" in text.lower():
            txt = text.replace("modulus", "%")
            speak("Your answer is "+eval(txt))
        elif "divide by" in text.lower():
            txt = text.replace("divide by", "/")
            speak("Your answer is "+eval(txt))
        else:
            speak("Your answer is "+eval(text))
    except:
        speak("try again")
        print_violet("invalid operation")


def caught(text):
    for i in OPERATION:
        if i in text.lower():
            return True
    else:
        return False


def call(text):
    action = "violet"
    text = text.lower()
    if text == action:
        return True
    return False


def get_person(text):
    speak(wikipedia.summary(text, sentences=2))


def get_search(text):
    text = text.lower().split(" ")
    s = ''
    for i in range(2, len(text)):
        s = s + text[i] + " "
    return s


def get_weather(text):
    text = text.lower().split(" ")
    city = text[-1]
    url = "https://www.google.com/search?q=" + "weather" + city
    html = requests.get(url).content
    # getting raw data
    soup = BeautifulSoup(html, 'html.parser')
    temp = soup.find('div', attrs={'class': 'BNeawe iBp4i AP7Wnd'}).text
    strg = soup.find('div', attrs={'class': 'BNeawe tAd8D AP7Wnd'}).text
    # formatting data
    data = strg.split('\n')
    tm = data[0]
    sky = data[1]
    listdiv = soup.findAll('div', attrs={'class': 'BNeawe s3v9rd AP7Wnd'})
    strd = listdiv[5].text
    pos = strd.find('Wind')
    other_data = strd[pos:]
    speak(temp + tm + sky + other_data)


def start_violet():
    SERVICE = authenticate()
    print_violet("Violet started")
    text = get_audio()
    allow = call(text)
    while allow:
        if text == "violet":
            speak("Hello sir!,What can i do for you")
        print_violet("listening.....")
        text = get_audio()

        if "open" in text:
            note(text)
        elif text in Intro:
            greet(text)
        elif "how are you" in text.lower():
            say = random.choice(HRU)
            speak(say)
        elif "tell me something" in text.lower():
            say = random.choice(["I have nothing to say...", "Hmm, you can ask me anything",
                                 "Hmm, you can ask me to tell a joke"])
            speak(say)
        elif "are you robot" in text.lower():
            say = random.choice(
                ["Of course I'm a kind of Robot", "I'm your friend", "Yes I'm a robot, but I'm a good one"])
            speak(say)
        elif "thank you" in text.lower():
            say = random.choice(TY)
            speak(say)
        elif "you are funny" in text.lower():
            say = random.choice(funny)
            speak(say)
        elif "your birthday" in text.lower():
            say = random.choice(["I don't celebrate my Birthday", "My birthday is on 5nd September, 2022"])
            speak(say)
        elif "i have a question" in text.lower():
            say = random.choice(["Ask me", "Ask me, I can help you", "Don't hesitate, ask me", "You can always ask me"])
            speak(say)
        elif "who are you" in text.lower():
            say = random.choice(WRU)
            speak(say)
        elif "which colour you like" in text.lower():
            say = random.choice(["I like all the 7 Colors of a rainbow", "All Colors are my favorite"])
            speak(say)
        elif "do you love me" in text.lower():
            say = random.choice(["Ya, I love you so much", "Ofcourse, I love you", "We're best friends"])
            speak(say)
        elif "are you single" in text.lower():
            say = random.choice(["Haha, I'm always be single", "I'm your Assistant, and I dont want any relationship",
                                 "I'm only for you"])
            speak(say)
        elif "you are smart" in text.lower():
            say = random.choice(["Yes, I'm smart", "Ofcourse I'm smart", "I'm a program, so I'm smart"])
            speak(say)
        elif "i am really sorry" in text.lower():
            say = random.choice(["It's Ok", "No problem"])
            speak(say)
        elif "i am alone" in text.lower():
            say = random.choice(["Don't feel lonely. I'm always with You",
                                 "I can make you feel happy, Just say tell me a joke"])
            speak(say)
        elif "i like your voice" in text.lower():
            say = random.choice(["Hope you love it...", "Thanks, I think this voice suits you the most",
                                 "Thank You So Much", "Ohh, that's good to know"])
            speak(say)
        elif "i am fine" in text.lower():
            speak("Good to hear that sir")
        elif text.lower() in ["how old", "what is your age", "how old are you", "age?"]:
            speak("I am 1 years old!")
        elif text.lower() in ["what is your name", "whats your name?"]:
            speak("I am Violet!")
        elif caught(text):
            calculate(text)
        elif "who is" in text.lower():
            get_person(text)
        elif "search about" in text.lower():
            search = get_search(text)
            webbrowser.open(search)
            break
        elif "say a joke" in text.lower():
            speak(pyjokes.get_joke())
        elif text.lower() in ["quit", "end", "bye"]:
            speak("Bye sir, take care")
            print_violet("Ending Violet")
            break
        elif caught(text):
            calculate(text)
        elif "weather of" in text.lower():
            get_weather(text)
        elif "time now" in text.lower():
            current_time = datetime.date.today().strftime("%I:%M %p")
            speak(current_time)
        elif "today's date" in text.lower():
            today = datetime.date.today().strftime("%B %d, %Y")
            speak(today)
        elif "play" in text.lower():
            song = text.replace("play", '')
            speak("playing" + song)
            pywhatkit.playonyt(song)
            break
        else:
            for k in CALENDER:
                if k in text.lower():
                    date = get_date(text)
                    if date:
                        get_event(date, SERVICE)
                    else:
                        speak("Please Try Again")
    else:
        speak("invalid access name"+text)
        root.destroy()


root = Tk()
vocal = StringVar()
vocal.set("Hello")
root.geometry("400x600")
root.title('VIOLET')
root.config(bg='#FFFFFF')
root.resizable(False, False)
mic_button = Button(root, width=44, height=2, text="START", bg='#7454c7', fg='white', border=2,
                    command=start_violet)
mic_button.place(x=20, y=500)
talk = Label(root, textvariable=vocal, fg='#7454c7', bg='white',
             font=('Microsoft YaHei UI Light', 11, 'bold'))
talk.place(x=180, y=250)

if __name__ == "__main__":
    root.mainloop()
