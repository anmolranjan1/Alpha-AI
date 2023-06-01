# import win32com.client
import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import random
import pyaudio
import PyPDF2
import pyjokes
import pywhatkit as kit
import psutil
import subprocess
import time
import password
import tic_tac_toe

engine = pyttsx3.init('sapi5')
# by using init method we will store engine instance into variable , sapi5 is Microsoft speech api
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

# speaker = win32com.client.Dispatch("SAPI.SpVoice")

def say(text):
    engine.say(text)
    engine.runAndWait()
    # os.system(f'say "{text}"')


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        say("Good morning Sir..")
    elif 12 <= hour < 17:
        say("Good afternoon Sir..")
    else:
        say("Good evening Sir")
    print("I am your personal assistant Alpha.")
    say("How may I help you? ")


def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold =  0.6
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            # return query
        except Exception as e:
            print("Say that again please...")
            return "None"
        return query


def quiz():
    questions = [
        ("When was the first known use of the word 'quiz'", "1781"),
        ("Which built-in function can get information from the user", "input"),
        ("Which keyword do you use to loop over a given list of elements", "for")
    ]
    for question, correct_answer in questions:
        answer = input(f"{question}? ")
        if answer == correct_answer:
            print("Correct!")
            say("Correct")
        else:
            print(f"The answer is {correct_answer!r}, not {answer!r}")
            say(f"The answer is {correct_answer!r}, not {answer!r}")


if __name__ == '__main__':
    print('Welcome to Alpha A.I')
    wishMe()
    # say("Alpha A I")
    # say("Alpha A.I")
    while True:
        print("Listening...")
        query = takecommand().lower()

        # todo: Add more sites
        sites = [["youtube", "https://youtube.com"], ["wikipedia", "https://wikipedia.com"],
                 ["google", "https://www.google.com"], ["stack overflow", "https://stackoverflow.com"],
                 ["twitter", "http://www.twitter.com"], ["linkedIn", "https://www.linkedin.com"],
                 ["instagram", "https://www.instagram.com"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])

        if "who made you" in query or "who created you" in query:
            say("I have been created by Anmol Ranjan.")

        elif 'how are you' in query:
            say("I am fine! What about you?")

        elif 'wikipedia' in query or 'search' in query:
            say('Searching Wikipedia....')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            say("According to Wikipedia")
            say(results)
            print(results)
            # to make a file with the content
            file = open('wikipedia.txt', 'w')
            file.write(results)

        elif "write a note" in query:
            say("What should I write sir?")
            note = takecommand()
            file = open('notes.txt', 'w')
            say("Should I include  date and time?")
            ans = takecommand()
            if 'yes' in ans:
                strTime = datetime.datetime.now().strftime(" %H %M %S")
                file.write(strTime)
                file.write("\n")
                file.write(note)
            else:
                file.write(note)

        elif "show the notes" in query:
            say("Showing the Note")
            file = open("notes.txt", "r")
            print(file.read())
            say(file.read(6))

        elif "joke" in query:
            print("Choose a category for joke")
            say("Choose a category for joke")
            print('1. Neutral', '2. chuck', '3. all')
            category = takecommand().lower()
            joke = pyjokes.get_joke(language='en', category=category)
            print(joke)
            say(joke)

        elif "open music" in query:
            # todo: Add path to the song you want to play
            musicPath = "sample-music.mp3"
            os.startfile(musicPath)
            # import subprocess, sys
            # opener = "open" if sys.platform == "darwin" else "xdg-open"
            # subprocess.call([opener, musicPath])

            # os.system(f"open {musicPath}")

        elif "the time" in query:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            # hour = datetime.datetime.now().strftime("%H")
            # min = datetime.datetime.now().strftime("%M")
            say(f"Sir the time is {strfTime}")
            # say(f"Sir the time is {hour} and {min} minutes")

        elif 'flip a coin' in query:
            coin = random.randrange(2)
            if coin == 1:
                say("Heads")
            else:
                say("Tails")

        # plays the video on YouTube
        elif 'play online' in query:
            say('What is the title of video')
            video = takecommand().lower()
            kit.playonyt(video)

        elif 'read pdf' in query.lower():
            book = open('sample-pdf.pdf', 'rb')
            pdf = PyPDF2.PdfReader(book)
            pages = len(pdf.pages)
            say(f"Total number of pages in this book are {pages}")
            say("Please enter the page number you want me to read.")
            pg = int(input("Page No.:"))
            page = pdf.pages[pg]
            text = page.extract_text()
            say(text)

        # generating a strong password
        elif 'generate password' in query.lower():
            password.generate(537)

        #checks password strength
        elif 'check password strength' in query.lower() or 'check' in query.lower():
            sample = input("Enter the password: ")
            result = password.checker(sample)
            # result = password.checker(sample)
            print(result)
            say(result)

        elif 'start countdown' in query.lower() or 'countdown' in query.lower():
            count = int(input("For how many seconds?? "))
            while count:
                mins, secs = divmod(count, 60)
                timeformat = '{:02d}:{:02d}'.format(mins, secs)
                print(timeformat, end='\r')
                say(timeformat)
                time.sleep(1)
                count -= 1
            say("stopped")

        # a small quiz
        elif 'quiz' in query.lower():
            quiz()

        # tic tac toe game
        elif 'play' in query.lower():
            tic_tac_toe.game()
            # os.system('tic_tac_toe.py')

        # temperature convertor
        elif 'convert' in query.lower():
            if 'into celsius' in query.lower():
                fahrenheit = float(input("Enter temperature in fahrenheit: "))
                celsius = (fahrenheit - 32) / 1.8
                say(f"{fahrenheit} degree Fahrenheit is equal to {celsius} degree Celsius.")
                print(f"{fahrenheit} degree Fahrenheit is equal to {celsius} degree Celsius.")
            else:
                celsius = float(input("Enter temperature in celsius: "))
                fahrenheit = (celsius * 1.8) + 32
                say(f"{celsius} degree Celsius is equal to {fahrenheit} degree Fahrenheit.")
                print(f"{fahrenheit} degree Fahrenheit is equal to {celsius} degree Celsius.")

        elif 'battery percentage' in query:
            battery = psutil.sensors_battery()
            bat = battery.percent
            say(f"Sir, your system has {bat} percent battery")
            print(f"Sir, your system has {bat} percent battery")

        elif 'shutdown' in query:
            say("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')

        elif "restart" in query:
            say("Restarting")
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            say("Hibernating")
            subprocess.call("shutdown / h")

        elif "log off" in query or "sign out" in query:
            say("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif 'quit' in query.lower():
            exit()

        else:
            print("...")
