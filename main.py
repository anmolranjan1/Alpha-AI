import win32com.client
import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import random
import smtplib
import pyaudio
import PyPDF2
import pyjokes
import pywhatkit as kit
import psutil
import subprocess
import time
# import openai
# from config import apikey

speaker = win32com.client.Dispatch("SAPI.SpVoice")
# def say(text):
#     os.system(f'say "{text}"')

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speaker.Speak("Good morning Sir..")
        # say("Good morning Sir...")
    elif hour >= 12 and hour < 17:
        speaker.Speak("Good afternoon Sir..")
        # say("Good afternoon Sir..")
    else:
        speaker.Speak("Good evening Sir")
        # say("Good evening Sir")
    print("I am your personal assistant Alpha.")
    speaker.Speak("How may I help you? ")
    # say("How may I help you? ")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold =  0.6
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            # query = input("Enter: ")
            print(f"User said: {query}")
            # return query
        except Exception as e:
            print("Say that again please...")
            return "None"
        return query

# def sendEmail(to, content):
#     server=smtplib.SMTP('smtp.gmail.com', 587)
#     server.ehlo()
#     server.starttls
#     server.login('myemail', 'password')
#     server.sendmail('myemail', to , content)

# chatStr = ""
# def chat(query):
#     global chatStr
#     print(chatStr)
#     openai.api_key = apikey
#     chatStr += f"User: {query}\n Alpha: "
#     response = openai.Completion.create(
#         model="text-davinci-003",
#         prompt= chatStr,
#         temperature=0.7,
#         max_tokens=256,
#         top_p=1,
#         frequency_penalty=0,
#         presence_penalty=0
#     )
#     # todo: Wrap this inside of a  try catch block
#     speaker.Speak(response["choices"][0]["text"])
#     chatStr += f"{response['choices'][0]['text']}\n"
#     return response["choices"][0]["text"]
#
# def ai(prompt):
#     openai.api_key = apikey
#     text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"
#
#     response = openai.Completion.create(
#       model="text-davinci-003",
#       prompt=prompt,
#       temperature=1,
#       max_tokens=256,
#       top_p=1,
#       frequency_penalty=0,
#       presence_penalty=0
#     )
#     # todo: Wrap this inside of a try catch block
#     # print(response["choices"][0]["text"])
#     text += response["choices"][0]["text"]
#     if not os.path.exists("Openai"):
#         os.mkdir("Openai")
#
#     # with open(f"Openai/prompt- {random.randint(1, 2343434356)}", "w") as f:
#     with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip()}.txt", "w") as f:
#         f.write(text)

if __name__ == '__main__':
    print('Welcome to Alpha A.I')
    wishMe()
    # speaker.Speak("Alpha A I")
    # say("Alpha A.I")
    while True:
        print("Listening...")
        query = takeCommand().lower()

        # todo: Add more sites
        sites = [["youtube", "https://youtube.com"], ["wikipedia", "https://wikipedia.com"],
                 ["google", "https://www.google.com"], ["stack overflow", "https://stackoverflow.com"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speaker.Speak(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])
                # speaker.Speak(query)
                # say(text)

        if "who made you" in query or "who created you" in query:
            speaker.Speak("I have been created by Anmol Ranjan.")

        elif 'how are you' in query:
            speaker.Speak("I am fine! What about you?")

        elif 'wikipedia' in query or 'search' in query:
            speaker.Speak('Searching Wikipedia....')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speaker.Speak("According to Wikipedia")
            speaker.Speak(results)
            print(results)

        elif "joke" in query:
            print("Choose a category for joke")
            speaker.Speak("Choose a category for joke")
            print('1. Neutral', '2. chuck', '3. all')
            category = takeCommand().lower()
            joke = pyjokes.get_joke(language='en', category=category)
            print(joke)
            speaker.Speak(joke)

        elif "open music" in query:
            # todo: Add path to the song you want to play
            musicPath = "C:\\Users\\anmol\\Downloads\\music"
            os.startfile(musicPath)
            # import subprocess, sys
            # opener = "open" if sys.platform == "darwin" else "xdg-open"
            # subprocess.call([opener, musicPath])

            # os.system(f"open {musicPath}")

        elif "the time" in query:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            # hour = datetime.datetime.now().strftime("%H")
            # min = datetime.datetime.now().strftime("%M")
            speaker.Speak(f"Sir the time is {strfTime}")
            # speaker.Speak(f"Sir the time is {hour} and {min} minutes")

        # elif "Using artificial intelligence".lower() in query.lower():
        #     ai(prompt=query)

        elif 'flip a coin' in query:
            coin = random.randrange(2)
            if (coin == 1):
                speaker.Speak("Heads")
            else:
                speaker.Speak("Tails")

        elif "write a note" in query:
            speaker.Speak("What should I write sir?")
            note = takeCommand()
            file = open('notes.txt', 'w')
            speaker.Speak("Should I include  date and time?")
            ans = takeCommand()
            if 'yes' in ans:
                strTime = datetime.datetime.now().strftime(" %H %M %S")
                file.write(strTime)
                file.write("-")
                file.write(note)
            else:
                file.write(note)

        elif "show notes" in query:
            speaker.Speak("Showing Notes")
            file = open("notes.txt", "r")
            print(file.read())
            speaker.Speak(file.read(6))

        # plays the video on youtube
        elif 'play online' in query:
            speaker.Speak('What is the title of video')
            video = takeCommand().lower()
            kit.playonyt(video)

        elif 'read pdf' in query:
            book = open('sample.pdf', 'rb')
            pdfReader = PyPDF2.PdfFileReader(book)
            pages = pdfReader.numPages
            speaker.Speak(f"Total number of pages in this book are {pages}")
            speaker.Speak("Please enter the page number you want me to read.")
            pg = int(input())
            page = pdfReader.getPage(pg)
            text = page.extractText()
            speaker.Speak(text)

        elif 'battery percentage' in query:
            battery = psutil.sensors_battery()
            bat = battery.percent
            speaker.Speak(f"Sir, your system has {bat} percent battery")
            print(f"Sir, your system has {bat} percent battery")

        elif 'shutdown system' in query:
            speaker.Speak("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speaker.Speak("Hibernating")
            subprocess.call("shutdown / h")

        elif "log off" in query or "sign out" in query:
            speaker.Speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        else:
            print("Please use the README file for more commands available to be executed..")

    #   elif "reset chat".lower() in query.lower():
    #       chatStr = ""
    #
    #   else:
    #     print("Chatting...")
    #     chat(query)

        # Speaker.speak(query)
