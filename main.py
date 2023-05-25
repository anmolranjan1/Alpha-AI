import win32com.client
import speech_recognition as sr
import os
import webbrowser
import datetime
import random
import openai
from config import apikey
import numpy as np

chatStr = ""
def chat(query):
    global chatStr
    print(chatStr)
    openai.api_key = apikey
    chatStr += f"Harry: {query}\n Jarvis: "
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt= chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    # todo: Wrap this inside of a  try catch block
    speaker.Speak(response["choices"][0]["text"])
    chatStr += f"{response['choices'][0]['text']}\n"
    return response["choices"][0]["text"]

def ai(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"

    response = openai.Completion.create(
      model="text-davinci-003",
      prompt=prompt,
      temperature=1,
      max_tokens=256,
      top_p=1,
      frequency_penalty=0,
      presence_penalty=0
    )
    # todo: Wrap this inside of a try catch block
    # print(response["choices"][0]["text"])
    text += response["choices"][0]["text"]
    if not os.path.exists("Openai"):
        os.mkdir("Openai")

    # with open(f"Openai/prompt- {random.randint(1, 2343434356)}", "w") as f:
    with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip()}.txt", "w") as f:
        f.write(text)

speaker = win32com.client.Dispatch("SAPI.SpVoice")
# def say(text):
#     os.system(f'say "{text}"')

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold =  0.6
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry from Jarvis"

if __name__ == '__main__':
    print('Welcome to Jarvis A.I')
    speaker.Speak("Jarvis A I")
    # say("Jarvis A.I")
    while True:
        print("Listening")
        query = takeCommand()
        # todo: Add more sites
        sites = [["youtube", "https://youtube.com"], ["wikipedia", "https://wikipedia.com"], ["google", "https://www.google.com"],]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speaker.Speak(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])
                # speaker.Speak(query)
                # say(text)
        # todo: Add a feature to play a specific song
        if "open music" in query:
            musicPath = "C:/Users/anmol/Downloads/downfall-21371.mp3"
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

        elif "open PDF".lower() in query.lower():
            os.startfile("C:/Users/anmol/OneDrive/Desktop/CAT 2 ADDA SOLUTION.pdf")

        elif "Using artificial intelligence".lower() in query.lower():
            ai(prompt=query)

        elif "Jarvis Quit".lower() in query.lower():
            exit()

        elif "reset chat".lower() in query.lower():
            chatStr = ""

        else:
            print("Chatting...")
            chat(query)