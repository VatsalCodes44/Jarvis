import win32com.client
import time
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import smtplib
from playsound import playsound
import requests


def speak(text):
    """ this function is defined so our jarvis can speak"""
    speaker=win32com.client.Dispatch("SAPI.SpVoice")
    speaker.speak(text)
    
    
def wishMe():
    """this function greets and telll about itself"""
    hour=int(time.strftime("%H"))
    if 0<=hour<12:
        speak("Good Morning")
    elif 12<=hour<18:
        speak("Good Afternoon")
    elif 18<=hour<21:
        speak("Good Evening")
    else:
        speak("Good Night")
    speak("I am jarvis, Vatsal, I am your personal voice assistant, How can i help you")
    
    
def takeCommand():
    """it inputs the microphone input from user and returns
    string output"""
    r=sr.Recognizer()
    query="speak again"
    with sr.Microphone() as source:
        print("listening......")
        r.pause_threshold=.8
        r.energy_threshold=300
        
        audio=r.listen(source)
    try:
        print("Recognizing")
        query=r.recognize_google(audio)
        print(f"The user said: {query}")
        
    except Exception:
        speak("please speak again")
        takeCommand()
    return query


def send_email(to,content):
    '''to send email'''
    f=open("gmail_password.txt","r")
    password=int(f.read())
    f.close()
    server=smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login("vatsalmahajan711@gmail.com", password)
    server.sendmail("vatsalmahajan711@gmail.com", to, content)
    print("Email sent successfully!")
    server.close()
    

if __name__=="__main__":
    wishMe()
    while True:
        command=takeCommand().lower().strip()
        if "wikipedia" in command:
            command=command.replace("wikipedia","")
            results=wikipedia.summary(command,sentences=2)
            speak("according to wikipedia")
            print(results)
            speak(results)
        elif "open youtube" in command:
            webbrowser.open("youtube.com")
        elif "open google" in command:
            webbrowser.open("google.com")
        elif"dog sound" in command:
            sound_file = "C:\\Users\\mahaj\\Downloads\\dog-barking-70772.mp3"
            playsound(sound_file)                   
        elif "play music" in command:
            webbrowser.open("https://www.youtube.com/watch?v=60ItHLz5WEA")
        elif "open spotify" in command:
            os.system("spotify")
        elif "the time" in command:
            speak("its"+time.strftime("%H:%M:%S"))
        elif "the date" in command:
            speak("its"+time.strftime("%d:%B:%Y"))
        elif "send email to vatsal" in command:
            try:
                speak("what should i send")
                content=takeCommand()
                to="mahajanvatsal44@gmail.com"
                send_email(to,content)
                speak("email has been sent")
            except:
                print("unable to send please try again")
        elif "exit" in command.strip():
            break
        else:
            command=command.replace("wikipedia","")
            results=wikipedia.summary(command,sentences=1)
            print(results)
            speak(results)