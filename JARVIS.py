import speech_recognition as sr
import os
import webbrowser
import win32com.client
import datetime

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def say(text):
    speaker.Speak(text)


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.6
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language = "en-in")
            print(f"User said: {query}")
            return query
        except sr.UnknownValueError:
            return "Sorry Sir, Can you Repeat Again"
        except sr.RequestError as e:
            return (f"Could not complete this request Sir; {e}")
 
   
say("Welcome Sir! I am JARVIS")
say("How may I help You")
while True:
    
    print("Listening...")
    query = takeCommand()
    sites = [
        ["my playlist","https://www.youtube.com/watch?v=Qrm7L5t2Ygc&list=PLjUo7wQwvXXylXpXZyYZbw5_ToLy-l-Go"],
        ["youtube","https://www.youtube.com"],
        ["songs","https://wynk.in/music"],
        ["linkedin","https://www.linkedin.com/feed/"],
        ["x","https://x.com/home"],
        ["google","https://www.google.com"]
        ]
    for site in sites:
        if f"Open {site[0]}".lower() in query.lower():
            say(f"Opening {site[0]} Sir...")
            webbrowser.open(site[1])
    
    if "open file" in query.lower():
        Path =r"D:\Ishu Katiyar Offer Letter.pdf"
        say("Opening your file Sir...")
        os.startfile(Path)
    
    elif "the time" in query:
        strfTime = datetime.datetime.now().strftime("%H:%M")
        say(f"Sir the time is {strfTime}")
    
    elif "microsoft edge".lower() in query.lower():
        Path = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk"
        say("Opening Microsoft edge Sir...")
        os.startfile(Path )
    
    elif "open brave".lower() in query.lower():
        Path = r"C:\Users\Ishu\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Brave.lnk"
        say("Opening Brave Sir...")
        os.startfile(Path )

    elif "Jarvis Quit".lower() in query.lower():
            say("Thank You Sir... Have a great day...")
            exit()
    
    elif "Hello jarvis".lower() in query.lower():
        say("Hello Sir...")
    
    elif "Creator".lower() in query.lower():
        say("Mr. Ishu Sir...")
    
    # say(query)
    