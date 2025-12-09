import speech_recognition as sr
import os
import webbrowser
import win32com.client
import datetime
from openai import OpenAI

speaker = win32com.client.Dispatch("SAPI.SpVoice")

from config import api_key
if not api_key:
    raise ValueError("API key not found")

client = OpenAI(api_key=api_key)

chatStr = ""
def say(text: str):
    """Speak and print text."""
    print(f"Jarvis: {text}")
    speaker.Speak(text)

def chat(query):
    global chatStr
    print(f"User: {query}")  
    chatStr += f"Ishu: {query}\nJarvis: "
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",   # or "gpt-4o" if your account has it
            messages=[
                {"role": "system", "content": "You are Jarvis, a helpful AI assistant."},
                {"role": "user", "content": query},
            ],
            temperature=0.7,
            max_tokens=256,
        )

        reply = response.choices[0].message.content.strip()
        say(reply)
        chatStr += reply + "\n"
        return reply
    except Exception as e:
        print(f"OpenAI Error: {e}")
        say("Sorry Sir, I couldn't process that request.")
        return str(e)


# def say(text):
#     speaker.Speak(text)


def takeCommand():
    """Listen from microphone and return recognized text."""
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 0.6
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except sr.UnknownValueError:
            say("Sorry Sir, can you repeat again?")
            return ""
        except sr.RequestError as e:
            say("Could not complete this request, Sir.")
            print(f"Speech API error: {e}")
            return ""
 
   
say("Welcome Sir! I am JARVIS")
say("How may I help You")

while True:
    
    query = takeCommand()
    if not query:
        continue  # if nothing recognized, listen again

    q_lower = query.lower()

    sites = [
        ["my playlist","https://www.youtube.com/watch?v=Qrm7L5t2Ygc&list=PLjUo7wQwvXXylXpXZyYZbw5_ToLy-l-Go"],
        ["youtube","https://www.youtube.com"],
        ["songs","https://wynk.in/music"],
        ["linkedin","https://www.linkedin.com/feed/"],
        ["x","https://x.com/home"],
        ["google","https://www.google.com"]
        ]
    handled = False

    for name, url in sites:
        if f"open {name}".lower() in q_lower:
            say(f"Opening {name} Sir...")
            webbrowser.open(url)
            handled = True
            break
    
    # ----- OTHER COMMANDS -----
    if not handled and "open file" in q_lower:
        path = r"D:\Ishu Katiyar Offer Letter.pdf"
        say("Opening your file Sir...")
        os.startfile(path)
        handled = True
    
    elif not handled and "the time" in q_lower:
        strfTime = datetime.datetime.now().strftime("%H:%M")
        say(f"Sir, the time is {strfTime}")
        handled = True
    
    elif not handled and "microsoft edge" in q_lower:
        path = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk"
        say("Opening Microsoft Edge, Sir...")
        os.startfile(path)
        handled = True
    
    elif not handled and "open brave" in q_lower:
        path = r"C:\Users\Ishu\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Brave.lnk"
        say("Opening Brave, Sir...")
        os.startfile(path)
        handled = True

    elif not handled and "jarvis quit" in q_lower:
        say("Thank you Sir... Have a great day...")
        break
    
    elif not handled and "hello jarvis" in q_lower:
        say("Hello Sir...")
        handled = True
    
    elif not handled and "creator" in q_lower:
        say("Mr. Ishu Sir...")
        handled = True

    # ----- DEFAULT: ASK CHATGPT -----
    if not handled:
        chat(query)
    
    # say(query)
    