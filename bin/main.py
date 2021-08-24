import speech_recognition as sr
import pyttsx3
import keyboard

import strings

# Constructs a new TTS engine instance
engine = pyttsx3.init()
# noinspection PyUnresolvedReferences
# Set voice type
engine.setProperty('voice', engine.getProperty('voices')[1].id)
# Set speed of speech (words per minute)
engine.setProperty('rate', 150)


def SwitchKeyboardLanguage():
    keyboard.send("shift+alt", do_press=True)


def CommandAnalysis(command):
    splitted_command = command.split(" ")
    if splitted_command[0] == "Sergey":
        if splitted_command[1] == "switch" and splitted_command[2] == "keyboard" and splitted_command[3] == "language":
            SwitchKeyboardLanguage()


def AssistantSays(text):
    print(f'{strings.name_of_assistant}: {text}')


# Voiceover a command
def Speak(audio):
    # Adds an utterance to speak to the queue
    engine.say(audio)
    engine.runAndWait()


# Speech recognition
def CommandRecognition():
    # Creates a new `Recognizer` instance
    r = sr.Recognizer()
    with sr.Microphone() as source:
        AssistantSays("Listening...")
        # Listening for speech
        audio = r.listen(source)
        # noinspection PyBroadException
        try:
            query = r.recognize_google(audio)
            return query
        except:
            AssistantSays("Try Again")


def Main():
    command = "Sergey switch keyboard language"
    CommandAnalysis(command)
