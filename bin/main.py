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
    # noinspection PyBroadException
    try:
        keyboard.send("shift+alt")
        AssistantSays("Language is switched")
    except:
        pass


def CommandAnalysis(command):
    splitted_command = command.split(" ")

    # Replace the wrong name
    if splitted_command[0] == "Sergei":
        splitted_command[0] = "Sergey"

    # Create the right command for print to user
    command = " ".join(map(str, splitted_command))
    UserSays(command)

    if splitted_command[0] == strings.name_of_assistant:
        if splitted_command[1] == "switch" or splitted_command[1] == "change":
            if splitted_command[2] == "keyboard" and splitted_command[3] == "language" or\
                    splitted_command[2] == "language":
                SwitchKeyboardLanguage()
    else:
        print(command)


def AssistantSays(text):
    print(f'{strings.name_of_assistant}: {text}')


def UserSays(text):
    print(f'You: {text}')


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
            command = r.recognize_google(audio)
            return command
        except:
            AssistantSays("Try Again")


def Start():
    print(strings.welcome_str)
    # command = CommandRecognition()
    # command = "Sergey switch language"
    command = "Sergei switch keyboard language"
    CommandAnalysis(command)


def Main():
    Start()
