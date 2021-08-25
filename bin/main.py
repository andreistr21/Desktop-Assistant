import speech_recognition as sr
import pyttsx3
import keyboard
import webbrowser

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


def URLCreator(quary):
    temp = "https://www.google.com/search?q="
    quary = quary.replace(" ", "+")
    url = f'{temp}{quary}'

    return url


def FirstLetterToUpperCase(list_of_words):
    first_word = []
    # Split first word
    first_word.extend(list_of_words[0])
    # Upper first letter
    first_word[0] = first_word[0].upper()
    # Make single world from list
    first_word = "".join(first_word)
    # Replace the word with a corrected one
    list_of_words[0] = first_word


def CommandAnalysis(command):
    splitted_command = command.split(" ")

    # Replace the wrong name
    if splitted_command[0] == "Sergei":
        splitted_command[0] = "Sergey"

    FirstLetterToUpperCase(splitted_command)

    # Create the right command for print to user
    command = " ".join(splitted_command)
    UserSays(command)

    # noinspection PyBroadException
    try:
        # Change keyboard language command
        if splitted_command[0] == strings.name_of_assistant:
            if splitted_command[1] == "switch" or splitted_command[1] == "change":
                if (splitted_command[2] == "keyboard" and splitted_command[3] == "language") or \
                        splitted_command[2] == "language":
                    SwitchKeyboardLanguage()

        # Search in the Internet
        if splitted_command[0] == "What" or splitted_command[0] == "Who":
            if (splitted_command[1] == "is" and splitted_command[2] == "it") or splitted_command[1] == "is":
                quary = " ".join(splitted_command)

                url = URLCreator(quary)
                webbrowser.open(url)
        elif splitted_command[0] == "Search":
            if splitted_command[1] == "about" or splitted_command[1] == "for":
                quary = ""
                # Create right quary
                for i in range(2, len(splitted_command)):
                    quary += splitted_command[i]
                    if i != len(splitted_command) - 1:
                        quary += " "

                url = URLCreator(quary)
                webbrowser.open(url)
            else:
                quary = ""
                # Create right quary
                for i in range(1, len(splitted_command)):
                    quary += splitted_command[i]
                    if i != len(splitted_command) - 1:
                        quary += " "

                url = URLCreator(quary)
                webbrowser.open(url)

    except:
        pass


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


def Main():
    print(strings.welcome_str)
    # command = CommandRecognition()
    # command = "Sergey switch language"
    command = "search huge microwave in the shop"
    CommandAnalysis(command)
