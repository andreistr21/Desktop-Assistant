import traceback
from os import system, popen
from subprocess import run
from webbrowser import open
from keyboard import send
import speech_recognition as sr
import dearpygui.dearpygui as dpg
from urllib.request import Request, urlopen
from bs4 import BeautifulSoup
from spacy import load
from regex import findall, IGNORECASE
from pynput.keyboard import Key, Controller

from resources import Strings
from bin.classes.MasterAudioController import MasterAudioController
from bin.common import Common as common
from bin.TextAnalyser import GetActionAndObject
from bin.classes.Dialog import Dialog
from bin.classes.Screen import Screen
from bin.Functions import SwitchKeyboardLanguage, SearchInTheInternet, OpenProgram, KillProgram, ChangeVolume, MultimediaControl, StopVoiceOver


nlp = load("en_core_web_trf")

keyboard = Controller()
audio_controller = MasterAudioController()
dialog = Dialog()
screen = Screen(dialog)


def AssistantSaysMethCall(text: str):
    """AssistantSays Methode call from class

    :param text: str
        Text to say by assistant
    :return: None
    """
    dialog.AssistantSays(text)


def ViewportResizeMethCall():
    screen.ViewportResize(dialog)


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


def CommandAnalysisCall(sender, app_data):
    # Stop voiceover process if it running now
    if common.voiceover_shared_list[0]:
        StopVoiceOver()

    CommandAnalysis(sender=sender, app_data=app_data)


def CommandAnalysis(sender="", app_data="", use_speech=False, command=""):
    """
    Analise command and run necessary function.

    :param sender:
    :param app_data:
    :param use_speech:
    :param command:
    :return:
    """
    # If use_speech True use command variable
    if not use_speech:
        dpg.set_value(sender, "")

        command = app_data
        # Focus text input item
        dpg.focus_item("Text_input_id")

    splitted_command = command.split(" ")

    FirstLetterToUpperCase(splitted_command)

    # Assembling the command in one string
    command = " ".join(splitted_command)

    splitted_command = command.split(" ")

    # Add "." at the end of the string if necessary
    if command[-1] != ".":
        command = "".join([command, "."])

    is_imperative, action, obj, doc = GetActionAndObject(nlp, command)

    dialog.UserSays(command)

    is_done = False

    # noinspection PyBroadException
    try:
        if is_imperative:
            # Change keyboard language command
            if (
                action == "Switch" or action == "Change"
            ) and obj == "keyboard language":
                SwitchKeyboardLanguage(dialog)
                is_done = True

            # Search in the Internet
            # Quary: Search\Find about\for ...
            elif action == "Search" or action == "Find":
                SearchInTheInternet(dialog, screen, obj)
                is_done = True

            # Open programs
            elif action == "Open" and obj != "you":
                OpenProgram(dialog, obj)
                is_done = True

            # Close programs
            elif action == "Close" and obj != "you":
                is_done = KillProgram(dialog, obj)

            # Change PC volume
            if action == "Set" and obj == "volume":
                volume_percents = None
                for token in doc:
                    if token.pos_ == "NUM":
                        volume_percents = int(token.text)

                if volume_percents is not None:
                    ChangeVolume(audio_controller, dialog, int(volume_percents))
                    is_done = True
            elif action == "Increase" and obj == "volume":
                ChangeVolume(audio_controller, dialog, "+", True)
                is_done = True
            elif action == "Decrease" and obj == "volume":
                ChangeVolume(audio_controller, dialog, "-", True)
                is_done = True

            # Shutdown\Restart PC
            elif action == "Shutdown":
                adv = None
                for token in doc:
                    if token.pos_ == "ADV":
                        adv = token.text

                if obj == "computer" and adv is None:
                    pass
                elif obj == "computer" and adv == "immediately":
                    system("shutdown /s /t 0")
                    is_done = True
            elif action == "Restart":
                adv = None
                for token in doc:
                    if token.pos_ == "ADV":
                        adv = token.text

                if obj == "computer" and adv is None:
                    system("shutdown /r")
                    is_done = True
                elif obj == "computer" and adv == "immediately":
                    system("shutdown /r /t 0")
                    is_done = True
            # Multimedia control
            elif (action == "Stop" or action == "Play") and (
                obj == "music" or obj == "track"
            ):
                MultimediaControl(keyboard, action.lower())
                dialog.AssistantSays("Audio stopped or resumed")

                is_done = True

            # Help menu
            elif action == "Help" and obj == "me":
                dialog.AssistantSays(
                    Strings.help_str,
                    Strings.help_str_for_voice_over,
                    another_text_for_voice_over=True,
                )
                is_done = True

        # Cases of non-recognition of the imperative mood
        if not is_imperative or not is_done:
            # Search in the Internet
            # Quary: What\who is (it) ...
            if splitted_command[0] == "What" or splitted_command[0] == "Who":
                if (
                    splitted_command[1] == "is" and splitted_command[2] == "it"
                ) or splitted_command[1] == "is":
                    quary = " ".join(splitted_command)

                    SearchInTheInternet(dialog, screen, quary)
                    is_done = True
            # Shutdown PC
            elif splitted_command[0] == "Shutdown":
                if splitted_command[1] == "computer" and len(splitted_command) == 2:
                    system("shutdown /s")
                    is_done = True

            # Open programs
            elif not is_done and splitted_command[0] == "Open":
                app_name = ""
                for i in range(1, len(splitted_command)):
                    app_name += splitted_command[i]
                    # Add spaces between words
                    if i != len(splitted_command) - 1:
                        app_name += " "

                OpenProgram(dialog, app_name)
                is_done = True

            # Close programs
            elif not is_done and splitted_command[0] == "Close":
                app_name = ""
                for i in range(1, len(splitted_command)):
                    app_name += splitted_command[i]
                    # Add spaces between words
                    if i != len(splitted_command) - 1:
                        app_name += " "

                is_done = KillProgram(dialog, app_name)
                if not is_done:
                    dialog.AssistantSays(
                        "I can't find an open app with that name"
                    )

            # Multimedia control
            elif (
                not is_done
                and splitted_command[0] == "Next"
                and (splitted_command[1] == "music" or splitted_command[1] == "track")
            ):
                MultimediaControl(keyboard, splitted_command[0].lower())
                dialog.AssistantSays("Next audio")

                is_done = True
            elif (
                not is_done
                and splitted_command[0] == "Previously"
                and (splitted_command[1] == "music" or splitted_command[1] == "track")
            ):
                MultimediaControl(keyboard, splitted_command[0].lower())
                dialog.AssistantSays("Previous audio")

                is_done = True

        if not is_done:
            dialog.AssistantSays("Sorry, I don't understand")
    except Exception as e:
        print(traceback.format_exc())
        print(e)


def TerminateVoiceover():
    if common.voiceover_shared_list[0]:
        StopVoiceOver()


# Speech recognition
def CommandRecognition():
    if common.voiceover_shared_list[0]:
        StopVoiceOver()
    else:
        # Creates a new `Recognizer` instance
        r = sr.Recognizer()
        with sr.Microphone() as source:
            dialog.AssistantSays("Listening...")
            # Listening for speech
            audio = r.listen(source)
            # noinspection PyBroadException
            try:
                command = r.recognize_google(audio)
                CommandAnalysis(use_speech=True, command=command)
            except Exception as _:
                dialog.AssistantSays("Try Again")
