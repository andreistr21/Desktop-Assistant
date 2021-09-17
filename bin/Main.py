import os
import subprocess
import threading
import webbrowser
import keyboard
import pyttsx3
import speech_recognition as sr
import math
import dearpygui.dearpygui as dpg

from resources import Strings
from bin.classes.MasterAudioController import MasterAudioController
from bin.common import Common as common
from bin.classes.Voice import Voice

# # Constructs a new TTS engine instance
# engine = pyttsx3.init()
# # noinspection PyUnresolvedReferences
# # Set voice type
# engine.setProperty("voice", engine.getProperty("voices")[1].id)
# # Set speed of speech (words per minute)
# engine.setProperty("rate", 150)
#
# audio_controller = MasterAudioController()

voice = Voice()
audio_controller = MasterAudioController()


def SwitchKeyboardLanguage():
    # noinspection PyBroadException
    try:
        keyboard.send("shift+alt")
        AssistantSays("Language is switched", common.pixels_y)
    except:
        AssistantSays("Can't changed keyboard language", common.pixels_y)


def URLCreator(quary):
    temp = "https://www.google.com/search?q="
    quary = quary.replace(" ", "+")
    url = f"{temp}{quary}"

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


def ChooseAppFromList(dictionary, app_name):
    # print(dictionary)
    temp_dict = dictionary.copy()
    if 1 < len(dictionary) < 4:
        # Delete app if this is web site
        for name in dictionary:
            if temp_dict[name].find("http://") != -1:
                del temp_dict[name]

        if len(temp_dict) == 0:
            # noinspection PyUnusedLocal
            temp_dict = dictionary.copy()
        elif len(temp_dict) == 1:
            return list(temp_dict.values())[0]
        elif len(temp_dict) > 1:
            if (
                len(app_name.split()) == 1
            ):  # Check the number of words in the application name
                shortest_name = ""
                size = 10000
                for name in temp_dict:
                    new_size = len(temp_dict[name])
                    if new_size < size:
                        size = new_size
                        shortest_name = name

                return dictionary[shortest_name]
            else:
                return list(dictionary.values())[0]

    elif len(dictionary) == 1:
        return list(dictionary.values())[0]


# Parsing output of powershell to dictionary
def PowerShellOutputParsing(string, app_name):
    apps_dictionary = {}
    name_pos = -8
    app_id_pos = -8
    while name_pos != -1:
        # noinspection PyUnusedLocal
        name = ""
        # noinspection PyUnusedLocal
        app_id = ""
        name_pos = string.find("Name  : ", name_pos + 8)

        if name_pos != -1:
            app_id_pos = string.find("AppID : ", app_id_pos + 8)
            next_name_pos = string.find("Name  : ", name_pos + 8)
            name = string[name_pos + 8 : app_id_pos]
            if next_name_pos != -1:
                app_id = string[app_id_pos + 8 : next_name_pos]
            else:
                app_id = string[app_id_pos + 8 : len(string)]

            # Replace redundant chars in name
            while (
                name.find("\n") != -1 or name.find("\r") != -1 or name.find("  ") != -1
            ):
                name = name.replace("\n", "")
                name = name.replace("\r", "")
                name = name.replace("  ", " ")
            # Replace redundant chars in app_id
            while (
                app_id.find("\n") != -1
                or app_id.find("\r") != -1
                or app_id.find("  ") != -1
            ):
                app_id = app_id.replace("\n", "")
                app_id = app_id.replace("\r", "")
                app_id = app_id.replace("  ", " ")

            apps_dictionary[name] = app_id

    app_id = ChooseAppFromList(apps_dictionary, app_name)

    return app_id


def OpenProgram(app_name):
    app_id = None
    # noinspection PyBroadException
    try:
        # Get all installed apps and theirs IDs in PC
        # fmt: off
        app_list = subprocess.run(
            [
                "powershell",
                "-Command",
                "get-StartApps | Where-Object { $_.Name -like '*" + app_name + "*' } | Format-List",
            ],
            capture_output=True,
        ).stdout.decode()
        # fmt: on

        # Return only app id
        app_id = PowerShellOutputParsing(app_list, app_name)
    except Exception as e:
        print(e)

    # Open the program
    os.system(f"start explorer shell:appsfolder\\{app_id}")

    AssistantSays(f'Opening "{app_name}" ...', common.pixels_y)


# def ChangeAssistantVolume(volume_rate):
#     engine.setProperty("rate", volume_rate)


def ChangeVolume(volume_percents, change=False):
    # This function will set or change volume level
    # volume_percents should be in volume range for set and "+" or "-" for increase or decrease
    if not change:
        audio_controller.SetVolumeScalar(volume_percents)
    else:
        volume = audio_controller.GetMasterVolume()
        half_volume = volume / 2
        if volume_percents == "+":
            audio_controller.SetVolumeScalar(volume * 2, True)
        elif volume_percents == "-":
            audio_controller.SetVolumeScalar(volume - half_volume, True)


def CommandAnalysisCall(sender, app_data):
    CommandAnalysis(sender=sender, app_data=app_data)


def CommandAnalysis(sender="", app_data="", use_command=False, command=""):
    if not use_command:
        dpg.set_value(sender, "")

        command = app_data

    splitted_command = command.split(" ")

    # Replace the wrong name
    if splitted_command[0] == "Sergei" or splitted_command[0] == "sergei":
        splitted_command[0] = "Sergey"

    FirstLetterToUpperCase(splitted_command)

    # Create the right command for print to user
    command = " ".join(splitted_command)
    UserSays(command, common.pixels_y)

    # noinspection PyBroadException
    try:
        # Change keyboard language command
        if splitted_command[0] == Strings.name_of_assistant:
            if splitted_command[1] == "switch" or splitted_command[1] == "change":
                if (
                    splitted_command[2] == "keyboard"
                    and splitted_command[3] == "language"
                ) or splitted_command[2] == "language":
                    SwitchKeyboardLanguage()

        # Search in the Internet
        # Quary: What\who is (it) ...
        if splitted_command[0] == "What" or splitted_command[0] == "Who":
            if (
                splitted_command[1] == "is" and splitted_command[2] == "it"
            ) or splitted_command[1] == "is":
                quary = " ".join(splitted_command)

                url = URLCreator(quary)
                webbrowser.open(url)
        # Quary: Search\Find about\for ...
        elif splitted_command[0] == "Search" or splitted_command[0] == " Find":
            if splitted_command[1] == "about" or splitted_command[1] == "for":
                quary = ""
                # Create right quary
                for i in range(2, len(splitted_command)):
                    quary += splitted_command[i]
                    # Add space between words
                    if i != len(splitted_command) - 1:
                        quary += " "

                url = URLCreator(quary)
                webbrowser.open(url)
            # Quary: Search ...
            else:
                quary = ""
                # Create right quary
                for i in range(1, len(splitted_command)):
                    quary += splitted_command[i]
                    # Add spaces between words
                    if i != len(splitted_command) - 1:
                        quary += " "

                url = URLCreator(quary)
                webbrowser.open(url)

        # Open programs
        if splitted_command[0] == "Open":
            app_name = ""
            for i in range(1, len(splitted_command)):
                app_name += splitted_command[i]
                # Add spaces between words
                if i != len(splitted_command) - 1:
                    app_name += " "

            OpenProgram(app_name)

        # Change assistant speech rate
        if splitted_command[0] == "Change":
            volume_rate = None
            if (
                splitted_command[1] == "rate"
                and splitted_command[2] == "of"
                and splitted_command[3] == "assistant"
                and splitted_command[4] == "speech"
                and splitted_command[5] == "to"
            ):
                volume_rate = int(splitted_command[6])
            elif (
                splitted_command[1] == "the"
                and splitted_command[2] == "rate"
                and splitted_command[3] == "of"
                and splitted_command[4] == "assistant"
                and splitted_command[5] == "speech"
                and splitted_command[6] == "to"
            ):
                volume_rate = int(splitted_command[7])

            elif (
                splitted_command[1] == "the"
                and splitted_command[2] == "assistant"
                and splitted_command[3] == "speech"
                and splitted_command[4] == "rate"
                and splitted_command[5] == "to"
            ):
                volume_rate = int(splitted_command[6])

            elif (
                splitted_command[1] == "assistant"
                and splitted_command[2] == "rate"
                and splitted_command[3] == "of"
                and splitted_command[4] == "speech"
                and splitted_command[5] == "to"
            ):
                volume_rate = int(splitted_command[6])

            if volume_rate is not None:
                voice.ChangeAssistantVolume(volume_rate)
                AssistantSays("Speech rate is changed", common.pixels_y)

        # Change PC volume
        if splitted_command[0] == "Set":
            if splitted_command[1] == "volume" and splitted_command[2] == "to":
                volume_percents = splitted_command[3]
                ChangeVolume(int(volume_percents))
        if splitted_command[0] == "Increase":
            if splitted_command[1] == "volume":
                ChangeVolume("+", True)
        if splitted_command[0] == "Decrease":
            if splitted_command[1] == "volume":
                ChangeVolume("-", True)

        # Shutdown\Restart PC
        if splitted_command[0] == "Shutdown":
            if splitted_command[1] == "computer" and len(splitted_command) == 2:
                os.system("shutdown /s")
            elif (
                splitted_command[1] == "computer"
                and splitted_command[2] == "immediately"
            ):
                os.system("shutdown /s /t 0")
        if splitted_command[0] == "Restart":
            if splitted_command[1] == "computer" and len(splitted_command) == 2:
                os.system("shutdown /r")
            elif (
                splitted_command[1] == "computer"
                and splitted_command[2] == "immediately"
            ):
                os.system("shutdown /r /t 0")

        # Help menu
        if splitted_command[0] == "Help":
            if splitted_command[1] == "me":
                AssistantSays(Strings.help_str, common.pixels_y)
    except Exception as e:
        print(e)


def TextDivisionIntoLines(text):
    max_letters_on_one_line = math.floor(common.one_line_max_pixels_text / 7)

    counter = 0
    for i in range(len(text)):
        if counter == max_letters_on_one_line:
            text = f"{text[:i]}\n{text[i:]}"
            counter = 0

        counter += 1

    return text


def NewLinesCounter(text):
    counter = 0

    for letter in text:
        if letter == "\n":
            counter += 1

    return counter


def AssistantSays(text, pixels):
    pre_edit_text = text

    text_len = len(text)
    text_len_pixels = (
        text_len * 7
    )  # 7 pixels for one letter, in one line max 310 pixels

    if text_len_pixels <= common.one_line_max_pixels_text:
        dpg.add_text(text, parent="Chat_window_id", pos=[15, pixels[0] + 10])

        with dpg.drawlist(
            width=text_len_pixels + 14,
            height=30,
            parent="Chat_window_id",
            pos=[9, pixels[0] + 6],
        ):
            dpg.draw_rectangle(
                pmin=[0, 0], pmax=[text_len_pixels + 14, 30], rounding=10
            )

        pixels[0] += 40

    else:
        number_of_lines = math.ceil(text_len_pixels / common.one_line_max_pixels_text)
        number_of_lines += NewLinesCounter(text)

        text = TextDivisionIntoLines(text)

        dpg.add_text(text, parent="Chat_window_id", pos=[15, pixels[0] + 10])

        with dpg.drawlist(
            width=common.one_line_max_pixels_text + 14,
            height=14 * number_of_lines + 15,
            parent="Chat_window_id",
            pos=[9, pixels[0] + 6],
        ):
            dpg.draw_rectangle(
                pmin=[0, 0],
                pmax=[common.one_line_max_pixels_text + 14, 14 * number_of_lines + 15],
                rounding=10,
            )

        pixels[0] += 14 * number_of_lines + 15 + 10

    voice.Speech(pre_edit_text)


def UserSays(text, pixels):
    text_len = len(text)
    text_len_pixels = (
        text_len * 7
    )  # 7 pixels for one letter, in one line max 310 pixels

    if text_len_pixels <= common.one_line_max_pixels_text:
        dpg.add_text(
            text, parent="Chat_window_id", pos=[387 - text_len_pixels, pixels[0] + 10]
        )

        with dpg.drawlist(
            width=text_len_pixels + 14,
            height=30,
            parent="Chat_window_id",
            pos=[393 - text_len_pixels - 12, pixels[0] + 6],
        ):
            dpg.draw_rectangle(
                pmin=[0, 0], pmax=[text_len_pixels + 14, 30], rounding=10
            )

        pixels[0] += 40
    else:
        number_of_lines = math.ceil(text_len_pixels / common.one_line_max_pixels_text)
        number_of_lines += NewLinesCounter(text)

        text = TextDivisionIntoLines(text)

        dpg.add_text(
            text,
            parent="Chat_window_id",
            pos=[387 - common.one_line_max_pixels_text, pixels[0] + 10],
        )

        with dpg.drawlist(
            width=common.one_line_max_pixels_text + 14,
            height=14 * number_of_lines + 15,
            parent="Chat_window_id",
            pos=[393 - common.one_line_max_pixels_text - 12, pixels[0] + 6],
        ):
            dpg.draw_rectangle(
                pmin=[0, 0],
                pmax=[common.one_line_max_pixels_text + 14, 14 * number_of_lines + 15],
                rounding=10,
            )

        pixels[0] += 14 * number_of_lines + 15 + 10


# def VoiceoverThread(text):
#     engine.say(text)
#     engine.runAndWait()
#
#
# # Voiceover a command
# def Speak(audio):
#     # Adds an utterance to speak to the queue
#     engine.say(audio)
#     engine.runAndWait()
#     # voiceover = threading.Thread(target=VoiceoverThread, args=(audio,))
#     # voiceover.start()


# Speech recognition
def CommandRecognition():
    # Creates a new `Recognizer` instance
    r = sr.Recognizer()
    with sr.Microphone() as source:
        AssistantSays("Listening...", common.pixels_y)
        # Listening for speech
        audio = r.listen(source)
        # noinspection PyBroadException
        try:
            command = r.recognize_google(audio)
            CommandAnalysis(use_command=True, command=command)
        except:
            AssistantSays("Try Again", common.pixels_y)

#
# def Main():
#     # AssistantSays(strings.welcome_str)
#     # command = CommandRecognition()
#     # command = input()
#     # command = "help me"
#     # CommandAnalysis(command)
#     pass
