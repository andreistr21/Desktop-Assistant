from multiprocessing import Process
from os import system
from subprocess import run
from webbrowser import open
from keyboard import send
import speech_recognition as sr
from math import floor, ceil
import dearpygui.dearpygui as dpg

from resources import Strings
from bin.classes.MasterAudioController import MasterAudioController
from bin.common import Common as common
from bin.VoiceoverCallback import VoiceOver


audio_controller = MasterAudioController()
assistant_speech_rate = 150
process = None

all_replicas = []

one_line_max_pixels_text = 310
text_x_pos = 387


def GUIChanger(
    chat_window_width,
    chat_window_height,
    input_text_pos,
    input_text_width,
    image_btn_pos,
):
    dpg.set_item_width("Chat_window_id", chat_window_width)
    dpg.set_item_height("Chat_window_id", chat_window_height)

    dpg.set_item_pos("Text_input_id", input_text_pos)
    dpg.set_item_width("Text_input_id", input_text_width)

    dpg.set_item_pos("microphone_btn_id", image_btn_pos)


def ViewportResize():
    global one_line_max_pixels_text
    global text_x_pos

    viewport_height = dpg.get_viewport_height()
    viewport_width = dpg.get_viewport_width()

    chat_window_height = viewport_height - 100
    chat_window_width = viewport_width - 18
    input_text_width = viewport_width - 75
    input_text_pos = [10, viewport_height - 70]
    one_line_max_pixels_text = viewport_width - 120
    image_btn_pos = [viewport_width - 60, viewport_height - 85]
    text_x_pos = viewport_width - 43

    GUIChanger(
        chat_window_width,
        chat_window_height,
        input_text_pos,
        input_text_width,
        image_btn_pos,
    )

    dpg.delete_item("Chat_window_id", children_only=True)
    common.pixels_y = [10]

    for item in all_replicas:
        if item[0] == "Assistant":
            AssistantSays(
                item[1],
                common.pixels_y,
                voiceover=False,
                logs=False,
                auto_scroll_to_bottom=False,
            )
        elif item[0] == "User":
            UserSays(item[1], common.pixels_y, logs=False, auto_scroll_to_bottom=False)

    # Scroll to the bottom
    dpg.render_dearpygui_frame()
    # Scroll to the bottom of the window
    dpg.set_y_scroll("Chat_window_id", dpg.get_y_scroll_max("Chat_window_id"))


def SwitchKeyboardLanguage():
    # noinspection PyBroadException
    try:
        send("shift+alt")
        AssistantSays("Language is switched.", common.pixels_y)
    except:
        AssistantSays("Can't changed keyboard language.", common.pixels_y)


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
        app_list = run(
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

    if app_id is not None:
        # Open the program
        system(f"start explorer shell:appsfolder\\{app_id}")

        AssistantSays(f'Opening "{app_name}" ...', common.pixels_y)
    else:
        AssistantSays("No apps found with this name", common.pixels_y)


def ChangeVolume(volume_percents, change=False):
    # This function will set or change volume level
    # volume_percents should be in volume range for set and "+" or "-" for increase or decrease
    if not change:
        audio_controller.SetVolumeScalar(volume_percents)
    else:
        volume = audio_controller.GetMasterVolume()

        if volume_percents == "+":
            # Volume cannot be more than 1
            if volume + 0.2 > 1:
                audio_controller.SetVolumeScalar(1, True)
            else:
                audio_controller.SetVolumeScalar(volume + 0.2, True)

            AssistantSays("Volume increased.", common.pixels_y)
        elif volume_percents == "-":
            # Volume cannot be less than 0
            if volume - 0.2 < 0:
                audio_controller.SetVolumeScalar(0, True)
            else:
                audio_controller.SetVolumeScalar(volume - 0.2, True)

            AssistantSays("Volume decreased.", common.pixels_y)


def CommandAnalysisCall(sender, app_data):
    global process

    # noinspection PyUnresolvedReferences
    if process is not None and process.is_alive():
        # noinspection PyUnresolvedReferences
        process.terminate()

    CommandAnalysis(sender=sender, app_data=app_data)


def CommandAnalysis(sender="", app_data="", use_speech=False, command=""):
    # If use_speech True use command variable
    if not use_speech:
        dpg.set_value(sender, "")

        command = app_data
        # Focus text input item
        dpg.focus_item("Text_input_id")

    splitted_command = command.split(" ")

    # Replace the wrong name
    if splitted_command[0] == "Sergei" or splitted_command[0] == "sergei":
        splitted_command[0] = "Sergey"

    FirstLetterToUpperCase(splitted_command)

    # Assembling the command in one string
    command = " ".join(splitted_command)

    splitted_command = command.split(" ")

    UserSays(command, common.pixels_y)

    is_done = False

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
                    is_done = True

        # Search in the Internet
        # Quary: What\who is (it) ...
        if splitted_command[0] == "What" or splitted_command[0] == "Who":
            if (
                splitted_command[1] == "is" and splitted_command[2] == "it"
            ) or splitted_command[1] == "is":
                quary = " ".join(splitted_command)

                url = URLCreator(quary)
                open(url)
                is_done = True
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
                open(url)
                is_done = True
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
                open(url)
                is_done = True

        # Open programs
        if splitted_command[0] == "Open":
            app_name = ""
            for i in range(1, len(splitted_command)):
                app_name += splitted_command[i]
                # Add spaces between words
                if i != len(splitted_command) - 1:
                    app_name += " "

            OpenProgram(app_name)
            is_done = True

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
                global assistant_speech_rate
                assistant_speech_rate = volume_rate
                AssistantSays("Speech rate is changed.", common.pixels_y)
                is_done = True

        # Change PC volume
        if splitted_command[0] == "Set":
            if splitted_command[1] == "volume" and splitted_command[2] == "to":
                volume_percents = splitted_command[3]
                ChangeVolume(int(volume_percents))
                is_done = True
        elif splitted_command[0] == "Increase":
            if splitted_command[1] == "volume":
                ChangeVolume("+", True)
                is_done = True
        elif splitted_command[0] == "Decrease":
            if splitted_command[1] == "volume":
                ChangeVolume("-", True)
                is_done = True

        # Shutdown\Restart PC
        if splitted_command[0] == "Shutdown":
            if splitted_command[1] == "computer" and len(splitted_command) == 2:
                system("shutdown /s")
                is_done = True
            elif (
                splitted_command[1] == "computer"
                and splitted_command[2] == "immediately"
            ):
                system("shutdown /s /t 0")
                is_done = True
        elif splitted_command[0] == "Restart":
            if splitted_command[1] == "computer" and len(splitted_command) == 2:
                system("shutdown /r")
                is_done = True
            elif (
                splitted_command[1] == "computer"
                and splitted_command[2] == "immediately"
            ):
                system("shutdown /r /t 0")
                is_done = True

        # Help menu
        if splitted_command[0] == "Help":
            if splitted_command[1] == "me":
                AssistantSays(
                    Strings.help_str,
                    common.pixels_y,
                    Strings.help_str_for_voice_over,
                    another_text_for_voice_over=True,
                )
                is_done = True

        if splitted_command[0] == "Test":
            test()
        if not is_done:
            AssistantSays("Sorry, I don't understand.", common.pixels_y)
    except Exception as e:
        print(e)


def test():
    print(dpg.get_viewport_height())
    print(dpg.get_viewport_width())


def TextDivisionIntoLines(text):
    global one_line_max_pixels_text

    max_letters_on_one_line = floor(one_line_max_pixels_text / 7)

    words_in_line_counter = 0
    word_start = 0
    # max_width = 0
    letter_index = 0
    new_lines_counter = 0

    # Create new line if there is max number of letters
    while letter_index < len(text):
        if words_in_line_counter == max_letters_on_one_line:
            if text[letter_index] != " ":
                if text[letter_index - 1] != " " or text[letter_index + 1] != " ":
                    # Search for start of word
                    for j in range(letter_index, 0, -1):
                        if text[j] != " ":
                            word_start = j
                        elif text[j] == " ":
                            break

            if word_start != 0:
                letter_index = word_start
                word_start = 0
            text = f"{text[:letter_index]}\n{text[letter_index:]}"
            new_lines_counter += 1
            words_in_line_counter = 0

        if text[letter_index] == "\n":
            words_in_line_counter = 0
        else:
            words_in_line_counter += 1

        letter_index += 1

    return text, new_lines_counter


def NewLinesCounter(text):
    counter = 0

    for letter in text:
        if letter == "\n":
            counter += 1

    return counter


def AssistantSays(
    text,
    pixels,
    voice_over_text="",
    another_text_for_voice_over=False,
    voiceover=True,
    logs=True,
    auto_scroll_to_bottom=True,
):
    global process
    global one_line_max_pixels_text

    pre_edit_text = text
    if logs:
        all_replicas.append(["Assistant", text])

    if not another_text_for_voice_over:
        voice_over_text = pre_edit_text

    if voiceover:
        # Start voiceover in background
        process = Process(
            target=VoiceOver, args=(voice_over_text, assistant_speech_rate)
        )
        process.start()

    text_len = len(text)
    text_len_pixels = (
        text_len * 7
    )  # 7 pixels for one letter, in one line max 310 pixels

    if text_len_pixels <= one_line_max_pixels_text:
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

        number_of_lines = NewLinesCounter(text)

        text, lines_to_add = TextDivisionIntoLines(text)

        number_of_lines += lines_to_add + 1

        dpg.add_text(text, parent="Chat_window_id", pos=[15, pixels[0] + 10])

        with dpg.drawlist(
            width=one_line_max_pixels_text + 14,
            height=14 * number_of_lines + 15,
            parent="Chat_window_id",
            pos=[9, pixels[0] + 6],
        ):
            dpg.draw_rectangle(
                pmin=[0, 0],
                pmax=[one_line_max_pixels_text + 14, 14 * number_of_lines + 15],
                rounding=10,
            )

        pixels[0] += 14 * number_of_lines + 15 + 10

    if auto_scroll_to_bottom:
        dpg.render_dearpygui_frame()
        # Scroll to the bottom of the window
        dpg.set_y_scroll("Chat_window_id", dpg.get_y_scroll_max("Chat_window_id"))


def UserSays(text, pixels, logs=True, auto_scroll_to_bottom=True):
    global one_line_max_pixels_text
    global text_x_pos

    if logs:
        all_replicas.append(["User", text])

    text_len = len(text)
    text_len_pixels = (
        text_len * 7
    )  # 7 pixels for one letter, in one line max 310 pixels

    if text_len_pixels <= one_line_max_pixels_text:
        dpg.add_text(
            text,
            parent="Chat_window_id",
            pos=[text_x_pos - text_len_pixels, pixels[0] + 10],
        )

        with dpg.drawlist(
            width=text_len_pixels + 14,
            height=30,
            parent="Chat_window_id",
            pos=[text_x_pos + 6 - text_len_pixels - 12, pixels[0] + 6],
        ):
            dpg.draw_rectangle(
                pmin=[0, 0], pmax=[text_len_pixels + 14, 30], rounding=10
            )

        pixels[0] += 40
    else:
        number_of_lines = ceil(text_len_pixels / one_line_max_pixels_text)
        number_of_lines += NewLinesCounter(text)

        text, lines_to_add = TextDivisionIntoLines(text)

        number_of_lines += lines_to_add

        dpg.add_text(
            text,
            parent="Chat_window_id",
            pos=[text_x_pos - one_line_max_pixels_text, pixels[0] + 10],
        )

        with dpg.drawlist(
            width=one_line_max_pixels_text + 14,
            height=14 * number_of_lines + 15,
            parent="Chat_window_id",
            pos=[text_x_pos + 6 - one_line_max_pixels_text - 12, pixels[0] + 6],
        ):
            dpg.draw_rectangle(
                pmin=[0, 0],
                pmax=[one_line_max_pixels_text + 14, 14 * number_of_lines + 15],
                rounding=10,
            )

        pixels[0] += 14 * number_of_lines + 15 + 10

    if auto_scroll_to_bottom:
        dpg.render_dearpygui_frame()
        # Scroll to the bottom of the window
        dpg.set_y_scroll("Chat_window_id", dpg.get_y_scroll_max("Chat_window_id"))


def TerminateVoiceover():
    global process

    # noinspection PyUnresolvedReferences
    if process is not None and process.is_alive():
        # noinspection PyUnresolvedReferences
        process.terminate()


# Speech recognition
def CommandRecognition():
    global process

    # noinspection PyUnresolvedReferences
    if process is not None and process.is_alive():
        # noinspection PyUnresolvedReferences
        process.terminate()
    else:
        # Creates a new `Recognizer` instance
        r = sr.Recognizer()
        with sr.Microphone() as source:
            AssistantSays("Listening...", common.pixels_y)
            # Listening for speech
            audio = r.listen(source)
            # noinspection PyBroadException
            try:
                command = r.recognize_google(audio)
                CommandAnalysis(use_speech=True, command=command)
            except:
                AssistantSays("Try Again.", common.pixels_y)
