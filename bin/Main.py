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

from bin.classes.Dialog import Dialog
from resources import Strings
from bin.classes.MasterAudioController import MasterAudioController
from bin.common import Common as common
from bin.CommandAnalyser import GetActionAndObject


nlp = load("en_core_web_trf")

keyboard = Controller()

audio_controller = MasterAudioController()

dialog = Dialog()


def AssistantSaysMethCall(text: str):
    """AssistantSays Methode call from class

    :param text: str
        Text to say by assistant
    :return: None
    """
    dialog.AssistantSays(text)


def StopVoiceOver():
    """Stops the voiceover process"""
    # Set is_speaking_now flag to true
    common.voiceover_shared_list[0] = False


def GUIChanger(
    chat_window_width,
    chat_window_height,
    input_text_pos,
    input_text_width,
    image_btn_pos,
):
    """Change resolution of each GUI element in the window."""

    dpg.set_item_width("Chat_window_id", chat_window_width)
    dpg.set_item_height("Chat_window_id", chat_window_height)

    dpg.set_item_pos("Text_input_id", input_text_pos)
    dpg.set_item_width("Text_input_id", input_text_width)

    dpg.set_item_pos("microphone_btn_id", image_btn_pos)


def ViewportResize():
    """Window resize handler. Restores dialog."""
    viewport_height = dpg.get_viewport_height()
    viewport_width = dpg.get_viewport_width()

    chat_window_height = viewport_height - 100
    chat_window_width = viewport_width - 18
    input_text_width = viewport_width - 75
    input_text_pos = [10, viewport_height - 70]
    dialog.one_line_max_pixels_text = viewport_width - 120
    image_btn_pos = [viewport_width - 60, viewport_height - 85]
    dialog.text_x_pos = viewport_width - 43

    GUIChanger(
        chat_window_width,
        chat_window_height,
        input_text_pos,
        input_text_width,
        image_btn_pos,
    )

    dpg.delete_item("Chat_window_id", children_only=True)
    dialog.pixels_y = 10

    for item in dialog.all_replicas:
        if item[0] == "Assistant":
            dialog.AssistantSays(
                item[1],
                voiceover=False,
                logs=False,
                auto_scroll_to_bottom=False,
            )
        elif item[0] == "User":
            dialog.UserSays(item[1], logs=False, auto_scroll_to_bottom=False)
        elif item[0] == "Button":
            ButtonCreate(item[1], logs=False)

    # Scroll to the bottom of the window
    dpg.set_y_scroll("Chat_window_id", dpg.get_y_scroll_max("Chat_window_id"))


def SwitchKeyboardLanguage():
    """Change keyboard language"""

    # noinspection PyBroadException
    try:
        send("shift+alt")
        dialog.AssistantSays("Language is switched")
    except Exception as _:
        dialog.AssistantSays("Can't changed keyboard language")


def QuaryCreator(quary: str) -> tuple:
    """Create suitable request
    Args:
        quary (String):
    Returns:
        tuple(str, str)
    """

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.89 Safari/537.36"
    }
    temp = "https://www.google.com/search?q="
    quary = quary.replace(" ", "+")
    url = f"{temp}{quary}&hl=en"

    return url, headers


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


def PowerShellOutputParsing(string, app_name):
    """Parsing output of powershell to dictionary"""

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

        dialog.AssistantSays(f'Opening "{app_name}" ...')
    else:
        dialog.AssistantSays("No apps found with this name")


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

            dialog.AssistantSays("Volume increased")
        elif volume_percents == "-":
            # Volume cannot be less than 0
            if volume - 0.2 < 0:
                audio_controller.SetVolumeScalar(0, True)
            else:
                audio_controller.SetVolumeScalar(volume - 0.2, True)

            dialog.AssistantSays("Volume decreased")


def CommandAnalysisCall(sender, app_data):
    # Stop voiceover process if it running now
    if common.voiceover_shared_list[0]:
        StopVoiceOver()

    CommandAnalysis(sender=sender, app_data=app_data)


def OpenInABrowserButtonClicked(_, __, url):
    open(url)
    TerminateVoiceover()


def ButtonCreate(
    url: str,
    logs=True,
):
    """Create button under the text
    Args:
        url (str):
        logs (bool): if True, remembers the button as created (new)
    """
    dpg.add_button(
        label="Open in a browser",
        width=dialog.one_line_max_pixels_text + 14,
        height=25,
        parent="Chat_window_id",
        callback=OpenInABrowserButtonClicked,
        user_data=url,
        pos=[9, dialog.pixels_y + 6],
    )

    dialog.pixels_y += 35

    if logs:
        dialog.all_replicas.append(["Button", url])

        # Update interface (render one frame)
        dpg.render_dearpygui_frame()
        # Scroll to the bottom of the window
        dpg.set_y_scroll("Chat_window_id", dpg.get_y_scroll_max("Chat_window_id"))


def SearchInTheInternet(quary):
    """Internet search for a given query
    Args:
        quary (String)
    Returns:
        None
    """
    url, headers = QuaryCreator(quary)
    request = Request(url, headers=headers)
    response = urlopen(request)

    soup = BeautifulSoup(response, "html.parser")
    if soup.find(class_="hgKElc"):
        dialog.AssistantSays(soup.find(class_="hgKElc").text)
        ButtonCreate(url)
    elif soup.find(class_="kno-rdesc"):
        dialog.AssistantSays(soup.find(class_="kno-rdesc").text[11:-10])
        ButtonCreate(url)
    else:
        dialog.AssistantSays("Opening in a browser...")
        open(url)


def GetActualProgramName(name: str):
    """
    Get actual program name from all opened programs

    :param name: str
        The approximate name of the program or process
    :returns: str, None
        Actual name of the program
    """

    # Get list of all processes
    programs_list = popen("wmic process get description, processid").read()

    # print(programs_list)

    # Find all speeches expressions
    programs_name = findall(fr"{name}.*", programs_list, flags=IGNORECASE)

    # Find a better match if necessary and kill
    if len(programs_name) > 0:
        return findall(r"^\S*\b", programs_name[0])[0]
    else:
        first_part = ""
        second_part = ""
        splitted_name = name.split()

        print(splitted_name)

        if len(splitted_name) == 2:
            print(f"splitted name: {splitted_name}")
            first_part = splitted_name[0]
            second_part = splitted_name[1]
            print(first_part)
            print(second_part)
        elif len(splitted_name) == 3:
            first_part = splitted_name[0:1]
            second_part = splitted_name[2]
        elif len(splitted_name) == 4:
            first_part = splitted_name[0:1]
            second_part = splitted_name[2:3]

        print(f"second part: {second_part}")

        first_part_names = findall(fr"{first_part}.*", programs_list, flags=IGNORECASE)
        second_part_names = findall(
            fr"{second_part}.*", programs_list, flags=IGNORECASE
        )

        print()
        # print(second_part_names)

        if len(first_part_names) > 0 and len(second_part_names) == 0:
            return findall(r"^\S*\b", first_part_names[0])[0]
        elif len(first_part_names) == 0 and len(second_part_names) > 0:
            return findall(r"^\S*\b", second_part_names[0])[0]
        elif len(first_part_names) > 0 and len(second_part_names) > 0:
            # Get names of programs
            first_app_name = findall(r"^\S*\b", first_part_names[0])[0]
            second_app_name = findall(r"^\S*\b", second_part_names[0])[0]

            # Split names by uppercase
            first_app_name_splitted = findall(r"[A-Z][^A-Z]*", first_app_name)
            second_app_name_splitted = findall(r"[A-Z][^A-Z]*", second_app_name)

            # Replace if str is empty after split (does not contain any uppercase letters)
            if len(first_app_name_splitted) == 0:
                first_app_name_splitted = [first_app_name]
            if len(second_app_name_splitted) == 0:
                second_app_name_splitted = [second_app_name]

            desired_len = len(splitted_name)
            first_app_name_len = len(first_app_name_splitted)
            second_app_name_len = len(second_app_name_splitted)

            # Choose best option
            if abs(desired_len - first_app_name_len) > abs(
                desired_len - second_app_name_len
            ):
                return second_app_name
            elif abs(desired_len - first_app_name_len) < abs(
                desired_len - second_app_name_len
            ):
                return first_app_name
            elif abs(desired_len - first_app_name_len) == abs(
                desired_len - second_app_name_len
            ):
                if first_app_name_len > second_app_name_len:
                    return second_app_name
                elif first_app_name_len < second_app_name_len:
                    return first_app_name
    return None


def KillProgram(name: str) -> bool:
    """
    Kill program by the name

    :param name: str
        The approximate name of the program or process
    :return: bool
        True if successful
    """

    actual_program_name = GetActualProgramName(name)

    print(actual_program_name)

    if actual_program_name is not None:
        try:
            system(f"TASKKILL /IM {actual_program_name} /F")

            dialog.AssistantSays("App is closed")

            return True
        except Exception as e:
            print(e)
    else:
        return False


def MultimediaControl(action):
    if action == "next":
        keyboard.press(Key.media_next)
    elif action == "previously":
        keyboard.press(Key.media_previous)
    elif action == "stop" or action == "play":
        keyboard.press(Key.media_play_pause)


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

    # # Replace the wrong name
    # if splitted_command[0] == "Sergei" or splitted_command[0] == "sergei":
    #     splitted_command[0] = "Sergey"

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
                SwitchKeyboardLanguage()
                is_done = True

            # Search in the Internet
            # Quary: Search\Find about\for ...
            elif action == "Search" or action == "Find":
                SearchInTheInternet(obj)
                is_done = True

            # Open programs
            elif action == "Open" and obj != "you":
                OpenProgram(obj)
                is_done = True

            # Close programs
            elif action == "Close" and obj != "you":
                is_done = KillProgram(obj)

            # Change PC volume
            if action == "Set" and obj == "volume":
                volume_percents = None
                for token in doc:
                    if token.pos_ == "NUM":
                        volume_percents = int(token.text)

                if volume_percents is not None:
                    ChangeVolume(int(volume_percents))
                    is_done = True
            elif action == "Increase" and obj == "volume":
                ChangeVolume("+", True)
                is_done = True
            elif action == "Decrease" and obj == "volume":
                ChangeVolume("-", True)
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
                MultimediaControl(action.lower())
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

                    SearchInTheInternet(quary)
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

                OpenProgram(app_name)
                is_done = True

            # Close programs
            elif not is_done and splitted_command[0] == "Close":
                app_name = ""
                for i in range(1, len(splitted_command)):
                    app_name += splitted_command[i]
                    # Add spaces between words
                    if i != len(splitted_command) - 1:
                        app_name += " "

                is_done = KillProgram(app_name)
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
                MultimediaControl(splitted_command[0].lower())
                dialog.AssistantSays("Next audio")

                is_done = True
            elif (
                not is_done
                and splitted_command[0] == "Previously"
                and (splitted_command[1] == "music" or splitted_command[1] == "track")
            ):
                MultimediaControl(splitted_command[0].lower())
                dialog.AssistantSays("Previous audio")

                is_done = True

        if not is_done:
            dialog.AssistantSays("Sorry, I don't understand")
    except Exception as e:
        print(traceback.format_exc())
        print(e)


# def TextDivisionIntoLines(text):
#     global one_line_max_pixels_text
# 
#     max_letters_on_one_line = floor(one_line_max_pixels_text / 7)
# 
#     words_in_line_counter = 0
#     word_start = 0
#     letter_index = 0
#     new_lines_counter = 0
# 
#     # Create new line if there is max number of letters
#     while letter_index < len(text):
#         if words_in_line_counter == max_letters_on_one_line:
#             if text[letter_index] != " ":
#                 if text[letter_index - 1] != " " or text[letter_index + 1] != " ":
#                     # Search for start of word
#                     for j in range(letter_index, 0, -1):
#                         if text[j] != " ":
#                             word_start = j
#                         elif text[j] == " ":
#                             break
# 
#             if word_start != 0:
#                 letter_index = word_start
#                 word_start = 0
#             text = f"{text[:letter_index]}\n{text[letter_index:]}"
#             new_lines_counter += 1
#             words_in_line_counter = 0
# 
#         if text[letter_index] == "\n":
#             words_in_line_counter = 0
#         else:
#             words_in_line_counter += 1
# 
#         letter_index += 1
# 
#     return text, new_lines_counter


# def NewLinesCounter(text):
#     counter = 0
# 
#     for letter in text:
#         if letter == "\n":
#             counter += 1
# 
#     return counter


# def CreateMP3File(text: str):
#     """
#     Create mp3 file for voiceover and save it in current directory.
# 
#     :param text: str
#         Text for voiceover
#     :return: None
#     """
#     start_time = time.process_time()
#     tts = gTTS(text, lang="en")
#     print(f"Voiceover received: {time.process_time() - start_time}")
# 
#     try:
#         start_time = time.process_time()
#         tts.save(f"resources/sounds/voiceover{common.voiceover_shared_list[1]}.mp3")
#         print(f"File saved: {time.process_time() - start_time}")
#         common.voiceover_shared_list[1] += 1
#     except PermissionError as _:
#         print("Can't save file: permission denied")


# def dialog.AssistantSays(
#     text,
#     pixels,
#     voice_over_text="",
#     another_text_for_voice_over=False,
#     voiceover=True,
#     logs=True,
#     auto_scroll_to_bottom=True,
# ):
#     global one_line_max_pixels_text
# 
#     pre_edit_text = text
#     if logs:
#         dialog.all_replicas.append(["Assistant", text])
# 
#     if not another_text_for_voice_over:
#         voice_over_text = pre_edit_text
# 
#     if voiceover:
#         CreateMP3File(voice_over_text)
#         common.voiceover_shared_list[0] = True
# 
#     text_len = len(text)
#     text_len_pixels = (
#         text_len * 7
#     )  # 7 pixels for one letter, in one line max 310 pixels
# 
#     if text_len_pixels <= one_line_max_pixels_text:
#         dpg.add_text(text, parent="Chat_window_id", pos=[15, pixels[0] + 10])
# 
#         with dpg.drawlist(
#             width=text_len_pixels + 14,
#             height=30,
#             parent="Chat_window_id",
#             pos=[9, pixels[0] + 6],
#         ):
#             dpg.draw_rectangle(
#                 pmin=[0, 0], pmax=[text_len_pixels + 14, 30], rounding=10
#             )
# 
#         pixels[0] += 40
# 
#     else:
# 
#         number_of_lines = NewLinesCounter(text)
# 
#         text, lines_to_add = TextDivisionIntoLines(text)
# 
#         number_of_lines += lines_to_add + 1
# 
#         dpg.add_text(text, parent="Chat_window_id", pos=[15, pixels[0] + 10])
# 
#         with dpg.drawlist(
#             width=one_line_max_pixels_text + 14,
#             height=14 * number_of_lines + 15,
#             parent="Chat_window_id",
#             pos=[9, pixels[0] + 6],
#         ):
#             dpg.draw_rectangle(
#                 pmin=[0, 0],
#                 pmax=[one_line_max_pixels_text + 14, 14 * number_of_lines + 15],
#                 rounding=10,
#             )
# 
#         pixels[0] += 14 * number_of_lines + 15 + 10
# 
#     if auto_scroll_to_bottom:
#         dpg.render_dearpygui_frame()
#         # Scroll to the bottom of the window
#         dpg.set_y_scroll("Chat_window_id", dpg.get_y_scroll_max("Chat_window_id"))


# def dialog.UserSays(text, pixels, logs=True, auto_scroll_to_bottom=True):
#     global one_line_max_pixels_text
#     global text_x_pos
# 
#     if logs:
#         dialog.all_replicas.append(["User", text])
# 
#     text_len = len(text)
#     text_len_pixels = (
#         text_len * 7
#     )  # 7 pixels for one letter, in one line max 310 pixels
# 
#     if text_len_pixels <= one_line_max_pixels_text:
#         dpg.add_text(
#             text,
#             parent="Chat_window_id",
#             pos=[text_x_pos - text_len_pixels, pixels[0] + 10],
#         )
# 
#         with dpg.drawlist(
#             width=text_len_pixels + 14,
#             height=30,
#             parent="Chat_window_id",
#             pos=[text_x_pos + 6 - text_len_pixels - 12, pixels[0] + 6],
#         ):
#             dpg.draw_rectangle(
#                 pmin=[0, 0], pmax=[text_len_pixels + 14, 30], rounding=10
#             )
# 
#         pixels[0] += 40
#     else:
#         number_of_lines = ceil(text_len_pixels / one_line_max_pixels_text)
#         number_of_lines += NewLinesCounter(text)
# 
#         text, lines_to_add = TextDivisionIntoLines(text)
# 
#         number_of_lines += lines_to_add
# 
#         dpg.add_text(
#             text,
#             parent="Chat_window_id",
#             pos=[text_x_pos - one_line_max_pixels_text, pixels[0] + 10],
#         )
# 
#         with dpg.drawlist(
#             width=one_line_max_pixels_text + 14,
#             height=14 * number_of_lines + 15,
#             parent="Chat_window_id",
#             pos=[text_x_pos + 6 - one_line_max_pixels_text - 12, pixels[0] + 6],
#         ):
#             dpg.draw_rectangle(
#                 pmin=[0, 0],
#                 pmax=[one_line_max_pixels_text + 14, 14 * number_of_lines + 15],
#                 rounding=10,
#             )
# 
#         pixels[0] += 14 * number_of_lines + 15 + 10
# 
#     if auto_scroll_to_bottom:
#         dpg.render_dearpygui_frame()
#         # Scroll to the bottom of the window
#         dpg.set_y_scroll("Chat_window_id", dpg.get_y_scroll_max("Chat_window_id"))


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
