from keyboard import send
from urllib.request import Request, urlopen
from bs4 import BeautifulSoup
from subprocess import run
from os import system, popen
from regex import findall, IGNORECASE
from pynput.keyboard import Key, Controller

from bin.common import Common as common


def SwitchKeyboardLanguage(dialog):
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


def SearchInTheInternet(dialog, screen, quary):
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
        screen.ButtonCreate(url, dialog)
    elif soup.find(class_="kno-rdesc"):
        dialog.AssistantSays(soup.find(class_="kno-rdesc").text[11:-10])
        screen.ButtonCreate(url, dialog)
    else:
        dialog.AssistantSays("Opening in a browser...")
        open(url)


def OpenProgram(dialog, app_name):
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


def KillProgram(dialog, name: str) -> bool:
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


def ChangeVolume(audio_controller, dialog, volume_percents, change=False):
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


def MultimediaControl(keyboard, action):
    if action == "next":
        keyboard.press(Key.media_next)
    elif action == "previously":
        keyboard.press(Key.media_previous)
    elif action == "stop" or action == "play":
        keyboard.press(Key.media_play_pause)


def StopVoiceOver():
    """Stops the voiceover process"""
    # Set is_speaking_now flag to true
    common.voiceover_shared_list[0] = False
