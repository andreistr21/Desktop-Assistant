from pygame import mixer
from os import remove
from time import time, sleep


def DeleteRedundantFile(voiceover_shared_list):
    """
    Delete mp3 file after voiceover.

    :return: None
    """

    try:
        remove(f"resources/temp/voiceover{voiceover_shared_list[1] - 1}.mp3")
    except PermissionError as e:
        print("Can't delete file: permission denied")


def StopVoiceover(music, voiceover_shared_list, only_flags=False):
    """
    Stop voiceover.

    :param music: module (mixer.music)
    :param voiceover_shared_list: multiprocessing.managers.ListProxy
        Shared list for voiceover.
    :param only_flags: bool
        Switch parameter. If True then only flags will change.
    :return: None
    """

    if not only_flags:
        music.stop()

    voiceover_shared_list[0] = False
    # voiceover_shared_list[1] = False


def VoiceOver(voiceover_shared_list, start_time):
    """
    This function is running as a second process and voices the text.

    :param voiceover_shared_list: Manager.list
        List of shared variables
        [0] - switch (bool) (True - is speaking now), [1] - voiceover files counter
    :param start_time:
    :return: None
    """

    print(f"Voiceover process started: {time() - start_time}")

    mixer.init()

    while 1:
        if voiceover_shared_list[0]:
            mixer.music.load(
                f"resources/temp/voiceover{voiceover_shared_list[1] - 1}.mp3"
            )
            mixer.music.play()

            # Loop for voiceover
            while mixer.music.get_busy():
                if not voiceover_shared_list[0]:
                    StopVoiceover(mixer.music, voiceover_shared_list)

                # Decrease load on CPU during voiceover
                sleep(0.01)
            else:
                StopVoiceover(None, voiceover_shared_list, only_flags=True)

                # Release previous file
                mixer.music.load("resources/sounds/silence.mp3")

                DeleteRedundantFile(voiceover_shared_list)

        # Decrease load on CPU during waiting for voiceover
        sleep(0.05)
