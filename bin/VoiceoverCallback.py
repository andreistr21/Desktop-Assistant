import traceback

from bin.classes.Voice import Voice

import time


def VoiceOver(voiceover_shared_list, start_time):
    """
    This function is running as a second process and voices the text.

    :param voiceover_shared_list: Manager.list
        List of shared variables
        [0] - assistant speech rate, [1] - voiceover text, [2] - switch (bool) (True - is speaking now),
        [3] - switch (True - need to stop voiceover)
    :param start_time:
    :return: None
    """

    print(f"Voiceover process started: {time.time() - start_time}")
    voice = Voice(voiceover_shared_list[0])
    previously_speech_rate = voiceover_shared_list[0]

    while 1:
        if voiceover_shared_list[2]:
            print(f"assistant speech rate: {voiceover_shared_list[0]}")
            if previously_speech_rate != voiceover_shared_list[0]:
                voice = Voice(voiceover_shared_list[0])
            voice.Speech(voiceover_shared_list[1])

            # Reset switch
            voiceover_shared_list[2] = False

        # Decrease load on CPU
        time.sleep(0.05)
