from bin.classes.Voice import Voice

import time


def VoiceOver(text, speech_rate, start_time):
    print(f"Voiceover precess started: {time.time() - start_time}")

    voice = Voice(speech_rate)
    voice.Speech(text)

