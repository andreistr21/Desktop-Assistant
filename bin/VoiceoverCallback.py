import pythoncom

from bin.classes.Voice import Voice


def VoiceOver(text, speech_rate):
    # noinspection PyUnresolvedReferences
    # COM initializing for thread
    pythoncom.CoInitialize()

    voice = Voice(speech_rate)
    voice.Speech(text)
