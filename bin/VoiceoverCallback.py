import pythoncom

from bin.classes.Voice import Voice


def VoiceOver(text, speech_rate):
    voice = Voice(speech_rate)
    voice.Speech(text)
