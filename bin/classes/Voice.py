import pyttsx3


class Voice(object):
    def __init__(self, speech_rate):
        # Constructs a new TTS engine instance
        self.engine = pyttsx3.init()
        # noinspection PyUnresolvedReferences
        # Set voice type
        self.engine.setProperty("voice", self.engine.getProperty("voices")[1].id)
        # Set speed of speech (words per minute)
        self.engine.setProperty("rate", speech_rate)

    def Speech(self, text):
        self.engine.say(text)
        self.engine.runAndWait()

    def ChangeAssistantVolumeRate(self, volume_rate):
        self.engine.setProperty("rate", volume_rate)
