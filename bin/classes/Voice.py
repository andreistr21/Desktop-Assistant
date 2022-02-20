from pyttsx3 import init

from bin.common.Common import voiceover_shared_list


class Voice(object):
    def __init__(self, speech_rate):
        required_id = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_DAVID_11.0"
        # Constructs a new TTS engine instance
        self.engine = init()
        self.engine.connect('started-word', self.OnWord())
        required_index = -1

        # Search for required voice
        voices_list = [self.engine.getProperty("voices")][0]
        # noinspection PyTypeChecker
        for voice_index in range(len(voices_list)):
            # noinspection PyUnresolvedReferences
            if self.engine.getProperty("voices")[voice_index].id == required_id:
                required_index = voice_index
                break

        # noinspection PyUnresolvedReferences
        # Set voice type
        self.engine.setProperty(
            "voice", self.engine.getProperty("voices")[required_index].id
        )

        # Set speed of speech (words per minute)
        self.engine.setProperty("rate", speech_rate)

    def OnWord(self):
        if voiceover_shared_list[3]:
            self.engine.stop()
            voiceover_shared_list[3] = False

    def Speech(self, text):
        self.engine.say(text)
        self.engine.runAndWait()

    def ChangeAssistantVolumeRate(self, volume_rate):
        self.engine.setProperty("rate", volume_rate)
