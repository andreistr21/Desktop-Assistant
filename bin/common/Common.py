from bin.classes.Voice import Voice

one_line_max_pixels_text = 310

# Pixels for the chat
pixels_y = [10]


voice = Voice()


def VoiceOver(text):
    voice.Speech(text)
