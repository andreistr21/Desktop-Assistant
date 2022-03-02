from math import ceil, floor
import dearpygui.dearpygui as dpg
from gtts import gTTS
from time import process_time

import bin.common.Common as common


class Dialog:
    def __init__(self):
        self.all_replicas = []

        self.one_line_max_pixels_text = 310
        self.text_x_pos = 387

        # Pixels for the chat
        self.pixels_y = 10

    def AssistantSays(
        self,
        text,
        voice_over_text="",
        another_text_for_voice_over=False,
        voiceover=True,
        logs=True,
        auto_scroll_to_bottom=True,
    ):
        pre_edit_text = text
        if logs:
            self.all_replicas.append(["Assistant", text])

        if not another_text_for_voice_over:
            voice_over_text = pre_edit_text

        if voiceover:
            self.CreateMP3File(voice_over_text)
            common.voiceover_shared_list[0] = True

        text_len = len(text)
        text_len_pixels = (
            text_len * 7
        )  # 7 pixels for one letter, in one line max 310 pixels

        if text_len_pixels <= self.one_line_max_pixels_text:
            dpg.add_text(text, parent="Chat_window_id", pos=[15, self.pixels_y + 10])

            with dpg.drawlist(
                width=text_len_pixels + 14,
                height=30,
                parent="Chat_window_id",
                pos=[9, self.pixels_y + 6],
            ):
                dpg.draw_rectangle(
                    pmin=[0, 0], pmax=[text_len_pixels + 14, 30], rounding=10
                )

            self.pixels_y += 40

        else:

            number_of_lines = self.NewLinesCounter(text)

            text, lines_to_add = self.TextDivisionIntoLines(text)

            number_of_lines += lines_to_add + 1

            dpg.add_text(text, parent="Chat_window_id", pos=[15, self.pixels_y + 10])

            with dpg.drawlist(
                width=self.one_line_max_pixels_text + 14,
                height=14 * number_of_lines + 15,
                parent="Chat_window_id",
                pos=[9, self.pixels_y + 6],
            ):
                dpg.draw_rectangle(
                    pmin=[0, 0],
                    pmax=[
                        self.one_line_max_pixels_text + 14,
                        14 * number_of_lines + 15,
                    ],
                    rounding=10,
                )

            self.pixels_y += 14 * number_of_lines + 15 + 10

        if auto_scroll_to_bottom:
            dpg.render_dearpygui_frame()
            # Scroll to the bottom of the window
            dpg.set_y_scroll("Chat_window_id", dpg.get_y_scroll_max("Chat_window_id"))

    def UserSays(self, text, logs=True, auto_scroll_to_bottom=True):
        if logs:
            self.all_replicas.append(["User", text])

        text_len = len(text)
        text_len_pixels = (
            text_len * 7
        )  # 7 pixels for one letter, in one line max 310 pixels

        if text_len_pixels <= self.one_line_max_pixels_text:
            dpg.add_text(
                text,
                parent="Chat_window_id",
                pos=[self.text_x_pos - text_len_pixels, self.pixels_y + 10],
            )

            with dpg.drawlist(
                width=text_len_pixels + 14,
                height=30,
                parent="Chat_window_id",
                pos=[self.text_x_pos + 6 - text_len_pixels - 12, self.pixels_y + 6],
            ):
                dpg.draw_rectangle(
                    pmin=[0, 0], pmax=[text_len_pixels + 14, 30], rounding=10
                )

            self.pixels_y += 40
        else:
            number_of_lines = ceil(text_len_pixels / self.one_line_max_pixels_text)
            number_of_lines += self.NewLinesCounter(text)

            text, lines_to_add = self.TextDivisionIntoLines(text)

            number_of_lines += lines_to_add

            dpg.add_text(
                text,
                parent="Chat_window_id",
                pos=[
                    self.text_x_pos - self.one_line_max_pixels_text,
                    self.pixels_y + 10,
                ],
            )

            with dpg.drawlist(
                width=self.one_line_max_pixels_text + 14,
                height=14 * number_of_lines + 15,
                parent="Chat_window_id",
                pos=[
                    self.text_x_pos + 6 - self.one_line_max_pixels_text - 12,
                    self.pixels_y + 6,
                ],
            ):
                dpg.draw_rectangle(
                    pmin=[0, 0],
                    pmax=[
                        self.one_line_max_pixels_text + 14,
                        14 * number_of_lines + 15,
                    ],
                    rounding=10,
                )

            self.pixels_y += 14 * number_of_lines + 15 + 10

        if auto_scroll_to_bottom:
            dpg.render_dearpygui_frame()
            # Scroll to the bottom of the window
            dpg.set_y_scroll("Chat_window_id", dpg.get_y_scroll_max("Chat_window_id"))

    @staticmethod
    def CreateMP3File(text: str):
        """
        Create mp3 file for voiceover and save it in current directory.

        :param text: str
            Text for voiceover
        :return: None
        """
        start_time = process_time()
        tts = gTTS(text, lang="en")
        print(f"Voiceover received: {process_time() - start_time}")

        try:
            start_time = process_time()
            tts.save(f"resources/temp/voiceover{common.voiceover_shared_list[1]}.mp3")
            print(f"File saved: {process_time() - start_time}")
            common.voiceover_shared_list[1] += 1
        except PermissionError as _:
            print("Can't save file: permission denied")

    @staticmethod
    def NewLinesCounter(text):
        counter = 0

        for letter in text:
            if letter == "\n":
                counter += 1

        return counter

    def TextDivisionIntoLines(self, text):
        max_letters_on_one_line = floor(self.one_line_max_pixels_text / 7)

        words_in_line_counter = 0
        word_start = 0
        letter_index = 0
        new_lines_counter = 0

        # Create new line if there is max number of letters
        while letter_index < len(text):
            if words_in_line_counter == max_letters_on_one_line:
                if text[letter_index] != " ":
                    if text[letter_index - 1] != " " or text[letter_index + 1] != " ":
                        # Search for start of word
                        for j in range(letter_index, 0, -1):
                            if text[j] != " ":
                                word_start = j
                            elif text[j] == " ":
                                break

                if word_start != 0:
                    letter_index = word_start
                    word_start = 0
                text = f"{text[:letter_index]}\n{text[letter_index:]}"
                new_lines_counter += 1
                words_in_line_counter = 0

            if text[letter_index] == "\n":
                words_in_line_counter = 0
            else:
                words_in_line_counter += 1

            letter_index += 1

        return text, new_lines_counter
