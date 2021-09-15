import dearpygui.dearpygui as dpg
import math

one_line_max_pixels_text = 310


def TextEdit(text):
    max_letters_on_one_line = math.floor(one_line_max_pixels_text / 7)

    counter = 0
    for i in range(len(text)):
        if counter == max_letters_on_one_line:
            text = f"{text[:i]}\n{text[i:]}"
            counter = 0

        counter += 1

    return text


def AssistantSays(text, pixels):
    text_len = len(text)
    # number_of_spaces = SpacesCount(text)
    # text_len -= number_of_spaces
    text_len_pixels = (
        text_len * 7
    )  # 7 pixels for one letter, in one line max 310 pixels

    if text_len_pixels <= one_line_max_pixels_text:
        dpg.add_text(text, parent="Chat_window_id", pos=[15, pixels[0] + 10])

        with dpg.drawlist(
            width=text_len_pixels + 14,
            height=30,
            parent="Chat_window_id",
            pos=[9, pixels[0] + 6],
        ):
            dpg.draw_rectangle(
                pmin=[0, 0], pmax=[text_len_pixels + 14, 30], rounding=10
            )

        pixels[0] += 40

    else:
        number_of_lines = math.ceil(text_len_pixels / one_line_max_pixels_text)

        text = TextEdit(text)

        dpg.add_text(text, parent="Chat_window_id", pos=[15, pixels[0] + 10])

        with dpg.drawlist(
            width=one_line_max_pixels_text + 14,
            height=14 * number_of_lines + 15,
            parent="Chat_window_id",
            pos=[9, pixels[0] + 6],
        ):
            dpg.draw_rectangle(
                pmin=[0, 0],
                pmax=[one_line_max_pixels_text + 14, 14 * number_of_lines + 15],
                rounding=10,
            )

        pixels[0] += 14 * number_of_lines + 15 + 10


def UserSays(text, pixels):
    text_len = len(text)
    text_len_pixels = (
        text_len * 7
    )  # 7 pixels for one letter, in one line max 310 pixels

    if text_len_pixels <= one_line_max_pixels_text:
        dpg.add_text(
            text, parent="Chat_window_id", pos=[387 - text_len_pixels, pixels[0] + 10]
        )

        with dpg.drawlist(
            width=text_len_pixels + 14,
            height=30,
            parent="Chat_window_id",
            pos=[393 - text_len_pixels - 12, pixels[0] + 6],
        ):
            dpg.draw_rectangle(
                pmin=[0, 0], pmax=[text_len_pixels + 14, 30], rounding=10
            )

        pixels[0] += 40
    else:
        number_of_lines = math.ceil(text_len_pixels / one_line_max_pixels_text)

        text = TextEdit(text)

        dpg.add_text(
            text, parent="Chat_window_id", pos=[387 - one_line_max_pixels_text, pixels[0] + 10]
        )

        with dpg.drawlist(
            width=one_line_max_pixels_text + 14,
            height=14 * number_of_lines + 15,
            parent="Chat_window_id",
            pos=[393 - one_line_max_pixels_text - 12, pixels[0] + 6],
        ):
            dpg.draw_rectangle(
                pmin=[0, 0],
                pmax=[one_line_max_pixels_text + 14, 14 * number_of_lines + 15],
                rounding=10,
            )

        pixels[0] += 14 * number_of_lines + 15 + 10


def main():
    vp = dpg.create_viewport(title="Sergey", width=430, height=750)

    width, height, channels, data = dpg.load_image(
        "../../resources/images/microphone.png"
    )

    # Create a theme
    with dpg.theme(default_theme=True):
        dpg.add_theme_style(
            dpg.mvStyleVar_FrameRounding, 5, category=dpg.mvThemeCat_Core
        )
        dpg.add_theme_style(
            dpg.mvStyleVar_WindowBorderSize, 0, category=dpg.mvThemeCat_Core
        )
        dpg.add_theme_style(
            dpg.mvStyleVar_ScrollbarSize, 9, category=dpg.mvThemeCat_Core
        )
        dpg.add_theme_style(
            dpg.mvStyleVar_ScrollbarRounding, 12, category=dpg.mvThemeCat_Core
        )

    # Add a font registry
    with dpg.font_registry():
        dpg.add_font(
            "../../resources/fonts/arial_monospaced_mt.ttf", 14, default_font=True
        )  # 6 pixels for each element

    # Add an image
    with dpg.texture_registry():
        dpg.add_static_texture(width, height, data, id="microphone_id")

    # Main window
    with dpg.window(id="Main_window_id"):
        dpg.add_input_text(id="Text input", pos=[10, 680], width=355, hint="Ask Sergey")
        dpg.add_image("microphone_id", pos=[375, 670])

    # Chat window
    with dpg.window(
        id="Chat_window_id",
        no_move=True,
        width=412,
        height=600,
        no_collapse=True,
        no_resize=True,
        no_title_bar=True,
    ):
        pass

    dpg.set_primary_window("Main_window_id", True)
    dpg.setup_dearpygui(viewport=vp)

    dpg.show_viewport(vp)
    dpg.start_dearpygui()


if __name__ == "__main__":
    main()
