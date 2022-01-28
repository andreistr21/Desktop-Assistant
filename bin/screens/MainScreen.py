import dearpygui.dearpygui as dpg
from multiprocessing import freeze_support
from spacy import load

from bin.Main import (
    CommandAnalysisCall,
    AssistantSays,
    CommandRecognition,
    TerminateVoiceover,
    ViewportResize,
)
import resources.Strings as strings
from bin.common import Common as common

nlp = load("en_core_web_trf")


def GUICreator(
    chat_window_width,
    chat_window_height,
    input_text_width,
    input_text_pos,
    image_btn_pos,
    image_width,
    image_height,
):
    # Main window
    with dpg.window(id="Main_window_id"):
        dpg.add_input_text(
            id="Text_input_id",
            pos=input_text_pos,
            width=input_text_width,
            hint="Ask Sergey",
            on_enter=True,
            callback=CommandAnalysisCall,
        )

        dpg.add_image_button(
            "microphone_image_id",
            id="microphone_btn_id",
            width=image_width,
            height=image_height,
            callback=CommandRecognition,
            pos=image_btn_pos,
        )

    # Chat window
    with dpg.window(
        id="Chat_window_id",
        no_move=True,
        width=chat_window_width,
        height=chat_window_height,
        no_collapse=True,
        no_resize=True,
        no_title_bar=True,
    ):
        pass


def main():
    freeze_support()

    vp = dpg.create_viewport(title="Sergey", width=430, height=750)
    dpg.set_viewport_small_icon("resources/images/icon_small.ico")
    dpg.set_viewport_large_icon("resources/images/icon_large.ico")

    image_width, image_height, channels, data = dpg.load_image(
        "resources/images/microphone.png"
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
            "resources/fonts/arial_monospaced_mt.ttf", 14, default_font=True
        )  # 6 pixels for each element

    # Add an image
    with dpg.texture_registry():
        dpg.add_static_texture(
            image_width, image_height, data, id="microphone_image_id"
        )

    GUICreator(412, 650, 355, [10, 680], [370, 665], image_width, image_height)

    dpg.set_primary_window("Main_window_id", True)
    dpg.setup_dearpygui(viewport=vp)

    dpg.set_viewport_resize_callback(ViewportResize)

    dpg.show_viewport(vp)
    AssistantSays(strings.welcome_str, common.pixels_y)
    dpg.start_dearpygui()
    # Terminate voiceover if exist
    TerminateVoiceover()
