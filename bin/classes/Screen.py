import dearpygui.dearpygui as dpg
from webbrowser import open


class Screen:
    def __init__(self, dialog):
        self.viewport_height = None
        self.viewport_width = None

        self.chat_window_height = None
        self.chat_window_width = None
        self.input_text_width = None
        self.input_text_pos = None
        # dialog.one_line_max_pixels_text = None
        self.image_btn_pos = None
        # dialog.text_x_pos = None

    def ViewportResize(self, dialog):
        """Window resize handler. Restores dialog."""
        self.viewport_height = dpg.get_viewport_height()
        self.viewport_width = dpg.get_viewport_width()

        self.chat_window_height = self.viewport_height - 100
        self.chat_window_width = self.viewport_width - 18
        self.input_text_width = self.viewport_width - 75
        self.input_text_pos = [10, self.viewport_height - 70]
        dialog.one_line_max_pixels_text = self.viewport_width - 120
        self.image_btn_pos = [self.viewport_width - 60, self.viewport_height - 85]
        dialog.text_x_pos = self.viewport_width - 43

        self.GUIChanger()

        dpg.delete_item("Chat_window_id", children_only=True)
        dialog.pixels_y = 10

        for item in dialog.all_replicas:
            if item[0] == "Assistant":
                dialog.AssistantSays(
                    item[1],
                    voiceover=False,
                    logs=False,
                    auto_scroll_to_bottom=False,
                )
            elif item[0] == "User":
                dialog.UserSays(item[1], logs=False, auto_scroll_to_bottom=False)
            elif item[0] == "Button":
                self.ButtonCreate(item[1], dialog, logs=False)

        # Scroll to the bottom of the window
        dpg.set_y_scroll("Chat_window_id", dpg.get_y_scroll_max("Chat_window_id"))

    def GUIChanger(self):
        """Change resolution of each GUI element in the window."""

        dpg.set_item_width("Chat_window_id", self.chat_window_width)
        dpg.set_item_height("Chat_window_id", self.chat_window_height)

        dpg.set_item_pos("Text_input_id", self.input_text_pos)
        dpg.set_item_width("Text_input_id", self.input_text_width)

        dpg.set_item_pos("microphone_btn_id", self.image_btn_pos)

    def ButtonCreate(
        self,
        url: str,
        dialog,
        logs=True,
    ):
        """Create button under the text
        Args:
            :param url: (str)
            :param logs:  (bool)
                if True, remembers the button as created (new)
            :param dialog:
        """
        dpg.add_button(
            label="Open in a browser",
            width=dialog.one_line_max_pixels_text + 14,
            height=25,
            parent="Chat_window_id",
            callback=self.OpenInABrowserButtonClicked,
            user_data=url,
            pos=[9, dialog.pixels_y + 6],
        )

        dialog.pixels_y += 35

        if logs:
            dialog.all_replicas.append(["Button", url])

            # Update interface (render one frame)
            dpg.render_dearpygui_frame()
            # Scroll to the bottom of the window
            dpg.set_y_scroll("Chat_window_id", dpg.get_y_scroll_max("Chat_window_id"))

    @staticmethod
    def OpenInABrowserButtonClicked(_, __, url):
        open(url)
