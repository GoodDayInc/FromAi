import customtkinter as ctk

class NavigationFrame(ctk.CTkFrame):
    def __init__(self, master, controller, **kwargs):
        super().__init__(master, **kwargs)
        self.controller = controller

        self.grid_rowconfigure(4, weight=1) # Pushes controls to the bottom

        self.title_label = ctk.CTkLabel(self, text="🗂️ Супер Скрипт",
                                        font=ctk.CTkFont(size=20, weight="bold"))
        self.title_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.file_ops_button = ctk.CTkButton(self, corner_radius=0, height=40, border_spacing=10,
                                             text="Файловые Операции",
                                             fg_color="transparent", text_color=("gray10", "gray90"),
                                             hover_color=("gray70", "gray30"),
                                             anchor="w", command=lambda: self.controller.select_view("file_ops"))
        self.file_ops_button.grid(row=1, column=0, sticky="ew")

        self.path_gen_button = ctk.CTkButton(self, corner_radius=0, height=40, border_spacing=10,
                                             text="Генератор Путей",
                                             fg_color="transparent", text_color=("gray10", "gray90"),
                                             hover_color=("gray70", "gray30"),
                                             anchor="w", command=lambda: self.controller.select_view("path_gen"))
        self.path_gen_button.grid(row=2, column=0, sticky="ew")

        self.folder_creator_button = ctk.CTkButton(self, corner_radius=0, height=40, border_spacing=10,
                                                   text="Создатель Папок",
                                                   fg_color="transparent", text_color=("gray10", "gray90"),
                                                   hover_color=("gray70", "gray30"),
                                                   anchor="w", command=lambda: self.controller.select_view("folder_creator"))
        self.folder_creator_button.grid(row=3, column=0, sticky="ew")

        self.article_converter_button = ctk.CTkButton(self, corner_radius=0, height=40, border_spacing=10,
                                                      text="Конвертер Артикулов",
                                                      fg_color="transparent", text_color=("gray10", "gray90"),
                                                      hover_color=("gray70", "gray30"),
                                                      anchor="w", command=lambda: self.controller.select_view("article_converter"))
        self.article_converter_button.grid(row=4, column=0, sticky="ew")

        self.theme_btn = ctk.CTkButton(self, text="Сменить тему", command=self.controller.toggle_theme)
        self.theme_btn.grid(row=5, column=0, padx=20, pady=10, sticky="s")

        self.help_btn = ctk.CTkButton(self, text="Справка", command=self.controller.show_help)
        self.help_btn.grid(row=6, column=0, padx=20, pady=(0, 20), sticky="s")

    def get_buttons(self):
        """Returns a dictionary of navigation buttons for highlighting."""
        return {
            "file_ops": self.file_ops_button,
            "path_gen": self.path_gen_button,
            "folder_creator": self.folder_creator_button,
            "article_converter": self.article_converter_button,
        }
