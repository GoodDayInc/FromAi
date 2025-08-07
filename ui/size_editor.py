import customtkinter as ctk
from tkinter import messagebox

class SizeEditor(ctk.CTkToplevel):
    """A Toplevel window for editing the size-to-article mapping."""

    def __init__(self, master, controller):
        super().__init__(master)
        self.controller = controller
        self.title("Редактор размеров")
        self.geometry("450x450")
        self.transient(master)
        self.grab_set()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Frame for the list
        list_frame = ctk.CTkFrame(self)
        list_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_columnconfigure(1, weight=1)

        # Header
        ctk.CTkLabel(list_frame, text="Размер", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=5, pady=2)
        ctk.CTkLabel(list_frame, text="Артикул", font=ctk.CTkFont(weight="bold")).grid(row=0, column=1, padx=5, pady=2)

        self.scrollable_frame = ctk.CTkScrollableFrame(list_frame, label_text="Текущие размеры")
        self.scrollable_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        self.scrollable_frame.grid_columnconfigure(0, weight=1)
        self.scrollable_frame.grid_columnconfigure(1, weight=1)
        list_frame.grid_rowconfigure(1, weight=1)

        # Frame for entries
        entry_frame = ctk.CTkFrame(self)
        entry_frame.grid(row=1, column=0, padx=10, pady=(0, 5), sticky="ew")
        entry_frame.grid_columnconfigure(1, weight=1)
        entry_frame.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(entry_frame, text="Размер:").grid(row=0, column=0, padx=(10, 5), pady=10)
        self.size_entry = ctk.CTkEntry(entry_frame)
        self.size_entry.grid(row=0, column=1, sticky="ew")

        ctk.CTkLabel(entry_frame, text="Артикул:").grid(row=0, column=2, padx=(10, 5), pady=10)
        self.article_entry = ctk.CTkEntry(entry_frame)
        self.article_entry.grid(row=0, column=3, padx=(0, 10), sticky="ew")

        # Frame for buttons
        btn_frame = ctk.CTkFrame(self)
        btn_frame.grid(row=2, column=0, padx=10, pady=(5, 10), sticky="ew")
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(btn_frame, text="Добавить/Обновить", command=self.add_or_update).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(btn_frame, text="Удалить выбранное", command=self.delete_selected).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.selected_row = None
        self.rows = {}

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.populate_list()

    def populate_list(self):
        """Fills the scrollable frame with the current size data."""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        self.rows = {}
        for i, (size, article) in enumerate(self.controller.size_to_article_map.items()):
            size_label = ctk.CTkLabel(self.scrollable_frame, text=size, anchor="w")
            size_label.grid(row=i, column=0, padx=5, pady=2, sticky="ew")

            article_label = ctk.CTkLabel(self.scrollable_frame, text=str(article), anchor="w")
            article_label.grid(row=i, column=1, padx=5, pady=2, sticky="ew")

            row_widgets = [size_label, article_label]
            self.rows[size] = row_widgets

            for widget in row_widgets:
                widget.bind("<Button-1>", lambda e, s=size: self.on_select(s))

    def on_select(self, size):
        """Handles row selection."""
        if self.selected_row and self.selected_row in self.rows:
            for widget in self.rows[self.selected_row]:
                widget.configure(fg_color="transparent")

        self.selected_row = size
        article = self.controller.size_to_article_map[size]

        for widget in self.rows[size]:
            widget.configure(fg_color="gray20")

        self.size_entry.delete(0, "end")
        self.size_entry.insert(0, size)
        self.article_entry.delete(0, "end")
        self.article_entry.insert(0, str(article))

    def add_or_update(self):
        """Adds a new or updates an existing size-article pair."""
        size = self.size_entry.get().strip()
        article_str = self.article_entry.get().strip()
        if not size or not article_str:
            messagebox.showwarning("Ошибка", "Оба поля должны быть заполнены.", parent=self)
            return
        try:
            article = int(article_str)
        except ValueError:
            messagebox.showwarning("Ошибка", "Артикул должен быть числом.", parent=self)
            return

        self.controller.size_to_article_map[size] = article
        self.controller.save_sizes()
        self.populate_list()
        self.size_entry.delete(0, "end")
        self.article_entry.delete(0, "end")
        self.selected_row = None

    def delete_selected(self):
        """Deletes the selected size-article pair."""
        if not self.selected_row:
            messagebox.showwarning("Ошибка", "Сначала выберите строку для удаления.", parent=self)
            return

        if messagebox.askyesno("Подтверждение", "Вы уверены?", parent=self):
            del self.controller.size_to_article_map[self.selected_row]
            self.controller.save_sizes()
            self.populate_list()
            self.selected_row = None

    def on_close(self):
        """Updates the main app before closing."""
        self.controller.update_converter_combobox()
        self.destroy()
