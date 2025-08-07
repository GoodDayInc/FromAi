import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import json
import threading
from pathlib import Path
import pandas as pd

try:
    import win32com.client as win32
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

from logic.config_manager import ConfigManager
from logic.file_operations import FileOperations
from utils.logger import Logger
from ui.navigation_frame import NavigationFrame
from ui.views import FileOperationsView, PathGeneratorView, FolderCreatorView, ArticleConverterView
from ui.size_editor import SizeEditor
from ui.widgets import PlaceholderEntry, Tooltip

SIZES_JSON_FILE = "sizes.json"
DEFAULT_SIZES = {
    "41 р": 1211561, "41.5 р": 1211562, "42 р": 1211563, "42.5 р": 1211564,
    "43 р": 1211565, "43.5 р": 1211566, "44 р": 1211567, "44.5 р": 1211568,
}

class MainWindow(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.current_thread: threading.Thread | None = None
        self.stop_event = threading.Event()
        self.operation_result_counter = 0
        self.operation_buttons = {}

        self.load_configuration()
        self._create_state_variables()
        self.setup_window()
        self.create_widgets()
        self.define_operations()
        self.setup_bindings()

        self.after(100, self.show_welcome_message)
        self.after(101, self.update_converter_combobox) # Populate combobox

    def _create_state_variables(self):
        """Creates all the tkinter variables for the application."""
        self.path_var = ctk.StringVar(value=self.last_path)
        self.dry_run_var = ctk.BooleanVar(value=self.last_dry_run)
        self.phrase_var = ctk.StringVar(value=self.last_phrase_to_remove)
        self.case_sensitive_phrase_var = ctk.BooleanVar(value=self.last_case_sensitive_phrase)
        self.use_regex_var = ctk.BooleanVar(value=self.last_use_regex)
        self.url_names_var = ctk.StringVar(value=self.last_url_names_to_delete)
        self.case_sensitive_url_var = ctk.BooleanVar(value=self.last_case_sensitive_url)
        self.folder_prefix_var = ctk.StringVar(value=self.last_folder_prefix)
        self.folder_suffix_var = ctk.StringVar(value=self.last_folder_suffix)
        self.folder_numbering_var = ctk.BooleanVar(value=self.last_folder_numbering)
        self.folder_start_num_var = ctk.IntVar(value=self.last_folder_start_num)
        self.folder_padding_var = ctk.IntVar(value=self.last_folder_padding)
        self.status_var = ctk.StringVar(value="Готов")

    def setup_window(self):
        self.title("🗂️ Супер Скрипт v3.0 (Refactored)")
        self.geometry("1100x800")
        self.minsize(900, 700)
        # Configure grid layout: 1 row, 2 columns (sidebar, main content)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        ctk.set_appearance_mode(self.current_theme_name)

    def load_configuration(self):
        self.config = ConfigManager.load_config()
        self.last_path = self.config.get("last_path", os.path.expanduser("~"))
        self.current_theme_name = self.config.get("theme", "dark")
        self.last_phrase_to_remove = self.config.get("last_phrase_to_remove", "")
        self.last_url_names_to_delete = self.config.get("last_url_names_to_delete", "")
        self.last_case_sensitive_phrase = self.config.get("last_case_sensitive_phrase", False)
        self.last_case_sensitive_url = self.config.get("last_case_sensitive_url", False)
        self.last_use_regex = self.config.get("last_use_regex", False)
        self.last_dry_run = self.config.get("last_dry_run", True)
        self.last_folder_prefix = self.config.get("last_folder_prefix", "")
        self.last_folder_suffix = self.config.get("last_folder_suffix", "")
        self.last_folder_numbering = self.config.get("last_folder_numbering", False)
        self.last_folder_start_num = self.config.get("last_folder_start_num", 1)
        self.last_folder_padding = self.config.get("last_folder_padding", 2)

        self.load_sizes()

    def create_widgets(self):
        # --- Sidebar ---
        self.navigation_frame = NavigationFrame(self, self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, sticky="nsw")
        self.nav_buttons = self.navigation_frame.get_buttons()

        # --- Main Content Area ---
        self.main_content_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_content_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_content_frame.grid_columnconfigure(0, weight=1)
        self.main_content_frame.grid_rowconfigure(1, weight=1)

        # --- Content Widgets ---
        self.create_path_panel(self.main_content_frame)
        self.create_views(self.main_content_frame)
        self.create_log_and_progress_panel(self.main_content_frame)

        # Select the default view
        self.select_view("file_ops")

    def create_views(self, parent):
        """Create all the view frames and store them."""
        self.views = {
            "file_ops": FileOperationsView(parent, self, fg_color="transparent"),
            "path_gen": PathGeneratorView(parent, self, fg_color="transparent"),
            "folder_creator": FolderCreatorView(parent, self, fg_color="transparent"),
            "article_converter": ArticleConverterView(parent, self, fg_color="transparent")
        }
        for view in self.views.values():
            view.grid(row=1, column=0, sticky='nsew', padx=0, pady=0)

    def select_view(self, view_name: str):
        """Show the selected view and highlight the corresponding button."""
        # Highlight the correct button
        for name, button in self.nav_buttons.items():
            button.configure(fg_color=("gray70", "gray30") if name == view_name else "transparent")

        # Show the correct frame
        for name, frame in self.views.items():
            if name == view_name:
                frame.grid(row=1, column=0, sticky='nsew')
                frame.tkraise()
            else:
                frame.grid_remove()

    def create_path_panel(self, parent):
        path_frame = ctk.CTkFrame(parent)
        path_frame.grid(row=0, column=0, sticky="new", pady=(0, 10))
        path_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(path_frame, text="Рабочая папка:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, sticky="w", padx=10, pady=10)

        self.path_var = ctk.StringVar(value=self.last_path)
        self.path_entry = PlaceholderEntry(path_frame, textvariable=self.path_var, placeholder="Введите или выберите путь...")
        self.path_entry.grid(row=0, column=1, sticky="ew", padx=(0, 5), pady=10)

        self.browse_btn = ctk.CTkButton(path_frame, text="Обзор...", command=self.browse_folder)
        self.browse_btn.grid(row=0, column=2, sticky="ew", padx=(0, 10), pady=10)

        self.operation_buttons["path_entry"] = self.path_entry
        self.operation_buttons["browse_btn"] = self.browse_btn

    def create_log_and_progress_panel(self, parent):
        bottom_frame = ctk.CTkFrame(parent)
        bottom_frame.grid(row=2, column=0, sticky="nsew", pady=(10,0))
        bottom_frame.grid_columnconfigure(0, weight=1)
        bottom_frame.grid_rowconfigure(0, weight=1)

        log_frame = ctk.CTkFrame(bottom_frame)
        log_frame.grid(row=0, column=0, sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)

        self.output_log = ctk.CTkTextbox(log_frame, wrap="word", font=("Consolas", 12))
        self.output_log.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.output_log.configure(state="disabled")
        self.logger = Logger(self.output_log)

        log_buttons_frame = ctk.CTkFrame(log_frame, fg_color="transparent")
        log_buttons_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)

        self.clear_log_btn = ctk.CTkButton(log_buttons_frame, text="🗑️ Очистить лог", command=self.clear_log, width=120)
        self.clear_log_btn.pack(side="left")

        self.save_log_btn = ctk.CTkButton(log_buttons_frame, text="📁 Сохранить лог", command=self.save_log_to_file, width=120)
        self.save_log_btn.pack(side="left", padx=5)

        self.stop_btn = ctk.CTkButton(log_buttons_frame, text="⏹️ Остановить", command=self.stop_current_operation, state="disabled", fg_color="red", hover_color="darkred")
        self.stop_btn.pack(side="right")

        progress_frame = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        progress_frame.grid(row=1, column=0, sticky="ew", pady=5, padx=5)
        progress_frame.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(progress_frame, mode="determinate")
        self.progress_bar.grid(row=0, column=0, sticky="ew")
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(progress_frame, text="", width=40)
        self.progress_label.grid(row=0, column=1, padx=(10, 0))

        # Status Bar
        self.status_label = ctk.CTkLabel(bottom_frame, textvariable=self.status_var, anchor="w")
        self.status_label.grid(row=2, column=0, sticky="ew", padx=10, pady=(5,0))

    def define_operations(self):
        """Defines all available operations in a structured dictionary."""
        self.operations = {
            "extract": {
                "name": "Извлечь из папок '1'", "function": FileOperations.organize_folders,
                "get_args": lambda: (self.path_var.get(), self.logger, self.stop_event, self.update_status, self.dry_run_var.get()),
                "is_file_op": True,
            },
            "rename_images": {
                "name": "Переименовать изображения 1-N", "function": FileOperations.rename_images_sequentially,
                "get_args": lambda: (self.path_var.get(), self.logger, self.stop_event, self.update_status, self.dry_run_var.get()),
                "is_file_op": True,
            },
            "remove_phrase": {
                "name": "Удалить фразу/RegEx из имен", "function": FileOperations.remove_phrase_from_names,
                "get_args": lambda: (self.path_var.get(), self.phrase_var.get(), self.logger, self.stop_event, self.update_status, self.case_sensitive_phrase_var.get(), self.use_regex_var.get(), self.dry_run_var.get()),
                "pre_check": lambda: self.phrase_var.get(), "pre_check_msg": "Пожалуйста, введите фразу или RegEx для удаления.",
                "is_file_op": True,
            },
            "delete_urls": {
                "name": "Удалить URL-ярлыки", "function": FileOperations.delete_url_shortcuts,
                "get_args": lambda: (self.path_var.get(), self.url_names_var.get(), self.logger, self.stop_event, self.update_status, self.case_sensitive_url_var.get(), self.dry_run_var.get()),
                "pre_check": lambda: self.url_names_var.get().strip(), "pre_check_msg": "Пожалуйста, введите имена URL-ярлыков.",
                "is_file_op": True,
            },
            "generate_paths": {
                "name": "Генерация путей для Excel", "function": FileOperations.generate_excel_paths,
                "get_args": lambda: (self.path_var.get(), self.path_gen_input_text.get("1.0", "end-1c"), self.logger, self.stop_event, self.update_status, self.path_gen_result_callback),
                "pre_check": lambda: self.path_gen_input_text.get("1.0", "end-1c").strip(), "pre_check_msg": "Пожалуйста, введите список моделей для генерации.",
                "is_file_op": False,
            },
            "create_folders": {
                "name": "Создание папок", "function": FileOperations.create_folders_from_list,
                "get_args": lambda: (self.path_var.get(), self.folder_creator_input_text.get("1.0", "end-1c"), self.folder_prefix_var.get(), self.folder_suffix_var.get(), self.folder_numbering_var.get(), self.folder_start_num_var.get(), self.folder_padding_var.get(), self.logger, self.stop_event, self.update_status, self.dry_run_var.get()),
                "pre_check": lambda: self.folder_creator_input_text.get("1.0", "end-1c").strip(), "pre_check_msg": "Пожалуйста, введите названия папок для создания.",
                "is_file_op": True,
            },
        }

    def setup_log_tags(self):
        colors = {"info": "#569CD6", "success": "#4EC9B0", "warning": "#FFCC00", "error": "#F44747"}
        for level, color in colors.items():
            self.output_log.tag_config(level, foreground=color)

    def setup_bindings(self):
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def run_operation(self, op_type: str):
        if self.current_thread and self.current_thread.is_alive():
            messagebox.showwarning("Операция выполняется", "Другая операция уже запущена.", parent=self)
            return

        op_details = self.operations.get(op_type)
        if not op_details:
            self.logger.error(f"Неизвестный тип операции: {op_type}")
            return

        if not self.validate_path(self.path_var.get(), op_details["name"]): return
        if op_details.get("pre_check") and not op_details["pre_check"]():
            messagebox.showwarning("Нет данных", op_details["pre_check_msg"], parent=self)
            return

        if op_details["is_file_op"] and not self.dry_run_var.get():
            if not self.confirm_operation(op_details["name"]):
                self.logger.info("Операция отменена пользователем.")
                self.update_status("Операция отменена.", 0)
                return

        self.operation_result_counter = 0
        self.last_op_type = op_type
        self.clear_log()
        self.title(f"🗂️ Выполняется: {op_details['name']}...")
        self.update_status(f"Запуск '{op_details['name']}'...", 0)
        self.stop_event.clear()
        self.set_ui_state(active=False)

        def operation_wrapper():
            args = op_details["get_args"]()
            result = op_details["function"](*args)
            if isinstance(result, int): self.operation_result_counter = result

        self.current_thread = threading.Thread(target=operation_wrapper, daemon=True)
        self.current_thread.start()
        self.after(100, self.check_thread_completion)

    def check_thread_completion(self):
        if self.current_thread and self.current_thread.is_alive():
            self.after(200, self.check_thread_completion)
        else:
            if self.current_thread and not self.stop_event.is_set() and not (op_type := self.operations.get(self.last_op_type)) is None and op_type["is_file_op"]:
                op_name = "Пробный запуск" if self.dry_run_var.get() else "Операция"
                messagebox.showinfo("Завершено", f"{op_name} завершена.\nОбработано: {self.operation_result_counter}", parent=self)
            self.set_ui_state(active=True)
            self.title("🗂️ Супер Скрипт v3.0")
            if self.stop_event.is_set(): self.update_status("Операция остановлена.", 0)
            self.current_thread = None

    def stop_current_operation(self):
        if self.current_thread and self.current_thread.is_alive():
            if messagebox.askyesno("Остановить?", "Прервать текущую операцию?", icon="warning", parent=self):
                self.stop_event.set()
                self.logger.warning("--- Попытка остановить операцию... ---")
                self.update_status("Остановка...", None)
                self.stop_btn.configure(state="disabled")

    def set_ui_state(self, active: bool):
        state = "normal" if active else "disabled"
        for widget in self.operation_buttons.values():
            widget.configure(state=state)
        self.stop_btn.configure(state="disabled" if active else "normal")
        self.navigation_frame.theme_btn.configure(state=state)
        self.navigation_frame.help_btn.configure(state=state)

    def show_welcome_message(self):
        self.logger.info("🎉 Добро пожаловать в Супер Скрипт v3.0!")
        self.logger.info("💡 Выберите вкладку, папку и операцию для начала работы.")
        self.update_status("Готов к работе.")
        self.setup_log_tags()

    def toggle_theme(self):
        new_mode = "light" if ctk.get_appearance_mode() == "Dark" else "dark"
        ctk.set_appearance_mode(new_mode)
        self.current_theme_name = new_mode.lower()
        self.logger.info(f"Тема изменена на: {self.current_theme_name.capitalize()}")

    def show_help(self):
        messagebox.showinfo("Справка", "Текст справки будет добавлен в будущем релизе.", parent=self)

    def browse_folder(self):
        initial_dir = self.path_var.get() if os.path.isdir(self.path_var.get()) else self.last_path
        folder = filedialog.askdirectory(initialdir=initial_dir, title="Выберите папку", parent=self)
        if folder:
            self.path_var.set(folder)
            self.logger.info(f"Выбрана папка: {folder}")
            self.update_status(f"Папка: {os.path.basename(folder)}")

    def clear_log(self):
        self.output_log.configure(state="normal")
        self.output_log.delete(1.0, "end")
        self.output_log.configure(state="disabled")
        self.logger.info("Журнал операций очищен.")

    def save_log_to_file(self):
        content = self.output_log.get(1.0, "end-1c")
        if not content.strip(): return
        path = filedialog.asksaveasfilename(defaultextension=".log", filetypes=[("Log files", "*.log")], parent=self)
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f: f.write(content)
                self.logger.info(f"Лог сохранен в {path}")
            except Exception as e:
                self.logger.error(f"Ошибка сохранения лога: {e}")

    def update_status(self, message: str, progress: int | None = None):
        self.status_var.set(message)
        if progress is not None:
            self.progress_bar.set(progress / 100)
            self.progress_label.configure(text=f"{progress}%")
        else:
            self.progress_bar.set(0)
            self.progress_label.configure(text="")

    def validate_path(self, path: str, op_name: str) -> bool:
        if not path or not os.path.isdir(path):
            messagebox.showerror("Ошибка", f"Путь '{path}' не существует или не папка.", parent=self)
            return False
        return True

    def confirm_operation(self, op_name: str) -> bool:
        return messagebox.askyesno("Подтверждение", f"Вы уверены, что хотите запустить '{op_name}'?", icon="warning", parent=self)

    def path_gen_result_callback(self, success_str: str, error_str: str):
        def update_ui():
            for widget, text in [(self.path_gen_output_text, success_str), (self.path_gen_error_text, error_str)]:
                widget.configure(state="normal")
                widget.delete("1.0", "end")
                widget.insert("1.0", text)
                widget.configure(state="disabled")
        self.after(0, update_ui)

    def copy_path_gen_results(self):
        content = self.path_gen_output_text.get("1.0", "end-1c").strip()
        if content:
            self.clipboard_clear()
            self.clipboard_append(content)
            self.logger.info("Результаты скопированы.")

    def on_closing(self):
        self.config["last_path"] = self.path_var.get()
        self.config["theme"] = self.current_theme_name
        self.config["last_phrase_to_remove"] = self.phrase_var.get()
        self.config["last_url_names_to_delete"] = self.url_names_var.get()
        self.config["last_case_sensitive_phrase"] = self.case_sensitive_phrase_var.get()
        self.config["last_case_sensitive_url"] = self.case_sensitive_url_var.get()
        self.config["last_use_regex"] = self.use_regex_var.get()
        self.config["last_dry_run"] = self.dry_run_var.get()
        self.config["last_folder_prefix"] = self.folder_prefix_var.get()
        self.config["last_folder_suffix"] = self.folder_suffix_var.get()
        self.config["last_folder_numbering"] = self.folder_numbering_var.get()
        self.config["last_folder_start_num"] = self.folder_start_num_var.get()
        self.config["last_folder_padding"] = self.folder_padding_var.get()
        ConfigManager.save_config(self.config)
        self.save_sizes()
        self.destroy()

    def load_sizes(self):
        try:
            with open(SIZES_JSON_FILE, "r", encoding="utf-8") as f:
                self.size_to_article_map = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            self.size_to_article_map = DEFAULT_SIZES
            self.save_sizes()

    def save_sizes(self):
        try:
            with open(SIZES_JSON_FILE, "w", encoding="utf-8") as f:
                json.dump(self.size_to_article_map, f, ensure_ascii=False, indent=4)
            if hasattr(self, "logger"): self.logger.info(f"Словарь размеров сохранен.")
        except Exception as e:
            if hasattr(self, "logger"): self.logger.error(f"Не удалось сохранить словарь: {e}")

    def update_converter_combobox(self):
        self.converter_size_combobox.configure(values=list(self.size_to_article_map.keys()))

    def open_size_editor(self):
        self.logger.info("Открыт редактор размеров.")
        editor = SizeEditor(self, self)

    def universal_file_reader(self, file_path: str) -> pd.DataFrame:
        try: return pd.read_excel(file_path, header=None, dtype=str)
        except Exception:
            try: return pd.read_csv(file_path, header=None, dtype=str, engine="python", encoding="utf-8-sig")
            except Exception:
                try: return pd.read_csv(file_path, header=None, dtype=str, engine="python", encoding="cp1251")
                except Exception as e: raise ValueError(f"Не удалось прочитать файл: {e}")

    def select_and_scan_converter_file(self):
        file_path = filedialog.askopenfilename(title="Выберите файл", filetypes=[("Таблицы", "*.xlsx *.xls *.csv")])
        if not file_path: return

        self.converter_input_file_path = file_path
        self.converter_detected_article = None
        self.converter_process_btn.configure(state="disabled")
        self.converter_size_combobox.configure(state="disabled")
        self.converter_size_combobox.set("")

        filename = os.path.basename(file_path)
        self.converter_file_label.configure(text=f"Выбран: {filename}")
        self.logger.info(f"Конвертер: выбран файл '{file_path}'")
        self.update_status(f"Сканирование {filename}...")

        try:
            df = self.universal_file_reader(self.converter_input_file_path)
            article_to_size_map = {str(v): k for k, v in self.size_to_article_map.items()}
            articles_set = {str(v) for v in self.size_to_article_map.values()}

            for col in df.columns:
                for cell in df[col].dropna():
                    if isinstance(cell, str) and (article := next((a for a in articles_set if a in cell), None)):
                        self.converter_detected_article = article
                        break
                if self.converter_detected_article: break

            if self.converter_detected_article:
                detected_size = article_to_size_map[self.converter_detected_article]
                self.converter_detected_label.configure(text=f"Найден размер: {detected_size}")
                self.converter_size_combobox.configure(state="readonly")
                self.converter_process_btn.configure(state="normal")
                self.update_status(f"Найден {detected_size}. Выберите новый размер.", 0)
            else:
                self.converter_detected_label.configure(text="Артикул не найден!")
                self.logger.error("В файле не найден известный артикул.")
                self.update_status("Артикул не найден.", 0)

        except Exception as e:
            self.logger.error(f"Ошибка чтения файла: {e}")
            messagebox.showerror("Ошибка", str(e), parent=self)

    def process_and_save_converter_file(self):
        new_size = self.converter_size_combobox.get()
        if not new_size: return

        self.update_status("Обработка...", 50)

        try:
            df = self.universal_file_reader(self.converter_input_file_path)
            new_article = str(self.size_to_article_map[new_size])
            df = df.applymap(lambda cell: cell.replace(self.converter_detected_article, new_article) if isinstance(cell, str) else cell)

            original = Path(self.converter_input_file_path)
            ext = original.suffix.lower()
            if ext not in [".xls", ".xlsx", ".csv"]: ext = ".xlsx"

            sugg_name = f"{original.stem}_{new_size.replace(' ', '')}{ext}"

            out_path_str = filedialog.asksaveasfilename(title="Сохранить как", defaultextension=ext, initialfile=sugg_name, parent=self)
            if not out_path_str: return

            output_path = Path(out_path_str)
            if output_path.suffix.lower() == ".xls" and HAS_WIN32:
                # Handle legacy .xls format via COM
                excel = win32.gencache.EnsureDispatch("Excel.Application")
                excel.DisplayAlerts = False
                temp_path = str(output_path.with_suffix(".tmp.xlsx"))
                df.to_excel(temp_path, index=False, header=False)
                wb = excel.Workbooks.Open(os.path.abspath(temp_path))
                wb.SaveAs(os.path.abspath(out_path_str), FileFormat=56) # 56 is for xlExcel8
                wb.Close()
                excel.Application.Quit()
                os.remove(temp_path)
            elif output_path.suffix.lower() == ".csv":
                df.to_csv(output_path, index=False, header=False, encoding="utf-8-sig")
            else: # Default to xlsx
                df.to_excel(output_path, index=False, header=False)

            self.logger.success(f"Файл для '{new_size}' создан: {output_path}")
            messagebox.showinfo("Успех!", f"Файл для '{new_size}' создан!", parent=self)
            self.update_status("Готово.", 100)

        except Exception as e:
            self.logger.error(f"Ошибка обработки: {e}")
            messagebox.showerror("Ошибка", f"Не удалось обработать файл.\n{e}", parent=self)
            self.update_status("Ошибка.", 0)
