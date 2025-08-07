import os
import shutil
import re
import json
import tempfile
import threading
import datetime
from pathlib import Path
from typing import Dict, Any, Callable, Optional, List

import customtkinter as ctk
from tkinter import filedialog
import pandas as pd

try:
    import win32com.client as win32
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# --- Constants ---
APP_NAME = "SuperScript"
APP_CONFIG_DIR = Path.home() / f".{APP_NAME.lower()}"
CONFIG_FILE = APP_CONFIG_DIR / "file_organizer_config.json"
SIZES_JSON_FILE = APP_CONFIG_DIR / "sizes.json"
IMAGE_EXTENSIONS = [
    ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".webp", ".svg", ".ico"
]
DEFAULT_SIZES = {
    "41 р": 1211561, "41.5 р": 1211562, "42 р": 1211563, "42.5 р": 1211564,
    "43 р": 1211565, "43.5 р": 1211566, "44 р": 1211567, "44.5 р": 1211568,
}

# --- Backend Logic ---

class ConfigManager:
    """Handles loading and saving of the application configuration."""

    @staticmethod
    def _ensure_config_dir():
        APP_CONFIG_DIR.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def load_config() -> Dict[str, Any]:
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                print(f"Error loading config file: {CONFIG_FILE}")
                return {}
        return {}

    @staticmethod
    def save_config(config: Dict[str, Any]) -> None:
        try:
            ConfigManager._ensure_config_dir()
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Error saving config: {e}")


class Logger:
    """Manages logging to the GUI's text widget."""

    def __init__(self, output_widget: ctk.CTkTextbox):
        self.output_widget = output_widget

    def log(self, message: str, level: str = "info") -> None:
        if not self.output_widget:
            return
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"

        self.output_widget.configure(state="normal")
        self.output_widget.insert("end", formatted_message, level)
        self.output_widget.configure(state="disabled")
        self.output_widget.see("end")

    def info(self, message: str) -> None:
        self.log(message, "info")

    def success(self, message: str) -> None:
        self.log(f"✓ {message}", "success")

    def warning(self, message: str) -> None:
        self.log(f"⚠ {message}", "warning")

    def error(self, message: str) -> None:
        self.log(f"✗ {message}", "error")


class FileOperations:
    """A collection of static methods for performing file system operations."""

    @staticmethod
    def _sanitize_folder_name(name: str) -> str:
        return re.sub(r'[?:"<>|*]', "", name).strip()

    @staticmethod
    def create_folders_from_list(
        base_path_str: str, folder_list_str: str, prefix: str, suffix: str,
        use_numbering: bool, start_num: int, padding: int,
        logger: Logger, stop_event: threading.Event,
        status_callback: Callable[[str, int], None], dry_run: bool = False
    ) -> int:
        op_prefix = "[DRY RUN] " if dry_run else ""
        op_name = "Создание папок"
        logger.info(f"{op_prefix}🏗️ Начинаем операцию: {op_name}")
        if not dry_run:
            logger.warning("⚠️ ВНИМАНИЕ: Операция создает папки на диске!")

        base_path = Path(base_path_str)
        logger.info(f"Целевая директория: {base_path}")

        folder_names = [name.strip() for name in folder_list_str.strip().split("\n") if name.strip()]

        if not folder_names:
            logger.warning("Список папок пуст. Операция прервана.")
            status_callback("Список папок пуст.", 0)
            return 0

        created_count = 0
        total_folders = len(folder_names)
        try:
            for i, name in enumerate(folder_names):
                if stop_event.is_set(): break
                progress = int((i + 1) / total_folders * 100)
                status_callback(f"{op_prefix}Создание: {name}", progress)

                path_parts = [FileOperations._sanitize_folder_name(part) for part in re.split(r"[\\/]", name)]
                if not any(path_parts):
                    logger.warning(f"Пропущено: имя '{name}' стало пустым после очистки.")
                    continue
                
                number_str = str(i + start_num).zfill(padding) + "_" if use_numbering else ""
                path_parts[-1] = f"{prefix}{number_str}{path_parts[-1]}{suffix}"
                final_path = base_path.joinpath(*path_parts)

                try:
                    if not dry_run: final_path.mkdir(parents=True, exist_ok=True)
                    logger.success(f"{op_prefix}Создана папка: '{final_path.relative_to(base_path)}'")
                    created_count += 1
                except OSError as e:
                    logger.error(f"Ошибка создания папки '{final_path.name}': {e}")
        finally:
            if not stop_event.is_set(): status_callback("Готово.", 100)
            logger.info(f"--- {op_prefix}Операция '{op_name}' завершена ---")
        return created_count

    @staticmethod
    def generate_excel_paths(
        base_path_str: str, model_list_str: str, logger: Logger,
        stop_event: threading.Event, status_callback: Callable[[str, int], None],
        result_callback: Callable[[str, str], None]
    ):
        # This function is not refactored with the helper because it has a custom result callback
        op_name = "Генерация путей для Excel"
        logger.info(f"📋 Начинаем операцию: {op_name}")
        base_path = Path(base_path_str)

        def natural_sort_key(p: Path):
            return [int(text) if text.isdigit() else text.lower() for text in re.split("([0-9]+)", p.name)]

        model_list = [name.strip() for name in model_list_str.strip().split("\n") if name.strip()]
        if not model_list:
            logger.warning("Список моделей пуст.")
            return

        success_output, error_output = [], []
        total_models = len(model_list)
        try:
            for i, model_name in enumerate(model_list):
                if stop_event.is_set(): break
                progress = int((i + 1) / total_models * 100)
                status_callback(f"Проверка: {model_name}", progress)
                model_path = base_path / model_name

                if not model_path.is_dir():
                    error_output.append(f"{model_name} -> ОШИБКА: Папка не найдена!")
                    logger.error(f"Папка для '{model_name}' не найдена: {model_path}")
                    continue
                
                try:
                    photo_paths = [p for p in model_path.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_EXTENSIONS]
                except OSError as e:
                    error_output.append(f"{model_name} -> ОШИБКА: Не удалось прочитать папку: {e}")
                    continue

                if photo_paths:
                    sorted_paths = sorted(photo_paths, key=natural_sort_key)
                    final_string = f'"[+\n+|'.join(map(str, sorted_paths)) + ']"'
                    success_output.append(final_string)
                    logger.success(f"Пути для '{model_name}' ({len(sorted_paths)} фото) сгенерированы.")
                else:
                    error_output.append(f"{model_name} -> ОШИБКА: Изображения не найдены.")
                    logger.warning(f"Для '{model_name}' не найдены изображения.")
        finally:
            result_callback("\n".join(success_output), "\n".join(error_output))
            if not stop_event.is_set(): status_callback("Готово.", 100)
            logger.info(f"--- Операция '{op_name}' завершена ---")
            
# --- GUI ---

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("🗂️ Супер Скрипт v3.0")
        self.geometry("1100x800")
        self.minsize(900, 700)
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # --- Backend Init ---
        self.config = ConfigManager.load_config()
        self.size_to_article_map = {}
        self.current_thread = None
        self.stop_event = threading.Event()
        self.operation_result_counter = 0
        
        self._create_widgets()
        self.logger = Logger(self.log_textbox)
        self.load_sizes()
        self.load_configuration()
        self.define_operations()

        self.logger.info("🎉 Добро пожаловать в Супер Скрипт v3.0!")
        self.logger.info("💡 Выберите вкладку, папку и операцию для начала работы.")

    def _create_widgets(self):
        self._create_top_frame()
        self._create_tab_view()
        self._create_bottom_frame()

    def _create_top_frame(self):
        self.top_frame = ctk.CTkFrame(self)
        self.top_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        self.top_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(self.top_frame, text="Рабочая папка:").grid(row=0, column=0, padx=10, pady=10)
        self.path_var = ctk.StringVar()
        self.path_entry = ctk.CTkEntry(self.top_frame, textvariable=self.path_var, placeholder_text="Выберите или введите путь...")
        self.path_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.browse_btn = ctk.CTkButton(self.top_frame, text="Обзор...", command=self.browse_folder)
        self.browse_btn.grid(row=0, column=2, padx=10, pady=10)

        self.theme_switch = ctk.CTkSwitch(self.top_frame, text="Темная тема", command=self.toggle_theme)
        self.theme_switch.grid(row=0, column=3, padx=10, pady=10)

    def _create_tab_view(self):
        self.tab_view = ctk.CTkTabview(self, anchor="w")
        self.tab_view.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        self.tab_view.add("🗂️ Файловые Операции")
        self.tab_view.add("📋 Генератор Путей")
        self.tab_view.add("🏗️ Создатель Папок")
        self.tab_view.add("🔄 Конвертер Артикулов")

        self._populate_file_ops_tab(self.tab_view.tab("🗂️ Файловые Операции"))
        self._populate_path_gen_tab(self.tab_view.tab("📋 Генератор Путей"))
        self._populate_folder_creator_tab(self.tab_view.tab("🏗️ Создатель Папок"))
        self._populate_article_converter_tab(self.tab_view.tab("🔄 Конвертер Артикулов"))

    def _populate_file_ops_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        
        op_frame = ctk.CTkFrame(tab)
        op_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        op_frame.grid_columnconfigure((0,1), weight=1)

        self.file_op_var = ctk.StringVar(value="")
        ops = [
            ("Извлечь из '1'", "extract"), ("Переименовать 1-N", "rename_images"),
            ("Удалить фразу/RegEx", "remove_phrase"), ("Удалить URL-ярлыки", "delete_urls")
        ]
        for i, (text, value) in enumerate(ops):
            ctk.CTkRadioButton(op_frame, text=text, variable=self.file_op_var, value=value).grid(
                row=i//2, column=i%2, padx=10, pady=5, sticky="w"
            )

        exec_frame = ctk.CTkFrame(tab)
        exec_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        exec_frame.grid_columnconfigure(0, weight=1)

        self.dry_run_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(exec_frame, text="Пробный запуск (Dry Run)", variable=self.dry_run_var).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.run_op_button = ctk.CTkButton(exec_frame, text="Выполнить", command=lambda: self.run_operation(self.file_op_var.get()))
        self.run_op_button.grid(row=0, column=1, padx=10, pady=10, sticky="e")

    def _populate_path_gen_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(tab, text="1. Введите названия моделей (каждое с новой строки):").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.path_gen_input = ctk.CTkTextbox(tab)
        self.path_gen_input.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        ctk.CTkButton(tab, text="Сгенерировать и проверить пути", command=lambda: self.run_operation("generate_paths")).grid(row=2, column=0, padx=10, pady=10, sticky="ew")

        results_frame = ctk.CTkFrame(tab)
        results_frame.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")
        results_frame.grid_columnconfigure(0, weight=1)
        results_frame.grid_rowconfigure(1, weight=1)
        
        self.path_gen_output = ctk.CTkTextbox(results_frame, state="disabled")
        self.path_gen_output.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        ctk.CTkButton(results_frame, text="Копировать успешные").grid(row=2, column=0, padx=10, pady=10, sticky="e")

    def _populate_folder_creator_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(tab, text="1. Введите названия папок (каждое с новой строки):").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.folder_creator_input = ctk.CTkTextbox(tab)
        self.folder_creator_input.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")

        options_frame = ctk.CTkFrame(tab)
        options_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        options_frame.grid_columnconfigure((1,3), weight=1)
        ctk.CTkLabel(options_frame, text="Префикс:").grid(row=0, column=0, padx=10, pady=5)
        self.folder_prefix_var = ctk.StringVar()
        ctk.CTkEntry(options_frame, textvariable=self.folder_prefix_var).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkLabel(options_frame, text="Суффикс:").grid(row=0, column=2, padx=10, pady=5)
        self.folder_suffix_var = ctk.StringVar()
        ctk.CTkEntry(options_frame, textvariable=self.folder_suffix_var).grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        
        self.folder_numbering_var = ctk.BooleanVar()
        ctk.CTkCheckBox(options_frame, text="Включить автонумерацию", variable=self.folder_numbering_var).grid(row=1, column=0, padx=10, pady=10, sticky="w")

        self.folder_start_num_var = ctk.IntVar(value=1)
        ctk.CTkEntry(options_frame, textvariable=self.folder_start_num_var, width=60).grid(row=1, column=1, padx=5, pady=5)
        self.folder_padding_var = ctk.IntVar(value=2)
        ctk.CTkEntry(options_frame, textvariable=self.folder_padding_var, width=60).grid(row=1, column=2, padx=5, pady=5)

        ctk.CTkButton(tab, text="Создать папки", command=lambda: self.run_operation("create_folders")).grid(row=3, column=0, padx=10, pady=10, sticky="ew")

    def _populate_article_converter_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(tab, text="1. Выбрать Excel/CSV файл").grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        ctk.CTkLabel(tab, text="Файл не выбран").grid(row=1, column=0, padx=20, pady=5)
        ctk.CTkLabel(tab, text="Найден размер в файле: -", font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, padx=20, pady=10)
        ctk.CTkLabel(tab, text="2. Выберите НОВЫЙ размер для замены:").grid(row=3, column=0, padx=20, pady=10)
        ctk.CTkComboBox(tab, values=[]).grid(row=4, column=0, padx=20, pady=5, sticky="ew")
        ctk.CTkButton(tab, text="3. Создать файл с новым размером", state="disabled").grid(row=5, column=0, padx=20, pady=10, sticky="ew")
        ctk.CTkButton(tab, text="⚙️ Редактор размеров").grid(row=6, column=0, padx=20, pady=20, sticky="ew")

    def _create_bottom_frame(self):
        self.bottom_frame = ctk.CTkFrame(self)
        self.bottom_frame.grid(row=2, column=0, padx=20, pady=(10, 20), sticky="ew")
        self.bottom_frame.grid_columnconfigure(0, weight=1)

        self.log_textbox = ctk.CTkTextbox(self.bottom_frame, state="disabled", height=150)
        self.log_textbox.grid(row=0, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        
        self.progress_bar = ctk.CTkProgressBar(self.bottom_frame)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

        self.stop_btn = ctk.CTkButton(self.bottom_frame, text="Остановить", state="disabled", command=self.stop_operation)
        self.stop_btn.grid(row=1, column=1, padx=10, pady=10)

    # --- Backend and UI Logic ---

    def define_operations(self):
        """Defines all available operations in a structured dictionary."""
        self.operations = {
            "extract": {
                "name": "Извлечь из папок '1'",
                "function": FileOperations.organize_folders,
                "get_args": lambda: (self.path_var.get(), self.logger, self.stop_event, self.update_status, self.dry_run_var.get()),
            },
            "rename_images": {
                "name": "Переименовать изображения 1-N",
                "function": FileOperations.rename_images_sequentially,
                "get_args": lambda: (self.path_var.get(), self.logger, self.stop_event, self.update_status, self.dry_run_var.get()),
            },
            "create_folders": {
                "name": "Создание папок",
                "function": FileOperations.create_folders_from_list,
                "get_args": lambda: (
                    self.path_var.get(), self.folder_creator_input.get("1.0", "end-1c"),
                    self.folder_prefix_var.get(), self.folder_suffix_var.get(),
                    self.folder_numbering_var.get(), self.folder_start_num_var.get(), self.folder_padding_var.get(),
                    self.logger, self.stop_event, self.update_status, self.dry_run_var.get()
                ),
            },
            "generate_paths": {
                "name": "Генерация путей для Excel",
                "function": FileOperations.generate_excel_paths,
                "get_args": lambda: (
                    self.path_var.get(), self.path_gen_input.get("1.0", "end-1c"), self.logger, self.stop_event,
                    self.update_status, self.path_gen_result_callback
                ),
            },
        }

    def load_configuration(self):
        self.path_var.set(self.config.get("last_path", str(Path.home())))
        if self.config.get("theme", "dark") == "dark":
            self.theme_switch.select()
            ctk.set_appearance_mode("dark")
        else:
            self.theme_switch.deselect()
            ctk.set_appearance_mode("light")

    def load_sizes(self):
        if not SIZES_JSON_FILE.exists():
            self.size_to_article_map = DEFAULT_SIZES
            self.save_sizes()
        else:
            try:
                with open(SIZES_JSON_FILE, "r", encoding="utf-8") as f:
                    self.size_to_article_map = json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                self.logger.error("Could not load sizes file, using defaults.")
                self.size_to_article_map = DEFAULT_SIZES

    def save_sizes(self):
        try:
            APP_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            with open(SIZES_JSON_FILE, "w", encoding="utf-8") as f:
                json.dump(self.size_to_article_map, f, ensure_ascii=False, indent=4)
        except Exception as e:
            self.logger.error(f"Не удалось сохранить словарь размеров: {e}")

    def browse_folder(self):
        initial_dir = self.path_var.get() if Path(self.path_var.get()).is_dir() else str(Path.home())
        folder_selected = filedialog.askdirectory(initialdir=initial_dir)
        if folder_selected:
            self.path_var.set(folder_selected)

    def toggle_theme(self):
        mode = "dark" if self.theme_switch.get() == 1 else "light"
        ctk.set_appearance_mode(mode)

    def update_status(self, message: str, progress: Optional[int] = None):
        if progress is not None:
            self.progress_bar.set(progress / 100)
        # TODO: Add a status label if desired

    def stop_operation(self):
        if self.current_thread and self.current_thread.is_alive():
            self.stop_event.set()
            self.logger.warning("--- Попытка остановить операцию... ---")
            self.stop_btn.configure(state="disabled")

    def run_operation(self, op_type: str):
        if self.current_thread and self.current_thread.is_alive():
            ctk.messagebox.showwarning("Операция выполняется", "Другая операция уже запущена.")
            return

        op_details = self.operations.get(op_type)
        if not op_details:
            self.logger.error(f"Неизвестный тип операции: {op_type}")
            return

        def operation_wrapper():
            self.stop_event.clear()
            self.stop_btn.configure(state="normal")
            args = op_details["get_args"]()
            result = op_details["function"](*args)
            if isinstance(result, int):
                self.operation_result_counter = result

            self.stop_btn.configure(state="disabled")
            if not self.stop_event.is_set():
                self.logger.success("🎉 Операция успешно завершена!")

        self.current_thread = threading.Thread(target=operation_wrapper, daemon=True)
        self.current_thread.start()

    def path_gen_result_callback(self, success_str: str, error_str: str):
        self.path_gen_output.configure(state="normal")
        self.path_gen_output.delete("1.0", "end")
        self.path_gen_output.insert("1.0", success_str)
        self.path_gen_output.configure(state="disabled")
        # TODO: Add error output box

    def _on_closing(self):
        # Save config on exit
        self.config["last_path"] = self.path_var.get()
        self.config["theme"] = "dark" if self.theme_switch.get() == 1 else "light"
        ConfigManager.save_config(self.config)
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
