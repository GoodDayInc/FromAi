import os
import shutil
import re
import json
import threading
import datetime
from pathlib import Path
from typing import Dict, Any, Callable, Optional, List

# --- GUI Imports ---
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# --- Third-party Imports ---
import pandas as pd

try:
    import win32com.client as win32

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# --- Constants ---
CONFIG_FILE = "file_organizer_config.json"
SIZES_JSON_FILE = "sizes.json"
IMAGE_EXTENSIONS = [
    ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".webp", ".svg", ".ico"
]
DEFAULT_SIZES = {
    "41 р": 1211561, "41.5 р": 1211562, "42 р": 1211563, "42.5 р": 1211564,
    "43 р": 1211565, "43.5 р": 1211566, "44 р": 1211567, "44.5 р": 1211568,
}

# --- Helper Classes ---

class ConfigManager:
    """Handles loading and saving of the application configuration."""

    @staticmethod
    def load_config() -> Dict[str, Any]:
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                print(f"Error loading config file or file is corrupted: {CONFIG_FILE}")
                return {}
        return {}

    @staticmethod
    def save_config(config: Dict[str, Any]) -> None:
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Error saving config: {e}")


class Logger:
    """Manages logging to the GUI's text widget."""

    def __init__(self, output_widget: scrolledtext.ScrolledText):
        self.output_widget = output_widget

    def log(self, message: str, level: str = "info") -> None:
        if not self.output_widget or not self.output_widget.winfo_exists():
            return
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        try:
            self.output_widget.configure(state="normal")
            self.output_widget.insert(tk.END, formatted_message, level)
            self.output_widget.configure(state="disabled")
            self.output_widget.see(tk.END)
        except tk.TclError:
            # This can happen if the widget is destroyed while logging.
            pass

    def info(self, message: str) -> None:
        self.log(message, "info")

    def success(self, message: str) -> None:
        self.log(f"✓ {message}", "success")

    def warning(self, message: str) -> None:
        self.log(f"⚠ {message}", "warning")

    def error(self, message: str) -> None:
        self.log(f"✗ {message}", "error")


class ModernTooltip:
    """A modern, theme-aware tooltip for tkinter widgets."""

    def __init__(
        self,
        widget,
        text: str,
        delay: int = 500,
        app_themes: Optional[Dict] = None,
        current_theme_name_getter: Optional[Callable] = None,
    ):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.app_themes = app_themes
        self.current_theme_name_getter = current_theme_name_getter
        self.tooltip_window = None
        self.id = None
        self.widget.bind("<Enter>", self.on_enter)
        self.widget.bind("<Leave>", self.on_leave)
        self.widget.bind("<ButtonPress>", self.on_leave)

    def on_enter(self, event=None):
        self.schedule_tooltip()

    def on_leave(self, event=None):
        self.hide_tooltip()

    def schedule_tooltip(self):
        self.hide_tooltip()
        self.id = self.widget.after(self.delay, self.show_tooltip)

    def show_tooltip(self):
        if self.tooltip_window or not self.text:
            return

        # Default theme settings
        bg_color, fg_color = "#2c3e50", "#ecf0f1"
        font_style = ("Segoe UI", 9)

        # Get theme-specific settings if available
        if self.app_themes and self.current_theme_name_getter:
            try:
                current_theme_name = self.current_theme_name_getter()
                theme_settings = self.app_themes.get(current_theme_name, {})
                bg_color = theme_settings.get("tooltip_bg", bg_color)
                fg_color = theme_settings.get("tooltip_fg", fg_color)
            except Exception:
                pass  # Ignore errors and use defaults

        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.configure(bg=bg_color)

        label = tk.Label(
            self.tooltip_window,
            text=self.text,
            background=bg_color,
            foreground=fg_color,
            font=font_style,
            relief="solid",
            borderwidth=1,
            padx=10,
            pady=6,
        )
        label.pack()

        self.tooltip_window.update_idletasks()

        # Position the tooltip
        widget_x = self.widget.winfo_rootx()
        widget_y = self.widget.winfo_rooty()
        widget_height = self.widget.winfo_height()
        tip_width = self.tooltip_window.winfo_width()
        tip_height = self.tooltip_window.winfo_height()

        x = widget_x + self.widget.winfo_width() // 2 - tip_width // 2
        y = widget_y + widget_height + 10

        # Adjust position to stay on screen
        screen_width = self.widget.winfo_screenwidth()
        screen_height = self.widget.winfo_screenheight()
        if x + tip_width > screen_width:
            x = screen_width - tip_width - 10
        if x < 0:
            x = 10
        if y + tip_height > screen_height:
            y = widget_y - tip_height - 10
        if y < 0:
            y = 10

        self.tooltip_window.wm_geometry(f"+{int(x)}+{int(y)}")

    def hide_tooltip(self):
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None


class PlaceholderEntry(ttk.Entry):
    """An entry widget that shows placeholder text."""

    def __init__(self, master=None, placeholder="", **kwargs):
        super().__init__(master, **kwargs)
        self.placeholder = placeholder
        self.bind("<FocusIn>", self._clear_placeholder)
        self.bind("<FocusOut>", self._add_placeholder)
        self._add_placeholder()

    def _clear_placeholder(self, e):
        if self.get() == self.placeholder:
            self.delete(0, tk.END)
            self.configure(style="Active.TEntry")

    def _add_placeholder(self, e=None):
        if not self.get():
            self.insert(0, self.placeholder)
            self.configure(style="Inactive.TEntry")


# --- File Operations ---

class FileOperations:
    """A collection of static methods for performing file system operations."""

    @staticmethod
    def _sanitize_folder_path(path: str) -> str:
        """Removes illegal characters from a path segment."""
        parts = re.split(r"[\\/]", path)
        sanitized_parts = [re.sub(r'[?:"<>|*]', "", part).strip() for part in parts]
        return os.path.join(*[p for p in sanitized_parts if p])

    @staticmethod
    def create_folders_from_list(
        base_path: str,
        folder_list_str: str,
        prefix: str,
        suffix: str,
        use_numbering: bool,
        start_num: int,
        padding: int,
        logger: Logger,
        stop_event: threading.Event,
        status_callback: Callable[[str, int], None],
        dry_run: bool = False,
    ) -> int:
        op_name = "[DRY RUN] " if dry_run else ""
        logger.info(f"{op_name}🏗️ Начинаем операцию: Создание папок")
        if not dry_run:
            logger.warning("⚠️ ВНИМАНИЕ: Операция создает папки на диске!")
        logger.info(f"Целевая директория: {base_path}")

        folder_names = [
            name.strip() for name in folder_list_str.strip().split("\n") if name.strip()
        ]
        if not folder_names:
            logger.warning("Список папок для создания пуст. Операция прервана.")
            status_callback("Список папок пуст.", 0)
            return 0

        total_folders = len(folder_names)
        created_count = 0

        try:
            for i, name in enumerate(folder_names):
                if stop_event.is_set():
                    logger.warning("Операция прервана пользователем.")
                    status_callback("Операция прервана.", 0)
                    break

                progress = int((i + 1) / total_folders * 100)
                status_callback(f"{op_name}Создание: {name}", progress)

                sanitized_path = FileOperations._sanitize_folder_path(name)
                if not sanitized_path:
                    logger.warning(f"Пропущено: имя '{name}' стало пустым после очистки.")
                    continue

                number_str = str(i + start_num).zfill(padding) + "_" if use_numbering else ""
                
                path_parts = list(os.path.split(sanitized_path))
                path_parts[-1] = f"{prefix}{number_str}{path_parts[-1]}{suffix}"
                final_name = os.path.join(*path_parts)
                
                full_path = os.path.join(base_path, final_name)

                try:
                    if not dry_run:
                        os.makedirs(full_path, exist_ok=True)
                    logger.success(f"{op_name}Создана папка: '{final_name}'")
                    created_count += 1
                except OSError as e:
                    logger.error(f"Ошибка создания папки '{final_name}': {e}")
            
            if created_count > 0:
                logger.success(f"✅ {op_name}Операция 'Создание папок' завершена! Всего создано: {created_count} папок.")
            else:
                logger.warning(f"{op_name}Ни одной папки не было создано.")

        except Exception as e:
            logger.error(f"Критическая ошибка во время создания папок: {e}")
            status_callback("Ошибка!", 0)
        finally:
            if not stop_event.is_set():
                status_callback("Готово.", 100)
            logger.info(f"--- {op_name}Операция 'Создание папок' завершена ---")

        return created_count

    @staticmethod
    def generate_excel_paths(
        base_path: str,
        model_list_str: str,
        logger: Logger,
        stop_event: threading.Event,
        status_callback: Callable[[str, int], None],
        result_callback: Callable[[str, str], None],
    ):
        logger.info("📋 Начинаем операцию: Генерация путей для Excel (Гибкий поиск)")
        logger.info(f"Используется базовый путь: {base_path}")

        def natural_sort_key(s: str):
            return [
                int(text) if text.isdigit() else text.lower()
                for text in re.split("([0-9]+)", os.path.basename(s))
            ]

        model_list = [name.strip() for name in model_list_str.strip().split("\n") if name.strip()]
        if not model_list:
            logger.warning("Список моделей пуст. Операция прервана.")
            status_callback("Список моделей пуст.", 0)
            result_callback("", "")
            return

        total_models = len(model_list)
        success_output, error_output = [], []
        try:
            for i, model_name in enumerate(model_list):
                if stop_event.is_set():
                    logger.warning("Операция прервана пользователем.")
                    status_callback("Операция прервана.", 0)
                    break
                
                progress = int((i + 1) / total_models * 100)
                status_callback(f"Проверка: {model_name}", progress)
                model_path = os.path.join(base_path, model_name)

                if not os.path.isdir(model_path):
                    error_output.append(f"{model_name} -> ОШИБКА: Папка не найдена!")
                    logger.error(f"Папка для модели '{model_name}' не найдена по пути: {model_path}")
                    continue
                
                photo_paths = []
                try:
                    for filename in os.listdir(model_path):
                        if os.path.splitext(filename)[1].lower() in IMAGE_EXTENSIONS:
                            full_path = os.path.join(model_path, filename)
                            if os.path.isfile(full_path):
                                photo_paths.append(full_path)
                except OSError as e:
                    error_output.append(f"{model_name} -> ОШИБКА: Не удалось прочитать папку: {e}")
                    logger.error(f"Не удалось прочитать папку для '{model_name}': {e}")
                    continue

                if photo_paths:
                    sorted_paths = sorted(photo_paths, key=natural_sort_key)
                    joined_paths = "|".join(sorted_paths)
                    final_string = f'"[+\n+{joined_paths}]"'
                    success_output.append(final_string)
                    logger.success(f"Пути для '{model_name}' ({len(sorted_paths)} фото) успешно сгенерированы.")
                else:
                    error_output.append(f"{model_name} -> ОШИБКА: Изображения не найдены в папке.")
                    logger.warning(f"Для модели '{model_name}' не найдены изображения в папке.")
            
            result_callback("\n".join(success_output), "\n".join(error_output))
            
            if not stop_event.is_set():
                logger.success("✅ Операция 'Генерация путей' завершена!")
                status_callback("Готово.", 100)
        except Exception as e:
            logger.error(f"Критическая ошибка во время генерации путей: {e}")
            status_callback("Ошибка!", 0)
        finally:
            logger.info("--- Операция 'Генерация путей для Excel' завершена ---")

    @staticmethod
    def organize_folders(
        root_path: str,
        logger: Logger,
        stop_event: threading.Event,
        status_callback: Callable[[str, int], None],
        dry_run: bool = False,
    ) -> int:
        op_name = "[DRY RUN] " if dry_run else ""
        logger.info(f"{op_name}🚀 Начинаем операцию: Извлечение из папок '1'")
        if not dry_run:
            logger.warning("⚠️ ВНИМАНИЕ: Операция изменяет структуру файлов!")
        logger.info(f"Целевая директория: {root_path}")
        
        processed_count = 0
        found_folders = False
        try:
            all_dirs = [dp for dp, dn, _ in os.walk(root_path) if "1" in dn]
            total_dirs = len(all_dirs)
            
            for i, dirpath in enumerate(sorted(all_dirs, key=lambda x: x.count(os.sep), reverse=True)):
                if stop_event.is_set():
                    logger.warning("Операция прервана пользователем.")
                    status_callback("Операция прервана.", 0)
                    return processed_count

                progress = int((i + 1) / total_dirs * 100) if total_dirs > 0 else 0
                folder_1_path = os.path.join(dirpath, "1")
                parent_path = dirpath
                
                logger.info(f"{op_name}📁 Найдена папка: {folder_1_path}")
                status_callback(f"{op_name}Обработка: {os.path.relpath(folder_1_path, root_path)}", progress)
                found_folders = True

                try:
                    items_in_folder_1 = os.listdir(folder_1_path)
                except OSError as e:
                    logger.error(f"Ошибка чтения содержимого {folder_1_path}: {e}")
                    continue

                if items_in_folder_1:
                    logger.info(f"{op_name}Перемещение {len(items_in_folder_1)} элементов из '{folder_1_path}' в '{parent_path}'...")
                    for item_name in items_in_folder_1:
                        if stop_event.is_set():
                            logger.warning("Операция прервана во время перемещения.")
                            status_callback("Операция прервана.", 0)
                            return processed_count
                        
                        src_item_path = os.path.join(folder_1_path, item_name)
                        dst_item_path = os.path.join(parent_path, item_name)
                        
                        if os.path.exists(dst_item_path):
                            logger.warning(f"Конфликт: Файл/папка '{item_name}' уже существует в '{parent_path}'. Пропуск.")
                            continue
                        try:
                            if not dry_run:
                                shutil.move(src_item_path, dst_item_path)
                            logger.success(f"{op_name}Перемещен: '{item_name}'")
                        except OSError as e:
                            logger.error(f"Ошибка перемещения '{item_name}': {e}")

                try:
                    if not os.listdir(folder_1_path):
                        if not dry_run:
                            os.rmdir(folder_1_path)
                        logger.success(f"{op_name}Удалена пустая папка: {folder_1_path}")
                        processed_count += 1
                    else:
                        logger.warning(f"Папка '{folder_1_path}' не пуста после попытки перемещения. Не удалена.")
                except OSError as e:
                    logger.error(f"Ошибка удаления папки '{folder_1_path}': {e}")

            if not found_folders:
                logger.warning("Папки с именем '1' не найдены в указанной директории и ее подпапках.")
            elif processed_count > 0:
                logger.success(f"✅ {op_name}Операция 'Извлечь из папок 1' завершена! Обработано и удалено папок '1': {processed_count}.")
            else:
                logger.warning(f"{op_name}Папки '1' были найдены, но ни одна не была удалена (возможно, из-за ошибок или конфликтов).")
        
        except Exception as e:
            logger.error(f"Критическая ошибка во время операции 'Извлечь из папок 1': {e}")
            status_callback("Ошибка!", 0)
        finally:
            if not stop_event.is_set():
                status_callback("Готово.", 100)
            logger.info(f"--- {op_name}Операция 'Извлечь из папок 1' завершена ---")
        return processed_count

    @staticmethod
    def rename_images_sequentially(
        directory: str,
        logger: Logger,
        stop_event: threading.Event,
        status_callback: Callable[[str, int], None],
        dry_run: bool = False,
    ) -> int:
        op_name = "[DRY RUN] " if dry_run else ""
        logger.info(f"{op_name}🔢 Начинаем операцию: Переименование изображений (1-N)")
        if not dry_run:
            logger.warning("⚠️ ВНИМАНИЕ: Операция изменяет имена файлов!")
        logger.info(f"Целевая директория: {directory}")
        
        total_renamed_files = 0
        processed_folders = 0
        try:
            subdirs = [
                root for root, _, files in os.walk(directory)
                if any(os.path.splitext(f)[1].lower() in IMAGE_EXTENSIONS for f in files)
            ]
            total_dirs = len(subdirs)
            
            for i, root in enumerate(subdirs):
                if stop_event.is_set():
                    logger.warning("Операция прервана пользователем.")
                    status_callback("Операция прервана.", 0)
                    return total_renamed_files
                
                progress = int((i + 1) / total_dirs * 100) if total_dirs > 0 else 0
                image_files = sorted([
                    f for f in os.listdir(root)
                    if os.path.isfile(os.path.join(root, f)) and os.path.splitext(f)[1].lower() in IMAGE_EXTENSIONS
                ])
                
                if not image_files:
                    continue

                processed_folders += 1
                rel_root = os.path.relpath(root, directory) or "."
                logger.info(f"{op_name}📂 Обрабатываем папку: {rel_root}")
                logger.info(f"Найдено изображений: {len(image_files)}")
                status_callback(f"{op_name}Обработка: {rel_root}", progress)
                
                renamed_in_folder = 0
                for index, filename in enumerate(image_files, 1):
                    if stop_event.is_set():
                        logger.warning("Операция прервана во время переименования.")
                        status_callback("Операция прервана.", 0)
                        return total_renamed_files

                    old_path = os.path.join(root, filename)
                    extension = os.path.splitext(filename)[1].lower()
                    new_filename = f"{index}{extension}"
                    new_path = os.path.join(root, new_filename)

                    if old_path == new_path:
                        logger.info(f"Файл '{filename}' уже имеет целевое имя. Пропуск.")
                        continue

                    if os.path.exists(new_path):
                        # Handle conflicts by adding a suffix
                        base, ext = os.path.splitext(new_filename)
                        conflict_count = 1
                        while os.path.exists(new_path):
                            new_filename_conflict = f"{base}_conflict_{conflict_count}{ext}"
                            new_path = os.path.join(root, new_filename_conflict)
                            conflict_count += 1
                            if conflict_count > 100:
                                logger.error(f"Слишком много конфликтов для {new_filename}. Пропуск {filename}.")
                                new_path = None
                                break
                        if new_path is None:
                            continue
                        logger.warning(f"Конфликт для '{new_filename}'. Переименовываю в '{os.path.basename(new_path)}'.")

                    try:
                        if not dry_run:
                            os.rename(old_path, new_path)
                        logger.success(f"{op_name}Переименован: '{filename}' → '{os.path.basename(new_path)}'")
                        total_renamed_files += 1
                        renamed_in_folder += 1
                    except OSError as e:
                        logger.error(f"Ошибка переименования '{filename}': {e}")
                
                logger.info(f"Переименовано в папке: {renamed_in_folder} файлов.")

            if processed_folders == 0:
                logger.warning("Изображения для переименования не найдены.")
            else:
                logger.success(f"✅ {op_name}Операция 'Переименовать изображения' завершена! Всего переименовано: {total_renamed_files} файлов в {processed_folders} папках.")
        
        except Exception as e:
            logger.error(f"Критическая ошибка во время операции 'Переименовать изображения': {e}")
            status_callback("Ошибка!", 0)
        finally:
            if not stop_event.is_set():
                status_callback("Готово.", 100)
            logger.info(f"--- {op_name}Операция 'Переименовать изображения' завершена ---")
        return total_renamed_files

    @staticmethod
    def remove_phrase_from_names(
        base_path_str: str,
        phrase: str,
        logger: Logger,
        stop_event: threading.Event,
        status_callback: Callable[[str, int], None],
        case_sensitive: bool,
        use_regex: bool,
        dry_run: bool = False,
    ) -> int:
        op_name = "[DRY RUN] " if dry_run else ""
        logger.info(f"{op_name}✂️ Начинаем операцию: Удаление фразы/шаблона '{phrase}'")
        if not dry_run:
            logger.warning("⚠️ ВНИМАНИЕ: Операция изменяет имена файлов и папок!")
        logger.info(f"Целевая директория: {base_path_str}")
        logger.info(f"Учитывать регистр: {'Да' if case_sensitive else 'Нет'} | Использовать RegEx: {'Да' if use_regex else 'Нет'}")

        if not phrase:
            logger.error("Фраза/шаблон для удаления не может быть пустой.")
            status_callback("Ошибка: Пустая фраза.", 0)
            return 0

        try:
            pattern = re.compile(phrase, 0 if case_sensitive else re.IGNORECASE) if use_regex else None
        except re.error as e:
            logger.error(f"Некорректное регулярное выражение: {e}")
            status_callback("Ошибка: некорректный RegEx!", 0)
            return 0

        processed_count = 0
        base_path = Path(base_path_str)
        try:
            # Get all items and sort by depth (deepest first) to avoid renaming parent before child
            items_to_process = sorted(list(base_path.rglob("*")), key=lambda p: len(str(p)), reverse=True)
            total_items = len(items_to_process)
            
            for i, item_path in enumerate(items_to_process):
                if stop_event.is_set():
                    logger.warning("Операция прервана пользователем.")
                    status_callback("Операция прервана.", 0)
                    return processed_count
                
                progress = int((i + 1) / total_items * 100) if total_items > 0 else 0
                status_callback(f"Проверка: {item_path.name}", progress)
                
                original_name = item_path.name
                
                if use_regex:
                    target_name_candidate = pattern.sub("", original_name).strip()
                else:
                    # Simple string replacement
                    if case_sensitive:
                        target_name_candidate = original_name.replace(phrase, "").strip()
                    else:
                        # Case-insensitive replacement
                        target_name_candidate = re.sub(re.escape(phrase), "", original_name, flags=re.IGNORECASE).strip()

                if not target_name_candidate:
                    if item_path.is_file() and item_path.suffix:
                        target_name_candidate = f"renamed_file{item_path.suffix}"
                        logger.warning(f"Имя файла '{original_name}' стало бы пустым. Будет '{target_name_candidate}'.")
                    elif item_path.is_dir():
                        target_name_candidate = "renamed_folder"
                        logger.warning(f"Имя папки '{original_name}' стало бы пустым. Будет '{target_name_candidate}'.")
                    else:
                        logger.warning(f"Пропуск: Имя '{original_name}' стало бы пустым после удаления.")
                        continue
                
                if target_name_candidate == original_name:
                    continue

                new_path = item_path.parent / target_name_candidate
                
                # Skip if a file/folder with the new name already exists
                if new_path.exists():
                    logger.warning(f"Конфликт: '{new_path}' уже существует. Пропуск переименования '{original_name}'.")
                    continue
                
                try:
                    if not dry_run:
                        item_path.rename(new_path)
                    logger.success(f"{op_name}Переименовано: '{original_name}' → '{target_name_candidate}'")
                    processed_count += 1
                except Exception as e:
                    logger.error(f"Ошибка переименования '{original_name}' в '{target_name_candidate}': {e}")

            if processed_count == 0:
                logger.warning("Фраза/шаблон не найдена ни в одном имени файла или папки.")
            else:
                logger.success(f"✅ {op_name}Операция 'Удалить фразу' завершена! Всего переименовано элементов: {processed_count}.")
        
        except Exception as e:
            logger.error(f"Критическая ошибка во время операции 'Удалить фразу': {e}")
            status_callback("Ошибка!", 0)
        finally:
            if not stop_event.is_set():
                status_callback("Готово.", 100)
            logger.info(f"--- {op_name}Операция 'Удалить фразу: {phrase}' завершена ---")
        return processed_count

    @staticmethod
    def delete_url_shortcuts(
        base_path_str: str,
        names_to_delete_str: str,
        logger: Logger,
        stop_event: threading.Event,
        status_callback: Callable[[str, int], None],
        case_sensitive: bool = False,
        dry_run: bool = False,
    ) -> int:
        op_name = "[DRY RUN] " if dry_run else ""
        logger.info(f"{op_name}🗑️ Начинаем операцию: Удаление URL-ярлыков")
        if not dry_run:
            logger.warning("⚠️ ВНИМАНИЕ: Операция удаляет файлы!")
        logger.info(f"Целевая директория: {base_path_str}")
        logger.info(f"Имена/части имен для удаления: '{names_to_delete_str}'")
        logger.info(f"Учитывать регистр: {'Да' if case_sensitive else 'Нет'}")

        names_list_raw = [name.strip() for name in names_to_delete_str.split(",") if name.strip()]
        if not names_list_raw:
            logger.warning("Не указаны имена или части имен ярлыков для удаления.")
            status_callback("Предупреждение: имена не указаны.", 0)
            return 0

        names_list = names_list_raw if case_sensitive else [name.lower() for name in names_list_raw]
        deleted_count = 0
        base_path = Path(base_path_str)
        
        try:
            url_files = list(base_path.rglob("*.url"))
            total_files = len(url_files)
            
            for i, file_path in enumerate(url_files):
                if stop_event.is_set():
                    logger.warning("Операция прервана пользователем.")
                    status_callback("Операция прервана.", 0)
                    return deleted_count
                
                progress = int((i + 1) / total_files * 100) if total_files > 0 else 0
                status_callback(f"Проверка: {file_path.name}", progress)

                file_name_to_check = file_path.stem if case_sensitive else file_path.stem.lower()
                
                if any(target_name in file_name_to_check for target_name in names_list):
                    try:
                        if not dry_run:
                            file_path.unlink()
                        logger.success(f"{op_name}Удален ярлык: '{file_path}'")
                        deleted_count += 1
                    except OSError as e:
                        logger.error(f"Ошибка удаления ярлыка '{file_path}': {e}")

            if deleted_count == 0:
                logger.warning("Интернет-ярлыки с указанными именами не найдены.")
            else:
                logger.success(f"✅ {op_name}Операция 'Удалить URL-ярлыки' завершена! Всего удалено ярлыков: {deleted_count}.")
        
        except Exception as e:
            logger.error(f"Критическая ошибка во время операции 'Удалить URL-ярлыки': {e}")
            status_callback("Ошибка!", 0)
        finally:
            if not stop_event.is_set():
                status_callback("Готово.", 100)
            logger.info(f"--- {op_name}Операция 'Удалить URL-ярлыки по именам: {names_to_delete_str}' завершена ---")
        return deleted_count


# --- GUI Components ---

class SizeEditor(tk.Toplevel):
    """A Toplevel window for editing the size-to-article mapping."""

    def __init__(self, master, controller):
        super().__init__(master)
        self.controller = controller
        self.title("Редактор размеров")
        self.geometry("450x400")
        self.transient(master)
        self.grab_set()

        self.tree = ttk.Treeview(self, columns=("Размер", "Артикул"), show="headings")
        self.tree.heading("Размер", text="Размер")
        self.tree.heading("Артикул", text="Артикул")
        self.tree.pack(pady=10, padx=10, fill="both", expand=True)
        
        entry_frame = ttk.Frame(self)
        entry_frame.pack(padx=10, pady=5, fill="x")
        ttk.Label(entry_frame, text="Размер:").pack(side="left", padx=(0, 5))
        self.size_entry = ttk.Entry(entry_frame)
        self.size_entry.pack(side="left", expand=True, fill="x")
        ttk.Label(entry_frame, text="Артикул:").pack(side="left", padx=(10, 5))
        self.article_entry = ttk.Entry(entry_frame)
        self.article_entry.pack(side="left", expand=True, fill="x")
        
        btn_frame = ttk.Frame(self)
        btn_frame.pack(padx=10, pady=10)
        ttk.Button(btn_frame, text="Добавить/Обновить", command=self.add_or_update).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Удалить выбранное", command=self.delete_selected).pack(side="left", padx=5)
        
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.populate_tree()

    def populate_tree(self):
        """Fills the treeview with the current size data."""
        self.tree.delete(*self.tree.get_children())
        for size, article in self.controller.size_to_article_map.items():
            self.tree.insert("", "end", values=(size, article))
    
    def on_select(self, event):
        """Populates entry fields when a tree item is selected."""
        if not self.tree.selection():
            return
        selected_item = self.tree.selection()[0]
        size, article = self.tree.item(selected_item, "values")
        self.size_entry.delete(0, "end")
        self.size_entry.insert(0, size)
        self.article_entry.delete(0, "end")
        self.article_entry.insert(0, article)

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
        self.populate_tree()
        self.size_entry.delete(0, "end")
        self.article_entry.delete(0, "end")

    def delete_selected(self):
        """Deletes the selected size-article pair."""
        if not self.tree.selection():
            messagebox.showwarning("Ошибка", "Сначала выберите строку для удаления.", parent=self)
            return
        
        if messagebox.askyesno("Подтверждение", "Вы уверены?", parent=self):
            selected_item = self.tree.selection()[0]
            size, _ = self.tree.item(selected_item, "values")
            del self.controller.size_to_article_map[size]
            self.controller.save_sizes()
            self.populate_tree()

    def on_close(self):
        """Updates the main app before closing."""
        self.controller.update_converter_combobox()
        self.destroy()


# --- Main Application ---

class ModernFileOrganizerApp:
    def __init__(self, master: tk.Tk):
        self.master = master
        self.current_thread: Optional[threading.Thread] = None
        self.stop_event = threading.Event()
        self.operation_result_counter = 0
        self.current_operation_is_path_gen = False
        self.operation_buttons: List[ttk.Button] = []

        self.setup_window()
        self.load_configuration()
        self.define_operations()
        self.setup_themes()
        self.setup_styles()
        self.create_widgets()
        self.apply_theme(self.current_theme_name)
        self.setup_bindings()
        
        self.master.after(100, self.show_welcome_message)

    def setup_window(self):
        self.master.title("🗂️ Супер Скрипт v2.7")
        self.master.geometry("1100x800")
        self.master.minsize(900, 700)
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_rowconfigure(1, weight=1)

    def load_configuration(self):
        self.config = ConfigManager.load_config()
        self.last_path = self.config.get("last_path", os.path.expanduser("~"))
        self.current_theme_name = self.config.get("theme", "light")
        # Load other last-used settings
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

    def define_operations(self):
        """Defines all available operations in a structured dictionary."""
        self.operations = {
            "extract": {
                "name": "Извлечь из папок '1'",
                "function": FileOperations.organize_folders,
                "get_args": lambda: (self.path_var.get(), self.logger, self.stop_event, self.update_status, self.dry_run_var.get()),
                "is_file_op": True,
            },
            "rename_images": {
                "name": "Переименовать изображения 1-N",
                "function": FileOperations.rename_images_sequentially,
                "get_args": lambda: (self.path_var.get(), self.logger, self.stop_event, self.update_status, self.dry_run_var.get()),
                "is_file_op": True,
            },
            "remove_phrase": {
                "name": "Удалить фразу/RegEx из имен",
                "function": FileOperations.remove_phrase_from_names,
                "get_args": lambda: (
                    self.path_var.get(), self.phrase_var.get(), self.logger, self.stop_event, 
                    self.update_status, self.case_sensitive_phrase_var.get(), self.use_regex_var.get(), self.dry_run_var.get()
                ),
                "pre_check": lambda: self.phrase_var.get(),
                "pre_check_msg": "Пожалуйста, введите фразу или RegEx для удаления.",
                "is_file_op": True,
            },
            "delete_urls": {
                "name": "Удалить URL-ярлыки",
                "function": FileOperations.delete_url_shortcuts,
                "get_args": lambda: (
                    self.path_var.get(), self.url_names_var.get(), self.logger, self.stop_event, 
                    self.update_status, self.case_sensitive_url_var.get(), self.dry_run_var.get()
                ),
                "pre_check": lambda: self.url_names_var.get().strip(),
                "pre_check_msg": "Пожалуйста, введите имена URL-ярлыков.",
                "is_file_op": True,
            },
            "generate_paths": {
                "name": "Генерация путей для Excel",
                "function": FileOperations.generate_excel_paths,
                "get_args": lambda: (
                    self.path_var.get(), self.path_gen_input_text.get("1.0", tk.END), self.logger, self.stop_event, 
                    self.update_status, self.path_gen_result_callback
                ),
                "pre_check": lambda: self.path_gen_input_text.get("1.0", tk.END).strip(),
                "pre_check_msg": "Пожалуйста, введите список моделей для генерации.",
                "is_file_op": False,
            },
            "create_folders": {
                "name": "Создание папок",
                "function": FileOperations.create_folders_from_list,
                "get_args": lambda: (
                    self.path_var.get(), self.folder_creator_input_text.get("1.0", tk.END),
                    self.folder_prefix_var.get(), self.folder_suffix_var.get(),
                    self.folder_numbering_var.get(), self.folder_start_num_var.get(), self.folder_padding_var.get(),
                    self.logger, self.stop_event, self.update_status, self.dry_run_var.get()
                ),
                "pre_check": lambda: self.folder_creator_input_text.get("1.0", tk.END).strip(),
                "pre_check_msg": "Пожалуйста, введите названия папок для создания.",
                "is_file_op": True,
            },
        }

    def setup_themes(self):
        self.themes = {
            "light": {
                "bg": "#F5F5F5", "fg": "#1E1E1E", "accent": "#0078D7", "accent_fg": "#FFFFFF",
                "secondary_bg": "#FFFFFF", "border": "#BDBDBD", "hover": "#005A9E",
                "disabled_bg": "#E0E0E0", "disabled_fg": "#A0A0A0",
                "log_bg": "#FFFFFF", "log_fg": "#1E1E1E", "log_info": "#0078D7",
                "log_success": "#107C10", "log_warning": "#FF8C00", "log_error": "#D83B01",
                "tooltip_bg": "#FFFFE0", "tooltip_fg": "#000000",
                "button_danger_bg": "#D83B01", "button_danger_fg": "#FFFFFF", "button_danger_hover": "#A82F00",
                "header_bg": "#E1E1E1", "footer_bg": "#E1E1E1",
                "labelframe_label_fg": "#005A9E", "progress_bar": "#0078D7", "entry_bg": "#F8F8F8",
            },
            "dark": {
                "bg": "#1E1E1E", "fg": "#F0F0F0", "accent": "#0078D7", "accent_fg": "#FFFFFF",
                "secondary_bg": "#2D2D2D", "border": "#505050", "hover": "#005A9E",
                "disabled_bg": "#3C3C3C", "disabled_fg": "#707070",
                "log_bg": "#252526", "log_fg": "#CCCCCC", "log_info": "#569CD6",
                "log_success": "#4EC9B0", "log_warning": "#FFCC00", "log_error": "#F44747",
                "tooltip_bg": "#3C3C3C", "tooltip_fg": "#F0F0F0",
                "button_danger_bg": "#F44747", "button_danger_fg": "#000000", "button_danger_hover": "#D33636",
                "header_bg": "#2D2D2D", "footer_bg": "#2D2D2D",
                "labelframe_label_fg": "#569CD6", "progress_bar": "#0078D7", "entry_bg": "#2A2A2A",
            },
        }

    def setup_styles(self):
        self.style = ttk.Style(self.master)
        try:
            current_themes = self.style.theme_names()
            if "clam" in current_themes:
                self.style.theme_use("clam")
        except tk.TclError:
            pass

    def create_widgets(self):
        self.create_header()
        self.create_main_content()
        self.create_footer()

    def create_header(self):
        theme = self.themes[self.current_theme_name]
        self.header_frame = tk.Frame(self.master, height=80, bg=theme.get("header_bg", theme["bg"]))
        self.header_frame.grid(row=0, column=0, sticky="ew")
        self.header_frame.grid_propagate(False)

        title_container = tk.Frame(self.header_frame, bg=self.header_frame.cget("bg"))
        title_container.pack(side="left", padx=20, pady=10)
        
        self.title_label = tk.Label(title_container, text="🗂️ Супер Скрипт", font=("Segoe UI", 22, "bold"), bg=self.header_frame.cget("bg"))
        self.title_label.pack(side="top", anchor="w")
        
        self.subtitle_label = tk.Label(title_container, text="Продвинутый инструмент для пакетной работы с файлами", font=("Segoe UI", 11), bg=self.header_frame.cget("bg"))
        self.subtitle_label.pack(side="top", anchor="w")
        
        controls_container = tk.Frame(self.header_frame, bg=self.header_frame.cget("bg"))
        controls_container.pack(side="right", padx=20, pady=10)

        self.theme_btn = ttk.Button(controls_container, text="Тема", command=self.toggle_theme, width=10, style="Header.TButton")
        self.theme_btn.pack(side="left", padx=(0, 10))
        ModernTooltip(self.theme_btn, "Переключить тему (Светлая/Темная)", app_themes=self.themes, current_theme_name_getter=self.get_current_theme_name)
        
        self.help_btn = ttk.Button(controls_container, text="❓ Справка", command=self.show_help, width=10, style="Header.TButton")
        self.help_btn.pack(side="left")
        ModernTooltip(self.help_btn, "Показать информацию о программе и операциях", app_themes=self.themes, current_theme_name_getter=self.get_current_theme_name)

    def create_main_content(self):
        self.main_pane = ttk.PanedWindow(self.master, orient="vertical")
        self.main_pane.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)

        # Top pane for controls
        top_frame = ttk.Frame(self.main_pane, style="TFrame")
        self.main_pane.add(top_frame, weight=3)
        top_frame.grid_columnconfigure(0, weight=1)
        top_frame.grid_rowconfigure(1, weight=1)
        
        # Bottom pane for logs
        bottom_frame = ttk.Frame(self.main_pane, style="TFrame")
        self.main_pane.add(bottom_frame, weight=2)
        bottom_frame.grid_columnconfigure(0, weight=1)
        bottom_frame.grid_rowconfigure(0, weight=1)

        self.create_path_panel(top_frame)
        self.create_notebook_panel(top_frame)
        self.create_log_panel(bottom_frame)
        self.create_progress_panel(bottom_frame)

    def create_path_panel(self, parent: tk.Frame):
        path_lf = ttk.LabelFrame(parent, text="📍 Общая рабочая папка", style="Controls.TLabelframe")
        path_lf.grid(row=0, column=0, sticky="new", pady=(0, 10))
        path_lf.grid_columnconfigure(1, weight=1)

        ttk.Label(path_lf, text="Рабочая папка:", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", padx=(10, 5), pady=10)
        
        self.path_var = tk.StringVar(value=self.last_path)
        self.path_entry = PlaceholderEntry(path_lf, textvariable=self.path_var, placeholder="Введите или выберите путь...")
        self.path_entry.grid(row=0, column=1, sticky="ew", padx=(0, 5), pady=10)
        ModernTooltip(self.path_entry, "Путь к папке, которую нужно обработать. Используется для всех операций.", app_themes=self.themes, current_theme_name_getter=self.get_current_theme_name)
        
        self.browse_btn = ttk.Button(path_lf, text="Обзор...", command=self.browse_folder, style="Accent.TButton")
        self.browse_btn.grid(row=0, column=2, sticky="ew", padx=(0, 10), pady=10)
        ModernTooltip(self.browse_btn, "Открыть диалог выбора папки.", app_themes=self.themes, current_theme_name_getter=self.get_current_theme_name)

    def create_notebook_panel(self, parent: tk.Frame):
        notebook = ttk.Notebook(parent, style="TNotebook")
        notebook.grid(row=1, column=0, sticky="nsew")
        
        tab_file_ops = ttk.Frame(notebook, style="TFrame", padding=15)
        tab_path_gen = ttk.Frame(notebook, style="TFrame", padding=15)
        tab_folder_creator = ttk.Frame(notebook, style="TFrame", padding=15)
        tab_article_converter = ttk.Frame(notebook, style="TFrame", padding=15)

        notebook.add(tab_file_ops, text="🗂️ Файловые Операции")
        notebook.add(tab_path_gen, text="📋 Генератор Путей Excel")
        notebook.add(tab_folder_creator, text="🏗️ Создатель Папок")
        notebook.add(tab_article_converter, text="🔄 Конвертер Артикулов")
        
        self.create_file_ops_panel(tab_file_ops)
        self.create_path_generator_panel(tab_path_gen)
        self.create_folder_creator_panel(tab_folder_creator)
        self.create_article_converter_panel(tab_article_converter)

    def create_file_ops_panel(self, parent: tk.Frame):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        # Frame for operation selection
        selection_lf = ttk.LabelFrame(parent, text="1. Выберите операцию", style="Controls.TLabelframe")
        selection_lf.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        self.selected_file_op = tk.StringVar()
        self.file_op_buttons = {}
        btn_configs = [
            ("extract", "📤 Извлечь из '1'"),
            ("rename_images", "🔢 Переименовать 1-N"),
            ("remove_phrase", "✂️ Удалить фразу/RegEx"),
            ("delete_urls", "🗑️ Удалить URL-ярлыки"),
        ]

        for i, (op_type, text) in enumerate(btn_configs):
            row, col = divmod(i, 2)
            selection_lf.grid_columnconfigure(col, weight=1)
            rb = ttk.Radiobutton(
                selection_lf,
                text=text,
                variable=self.selected_file_op,
                value=op_type,
                command=self._on_file_op_selected,
                style="Toggle.TButton"
            )
            rb.grid(row=row, column=col, padx=5, pady=5, sticky="ew")
            self.file_op_buttons[op_type] = rb
            self.operation_buttons.append(rb) # Add to list for state changes

        # Frame for contextual options
        self.file_ops_options_frame = ttk.Frame(parent, style="Sub.TFrame")
        self.file_ops_options_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        self._create_file_op_option_widgets(self.file_ops_options_frame)

        # --- Execution and Dry Run Frame ---
        exec_lf = ttk.LabelFrame(parent, text="2. Запуск", style="Controls.TLabelframe")
        exec_lf.grid(row=2, column=0, sticky="ew", pady=(5, 0))
        exec_lf.grid_columnconfigure(1, weight=1)
        
        self.dry_run_var = tk.BooleanVar(value=self.last_dry_run)
        self.dry_run_cb = ttk.Checkbutton(exec_lf, text="✅ Пробный запуск (Dry Run)", variable=self.dry_run_var, style="Sub.TCheckbutton")
        self.dry_run_cb.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        ModernTooltip(self.dry_run_cb, "Симулировать операцию в логе без реального изменения файлов. Настоятельно рекомендуется!", app_themes=self.themes, current_theme_name_getter=self.get_current_theme_name)

        self.file_op_run_btn = ttk.Button(exec_lf, text="Выполнить", style="Accent.TButton", state="disabled", command=self._run_selected_file_op)
        self.file_op_run_btn.grid(row=0, column=1, padx=10, pady=10, sticky="e")
        self.operation_buttons.append(self.file_op_run_btn)

    def _create_file_op_option_widgets(self, parent: tk.Frame):
        """Creates the widgets for file operation options, initially hidden."""
        parent.grid_columnconfigure(0, weight=1)
        
        # --- Remove Phrase Options ---
        self.remove_phrase_options = ttk.Frame(parent, style="Sub.TFrame")
        self.remove_phrase_options.grid_columnconfigure(1, weight=1)
        
        ttk.Label(self.remove_phrase_options, text="Фраза / RegEx:").grid(row=0, column=0, sticky="w", padx=(0, 5), pady=5)
        self.phrase_var = tk.StringVar(value=self.last_phrase_to_remove)
        self.phrase_entry = PlaceholderEntry(self.remove_phrase_options, textvariable=self.phrase_var, placeholder="Введите фразу или регулярное выражение")
        self.phrase_entry.grid(row=0, column=1, sticky="ew", pady=5)
        
        phrase_opts_frame = ttk.Frame(self.remove_phrase_options, style="Sub.TFrame")
        phrase_opts_frame.grid(row=0, column=2, sticky="w", padx=(10, 0))
        self.case_sensitive_phrase_var = tk.BooleanVar(value=self.last_case_sensitive_phrase)
        ttk.Checkbutton(phrase_opts_frame, text="Регистр", variable=self.case_sensitive_phrase_var, style="Sub.TCheckbutton").pack(side="left")
        self.use_regex_var = tk.BooleanVar(value=self.last_use_regex)
        ttk.Checkbutton(phrase_opts_frame, text="RegEx", variable=self.use_regex_var, style="Sub.TCheckbutton").pack(side="left", padx=5)

        # --- Delete URLs Options ---
        self.delete_urls_options = ttk.Frame(parent, style="Sub.TFrame")
        self.delete_urls_options.grid_columnconfigure(1, weight=1)
        
        ttk.Label(self.delete_urls_options, text="Имена URL (через ','):").grid(row=0, column=0, sticky="w", padx=(0, 5), pady=5)
        self.url_names_var = tk.StringVar(value=self.last_url_names_to_delete)
        self.url_names_entry = PlaceholderEntry(self.delete_urls_options, textvariable=self.url_names_var, placeholder="имя1, частьимени2")
        self.url_names_entry.grid(row=0, column=1, sticky="ew", pady=5)
        
        self.case_sensitive_url_var = tk.BooleanVar(value=self.last_case_sensitive_url)
        ttk.Checkbutton(self.delete_urls_options, text="Регистр", variable=self.case_sensitive_url_var, style="Sub.TCheckbutton").grid(row=0, column=2, sticky="w", padx=10)

        # --- Store frames for easy access ---
        self.file_op_option_frames = {
            "remove_phrase": self.remove_phrase_options,
            "delete_urls": self.delete_urls_options,
        }

    def _on_file_op_selected(self):
        """Callback when a file operation radio button is selected."""
        selected_op = self.selected_file_op.get()
        if not selected_op:
            return

        # Update run button
        op_name = self.operations.get(selected_op, {}).get("name", "Выполнить")
        self.file_op_run_btn.config(text=f"Выполнить: {op_name}", state="normal")

        # Show/hide contextual option frames
        for op_type, frame in self.file_op_option_frames.items():
            if op_type == selected_op:
                frame.grid(row=0, column=0, sticky="ew")
            else:
                frame.grid_remove()
    
    def _run_selected_file_op(self):
        op_type = self.selected_file_op.get()
        if op_type:
            self.run_operation(op_type)
        else:
            messagebox.showwarning("Нет выбора", "Пожалуйста, сначала выберите операцию.", parent=self.master)

    def create_path_generator_panel(self, parent: tk.Frame):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(2, weight=1) # Let the result area expand

        input_lf = ttk.LabelFrame(parent, text="1. Введите названия моделей (каждое с новой строки)", style="Controls.TLabelframe")
        input_lf.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        input_lf.grid_columnconfigure(0, weight=1)
        input_lf.grid_rowconfigure(0, weight=1)
        self.path_gen_input_text = scrolledtext.ScrolledText(input_lf, wrap="word", font=("Consolas", 10), height=5)
        self.path_gen_input_text.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        ModernTooltip(self.path_gen_input_text, "Вставьте сюда список моделей. Каждая модель на новой строке.", app_themes=self.themes, current_theme_name_getter=self.get_current_theme_name)
        
        generate_btn = ttk.Button(parent, text="✅ Сгенерировать и проверить пути", command=lambda: self.run_operation("generate_paths"), style="Accent.TButton")
        generate_btn.grid(row=1, column=0, sticky="ew", pady=5)
        self.operation_buttons.append(generate_btn)
        
        output_lf = ttk.LabelFrame(parent, text="2. Результат", style="Controls.TLabelframe")
        output_lf.grid(row=2, column=0, sticky="nsew")
        output_lf.grid_columnconfigure(0, weight=1)
        output_lf.grid_rowconfigure(0, weight=3) # Success gets more space
        output_lf.grid_rowconfigure(1, weight=1) # Error gets less space

        self.path_gen_output_text = scrolledtext.ScrolledText(output_lf, wrap="none", font=("Consolas", 10), height=6, background="#f0fff0")
        self.path_gen_output_text.grid(row=0, column=0, sticky="nsew", padx=5, pady=(5, 0))
        self.path_gen_output_text.configure(state="disabled")
        
        self.path_gen_error_text = scrolledtext.ScrolledText(output_lf, wrap="none", font=("Consolas", 10), height=3, background="#fff0f0", foreground="red")
        self.path_gen_error_text.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.path_gen_error_text.configure(state="disabled")
        
        copy_btn = ttk.Button(output_lf, text="Копировать успешные результаты", command=self.copy_path_gen_results)
        copy_btn.grid(row=2, column=0, sticky="e", padx=5, pady=5)

    def create_folder_creator_panel(self, parent: tk.Frame):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        input_lf = ttk.LabelFrame(parent, text="1. Введите названия папок (каждое с новой строки)", style="Controls.TLabelframe")
        input_lf.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        input_lf.grid_columnconfigure(0, weight=1)
        input_lf.grid_rowconfigure(0, weight=1)
        self.folder_creator_input_text = scrolledtext.ScrolledText(input_lf, wrap="word", font=("Consolas", 10), height=5)
        self.folder_creator_input_text.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        ModernTooltip(self.folder_creator_input_text, "Можно создавать вложенные папки, например: ProjectA/assets", app_themes=self.themes, current_theme_name_getter=self.get_current_theme_name)

        options_lf = ttk.LabelFrame(parent, text="2. Опции создания", style="Controls.TLabelframe")
        options_lf.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        options_lf.grid_columnconfigure(1, weight=1)
        options_lf.grid_columnconfigure(3, weight=1)

        ttk.Label(options_lf, text="Префикс:").grid(row=0, column=0, sticky="w", padx=(10, 5), pady=5)
        self.folder_prefix_var = tk.StringVar(value=self.last_folder_prefix)
        self.folder_prefix_entry = PlaceholderEntry(options_lf, textvariable=self.folder_prefix_var)
        self.folder_prefix_entry.grid(row=0, column=1, sticky="ew", padx=(0, 5), pady=5)
        
        ttk.Label(options_lf, text="Суффикс:").grid(row=0, column=2, sticky="w", padx=(10, 5), pady=5)
        self.folder_suffix_var = tk.StringVar(value=self.last_folder_suffix)
        self.folder_suffix_entry = PlaceholderEntry(options_lf, textvariable=self.folder_suffix_var)
        self.folder_suffix_entry.grid(row=0, column=3, sticky="ew", padx=(0, 10), pady=5)

        self.folder_numbering_var = tk.BooleanVar(value=self.last_folder_numbering)
        self.folder_numbering_cb = ttk.Checkbutton(options_lf, text="Включить автонумерацию", variable=self.folder_numbering_var, style="Sub.TCheckbutton")
        self.folder_numbering_cb.grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=5)
        
        num_opts_frame = ttk.Frame(options_lf, style="Sub.TFrame")
        num_opts_frame.grid(row=1, column=2, columnspan=2, sticky="w", padx=(10, 0))
        
        self.folder_start_num_var = tk.IntVar(value=self.last_folder_start_num)
        ttk.Label(num_opts_frame, text="Начать с:").pack(side="left")
        ttk.Spinbox(num_opts_frame, from_=0, to=9999, textvariable=self.folder_start_num_var, width=6).pack(side="left", padx=(2, 10))

        self.folder_padding_var = tk.IntVar(value=self.last_folder_padding)
        ttk.Label(num_opts_frame, text="Цифр (padding):").pack(side="left")
        ttk.Spinbox(num_opts_frame, from_=1, to=10, textvariable=self.folder_padding_var, width=4).pack(side="left", padx=2)

        create_btn = ttk.Button(parent, text="✅ Создать папки", command=lambda: self.run_operation("create_folders"), style="Accent.TButton")
        create_btn.grid(row=2, column=0, sticky="ew", pady=5)
        self.operation_buttons.append(create_btn)

    def create_article_converter_panel(self, parent: tk.Frame):
        parent.grid_columnconfigure(0, weight=1)
        container = ttk.Frame(parent)
        container.grid(sticky="nsew", padx=20, pady=20)
        container.grid_columnconfigure(0, weight=1)

        self.converter_select_btn = ttk.Button(container, text="1. Выбрать Excel/CSV файл", command=self.select_and_scan_converter_file)
        self.converter_select_btn.grid(row=0, column=0, pady=5, ipady=5, sticky="ew")
        
        self.converter_file_label = ttk.Label(container, text="Файл не выбран", anchor="center")
        self.converter_file_label.grid(row=1, column=0, pady=2, sticky="ew")

        self.converter_detected_label = ttk.Label(container, text="", font=("Segoe UI", 10, "bold"), anchor="center")
        self.converter_detected_label.grid(row=2, column=0, pady=5, sticky="ew")
        
        ttk.Label(container, text="2. Выберите НОВЫЙ размер для замены:", anchor="center").grid(row=3, column=0, pady=(10, 0), sticky="ew")
        
        self.converter_size_combobox = ttk.Combobox(container, state="disabled")
        self.converter_size_combobox.grid(row=4, column=0, pady=5, ipady=3, sticky="ew")

        self.converter_process_btn = ttk.Button(container, text="3. Создать файл с новым размером", command=self.process_and_save_converter_file, state="disabled")
        self.converter_process_btn.grid(row=5, column=0, pady=5, ipady=5, sticky="ew")
        
        ttk.Separator(container, orient="horizontal").grid(row=6, column=0, sticky="ew", pady=20)
        
        self.converter_edit_btn = ttk.Button(container, text="⚙️ Редактор размеров", command=self.open_size_editor)
        self.converter_edit_btn.grid(row=7, column=0, pady=10, sticky="ew")

    def create_log_panel(self, parent: tk.Frame):
        log_lf = ttk.LabelFrame(parent, text="📋 Журнал операций", style="Controls.TLabelframe")
        log_lf.grid(row=0, column=0, sticky="nsew", pady=(5, 0))
        log_lf.grid_columnconfigure(0, weight=1)
        log_lf.grid_rowconfigure(0, weight=1)
        
        self.output_log = scrolledtext.ScrolledText(log_lf, wrap="word", font=("Consolas", 10), relief="flat", borderwidth=0)
        self.output_log.grid(row=0, column=0, sticky="nsew", padx=5, pady=(5, 0))
        self.output_log.configure(state="disabled")
        self.output_log.bind("<Control-c>", lambda e: self.master.clipboard_append(self.output_log.selection_get()))
        
        # Log context menu
        log_context_menu = tk.Menu(self.master, tearoff=0)
        log_context_menu.add_command(label="Копировать", command=self.copy_selected_log)
        log_context_menu.add_command(label="Копировать всё", command=self.copy_all_log)
        self.output_log.bind("<Button-3>", lambda e: log_context_menu.tk_popup(e.x_root, e.y_root))
        
        self.logger = Logger(self.output_log)
        
        log_buttons_frame = ttk.Frame(log_lf, style="Sub.TFrame")
        log_buttons_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=(5, 10))
        
        self.clear_log_btn = ttk.Button(log_buttons_frame, text="🗑️ Очистить лог", command=self.clear_log)
        self.clear_log_btn.pack(side="left")
        
        self.save_log_btn = ttk.Button(log_buttons_frame, text="📁 Сохранить лог", command=self.save_log_to_file)
        self.save_log_btn.pack(side="left", padx=5)
        
        self.stop_btn = ttk.Button(log_buttons_frame, text="⏹️ Остановить операцию", command=self.stop_current_operation, state="disabled", style="Danger.TButton")
        self.stop_btn.pack(side="right")
        ModernTooltip(self.stop_btn, "Прервать выполнение текущей длительной операции.", app_themes=self.themes, current_theme_name_getter=self.get_current_theme_name)

    def create_progress_panel(self, parent: tk.Frame):
        progress_frame = ttk.Frame(parent, style="TFrame")
        progress_frame.grid(row=1, column=0, sticky="ew", pady=(5, 0))
        progress_frame.grid_columnconfigure(0, weight=1)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode="determinate", style="Custom.Horizontal.TProgressbar")
        self.progress_bar.grid(row=0, column=0, sticky="ew")
        
        self.progress_label = ttk.Label(progress_frame, text="", width=10, anchor="e")
        self.progress_label.grid(row=0, column=1, padx=(10, 0))

    def create_footer(self):
        theme = self.themes[self.current_theme_name]
        self.footer_frame = tk.Frame(self.master, height=30, bg=theme.get("footer_bg", theme["bg"]))
        self.footer_frame.grid(row=2, column=0, sticky="ew")
        self.footer_frame.grid_propagate(False)
        self.footer_frame.grid_columnconfigure(0, weight=1)
        
        self.status_var = tk.StringVar(value="Готов")
        self.status_label = tk.Label(self.footer_frame, textvariable=self.status_var, font=("Segoe UI", 9), anchor="w", bg=self.footer_frame.cget("bg"))
        self.status_label.pack(side="left", fill="x", padx=20, pady=5)

    def apply_theme(self, theme_name: str):
        if not hasattr(self, "themes") or not hasattr(self, "style"):
            return
        
        theme = self.themes[theme_name]
        bg = theme["bg"]
        fg = theme["fg"]
        self.master.configure(bg=bg)

        # Configure styles
        self.style.configure(".", background=bg, foreground=fg, fieldbackground=theme["secondary_bg"])
        self.style.configure("TFrame", background=bg)
        self.style.configure("Sub.TFrame", background=bg)
        self.style.configure("TLabel", background=bg, foreground=fg, font=("Segoe UI", 10))
        self.style.configure("TLabelframe", background=bg, bordercolor=theme["border"], relief="solid")
        self.style.configure("TLabelframe.Label", background=bg, foreground=theme.get("labelframe_label_fg", theme["accent"]), font=("Segoe UI", 11, "bold"))
        self.style.configure("Controls.TLabelframe", background=bg, bordercolor=theme["border"], relief="solid")
        self.style.configure("Controls.TLabelframe.Label", background=bg, foreground=theme.get("labelframe_label_fg", theme["accent"]), font=("Segoe UI", 11, "bold"))
        
        # Button styles
        self.style.configure("TButton", font=("Segoe UI", 10), padding=(10, 5), relief="flat")
        self.style.map("TButton", background=[("active", theme["hover"]), ("disabled", theme["disabled_bg"])], foreground=[("disabled", theme["disabled_fg"])])
        self.style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(10, 5), background=theme["accent"], foreground=theme["accent_fg"])
        self.style.map("Accent.TButton", background=[("active", theme["hover"]), ("disabled", theme["disabled_bg"])], foreground=[("disabled", theme["disabled_fg"])])
        self.style.configure("Header.TButton", font=("Segoe UI", 9), padding=(8, 4))
        self.style.map("Header.TButton", background=[("active", theme["hover"])])
        self.style.configure("Danger.TButton", font=("Segoe UI", 10, "bold"), padding=(10, 5), background=theme["button_danger_bg"], foreground=theme["button_danger_fg"])
        self.style.map("Danger.TButton", background=[("active", theme["button_danger_hover"]), ("disabled", theme["disabled_bg"])], foreground=[("disabled", theme["disabled_fg"])])
        
        self.style.configure("Toggle.TButton", font=("Segoe UI", 10), padding=(10, 5), relief="flat")
        self.style.map("Toggle.TButton", background=[("selected", theme["accent"]), ("active", theme["hover"]), ("disabled", theme["disabled_bg"])], foreground=[("selected", theme["accent_fg"]), ("disabled", theme["disabled_fg"])])

        # Other widget styles
        self.style.configure("TEntry", font=("Segoe UI", 10), padding=5, bordercolor=theme["border"], insertcolor=fg, relief="flat")
        self.style.map("TEntry", bordercolor=[("focus", theme["accent"])], fieldbackground=[("disabled", theme["disabled_bg"])])
        self.style.configure("Inactive.TEntry", fieldbackground=theme["entry_bg"])
        self.style.configure("Active.TEntry", fieldbackground=theme["secondary_bg"])
        self.style.configure("TCheckbutton", background=bg, foreground=fg, font=("Segoe UI", 9))
        self.style.configure("Sub.TCheckbutton", background=bg, foreground=fg, font=("Segoe UI", 9))
        self.style.configure("Custom.Horizontal.TProgressbar", troughcolor=theme["secondary_bg"], background=theme["progress_bar"], thickness=20)
        self.style.map("TCombobox", fieldbackground=[("readonly", theme["secondary_bg"])], foreground=[("readonly", fg)])

        # Update individual widgets
        if hasattr(self, "header_frame"):
            self.header_frame.config(bg=theme.get("header_bg", bg))
            self.title_label.config(bg=self.header_frame.cget("bg"), fg=fg)
            self.subtitle_label.config(bg=self.header_frame.cget("bg"), fg=fg)
        if hasattr(self, "footer_frame"):
            self.footer_frame.config(bg=theme.get("footer_bg", bg))
            self.status_label.config(bg=self.footer_frame.cget("bg"), fg=fg)
        if hasattr(self, "output_log"):
            self.output_log.config(background=theme["log_bg"], foreground=theme["log_fg"], insertbackground=fg, selectbackground=theme["accent"], selectforeground=theme["accent_fg"])
            self.setup_log_tags()
        if hasattr(self, "theme_btn"):
            theme_icon = "🌙" if theme_name == "light" else "☀️"
            self.theme_btn.config(text=f"{theme_icon} Тема")
        if hasattr(self, "converter_detected_label"):
            self.converter_detected_label.config(foreground=theme["accent"])
        
        self.master.update_idletasks()

    def setup_log_tags(self):
        if not hasattr(self, "output_log") or not self.output_log:
            return
        theme = self.themes[self.current_theme_name]
        self.output_log.tag_config("info", foreground=theme.get("log_info", theme["fg"]))
        self.output_log.tag_config("success", foreground=theme.get("log_success", "green"), font=("Consolas", 10, "bold"))
        self.output_log.tag_config("warning", foreground=theme.get("log_warning", "orange"))
        self.output_log.tag_config("error", foreground=theme.get("log_error", "red"), font=("Consolas", 10, "bold"))

    def setup_bindings(self):
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.master.bind("<Control-o>", lambda e: self.browse_folder())
        self.master.bind("<Control-l>", lambda e: self.clear_log())
        self.master.bind("<Control-h>", lambda e: self.show_help())

    def run_operation(self, op_type: str):
        if self.current_thread and self.current_thread.is_alive():
            messagebox.showwarning("Операция выполняется", "Другая операция уже запущена.", parent=self.master)
            return

        op_details = self.operations.get(op_type)
        if not op_details:
            self.logger.error(f"Неизвестный тип операции: {op_type}")
            return

        source_path = self.path_var.get().strip()
        if not self.validate_path(source_path, op_details["name"]):
            return
        
        if op_details.get("pre_check") and not op_details["pre_check"]():
            messagebox.showwarning("Нет данных", op_details["pre_check_msg"], parent=self.master)
            return

        dry_run = self.dry_run_var.get() if op_details["is_file_op"] else False
        if op_details["is_file_op"] and not dry_run:
            if not self.confirm_operation(op_details["name"]):
                self.logger.info("Операция отменена пользователем.")
                self.update_status("Операция отменена.", 0)
                return

        self.operation_result_counter = 0
        self.current_operation_is_path_gen = (op_type == "generate_paths")

        self.clear_log()
        self.master.title(f"🗂️ Выполняется: {op_details['name']}...")
        self.update_status(f"Запуск '{op_details['name']}'...", 0)
        self.stop_event.clear()
        self.set_ui_state(active=False)

        def operation_wrapper():
            args = op_details["get_args"]()
            result = op_details["function"](*args)
            if isinstance(result, int):
                self.operation_result_counter = result

        self.current_thread = threading.Thread(target=operation_wrapper, daemon=True)
        self.current_thread.start()
        self.master.after(100, self.check_thread_completion)

    # --- UI Actions & Helpers ---

    def show_welcome_message(self):
        if hasattr(self, "logger"):
            self.logger.info("🎉 Добро пожаловать в Супер Скрипт v2.7!")
            self.logger.info("💡 Выберите вкладку, папку и операцию для начала работы.")
            self.update_status("Готов к работе. Выберите вкладку и операцию.", 0)

    def get_current_theme_name(self) -> str:
        return self.current_theme_name

    def toggle_theme(self):
        self.current_theme_name = "dark" if self.current_theme_name == "light" else "light"
        self.logger.info(f"Тема изменена на: {self.current_theme_name.capitalize()}")
        self.apply_theme(self.current_theme_name)

    def show_help(self):
        help_text = """Супер Скрипт v2.7
---------------------------------------------
Этот инструмент предназначен для автоматизации общих задач по организации файлов, созданию папок и генерации данных.

**Вкладка '🗂️ Файловые Операции'**:
  📤 **Извлечь из '1'**: Ищет папки с именем '1', перемещает их содержимое в родительскую папку и удаляет пустую папку '1'.
  🔢 **Переименовать 1-N**: Находит все изображения в каждой подпапке и переименовывает их в числовую последовательность (1.jpg, 2.jpg...).
  ✂️ **Удалить фразу/RegEx**: Удаляет указанную фразу или регулярное выражение из имен всех файлов и папок.
  🗑️ **Удалить URL-ярлыки**: Удаляет файлы .url, имена которых содержат указанные строки.
  ✅ **Пробный запуск (Dry Run)**: **САМАЯ ВАЖНАЯ ОПЦИЯ!** Позволяет симулировать операцию без реального изменения файлов. Все действия будут показаны в логе.

**Вкладка '📋 Генератор Путей Excel'**:
  - ✅ **Сгенерировать и проверить пути**: На основе списка введенных моделей создает строки с путями ко ВСЕМ найденным изображениям для каждой модели.
  - Эта операция является безопасной и не изменяет файлы.

**Вкладка '🏗️ Создатель Папок'**:
  - ✅ **Создать папки**: Создает папки в общей рабочей директории на основе введенного списка.
  - Поддерживает создание вложенных папок (например, `Проект/Ресурсы`).
  - Можно добавлять префиксы, суффиксы и автоматическую нумерацию.
  - Эта операция также поддерживает 'Пробный запуск'.

**Вкладка '🔄 Конвертер Артикулов'**:
  - Предназначен для быстрой замены артикулов размеров в файлах Excel/CSV.
  - 1. Выберите файл. Скрипт автоматически найдет в нем известный артикул.
  - 2. Выберите из списка новый размер, на который нужно произвести замену.
  - 3. Нажмите "Создать файл", чтобы сохранить новую копию файла с замененным артикулом.
  - ⚙️ **Редактор размеров**: Позволяет добавлять, изменять и удалять пары "Размер-Артикул" в вашем словаре. Изменения сохраняются в `sizes.json`.

⚠️ **ВАЖНО**: Перед запуском любой операции, которая изменяет файлы (без галочки 'Пробный запуск'), настоятельно рекомендуется создать резервную копию ваших данных!"""
        
        help_window = tk.Toplevel(self.master)
        help_window.title("Справка - Супер Скрипт v2.7")
        help_window.geometry("800x650")
        help_window.transient(self.master)
        help_window.grab_set()
        
        theme = self.themes[self.current_theme_name]
        help_window.configure(bg=theme["bg"])
        
        text_area = scrolledtext.ScrolledText(help_window, wrap="word", font=("Segoe UI", 10), bg=theme["log_bg"], fg=theme["log_fg"], relief="flat")
        text_area.pack(fill="both", expand=True, padx=10, pady=10)
        text_area.insert(tk.END, help_text)
        text_area.configure(state="disabled")
        
        ttk.Button(help_window, text="Закрыть", command=help_window.destroy, style="Accent.TButton").pack(pady=10)

    def browse_folder(self):
        initial_dir = self.path_var.get() if os.path.isdir(self.path_var.get()) else self.last_path
        folder_selected = filedialog.askdirectory(initialdir=initial_dir, title="Выберите рабочую папку", parent=self.master)
        if folder_selected:
            self.path_var.set(folder_selected)
            self.logger.info(f"Выбрана рабочая папка: {folder_selected}")
            self.update_status(f"Рабочая папка: {os.path.basename(folder_selected)}", 0)

    def clear_log(self):
        if hasattr(self, "output_log") and self.output_log:
            self.output_log.configure(state="normal")
            self.output_log.delete(1.0, tk.END)
            self.output_log.configure(state="disabled")
            self.logger.info("Журнал операций очищен.")
            self.update_status("Журнал очищен.", 0)

    def save_log_to_file(self):
        content = self.output_log.get(1.0, tk.END)
        if not content.strip():
            self.logger.warning("Лог пуст, нечего сохранять.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".log", filetypes=[("Log files", "*.log"), ("All files", "*.*")])
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)
                self.logger.info(f"Лог сохранен в {path}")
            except Exception as e:
                self.logger.error(f"Ошибка сохранения лога: {e}")

    def copy_selected_log(self):
        try:
            if self.output_log.tag_ranges("sel"):
                self.master.clipboard_clear()
                self.master.clipboard_append(self.output_log.get(tk.SEL_FIRST, tk.SEL_LAST))
        except tk.TclError:
            pass  # Ignore error if no selection

    def copy_all_log(self):
        content = self.output_log.get("1.0", tk.END).strip()
        if content:
            self.master.clipboard_clear()
            self.master.clipboard_append(content)
            self.logger.info("Содержимое лога скопировано.")

    def update_status(self, message: str, progress: Optional[int] = None):
        if hasattr(self, "status_var"):
            self.status_var.set(message)
        if hasattr(self, "progress_bar"):
            if progress is not None:
                self.progress_bar["value"] = progress
                self.progress_label.config(text=f"{progress}%")
            else:
                self.progress_bar["value"] = 0
                self.progress_label.config(text="")

    def set_ui_state(self, active: bool):
        state = "normal" if active else "disabled"
        readonly_state = "readonly" if not active else "normal"
        
        # Path entry and browse button
        if hasattr(self, "path_entry"): self.path_entry.config(state=readonly_state)
        if hasattr(self, "browse_btn"): self.browse_btn.config(state=state)
        
        # Operation buttons
        for btn in self.operation_buttons:
            btn.config(state=state)
            
        # Options entries and checkboxes
        for widget_attr in ["phrase_entry", "url_names_entry", "folder_creator_input_text", "folder_prefix_entry", "folder_suffix_entry"]:
            if hasattr(self, widget_attr): getattr(self, widget_attr).config(state=readonly_state)
        
        for cb_attr in ["case_sensitive_phrase_cb", "use_regex_cb", "case_sensitive_url_cb", "dry_run_cb", "folder_numbering_cb"]:
            if hasattr(self, cb_attr): getattr(self, cb_attr).config(state=state)
            
        # Log buttons
        if hasattr(self, "clear_log_btn"): self.clear_log_btn.config(state=state)
        if hasattr(self, "save_log_btn"): self.save_log_btn.config(state=state)
        if hasattr(self, "stop_btn"): self.stop_btn.config(state="disabled" if active else "normal")

        # Header buttons
        if hasattr(self, "theme_btn"): self.theme_btn.config(state=state)
        if hasattr(self, "help_btn"): self.help_btn.config(state=state)

    def validate_path(self, path: str, operation_name: str) -> bool:
        if not path or not path.strip():
            msg = f"Путь к папке не указан для '{operation_name}'."
            self.logger.error(msg)
            messagebox.showerror("Ошибка пути", "Пожалуйста, выберите или введите путь к рабочей папке.", parent=self.master)
            return False
        if not os.path.isdir(path):
            msg = f"Указанный путь '{path}' не существует или не является папкой для '{operation_name}'."
            self.logger.error(msg)
            messagebox.showerror("Ошибка пути", f"Указанный путь не существует или не является папкой:\n{path}", parent=self.master)
            return False
        return True

    def confirm_operation(self, operation_name: str) -> bool:
        confirm_msg = f"""Вы уверены, что хотите запустить операцию:
'{operation_name}'?

Рабочая папка:
'{self.path_var.get()}'

⚠️ Эта операция может необратимо изменить или удалить файлы.
Настоятельно рекомендуется создать резервную копию данных!"""
        return messagebox.askyesno("Подтверждение операции", confirm_msg, icon="warning", parent=self.master)

    def check_thread_completion(self):
        if self.current_thread and self.current_thread.is_alive():
            self.master.after(200, self.check_thread_completion)
        else:
            if self.current_thread and not self.stop_event.is_set():
                if not self.current_operation_is_path_gen:
                    op_type = "Пробный запуск" if self.dry_run_var.get() else "Операция"
                    summary_msg = f"{op_type} завершена.\n\nОбработано элементов: {self.operation_result_counter}"
                    messagebox.showinfo("Операция завершена", summary_msg, parent=self.master)

            self.set_ui_state(active=True)
            self.master.title("🗂️ Супер Скрипт v2.7")
            if self.stop_event.is_set():
                self.update_status("Операция остановлена.", 0)
            
            self.current_thread = None
            self.stop_event.clear()

    def stop_current_operation(self):
        if self.current_thread and self.current_thread.is_alive():
            msg = "Вы уверены, что хотите прервать текущую операцию?\nНекоторые изменения могут быть уже применены."
            if messagebox.askyesno("Остановить операцию?", msg, icon="warning", parent=self.master):
                self.stop_event.set()
                self.logger.warning("--- Попытка остановить операцию... Ожидайте завершения текущего шага. ---")
                self.update_status("Остановка операции... Пожалуйста, подождите.", None)
                self.stop_btn.config(state="disabled")
        else:
            self.logger.info("Нет активной операции для остановки.")
            self.update_status("Нет активных операций.", 0)

    # --- Article Converter Methods ---

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
            if hasattr(self, "logger"):
                self.logger.info(f"Словарь размеров сохранен в '{SIZES_JSON_FILE}'.")
        except Exception as e:
            if hasattr(self, "logger"):
                self.logger.error(f"Не удалось сохранить словарь размеров: {e}")

    def update_converter_combobox(self):
        if hasattr(self, "converter_size_combobox"):
            self.converter_size_combobox["values"] = list(self.size_to_article_map.keys())
            self.logger.info("Список размеров в комбобоксе обновлен.")

    def open_size_editor(self):
        self.logger.info("Открыт редактор размеров.")
        editor = SizeEditor(self.master, self)
    
    def universal_file_reader(self, file_path: str) -> pd.DataFrame:
        try:
            return pd.read_excel(file_path, header=None, dtype=str)
        except Exception:
            try:
                return pd.read_csv(file_path, header=None, dtype=str, engine="python", encoding="utf-8-sig")
            except Exception:
                try:
                    return pd.read_csv(file_path, header=None, dtype=str, engine="python", encoding="cp1251")
                except Exception as e:
                    raise ValueError(f"Не удалось прочитать файл. Детали: {e}")

    def select_and_scan_converter_file(self):
        file_path = filedialog.askopenfilename(title="Выберите исходный файл", filetypes=[("Таблицы", "*.xlsx *.xls *.csv"), ("Все файлы", "*.*")])
        if not file_path:
            return

        self.converter_input_file_path = file_path
        self.converter_detected_article = None
        self.converter_process_btn.config(state="disabled")
        self.converter_size_combobox.config(state="disabled")
        self.converter_size_combobox.set("")
        
        filename = os.path.basename(file_path)
        self.converter_file_label.config(text=f"Выбран: {filename}")
        self.logger.info(f"Конвертер: выбран файл '{file_path}'")
        self.update_status(f"Сканирование файла {filename}...", None)

        try:
            df = self.universal_file_reader(self.converter_input_file_path)
            article_to_size_map = {str(v): k for k, v in self.size_to_article_map.items()}
            articles_set = {str(v) for v in self.size_to_article_map.values()}
            
            # Find the first known article in the file
            for col in df.columns:
                for cell_content in df[col].dropna():
                    if isinstance(cell_content, str):
                        for article in articles_set:
                            if article in cell_content:
                                self.converter_detected_article = article
                                break
                    if self.converter_detected_article: break
                if self.converter_detected_article: break
            
            if self.converter_detected_article:
                detected_size = article_to_size_map[self.converter_detected_article]
                self.converter_detected_label.config(text=f"Найден размер в файле: {detected_size}")
                self.logger.success(f"В файле найден артикул '{self.converter_detected_article}' (размер: {detected_size})")
                self.converter_size_combobox.config(state="readonly")
                self.converter_process_btn.config(state="normal")
                self.update_status(f"Найден размер {detected_size}. Выберите новый размер.", 0)
            else:
                self.converter_detected_label.config(text="В файле не найден ни один известный артикул!")
                self.logger.error("В файле не найден ни один известный артикул из словаря.")
                self.update_status("Артикул не найден. Проверьте файл или редактор размеров.", 0)

        except Exception as e:
            self.logger.error(f"Ошибка чтения файла: {e}")
            messagebox.showerror("Ошибка чтения файла", str(e), parent=self.master)
            self.update_status("Ошибка чтения файла.", 0)

    def process_and_save_converter_file(self):
        newly_selected_size = self.converter_size_combobox.get()
        if not newly_selected_size:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите новый размер!", parent=self.master)
            return
        
        self.logger.info(f"Запущена обработка. Старый артикул: {self.converter_detected_article}, новый размер: {newly_selected_size}")
        self.update_status("Обработка файла...", 50)
        
        try:
            df = self.universal_file_reader(self.converter_input_file_path)
            new_article = str(self.size_to_article_map[newly_selected_size])
            df = df.applymap(lambda cell: cell.replace(self.converter_detected_article, new_article) if isinstance(cell, str) else cell)
            
            original_path = Path(self.converter_input_file_path)
            original_extension = original_path.suffix.lower()
            if original_extension not in [".xls", ".xlsx", ".csv"]:
                original_extension = ".xlsx"
            
            # Suggest a new filename
            size_for_filename = newly_selected_size.replace(" ", "").replace(".", "_")
            initial_name = original_path.stem
            suggested_filename = f"{initial_name.strip()}_{size_for_filename}{original_extension}"
            
            output_path_str = filedialog.asksaveasfilename(
                title="Куда сохранить готовый файл?",
                defaultextension=original_extension,
                filetypes=[("Исходный формат", f"*{original_extension}"), ("Книга Excel", "*.xlsx")],
                initialfile=suggested_filename,
                parent=self.master
            )
            
            if not output_path_str:
                self.logger.warning("Операция сохранения отменена пользователем.")
                self.update_status("Сохранение отменено.", 0)
                return

            output_path = Path(output_path_str)
            if output_path.suffix.lower() == ".xls" and HAS_WIN32:
                # Special handling for legacy .xls format via COM
                temp_xlsx_path = str(output_path.with_suffix(".tmp.xlsx"))
                df.to_excel(temp_xlsx_path, index=False, header=False)
                excel = win32.gencache.EnsureDispatch("Excel.Application")
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(os.path.abspath(temp_xlsx_path))
                wb.SaveAs(os.path.abspath(output_path_str), FileFormat=56) # 56 is for xlExcel8
                wb.Close()
                excel.Application.Quit()
                os.remove(temp_xlsx_path)
            elif output_path.suffix.lower() == ".csv":
                df.to_csv(output_path, index=False, header=False, encoding="utf-8-sig")
            else: # Default to xlsx
                df.to_excel(output_path, index=False, header=False)
            
            self.logger.success(f"Готовый файл для размера {newly_selected_size} успешно создан: {output_path}")
            messagebox.showinfo("Успех!", f"Готовый файл для размера {newly_selected_size} успешно создан!", parent=self.master)
            self.update_status("Готово.", 100)
        
        except Exception as e:
            self.logger.error(f"Ошибка при обработке файла: {e}")
            messagebox.showerror("Ошибка при обработке", f"Не удалось обработать файл.\n\nДетали: {e}", parent=self.master)
            self.update_status("Ошибка обработки.", 0)

    def path_gen_result_callback(self, success_str: str, error_str: str):
        """Callback to update the UI with path generation results."""
        def update_ui():
            for widget, text in [(self.path_gen_output_text, success_str), (self.path_gen_error_text, error_str)]:
                widget.config(state="normal")
                widget.delete("1.0", tk.END)
                widget.insert("1.0", text)
                widget.config(state="disabled")
        self.master.after(0, update_ui)

    def copy_path_gen_results(self):
        content = self.path_gen_output_text.get("1.0", tk.END).strip()
        if content:
            self.master.clipboard_clear()
            self.master.clipboard_append(content)
            self.logger.info("Результаты генератора путей скопированы в буфер обмена.")
            messagebox.showinfo("Скопировано", "Успешные результаты были скопированы.", parent=self.master)
        else:
            messagebox.showwarning("Нечего копировать", "Поле успешных результатов пусто.", parent=self.master)

    def on_closing(self):
        """Save configuration before exiting the application."""
        self.config["last_path"] = self.path_var.get()
        self.config["theme"] = self.current_theme_name
        self.config["last_phrase_to_remove"] = self.phrase_var.get()
        self.config["last_url_names_to_delete"] = self.url_names_var.get()
        self.config["last_case_sensitive_phrase"] = self.case_sensitive_phrase_var.get()
        self.config["last_case_sensitive_url"] = self.case_sensitive_url_var.get()
        self.config["last_use_regex"] = self.use_regex_var.get()
        self.config["last_dry_run"] = self.dry_run_var.get()
        
        if hasattr(self, "folder_prefix_var"):
            self.config["last_folder_prefix"] = self.folder_prefix_var.get()
            self.config["last_folder_suffix"] = self.folder_suffix_var.get()
            self.config["last_folder_numbering"] = self.folder_numbering_var.get()
            self.config["last_folder_start_num"] = self.folder_start_num_var.get()
            self.config["last_folder_padding"] = self.folder_padding_var.get()

        ConfigManager.save_config(self.config)
        self.save_sizes()

        if self.current_thread and self.current_thread.is_alive():
            msg = "Активная операция еще не завершена. Вы уверены, что хотите выйти?"
            if messagebox.askyesno("Операция выполняется", msg, icon="warning", parent=self.master):
                self.stop_event.set()
                self.master.destroy()
        else:
            self.master.destroy()

# --- Entry Point ---
if __name__ == "__main__":
    root = tk.Tk()
    app = ModernFileOrganizerApp(root)
    root.mainloop()
