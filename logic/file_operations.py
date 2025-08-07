import os
import shutil
import re
import threading
from pathlib import Path
from typing import Callable

# This will be created shortly in utils/logger.py
from utils.logger import Logger

IMAGE_EXTENSIONS = [
    ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".webp", ".svg", ".ico"
]


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
