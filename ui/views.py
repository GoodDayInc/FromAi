import customtkinter as ctk
from tkinter import messagebox
from ui.widgets import PlaceholderEntry, Tooltip

class FileOperationsView(ctk.CTkFrame):
    def __init__(self, master, controller, **kwargs):
        super().__init__(master, **kwargs)
        self.controller = controller

        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text="Файловые Операции", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, padx=20, pady=(10, 20), sticky="w")

        # --- Operation Selection Group ---
        selection_group = ctk.CTkFrame(self)
        selection_group.grid(row=1, column=0, sticky="ew", padx=20, pady=10)
        selection_group.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(selection_group, text="1. Выберите операцию", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=15, pady=(15, 10), sticky="w")

        self.selected_file_op = ctk.StringVar()
        btn_configs = [
            ("extract", "📤 Извлечь из '1'"), ("rename_images", "🔢 Переименовать 1-N"),
            ("remove_phrase", "✂️ Удалить фразу/RegEx"), ("delete_urls", "🗑️ Удалить URL-ярлыки"),
        ]
        radio_frame = ctk.CTkFrame(selection_group, fg_color="transparent")
        radio_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        for i, (op_type, text) in enumerate(btn_configs):
            radio_frame.grid_columnconfigure(i, weight=1)
            rb = ctk.CTkRadioButton(radio_frame, text=text, variable=self.selected_file_op, value=op_type, command=self._on_file_op_selected)
            rb.grid(row=0, column=i, padx=5, pady=5, sticky="ew")
            self.controller.operation_buttons[f"op_{op_type}"] = rb

        # --- Contextual Options Group ---
        self.options_group = ctk.CTkFrame(self)
        self.options_group.grid(row=2, column=0, sticky="ew", padx=20, pady=10)
        self.options_group.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.options_group, text="2. Настройки операции", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=15, pady=(15, 10), sticky="w")
        self._create_file_op_option_widgets(self.options_group)

        # --- Execution Group ---
        exec_group = ctk.CTkFrame(self)
        exec_group.grid(row=3, column=0, sticky="ew", padx=20, pady=10)
        exec_group.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(exec_group, text="3. Запуск", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=2, padx=15, pady=(15, 10), sticky="w")

        dry_run_cb = ctk.CTkCheckBox(exec_group, text="✅ Пробный запуск (Dry Run)", variable=self.controller.dry_run_var)
        dry_run_cb.grid(row=1, column=0, padx=15, pady=15, sticky="w")
        Tooltip(dry_run_cb, "Симулировать операцию в логе без реального изменения файлов.")
        self.controller.operation_buttons["dry_run_cb"] = dry_run_cb

        self.file_op_run_btn = ctk.CTkButton(exec_group, text="Выполнить", state="disabled", command=self._run_selected_file_op)
        self.file_op_run_btn.grid(row=1, column=1, padx=15, pady=15, sticky="e")
        self.controller.operation_buttons["run_file_op"] = self.file_op_run_btn

    def _create_file_op_option_widgets(self, parent: ctk.CTkFrame):
        options_container = ctk.CTkFrame(parent, fg_color="transparent")
        options_container.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        options_container.grid_columnconfigure(0, weight=1)

        self.remove_phrase_options = ctk.CTkFrame(options_container, fg_color="transparent")
        self.remove_phrase_options.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(self.remove_phrase_options, text="Фраза / RegEx:").grid(row=0, column=0, sticky="w", padx=(0, 5), pady=5)
        phrase_entry = PlaceholderEntry(self.remove_phrase_options, textvariable=self.controller.phrase_var, placeholder="Введите фразу или регулярное выражение")
        phrase_entry.grid(row=0, column=1, sticky="ew", pady=5)
        self.controller.operation_buttons["phrase_entry"] = phrase_entry
        case_cb = ctk.CTkCheckBox(self.remove_phrase_options, text="Регистр", variable=self.controller.case_sensitive_phrase_var)
        case_cb.grid(row=0, column=2, padx=10)
        regex_cb = ctk.CTkCheckBox(self.remove_phrase_options, text="RegEx", variable=self.controller.use_regex_var)
        regex_cb.grid(row=0, column=3, padx=5)
        self.controller.operation_buttons.update({"phrase_case_cb": case_cb, "phrase_regex_cb": regex_cb})

        self.delete_urls_options = ctk.CTkFrame(options_container, fg_color="transparent")
        self.delete_urls_options.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(self.delete_urls_options, text="Имена URL (через ','):").grid(row=0, column=0, sticky="w", padx=(0, 5), pady=5)
        url_names_entry = PlaceholderEntry(self.delete_urls_options, textvariable=self.controller.url_names_var, placeholder="имя1, частьимени2")
        url_names_entry.grid(row=0, column=1, sticky="ew", pady=5)
        self.controller.operation_buttons["url_entry"] = url_names_entry
        url_case_cb = ctk.CTkCheckBox(self.delete_urls_options, text="Регистр", variable=self.controller.case_sensitive_url_var)
        url_case_cb.grid(row=0, column=2, sticky="w", padx=10)
        self.controller.operation_buttons["url_case_cb"] = url_case_cb

        self.file_op_option_frames = {"remove_phrase": self.remove_phrase_options, "delete_urls": self.delete_urls_options}
        self.options_group.grid_remove() # Hide by default

    def _on_file_op_selected(self):
        selected_op = self.selected_file_op.get()
        if not selected_op: return

        self.options_group.grid()
        op_name = self.controller.operations.get(selected_op, {}).get("name", "Выполнить")
        self.file_op_run_btn.configure(text=f"Выполнить: {op_name}", state="normal")
        for op, frame in self.file_op_option_frames.items():
            if op == selected_op: frame.grid(row=0, column=0, sticky="ew")
            else: frame.grid_remove()

        if selected_op not in self.file_op_option_frames:
            self.options_group.grid_remove()

    def _run_selected_file_op(self):
        if op_type := self.selected_file_op.get():
            self.controller.run_operation(op_type)
        else:
            messagebox.showwarning("Нет выбора", "Пожалуйста, сначала выберите операцию.", parent=self.controller)

class PathGeneratorView(ctk.CTkFrame):
    def __init__(self, master, controller, **kwargs):
        super().__init__(master, **kwargs)
        self.controller = controller
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(3, weight=2)

        ctk.CTkLabel(self, text="Генератор Путей для Excel", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, padx=20, pady=(10, 20), sticky="w")

        input_group = ctk.CTkFrame(self)
        input_group.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        input_group.grid_columnconfigure(0, weight=1)
        input_group.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(input_group, text="1. Введите названия моделей (каждое с новой строки)", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=15, pady=(15, 10), sticky="w")
        self.controller.path_gen_input_text = ctk.CTkTextbox(input_group, wrap="word")
        self.controller.path_gen_input_text.grid(row=1, column=0, sticky="nsew", padx=15, pady=(0, 15))

        generate_btn = ctk.CTkButton(self, text="✅ Сгенерировать и проверить пути", height=40, command=lambda: self.controller.run_operation("generate_paths"))
        generate_btn.grid(row=2, column=0, sticky="ew", pady=10, padx=20)
        self.controller.operation_buttons["gen_paths"] = generate_btn

        output_group = ctk.CTkFrame(self)
        output_group.grid(row=3, column=0, sticky="nsew", padx=20, pady=10)
        output_group.grid_columnconfigure(0, weight=1)
        output_group.grid_rowconfigure(1, weight=3)
        output_group.grid_rowconfigure(3, weight=1)
        ctk.CTkLabel(output_group, text="2. Результат", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=15, pady=(15, 10), sticky="w")
        self.controller.path_gen_output_text = ctk.CTkTextbox(output_group, wrap="none")
        self.controller.path_gen_output_text.grid(row=1, column=0, sticky="nsew", padx=15, pady=0)
        self.controller.path_gen_output_text.configure(state="disabled")
        ctk.CTkLabel(output_group, text="Ошибки", font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, padx=15, pady=(10, 5), sticky="w")
        self.controller.path_gen_error_text = ctk.CTkTextbox(output_group, wrap="none", height=80)
        self.controller.path_gen_error_text.grid(row=3, column=0, sticky="nsew", padx=15, pady=(0, 15))
        self.controller.path_gen_error_text.configure(state="disabled")
        copy_btn = ctk.CTkButton(output_group, text="Копировать успешные результаты", command=self.controller.copy_path_gen_results)
        copy_btn.grid(row=4, column=0, sticky="e", padx=15, pady=15)

class FolderCreatorView(ctk.CTkFrame):
    def __init__(self, master, controller, **kwargs):
        super().__init__(master, **kwargs)
        self.controller = controller
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self, text="Создатель Папок", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, padx=20, pady=(10, 20), sticky="w")

        input_group = ctk.CTkFrame(self)
        input_group.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        input_group.grid_columnconfigure(0, weight=1)
        input_group.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(input_group, text="1. Введите названия папок (каждое с новой строки)", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=15, pady=(15, 10), sticky="w")
        self.controller.folder_creator_input_text = ctk.CTkTextbox(input_group, wrap="word")
        self.controller.folder_creator_input_text.grid(row=1, column=0, sticky="nsew", padx=15, pady=(0, 15))
        Tooltip(self.controller.folder_creator_input_text, "Можно создавать вложенные папки, например: ProjectA/assets")

        options_group = ctk.CTkFrame(self)
        options_group.grid(row=2, column=0, sticky="ew", padx=20, pady=10)
        options_group.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(options_group, text="2. Опции", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=2, padx=15, pady=(15, 10), sticky="w")

        ctk.CTkLabel(options_group, text="Префикс:").grid(row=1, column=0, sticky="w", padx=(15, 5), pady=5)
        folder_prefix_entry = PlaceholderEntry(options_group, textvariable=self.controller.folder_prefix_var)
        folder_prefix_entry.grid(row=1, column=1, sticky="ew", padx=(0, 15), pady=5)

        ctk.CTkLabel(options_group, text="Суффикс:").grid(row=2, column=0, sticky="w", padx=(15, 5), pady=5)
        folder_suffix_entry = PlaceholderEntry(options_group, textvariable=self.controller.folder_suffix_var)
        folder_suffix_entry.grid(row=2, column=1, sticky="ew", padx=(0, 15), pady=5)

        numbering_cb = ctk.CTkCheckBox(options_group, text="Включить автонумерацию", variable=self.controller.folder_numbering_var)
        numbering_cb.grid(row=3, column=0, columnspan=2, sticky="w", padx=15, pady=10)

        num_opts_frame = ctk.CTkFrame(options_group, fg_color="transparent")
        num_opts_frame.grid(row=4, column=0, columnspan=2, sticky="w", padx=15, pady=5)
        ctk.CTkLabel(num_opts_frame, text="Начать с:").pack(side="left")
        start_spinbox = ctk.CTkEntry(num_opts_frame, textvariable=self.controller.folder_start_num_var, width=80)
        start_spinbox.pack(side="left", padx=(5, 20))
        ctk.CTkLabel(num_opts_frame, text="Цифр (padding):").pack(side="left")
        padding_spinbox = ctk.CTkEntry(num_opts_frame, textvariable=self.controller.folder_padding_var, width=60)
        padding_spinbox.pack(side="left", padx=5)

        create_btn = ctk.CTkButton(self, text="✅ Создать папки", height=40, command=lambda: self.controller.run_operation("create_folders"))
        create_btn.grid(row=3, column=0, sticky="ew", pady=20, padx=20)

        self.controller.operation_buttons.update({"folder_create_input": self.controller.folder_creator_input_text, "folder_prefix_entry": folder_prefix_entry, "folder_suffix_entry": folder_suffix_entry, "folder_numbering_cb": numbering_cb, "folder_start_spinbox": start_spinbox, "folder_padding_spinbox": padding_spinbox, "folder_create_btn": create_btn})

class ArticleConverterView(ctk.CTkFrame):
    def __init__(self, master, controller, **kwargs):
        super().__init__(master, **kwargs)
        self.controller = controller
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text="Конвертер Артикулов", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, padx=20, pady=(10, 20), sticky="w")

        container = ctk.CTkFrame(self)
        container.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        container.grid_columnconfigure(0, weight=1)

        self.controller.converter_select_btn = ctk.CTkButton(container, text="1. Выбрать Excel/CSV файл", height=40, command=self.controller.select_and_scan_converter_file)
        self.controller.converter_select_btn.grid(row=0, column=0, pady=15, padx=15, sticky="ew")

        self.controller.converter_file_label = ctk.CTkLabel(container, text="Файл не выбран", anchor="center")
        self.controller.converter_file_label.grid(row=1, column=0, pady=(0, 10), padx=15, sticky="ew")

        self.controller.converter_detected_label = ctk.CTkLabel(container, text="", font=ctk.CTkFont(weight="bold"), anchor="center")
        self.controller.converter_detected_label.grid(row=2, column=0, pady=(0, 15), padx=15, sticky="ew")

        ctk.CTkLabel(container, text="2. Выберите НОВЫЙ размер для замены:", anchor="w").grid(row=3, column=0, pady=(10, 5), padx=15, sticky="ew")

        self.controller.converter_size_combobox = ctk.CTkComboBox(container, state="disabled", values=[])
        self.controller.converter_size_combobox.grid(row=4, column=0, pady=(0, 15), padx=15, ipady=4, sticky="ew")

        self.controller.converter_process_btn = ctk.CTkButton(container, text="3. Создать файл с новым размером", height=40, command=self.controller.process_and_save_converter_file, state="disabled")
        self.controller.converter_process_btn.grid(row=5, column=0, pady=15, padx=15, sticky="ew")

        ctk.CTkFrame(container, height=2, fg_color="gray50").grid(row=6, column=0, sticky="ew", pady=15, padx=15)

        self.controller.converter_edit_btn = ctk.CTkButton(container, text="⚙️ Редактор размеров", command=self.controller.open_size_editor)
        self.controller.converter_edit_btn.grid(row=7, column=0, pady=(0, 15), padx=15, sticky="ew")

        self.controller.operation_buttons.update({"conv_select_btn": self.controller.converter_select_btn, "conv_process_btn": self.controller.converter_process_btn, "conv_edit_btn": self.controller.converter_edit_btn, "conv_combobox": self.controller.converter_size_combobox})
