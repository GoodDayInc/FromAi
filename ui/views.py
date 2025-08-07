import customtkinter as ctk
from ui.widgets import PlaceholderEntry, Tooltip

class FileOperationsView(ctk.CTkFrame):
    def __init__(self, master, controller, **kwargs):
        super().__init__(master, **kwargs)
        self.controller = controller

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Frame for operation selection
        selection_lf = ctk.CTkFrame(self)
        selection_lf.grid(row=0, column=0, sticky="ew", pady=(0, 10), padx=10)
        ctk.CTkLabel(selection_lf, text="1. Выберите операцию").pack(pady=5)

        self.selected_file_op = ctk.StringVar()
        btn_configs = [
            ("extract", "📤 Извлечь из '1'"),
            ("rename_images", "🔢 Переименовать 1-N"),
            ("remove_phrase", "✂️ Удалить фразу/RegEx"),
            ("delete_urls", "🗑️ Удалить URL-ярлыки"),
        ]

        radio_frame = ctk.CTkFrame(selection_lf)
        radio_frame.pack(fill="x", expand=True, padx=5, pady=5)
        for i, (op_type, text) in enumerate(btn_configs):
            radio_frame.grid_columnconfigure(i, weight=1)
            rb = ctk.CTkRadioButton(
                radio_frame,
                text=text,
                variable=self.selected_file_op,
                value=op_type,
                command=self._on_file_op_selected
            )
            rb.grid(row=0, column=i, padx=5, pady=5, sticky="ew")
            self.controller.operation_buttons[f"op_{op_type}"] = rb

        # Frame for contextual options
        self.file_ops_options_frame = ctk.CTkFrame(self)
        self.file_ops_options_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 10), padx=10)
        self._create_file_op_option_widgets(self.file_ops_options_frame)

        # --- Execution and Dry Run Frame ---
        exec_lf = ctk.CTkFrame(self)
        exec_lf.grid(row=2, column=0, sticky="ew", pady=(5, 0), padx=10)
        exec_lf.grid_columnconfigure(1, weight=1)

        self.controller.dry_run_var = ctk.BooleanVar(value=self.controller.last_dry_run)
        dry_run_cb = ctk.CTkCheckBox(exec_lf, text="✅ Пробный запуск (Dry Run)", variable=self.controller.dry_run_var)
        dry_run_cb.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        Tooltip(dry_run_cb, "Симулировать операцию в логе без реального изменения файлов. Настоятельно рекомендуется!")
        self.controller.operation_buttons["dry_run_cb"] = dry_run_cb

        self.file_op_run_btn = ctk.CTkButton(exec_lf, text="Выполнить", state="disabled", command=self._run_selected_file_op)
        self.file_op_run_btn.grid(row=0, column=1, padx=10, pady=10, sticky="e")
        self.controller.operation_buttons["run_file_op"] = self.file_op_run_btn

    def _create_file_op_option_widgets(self, parent: ctk.CTkFrame):
        parent.grid_columnconfigure(0, weight=1)

        # --- Remove Phrase Options ---
        self.remove_phrase_options = ctk.CTkFrame(parent, fg_color="transparent")
        self.remove_phrase_options.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.remove_phrase_options, text="Фраза / RegEx:").grid(row=0, column=0, sticky="w", padx=(0, 5), pady=5)
        self.controller.phrase_var = ctk.StringVar(value=self.controller.last_phrase_to_remove)
        phrase_entry = PlaceholderEntry(self.remove_phrase_options, textvariable=self.controller.phrase_var, placeholder="Введите фразу или регулярное выражение")
        phrase_entry.grid(row=0, column=1, sticky="ew", pady=5)
        self.controller.operation_buttons["phrase_entry"] = phrase_entry

        phrase_opts_frame = ctk.CTkFrame(self.remove_phrase_options, fg_color="transparent")
        phrase_opts_frame.grid(row=0, column=2, sticky="w", padx=(10, 0))
        self.controller.case_sensitive_phrase_var = ctk.BooleanVar(value=self.controller.last_case_sensitive_phrase)
        case_cb = ctk.CTkCheckBox(phrase_opts_frame, text="Регистр", variable=self.controller.case_sensitive_phrase_var)
        case_cb.pack(side="left")
        self.controller.use_regex_var = ctk.BooleanVar(value=self.controller.last_use_regex)
        regex_cb = ctk.CTkCheckBox(phrase_opts_frame, text="RegEx", variable=self.controller.use_regex_var)
        regex_cb.pack(side="left", padx=5)
        self.controller.operation_buttons["phrase_case_cb"] = case_cb
        self.controller.operation_buttons["phrase_regex_cb"] = regex_cb

        # --- Delete URLs Options ---
        self.delete_urls_options = ctk.CTkFrame(parent, fg_color="transparent")
        self.delete_urls_options.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.delete_urls_options, text="Имена URL (через ','):").grid(row=0, column=0, sticky="w", padx=(0, 5), pady=5)
        self.controller.url_names_var = ctk.StringVar(value=self.controller.last_url_names_to_delete)
        url_names_entry = PlaceholderEntry(self.delete_urls_options, textvariable=self.controller.url_names_var, placeholder="имя1, частьимени2")
        url_names_entry.grid(row=0, column=1, sticky="ew", pady=5)
        self.controller.operation_buttons["url_entry"] = url_names_entry

        self.controller.case_sensitive_url_var = ctk.BooleanVar(value=self.controller.last_case_sensitive_url)
        url_case_cb = ctk.CTkCheckBox(self.delete_urls_options, text="Регистр", variable=self.controller.case_sensitive_url_var)
        url_case_cb.grid(row=0, column=2, sticky="w", padx=10)
        self.controller.operation_buttons["url_case_cb"] = url_case_cb

        self.file_op_option_frames = {
            "remove_phrase": self.remove_phrase_options,
            "delete_urls": self.delete_urls_options,
        }

    def _on_file_op_selected(self):
        selected_op = self.selected_file_op.get()
        op_name = self.controller.operations.get(selected_op, {}).get("name", "Выполнить")
        self.file_op_run_btn.configure(text=f"Выполнить: {op_name}", state="normal")

        for op_type, frame in self.file_op_option_frames.items():
            if op_type == selected_op:
                frame.grid(row=0, column=0, sticky="ew")
            else:
                frame.grid_remove()

    def _run_selected_file_op(self):
        op_type = self.selected_file_op.get()
        if op_type:
            self.controller.run_operation(op_type)
        else:
            messagebox.showwarning("Нет выбора", "Пожалуйста, сначала выберите операцию.", parent=self.controller)

class PathGeneratorView(ctk.CTkFrame):
    def __init__(self, master, controller, **kwargs):
        super().__init__(master, **kwargs)
        self.controller = controller

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)  # Make input area resizable
        self.grid_rowconfigure(2, weight=2)  # Make output area resizable (with more weight)

        input_lf = ctk.CTkFrame(self)
        input_lf.grid(row=0, column=0, sticky="nsew", pady=(0, 10), padx=10)
        input_lf.grid_columnconfigure(0, weight=1)
        input_lf.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(input_lf, text="1. Введите названия моделей (каждое с новой строки)").pack(pady=5)

        self.controller.path_gen_input_text = ctk.CTkTextbox(input_lf, wrap="word") # Removed fixed height
        self.controller.path_gen_input_text.pack(fill="both", expand=True, padx=5, pady=5)
        Tooltip(self.controller.path_gen_input_text, "Вставьте сюда список моделей. Каждая модель на новой строке.")

        generate_btn = ctk.CTkButton(self, text="✅ Сгенерировать и проверить пути", command=lambda: self.controller.run_operation("generate_paths"))
        generate_btn.grid(row=1, column=0, sticky="ew", pady=5, padx=10)
        self.controller.operation_buttons["gen_paths"] = generate_btn

        output_lf = ctk.CTkFrame(self)
        output_lf.grid(row=2, column=0, sticky="nsew", padx=10)
        output_lf.grid_columnconfigure(0, weight=1)
        output_lf.grid_rowconfigure(0, weight=3)
        output_lf.grid_rowconfigure(1, weight=1)

        self.controller.path_gen_output_text = ctk.CTkTextbox(output_lf, wrap="none", height=6)
        self.controller.path_gen_output_text.grid(row=0, column=0, sticky="nsew", padx=5, pady=(5, 0))
        self.controller.path_gen_output_text.configure(state="disabled")

        self.controller.path_gen_error_text = ctk.CTkTextbox(output_lf, wrap="none", height=3)
        self.controller.path_gen_error_text.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.controller.path_gen_error_text.configure(state="disabled")

        copy_btn = ctk.CTkButton(output_lf, text="Копировать успешные результаты", command=self.controller.copy_path_gen_results)
        copy_btn.grid(row=2, column=0, sticky="e", padx=5, pady=5)

class FolderCreatorView(ctk.CTkFrame):
    def __init__(self, master, controller, **kwargs):
        super().__init__(master, **kwargs)
        self.controller = controller

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        input_lf = ctk.CTkFrame(self)
        input_lf.grid(row=0, column=0, sticky="nsew", pady=(0, 10), padx=10)
        input_lf.grid_columnconfigure(0, weight=1)
        input_lf.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(input_lf, text="1. Введите названия папок (каждое с новой строки)").pack(pady=5)

        self.controller.folder_creator_input_text = ctk.CTkTextbox(input_lf, wrap="word", height=5)
        self.controller.folder_creator_input_text.pack(fill="both", expand=True, padx=5, pady=5)
        Tooltip(self.controller.folder_creator_input_text, "Можно создавать вложенные папки, например: ProjectA/assets")

        options_lf = ctk.CTkFrame(self)
        options_lf.grid(row=1, column=0, sticky="ew", pady=(0, 10), padx=10)
        options_lf.grid_columnconfigure(1, weight=1)
        options_lf.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(options_lf, text="Префикс:").grid(row=0, column=0, sticky="w", padx=(10, 5), pady=5)
        self.controller.folder_prefix_var = ctk.StringVar(value=self.controller.last_folder_prefix)
        self.controller.folder_prefix_entry = PlaceholderEntry(options_lf, textvariable=self.controller.folder_prefix_var)
        self.controller.folder_prefix_entry.grid(row=0, column=1, sticky="ew", padx=(0, 5), pady=5)

        ctk.CTkLabel(options_lf, text="Суффикс:").grid(row=0, column=2, sticky="w", padx=(10, 5), pady=5)
        self.controller.folder_suffix_var = ctk.StringVar(value=self.controller.last_folder_suffix)
        self.controller.folder_suffix_entry = PlaceholderEntry(options_lf, textvariable=self.controller.folder_suffix_var)
        self.controller.folder_suffix_entry.grid(row=0, column=3, sticky="ew", padx=(0, 10), pady=5)

        self.controller.folder_numbering_var = ctk.BooleanVar(value=self.controller.last_folder_numbering)
        numbering_cb = ctk.CTkCheckBox(options_lf, text="Включить автонумерацию", variable=self.controller.folder_numbering_var)
        numbering_cb.grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=5)

        num_opts_frame = ctk.CTkFrame(options_lf, fg_color="transparent")
        num_opts_frame.grid(row=1, column=2, columnspan=2, sticky="w", padx=(10, 0))

        self.controller.folder_start_num_var = ctk.IntVar(value=self.controller.last_folder_start_num)
        ctk.CTkLabel(num_opts_frame, text="Начать с:").pack(side="left")
        start_spinbox = ctk.CTkEntry(num_opts_frame, textvariable=self.controller.folder_start_num_var, width=80)
        start_spinbox.pack(side="left", padx=(2, 10))

        self.controller.folder_padding_var = ctk.IntVar(value=self.controller.last_folder_padding)
        ctk.CTkLabel(num_opts_frame, text="Цифр (padding):").pack(side="left")
        padding_spinbox = ctk.CTkEntry(num_opts_frame, textvariable=self.controller.folder_padding_var, width=60)
        padding_spinbox.pack(side="left", padx=2)

        create_btn = ctk.CTkButton(self, text="✅ Создать папки", command=lambda: self.controller.run_operation("create_folders"))
        create_btn.grid(row=2, column=0, sticky="ew", pady=5, padx=10)

        self.controller.operation_buttons.update({
            "folder_create_input": self.controller.folder_creator_input_text,
            "folder_prefix_entry": self.controller.folder_prefix_entry,
            "folder_suffix_entry": self.controller.folder_suffix_entry,
            "folder_numbering_cb": numbering_cb,
            "folder_start_spinbox": start_spinbox,
            "folder_padding_spinbox": padding_spinbox,
            "folder_create_btn": create_btn,
        })

class ArticleConverterView(ctk.CTkFrame):
    def __init__(self, master, controller, **kwargs):
        super().__init__(master, **kwargs)
        self.controller = controller

        self.grid_columnconfigure(0, weight=1)

        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=20, pady=20)
        container.grid_columnconfigure(0, weight=1)

        self.controller.converter_select_btn = ctk.CTkButton(container, text="1. Выбрать Excel/CSV файл", command=self.controller.select_and_scan_converter_file)
        self.controller.converter_select_btn.grid(row=0, column=0, pady=5, ipady=5, sticky="ew")

        self.controller.converter_file_label = ctk.CTkLabel(container, text="Файл не выбран", anchor="center")
        self.controller.converter_file_label.grid(row=1, column=0, pady=2, sticky="ew")

        self.controller.converter_detected_label = ctk.CTkLabel(container, text="", font=ctk.CTkFont(weight="bold"), anchor="center")
        self.controller.converter_detected_label.grid(row=2, column=0, pady=5, sticky="ew")

        ctk.CTkLabel(container, text="2. Выберите НОВЫЙ размер для замены:", anchor="center").grid(row=3, column=0, pady=(10, 0), sticky="ew")

        self.controller.converter_size_combobox = ctk.CTkComboBox(container, state="disabled", values=[])
        self.controller.converter_size_combobox.grid(row=4, column=0, pady=5, ipady=3, sticky="ew")

        self.controller.converter_process_btn = ctk.CTkButton(container, text="3. Создать файл с новым размером", command=self.controller.process_and_save_converter_file, state="disabled")
        self.controller.converter_process_btn.grid(row=5, column=0, pady=5, ipady=5, sticky="ew")

        # Separator is just a frame with height
        ctk.CTkFrame(container, height=2, fg_color="gray50").grid(row=6, column=0, sticky="ew", pady=20)

        self.controller.converter_edit_btn = ctk.CTkButton(container, text="⚙️ Редактор размеров", command=self.controller.open_size_editor)
        self.controller.converter_edit_btn.grid(row=7, column=0, pady=10, sticky="ew")

        self.controller.operation_buttons.update({
            "conv_select_btn": self.controller.converter_select_btn,
            "conv_process_btn": self.controller.converter_process_btn,
            "conv_edit_btn": self.controller.converter_edit_btn,
            "conv_combobox": self.controller.converter_size_combobox,
        })
