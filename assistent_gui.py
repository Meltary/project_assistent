"""Модуль с графическим интерфейсом для сравнения файлов 1С и Лоцман."""
from tkinter import filedialog
from ttkbootstrap import Window, Frame, Label, Entry, Button, StringVar, Treeview, Notebook, Checkbutton
from ttkbootstrap.dialogs import MessageDialog

from assistent_logic import (
    compare_files_1c_locman,
    convert_xlsx_to_csv,
)
from assistent_db import (
    save_mappings,
    load_mappings,
    get_mappings_for_locman,
    init_database,
)


def show_error_dialog(title: str, message: str) -> None:
    """Показывает диалог с ошибкой."""
    MessageDialog(title=title, message=message, buttons=["OK"]).show()


def main() -> None:
    """Простой мастер из трёх шагов с переходами между окнами."""

    # Инициализируем базу данных при запуске
    init_database()
    
    # Загружаем сохраненные соответствия из БД
    saved_mappings = load_mappings()

    win = Window(title='Ассистент', size=[800, 600], themename='litera')
    win.place_window_center()

    # Служебные переменные
    file1_var = StringVar(value="")
    locman_display = StringVar(value="")
    locman_vars: list[StringVar] = []

    frames: list[Frame] = []  # шаги мастера (фреймы)
    matches_tree = None       # таблица совпадений на шаге 2
    nomatch_1c_tree = None    # таблица несовпадений 1С на шаге 2
    locman_list_tree = None   # список Лоцман во вкладке выбора
    # Данные для вкладки выбора соответствий
    locman_to_1c_options: dict[tuple[str, str], list[str]] = {}  # (name_loc, loc_path) -> [name_1c, ...]
    selected_mappings: dict[tuple[str, str], set[str]] = saved_mappings.copy()  # (name_loc, loc_path) -> {name_1c, ...}

    # Показывает нужный шаг и скрывает остальные
    def show_frame(index: int) -> None:
        for i, frame in enumerate(frames):
            frame.pack_forget()
            if i == index:
                frame.pack(fill='both', expand=True, padx=20, pady=20)

    # Диалог выбора файла и запись пути в StringVar
    def choose_file(target: StringVar, title: str) -> None:
        path = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if path:
            target.set(path)

    # Обновляет текстовое отображение списка файлов Лоцман
    def update_locman_display() -> None:
        locman_display.set("\n".join([v.get() for v in locman_vars if v.get().strip()]))

    # Выбор файла Лоцман для конкретной строки
    def choose_locman_file(target: StringVar) -> None:
        choose_file(target, "Файл Лоцман")
        update_locman_display()

    # Финальный диалог с перечнем выбранных файлов
    def finish() -> None:
        locman_files = [v.get() for v in locman_vars if v.get().strip()]
        MessageDialog(
            title="Готово",
            message=f"Файл 1: {file1_var.get()}\nФайлы Лоцман:\n" + "\n".join(locman_files),
            buttons=["OK"],
        ).show()

    # Конвертирует XLSX -> CSV и подготавливает результат сравнения (шаг 2)
    def step1_next() -> None:
        locman_files = [v.get() for v in locman_vars if v.get().strip()]
        if not file1_var.get().strip() or not locman_files:
            MessageDialog(title="Ошибка", message="Выберите файлы 1С и Лоцман.", buttons=["OK"]).show()
            return

        # 1С -> CSV
        file1_csv = convert_xlsx_to_csv(file1_var.get(), error_callback=show_error_dialog)
        file1_var.set(file1_csv)

        # Лоцман -> CSV
        locman_csvs: list[str] = []
        for path in locman_files:
            locman_csvs.append(convert_xlsx_to_csv(path, error_callback=show_error_dialog))

        # записываем обратно в vars (только для заполненных)
        out_idx = 0
        for v in locman_vars:
            if v.get().strip():
                v.set(locman_csvs[out_idx])
                out_idx += 1
        update_locman_display()

        # считаем совпадения/несовпадения и обновляем таблицы на шаге 2
        nonlocal matches_tree, nomatch_1c_tree
        if matches_tree is not None and nomatch_1c_tree is not None:
            # очищаем таблицы
            for item in matches_tree.get_children():
                matches_tree.delete(item)
            for item in nomatch_1c_tree.get_children():
                nomatch_1c_tree.delete(item)

            matches, nomatches, nomatches_1c, locman_options_dict = compare_files_1c_locman(
                file1_csv, locman_csvs, error_callback=show_error_dialog
            )
            
            # сохраняем опции для вкладки выбора
            nonlocal locman_to_1c_options
            locman_to_1c_options = locman_options_dict

            for name_1c, name_loc, loc_path in matches:
                matches_tree.insert("", "end", values=(name_1c, name_loc, loc_path))
            for name_1c in nomatches_1c:
                nomatch_1c_tree.insert("", "end", values=(name_1c,))
            
            # Заполняем список Лоцман во вкладке выбора
            nonlocal locman_list_tree
            if locman_list_tree is not None:
                # Очищаем список
                for item in locman_list_tree.get_children():
                    locman_list_tree.delete(item)

                # Показываем только те наименования Лоцман, которых нет во вкладке "Совпадения"
                # (т.е. только элементы без совпадений)
                for name_loc, loc_path in nomatches:
                    locman_list_tree.insert("", "end", values=(name_loc, loc_path))

        show_frame(1)

    # --- Шаг 1: выбор обоих файлов ---
    step1 = Frame(win)
    Label(step1, text="Шаг 1: выберите файлы").pack(anchor='w', pady=(0, 10))

    # Файл 1С
    Label(step1, text="Файл (1С):").pack(anchor='w')
    Entry(step1, textvariable=file1_var, width=70).pack(anchor='w', fill='x')
    Button(step1, text="Обзор 1С...", command=lambda: choose_file(file1_var, "Файл 1С")).pack(anchor='w', pady=(10, 10))

    # Файлы Лоцман (каждое нажатие "Добавить" создаёт новую строку)
    Label(step1, text="Файлы Лоцман:").pack(anchor='w')

    locman_container = Frame(step1)
    locman_container.pack(anchor="w", fill="x", pady=(0, 5))

    def add_locman_row(default_value: str = "") -> None:
        v = StringVar(value=default_value)
        locman_vars.append(v)

        row = Frame(locman_container)
        row.pack(fill="x", pady=3)

        Entry(row, textvariable=v, width=70).pack(side="left", fill="x", expand=True)
        Button(row, text="Обзор...", command=lambda vv=v: choose_locman_file(vv)).pack(side="left", padx=5)

        update_locman_display()

    Button(step1, text="Добавить файл Лоцман +", command=add_locman_row).pack(anchor='w', pady=(0, 10))
    add_locman_row("")

    Button(step1, text="Далее →", bootstyle="success", command=step1_next).pack(anchor='e', pady=10)
    frames.append(step1)

    # --- Шаг 2: проверка и сравнение файлов ---
    step2 = Frame(win)
    Label(step2, text="Шаг 2: сравнение файлов").pack(anchor='w', pady=(0, 10))
    Label(step2, text="Файл 1 (1С):").pack(anchor='w')
    Label(step2, textvariable=file1_var, bootstyle="secondary").pack(anchor='w', pady=(0, 5))
    Label(step2, text="Файлы Лоцман:").pack(anchor='w')
    Label(step2, textvariable=locman_display, bootstyle="secondary", justify="left").pack(anchor='w', pady=(0, 10))

    # вкладки "Совпадения" / "Не совпадения (Лоцман)" / "Не совпадения (1С)"
    notebook = Notebook(step2)
    notebook.pack(fill="both", expand=True, pady=(0, 10))

    tab_match = Frame(notebook)
    tab_nomatch_1c = Frame(notebook)
    tab_selection = Frame(notebook)  # вкладка для выбора соответствий
    notebook.add(tab_match, text="Совпадения")
    notebook.add(tab_selection, text="Выбор соответствий")
    notebook.add(tab_nomatch_1c, text="Не совпадения (1С)")

    # таблица совпадений
    matches_tree = Treeview(
        tab_match,
        columns=("name_1c", "name_loc"),
        show="headings",
        height=8,
    )
    matches_tree.heading("name_1c", text="Наименование (1С)")
    matches_tree.heading("name_loc", text="Наименование (Лоцман)")
    matches_tree.pack(fill="both", expand=True)

    # таблица не совпадений 1С
    nomatch_1c_tree = Treeview(
        tab_nomatch_1c,
        columns=("name_1c",),
        show="headings",
        height=8,
    )
    nomatch_1c_tree.heading("name_1c", text="Наименование (1С)")
    nomatch_1c_tree.pack(fill="both", expand=True)

    # --- Вкладка "Выбор соответствий" ---
    # Создаём контейнер с двумя панелями
    selection_container = Frame(tab_selection)
    selection_container.pack(fill="both", expand=True, padx=10, pady=10)

    # Левая панель: список Лоцман
    left_panel = Frame(selection_container)
    left_panel.pack(side="left", fill="both", expand=True, padx=(0, 5))
    Label(left_panel, text="Список из Лоцман:").pack(anchor='w', pady=(0, 5))
    
    locman_list_tree = Treeview(
        left_panel,
        columns=("name_loc", "file_loc"),
        show="headings",
        selectmode="browse",
    )
    locman_list_tree.heading("name_loc", text="Наименование (Лоцман)")
    locman_list_tree.heading("file_loc", text="Файл")
    locman_list_tree.column("name_loc", width=300)
    locman_list_tree.column("file_loc", width=200)
    locman_list_tree.pack(fill="both", expand=True)

    # Правая панель: варианты из 1С с чекбоксами
    right_panel = Frame(selection_container)
    right_panel.pack(side="right", fill="both", expand=True, padx=(5, 0))
    Label(right_panel, text="Варианты из 1С:").pack(anchor='w', pady=(0, 5))
    
    # Фрейм для чекбоксов
    checkboxes_frame = Frame(right_panel)
    checkboxes_frame.pack(fill="both", expand=True)
    
    # Внутренний фрейм для чекбоксов (будет прокручиваться при необходимости)
    checkboxes_inner = Frame(checkboxes_frame)
    checkboxes_inner.pack(fill="both", expand=True)
    
    # Словарь для хранения чекбоксов: (name_loc, loc_path) -> {name_1c: Checkbutton}
    checkboxes_dict: dict[tuple[str, str], dict[str, tuple[Checkbutton, StringVar]]] = {}

    def on_locman_select(event):
        """Обработчик выбора элемента Лоцман"""
        selection = locman_list_tree.selection()
        if not selection:
            return
        
        # Очищаем предыдущие чекбоксы
        for widget in checkboxes_inner.winfo_children():
            widget.destroy()
        checkboxes_dict.clear()
        
        # Получаем выбранный элемент
        item = locman_list_tree.item(selection[0])
        values = item["values"]
        if len(values) < 2:
            return
        
        name_loc = values[0]
        loc_path = values[1]
        loc_key = (name_loc, loc_path)
        
        # Получаем варианты из 1С для этого Лоцман
        options = locman_to_1c_options.get(loc_key, [])
        
        if not options:
            Label(checkboxes_inner, text="Нет вариантов из 1С", bootstyle="secondary").pack(anchor='w', pady=5)
            return
        
        # Инициализируем множество выбранных, если его ещё нет
        if loc_key not in selected_mappings:
            # Пытаемся загрузить из БД, если нет в памяти
            db_mappings = get_mappings_for_locman(name_loc, loc_path)
            selected_mappings[loc_key] = db_mappings.copy() if db_mappings else set()
        
        # Создаём чекбоксы для каждого варианта
        checkboxes_dict[loc_key] = {}
        for name_1c in options:
            var = StringVar()
            # Проверяем, был ли этот вариант уже выбран (в памяти или в БД)
            if name_1c in selected_mappings[loc_key]:
                var.set("1")
            
            cb = Checkbutton(
                checkboxes_inner,
                text=name_1c,
                variable=var,
                bootstyle="round-toggle",
            )
            cb.pack(anchor='w', pady=2)
            checkboxes_dict[loc_key][name_1c] = (cb, var)
    
    # Привязываем обработчик выбора
    locman_list_tree.bind("<<TreeviewSelect>>", on_locman_select)
    
    # Функция для сохранения выбранных соответствий
    def save_selections():
        """Сохраняет выбранные соответствия в память и базу данных"""
        # Собираем все выбранные соответствия из текущих чекбоксов
        for loc_key, checkboxes in checkboxes_dict.items():
            selected = set()
            for name_1c, (cb, var) in checkboxes.items():
                if var.get() == "1":
                    selected.add(name_1c)
            selected_mappings[loc_key] = selected
        
        # Сохраняем в базу данных
        try:
            saved_count = save_mappings(selected_mappings)
            MessageDialog(
                title="Сохранено", 
                message=f"Выбранные соответствия сохранены в базу данных.\nСохранено записей: {saved_count}", 
                buttons=["OK"]
            ).show()
        except Exception as e:
            MessageDialog(
                title="Ошибка сохранения", 
                message=f"Не удалось сохранить в базу данных:\n{str(e)}", 
                buttons=["OK"]
            ).show()
    
    Button(right_panel, text="Сохранить выбор", command=save_selections, bootstyle="success").pack(pady=10)

    nav2 = Frame(step2)
    nav2.pack(fill='x', pady=10)
    Button(nav2, text="← Назад", command=lambda: show_frame(0)).pack(side='left')
    Button(nav2, text="Далее →", bootstyle="success", command=lambda: show_frame(2)).pack(side='right')
    frames.append(step2)

    # --- Шаг 3 ---
    step3 = Frame(win)
    Label(step3, text="Шаг 3: подтвердите и запустите сравнение").pack(anchor='w', pady=(0, 10))
    Label(step3, textvariable=file1_var).pack(anchor='w')
    Label(step3, textvariable=locman_display, justify="left").pack(anchor='w', pady=(0, 10))
    nav3 = Frame(step3)
    nav3.pack(fill='x', pady=10)
    Button(nav3, text="← Назад", command=lambda: show_frame(1)).pack(side='left')
    Button(nav3, text="Сравнить", bootstyle="primary", command=finish).pack(side='right')
    frames.append(step3)

    # Показать первый шаг
    show_frame(0)

    # Запуск цикла обработки событий
    win.mainloop()


if __name__ == "__main__":
    main()
