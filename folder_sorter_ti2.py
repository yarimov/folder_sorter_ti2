import os
import glob
import shutil
import ctypes

import customtkinter
from customtkinter import filedialog as fd
import yaml
from yaml.loader import SafeLoader

customtkinter.set_appearance_mode("System")
# Цветовая схема окон приложения: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")
# Основной цвет оформления приложения: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # Конфигурация окна приложения
        self.title("Титан2 - Сортировщик папок")
        self.geometry(f"{600}x{580}")
        self.resizable(width=False, height=False)
        self.iconbitmap('logo_app.ico')

        # Свойство для отслеживания комбинаций "copy", "paste" на кириллице в customtkinter
        self.bind("<Control-KeyPress>", self.keypress)

        # Первые символы в названии файлов, которые будут исключены из проверок и сортировок (Служебные файлы)
        self.ignored_in_filename = ('~', '!')
        # Форматы файлов и соответствующие им наименования папок для их последующей сортировки (по умолчанию)
        self.format_directory = {
            'doc': 'WORD',
            'docx': 'WORD',
            'dot': 'WORD',
            'dotx': 'WORD',
            'xlsx': 'EXCEL',
            'xlsm': 'EXCEL',
            'dwg': 'DWG',
            'pdf': 'PDF'
        }

        # Создание основных форм и кнопок
        self.entry_folder_from = customtkinter.CTkEntry(self, placeholder_text="Введите путь к папке проекта",
                                                        width=400)
        self.entry_folder_from.grid(row=0, column=0, columnspan=1, padx=(20, 0), pady=(20, 0), sticky="n")
        self.entry_folder_to = customtkinter.CTkEntry(self, placeholder_text="Введите путь для сохранения", width=400)
        self.entry_folder_to.grid(row=0, column=0, columnspan=1, padx=(20, 0), pady=(60, 0), sticky="n")
        self.button_select_folder_from = customtkinter.CTkButton(master=self, text="Обзор",
                                                                 command=self.select_folder_from)
        self.button_select_folder_from.grid(row=0, column=2, padx=(20, 0), pady=(20, 0), sticky="n")
        self.button_select_folder_to = customtkinter.CTkButton(master=self, text="Обзор", command=self.select_folder_to)
        self.button_select_folder_to.grid(row=0, column=2, padx=(20, 0), pady=(60, 0), sticky="n")
        self.button_check_project = customtkinter.CTkButton(master=self, text="Проверить",
                                                            command=self.button_check_project_event)
        self.button_check_project.grid(row=1, column=2, padx=(20, 0), pady=(20, 0), sticky="n")
        self.button_select_sort = customtkinter.CTkButton(master=self, text="Сортировать",
                                                          command=self.button_select_sort_event)
        self.button_select_sort.grid(row=1, column=2, padx=(20, 0), pady=(60, 0), sticky="n")
        self.checkbox_default_folder = customtkinter.CTkCheckBox(master=self, text="Путь сохранения по умолчанию",
                                                                 command=self.checkbox_default_folder_event)
        self.checkbox_default_folder.grid(row=0, column=0, columnspan=1, pady=(100, 0), padx=(0, 150))
        self.textbox = customtkinter.CTkTextbox(self, width=350, height=420)
        self.textbox.grid(row=1, column=0, padx=(20, 0), pady=(20, 0), sticky="nsew")

        # Установка значений по умолчанию
        self.textbox.insert("0.0", "")
        self.textbox.configure(state="disabled")

    @staticmethod
    def keypress(event):
        """ Метод отслеживает сочетания клавиш 'CTL + C', 'CTL + V' и 'CTL + X'
        и выдает соответствующие события '<<Copy>>', '<<Paste>>, '<<Cut>>'.

        """

        usr = ctypes.windll.LoadLibrary("user32.dll")
        pf = getattr(usr, "GetKeyboardLayout")
        # Запрос текущей раскладки, '0x4190419' - это кириллица
        if hex(pf(0)) == '0x4190419':
            if event.keycode == 86:
                event.widget.event_generate('<<Paste>>')
            elif event.keycode == 67:
                event.widget.event_generate('<<Copy>>')
            elif event.keycode == 88:
                event.widget.event_generate('<<Cut>>')

    @staticmethod
    def create_dir(path_to_dir: str) -> None:
        """ Метод создает новую папку, если она не была создана ранее. """
        if not os.path.exists(path_to_dir):
            os.mkdir(path_to_dir)

    @staticmethod
    def sort_by_format(file_format: str, folder_from: str) -> list:
        """ Метод формирует список файлов из паки 'folder_from', соответствующих формату 'file_format'.
        * - для выбора всех форматов.

        """
        list_files = []
        for filename in glob.glob(f"{folder_from}/*.{file_format}"):
            list_files.append(filename)
        return list_files

    @staticmethod
    def load_config() -> dict:
        """ Метод осуществляет загрузку параметров из файла конфигурации. """
        try:
            with open('config.yml') as f:
                return yaml.load(f, Loader=SafeLoader)
        except FileNotFoundError:
            return {"DEFAULT_FOLDER": '', 'IGNORED_IN_FILENAME': ('*',)}

    def check_project(self, folder_from: str) -> dict:
        """ Метод осуществляет проверку наличия PDF файлов к файлам из папки 'folder_from'.
        При наличии PDF файлов к файлам из папки 'folder_from' присваивается статус [Ok] или [None].

        """
        list_dir_pdf_files = []
        for dir_pdf_file in self.sort_by_format('pdf', folder_from):
            list_dir_pdf_files.append(os.path.splitext(os.path.basename(dir_pdf_file))[0])
        check_list_all_files = {}
        for dir_all_file in self.sort_by_format('*', folder_from):
            file_format = os.path.splitext(os.path.basename(dir_all_file))[1].upper()
            file_name = os.path.splitext(os.path.basename(dir_all_file))[0].upper()
            if file_format != ".PDF" and file_name[0] not in self.ignored_in_filename:
                # Исключает служебные документы в начале которых содержатся спец. символы из "ignored_in_filename"
                if os.path.splitext(os.path.basename(dir_all_file))[0] in list_dir_pdf_files:
                    check_list_all_files[os.path.basename(dir_all_file)] = "Ok"
                else:
                    check_list_all_files[os.path.basename(dir_all_file)] = "None"
        return check_list_all_files

    def display_info_to_textbox(self, text_info: str) -> None:
        """ Метод отображает информационное сообщение 'text_info' в текстовом поле 'textbox' приложения. """
        self.textbox.configure(state="normal")
        self.textbox.delete('0.0', 'end')
        self.textbox.insert('0.0', f'{text_info}')
        self.textbox.configure(state="disabled")

    def get_check_results(self, check_list_all_files: dict) -> str:
        """ Метод формирует строку с отчетом о результате проверок о наличии PDF к файлам из проверяемой папки. """
        self.textbox.configure(state="normal")
        self.textbox.delete('0.0', 'end')
        text_info = 'Отчет о наличии PDF к следующим файлам: \n \n'
        for file_name in check_list_all_files:
            text_info += f'{file_name}   [{check_list_all_files.get(file_name)}]\n'
        return text_info

    def sort_files(self, folder_from: str, folder_to: str) -> str:
        """ Метод сортирует файлы из папки 'folder_from' в 'folder_to' по следующему алгоритму:
        в папке 'folder_to' создается папка с именем 'folder_from'. В ней создаются папки 'WORD' 'EXCEL', 'DWG', 'PDF',
        соответствующие форматам файлов ('.doc','.docx','.dot','.dotx'),
        ('.xlsx','.xlsm'), ('.dwg') и ('pdf') соответственно. Все файлы сортируются в папки соответствующим по формату,
        кроме PDF файлов содержащих в названии '-UL', '_UL', '-RVI', '_RVI'. Они сохраняются в коревой
        папке 'folder_from'.

        """

        for file_format in self.format_directory:
            for filename in glob.glob(f"{folder_from}/*.{file_format}"):
                path_to_dir = f"{folder_to}\\{self.format_directory.get(file_format)}"
                self.create_dir(path_to_dir)
                if os.path.basename(filename)[0] not in self.ignored_in_filename:
                    if os.path.isfile(os.path.join(folder_from, os.path.basename(filename))):
                        if file_format == "pdf":
                            if "_UL" in filename.upper() or "-UL" in filename.upper() or \
                                    "_RVI" in filename.upper() or "-RVI" in filename.upper():
                                shutil.copy(os.path.join(folder_from, os.path.basename(filename)),
                                            os.path.join(folder_to, os.path.basename(filename)))
                            else:
                                shutil.copy(os.path.join(folder_from, os.path.basename(filename)),
                                            os.path.join(path_to_dir, os.path.basename(filename)))
                        else:
                            shutil.copy(os.path.join(folder_from, os.path.basename(filename)),
                                        os.path.join(path_to_dir, os.path.basename(filename)))
        return "Копирование завершено! \n \n"

    def select_folder_from(self) -> None:
        """ Метод выбора директории 'folder_from' и отображения в поле 'entry_folder_from' """
        folder_from = fd.askdirectory()
        if folder_from:
            self.entry_folder_from.delete(0, 'end')
            self.entry_folder_from.insert(0, folder_from)

    def select_folder_to(self) -> None:
        """ Метод выбора директории 'folder_to' и отображения в поле 'entry_folder_to.insert' """
        folder_to = fd.askdirectory()
        if folder_to:
            self.entry_folder_to.delete(0, 'end')
            self.entry_folder_to.insert(0, folder_to)

    def button_check_project_event(self):
        """ Метод запускает логику проверки файлов проекта при нажатии на кнопку "ПРОВЕРИТЬ".

        """

        self.ignored_in_filename = self.load_config().get('IGNORED_IN_FILENAME', ('*',))
        # Обновляет параметр из файла конфигурации. Символы для служебных файлов в начале имени файла.

        if self.load_config().get('FORMAT_DIRECTORY', False):
            self.format_directory = self.load_config().get('FORMAT_DIRECTORY', {})
            # Обновляет параметр из файла конфигурации. Таблица соотношения форматов файлов и директорий для них.

        folder_from = self.entry_folder_from.get()
        if folder_from:
            # Получает путь к проекту и проверяет на корректность ввода.
            if os.path.exists(folder_from):
                check_list_all_files = self.check_project(folder_from)
                # Результат проверки о наличии PDF файлов к файлам из папки 'folder_from'
                text_info = self.get_check_results(check_list_all_files)
                # Результат проверки о наличии PDF файлов в формате строки для вывода в тестовую форму 'textbox'
                self.display_info_to_textbox(text_info)
                # Отображает текстовое сообщение пользователю
            else:
                self.display_info_to_textbox('Ошибка! Указанный путь к проекту не корректен.')
        else:
            self.display_info_to_textbox('Ошибка! Укажите путь к проекту.')

    def button_select_sort_event(self):
        """ Метод запускает логику сортировки файлов проекта при нажатии на кнопку "СОРТИРОВАТЬ".

        """
        self.ignored_in_filename = self.load_config().get('IGNORED_IN_FILENAME', ('*',))
        # Обновляет параметр из файла конфигурации. Символы для служебных файлов в начале имени файла.
        if self.load_config().get('FORMAT_DIRECTORY', False):
            self.format_directory = self.load_config().get('FORMAT_DIRECTORY', {})
            # Обновляет параметр из файла конфигурации. Таблица соотношения форматов файлов и директорий для них.

        folder_from = self.entry_folder_from.get()
        folder_to = self.entry_folder_to.get()

        # Проверки на правильность заполнения форм 'folder_from' и 'folder_to'. Вывод сообщений в случае ошибок.
        # В случае успешных проверок выполняется сортировка файлов.
        if folder_from and os.path.exists(folder_from) and folder_from != folder_to:
            if folder_to or folder_to == "":
                if os.path.exists(folder_to) or folder_to == "":
                    if folder_to == "":
                        folder_to = os.path.abspath(
                            f'{folder_from}\\{os.path.basename(folder_from)}\\..\\{os.path.basename(folder_from)}')
                        self.create_dir(folder_to)
                        self.sort_files(folder_from=folder_from, folder_to=folder_to)
                        self.display_info_to_textbox(f'Файлы скопированы в папку: \n {folder_to}')
                    else:
                        folder_to = os.path.abspath(
                            f'{folder_to}\\{os.path.basename(folder_from)}')
                        self.create_dir(folder_to)
                        self.sort_files(folder_from=folder_from, folder_to=folder_to)
                        self.display_info_to_textbox(f'Файлы скопированы в папку: \n {folder_to}')
                else:
                    self.display_info_to_textbox('Ошибка! Указанный путь для сохранения не корректен.')
            else:
                self.display_info_to_textbox('Ошибка! Укажите путь для сохранения проекта.')
        elif not os.path.exists(folder_from):
            self.display_info_to_textbox('Ошибка! Указанный путь к проекту не корректен.')
        elif folder_from == folder_to:
            folder_to = os.path.abspath(
                f'{folder_from}\\{os.path.basename(folder_from)}\\{os.path.basename(folder_from)}')
            self.create_dir(folder_to)
            self.sort_files(folder_from=folder_from, folder_to=folder_to)
            self.display_info_to_textbox(f'Файлы скопированы в папку: \n {folder_to}')
        else:
            self.display_info_to_textbox('Ошибка! Укажите путь к проекту.')

    def checkbox_default_folder_event(self):

        if self.entry_folder_to.cget('state') == 'normal':
            self.entry_folder_to.delete(0, 'end')
            self.entry_folder_to.insert(0, (self.load_config().get("DEFAULT_FOLDER", '')).replace("\\", "/"))
            self.entry_folder_to.configure(state="disabled", text_color='#999999')
            self.button_select_folder_to.configure(state="disabled")
        else:
            self.entry_folder_to.configure(state="normal", text_color='#DCE4EE')
            self.button_select_folder_to.configure(state="normal")
            self.entry_folder_to.delete(0, 'end')
            self.entry_folder_to.configure(placeholder_text="Введите путь к папке")


if __name__ == "__main__":
    app = App()
    app.mainloop()
