"""
    Программный код интерфейса для программы "Автоматические Документы"
    Автор программного кода: Кирилл Жосул.
"""

# Импортация модулей.

# Встроенные модули
from os.path import exists  # Функция для проверки существования файла или папки.
from os import getcwd  # Функция получения пути приложения.

# Модули внешние.

# Виджеты PyQt5.
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QLabel
from PyQt5.QtWidgets import QTextEdit
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QCheckBox
from PyQt5.QtWidgets import QFormLayout
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtCore import QRect
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt

# Модули ядра.
import generator  # Модуль для работы с генерацией файлов.


class Window(QWidget):
    # Класс окна, наследованного от QWidget как виджет PyQt5.

    def __init__(self):
        # Класс конструктора.

        # Вызов инициализации родителя.
        super().__init__()

        # Переменные выбора.
        self.__selected_file_table = ""
        self.__selected_file_image = ""
        self.__selected_directory_out = ""

        # Поток генерации.
        self.__generator_thread = None

        # Другое.
        self.__checkbox_switch_overwrite = None
        self.__textedit_directory_rule = None
        self.__textedit_filename_rule = None
        self.__label_select_file_table = None
        self.__label_select_file_image = None
        self.__label_select_directory_out = None

        # Инициализация интерфейса.
        self.initialise_interface()

    def initialise_interface(self):
        # Функция инициализации интерфейса.

        # Инициализация.
        self.interface_window()
        self.interface_layout()

    def interface_window(self):
        # Функция настройки окна.

        # Настройка окна.
        self.setFixedSize(254, 522)
        self.setWindowTitle("Автоматические документы")
        self.setWindowIcon(QIcon("icon.png"))

    def interface_layout(self):
        # Функция настройки формы.

        # Создание формы.

        # Объект формы.
        layout = QFormLayout()

        # Добавление заголовка.
        layout.addRow(QLabel("Автоматические документы."))
        layout.addRow(QLabel("Версия: 1.0, Автор: Кирилл Жосул."))
        layout.addRow(QLabel("Создано для Самарского Хакатона 2021."))

        # Пустое пространство.
        layout.addRow(QLabel())

        # Заголовок.
        label = QLabel("Настройка перед запуском")
        label.setAlignment(Qt.AlignCenter)
        layout.addRow(label)

        # Пустое пространство.
        layout.addRow(QLabel())

        # Кнопка выбора пути к таблице.
        button_select_file_table = QPushButton("...")
        self.__label_select_file_table = QLabel("Выберите файл таблицы!")
        layout.addRow(button_select_file_table, self.__label_select_file_table)
        button_select_file_table.clicked.connect(self.button_file_table)

        # Кнопка выбора пути к шаблону.
        button_select_file_image = QPushButton("...")
        self.__label_select_file_image = QLabel("Выберите файл шаблона!")
        layout.addRow(button_select_file_image, self.__label_select_file_image)
        button_select_file_image.clicked.connect(self.button_file_image)

        # Кнопка выбора пути к папке вывода.
        button_select_directory_out = QPushButton("...")
        self.__label_select_directory_out = QLabel("Выберите папку результата!")
        layout.addRow(button_select_directory_out, self.__label_select_directory_out)
        button_select_directory_out.clicked.connect(self.button_directory_out)

        # Ввод правила именований.
        layout.addRow(QLabel("Правило именования папок."))
        self.__textedit_directory_rule = QTextEdit("{name}")
        layout.addRow(self.__textedit_directory_rule)
        layout.addRow(QLabel("Правило именования файлов."))
        self.__textedit_filename_rule = QTextEdit("{name} {type} {givenfor} ({date})")
        layout.addRow(self.__textedit_filename_rule)

        # Чекбокс перезаписи.
        self.__checkbox_switch_overwrite = QCheckBox("Включить перезапись существуюших")
        layout.addRow(self.__checkbox_switch_overwrite)

        # Кнопка справки.
        button_documentation_rules = QPushButton("Справка по правилам именования")
        layout.addRow(button_documentation_rules)
        button_documentation_rules.clicked.connect(self.button_documentation_rules)

        # Кнопка подтверждение.
        button_generate = QPushButton("Запустить генерацию")
        layout.addRow(button_generate)
        button_generate.clicked.connect(self.button_generate)

        # Добавление формы.
        self.setLayout(layout)

    def button_directory_out(self):
        # Функция обработки кнопки выбора папки вывода.

        # Получение папки от пользователя.
        file = QFileDialog.getExistingDirectory(self, caption='Выберите папку результата', directory=getcwd())

        if len(file) != 0:
            # Если мы выбрали папку.

            # Выбор папки.
            self.__selected_directory_out = file

            # Установка текста.
            directories = file.split("/")
            self.__label_select_directory_out.setText("/" + directories[len(directories) - 1])

    def button_file_table(self):
        # Функция обработки кнопки выбора файла таблицы.

        # Получение файлов от пользователя.
        file = QFileDialog.getOpenFileName(self, caption='Выберите файла таблицы',
                                           directory=getcwd(), filter="Excel files (*.xlsx)")
        file = file[0]

        if len(file) != 0:
            # Если мы выбрали файл.

            # Выбор файла.
            self.__selected_file_table = file

            # Установка текста.
            directories = file.split("/")
            self.__label_select_file_table.setText(directories[len(directories) - 1])

    @staticmethod
    def button_documentation_rules():
        # Функция пока справки по документации.
        information("Справка по правилам именования:\n"
                    "{type} - Тип документа (Грамота или Другое),\n"
                    "{name} - Имя получателя,\n"
                    "{school} - Учреждение,\n"
                    "{event} - Мероприятие,\n"
                    "{givenfor} - За что дано,\n"
                    "{date} - Дата.")

    def button_file_image(self):
        # Функция обработки кнопки выбора файла шаблона.

        # Получение файлов от пользователя.
        file = QFileDialog.getOpenFileName(self, caption='Выберите файла шаблона',
                                           directory=getcwd(), filter="Image files (*.png *.jpg)")
        file = file[0]

        if len(file) != 0:
            # Если мы выбрали файл.

            # Выбор файла.
            self.__selected_file_image = file

            # Установка текста.
            directories = file.split("/")
            self.__label_select_file_image.setText(directories[len(directories) - 1])

    def button_generate(self):
        # Функция кнопки генерации.
        if self.__generator_thread is None:
            # Если поток ещё не был запущен.
            if exists(self.__selected_file_table):
                if exists(self.__selected_file_image):
                    if exists(self.__selected_directory_out):
                        # Если всё ОК.

                        # Получение потока генерации (не запущенного).
                        self.__generator_thread = generator.generate(self.__selected_file_table,
                                                                     self.__selected_file_image,
                                                                     self.__selected_directory_out,
                                                                     self.__textedit_directory_rule.toPlainText(),
                                                                     self.__textedit_filename_rule.toPlainText(),
                                                                     self.__checkbox_switch_overwrite.isChecked()
                                                                     )

                        # Запуск потока генерации.
                        self.__generator_thread.start()
                    else:
                        # Показ ошибки.
                        error("Вы не выбрали папку результата,\nили она была перемещёна.")
                else:
                    # Показ ошибки.
                    error("Вы не выбрали файл шаблона,\nили он был перемещён.")
            else:
                # Показ ошибки.
                error("Вы не выбрали файл таблицы,\nили он был перемещён.")
        else:
            # Перезапуск если поток окончен.
            if not self.__generator_thread.is_alive():
                self.__generator_thread = None
                self.button_generate()
                return

            # Показ информации.
            information("Генерация уже запущена.")

    def close_thread(self):
        # Закрывает поток, если он работает.
        if self.__generator_thread is not None:
            if self.__generator_thread.is_alive():
                self.__generator_thread.terminate()


def information(text):
    # Функция отображения информации.
    message = QMessageBox()
    message.setWindowTitle("Информация")
    message.setText(text)
    message.setIcon(QMessageBox.Information)
    message.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
    message.exec_()


def error(text):
    # Функция отображения ошибки.
    message = QMessageBox()
    message.setWindowTitle("Произошла ошибка!")
    message.setText(text)
    message.setIcon(QMessageBox.Warning)
    message.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
    message.exec_()


def start(args=None):
    # Функия для старта интерфейса.

    # Обработка параметров.
    if args is None:
        args = []

    # Создание приложения PyQt5
    global application
    application = QApplication(args)

    # Создание окна и его отображения.
    global window
    window = Window()
    window.show()

    # Запуск приложения.
    application.exec()

    # Код дойдёт до сюда, когда программа закроется.

    # Остановка потока обработки.
    window.close_thread()