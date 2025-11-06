import pyautogui
import pygetwindow as gw
import time
import os
import pyperclip
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QLabel, QMessageBox
from PyQt6.QtCore import QThread, pyqtSignal, pyqtSlot
import sys
from openpyxl import load_workbook
import keyboard

COORDS_FILE = 'coords/button_coords.txt'
LOGIN_BUTTON_FILE = 'coords/login_button_coords.txt'
LOGIN_FIELD_FILE = 'coords/login_field_coords.txt'
PASSWORD_FIELD_FILE = 'coords/password_field_coords.txt'
SECOND_LOGIN_BUTTON_FILE = 'coords/second_login_button_coords.txt'
ERROR_PASSWORD_FIELD_FILE = 'coords/error_password_field_coords.txt'
ERROR_SECOND_LOGIN_BUTTON_FILE = 'coords/error_second_login_button_coords.txt'
MIMIC_WINDOW_TITLE = 'Mimic'
URL_TO_OPEN = 'https://promo.ladbrokes.com/en/promo/bspin/INSTANTSPINS'
EXCEL_FILE = 'excel/users.xlsx'
ERROR_IMAGE = 'screenshots/error_image.png'
ERROR_COORDS_FILE = 'coords/error_coords.txt'


def save_coords(coords, file_name):
    try:
        os.makedirs(os.path.dirname(file_name), exist_ok=True)
        with open(file_name, 'w') as f:
            f.write(f'{coords[0]},{coords[1]}')
            print(f'Координаты сохранены в файл {file_name}: {coords}')
    except Exception as e:
        print(f'Ошибка при сохранении координат в {file_name}: {e}')


def load_coords(file_name):
    if not os.path.exists(file_name):
        print(f'Файл {file_name} не существует.')
        return None
    try:
        with open(file_name, 'r') as f:
            x, y = map(int, f.read().strip().split(','))
            return (x, y)
    except Exception as e:
        print(f'Ошибка при загрузке координат из {file_name}: {e}')
        return None


def click_button(file_name):
    coords = load_coords(file_name)
    if coords:
        print(f'Нажимаем на кнопку или поле по координатам: {coords}')
        pyautogui.click(coords)
    return None


def read_excel(file_name):
    from openpyxl import load_workbook
    workbook = load_workbook(filename=file_name)
    sheet = workbook.active
    users = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        login = row[0] if row[0] else ''
        password = row[1] if row[1] else ''
        column_c = row[2] if row[2] else ''
        column_e = row[4] if row[4] else ''
        users.append({'login': login, 'password': password, 'column_c': column_c, 'column_e': column_e})
    return users


def paste_data_to_excel(file_path, current_login):
    try:
        wb = load_workbook(file_path)
        sheet = wb.active
        copied_data = pyperclip.paste()
        if not copied_data:
            print('Буфер обмена пуст. Скопируйте данные перед запуском.')
        return None
    except Exception as e:
        print(f'Ошибка при работе с Excel: {e}')


def paste_data_to_excel2(file_path, current_login):
    try:
        wb = load_workbook(file_path)
        sheet = wb.active
        copied_data = pyperclip.paste()
        if not copied_data:
            print('Буфер обмена пуст. Скопируйте данные перед запуском.')
        return None
    except Exception as e:
        print(f'Ошибка при работе с Excel: {e}')


def enter_credentials(login, password, error_coords=False):
    if not error_coords:
        print('Нажимаем первую кнопку Login')
        click_button(LOGIN_BUTTON_FILE)
        time.sleep(1)
        print(f'Вводим логин: {login}')
        click_button(LOGIN_FIELD_FILE)
        pyautogui.typewrite(login, interval=0.1)
    print(f'Вводим пароль: {password}')
    if error_coords:
        click_button(ERROR_PASSWORD_FIELD_FILE)
    pyautogui.typewrite(password, interval=0.1)
    print('Нажимаем вторую кнопку Login')
    if error_coords:
        click_button(ERROR_SECOND_LOGIN_BUTTON_FILE)
    return None


def request_new_coords(file_name, parent=None):
    print(f'Пожалуйста, укажите новые координаты для {file_name}.')
    print('Сначала наведите курсор на нужную область и нажмите \'Enter\'.')
    input('Нажмите Enter, когда будете готовы задать координаты.')
    time.sleep(1)
    x, y = pyautogui.position()
    save_coords((x, y), file_name)
    print(f'Координаты для {file_name} сохранены: {x}, {y}')
    if parent:
        QMessageBox.information(parent, 'Успех', f'Координаты для {file_name} обновлены.')


def request_krestik_coords(parent=None):
    reply = QMessageBox.question(parent, 'Обновление координат',
                                 'Наведите курсор на крестик модального окна и нажмите \'Enter\'.\n\nХотите продолжить?',
                                 QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
    if reply == QMessageBox.StandardButton.Ok:
        request_new_coords('coords/krestik.txt', parent)
    else:
        QMessageBox.information(parent, 'Отмена', 'Обновление координат отменено.')


def request_multilogin_button_coords(parent=None):
    reply = QMessageBox.question(parent, 'Обновление координат',
                                 'Наведите курсор на кнопку запуска разового профиля и нажмите \'Enter\'.\n\nХотите продолжить?',
                                 QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
    if reply == QMessageBox.StandardButton.Ok:
        request_new_coords(COORDS_FILE, parent)
    else:
        QMessageBox.information(parent, 'Отмена', 'Обновление координат отменено.')


def request_login_button_coords(parent=None):
    reply = QMessageBox.question(parent, 'Обновление координат',
                                 'Наведите курсор на кнопку логина и нажмите \'Enter\'.\n\nХотите продолжить?',
                                 QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
    if reply == QMessageBox.StandardButton.Ok:
        request_new_coords(LOGIN_BUTTON_FILE, parent)
    else:
        QMessageBox.information(parent, 'Отмена', 'Обновление координат отменено.')


def request_login_field_coords(parent=None):
    reply = QMessageBox.question(parent, 'Обновление координат',
                                 'Наведите курсор на поле логина и нажмите \'Enter\'.\n\nХотите продолжить?',
                                 QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
    if reply == QMessageBox.StandardButton.Ok:
        request_new_coords(LOGIN_FIELD_FILE, parent)
    else:
        QMessageBox.information(parent, 'Отмена', 'Обновление координат отменено.')


def request_password_field_coords(parent=None):
    reply = QMessageBox.question(parent, 'Обновление координат',
                                 'Наведите курсор на поле пароля и нажмите \'Enter\'.\n\nХотите продолжить?',
                                 QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
    if reply == QMessageBox.StandardButton.Ok:
        request_new_coords(PASSWORD_FIELD_FILE, parent)
    else:
        QMessageBox.information(parent, 'Отмена', 'Обновление координат отменено.')


def request_secondlogin_field_coords(parent=None):
    reply = QMessageBox.question(parent, 'Обновление координат',
                                 'Наведите курсор на вторую кнопку логин и нажмите \'Enter\'.\n\nХотите продолжить?',
                                 QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
    if reply == QMessageBox.StandardButton.Ok:
        request_new_coords(SECOND_LOGIN_BUTTON_FILE, parent)
    else:
        QMessageBox.information(parent, 'Отмена', 'Обновление координат отменено.')


def request_firstwheel_coords(parent=None):
    reply = QMessageBox.question(parent, 'Обновление координат',
                                 'Наведите курсор на первое колесо и нажмите \'Enter\'.\n\nХотите продолжить?',
                                 QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
    if reply == QMessageBox.StandardButton.Ok:
        request_new_coords('coords/first_wheel.txt', parent)
    else:
        QMessageBox.information(parent, 'Отмена', 'Обновление координат отменено.')


def request_secondwheel_coords(parent=None):
    reply = QMessageBox.question(parent, 'Обновление координат',
                                 'Наведите курсор на второе колесо и нажмите \'Enter\'.\n\nХотите продолжить?',
                                 QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
    if reply == QMessageBox.StandardButton.Ok:
        request_new_coords('coords/second_wheel.txt', parent)
    else:
        QMessageBox.information(parent, 'Отмена', 'Обновление координат отменено.')


def request_thirdwheel_coords(parent=None):
    reply = QMessageBox.question(parent, 'Обновление координат',
                                 'Наведите курсор на третье колесо и нажмите \'Enter\'.\n\nХотите продолжить?',
                                 QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
    if reply == QMessageBox.StandardButton.Ok:
        request_new_coords('coords/third_wheel.txt', parent)
    else:
        QMessageBox.information(parent, 'Отмена', 'Обновление координат отменено.')


def request_presswheel_coords(parent=None):
    reply = QMessageBox.question(parent, 'Обновление координат',
                                 'Наведите курсор на колесо для прокручивания и нажмите \'Enter\'.\n\nХотите продолжить?',
                                 QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
    if reply == QMessageBox.StandardButton.Ok:
        request_new_coords('coords/third_wheel2.txt', parent)  # Assuming this is for press wheel, adjust if needed
    else:
        QMessageBox.information(parent, 'Отмена', 'Обновление координат отменено.')


def request_passwordmodal_coords(parent=None):
    reply = QMessageBox.question(parent, 'Обновление координат',
                                 'Наведите курсор на закрытие модального окна с сохранением пароля и нажмите \'Enter\'.\n\nХотите продолжить?',
                                 QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
    if reply == QMessageBox.StandardButton.Ok:
        request_new_coords('coords/target_point2.txt', parent)
    else:
        QMessageBox.information(parent, 'Отмена', 'Обновление координат отменено.')


def request_funt_coords(parent=None):
    reply = QMessageBox.question(parent, 'Обновление координат',
                                 'Наведите курсор на фунт и нажмите \'Enter\'.\n\nХотите продолжить?',
                                 QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
    if reply == QMessageBox.StandardButton.Ok:
        request_new_coords('coords/target_point.txt', parent)
    else:
        QMessageBox.information(parent, 'Отмена', 'Обновление координат отменено.')


def request_krest_coords(parent=None):
    reply = QMessageBox.question(parent, 'Обновление координат',
                                 'Наведите курсор на крестик в поисковой строке и нажмите \'Enter\'.\n\nХотите продолжить?',
                                 QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
    if reply == QMessageBox.StandardButton.Ok:
        request_new_coords('coords/target_point5.txt', parent)
    else:
        QMessageBox.information(parent, 'Отмена', 'Обновление координат отменено.')


def request_ladbucksopen_coords(parent=None):
    reply = QMessageBox.question(parent, 'Обновление координат',
                                 'Наведите курсор на место открытия ледбаксов и нажмите \'Enter\'.\n\nХотите продолжить?',
                                 QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
    if reply == QMessageBox.StandardButton.Ok:
        request_new_coords('coords/target_point3.txt', parent)
    else:
        QMessageBox.information(parent, 'Отмена', 'Обновление координат отменено.')


def request_ladbuckscopy_coords(parent=None):
    reply = QMessageBox.question(parent, 'Обновление координат',
                                 'Наведите курсор на количество ледбаксов и нажмите \'Enter\'.\n\nХотите продолжить?',
                                 QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
    if reply == QMessageBox.StandardButton.Ok:
        request_new_coords('coords/target_point4.txt', parent)
    else:
        QMessageBox.information(parent, 'Отмена', 'Обновление координат отменено.')


def handle_error_and_retry(login, password):
    print(f'Пытаемся выполнить вход для {login}...')
    retries = 0

    # повторяем попытки пока не исчерпаем лимит
    while retries < 5:
        print('Ждём 7 секунд перед проверкой на ошибку...')
        time.sleep(7)

        print('Проверяем наличие ошибки на экране...')
        error_coords = load_coords(ERROR_COORDS_FILE)

        if error_coords:
            try:
                screenshot = pyautogui.screenshot(
                    region=(error_coords[0], error_coords[1], 599, 592)
                )
                screenshot.save('screenshots/screenshot_error_area.png')

                error_location = pyautogui.locate(
                    ERROR_IMAGE, screenshot, confidence=0.8
                )

                if error_location:
                    print(
                        "Ошибка найдена на экране! Пожалуйста, введите новые координаты "
                        "для полей 'Пароль' и 'Второй Login'."
                    )
                    enter_credentials(login, password, error_coords=True)
                    retries += 1
                    time.sleep(3)
                    continue

                # если дошли сюда — ошибки не найдено
                print(f'Успешный вход для {login}. Ошибка не найдена на экране.')
                return True

            except pyautogui.ImageNotFoundException:
                print(
                    'Не удалось найти изображение ошибки, возможно, ошибка отсутствует. '
                    'Продолжаем выполнение.'
                )
                return True

        else:
            # координаты ошибки не загружены/не найдены
            return False

    # если исчерпали retries
    return False


def close_modal_window_and_click_wheel():
    time.sleep(10)
    krestik_coords = load_coords('coords/krestik.txt')
    print(f'Нажимаем на крестик по координатам: {krestik_coords}')
    pyautogui.click(krestik_coords)


def click_third_wheel():
    wheel_coords = load_coords('coords/third_wheel.txt')
    print(f'Нажимаем на третье колесо по координатам: {wheel_coords}')
    pyautogui.click(wheel_coords)
    time.sleep(2)


def click_second_wheel():
    wheel_coords = load_coords('coords/second_wheel.txt')
    print(f'Нажимаем на второе колесо по координатам: {wheel_coords}')
    pyautogui.click(wheel_coords)
    time.sleep(2)


def click_first_wheel():
    wheel_coords = load_coords('coords/first_wheel.txt')
    print(f'Нажимаем на первое колесо по координатам: {wheel_coords}')
    pyautogui.click(wheel_coords)
    time.sleep(2)


def second_click_to_wheel():
    """Выполняет повторный клик по третьему колесу с новыми координатами"""
    third_wheel2_coords = load_coords('coords/third_wheel2.txt')
    print(f'Нажимаем на колесо по новым координатам: {third_wheel2_coords}')
    pyautogui.click(third_wheel2_coords)
    time.sleep(10)


def press_f12():
    print('Нажимаем клавишу F12.')
    pyautogui.press('f12')
    time.sleep(5)


def press_ctrl_f():
    print('Нажимаем комбинацию Ctrl + F для поиска.')
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(3)


def print_funt():
    print('Вводим символ \'£\' в строку поиска.')
    pyperclip.copy('£')
    pyautogui.hotkey('ctrl', 'v')


def press_enter():
    print('Нажимаем Enter для выполнения поиска.')
    pyautogui.press('enter')
    time.sleep(1)


def denaid_password_window():
    print('Закрываем окно пароля.')
    target_point2_coords = load_coords('coords/target_point2.txt')
    print(f'Нажимаем на колесо по новым координатам: {target_point2_coords}')
    pyautogui.click(target_point2_coords)
    time.sleep(2)


def click_and_copy_funt():
    print('Загружаем координаты для новой точки.')
    target_point_coords = load_coords('coords/target_point.txt')
    print(f'Наводим курсор на точку по координатам: {target_point_coords}')
    pyautogui.moveTo(target_point_coords)
    print('Делаем двойной клик.')
    pyautogui.doubleClick()
    time.sleep(1)
    print('Копируем содержимое с помощью Ctrl + C.')
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(1)


def enter_data_to_excel(current_login):
    copied_data = pyperclip.paste()
    if copied_data:
        print(f'Сохраняем данные для: {current_login}')
        paste_data_to_excel(EXCEL_FILE, current_login)
    return None


def press_krestik():
    print('Загружаем координаты для новой точки.')
    target_point5_coords = load_coords('coords/target_point5.txt')
    print(f'Нажимаем на крестик по координатам: {target_point5_coords}')
    pyautogui.click(target_point5_coords)
    time.sleep(1)


def print_coinsbalance():
    print('Вводим символ \'coinsbalance\' в строку поиска.')
    pyperclip.copy('coinsbalance')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)


def open_ladbucks():
    print('Загружаем координаты для новой точки.')
    target_point3_coords = load_coords('coords/target_point3.txt')
    print(f'Нажимаем на крестик по координатам: {target_point3_coords}')
    pyautogui.click(target_point3_coords)


def copy_ladbucks():
    print('Загружаем координаты для новой точки.')
    target_point4_coords = load_coords('coords/target_point4.txt')
    print(f'Наводим курсор на точку по координатам: {target_point4_coords}')
    pyautogui.moveTo(target_point4_coords)
    time.sleep(1)
    print('Делаем двойной клик.')
    pyautogui.doubleClick()
    time.sleep(1)
    print('Копируем содержимое с помощью Ctrl + C.')
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(1)


def paste_ladbucks(current_login):
    copied_data = pyperclip.paste()
    if copied_data:
        print(f'Сохраняем данные для: {current_login}')
        paste_data_to_excel2(EXCEL_FILE, current_login)
    return None


def activate_multilogin_window():
    windows = [w for w in gw.getWindowsWithTitle('Multilogin') if not w.isActive]
    if not windows:
        print('Окно Multilogin не найдено.')
    return False


def wait_for_mimic_window():
    print('Ожидаем открытия окна браузера Mimic...')
    for i in range(30):
        windows = list(gw.getWindowsWithTitle(MIMIC_WINDOW_TITLE))
        if windows:
            print('Окно браузера Mimic обнаружено!')
            windows[0].activate()
            time.sleep(2)
            return True
        else:
            print('Окно браузера Mimic не найдено.')
            return False


def enter_url_in_browser():
    print(f'Вставляем URL: {URL_TO_OPEN}')
    pyperclip.copy(URL_TO_OPEN)
    pyautogui.hotkey('ctrl', 'l')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    print('Ссылка вставлена и подтверждена.')


def wait_for_page_load():
    print('Ожидаем загрузки страницы...')
    time.sleep(10)
    print('Страница загружена.')


def close_browser_window():
    windows = list(gw.getWindowsWithTitle(MIMIC_WINDOW_TITLE))
    if windows:
        print('Закрываем окно браузера.')
        windows[0].close()
        time.sleep(2)
    return None


def activate_new_profile():
    print('Запускаем разовый профиль в мультилогине...')
    click_button(COORDS_FILE)


def login_to_site():
    if wait_for_mimic_window():
        enter_url_in_browser()
        wait_for_page_load()
    return None


def process_user_account(user):
    enter_credentials(user['login'], user['password'])
    success = handle_error_and_retry(user['login'], user['password'])
    if success:
        print(f"Успешный вход для пользователя: {user['login']}")
    return None


def wait_for_browser_to_close():
    print('Ожидаем закрытия браузера...')
    time.sleep(120)


def main_step(user):
    close_browser_window()
    activate_multilogin_window()
    activate_new_profile()
    login_to_site()
    process_user_account(user)
    close_modal_window_and_click_wheel()


def wait_if_paused(self):
    if self._is_paused:
        self.msleep(100)
    if not self._is_running:
        pass  # postinserted
    return None


class WorkerThread(QThread):
    update_label = pyqtSignal(str)

    def __init__(self, selected_wheel=None):
        super().__init__()
        self._is_paused = False
        self._is_running = True
        self.selected_wheel = selected_wheel

    def run(self):
        if not self.selected_wheel:
            self.update_label.emit('Колесо не выбрано!')
        return None

    def execute_first_wheel_code(self):
        print('Код для первого колеса выполняется...')
        self.update_label.emit('Первое колесо нажато')
        click_first_wheel()

    def execute_second_wheel_code(self):
        print('Код для второго колеса выполняется...')
        self.update_label.emit('Второе колесо нажато')
        click_second_wheel()

    def execute_third_wheel_code(self):
        print('Код для третьего колеса выполняется...')
        self.update_label.emit('Третье колесо нажато')
        click_third_wheel()

    def wait_if_paused(self):
        if self._is_paused:
            self.msleep(100)

    def pause(self):
        if not self._is_paused:
            self._is_paused = True
            print('Поток приостановлен.')
        return None

    def resume(self):
        if self._is_paused:
            self._is_paused = False
            print('Поток возобновлен.')
        return None


class WheelSelectionWindow(QWidget):
    wheel_selected = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle('Выберите колесо')
        self.setFixedSize(400, 200)
        layout = QVBoxLayout()
        self.first_wheel_button = QPushButton('Первое колесо')
        self.first_wheel_button.clicked.connect(self.select_first_wheel)
        layout.addWidget(self.first_wheel_button)
        self.second_wheel_button = QPushButton('Второе колесо')
        self.second_wheel_button.clicked.connect(self.select_second_wheel)
        layout.addWidget(self.second_wheel_button)
        self.third_wheel_button = QPushButton('Третье колесо')
        self.third_wheel_button.clicked.connect(self.select_third_wheel)
        layout.addWidget(self.third_wheel_button)
        self.setLayout(layout)

    def select_first_wheel(self):
        self.wheel_selected.emit('Первое колесо')
        self.close()

    def select_second_wheel(self):
        self.wheel_selected.emit('Второе колесо')
        self.close()

    def select_third_wheel(self):
        self.wheel_selected.emit('Третье колесо')
        self.close()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Ladbrokes')
        self.setFixedSize(400, 200)
        self.selected_wheel = None
        self.excel_file_path = 'excel/users.xlsx'
        self.thread = WorkerThread()
        self.thread.update_label.connect(self.update_status)
        self.init_ui()
        self.setup_global_shortcuts()

    def init_ui(self):
        """Инициализация пользовательского интерфейса"""  # inserted
        layout = QVBoxLayout()
        self.label = QLabel('Нажмите \'Старт\' для начала работы')
        layout.addWidget(self.label)
        self.start_button = QPushButton('Старт')
        self.start_button.clicked.connect(self.start_process)
        layout.addWidget(self.start_button)
        self.pause_button = QPushButton('Пауза')
        self.pause_button.clicked.connect(self.pause_process)
        layout.addWidget(self.pause_button)
        self.clear_button = QPushButton('Очистить данные')
        self.clear_button.clicked.connect(self.clear_excel_data)
        layout.addWidget(self.clear_button)
        self.update_coords_button = QPushButton('Обновить координаты кнопок')
        self.update_coords_button.clicked.connect(self.update_button_coordinates)
        layout.addWidget(self.update_coords_button)
        self.select_wheel_button = QPushButton('Выбрать колесо')
        self.select_wheel_button.clicked.connect(self.show_wheel_selection)
        layout.addWidget(self.select_wheel_button)
        self.status_label = QLabel('Можно приступать к работе')
        layout.addWidget(self.status_label)
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def show_wheel_selection(self):
        """Показать окно выбора колеса"""  # inserted
        self.wheel_selection_window = WheelSelectionWindow()
        self.wheel_selection_window.wheel_selected.connect(self.on_wheel_selected)
        self.wheel_selection_window.show()

    def on_wheel_selected(self, wheel):
        """Обработчик выбора колеса"""  # inserted
        self.selected_wheel = wheel
        self.label.setText(f'Вы выбрали: {wheel}')

    def execute_wheel_code(self):
        """Выполнение кода для выбранного колеса"""  # inserted
        if not self.selected_wheel:
            QMessageBox.warning(self, 'Ошибка', 'Колесо не выбрано!')
        return None

    def update_label(self, message):
        """Обновление текста метки"""  # inserted
        self.label.setText(message)

    @pyqtSlot(str)
    def update_status(self, message):
        """Обновление статуса (слот)"""  # inserted
        self.label.setText(message)

    def setup_global_shortcuts(self):
        """Настройка глобальных горячих клавиш"""  # inserted
        keyboard.add_hotkey('f9', self.start_process)
        keyboard.add_hotkey('f8', self.pause_process)

    def start_process(self):
        """Запуск процесса"""  # inserted
        if not self.thread.isRunning():
            self.thread = WorkerThread(self.selected_wheel)
            self.thread.update_label.connect(self.update_status)
            self.thread.start()
        return None

    def pause_process(self):
        """Приостановка процесса"""  # inserted
        if self.thread.isRunning():
            self.thread.pause()
        return None

    def clear_excel_data(self):
        """Очистка данных в Excel"""  # inserted
        if not self.excel_file_path:
            QMessageBox.critical(self, 'Ошибка', 'Файл Excel не указан.')
        return None

    def update_button_coordinates(self):
        """Обновление координат кнопок"""  # inserted
        try:
            request_multilogin_button_coords(self)
            request_login_button_coords(self)
            request_login_field_coords(self)
            request_password_field_coords(self)
            request_secondlogin_field_coords(self)
            request_krestik_coords(self)
            request_firstwheel_coords(self)
            request_secondwheel_coords(self)
            request_thirdwheel_coords(self)
            request_presswheel_coords(self)
            request_passwordmodal_coords(self)
            request_funt_coords(self)
            request_krest_coords(self)
            request_ladbucksopen_coords(self)
            request_ladbuckscopy_coords(self)
            QMessageBox.information(self, 'Успех', 'Координаты обновлены и сохранены.')
        except Exception as e:
            QMessageBox.critical(self, 'Ошибка', f'Не удалось обновить координаты: {e}')

    def closeEvent(self, event):
        """Обработчик закрытия окна"""
        keyboard.unhook_all_hotkeys()
        self.thread.terminate() if self.thread.isRunning() else None
        QApplication.quit()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())