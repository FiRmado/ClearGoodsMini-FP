# Импорты
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox
import win32com.client
import re


# Создание интерфейса с использованием ttkbootstrap
root = ttk.Window(themename="superhero")
root.title("Очищення бази")
root.geometry("600x450")
root.resizable(False, False)

# Меню
menu_bar = ttk.Menu(root)
root.config(menu=menu_bar)

info_menu = ttk.Menu(menu_bar, tearoff=0)
# Подменю для информации о программе
about_menu = ttk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Про програму", menu=about_menu)
about_menu.add_command(label="Автор: @FiRmado")
about_menu.add_command(label="Дата релізу: 10.12.2024")

# Поле ввода IP-адреса
ip_label = ttk.Label(root, text="Введіть IP-адресу каси:")
ip_label.pack(pady=5)

ip_entry = ttk.Entry(root, width=30)
ip_entry.pack(pady=5)

# Кнопка подключения
def connect_to_device():
    ip_address = ip_entry.get()
    if not ip_address:
        log_message("Помилка: Введіть IP-адресу!")
        return

    try:
        ecr = win32com.client.Dispatch("ecrmini.t400")
        command = f"connect_tcp;{ip_address};21700;1111;"
        if ecr.t400me(command):
            log_message(f"Підключення до {ip_address} встановлено успішно!")
            messagebox.showinfo(f"{ip_address}", f"Підключення до {ip_address} встановлено успішно!")
        else:
            log_message("Помилка підключення до каси.")
    except Exception as e:
        log_message(f"Помилка: {e}")

connect_button = ttk.Button(root, text="Підключитися", command=connect_to_device, style="success", width=30)
connect_button.pack(pady=10)
##################################################################################
# ОПРОС СЕТИ НА НАЛИЧИЕ КАСС
def discover_devices():
    try:
        # Подключение к OLE-серверу
        ecr = win32com.client.Dispatch("ecrmini.t400")
        log_message("Підключення до .dll бібліотеки встановлено!")

        # Выполнение команды rro_request
        command = "rro_request;"
        if ecr.t400me(command):
            response = ecr.get_last_result
            if response:
                devices = response.split(";")
                if devices[0] == "0":  # Проверяем, что нет ошибок
                    log_message("Знайдені пристрої:")
                    for device in devices[1:]:
                        device = device.strip()  # Убираем лишние пробелы
                        if not device:
                            continue  # Игнорируем пустые строки

                        # Используем регулярное выражение для извлечения данных
                        match = re.search(
                            r'"(?P<type>MINI-[^"]+)"\s+"[^"]*"\s+"[^"]*"\s+"(?P<ip>[\d\.]+)(?::\d+)?"\s+"(?P<status>READY|BUSY|NONE)"',
                            device,
                        )
                        if match:
                            device_type = match.group("type")
                            ip_address = match.group("ip")
                            status = match.group("status")
                            log_message(f"Тип: {device_type}, IP: {ip_address}, Статус: {status}")
                        else:
                            log_message(f"Неможливо обробити пристрій: {device}")
                else:
                    log_message(f"Помилка в команді rro_request: {response}")
            else:
                log_message("Помилка: Відповідь відсутня.")
        else:
            log_message("Помилка виконання команди rro_request.")
    except Exception as e:
        log_message(f"Помилка: {e}")

# Кнопка для опроса сети
discover_button = ttk.Button(root, text="Опрос мережі", command=discover_devices, style="primary", width=30)
discover_button.pack(pady=10)

##################################################################################
# ФУНКЦИЯ ОЧИЩЕНИЯ БАЗЫ ТОВАРОВ
def clear_database():
    try:
        ip_address = ip_entry.get()
        if not ip_address:
            log_message("Помилка: Введіть IP-адресу!")
            messagebox.showwarning("Увага", "Помилка: Введіть IP-адресу!")
            return

        # Подключение к OLE-серверу
        ecr = win32com.client.Dispatch("ecrmini.t400")

        # Подключение к кассе
        connect_command = f"connect_tcp;{ip_address};21700;1111;"
        if not ecr.t400me(connect_command):
            log_message("Помилка підключення до каси.")
            messagebox.showerror("Помилка", "Не вдалося підключитися до каси.")
            return

        log_message(f"Підключення до {ip_address} встановлено успішно!")

        # Выполнение команд
        commands = [
            #"write_table;18;7;1",
            "paper_feed;10;",
            #"del_plu;2;1"
        ]
        for command in commands:
            if not ecr.t400me(command):
                log_message(f"Помилка виконання команди: {command}")
            else:
                log_message(f"Команда виконана успішно: {command}")

        # Уведомление о завершении
        messagebox.showinfo(
            "Очищення бази",
            "База товарів очищенна, необхідно закрити цю програму і можна розпочинати роботу на своїй програмі!"
        )

    except Exception as e:
        log_message(f"Помилка: {e}")

    finally:
        # Закрытие порта
        close_command = "close_port;"
        if not ecr.t400me(close_command):
            log_message(f"Помилка виконання команди: {close_command}")
        else:
            log_message("Порт закрито успішно.")


###############################################################################################################
clear_button = ttk.Button(root, text="Очистити базу товарів", command=clear_database, style="danger", width=30)
clear_button.pack(pady=10)

# Поле вывода сообщений
def log_message(message):
    output_text.insert("end", message + "\n")
    output_text.see("end")

output_text = ttk.ScrolledText(root, width=70, height=15)
output_text.pack(pady=10)

# Запуск приложения
root.mainloop()