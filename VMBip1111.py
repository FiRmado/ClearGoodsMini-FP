# Импорты
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox
import win32com.client
import re

# Функция для логирования сообщений
def log_message(message):
    output_text.insert("end", message + "\n")
    output_text.see("end")

# Функция для опроса сети на наличие касс
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
                    available_ips = []
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
                            if status == "READY":
                                available_ips.append(ip_address)
                    # Заполняем выпадающий список IP-адресами
                    ip_combo["values"] = available_ips
                    if available_ips:
                        ip_combo.set(available_ips[0])  # Устанавливаем первый IP по умолчанию
                else:
                    log_message(f"Помилка в команді rro_request: {response}")
            else:
                log_message("Помилка: Відповідь відсутня.")
        else:
            log_message("Помилка виконання команди rro_request.")
    except Exception as e:
        log_message(f"Помилка: {e}")

# Функция подключения к кассе
def connect_to_device():
    ip_address = ip_combo.get()
    password = password_entry.get()
    if not ip_address:
        log_message("Помилка: Виберіть IP-адресу!")
        return

    try:
        ecr = win32com.client.Dispatch("ecrmini.t400")
        # Формируем команду с учётом пароля
        command = f"connect_tcp;{ip_address};21700;{password};" if password else f"connect_tcp;{ip_address};21700;;"
        if ecr.t400me(command):
            log_message(f"Підключення до {ip_address} встановлено успішно!")
            messagebox.showinfo(f"{ip_address}", f"Підключення до {ip_address} встановлено успішно!")
        else:
            log_message("Помилка підключення до каси.")
    except Exception as e:
        log_message(f"Помилка: {e}")

# Функция очистки базы товаров
def clear_database():
    ip_address = ip_combo.get()
    password = password_entry.get()
    if not ip_address:
        log_message("Помилка: Виберіть IP-адресу!")
        messagebox.showwarning("Увага", "Помилка: Виберіть IP-адресу!")
        return

    try:
        ecr = win32com.client.Dispatch("ecrmini.t400")

        # Подключение к кассе
        connect_command = f"connect_tcp;{ip_address};21700;{password};" if password else f"connect_tcp;{ip_address};21700;;"
        if not ecr.t400me(connect_command):
            log_message("Помилка підключення до каси.")
            messagebox.showerror("Помилка", "Не вдалося підключитися до каси.")
            return

        log_message(f"Підключення до {ip_address} встановлено успішно!")

        # Выполнение команд
        commands = [
            "write_table;18;7;1;",
            "del_plu;2;1;",
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

# Создание интерфейса с использованием ttkbootstrap
root = ttk.Window(themename="superhero")
root.title("Очищення бази")
root.geometry("600x450")
root.resizable(False, False)
# Меню программы
menu_bar = ttk.Menu(root)
root.config(menu=menu_bar)
info_menu = ttk.Menu(menu_bar, tearoff=0)
# Подменю для информации о программе
about_menu = ttk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Про програму", menu=about_menu)
about_menu.add_command(label="Автор: @FiRmado")
about_menu.add_command(label="Дата релізу: 10.12.2024")
# Поле выбора IP-адреса
ip_label = ttk.Label(root, text="Виберіть IP-адресу каси:")
ip_label.pack(pady=5)

ip_frame = ttk.Frame(root)
ip_frame.pack(pady=5)

ip_combo = ttk.Combobox(ip_frame, state="readonly", width=20)
ip_combo.grid(row=0, column=0, padx=5)

password_label = ttk.Label(ip_frame, text="Пароль:")
password_label.grid(row=0, column=1, padx=5)

password_entry = ttk.Entry(ip_frame, width=10)
password_entry.grid(row=0, column=2, padx=5)

# Кнопки
discover_button = ttk.Button(root, text="Опрос мережі", command=discover_devices, style="primary", width=30)
discover_button.pack(pady=10)

connect_button = ttk.Button(root, text="Підключитися", command=connect_to_device, style="success", width=30)
connect_button.pack(pady=10)

clear_button = ttk.Button(root, text="Очистити базу товарів", command=clear_database, style="danger", width=30)
clear_button.pack(pady=10)

# Поле вывода сообщений
output_text = ttk.ScrolledText(root, width=70, height=15)
output_text.pack(pady=10)

# Запуск приложения
root.mainloop()