import os
import sys
import tarfile
import py7zr
import rarfile
import pyzipper
from tkinter import ttk, filedialog, messagebox
import tkinter as tk
import tkinter.font as tkfont
import webbrowser
import random
import threading
from tkinterdnd2 import *

def resource_path(relative_path):
    """ путь к ресурсу и работа """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

class Archiver:
    def __init__(self):
        self.FILE_TYPES = {
            '.jpg': 'Фото', '.png': 'Фото',
            '.mp4': 'Видео',
            '.doc': 'Документ', '.docx': 'Документ', '.pdf': 'ПДФ Документ',
            '.mp3': 'Аудио'
        }

    def get_file_type(self, filename):
        ext = os.path.splitext(filename)[1].lower()
        return self.FILE_TYPES.get(ext, 'Файл')

    def create_archive(self, files, name, ext, password=None):
        if not name.endswith(ext):
            name += ext
        try:
            if ext == '.zip':
                with pyzipper.AESZipFile(name, 'w',
                                         compression=pyzipper.ZIP_DEFLATED,
                                         encryption=pyzipper.WZ_AES if password else None) as z:
                    if password:
                        z.pwd = password.encode('utf-8')
                    for f in files:
                        z.write(f, os.path.basename(f))
            elif ext == '.7z':
                with py7zr.SevenZipFile(name, 'w', password=password) as z:
                    for f in files:
                        z.write(f, os.path.basename(f))
            elif ext == '.tar':
                with tarfile.open(name, 'w') as t:
                    for f in files:
                        t.add(f, os.path.basename(f))
            size = os.path.getsize(name) / (1024 * 1024)
            password_note = "\nПароль установлен" if password else ""
            return f"Создан:\nПуть: {name}\nРазмер: {size:.2f} МБ{password_note}"
        except Exception as e:
            return f"Ошибка при создании: {str(e)}"

    def extract_archive(self, archive, folder, password=None):
        try:
            if archive.endswith('.zip'):
                with pyzipper.AESZipFile(archive, 'r') as z:
                    if password:
                        z.pwd = password.encode('utf-8')
                    z.extractall(folder)
            elif archive.endswith('.7z'):
                with py7zr.SevenZipFile(archive, 'r', password=password) as z:
                    z.extractall(folder)
            elif archive.endswith('.rar'):
                with rarfile.RarFile(archive, 'r') as r:
                    if password:
                        r.setpassword(password)
                    r.extractall(folder)
            elif archive.endswith('.tar'):
                with tarfile.open(archive, 'r') as t:
                    t.extractall(folder)
            return f"Извлечено в {folder}"
        except pyzipper.BadZipFile:
            return "Ошибка: Неверный пароль или поврежденный архив"
        except Exception as e:
            return f"Ошибка при извлечении: {str(e)}"

    def get_archive_info(self, archive, password=None):
        try:
            info = f"Содержимое {archive}:\n{'-' * 20}\n"
            if archive.endswith('.zip'):
                with pyzipper.AESZipFile(archive, 'r') as z:
                    if password:
                        z.pwd = password.encode('utf-8')
                    for f in z.infolist():
                        size = f.file_size / (1024 * 1024)
                        typ = self.get_file_type(f.filename)
                        info += f"{f.filename} - {size:.2f} МБ ({typ})\n"
            elif archive.endswith('.7z'):
                with py7zr.SevenZipFile(archive, 'r', password=password) as z:
                    for f in z.getnames():
                        typ = self.get_file_type(f)
                        info += f"{f} - (размер неизвестен) ({typ})\n"
            elif archive.endswith('.rar'):
                with rarfile.RarFile(archive, 'r') as r:
                    if password:
                        r.setpassword(password)
                    for f in r.infolist():
                        size = f.file_size / (1024 * 1024)
                        typ = self.get_file_type(f.filename)
                        info += f"{f.filename} - {size:.2f} МБ ({typ})\n"
            elif archive.endswith('.tar'):
                with tarfile.open(archive, 'r') as t:
                    for f in t.getmembers():
                        size = f.size / (1024 * 1024)
                        typ = self.get_file_type(f.name)
                        info += f"{f.name} - {size:.2f} МБ ({typ})\n"
            else:
                return f"Ошибка: Неизвестный формат архива {archive}"
            return info
        except pyzipper.BadZipFile:
            return "Ошибка: Неверный пароль или поврежденный архив"
        except Exception as e:
            return f"Ошибка при получении информации: {str(e)}"

    def preview_existing_archive(self, archive, password=None):
        try:
            preview = f"Предпросмотр {archive}:\n{'-' * 20}\n"
            if archive.endswith('.zip'):
                with pyzipper.AESZipFile(archive, 'r') as z:
                    if password:
                        z.pwd = password.encode('utf-8')
                    for f in z.infolist():
                        size = f.file_size / (1024 * 1024)
                        typ = self.get_file_type(f.filename)
                        preview += f"{f.filename} - {size:.2f} МБ ({typ})\n"
            elif archive.endswith('.7z'):
                with py7zr.SevenZipFile(archive, 'r', password=password) as z:
                    for f in z.getnames():
                        typ = self.get_file_type(f)
                        preview += f"{f} - (размер неизвестен) ({typ})\n"
            elif archive.endswith('.rar'):
                with rarfile.RarFile(archive, 'r') as r:
                    if password:
                        r.setpassword(password)
                    for f in r.infolist():
                        size = f.file_size / (1024 * 1024)
                        typ = self.get_file_type(f.filename)
                        preview += f"{f.filename} - {size:.2f} МБ ({typ})\n"
            elif archive.endswith('.tar'):
                with tarfile.open(archive, 'r') as t:
                    for f in t.getmembers():
                        size = f.size / (1024 * 1024)
                        typ = self.get_file_type(f.name)
                        preview += f"{f.name} - {size:.2f} МБ ({typ})\n"
            return preview
        except pyzipper.BadZipFile:
            return "Ошибка: Неверный пароль или поврежденный архив"
        except Exception as e:
            return f"Ошибка при предпросмотре: {str(e)}"

    def preview_archive(self, files):
        preview = "Предпросмотр:\n" + "-" * 20 + "\n"
        total_size = 0
        for f in files:
            size = os.path.getsize(f) / (1024 * 1024)
            typ = self.get_file_type(f)
            preview += f"{os.path.basename(f)} - {size:.2f} МБ ({typ})\n"
            total_size += size
        preview += f"Всего: {total_size:.2f} МБ"
        return preview

class ArchiverGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("whoisarchiver?")
        self.root.geometry("350x350")
        self.root.configure(bg="#f0f0f0")
        self.archiver = Archiver()
        self.hover_count = 0

        # Иницилизация TkinterDnD2
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.handle_drop)

        available_fonts = tkfont.families()
        self.default_font = "Montserrat" if "Montserrat" in available_fonts else "Arial"

        self.style = ttk.Style()
        self.style.configure('TLabel', background='#f0f0f0', font=(self.default_font, 10))
        self.style.configure('TButton', font=(self.default_font, 9, 'bold'), padding=3, width=8)
        self.style.configure('TCombobox', font=(self.default_font, 10))

        self.tabs = ttk.Notebook(root)
        self.tabs.pack(fill="both", expand=True, pady=5)

        self.main_tab = ttk.Frame(self.tabs)
        self.tabs.add(self.main_tab, text="Архиватор")

        self.about_tab = ttk.Frame(self.tabs)
        self.tabs.add(self.about_tab, text="О создателе")

        self.secret_tab = ttk.Frame(self.tabs)
        self.tabs.add(self.secret_tab, text="Секрет?")

        self.title = ttk.Label(self.main_tab, text="whoisarchiver?", font=(self.default_font, 14, 'bold'))
        self.title.pack(pady=5)

        # Добавляем drag and drop
        self.drop_frame = ttk.Frame(self.main_tab, relief="groove", borderwidth=2)
        self.drop_frame.pack(pady=5, padx=5, fill="x")
        self.drop_label = ttk.Label(self.drop_frame, text="Перетащите файлы или архив сюда", font=(self.default_font, 10))
        self.drop_label.pack(pady=10, padx=10)
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.handle_drop)
        self.drop_frame.dnd_bind('<<DragEnter>>', self.highlight_drop_area)
        self.drop_frame.dnd_bind('<<DragLeave>>', self.unhighlight_drop_area)

        self.ext_var = tk.StringVar(value='.zip')
        self.ext_combo = ttk.Combobox(self.main_tab, textvariable=self.ext_var,
                                      values=['.zip', '.7z', '.tar'], state="readonly", width=10)
        self.ext_combo.pack(pady=5)

        self.info = tk.Text(self.main_tab, height=8, width=40, font=(self.default_font, 10))
        self.info.pack(pady=5)

        self.btn_frame = ttk.Frame(self.main_tab)
        self.btn_frame.pack(pady=5)
        ttk.Button(self.btn_frame, text="Создать", command=self.create).pack(side="left", padx=5)
        ttk.Button(self.btn_frame, text="Извлечь", command=self.extract).pack(side="left", padx=5)
        ttk.Button(self.btn_frame, text="Инфо", command=self.show_archive_info).pack(side="left", padx=5)

        self.create_about_tab()

        self.secret_text = tk.Text(self.secret_tab, height=12, width=40, font=(self.default_font, 10))
        self.secret_text.pack(pady=5, padx=5)
        self.secret_button = ttk.Button(self.secret_tab, text="Не нажимать", command=self.secret_click)
        self.secret_button.pack(pady=10)
        self.secret_button.bind("<Enter>", self.move_button)

    def highlight_drop_area(self, event):
        self.drop_frame.configure(relief="solid", style="Highlight.TFrame")
        self.style.configure("Highlight.TFrame", bordercolor="blue", borderwidth=2)

    def unhighlight_drop_area(self, event):
        self.drop_frame.configure(relief="groove", style="TFrame")
        self.style.configure("TFrame", bordercolor="black", borderwidth=2)

    def handle_drop(self, event):
        files = self.root.splitlist(event.data)
        archive_extensions = ['.zip', '.7z', '.rar', '.tar']
        
        # Проверяем архивы ли
        is_archive = any(os.path.splitext(f)[1].lower() in archive_extensions for f in files)
        
        if is_archive:
            if len(files) > 1:
                self.show_info("Пожалуйста, перетащите только один архив")
                return
            archive = files[0]
            if not os.path.splitext(archive)[1].lower() in archive_extensions:
                self.show_info("Неподдерживаемый формат архива!")
                return
            self.process_dropped_archive(archive)
        else:
            self.process_dropped_files(files)

    def process_dropped_files(self, files):
        if not files:
            self.show_info("Файлы не выбраны")
            return
        preview = self.archiver.preview_archive(files)
        self.show_info(preview)
        name = filedialog.asksaveasfilename(title="Сохраните архив", defaultextension=self.ext_var.get())
        if not name:
            self.show_info("Укажите имя архива")
            return

        password = None
        if self.ext_var.get() != '.tar':
            password = self.ask_password("Установка пароля")

        threading.Thread(target=self.create_async, args=(files, name, self.ext_var.get(), password),
                         daemon=True).start()

    def process_dropped_archive(self, archive):
        password = None
        if not archive.endswith('.tar'):
            password = self.ask_password("Введите пароль")

        preview = self.archiver.preview_existing_archive(archive, password)
        self.show_info(preview)

        folder = filedialog.askdirectory(title="Выберите папку")
        if not folder:
            self.show_info("Выберите папку!")
            return

        result = self.archiver.extract_archive(archive, folder, password)
        self.show_info(result)

    def create_about_tab(self):
        about_frame = ttk.Frame(self.about_tab)
        about_frame.pack(pady=10, padx=10, fill="both", expand=True)

        title_label = ttk.Label(about_frame, text="О создателе",
                                font=(self.default_font, 14, 'bold'))
        title_label.pack(pady=(0, 10))

        about_text = """Эта программа создана для регионального конкурса проектных работ 
по программированию от Центра цифрового образования детей «IT-куб» 
в Калининграде.

Меня зовут Максимка ( в интернете я более известнен как Fazenda), я начинающий разработчик, изучающий Python. 
В приоритете у меня другие языки программирования, но я стараюсь 
совершенствоваться и в этом направлении. 

Кроме программирования, я увлекаюсь музыкой. 
Надеюсь, вам понравится мой архиватор!

Связаться со мной: whoisfazenda@gmail.com"""

        about_label = ttk.Label(about_frame, text=about_text,
                                font=(self.default_font, 10),
                                wraplength=360,
                                justify="left")
        about_label.pack()

    def ask_password(self, title="Установка пароля"):
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Введите пароль (если требуется):").pack(pady=10)
        password_var = tk.StringVar()
        password_entry = ttk.Entry(dialog, textvariable=password_var, show="*")
        password_entry.pack(pady=5)

        result = [None]

        def on_ok():
            result[0] = password_var.get() if password_var.get() else None
            dialog.destroy()

        def on_cancel():
            result[0] = None
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="ОК", command=on_ok).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Отмена", command=on_cancel).pack(side="left", padx=5)

        self.root.wait_window(dialog)
        return result[0]

    def show_info(self, text):
        self.info.delete(1.0, tk.END)
        self.info.insert(tk.END, text)
        self.root.update()

    def create_async(self, files, name, ext, password):
        self.show_info("Создание архива, подождите...")
        result = self.archiver.create_archive(files, name, ext, password)
        self.show_info(result)

    def create(self):
        files = filedialog.askopenfilenames(title="Выберите файлы")
        if not files:
            self.show_info("Выберите файлы!")
            return
        preview = self.archiver.preview_archive(files)
        self.show_info(preview)
        name = filedialog.asksaveasfilename(title="Сохраните архив", defaultextension=self.ext_var.get())
        if not name:
            self.show_info("Укажите имя архива!")
            return

        password = None
        if self.ext_var.get() != '.tar':
            password = self.ask_password("Установка пароля")

        threading.Thread(target=self.create_async, args=(files, name, self.ext_var.get(), password),
                         daemon=True).start()

    def extract(self):
        archive = filedialog.askopenfilename(title="Выберите архив",
                                             filetypes=[("Архивы", "*.zip *.7z *.rar *.tar")])
        if not archive:
            self.show_info("Выберите архив!")
            return

        password = None
        if not archive.endswith('.tar'):
            password = self.ask_password("Введите пароль")

        preview = self.archiver.preview_existing_archive(archive, password)
        self.show_info(preview)

        folder = filedialog.askdirectory(title="Выберите папку")
        if not folder:
            self.show_info("Выберите папку!")
            return

        result = self.archiver.extract_archive(archive, folder, password)
        self.show_info(result)

    def show_archive_info(self):
        archive = filedialog.askopenfilename(title="Выберите архив",
                                             filetypes=[("Архивы", "*.zip *.7z *.rar *.tar")])
        if not archive:
            self.show_info("Архив не выбран!")
            return

        password = None
        if not archive.endswith('.tar'):
            password = self.ask_password("Введите пароль")

        self.show_info("Получение информации об архиве...")
        result = self.archiver.get_archive_info(archive, password)
        self.show_info(result)

    def move_button(self, event):
        if self.hover_count < 3:
            self.hover_count += 1
            x = random.randint(0, 350)
            y = random.randint(0, 250)
            self.secret_button.place(x=x, y=y)
        else:
            self.secret_button.place_forget()
            center_x = (400 - self.secret_button.winfo_reqwidth()) // 2
            center_y = (300 - self.secret_button.winfo_reqheight()) // 2
            self.secret_button.place(x=center_x, y=center_y)
            self.secret_button.unbind("<Enter>")

    def secret_click(self):
        if self.hover_count >= 3:
            self.secret_button.destroy()
            self.secret_label = tk.Label(self.secret_tab, text="Просто ради шутки. Я смешной",
                                         font=(self.default_font, 12, 'bold'), bg="#f0f0f0")
            center_x = (400 - self.secret_label.winfo_reqwidth()) // 2
            center_y = (300 - self.secret_label.winfo_reqheight()) // 2
            self.secret_label.place(x=center_x, y=center_y)
            webbrowser.open("https://www.youtube.com/watch?v=dQw4w9WgXcQ")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ArchiverGUI(root)

    try:
        if sys.platform.startswith('win'):
            root.iconbitmap(resource_path("icon.ico"))
        else:
            icon = tk.PhotoImage(file=resource_path("icon.png"))
            root.iconphoto(True, icon)
    except tk.TclError:
        print("Не удалось загрузить иконку. Убедитесь, что файлы 'icon.ico' или 'icon.png' существуют.")

    root.mainloop()