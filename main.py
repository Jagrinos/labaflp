import tkinter as tk
from tkinter import messagebox
from docx import Document
from docx2pdf import convert
import fitz
import os
from PIL import Image, ImageTk
import io

# Функция для сохранения данных в документ
def save_to_docx():
    name = entry_name.get()
    surname = entry_surname.get()

    if not name or not surname:
        messagebox.showwarning("Ошибка", "Пожалуйста, введите имя и фамилию")
        return

    # Показываем сообщение о загрузке
    loading_label.config(text="Идёт загрузка...")
    root.update()  # Обновляем интерфейс, чтобы сообщение отобразилось

    # Запускаем сохранение и обновление предпросмотра в отдельном поток
    save_and_update_preview(name, surname)

# Функция для сохранения данных и обновления предпросмотра
def save_and_update_preview(name, surname):
    try:
        # Создание нового документа или открытие существующего

        doc = Document()

        # Добавление данных в документ
        doc.add_paragraph(f"Имя: {name}")
        doc.add_paragraph(f"Фамилия: {surname}")
        doc.add_paragraph("")  # Пустая строка для разделения записей

        # Сохранение документа
        doc.save('output.docx')

        # Конвертация DOCX в PDF
        convert('output.docx', 'output.pdf')

        # Открытие PDF и отображение первой страницы
        pdf_document = fitz.open('output.pdf')
        first_page = pdf_document.load_page(0)  # Загружаем первую страницу
        pix = first_page.get_pixmap()  # Преобразуем страницу в изображение

        # Преобразуем изображение в формат, подходящий для Tkinter
        img = Image.open(io.BytesIO(pix.tobytes()))
        img = ImageTk.PhotoImage(img)

        # Обновляем изображение в интерфейсе
        pdf_preview_label.config(image=img)
        pdf_preview_label.image = img  # Сохраняем ссылку, чтобы изображение не удалялось сборщиком мусора

        # Скрываем сообщение о загрузке
        loading_label.config(text="")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")
    finally:
        # Скрываем сообщение о загрузке в любом случае
        loading_label.config(text="")

# Создание основного окна
root = tk.Tk()
root.title("Ввод данных в DOCX с предпросмотром PDF")

# Левая часть окна: поля ввода и кнопки
left_frame = tk.Frame(root)
left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="n")

# Поля для ввода имени и фамилии
tk.Label(left_frame, text="Имя:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
entry_name = tk.Entry(left_frame)
entry_name.grid(row=0, column=1, padx=10, pady=10)

tk.Label(left_frame, text="Фамилия:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
entry_surname = tk.Entry(left_frame)
entry_surname.grid(row=1, column=1, padx=10, pady=10)

# Кнопка для сохранения данных
save_button = tk.Button(left_frame, text="Сохранить", command=save_to_docx)
save_button.grid(row=2, column=0, columnspan=2, pady=10)

# Метка для отображения сообщения о загрузке
loading_label = tk.Label(left_frame, text="", fg="blue")
loading_label.grid(row=3, column=0, columnspan=2, pady=10)

# Правая часть окна: предпросмотр PDF
right_frame = tk.Frame(root)
right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="n")

# Метка для отображения предпросмотра PDF
pdf_preview_label = tk.Label(right_frame)
pdf_preview_label.grid(row=0, column=0, padx=10, pady=10)

# Первоначальное обновление предпросмотра (если файл уже существует)
if os.path.exists('output.docx'):
    save_and_update_preview("name", "surname")

# Запуск основного цикла
root.mainloop()