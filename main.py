import tkinter as tk
from tkinter import messagebox

from Demos.mmapfile_demo import fsize
from docx import Document
from docx2pdf import convert
import fitz
import os
from PIL import Image, ImageTk
import io

from markup import markupdocx


# Функция для сохранения данных в документ
def save_to_docx():
    rec = recipient.get("1.0", tk.END).strip()
    tit = title.get()
    txt = text.get("1.0", tk.END).strip()
    sp = sender_profession.get()
    s = sender.get()

    if not rec or not tit or not txt or not sp or not s:
        messagebox.showwarning("Ошибка", "Пожалуйста, введите все поля")
        return

    # Показываем сообщение о загрузке
    loading_label.config(text="Идёт загрузка...")
    root.update()  # Обновляем интерфейс, чтобы сообщение отобразилось

    # Запускаем сохранение и обновление предпросмотра в отдельном поток
    save_and_update_preview(rec, tit, txt, sp, s)

# Функция для сохранения данных и обновления предпросмотра
def save_and_update_preview(rec, tit, txt, sp, s):
    try:
        markupdocx(rec, tit, txt, sp, s)

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

# Поля для ввода
tk.Label(left_frame, text="Получатель:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
recipient = tk.Text(left_frame)
recipient.grid(row=0, column=1, padx=10, pady=10)

tk.Label(left_frame, text="Заголовок:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
title = tk.Entry(left_frame)
title.grid(row=1, column=1, padx=10, pady=10)

tk.Label(left_frame, text="Текст:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
text = tk.Text(left_frame)
text.grid(row=2, column=1, padx=10, pady=10)

tk.Label(left_frame, text="Ваша должность:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
sender_profession = tk.Entry(left_frame)
sender_profession.grid(row=3, column=1, padx=10, pady=10)

tk.Label(left_frame, text="Ваше ФИО:").grid(row=4, column=0, padx=10, pady=10, sticky="w")
sender = tk.Entry(left_frame)
sender.grid(row=4, column=1, padx=10, pady=10)


# Кнопка для сохранения данных
save_button = tk.Button(left_frame, text="Сохранить", command=save_to_docx)
save_button.grid(row=5, column=0, columnspan=2, pady=10)



# Правая часть окна: предпросмотр PDF
right_frame = tk.Frame(root)
right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="n")

# Метка для отображения предпросмотра PDF
pdf_preview_label = tk.Label(right_frame)
pdf_preview_label.grid(row=0, column=0, padx=10, pady=10)
# Метка для отображения сообщения о загрузке
loading_label = tk.Label(right_frame, text="", fg="blue",font=("Arial", 36))
loading_label.grid(row=1, column=0, columnspan=2, pady=10)

# Первоначальное обновление предпросмотра (если файл уже существует)
if os.path.exists('output.docx'):
    save_and_update_preview(
        "Генеральному директору\nООО \"Умный склад\"\nПодгорному М.Ю",
        "Благодарственное письмо",
        "ТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекст"
        "ТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекс\n"
        "ТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекст\n"
        "ТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекстТекст",
        "Генеральный директор",
        "Д.С. Шербаков"
    )

# Запуск основного цикла
root.mainloop()
