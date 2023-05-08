import tkinter as tk
import tkinter.ttk as ttk
from tkinter import *
from tkinter.ttk import *
import sqlite3
import os
from openpyxl import Workbook
from tkinter import messagebox



if not os.path.exists("employees.db"):
    conn = sqlite3.connect('employees.db')
    conn.execute('''CREATE TABLE employees
                     (name TEXT NOT NULL,
                     employee_id INTEGER NOT NULL,
                     profession TEXT NOT NULL,
                     hire_date TEXT NOT NULL,
                     labor_book_num_serial TEXT NOT NULL,
                     order_date TEXT NOT NULL,
                     order_num TEXT NOT NULL,
                     phone_num TEXT NOT NULL,
                     address TEXT NOT NULL,
                     family_data TEXT NOT NULL,
                     child_num TEXT NOT NULL,
                     child_name_birthday TEXT NOT NULL,
                     name_husband_wife);''')
else:
    pass


class EmployeeDatabase:
    def __init__(self, master):
        self.master = master
        master.title("Employee Database")



        root.geometry("800x600")

        # Робимо фрейм для таблиці и полос прокрутки
        table_frame = Frame(root)
        table_frame.pack(fill=BOTH, expand=1)

        self.table = ttk.Treeview(table_frame, columns=("col1", "col2", "col3", "col4", "col5", "col6", "col7", "col8", "col9", "col10", "col11", "col12", "col13"))
        self.table.column("#0", width=0)
        self.table.heading("col1", text="ПІБ")
        self.table.heading("col2", text="Табельний номер")
        self.table.heading("col3", text="Професія")
        self.table.heading("col4", text="Дата прийому на роботу")
        self.table.heading("col5", text="Серія/номер трудової книжки")
        self.table.heading("col6", text="Дата приказу про прийом на роботу")
        self.table.heading("col7", text="Номер приказу про прийом на роботу")
        self.table.heading("col8", text="Номер телефону")
        self.table.heading("col9", text="Адреса")
        self.table.heading("col10", text="Дані про шлюб")
        self.table.heading("col11", text="Кількість дітей")
        self.table.heading("col12", text="Ім'я дитини/дата народження")
        self.table.heading("col13", text="Ім’я жінки(чоловіка)")


        # Отримуємо інформацію з БД
        conn = sqlite3.connect("employees.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM employees")
        rows = cursor.fetchall()

        # Заповнюємо таблицю інформацією
        for row in rows:
            self.table.insert("", "end", values=row)

        # Закриваємо з єднання з БД
        cursor.close()
        conn.close()

        self.table.pack(side="bottom", padx=10, pady=10)

        # полоси прокрутки
        xscrollbar = Scrollbar(table_frame, orient=HORIZONTAL, command=self.table.xview)
        yscrollbar = Scrollbar(table_frame, orient=VERTICAL, command=self.table.yview)
        self.table.configure(xscrollcommand=xscrollbar.set, yscrollcommand=yscrollbar.set)

        # упаковуємо таблицю и полоси прокрутки
        yscrollbar.pack(side=RIGHT, fill=Y)
        xscrollbar.pack(side=BOTTOM, fill=X)
        self.table.pack(fill=BOTH, expand=1)


        #Кнопки
        self.add_button = tk.Button(master, text="Додати інформацію", command=self.add_employee)
        self.edit_button = tk.Button(master, text="Змінити інформацію", command=self.update_employee)
        self.delete_button = tk.Button(master, text="Видалити інформацію", command=self.delete_employee)
        self.download_button = tk.Button(master, text="Завантажити інформацію", command=self.download_employee)

        # Разміщення кнопок на екрані
        self.add_button.pack(side="left", fill="x", padx=10, pady=10)
        self.edit_button.pack(side="left", fill="x", padx=10, pady=10)
        self.delete_button.pack(side="left", fill="x", padx=10, pady=10)
        self.download_button.pack(side="left", fill="x", padx=10, pady=10)





    # Функція, яка спрацьовує при натисненні на кнопку "Додати інформацію"
    def add_employee(self):
        add_window = tk.Toplevel(self.master)
        add_window.geometry("770x250")
        add_window.title("Додати інформацію")

        # Створення текстового поля и помітки
        Label(add_window, text="ПІБ").grid(row=1, column=0)
        pib_entry = Entry(add_window)
        pib_entry.grid(row=1, column=1)

        Label(add_window, text="Професія").grid(row=2, column=0)
        prof_entry = Entry(add_window)
        prof_entry.grid(row=2, column=1)

        Label(add_window, text="Табельний номер").grid(row=3, column=0)
        tnum_entry = Entry(add_window)
        tnum_entry.grid(row=3, column=1)

        Label(add_window, text="Дата прийому на роботу").grid(row=4, column=0)
        date_entry = Entry(add_window)
        date_entry.grid(row=4, column=1)

        Label(add_window, text="Серія і номер трудової книжки").grid(row=5, column=0)
        trudovik_entry = Entry(add_window)
        trudovik_entry.grid(row=5, column=1)

        Label(add_window, text="Дата приказу про прийом на роботу").grid(row=6, column=0)
        orderdate_entry = Entry(add_window)
        orderdate_entry.grid(row=6, column=1)

        Label(add_window, text="Номер приказу про прийом на роботу").grid(row=7, column=0)
        ordernum_entry = Entry(add_window)
        ordernum_entry.grid(row=7, column=1)

        Label(add_window, text="Адреса").grid(row=1, column=2)
        address_entry = Entry(add_window)
        address_entry.grid(row=1, column=3)

        Label(add_window, text="Телефон").grid(row=2, column=2)
        phone_entry = Entry(add_window)
        phone_entry.grid(row=2, column=3)

        Label(add_window, text="Дані про шлюб").grid(row=3, column=2)
        marriage_entry = Entry(add_window)
        marriage_entry.grid(row=3, column=3)

        Label(add_window, text="Кількість дітей").grid(row=4, column=2)
        kids_entry = Entry(add_window)
        kids_entry.grid(row=4, column=3)

        Label(add_window, text="Ім'я дитини/дата народження").grid(row=5, column=2)
        spouse_entry = Entry(add_window)
        spouse_entry.grid(row=5, column=3)

        Label(add_window, text="Прізвище та ім'я жінки/чоловіка").grid(row=6, column=2)
        spouse_name_entry = Entry(add_window)
        spouse_name_entry.grid(row=6, column=3)


        def save_info():
            # Отримуємо інфо з текстових полей
            name = pib_entry.get()
            employee_id = tnum_entry.get()
            profession = prof_entry.get()
            hire_date = date_entry.get()
            book_series_num = trudovik_entry.get()
            order_date = orderdate_entry.get()
            order_num = ordernum_entry.get()
            phone_num = phone_entry.get()
            address = address_entry.get()
            marriage_data = marriage_entry.get()
            children_num = kids_entry.get()
            child_name_birth = spouse_entry.get()
            spouse_name = spouse_name_entry.get()
            if not name or not employee_id or not profession or not hire_date or not book_series_num or not order_date or not order_num or not phone_num or not address or not marriage_data or not children_num or not child_name_birth or not spouse_name:
                messagebox.showerror(title="Помилка", message="Усі поля повинні бути заповнені (у випадку відсутності інформації ставте знак '-')")
            else:
                conn = sqlite3.connect('employees.db')
                cursor = conn.cursor()

                # Виконуємо SQL-запит на додавання інформації в таблицю
                cursor.execute('INSERT INTO employees VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                               (name, employee_id, profession, hire_date, book_series_num, order_date,
                                order_num, phone_num, address, marriage_data, children_num, child_name_birth, spouse_name))
                # Зберігаємо
                conn.commit()

                notif_window = tk.Tk()
                notif_window.geometry("480x70")
                notif_window.title("Notification")
                label = tk.Label(notif_window, text="Інформація обновиться при наступному запуску додатку", font=("Arial", 13))
                label.pack()

                def notif_info():
                    notif_window.destroy()

                agree_button = Button(notif_window, text="Ок", command=notif_info)
                agree_button.pack()



                # закриваємо з єднання
                conn.close()

                add_window.destroy()



        save_button = Button(add_window, text="Зберегти", command=save_info)
        save_button.grid(row=8, column=1, rowspan=8, padx=50, pady=20, sticky=N + S + E)


    # Функція, яка спрацьовує при натисненні на кнопку"Змінити інформацію"
    def update_employee(self):
        # Створення вікна
        window = tk.Tk()
        window.title("Зміна інформації про співробітника")
        window.geometry("470x600")

        # Функція для зміни інформації про співробітника
        def update_employee_db():
            # Получение табельного номера сотрудника из поля ввода
            emp_id = entry_id.get()

            # Коннект до БД и отримання наявних значень полів співробітника
            conn = sqlite3.connect('employees.db')
            c = conn.cursor()
            c.execute("SELECT * FROM employees WHERE employee_id = ?", (emp_id,))
            current_values = c.fetchone()
            if current_values is None:
                # Якщо співробітника з таким табельним номером не знайдено, вивести повідомлення про помилку
                messagebox.showerror("Помилка", f"Співробітника з табельним номером {emp_id} не знайдено")
            else:

                # Отримання нових значень поля из полей ввода або наявних значень, якщо поля залишили порожніми
                name = entry_name.get() if entry_name.get() else current_values[0]
                profession = entry_pos.get() if entry_pos.get() else current_values[2]
                hire_date = entry_hire_date.get() if entry_hire_date.get() else current_values[3]
                labor_book = entry_labor_book.get() if entry_labor_book.get() else current_values[4]
                order_date = entry_order_date.get() if entry_order_date.get() else current_values[5]
                order_num = entry_order_num.get() if entry_order_num.get() else current_values[6]
                phone = entry_phone.get() if entry_phone.get() else current_values[7]
                address = entry_address.get() if entry_address.get() else current_values[8]
                marriage = entry_marriage.get() if entry_marriage.get() else current_values[9]
                child_num = entry_child_num.get() if entry_child_num.get() else current_values[10]
                child_name_bd = entry_child_name_bd.get() if entry_child_name_bd.get() else current_values[11]
                suprug_name = entry_suprug_name.get() if entry_suprug_name.get() else current_values[12]

                c = conn.cursor()
                c.execute(
                    "UPDATE employees SET name = ?, profession = ?, hire_date = ?, labor_book_num_serial = ?, order_date = ?, order_num = ?, phone_num = ?, address = ?, family_data = ?, child_num = ?, child_name_birthday = ?, name_husband_wife = ?  WHERE employee_id = ?",
                    (
                    name, profession, hire_date, labor_book, order_date, order_num, phone, address, marriage, child_num,
                    child_name_bd, suprug_name, emp_id))

                # Підтвердження внесення змін до бази даних
                conn.commit()

                # Закриття з'єднання з базою даних
                conn.close()

                # Створення вікна підтвердження
                notifi_window = tk.Tk()
                notifi_window.geometry("480x70")
                notifi_window.title("Notification")

                # Функція для закриття вікна підтвердження
                def close_noif():
                    notifi_window.destroy()

                label = tk.Label(notifi_window, text="Інформація успішно оновлена!", font=("Arial", 13))
                agree_button = tk.Button(notifi_window, text="Ок", command=close_noif)
                label.pack()
                agree_button.pack()

        # Створення віджетів для вводу нової інформації про співробітника та кнопки зміни
        label_id = tk.Label(window, text="Табельний номер:")
        label_id.pack()
        entry_id = tk.Entry(window)
        entry_id.pack()

        label_name = tk.Label(window, text="ПІБ:")
        label_name.pack()
        entry_name = tk.Entry(window)
        entry_name.pack()

        label_pos = tk.Label(window, text="Посада:")
        label_pos.pack()
        entry_pos = tk.Entry(window)
        entry_pos.pack()

        label_salary = tk.Label(window, text="Дата прийому на роботу:")
        label_salary.pack()
        entry_hire_date = tk.Entry(window)
        entry_hire_date.pack()

        label_salary = tk.Label(window, text="Серія\номер трудової книжки:")
        label_salary.pack()
        entry_labor_book = tk.Entry(window)
        entry_labor_book.pack()

        label_salary = tk.Label(window, text="Дата приказу про прийом на роботу:")
        label_salary.pack()
        entry_order_date = tk.Entry(window)
        entry_order_date.pack()

        label_salary = tk.Label(window, text="Номер приказу про прийом на роботу:")
        label_salary.pack()
        entry_order_num = tk.Entry(window)
        entry_order_num.pack()

        label_salary = tk.Label(window, text="Номер телефону:")
        label_salary.pack()
        entry_phone = tk.Entry(window)
        entry_phone.pack()

        label_salary = tk.Label(window, text="Адреса:")
        label_salary.pack()
        entry_address = tk.Entry(window)
        entry_address.pack()

        label_salary = tk.Label(window, text="Дані про шлюб:")
        label_salary.pack()
        entry_marriage = tk.Entry(window)
        entry_marriage.pack()

        label_salary = tk.Label(window, text="Кількість дітей:")
        label_salary.pack()
        entry_child_num = tk.Entry(window)
        entry_child_num.pack()

        label_salary = tk.Label(window, text="І'мя дитини\дата народження:")
        label_salary.pack()
        entry_child_name_bd= tk.Entry(window)
        entry_child_name_bd.pack()

        label_salary = tk.Label(window, text="І'мя жінки\чоловіка:")
        label_salary.pack()
        entry_suprug_name = tk.Entry(window)
        entry_suprug_name.pack()

        button = tk.Button(window, text="Змінити", command=update_employee_db)
        button.pack()

        text_widget = tk.Text(window, height=10, width=55)
        text_widget.insert(tk.END, "Для зміни інформаціЇ введіть табельний номер працівника(Обов'язково!) та заповніть інформацією поле, яке хочете змінити")
        text_widget.pack()

    # Функція, яка спрацьовує при натисненні на кнопку "Видалити інформацію"
    def delete_employee(self):
        # Створення вікна
        window = tk.Toplevel()
        window.title("Видалення співробітника")
        window.geometry("400x150")

        # Функція для видалення співробітника
        def delete_employee():
            # Підключення до бази даних
            conn = sqlite3.connect('employees.db')
            c = conn.cursor()

            # Отримання табельного номера співробітника з поля введення
            emp_id = entry.get()

            # Перевірка наявності табельного номера у базі даних
            c.execute("SELECT * FROM employees WHERE employee_id = ?", (emp_id,))
            row = c.fetchone()
            if row is None:
                # Якщо співробітника з таким табельним номером не знайдено, вивести повідомлення про помилку
                messagebox.showerror("Помилка", f"Співробітника з табельним номером {emp_id} не знайдено")
            else:
                # Якщо співробітника з таким табельним номером знайдено, виконати запит на його видалення з бази даних
                c.execute("DELETE FROM employees WHERE employee_id = ?", (emp_id,))

                # Підтвердження внесення змін до бази даних
                conn.commit()

                # Закриття з'єднання з базою даних
                conn.close()

                # Вивести повідомлення про успішне видалення співробітника
                messagebox.showinfo("Успіх",f"Інформація про співробітника з табельним номером {emp_id} успішно видалена з бази даних")

        # Створення віджетів для вводу табельного номера та кнопки видалення
        label = tk.Label(window, text="Табельний номер:")
        label.pack()
        entry = tk.Entry(window)
        entry.pack()
        button = tk.Button(window, text="Видалити", command=delete_employee)
        button.pack()

        window.mainloop()

        # Функція, яка спрацьовує при натисненні на кнопку "Видалити інформацію"
    def download_employee(self):
        conn = sqlite3.connect('employees.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM employees")
        employees = cursor.fetchall()

        # Створення нового файлу Excel
        workbook = Workbook()
        worksheet = workbook.active

        worksheet.append(['ПІБ', 'Табельний номер', 'Професія', 'Дата прийому на роботу',
                          'Серія і номер трудової книжки', 'Дата приказу про прийом на роботу',
                          'Номер приказу про прийом на роботу', 'Номер телефону', 'Адреса', 'Дані про шлюб','Кількість дітей',
                          'Ім’я та дата народження дітей', 'Прізвище та ім’я жінки (чоловіка)'])

        for employee in employees:
            # Додавання рядка з даними працівника до файлу Excel
            row = list(employee)
            worksheet.append(row)

        # Збереження файлу Excel
        workbook.save('employees.xlsx')

        notif_window = tk.Tk()
        notif_window.geometry("480x70")
        notif_window.title("Notification")
        label = tk.Label(notif_window, text="Успішно завантажено", font=("Arial", 13))
        label.pack()

        conn.close()

root = tk.Tk()
EmployeeDatabase(root)
root.mainloop()
