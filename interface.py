#interface.py

import csv
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, StringVar, Toplevel
import numpy as np
import data  # Предполагается, что ваши данные находятся в файле data.py
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from docx import Document
import tkinter.filedialog
import re
import math


class ToolTip(object):

    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0

    def showtip(self, text, new_x, new_y):
        "Display text in tooltip window"
        self.text = text
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() - 750 + new_x
        y = y + cy + self.widget.winfo_rooty() - 200 + new_y
        self.tipwindow = tw = Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(tw, text=self.text, justify="left",
                      background="#ffffff", relief="solid", borderwidth=1,
                      font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

def CreateToolTip(widget, text, new_x, new_y):
    toolTip = ToolTip(widget)
    def enter(event):
        toolTip.showtip(text, new_x, new_y)
    def leave(event):
        toolTip.hidetip()
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)

class ConsumerCalculator:
    def __init__(self, master):
        self.amount_users = []
        self.master = master
        self.master.title("Расчёт потребителей")

        # Создание фрейма для поля ввода t и кнопок
        input_frame = tk.Frame(master)
        input_frame.pack(side=tk.TOP, fill=tk.X)

        # Добавление метки перед полем ввода t
        self.label_t = tk.Label(input_frame, text="№ потреб.:")
        self.label_t.pack(side=tk.LEFT, padx=5, pady=5)
        CreateToolTip(self.label_t, text= 'Жилые дома квартирного типа:\n'
                      '     1) с водопроводом и канализацией без ванн\n'
                      '     2) с водопроводом, канализацией и ваннами с водонагревателями, работающими на твердом топливе\n'
                      '     3) с водопроводом, канализацией и ваннами с газовыми водонагревателями\n'
                      '     4) с централизованным горячим водоснабжением, оборудованные умывальниками, мойками и душами\n'
                      '     5) с сидячими ваннами, оборудованными душами\n'
                      '     6) с ваннами длиной от 1500 мм, оборудованными душами\n'
                      'Общежития: \n'
                      '     7) с общими душевыми\n'
                      '     8) с душами при всех жилых комнатах\n'
                      '     9) с общими кухнями и блоками душевых на этажах при жилых комнатах в каждой секции здания\n'
                      'Гостиницы пансионаты и мотели\n'
                      '     10) с общими ваннами и душами\n'
                      '     11) с душами во всех отдельных номерах\n'
                      '     12) с ваннами в отдельных номерах процент общего числа номеров до 25\n'
                      '     13) с ваннами в отдельных номерах процент общего числа номеров до 75\n'
                      '     14) с ваннами в отдельных номерах процент общего числа номеров до 100\n'
                      'Больницы \n'
                      '     15) с общими ваннами и душевыми\n'
                      '     16)с санузлами приближенными к палатам\n'
                      '     17) инфекционные\n' 
                      'Санатории и дома отдыха\n'
                      '     18) с общими душами\n'
                      '     19) с душами при всех жилых комнатах\n'
                      '     20) с ваннами при всех жилых комнатах\n'
                      'Поликлиники и амбулатории\n'
                      '     21) Поликлиники и амбулатории\n'
                      'Дошкольные образовательные организации\n'
                      '  C дневным пребыванием детей\n'
                      '     22)  со столовыми работающими на полуфабрикатах\n'
                      '     23) с дневным пребыванием детей со столовыми работающими на сырье и прачечными оборудованными автоматическими стиральными машинами\n'
                      '  C круглосуточным пребыванием детей\n'
                      '     24) со столовыми работающими на полуфабрикатах\n'
                      '     25) со столовыми работающими на сырье и прачечными оборудованными автоматическими стиральными машинами\n'
                      '     26) со столовыми работающими на полуфабрикатах\n'
                      '     27) со столовыми работающими на сырье и прачечными оборудованными автоматическими стиральными машинами\n'
                      'Прачечные \n'
                      '     28) механизированные\n'
                      '     29) немеханизированные\n'
                      'Административные здания\n'
                      '     30) Административные здания\n'
                      'Образовательные организации\n'
                      '     31) профессионального и высшего образования с душевыми при гимнастических залах и буфетами реализующими готовую продукцию\n'
                      'Лаборатории\n'
                      '     32) общеобразовательных организаций и организаций профессиональных и высшего образования\n'
                      'Общеобразовательные организации\n'
                      '     33) с душевыми при гимнастических залах и столовыми работающими на полуфабрикатах\n'
                      '     34) с продленным днем\n'
                      'Общеобразовательные организации интернаты\n'
                      '     35) с помещениями учебными с душевыми при гимнастических залах\n'
                      '     36) с помещениями спальными\n'
                      'Аптеки\n'
                      '     37) торговый зал и подсобные помещения\n'
                      '     38) лаборатория приготовления лекарств\n'
                      'Предприятия общественного питания для приготовления пищи реализуемой\n'
                      '     39) в обеденном зале\n'
                      '     40) на дом\n'
                      'Магазины\n'
                      '     41) продовольственные\n'
                      '     42) промтоварные\n'
                      'Парикмахерские\n'
                      '     43) Парикмахерские\n'
                      'Кинотеатры\n'
                      '     44) Кинотеатры\n'
                      'Клубы\n'
                      '     45) Клубы\n'
                      'Театры\n'
                      '     46) для зрителей\n'
                      '     47) для артистов\n'
                      'Стадионы и спортзалы'
                      '     48) для зрителей\n'
                      '     49) для физкультурников с учетом приема душа\n' 
                      '     50) для спортсменов с учетом приема душа\n'
                      'Плавательные бассейны\n'
                      '     51) пополнение бассейна\n' 
                      '     52) для зрителей\n'
                      '     53) для спортсменов с учетом приема душа\n'
                      'Бани\n'
                      '     54) для мытья в мыльной с тазами на скамьях и ополаскиванием в душе\n'
                      '     55) с приемом оздоровительных процедур и ополаскиванием в душе\n'
                      '     56) душевая кабина\n'
                      '     57) ванная кабина\n'
                      'Душевые в бытовых помещениях промышленных предприятий\n'
                      '     58) Душевые в бытовых помещениях промышленных предприятий\n'
                      'Цеха\n'
                      '     59) с тепловыделениями свыше 84 кДж на 1 м куб в час\n'
                      '     60) остальные цеха\n'
                      'Расход воды на поливку\n'
                      '     61) травяного покрова\n'
                      '     62) футбольного поля\n' 
                      '     63) остальных спортивных сооружений\n'
                      '     64) совершенствованных покрытий тротуаров площадей заводских проездов\n' 
                      '     65) зеленых насаждений газонов и цветников\n'
                      'Заливка\n'
                      '     66) поверхности катка', new_x = 0, new_y = 0)
        # Поле для ввода t
        self.t_entry = tk.Entry(input_frame)
        self.t_entry.pack(side=tk.LEFT, padx=5, pady=5)
        self.t_entry.insert(0, "Введите номер потребителя (t)")

        # Кнопки для добавления и удаления участка, размещенные рядом с выбором диаметра
        add_section_button = tk.Button(input_frame, text="Добавить новый участок", command=lambda: self.add_section())
        add_section_button.pack(side=tk.LEFT)

        add_section_label = tk.Label(input_frame, text="Добавить к")
        add_section_label.pack(side=tk.LEFT)
        CreateToolTip(add_section_label, text='Выбор положения для добавления участка.\n'
                      'Добавить к концу - добавляет участок в конец\n'
                      'При выборе числового значения - добавляет участок после выбранного значения', new_x = 325, new_y = 200)
        
        self.add_combobox = ttk.Combobox(input_frame, state="readonly", values = "Концу", width=20)
        self.values = tuple(self.add_combobox["values"])
        self.add_combobox.pack(side=tk.LEFT, padx=5, pady=5)
        self.add_combobox.current(0)
        

        # Кнопка для расчета
        self.calculate_table_button = tk.Button(input_frame, text='Рассчитать по\n табличным данным', command=self.calculate_table, width=20)
        self.calculate_table_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.calculate_gidr_button = tk.Button(input_frame, text='Рассчитать по\n формуле гидравлики', command=self.calculate_gidr, width=20)
        self.calculate_gidr_button.pack(side=tk.LEFT, padx=5, pady=5)

        # Кнопка для отображения результатов расчета
        self.results_button = tk.Button(input_frame, text="Результаты расчёта", command=self.show_results, width=20)
        self.results_button.pack(side=tk.LEFT, padx=5, pady=5)

        # Кнопка для сохранения результатов
        self.save_button = tk.Button(input_frame, text="Сохранить результаты", command=self.save_results, width=20)
        self.save_button.pack(side=tk.LEFT, padx=5, pady=5)

        # Создание фрейма для прокрутки
        self.frame = tk.Frame(self.master)
        self.frame.pack(fill=tk.BOTH, expand=True)

        # Создание Canvas для вертикальной и горизонтальной прокрутки
        self.canvas = tk.Canvas(self.frame)
        self.scrollbar_y = tk.Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar_x = tk.Scrollbar(self.frame, orient="horizontal", command=self.canvas.xview)
        self.scrollable_frame = tk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self.canvas.configure(yscrollcommand=self.scrollbar_y.set)
        self.canvas.configure(xscrollcommand=self.scrollbar_x.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.entries = []  # Список для хранения полей ввода
        self.results_data = []  # Список для хранения результатов расчетов
        self.t_string_value = ""  # Переменная для хранения значения t_string
        self.among_entries = []
        self.added = []
        self.u_entry = None
        self.diam_var = None
        # # Добавляем первый участок по умолчанию
        self.add_section()

    def add_section(self):
        frame = tk.Frame(self.scrollable_frame, name = "frame"+str(len(self.amount_users)+1))
        # Поля для ввода номера потребителя
        i_entry = tk.Label(frame, text= "Номер участка: ")
        i_entry.pack(side=tk.LEFT)
        k_label = tk.Label(frame, text= str(len(self.amount_users)+1))
        k_label.pack(side=tk.LEFT)

        # Поля для ввода U и диаметра
        self.u_entry = tk.Entry(frame)
        self.u_entry.pack(side=tk.LEFT)
        self.u_entry.insert(0, "Введите U")

        self.diam_var = StringVar(frame)
        self.diam_var.set(data.diam_values[0])  # Устанавливаем значение по умолчанию
        diam_menu = tk.OptionMenu(frame, self.diam_var, *data.diam_values)
        diam_menu.pack(side=tk.LEFT)

        delete_button = tk.Button(frame, text="Удалить участок", command=lambda: self.remove_section(frame))
        delete_button.pack(side=tk.LEFT)
        enumerate(str(self.entries))
        
        if self.add_combobox.get().isdigit() and int(self.add_combobox.get())>1:
            indx = 0
            puth = str(self.entries[0][2])
            while puth[-1].isdigit():
                puth = puth[:-1]
            found = str(puth+str(int(self.add_combobox.get())))
            for i in range(len(self.entries)):
                if found in str(self.entries[i][2]):
                    indx = i
            frame.pack(after=found)
            self.entries.insert(indx+1, (self.u_entry, self.diam_var, frame, k_label))
        elif self.add_combobox.get().isdigit() and int(self.add_combobox.get()) == 1:
            frame.pack(after=str(self.entries[0][2]))
            self.entries.insert(1, (self.u_entry, self.diam_var, frame, k_label))
        else:
            frame.pack()
            self.entries.append((self.u_entry, self.diam_var, frame, k_label))
        
        self.amount_users.append(len(self.amount_users)+1)
        self.add_combobox["values"] = self.values + tuple(self.amount_users)
                
    def remove_section(self, frame):
        # Удаление указанного участка и обновление интерфейса
        frame.pack_forget()  # Скрываем фрейм
        self.entries = [entry for entry in self.entries if entry[2] != frame]
        puth = re.findall(r'\d+', str(frame))
        self.amount_users.remove(int(puth[1]))      
        self.add_combobox["values"] = ''
        self.add_combobox["values"] = self.values + tuple(self.amount_users)

    def clear_entries(self):
        for entry, _, _ in self.entries:
            entry.delete(0, tk.END)
            entry.insert(0, "")
        self.t_entry.delete(0, tk.END)
        self.t_entry.insert(0, "Введите номер потребителя (t)")

    def calculate_gidr(self):
        self.results_data.clear()  # Очищаем предыдущие результаты
        try:
            # Получаем значение t из поля ввода
            t_input = float(self.t_entry.get())
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректное значение для t.")
            return

        for u_entry, diam_var, frame, k_label in self.entries:
            try:
                # Получаем значения U и диаметр
                U = float(u_entry.get())
                diam_input = int(diam_var.get())
                k_input = int(k_label.cget("text"))

                if t_input not in data.t_values:
                    raise ValueError(f"Значение t={t_input} не найдено в массиве.")

                index = np.where(data.t_values == t_input)[0][0]
                q_c_hru = data.q_c_hru_values[index]
                q_c_0 = data.q_c_0_values[index]
                t_num = data.t_num_values[index]
                t_string = data.t_string_values[index]

                # Сохраняем значение t_string для отображения
                self.t_string_value = t_string

                # Получаем значения q и v для выбранного диаметра
                q_values = data.q_dict[diam_input]
                v_values = data.v_dict[diam_input]

                x_input = (q_c_hru * U) / (3600 * q_c_0)
                y_output = self.interpolate_or_extrapolate(x_input, data.x_values, data.y_values)
                Q_output = 5 * (q_c_0) * (y_output)

                for i in range(len(q_values) - 1):
                    v_interpolated = (4*(Q_output/1000))/(math.pi*(diam_input/1000)**2)

                if v_interpolated is None:
                    raise ValueError("Не удалось вычислить значение скорости.")

                # Сохраняем результаты
                result = [k_input, t_input, U, diam_input, q_c_hru, q_c_0, t_num, t_string, round(x_input, 4),
                          round(y_output, 4), round(Q_output, 4), round(v_interpolated, 4)]
                self.results_data.append(result)

            except ValueError as e:
                messagebox.showerror("Ошибка", str(e))
                return  # Прерываем выполнение функции при ошибке
            messagebox.showinfo("Успех", "Расчет по гидравлике выполнен успешно.")

    def calculate_table(self):
        self.results_data.clear()  # Очищаем предыдущие результаты
        try:
            # Получаем значение t из поля ввода
            t_input = float(self.t_entry.get())
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректное значение для t.")
            return

        for u_entry, diam_var, frame, k_label in self.entries:
            try:
                # Получаем значения U и диаметр
                U = float(u_entry.get())
                diam_input = int(diam_var.get())
                k_input = int(k_label.cget("text"))

                if t_input not in data.t_values:
                    raise ValueError(f"Значение t={t_input} не найдено в массиве.")

                index = np.where(data.t_values == t_input)[0][0]
                q_c_hru = data.q_c_hru_values[index]
                q_c_0 = data.q_c_0_values[index]
                t_num = data.t_num_values[index]
                t_string = data.t_string_values[index]

                # Сохраняем значение t_string для отображения
                self.t_string_value = t_string

                # Получаем значения q и v для выбранного диаметра
                q_values = data.q_dict[diam_input]
                v_values = data.v_dict[diam_input]

                x_input = (q_c_hru * U) / (3600 * q_c_0)
                y_output = self.interpolate_or_extrapolate(x_input, data.x_values, data.y_values)
                Q_output = 5 * (q_c_0) * (y_output)

                # Линейная интерполяция или экстраполяция для v
                v_interpolated = None
                if Q_output < q_values[0]:
                    q0, q1 = q_values[0], q_values[1]
                    v0, v1 = v_values[0], v_values[1]
                    v_interpolated = v0 + (v1 - v0) * (Q_output - q0) / (q1 - q0)
                elif Q_output > q_values[-1]:
                    q0, q1 = q_values[-2], q_values[-1]
                    v0, v1 = v_values[-2], v_values[-1]
                    v_interpolated = v0 + (v1 - v0) * (Q_output - q0) / (q1 - q0)
                else:
                    for i in range(len(q_values) - 1):
                        if q_values[i] <= Q_output <= q_values[i + 1]:
                            q0, q1 = q_values[i], q_values[i + 1]
                            v0, v1 = v_values[i], v_values[i + 1]
                            v_interpolated = v0 + (v1 - v0) * (Q_output - q0) / (q1 - q0)
                            break

                if v_interpolated is None:
                    raise ValueError("Не удалось вычислить значение скорости.")

                # Сохраняем результаты
                result = [k_input, t_input, U, diam_input, q_c_hru, q_c_0, t_num, t_string, round(x_input, 4),
                          round(y_output, 4), round(Q_output, 4), round(v_interpolated, 4)]
                self.results_data.append(result)

            except ValueError as e:
                messagebox.showerror("Ошибка", str(e))
                return  # Прерываем выполнение функции при ошибке

        # Сообщение о выполненном расчете
        messagebox.showinfo("Успех", "Расчет по таблице выполнен успешно.")

    def interpolate_or_extrapolate(self, x_input, x_values, y_values):
        if x_input < x_values[0]:
            x0, x1 = x_values[0], x_values[1]
            y0, y1 = y_values[0], y_values[1]
            return y0 + (y1 - y0) * (x_input - x0) / (x1 - x0)
        elif x_input > x_values[-1]:
            x0, x1 = x_values[-2], x_values[-1]
            y0, y1 = y_values[-2], y_values[-1]
            return y0 + (y1 - y0) * (x_input - x0) / (x1 - x0)
        else:
            for i in range(len(x_values) - 1):
                if x_values[i] <= x_input <= x_values[i + 1]:
                    x0, x1 = x_values[i], x_values[i + 1]
                    y0, y1 = y_values[i], y_values[i + 1]
                    return y0 + (y1 - y0) * (x_input - x0) / (x1 - x0)

    def show_results(self):
        # Открываем новое окно для отображения результатов
        results_window = Toplevel(self.master)
        results_window.title("Результаты расчёта")

        # Получаем значение t_string из последнего результата
        if self.results_data:
            self.t_string_value = self.results_data[0][7]  # Предполагаем, что t_string одинаков для всех результатов

        # Создаем метку для t_string
        t_string_label = tk.Label(results_window, text=f"Потребитель: {self.t_string_value}", font=("Arial", 12))
        t_string_label.pack(pady=5)

        # Создание заголовков таблицы
        headers = ["Участок", "t", "U", "D", "q_(c)hru", "q_(c)0", "Группа (t_num)", "Потребитель", "Pc*N", "alpha", "Q", "Скорость"]

        header_frame = tk.Frame(results_window)
        for header in headers:
            label = tk.Label(header_frame, text=header, borderwidth=1, relief="solid", width=15)
            label.pack(side=tk.LEFT)
        header_frame.pack()

        # Создание области для результатов
        results_frame = tk.Frame(results_window)
        results_frame.pack()

        for result in self.results_data:
            result_row = tk.Frame(results_frame)
            result_row.pack()

            for res in result:
                label = tk.Label(result_row, text=str(res), borderwidth=1, relief="solid", width=15)
                label.pack(side=tk.LEFT)

    def save_results(self):
        file_path = tk.filedialog.asksaveasfilename(defaultextension=".csv",
                                                    filetypes=[("CSV files", ".csv"), ("Excel files", ".xlsx"),
                                                               ("Word documents", ".docx")])
        if file_path:
            extension = file_path.split('.')[-1].lower()
            if extension == 'csv':
                self.save_to_csv(file_path)
            elif extension == 'xlsx':
                self.save_to_excel(file_path)
            elif extension == 'docx':
                self.save_to_docx(file_path)
            else:
                messagebox.showerror("Неверный формат файла.",
                                     "Файл должен быть сохранен в формате CSV, XLSX или DOCX.")
        else:
            messagebox.showwarning("Файл не сохранён.", "Не удалось сохранить файл.")

    def save_to_excel(self, file_path):
        wb = Workbook()
        ws = wb.active
        headers = ["t", "U", "D", "q_(c)hru", "q_(c)0", "Группа (t_num)", "Потребитель", "Pc*N", "alpha", "Q", "Скорость"]
        ws.append(headers)
        for result in self.results_data:
            ws.append(result)
        wb.save(file_path)
        messagebox.showinfo("Успех", "Результаты сохранены в Excel-файле.")

    def save_to_docx(self, file_path):
        document = Document()
        table = document.add_table(rows=len(self.results_data)+1, cols=11)
        headers = ["t", "U", "D", "q_(c)hru", "q_(c)0", "Группа (t_num)", "Потребитель", "Pc*N", "alpha", "Q", "Скорость"]
        hdr_cells = table.rows[0].cells
        for i, header in enumerate(headers):
            hdr_cells[i].text = header
        for row_idx, result in enumerate(self.results_data):
            row_cells = table.rows[row_idx+1].cells
            for col_idx, item in enumerate(result):
                row_cells[col_idx].text = str(item)
        document.save(file_path)
        messagebox.showinfo("Успех", "Результаты сохранены в DOCX-файле.")

    def save_to_csv(self, file_path):
        with open(file_path, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            headers = ["t", "U", "D", "q_(c)hru", "q_(c)0", "Группа (t_num)", "Потребитель", "Pc*N", "alpha", "Q", "Скорость"]
            writer.writerow(headers)
            for result in self.results_data:
                writer.writerow(result)
        messagebox.showinfo("Успех", "Результаты сохранены в CSV-файле.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ConsumerCalculator(root)
    root.mainloop()