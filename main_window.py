from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import os
from tkinter.messagebox import showerror, showwarning, showinfo
import xml_generator
import code_generator

class MainWindows:
    __excel_path = str()
  
    def __init__(self, master):
        self.encoding = ['UTF-8', 'cp1251']
        # по умолчанию будет выбран первый элемент из languages
        self.encoding_var = StringVar(value=self.encoding[0]) 

        self.excel_path = str()
        self.master = master
        master.title("Кодогенератор")
        master.geometry("500x300")
        master.resizable(0,0)

        self.label = Label(master, text="Выбрать папку с файлом Excel")
        self.label.pack()

        self.greet_button = Button(master, text="Открыть проводник", command=self.path_select)
        self.greet_button.pack(ipadx=50, ipady=10)

        self.label_path = Label(master, text="Путь до файла", bg="#FFFFFF")
        self.label_path.pack(ipadx=100, ipady=10, pady=5)

        self.label_db_func = Label(master, text="Задать номер блока данных и функции")
        self.label_db_func.pack()

        self.number = Entry(master, bg="#FFFFFF")
        self.number.pack(ipadx=80, ipady=10, pady=5)
        self.number.insert(0, '1000')

        self.label_encoding = Label(master, text="Задать кодировку файла")
        self.label_encoding.pack()
        
        self.combobox = ttk.Combobox(textvariable=self.encoding_var, values=self.encoding)
        self.combobox.pack()

        self.close_button = Button(master, text="Генерировать .scl и .xml файлы", command=self.generate_source)
        self.close_button.pack(ipadx=50, ipady=10, pady=5)

    def show_win_error(self, error_message:str) -> None: 
        showerror(title="Ошибка", message= error_message)
    
    def show_win_info(self) -> None: 
        showinfo(title="Информация", message="Файл успешно создан!")
        
    def path_select(self) -> None:
        excel_path = filedialog.askopenfilenames(filetypes = (('xlsx files','*.xlsx'),))
        self.__excel_path  = excel_path[0]
        self.label_path["text"] = excel_path[0]
    
    def generate_source(self) -> None:
        code_gen = code_generator.CodeGenerator()
        xml_gen = xml_generator.XMLGenerator()
        try:
            excel_source = self.__excel_path
            if excel_source == '':
                raise Exception('Неправильно указан путь до файла')
            
            number = self.number.get()
            if number == '':
                raise Exception('Неправильно указан номер DB')
            db_number = int(number)
            if  (db_number < 0 or db_number > 5000):
                raise Exception('Неправильно указан номер DB')
            
            source = code_gen.generate_source(db_number, excel_source)
            current_directory = os.path.dirname(excel_source)
            source_path = f'{current_directory}\\Inputs emulation{code_gen.FILE_EXTENSION}'
            encoding = self.combobox.get()
            code_gen.write_code_to_file(source_path, source, encoding)

            if code_gen.tags_list is []:
                raise Exception("Не удалось создать .xml файл")
            xml_gen.generate_xml(db_number, code_gen.tags_list, excel_source)
            self.show_win_info()
        except Exception as e:
            self.show_win_error(f'{e}')