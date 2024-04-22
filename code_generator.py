import pandas as pd
import numpy as np

class CodeGenerator:
    tags_list = list() # список для хранения данных о тэгах [Name, Data Type, Comment]

    FILE_EXTENSION= '.scl' # расширение файла с исходным кодом для контроллера
    DATA_TYPES= ['Bool', 'Word', 'Byte', 'DWord'] # допустимые типы данных тэгов
    NAME_COLUMN = 0 # номер столбца в файле .xlsx с названием тэга
    DATA_TYPE_COLUMN = 2 # номер столбца в файле .xlsx с типом тэга
    LOGICAL_ADDRESS_COLUMN = 3 # номер столбца в файле .xlsx с адресом тэга
    COMMENT_COLUMN = 4

    def __read_excel(self, excel_file:str) -> list:
        try:
            df = pd.read_excel(excel_file)
            rows = df.values.tolist()
            tags = list()
            tags = [r[0:5] for r in rows if r[self.LOGICAL_ADDRESS_COLUMN][1] == 'I' and r[self.DATA_TYPE_COLUMN] in self.DATA_TYPES]
            return tags
        except Exception as e:
            raise Exception(e)

    def generate_source(self, db_num:int, excel_path:str) -> str:
        '''Создается исходный код для блока данных DB и функции FC.
           В блоке данных создаётся копия тэгов входов контроллера.
           В функции происходит присваивание данных из блока данных в тэги входов.
           Метод принимает номер блока данных, функции и путь до файла .xlsx.
        '''
        tags = self.__read_excel(excel_path)
        var = str() # данных блока данных DB
        function_code = str() # код функции FC

        for tag in tags:
            if tag[self.COMMENT_COLUMN] is np.nan: # по умолчанию pandas вставляет nan, когда ячейка пустая
                tag[self.COMMENT_COLUMN] = ''
            self.tags_list.append([tag[self.NAME_COLUMN], tag[self.DATA_TYPE_COLUMN], tag[self.COMMENT_COLUMN]]) # формируется список [Name, Data Type, Comment]
            var += f'"{str(tag[self.NAME_COLUMN])}": {str(tag[self.DATA_TYPE_COLUMN]).upper()};' # код для блока данных DB '"Название" := "Тип";'
            function_code += f'{str(tag[self.LOGICAL_ADDRESS_COLUMN])} := DB{db_num}."{str(tag[self.NAME_COLUMN])}";' # код для функции FC
        
        # Исходный код для блока данных DB
        code = f'''DATA_BLOCK DB {db_num}
{"{S7_Optimized_Access := 'FALSE'}"}
VAR
{var}
END_VAR
BEGIN
END_DATA_BLOCK'''
        
        code += '\r\n' # переход на новую строку

        # Исходный код для функции FC
        code +=  f'''FUNCTION FC {db_num}:VOID
VAR_INPUT
    Enable: BOOL;
END_VAR
BEGIN
IF Enable Then
{function_code}
END_IF;
END_FUNCTION'''
        return code

    def write_code_to_file(self, path:str, code:str, encoding:str) -> None:
        try:
            #with open(path, 'w', encoding='cp1251', errors='replace') as file:
            with open(path, 'w', encoding=encoding, errors='replace') as file:
                file.flush()
                file.write(code)
        except Exception as e:
            raise Exception(e)