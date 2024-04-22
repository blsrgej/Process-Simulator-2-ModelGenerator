from tkinter.messagebox import showerror, showwarning, showinfo
import xml.etree.ElementTree as ET
import os

NAME = 0
DATA_TYPE = 1
COMMENT = 2

class XMLGenerator:
    def generate_xml(self, db_number:int, tags_list:list, excel_source:str) -> None:
        '''Создается .xml файл для программы Process-Simulator 2.
           В файле содержится информации о блоке данных DB в контроллере,
           настройка подключения к контроллеру.
        '''
        index_byte = 0
        index_bit = 0
        edge_bit = False
        try:
            items = ET.Element('Items')

            for tag in tags_list:
                if tag[NAME] == '' or tag[DATA_TYPE] == '':
                    continue

                item = ET.Element('Item', Name=tag[NAME], Comment=tag[COMMENT])

                if tag[DATA_TYPE] == 'Bool':
                    prop = ET.Element('Properties', DataType="S7_Bit", Byte=str(index_byte), Bit=str(index_bit), Signed="False", MemoryType="DB", DB=str(db_number))
                    item.append(prop)
                    items.append(item)
                    index_bit += 1
                    edge_bit = True
                    if index_bit >= 8:
                        index_bit = 0
                        index_byte += 1
                        edge_bit = False
                elif tag[DATA_TYPE] == 'Word' or tag[DATA_TYPE] == 'Int':
                    if edge_bit == True:
                        index_byte += 1
                        index_bit = 0
                        edge_bit = False
                    prop = ET.Element('Properties', DataType="S7_Word", Byte=str(index_byte), Signed="False",  MemoryType="DB", DB=str(db_number))
                    item.append(prop)
                    items.append(item)
                    index_byte += 2
                elif tag[DATA_TYPE] == 'Real':
                    if edge_bit == True:
                        index_byte += 1
                        index_bit = 0
                        edge_bit = False
                    prop = ET.Element('Properties', DataType="S7_DoubleWord", Byte=str(index_byte), Signed="False", MemoryType="DB", DB=str(db_number))
                    item.append(prop)
                    items.append(item)
                    index_byte += 4
    
            description = ET.Element('Description')
            root = ET.Element('ProcessSimulator', WindowState="Normal", Top="72", Left="66", Height="954", Width="1936", StayOnTop="False")
            description.text = '![CDATA[]]'
            root.append(description)
    
            communication = ET.Element('Communication')
            connection = ET.Element('Connection', Name='PLC', Type='S7IsoTCP')
            properties = ET.Element('Properties', IP='0.0.0.0', Rack="0", Slot="1", Type="PG", Slowdown="0", ErrorsBeforeDisconnect="3")
            connection.append(properties)
            connection.append(items)
            communication.append(connection)
            root.append(communication)
                
            # Создание файла .xml
            current_directory = os.path.dirname(excel_source)
            my_file = open(f"{current_directory}\\Inputs emulation.xml" , "wb")
            etree = ET.ElementTree(root)
            etree.write(my_file, encoding='utf-8', xml_declaration=False)
            my_file.close()
        except Exception as e:
            raise Exception(e)