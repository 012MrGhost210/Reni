import os
import glob
import pandas as pd
import xml.etree.ElementTree as ET

def simple_xml_to_excel_converter():
    """
    Простой конвертер - берет все XML из папки 'xml' и сохраняет в папку 'excel'
    """
    # Определяем пути
    script_dir = os.path.dirname(os.path.abspath(__file__))
    xml_folder = os.path.join(script_dir, "xml")
    excel_folder = os.path.join(script_dir, "excel")
    
    # Создаем папки если их нет
    os.makedirs(xml_folder, exist_ok=True)
    os.makedirs(excel_folder, exist_ok=True)
    
    # Находим все XML файлы
    xml_files = glob.glob(os.path.join(xml_folder, "*.xml"))
    
    if not xml_files:
        print(f"⚠️ Поместите XML файлы в папку: {xml_folder}")
        return
    
    print(f"Найдено {len(xml_files)} XML файлов")
    
    # Обрабатываем каждый файл
    for xml_file in xml_files:
        filename = os.path.basename(xml_file)
        excel_name = os.path.splitext(filename)[0] + ".xlsx"
        excel_file = os.path.join(excel_folder, excel_name)
        
        try:
            # Читаем XML
            tree = ET.parse(xml_file)
            root = tree.getroot()
            
            # Собираем данные
            data = []
            for item in root:
                row = {}
                for elem in item:
                    if len(elem) == 0:  # Простые элементы
                        row[elem.tag] = elem.text
                    else:  # Вложенные элементы
                        for sub_elem in elem:
                            row[f"{elem.tag}_{sub_elem.tag}"] = sub_elem.text
                if row:
                    data.append(row)
            
            if data:
                # Сохраняем в Excel
                df = pd.DataFrame(data)
                df.to_excel(excel_file, index=False)
                print(f"✅ {filename} -> {excel_name}")
            else:
                print(f"⚠️ {filename}: нет данных для конвертации")
                
        except Exception as e:
            print(f"❌ Ошибка при обработке {filename}: {str(e)}")
    
    print(f"\n🎉 Готово! Excel файлы сохранены в: {excel_folder}")

# Автоматический запуск при выполнении скрипта
if __name__ == "__main__":
    simple_xml_to_excel_converter()
    input("\nНажмите Enter для выхода...")
