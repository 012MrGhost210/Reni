import pandas as pd
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def xml_to_excel_advanced(xml_file, excel_file, root_tag=None):
    """
    Универсальный конвертер XML в Excel
    """
    tree = ET.parse(xml_file)
    root = tree.getroot()
    
    # Определяем корневой элемент для данных
    if root_tag:
        data_root = root.find(root_tag)
    else:
        data_root = root
    
    # Собираем все уникальные теги
    all_tags = set()
    rows = []
    
    for item in data_root:
        row_data = {}
        for element in item.iter():
            if element != item:  # Пропускаем сам элемент-родитель
                if len(element) == 0:  # Только листовые элементы
                    row_data[element.tag] = element.text
                    all_tags.add(element.tag)
        rows.append(row_data)
    
    # Создаем DataFrame
    df = pd.DataFrame(rows, columns=list(all_tags))
    
    # Сохраняем в Excel с настройкой ширины колонок
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        
        # Автонастройка ширины колонок
        worksheet = writer.sheets['Data']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f"Конвертация завершена. Файл сохранен: {excel_file}")
    return df
