import re
import xlrd

def clean_number(value):
    """Преобразует '1 784 222,90 руб.' в число 1784222.90"""
    if value is None:
        return None
    
    if isinstance(value, (int, float)):
        return float(value)
    
    value_str = str(value).strip()
    print(f"  raw: '{value_str}'")  # Отладка
    
    if 'руб' in value_str.lower():
        value_str = value_str.lower().split('руб')[0].strip()
        print(f"  после удаления руб: '{value_str}'")
    
    value_str = value_str.replace(' ', '')
    print(f"  после удаления пробелов: '{value_str}'")
    
    value_str = value_str.replace(',', '.')
    print(f"  после замены запятой: '{value_str}'")
    
    value_str = re.sub(r'[^\d.]', '', value_str)
    print(f"  после очистки: '{value_str}'")
    
    try:
        return float(value_str)
    except:
        return None

# Путь к файлу
file_path = r"\\fs-01.renlife.com\alldocs\Инвестиционный департамент\7.0 Treasury\Фонд СЧА\test4444.xls"

# Открываем файл
wb = xlrd.open_workbook(file_path, formatting_info=False)
sheet = wb.sheet_by_index(0)

# Ищем строку с фразой
search_text = "Итого стоимость чистых активов"

for row_idx in range(sheet.nrows):
    cell_value = sheet.cell(row_idx, 0).value
    if cell_value and isinstance(cell_value, str):
        if search_text.lower() in cell_value.lower():
            print(f"Найдено в строке {row_idx}")
            
            # Берем значение из столбца P (индекс 15)
            raw_value = sheet.cell(row_idx, 15).value
            print(f"Значение в P: '{raw_value}'")
            print(f"Тип значения: {type(raw_value)}")
            
            # Преобразуем в число
            number = clean_number(raw_value)
            print(f"Результат: {number}")
            break
