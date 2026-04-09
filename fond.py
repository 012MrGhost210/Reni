import re
import xlrd

def clean_number(value):
    """Преобразует '1 784 222,90 руб.' в число 1784222.90"""
    if value is None:
        return None
    
    # Если уже число
    if isinstance(value, (int, float)):
        return float(value)
    
    # Преобразуем в строку
    value_str = str(value).strip()
    
    # Удаляем всё после "руб"
    if 'руб' in value_str.lower():
        value_str = value_str.lower().split('руб')[0].strip()
    
    # Удаляем все пробелы
    value_str = value_str.replace(' ', '')
    
    # Заменяем запятую на точку
    value_str = value_str.replace(',', '.')
    
    # Удаляем всё кроме цифр и точки
    value_str = re.sub(r'[^\d.]', '', value_str)
    
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
            
            # Преобразуем в число
            number = clean_number(raw_value)
            print(f"Число: {number}")
            break
   🔍 Ищу: 'Итого стоимость чистых активов' в любом столбце
   📍 Беру число из столбца: P (индекс 15)
      ✅ Найдено ключевое слово в строке 155, столбец 0
      📍 Значение в столбце 15 (P): '628 487,91 руб.'
      ⚠️ Не удалось преобразовать в число: '628 487,91 руб.'
   ⚠️ Ключевое слово не найдено
