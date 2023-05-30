import pandas as pd
import glob
import re

# Путь к папке с файлами
folder_path = r'C:\Users\Dmitry\Downloads\сборки_xls'

# Получаем список файлов с расширением .xlsx в папке
file_list = glob.glob(folder_path + '\\*.xlsx')

# Проходимся по каждому файлу
for file_path in file_list:
    # Загружаем файл в DataFrame
    df = pd.read_excel(file_path)

    # Получаем последнюю строку столбца 'заказ на'
    last_row_value = df['заказ на'].iloc[-1]

    # Преобразуем last_row_value в строку
    string = str(last_row_value)

    if "(offset)" in string:
        matches = re.findall(r"_([\d]+(?:mat)?)\(offset\)?_", string)
    else:
        matches = re.findall(r"_([\d]+(?:mat)?)_\(", string)

    # Преобразуем строковые значения в числа
    numbers = [int(match.rstrip("mat")) for match in matches]

    # Записываем значения в столбец "Плотность" той же строки
    df.loc[df.index[-1], 'Плотность'] = ', '.join(map(str, numbers))

    # Преобразуем тип данных столбца "Плотность" в числовой формат
    df['Плотность'] = pd.to_numeric(df['Плотность'], errors='coerce')

    # Сохраняем изменения в файле
    df.to_excel(file_path, index=False)

    # Продолжайте обработку данных по своим потребностям
