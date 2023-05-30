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
    matches = re.findall(r"\d+x\d+_(\d+)", string)

    # Записываем значения в столбец "тираж" той же строки
    df.loc[df.index[-1], 'тираж'] = ', '.join(map(str, matches))

    # Преобразуем тип данных столбца "тираж" в числовой формат
    df['тираж'] = pd.to_numeric(df['тираж'], errors='coerce')

    # Сохраняем изменения в файле
    df.to_excel(file_path, index=False)

