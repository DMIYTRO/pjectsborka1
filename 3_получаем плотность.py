import pandas as pd
import glob

# Путь к папке с файлами
folder_path = r'C:\Users\Dmitry\Downloads\сборки_xls'

# Получаем список файлов с расширением .xlsx в папке
file_list = glob.glob(folder_path + '\\*.xlsx')

# Проходимся по каждому файлу
for file_path in file_list:
    # Загружаем файл в DataFrame
    df = pd.read_excel(file_path)

    # Извлекаем числа из столбца 'имя' с помощью регулярного выражения
    df['Плотность'] = df['имя'].str.extract(r"_(\d+)_")

    # Преобразуем столбец 'Плотность' в числовой формат
    df['Плотность'] = pd.to_numeric(df['Плотность'])

    # Определяем индекс столбца 'место'
    column_index = df.columns.get_loc('мест')

    # # Перемещаем столбец 'Плотность' после столбца 'место'
    # columns = df.columns.tolist()
    # columns.insert(column_index, 'Плотность')
    # df = df.reindex(columns=columns)

    # Записываем результаты обратно в исходный файл
    df.to_excel(file_path, index=False)

# Выводим сообщение о завершении
print("Данные успешно сохранены в файлы.")
