import pandas as pd
import glob

folder_path = r'C:\Users\Dmitry\Downloads\сборки_xls'
file_list = glob.glob(folder_path + '\\*.xlsx')

for file_path in file_list:
    # Загрузка файла в DataFrame
    df = pd.read_excel(file_path)
    # Получение значения последней ячейки столбца "тираж"
    last_value = df['тираж'].iloc[-1]

    # Получение значения последней ячейки столбца "Плотность"
    last_density = df['Плотность'].iloc[-1]
    # Создание объекта ExcelWriter с использованием движка записи xlsxwriter
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    # Запись DataFrame в Excel
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    # Получение объекта Workbook и Worksheet
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    # Создание объектов форматирования для заливки ячеек и жирного шрифта
    green_bold_format = workbook.add_format({'bg_color': '#3CB371', 'bold': True})
    red_bold_format = workbook.add_format({'bg_color': '#4682B4', 'bold': True})
    green_format = workbook.add_format({'bg_color': '#6495ED', 'bold': True})
    last_row_format = workbook.add_format({'bg_color': '#808080'})
    # Применение форматирования к строкам, где значение тиража или Плотность соответствует условию
    for row in range(1, len(df) + 1):
        cell_value = df.loc[row - 1, 'тираж']
        density_value = df.loc[row - 1, 'Плотность']
        if cell_value < last_value:
            worksheet.set_row(row, cell_format=green_bold_format)
        if density_value < last_density:
            worksheet.set_row(row, cell_format=green_format)
        if density_value > last_density:
            worksheet.set_row(row, cell_format=red_bold_format)
        last_paid_value = df['тираж'].iloc[-1]
# Заливка ячеек цветом 4b9d7a, если значение в столбце "мест" больше значения в столбце "тираж"
        for row in range(len(df) - 1):  # Исключаем последнюю строку
            cell_value = df.loc[row, 'мест'] * last_paid_value
            if cell_value > df.loc[row, 'тираж']:
                worksheet.set_row(row + 1, cell_format=green_format)
    # Залитие последних двух строк цветом a6a8a4
    last_row_index = len(df) - 1
    for row in range(last_row_index, last_row_index + 2):
        worksheet.set_row(row, cell_format=last_row_format)
    # Auto-adjust columns' width
    for column in df:
        column_width = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        worksheet.set_column(col_idx, col_idx, column_width)
    # Закрытие файла
    writer.close()
