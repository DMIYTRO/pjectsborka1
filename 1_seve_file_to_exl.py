import os
import shutil
import pandas as pd

"""этот скрип меняет раширение файлов в указанной директории"""

def change_extension(directory):
    originals_dir = os.path.join(directory, "originals")
    if not os.path.exists(originals_dir):
        os.makedirs(originals_dir)

    for file in os.listdir(directory):
        if file.endswith(".xls"):
            old_path = os.path.join(directory, file)
            new_path = os.path.join(directory, os.path.splitext(file)[0] + ".html")

            # Создаем копию исходного файла в папке "originals"
            original_path = os.path.join(originals_dir, file)
            shutil.copy2(old_path, original_path)

            # Переименовываем файл в текущей директории
            os.rename(old_path, new_path)
            print(f"Файл {file} переименован в {os.path.basename(new_path)}.")

# Используем указанную директорию
directory_path = r"C:\Users\Dmitry\Downloads\сборки_xls"

change_extension(directory_path)

"""этот скрипт конвертирует файлы из html в xls"""

def convert_to_xls(directory):
    originals_dir = os.path.join(directory, "originals")
    if not os.path.exists(originals_dir):
        os.makedirs(originals_dir)

    for file in os.listdir(directory):
        if file.endswith(".html"):
            html_path = os.path.join(directory, file)
            xls_path = os.path.join(directory, os.path.splitext(file)[0] + ".xlsx")

            # Чтение HTML файла и сохранение в XLS
            df = pd.read_html(html_path)[0]
            with pd.ExcelWriter(xls_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)

            # Переносим исходный HTML файл в папку "originals"
            original_path = os.path.join(originals_dir, file)
            shutil.move(html_path, original_path)

            print(f"Файл {file} сохранен в формате XLS, исходный файл перемещен в папку originals.")

# Используем указанную директорию
directory_path = r"C:\Users\Dmitry\Downloads\сборки_xls"

change_extension(directory_path)
convert_to_xls(directory_path)