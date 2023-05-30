import subprocess
import sys
script_paths = ['1_seve_file_to_exl.py', '2_промежуточны_формат_xlsx.py', '3_получаем плотность.py',
                '4_сохранили_плотность.py', '5_получаем тираж.py', '6_готовый_файл.py', '7_все объеденить.py']

# for script_path in script_paths:
#     try:
#         subprocess.run(['python', script_path], check=True)
#     except subprocess.CalledProcessError:
#         print(f'Ошибка при выполнении скрипта: {script_path}')

for script_path in script_paths:
    subprocess.run([sys.executable, script_path], check=True)