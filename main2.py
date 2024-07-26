import os
import pandas as pd

def merge_excel_files(output_file):
    # Получаем текущую директорию
    current_directory = os.getcwd()
    
    # Создаем пустой DataFrame для объединения всех данных
    combined_df = pd.DataFrame()

    # Перебираем все файлы в текущей директории
    for filename in os.listdir(current_directory):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            file_path = os.path.join(current_directory, filename)
            # Читаем текущий Excel файл
            try:
                df = pd.read_excel(file_path)
                # Проверяем, пустой ли DataFrame
                if not df.empty:
                    # Добавляем данные в общий DataFrame
                    combined_df = pd.concat([combined_df, df], ignore_index=True)
                else:
                    print(f"Warning: {filename} is empty and will be skipped.")
            except Exception as e:
                print(f"Error reading {filename}: {e}")

    # Сохраняем объединенные данные в новый Excel файл
    if not combined_df.empty:
        combined_df.to_excel(output_file, index=False)
        print(f"Data successfully combined and saved to {output_file}")
    else:
        print("No data found to combine.")

# Указываем имя выходного файла
output_file = "combined_output.xlsx"

# Запускаем функцию объединения файлов
merge_excel_files(output_file)
