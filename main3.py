import os
import pandas as pd
from datetime import datetime

def merge_excel_files():
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
                    file_size = os.path.getsize(file_path)
                    print(f"Warning: {filename} is empty (size: {file_size} bytes) and will be skipped.")
            except Exception as e:
                print(f"Error reading {filename}: {e}")

    # Добавляем текущую дату и время к имени выходного файла
    current_datetime = datetime.now().strftime("%d.%m.%Y_%H-%M")
    output_file = f"combined_output_{current_datetime}.xlsx"

    # Подсчитываем количество строк и непустых строк в столбце "Телефон контакта"
    total_rows = combined_df.shape[0]
    non_empty_phone_rows = combined_df['Телефон контакта'].dropna().shape[0]

    # Сохраняем объединенные данные в новый Excel файл
    if not combined_df.empty:
        combined_df.to_excel(output_file, index=False)
        print(f"Data successfully combined and saved to {output_file}")
        print(f"Total rows in 'Телефон контакта': {total_rows}")
        print(f"Non-empty rows in 'Телефон контакта': {non_empty_phone_rows}")
    else:
        print("No data found to combine.")

# Запускаем функцию объединения файлов
merge_excel_files()
