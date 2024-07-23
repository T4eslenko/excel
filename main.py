import os
import pandas as pd

# Путь к директории с файлами
directory = '.'

# Инициализируем пустой DataFrame для объединения данных
combined_df = pd.DataFrame()

# Проходим по всем файлам в директории
for filename in os.listdir(directory):
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        # Читаем данные из текущего файла Excel
        file_path = os.path.join(directory, filename)
        df = pd.read_excel(file_path)
        
        # Добавляем данные в общий DataFrame
        combined_df = pd.concat([combined_df, df], ignore_index=True)

# Сохраняем общий DataFrame в новый файл Excel
output_file = 'combined_data.xlsx'
combined_df.to_excel(output_file, index=False)

print(f'Данные успешно объединены и сохранены в {output_file}')
