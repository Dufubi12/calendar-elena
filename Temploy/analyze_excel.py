import pandas as pd
import json
import os

file_path = 'Посещаемость кружки и ИМ 2025_2026.xlsx'

if not os.path.exists(file_path):
    print(f"Ошибка: Файл '{file_path}' не найден.")
    exit(1)

try:
    # Read the Excel file
    xls = pd.ExcelFile(file_path)
    
    print(f"Загружен файл: {file_path}")
    print(f"Доступные листы: {xls.sheet_names}")

    # Analyze 'Елена ИМ'
    if 'Елена ИМ' in xls.sheet_names:
        df_elena = pd.read_excel(xls, sheet_name='Елена ИМ')
        print("\n--- Анализ листа 'Елена ИМ' ---")
        print(f"Количество строк: {len(df_elena)}")
        print(f"Колонки: {list(df_elena.columns)}")
        print("Первые 5 строк данных:")
        print(df_elena.head().to_string())
        
        # Convert to JSON for potential usage
        elena_json = df_elena.head().to_json(orient='records', force_ascii=False)
        print("\nПример данных в JSON (для сайта):")
        print(elena_json)
    else:
        print("\nЛист 'Елена ИМ' не найден!")

    # Analyze 'Посещаемость'
    if 'Посещаемость' in xls.sheet_names:
        df_attendance = pd.read_excel(xls, sheet_name='Посещаемость')
        print("\n--- Анализ листа 'Посещаемость' ---")
        print(f"Количество строк: {len(df_attendance)}")
        print(f"Колонки: {list(df_attendance.columns)}")
        print("Первые 5 строк данных:")
        print(df_attendance.head().to_string())
    else:
        print("\nЛист 'Посещаемость' не найден!")

except Exception as e:
    print(f"Произошла ошибка при чтении файла: {e}")
    print("Убедитесь, что установлены необходимые библиотеки: pip install pandas openpyxl")
