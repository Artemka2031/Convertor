import sys
from pathlib import Path
from xml_to_excel import xml_to_excel
from excel_to_xml import excel_to_xml

# Определяем корневую директорию проекта
if getattr(sys, 'frozen', False):
    # Если скрипт запущен как исполняемый файл (например, через PyInstaller)
    project_dir = Path(sys.executable).parent
else:
    # Если скрипт запущен как обычный Python-файл
    project_dir = Path(__file__).parent

def main():
    # Пути к папкам относительно корневой директории проекта
    incoming_dir = project_dir / 'incoming_data'
    processed_dir = project_dir / 'processed_data'

    # Проверяем существование папок, если их нет — создаем
    if not incoming_dir.exists():
        print(f"Папка {incoming_dir} не найдена. Создаю...")
        incoming_dir.mkdir(parents=True, exist_ok=True)
        print(f"Папка {incoming_dir} создана. Пожалуйста, поместите файлы для обработки в эту папку.")

    if not processed_dir.exists():
        print(f"Папка {processed_dir} не найдена. Создаю...")
        processed_dir.mkdir(parents=True, exist_ok=True)
        print(f"Папка {processed_dir} создана.")

    while True:
        # Приветствие и выбор действия
        print("\nДобро пожаловать в приложение для обработки данных!")
        print("Выберите действие:")
        print("1. Сформировать Excel из XML")
        print("2. Сформировать XML из Excel")
        print("3. Завершить программу")
        choice = input("Введите 1, 2 или 3: ").strip()

        # Проверяем, хочет ли пользователь выйти
        if choice == '3':
            print("Программа завершена.")
            break

        # Определяем тип файла и функцию обработки
        if choice == '1':
            file_type = 'xml'
            process_func = xml_to_excel
            output_ext = 'xlsx'
        elif choice == '2':
            file_type = 'xlsx'
            process_func = excel_to_xml
            output_ext = 'xml'
        else:
            print("Неверный выбор. Пожалуйста, введите 1, 2 или 3.")
            continue

        # Получаем список файлов нужного формата
        files = list(incoming_dir.glob(f'*.{file_type}'))
        if not files:
            print(f"Нет файлов формата .{file_type} в папке {incoming_dir}.")
            print("Пожалуйста, поместите файлы в папку incoming_data и попробуйте снова.")
            continue

        # Показываем доступные файлы
        print(f"\nДоступные файлы ({file_type}):")
        for i, file in enumerate(files, 1):
            print(f"{i}. {file.name}")

        # Выбор файла
        try:
            file_choice = int(input("\nВыберите номер файла: ")) - 1
            if file_choice < 0 or file_choice >= len(files):
                print("Неверный номер файла. Пожалуйста, выберите номер из списка.")
                continue
        except ValueError as e:
            print(f"Ошибка: Введите корректный номер файла (целое число). Подробности: {e}")
            continue

        # Определяем входной и выходной файлы
        input_file = files[file_choice]
        base_name = input_file.stem
        output_file = processed_dir / f"{base_name}.{output_ext}"

        # Проверка на существование выходного файла
        if output_file.exists():
            overwrite = input(f"Файл {output_file.name} уже существует в папке {processed_dir}. Перезаписать? (y/n): ").strip().lower()
            if overwrite != 'y':
                print("Обработка отменена.")
                continue

        # Обработка файла
        try:
            process_func(str(input_file), str(output_file))
        except Exception as e:
            print(f"Произошла ошибка при обработке файла: {e}")

        # После обработки или ошибки предлагаем выбрать действие снова
        print("\nОбработка завершена (успешно или с ошибкой). Хотите продолжить?")

if __name__ == "__main__":
    main()