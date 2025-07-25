# UTD Converter

Конвертер файлов между форматами XML и Excel.

## Требования

- Windows 7 или выше.
- Установленный Python 3.8 или выше (рекомендуется Python 3.10).
- Доступ в интернет для загрузки Bootstrap (для интерфейса).

## Установка и запуск

1. **Распакуйте архив**:Распакуйте архив `Convertor.zip` в любую папку на вашем компьютере.

2. **Проверьте наличие Python**:Убедитесь, что Python установлен. Откройте командную строку (cmd) и введите:

   ```
   python --version
   ```

   Если Python не установлен, скачайте и установите его с официального сайта: https://www.python.org/downloads/

3. **Запустите приложение:**

   - Перейдите в папку с распакованным проектом.
   - Дважды щелкните по файлу `run_gui.bat`.
   - Если файл не запускается, откройте командную строку (cmd), перейдите в папку проекта и выполните:

     ```
     .\run_gui.bat
     ```

4. **Используйте приложение:**

   - После запуска в командной строке появится сообщение о запуске сервера.
   - Откройте браузер и перейдите по адресу: http://127.0.0.1:8000
   - Вы увидите интерфейс конвертера. Выберите файл (XML или Excel), дождитесь конвертации и скачайте результат.

## Устранение неполадок

- Если приложение не запускается, убедитесь, что все зависимости установлены:

  ```
  .\venv\Scripts\activate
  pip install -r requirements.txt
  ```
- Если возникают ошибки с интернетом, убедитесь, что у вас есть доступ в сеть (интерфейс загружает Bootstrap через CDN).

## Поддерживаемые форматы

- Входные файлы: `.xlsx` (Excel), `.xml`.
- Выходные файлы: `.xml` (если входной файл Excel), `.xlsx` (если входной файл XML).