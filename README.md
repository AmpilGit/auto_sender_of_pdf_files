## 📄 README.md для проекта: **PDF Processor**

Это приложение на Python, использующее библиотеку `tkinter`, позволяет автоматизировать процесс:
- **Извлечения текста из PDF-файлов**
- **Поиска совпадений с данными из Excel-таблицы (`patterns.xlsx`)**
- **Автоматической рассылки писем по электронной почте**
- **Создания лог-файла с результатами обработки**

Приложение предназначено для работы с документами в формате PDF и отправки их по указанным адресам электронной почты.

---

### 🛠️ Технологии

| Компонент | Использование |
|-----------|---------------|
| **Python** | 3.10+ |
| **tkinter** | GUI-интерфейс |
| **pdf2image** | Конвертация PDF в изображения для OCR |
| **pytesseract** | Распознавание текста на изображении (OCR) |
| **openpyxl** | Чтение данных из Excel-файла (`patterns.xlsx`) |
| **reportlab** | Создание PDF-отчётов (не используется в текущей версии) |
| **smtplib** | Отправка электронных писем |
| **imaplib** | Не используется, но доступен для расширения |


### 🧰 Как запустить

#### 1. Установите зависимости:

```bash
pip install pytesseract pdf2image openpyxl reportlab python-dotenv
```

> ⚠️ Убедитесь, что установлен **Tesseract OCR** и добавлен в переменные окружения.  
> Пример пути: `C:\Program Files (x86)\tesseract.exe`

#### 2. Настройте файл `patterns.xlsx`:

Формат файла:
| Номер счета | Email менеджера |
|-------------|----------------|
| 12345       | manager@example.com |
| 67890       | another.manager@example.com |

> Важно: этот файл должен находиться в **папке вывода**, которую вы выберете через интерфейс.

---

### 🔍 Работа программы

1. Выберите папку с PDF-файлами.
2. Выберите папку вывода (где будет сохранён `result.txt`).
3. Запустите обработку.
4. Приложение:
   - Проходит по всем PDF-файлам в выбранной папке
   - Извлекает текст с первой страницы
   - Проверяет наличие номеров счетов из `patterns.xlsx`
   - Если найдены — отправляет PDF-файл по указанному email
   - Если не найдены — записывает в `result.txt`
   - Также отправляет копию письма модератору

---

### 🧾 Функционал

| Функция | Описание |
|---------|----------|
| `extract_text_from_pdf` | Извлекает текст из PDF-документа с помощью OCR |
| `load_data_from_excel` | Читает данные из `patterns.xlsx` |
| `combine_pdfs_to_one` | Обрабатывает все PDF-файлы в папке и отправляет письма |
| `select_folder` | Выбор папки с PDF-файлами |
| `select_result_folder` | Выбор папки для сохранения результата и `patterns.xlsx` |

---

### 📝 Требования

- **Tesseract OCR** должен быть установлен и добавлен в системные переменные.
- **PDF-файлы должны содержать текст на первой странице**.
- **Excel-файл `patterns.xlsx` должен содержать две колонки**: `Номер счета`, `Email менеджера`.
- **Доступ к SMTP-серверу mail.ru** (настроены в коде).

---

### 🧩 Настройка SMTP

В коде уже заданы параметры для отправки писем через `smtp.mail.ru`.  
Если нужно изменить — отредактируйте:

> Убедитесь, что у вас есть права на отправку писем с этого аккаунта.

---

### 📦 Дополнительно

- Приложение может быть расширено для работы с несколькими страницами.
- Можно добавить поддержку других форматов документов (например, `.docx`, `.jpg`, `.png`).
- Возможна интеграция с базой данных или API.

---

### 🧩 Пример использования

1. Откройте приложение.
2. Выберите папку с PDF-файлами.
3. Выберите папку вывода.
4. Нажмите "Выбрать папку".
5. Приложение начнёт обработку и отправку писем.

---

### 🧪 Тестирование

- Проверьте, как работает поиск номеров счетов.
- Убедитесь, что письма отправляются корректно.
- Проверьте лог-файл `result.txt`.

---

### 📌 Автор

**Кудрявцев Данил**  
Системный администратор IT-отдела  
Email: mikushkinodanil4@gmail.com

---

### 📝 Лицензия

MIT License

Copyright (c) 2025 Кудрявцев Данил

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

