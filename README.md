# Excel-to-Word Merge Macro Rus

Этот макрос позволяет автоматически создавать Word-документ с подставленными названиями фирм из Excel.

## 📌 Что делает скрипт?

- Берёт список фирм из Excel, начиная с 4 строки, 2 столбца.
- Загружает шаблон Word (`Шаблон.docx`) с плейсхолдером `{ФИРМА}`.
- Вставляет шаблон в итоговый документ столько раз, сколько фирм.
- Заменяет плейсхолдер на конкретное название фирмы.
- Сохраняет результат как `Итоговый_документ.docx` на рабочем столе.

## 📂 Что нужно?

- Excel файл с названиями фирм в 4 строке и 2 столбце (пример — `Пример.xlsx`)
- Шаблон Word с текстом `{ФИРМА}` внутри (пример — `Шаблон.docx`)
- Установленный Microsoft Word и Excel

## 🚀 Как пользоваться?

1. Скачай и открой Excel-файл с названиями фирм.
2. Нажми `Alt + F11` чтобы открыть редактор VBA.
3. Вставь код из файла `macro.bas` в модуль.
4. Запусти макрос: `СоздатьДокументСоСпискомФирм`
5. Готово! Итог будет на рабочем столе.

## 💡 Примечания

- Удалённый разрыв параграфа в конце блока (можно добавить вручную при необходимости).
- Если хочешь разделять фирмы по страницам — легко добавить `InsertBreak`.

## 🛠️ Автор

Создан для автоматизации однотипной офисной работы.  
С радостью приму предложения и идеи для доработки!

# Excel-to-Word Merge Macro Eng

This macro allows you to automatically create a Word document with substituted business names from Excel.

## 📌 What does the script do?

- Takes a list of firms from Excel, starting at row 4, 2 columns.
- Loads a Word template (`Template.docx`) with placeholder `{FIRM}`.
- Inserts the template into the final document as many times as there are firms.
- Replaces the placeholder with a specific firm name.
- Saves the result as ``Total_document.docx`` on the desktop.

## 📂 What is needed?

- Excel file with firm names in row 4 and column 2 (example - `Example.xlsx`)
- A Word template with the text `{FIRM}` inside (example - `Template.docx`)
- Microsoft Word and Excel installed

## 🚀 How to use?

1. Download and open the Excel file with the firm names.
2. Press `Alt + F11` to open the VBA editor.
3. Paste the code from the `macro.bas` file into the module.
4. Run the macro: `CreateDocumentCompanyList`.
5. Done! The output will be on your desktop.

## 💡 Notes

- Removed paragraph break at the end of the block (can be added manually if needed).
- If you want to separate firms by page, it's easy to add `InsertBreak`.

## 🛠️ Author

Created to automate the same type of office work.  
Will gladly accept suggestions and ideas for improvements!
