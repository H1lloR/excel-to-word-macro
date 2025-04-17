Sub СоздатьДокументСоСпискомФирм()
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim TemplateDoc As Object
    Dim ws As Worksheet
    Dim firmName As String
    Dim i As Long
    Dim templatePath As String
    Dim savePath As String
    Dim pasteRange As Object

    ' Пути к файлам
    templatePath = Environ("USERPROFILE") & "\Desktop\Шаблон.docx"
    savePath = Environ("USERPROFILE") & "\Desktop\Итоговый_документ.docx"

    ' Проверка существования шаблона
    If Dir(templatePath) = "" Then
        MsgBox "Файл шаблона не найден на рабочем столе: Шаблон.docx", vbCritical
        Exit Sub
    End If

    ' Получаем рабочий лист
    Set ws = ThisWorkbook.Sheets(1)

    ' Запускаем Word
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = False

    ' Открываем шаблон один раз
    Set TemplateDoc = WordApp.Documents.Open(templatePath, ReadOnly:=True)

    ' Копируем содержимое шаблона
    TemplateDoc.Content.Copy
    TemplateDoc.Close SaveChanges:=False

    ' Создаём итоговый документ
    Set WordDoc = WordApp.Documents.Add

    ' Цикл по названиям фирм
    For i = 4 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        firmName = Trim(ws.Cells(i, 2).Value)

        If Len(firmName) > 0 Then
            ' Вставляем копию шаблона
            Set pasteRange = WordDoc.Range
            pasteRange.Collapse Direction:=0 ' В конец документа
            pasteRange.PasteAndFormat Type:=16 ' wdFormatOriginalFormatting

            ' Заменяем плейсхолдер на имя фирмы только в вставленном участке
            With pasteRange.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "{ФИРМА}"
                .Replacement.Text = firmName
                .Wrap = 0
                .Execute Replace:=2 ' wdReplaceAll
            End With

            ' Добавим разрыв параграфа
            ' pasteRange.InsertParagraphAfter опциально для каждой задачи
        End If
    Next i

    ' Сохраняем результат
    WordDoc.SaveAs2 savePath
    WordDoc.Close SaveChanges:=False
    WordApp.Quit

    MsgBox "Готово! Итоговый документ сохранён на рабочем столе."

    ' Очистка
    Set WordDoc = Nothing
    Set WordApp = Nothing
    Set TemplateDoc = Nothing
End Sub
