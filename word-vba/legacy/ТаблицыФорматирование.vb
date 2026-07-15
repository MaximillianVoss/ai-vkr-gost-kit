Sub ТаблицыФорматирование()
    Dim tbl As Table
    Dim doc As Document
    Dim captionText As String
    Dim tableIndex As Integer
    Dim para As Paragraph
    Dim captionExists As Boolean
    Set doc = ActiveDocument
    tableIndex = 1 ' Счетчик таблиц для заголовков

    For Each tbl In doc.Tables
        ' Устанавливаем ширину таблицы на ширину страницы
        tbl.PreferredWidthType = wdPreferredWidthPercent
        tbl.PreferredWidth = 100 ' 100% ширины страницы
        
        ' Применяем стиль, если он существует и не установлен
        On Error Resume Next
        If tbl.Style <> "ГОСТ.Черный" Then
            tbl.Style = "ГОСТ.Черный"
        End If
        On Error GoTo 0
        
        ' Повторяем строки заголовков на новой странице
        tbl.Rows(1).HeadingFormat = True
        
        ' Отключаем проверку правописания для содержимого таблицы
        tbl.Range.NoProofing = True
        
        ' Проверяем, существует ли уже описание (подпись) для таблицы
        captionExists = False
        For Each para In doc.Paragraphs
            If para.Style = wdStyleCaption And para.Range.Text Like "Таблица " & tableIndex & "*" Then
                captionExists = True
                Exit For
            End If
        Next para
        
        ' Добавляем подпись, если она еще не существует
        If Not captionExists Then
            ' captionText = "Таблица " & tableIndex & " Описание"
            captionText = " Описание"
            tbl.Select ' Выбираем таблицу, чтобы добавить подпись
            Selection.InsertCaption Label:="Таблица", Title:=captionText, Position:=wdCaptionPositionAbove
            Selection.Range.NoProofing = True ' Отключаем проверку правописания для подписи
        End If

        
        tableIndex = tableIndex + 1
    Next tbl
    
    MsgBox "Все таблицы отформатированы, заголовки добавлены (если необходимо), и проверка правописания отключена.", vbInformation
End Sub




Sub ConvertDoc()
'
' latex2word.ConvertDoc Макрос
'
'

End Sub
