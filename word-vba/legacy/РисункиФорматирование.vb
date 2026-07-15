Sub РисункиФорматирование()
    Dim shp As InlineShape
    Dim captionText As String
    Dim picIndex As Integer
    Dim captionExists As Boolean
    Dim para As Paragraph
    Dim pageWidth As Single
    
    ' Получаем ширину страницы (учитываем поля документа)
    pageWidth = ActiveDocument.PageSetup.PageWidth - ActiveDocument.PageSetup.LeftMargin - ActiveDocument.PageSetup.RightMargin
    picIndex = 1 ' Счетчик рисунков для заголовков

    For Each shp In ActiveDocument.InlineShapes
        ' Центрируем рисунок
        shp.Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        ' Проверяем размер рисунка, если он занимает более 40% ширины страницы
        If shp.Width > 0.4 * pageWidth Then
            ' Масштабируем рисунок пропорционально до ширины страницы
            shp.LockAspectRatio = msoTrue
            shp.Width = pageWidth
        End If

        ' Проверяем, существует ли уже описание (подпись) для рисунка
        captionExists = False
        For Each para In ActiveDocument.Paragraphs
            If para.Style = wdStyleCaption And para.Range.Text Like "Рисунок " & picIndex & "*" Then
                captionExists = True
                Exit For
            End If
        Next para

        ' Добавляем подпись, если она еще не существует
        If Not captionExists Then
            captionText = "Описание"
            shp.Select ' Выбираем рисунок, чтобы добавить подпись
            Selection.InsertCaption Label:="Рисунок", Title:=" " & captionText, Position:=wdCaptionPositionBelow
            Selection.Range.NoProofing = True ' Отключаем проверку правописания для подписи
        End If
        
        picIndex = picIndex + 1
    Next shp

    MsgBox "Все рисунки отформатированы, центрированы, масштабированы (если необходимо), и подписаны.", vbInformation
End Sub
