Option Explicit

Private Const DEFAULT_MAX_FIGURE_WIDTH_PERCENT As Double = 100#
Private Const CAPTION_LOOKUP_DEPTH As Long = 3

Public Sub РисункиФорматирование()
    Dim alignment As WdParagraphAlignment
    Dim maxWidthPercent As Double
    Dim includeChapterNumber As Boolean
    Dim inlinePicture As InlineShape
    Dim pageWidth As Single
    Dim maxWidth As Single
    Dim formattedCount As Long
    Dim insertedCount As Long

    alignment = PromptFigureAlignment()
    maxWidthPercent = PromptMaxFigureWidthPercent(DEFAULT_MAX_FIGURE_WIDTH_PERCENT)
    includeChapterNumber = PromptCaptionNumberingMode("Нумерация рисунков")
    ConfigureCaptionLabel "Рисунок", includeChapterNumber

    For Each inlinePicture In ActiveDocument.InlineShapes
        inlinePicture.Range.ParagraphFormat.Alignment = alignment
        inlinePicture.Range.ParagraphFormat.LeftIndent = 0
        inlinePicture.Range.ParagraphFormat.RightIndent = 0
        inlinePicture.Range.ParagraphFormat.FirstLineIndent = 0

        pageWidth = UsablePageWidth(inlinePicture.Range)
        maxWidth = pageWidth * CSng(maxWidthPercent / 100#)
        If inlinePicture.Width > maxWidth Then
            inlinePicture.LockAspectRatio = msoTrue
            inlinePicture.Width = maxWidth
        End If

        If Not HasAdjacentCaption(inlinePicture.Range, "Рисунок", False) Then
            inlinePicture.Range.Select
            Selection.InsertCaption Label:="Рисунок", Title:=" Описание", Position:=wdCaptionPositionBelow
            Selection.Range.NoProofing = True
            insertedCount = insertedCount + 1
        End If
        formattedCount = formattedCount + 1
    Next inlinePicture

    ActiveDocument.Fields.Update
    MsgBox "Рисунки отформатированы: " & formattedCount & ". Добавлено подписей: " & insertedCount & ".", vbInformation
End Sub

Public Sub ТаблицыФорматирование()
    Dim includeChapterNumber As Boolean
    Dim tableItem As Table
    Dim formattedCount As Long
    Dim insertedCount As Long

    includeChapterNumber = PromptCaptionNumberingMode("Нумерация таблиц")
    ConfigureCaptionLabel "Таблица", includeChapterNumber

    For Each tableItem In ActiveDocument.Tables
        tableItem.PreferredWidthType = wdPreferredWidthPercent
        tableItem.PreferredWidth = 100
        tableItem.Rows.Alignment = wdAlignRowCenter
        tableItem.Range.NoProofing = True

        If tableItem.Rows.Count > 0 Then
            tableItem.Rows(1).HeadingFormat = True
        End If

        On Error Resume Next
        tableItem.Style = "ГОСТ.Черный"
        On Error GoTo 0

        If Not HasAdjacentCaption(tableItem.Range, "Таблица", True) Then
            tableItem.Range.Select
            Selection.InsertCaption Label:="Таблица", Title:=" Описание", Position:=wdCaptionPositionAbove
            Selection.Range.NoProofing = True
            insertedCount = insertedCount + 1
        End If
        formattedCount = formattedCount + 1
    Next tableItem

    ActiveDocument.Fields.Update
    MsgBox "Таблицы отформатированы: " & formattedCount & ". Добавлено подписей: " & insertedCount & ".", vbInformation
End Sub

Public Sub ОбновитьНумерациюЗаголовков()
    Dim addChapterPrefix As Boolean
    Dim caseMode As Long
    Dim paragraphItem As Paragraph
    Dim headingText As String
    Dim normalizedText As String
    Dim chapterIndex As Long
    Dim changedCount As Long
    Dim contentRange As Range

    addChapterPrefix = PromptChapterPrefixMode()
    caseMode = PromptHeadingCaseMode()
    chapterIndex = 1

    For Each paragraphItem In ActiveDocument.Paragraphs
        If IsTopLevelHeading(paragraphItem) Then
            headingText = CleanParagraphText(paragraphItem.Range.Text)
            headingText = StripChapterPrefix(headingText)
            normalizedText = NormalizeText(headingText)

            If Len(headingText) > 0 And Not IsSpecialHeading(normalizedText) Then
                headingText = ApplyHeadingCase(headingText, caseMode)
                If addChapterPrefix Then
                    headingText = "Глава " & chapterIndex & ". " & headingText
                End If

                Set contentRange = paragraphItem.Range.Duplicate
                contentRange.End = contentRange.End - 1
                contentRange.Text = headingText
                paragraphItem.Range.ListFormat.RemoveNumbers
                changedCount = changedCount + 1
                chapterIndex = chapterIndex + 1
            End If
        End If
    Next paragraphItem

    ActiveDocument.Fields.Update
    MsgBox "Обновлено заголовков верхнего уровня: " & changedCount & ".", vbInformation
End Sub

Public Sub AlignTablesAndFigures()
    ТаблицыФорматирование
    РисункиФорматирование
End Sub

Public Sub UpdateGostHeadingNumbering()
    ОбновитьНумерациюЗаголовков
End Sub

Private Function PromptFigureAlignment() As WdParagraphAlignment
    Dim answer As String
    answer = InputBox("Как выровнять рисунки?" & vbCrLf & _
        "1 - слева" & vbCrLf & _
        "2 - по центру" & vbCrLf & _
        "3 - справа", "Выравнивание рисунков", "2")

    Select Case Trim$(answer)
        Case "1"
            PromptFigureAlignment = wdAlignParagraphLeft
        Case "3"
            PromptFigureAlignment = wdAlignParagraphRight
        Case Else
            PromptFigureAlignment = wdAlignParagraphCenter
    End Select
End Function

Private Function PromptMaxFigureWidthPercent(ByVal defaultPercent As Double) As Double
    Dim answer As String
    Dim result As Double

    answer = InputBox("Максимальная ширина рисунка, % от полезной ширины страницы. Рисунки не будут растягиваться.", _
        "Размер рисунков", CStr(defaultPercent))
    If Len(Trim$(answer)) = 0 Or Not IsNumeric(answer) Then
        result = defaultPercent
    Else
        result = CDbl(answer)
    End If

    If result < 10# Then result = 10#
    If result > 100# Then result = 100#
    PromptMaxFigureWidthPercent = result
End Function

Private Function PromptCaptionNumberingMode(ByVal title As String) As Boolean
    Dim answer As VbMsgBoxResult
    answer = MsgBox(title & vbCrLf & vbCrLf & _
        "Да - сквозная нумерация по всему документу." & vbCrLf & _
        "Нет - нумерация по главам, если заголовки Word это поддерживают.", _
        vbYesNoCancel + vbQuestion, title)

    If answer = vbCancel Then Err.Raise vbObjectError + 513, , "Операция отменена пользователем."
    PromptCaptionNumberingMode = (answer = vbNo)
End Function

Private Function PromptChapterPrefixMode() As Boolean
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Добавить к заголовкам верхнего уровня префикс вида 'Глава N.'?", _
        vbYesNoCancel + vbQuestion, "Нумерация заголовков")

    If answer = vbCancel Then Err.Raise vbObjectError + 514, , "Операция отменена пользователем."
    PromptChapterPrefixMode = (answer = vbYes)
End Function

Private Function PromptHeadingCaseMode() As Long
    Dim answer As String
    answer = InputBox("Как изменить регистр текста заголовков верхнего уровня?" & vbCrLf & _
        "0 - не менять" & vbCrLf & _
        "1 - ПРОПИСНЫЕ" & vbCrLf & _
        "2 - строчные" & vbCrLf & _
        "3 - Заглавная первая буква" & vbCrLf & _
        "4 - Каждое Слово С Заглавной", "Регистр заголовков", "0")

    If Not IsNumeric(answer) Then
        PromptHeadingCaseMode = 0
        Exit Function
    End If

    PromptHeadingCaseMode = CLng(answer)
    If PromptHeadingCaseMode < 0 Or PromptHeadingCaseMode > 4 Then PromptHeadingCaseMode = 0
End Function

Private Sub ConfigureCaptionLabel(ByVal labelName As String, ByVal includeChapterNumber As Boolean)
    On Error Resume Next
    Application.CaptionLabels.Add Name:=labelName
    On Error GoTo 0

    On Error Resume Next
    With Application.CaptionLabels(labelName)
        .IncludeChapterNumber = includeChapterNumber
        If includeChapterNumber Then
            .ChapterStyleLevel = 1
            .Separator = wdSeparatorPeriod
        End If
    End With
    On Error GoTo 0
End Sub

Private Function HasAdjacentCaption(ByVal sourceRange As Range, ByVal labelName As String, ByVal lookAbove As Boolean) As Boolean
    Dim paragraphItem As Paragraph
    Set paragraphItem = FindAdjacentNonEmptyParagraph(sourceRange, lookAbove, CAPTION_LOOKUP_DEPTH)
    If paragraphItem Is Nothing Then Exit Function
    HasAdjacentCaption = ParagraphMatchesCaption(paragraphItem, labelName)
End Function

Private Function FindAdjacentNonEmptyParagraph(ByVal sourceRange As Range, ByVal lookAbove As Boolean, ByVal maxDepth As Long) As Paragraph
    Dim currentRange As Range
    Dim nextRange As Range
    Dim paragraphItem As Paragraph
    Dim index As Long

    Set currentRange = sourceRange.Duplicate
    For index = 1 To maxDepth
        Set nextRange = Nothing
        On Error Resume Next
        If lookAbove Then
            Set nextRange = currentRange.Previous(wdParagraph, 1)
        Else
            Set nextRange = currentRange.Next(wdParagraph, 1)
        End If
        On Error GoTo 0

        If nextRange Is Nothing Then Exit Function
        Set paragraphItem = nextRange.Paragraphs(1)
        If Len(CleanParagraphText(paragraphItem.Range.Text)) > 0 Then
            Set FindAdjacentNonEmptyParagraph = paragraphItem
            Exit Function
        End If
        Set currentRange = paragraphItem.Range.Duplicate
    Next index
End Function

Private Function ParagraphMatchesCaption(ByVal paragraphItem As Paragraph, ByVal labelName As String) As Boolean
    Dim text As String
    text = CleanParagraphText(paragraphItem.Range.Text)

    If CaptionTextMatchesLabel(text, labelName) Then
        ParagraphMatchesCaption = True
        Exit Function
    End If

    If CaptionTextMatchesAnyLabel(text) Then Exit Function

    On Error Resume Next
    ParagraphMatchesCaption = (paragraphItem.Style = wdStyleCaption)
    On Error GoTo 0
End Function

Private Function CaptionTextMatchesAnyLabel(ByVal text As String) As Boolean
    CaptionTextMatchesAnyLabel = CaptionTextMatchesLabel(text, "Таблица") Or CaptionTextMatchesLabel(text, "Рисунок")
End Function

Private Function CaptionTextMatchesLabel(ByVal text As String, ByVal labelName As String) As Boolean
    Dim normalized As String
    normalized = NormalizeText(text)

    If labelName = "Таблица" Then
        CaptionTextMatchesLabel = StartsWith(normalized, "таблица ") Or StartsWith(normalized, "табл. ")
    ElseIf labelName = "Рисунок" Then
        CaptionTextMatchesLabel = StartsWith(normalized, "рисунок ") Or StartsWith(normalized, "рис. ")
    End If
End Function

Private Function StartsWith(ByVal value As String, ByVal prefix As String) As Boolean
    StartsWith = (Left$(value, Len(prefix)) = prefix)
End Function

Private Function UsablePageWidth(ByVal sourceRange As Range) As Single
    Dim sectionItem As Section

    On Error Resume Next
    Set sectionItem = sourceRange.Sections(1)
    On Error GoTo 0

    If Not sectionItem Is Nothing Then
        UsablePageWidth = sectionItem.PageSetup.PageWidth - sectionItem.PageSetup.LeftMargin - sectionItem.PageSetup.RightMargin
    Else
        UsablePageWidth = ActiveDocument.PageSetup.PageWidth - ActiveDocument.PageSetup.LeftMargin - ActiveDocument.PageSetup.RightMargin
    End If
End Function

Private Function IsTopLevelHeading(ByVal paragraphItem As Paragraph) As Boolean
    Dim styleName As String

    On Error Resume Next
    styleName = LCase$(paragraphItem.Style.NameLocal)
    On Error GoTo 0

    IsTopLevelHeading = (styleName = "гост заголовок 1" Or styleName = "heading 1" Or paragraphItem.OutlineLevel = wdOutlineLevel1)
End Function

Private Function IsSpecialHeading(ByVal normalizedText As String) As Boolean
    Select Case normalizedText
        Case "реферат", "введение", "заключение", "список литературы", "список использованных источников", "список источников"
            IsSpecialHeading = True
        Case Else
            IsSpecialHeading = StartsWith(normalizedText, "приложение")
    End Select
End Function

Private Function StripChapterPrefix(ByVal text As String) As String
    Dim expression As Object
    Set expression = CreateObject("VBScript.RegExp")

    expression.Global = False
    expression.IgnoreCase = True
    expression.Pattern = "^\s*(глава\s+)?\d+([\.\)\-]\s*)?"
    StripChapterPrefix = Trim$(expression.Replace(text, ""))
End Function

Private Function ApplyHeadingCase(ByVal text As String, ByVal mode As Long) As String
    Select Case mode
        Case 1
            ApplyHeadingCase = UCase$(text)
        Case 2
            ApplyHeadingCase = LCase$(text)
        Case 3
            ApplyHeadingCase = CapitalizeFirst(text)
        Case 4
            ApplyHeadingCase = CapitalizeWords(text)
        Case Else
            ApplyHeadingCase = text
    End Select
End Function

Private Function CapitalizeFirst(ByVal text As String) As String
    Dim lowered As String
    lowered = LCase$(text)
    If Len(lowered) = 0 Then
        CapitalizeFirst = lowered
    Else
        CapitalizeFirst = UCase$(Left$(lowered, 1)) & Mid$(lowered, 2)
    End If
End Function

Private Function CapitalizeWords(ByVal text As String) As String
    Dim parts() As String
    Dim index As Long
    Dim item As String

    parts = Split(LCase$(text), " ")
    For index = LBound(parts) To UBound(parts)
        item = parts(index)
        If Len(item) > 0 Then parts(index) = UCase$(Left$(item, 1)) & Mid$(item, 2)
    Next index
    CapitalizeWords = Join(parts, " ")
End Function

Private Function CleanParagraphText(ByVal text As String) As String
    CleanParagraphText = Trim$(Replace(Replace(text, Chr$(13), ""), Chr$(7), ""))
End Function

Private Function NormalizeText(ByVal text As String) As String
    NormalizeText = LCase$(Replace(CleanParagraphText(text), "ё", "е"))
End Function
