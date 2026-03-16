Attribute VB_Name = "Module1"
' ===== DetectLists.bas =====
Public ResultText As String

Type ListBlockData
    listType As String
    StartIndex As Long
    EndIndex As Long
    ItemCount As Integer
    Items(100) As String
    Levels(100) As Integer
End Type

Sub DetectLists()
    Dim doc As Document
    Dim para As Paragraph
    Dim i As Long
    Dim report As String
    Dim inList As Boolean
    Dim currentBlock As ListBlockData
    Dim listType As String
    Dim level As Integer
    Dim cleanText As String
    Dim j As Integer
    Dim allBlocks(50) As ListBlockData
    Dim allCount As Integer

    Set doc = ActiveDocument
    inList = False
    allCount = 0
    i = 1

    For Each para In doc.Paragraphs
        If para.Range.ListFormat.listType <> wdListNoNumbering Then

            ' --- Определяем тип ---
            Select Case para.Range.ListFormat.listType
                Case wdListBullet:                                   listType = "Bullet"
                Case wdListSimpleNumbering, wdListOutlineNumbering:  listType = "Numbered"
                Case wdListMixedNumbering:                           listType = "Multilevel"
                Case Else:                                           listType = "Other"
            End Select

            ' --- Уровень берём только из ListLevelNumber ---
            level = para.Range.ListFormat.ListLevelNumber
            If level < 1 Then level = 1

            ' --- Чистим текст ---
            cleanText = Replace(Replace(Replace(Replace( _
                para.Range.Text, Chr(13), ""), Chr(7), ""), Chr(11), ""), Chr(12), "")
            cleanText = Trim(cleanText)

            If Not inList Then
                ' Начинаем новый блок
                inList = True
                currentBlock.listType = listType
                currentBlock.StartIndex = i
                currentBlock.ItemCount = 0
            ElseIf listType <> currentBlock.listType Then
                ' *** ИСПРАВЛЕНИЕ кейса 11 ***
                ' Тип сменился — сохраняем текущий блок и начинаем новый
                allBlocks(allCount) = currentBlock
                allCount = allCount + 1

                currentBlock.listType = listType
                currentBlock.StartIndex = i
                currentBlock.ItemCount = 0
            End If

            currentBlock.Items(currentBlock.ItemCount) = cleanText
            currentBlock.Levels(currentBlock.ItemCount) = level
            currentBlock.ItemCount = currentBlock.ItemCount + 1
            currentBlock.EndIndex = i

        Else
            ' Не список — закрываем блок если был открыт
            If inList Then
                allBlocks(allCount) = currentBlock
                allCount = allCount + 1
                inList = False
                currentBlock.ItemCount = 0
            End If
        End If

        i = i + 1
    Next para

    ' Закрываем последний блок если документ закончился
    If inList Then
        allBlocks(allCount) = currentBlock
        allCount = allCount + 1
    End If

    ' --- Формируем отчёт ---
    report = "=== DetectLists: Word lists ===" & vbNewLine & vbNewLine

    Dim b As Integer
    For b = 0 To allCount - 1
        report = report & "--- Block " & (b + 1) & " ---" & vbNewLine
        report = report & "Type: " & allBlocks(b).listType & vbNewLine
        report = report & "Paragraphs: " & allBlocks(b).StartIndex & " - " & allBlocks(b).EndIndex & vbNewLine
        If allBlocks(b).ItemCount > 0 Then
            report = report & "Items: " & allBlocks(b).ItemCount & vbNewLine
            For j = 0 To allBlocks(b).ItemCount - 1
                report = report & "  [Lvl." & allBlocks(b).Levels(j) & "] " & allBlocks(b).Items(j) & vbNewLine
            Next j
        Else
            report = report & "Items: 0 (empty)" & vbNewLine
        End If
        report = report & vbNewLine
    Next b

    report = report & "Total blocks: " & allCount

    ResultText = report
    UserForm1.Show
End Sub

