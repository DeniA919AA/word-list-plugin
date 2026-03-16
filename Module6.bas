Attribute VB_Name = "Module6"
' ===== ListBlock.bas =====
Type ListItem
    Text As String
    level As Integer
End Type

Type ListBlock
    BlockType As String
    Source As String
    RangeStart As Long
    RangeEnd As Long
    ItemCount As Integer
    Items(100) As ListItem
End Type

Sub BuildListBlocks()
    Dim doc As Document
    Dim para As Paragraph
    Dim i As Long
    Dim blocks(50) As ListBlock
    Dim blockCount As Integer
    Dim listType As String
    Dim level As Integer
    Dim cleanText As String
    Dim inWord As Boolean
    Dim currentWord As ListBlock
    Dim inPT As Boolean
    Dim currentPT As ListBlock
    Dim txt As String
    Dim rawTxt As String
    Dim ltrimmed As String
    Dim ptType As String
    Dim report As String
    Dim b As Integer
    Dim it As Integer
    Dim pos As Long
    Dim firstChar As String
    Dim hasLeadingSpace As Boolean
    Dim prevPTType As String

    Set doc = ActiveDocument
    blockCount = 0
    inWord = False
    inPT = False
    prevPTType = ""
    i = 1

    ' --- Один проход по всем параграфам ---
    For Each para In doc.Paragraphs

        ' ========== Word list ==========
        If para.Range.ListFormat.listType <> wdListNoNumbering Then

            ' Если был открыт PlainText блок — закрываем
            If inPT Then
                blocks(blockCount) = currentPT
                blockCount = blockCount + 1
                inPT = False
                currentPT.ItemCount = 0
                prevPTType = ""
            End If

            Select Case para.Range.ListFormat.listType
                Case wdListBullet:                                  listType = "Bullet"
                Case wdListSimpleNumbering, wdListOutlineNumbering: listType = "Numbered"
                Case wdListMixedNumbering:                          listType = "Multilevel"
                Case Else:                                          listType = "Other"
            End Select

            level = para.Range.ListFormat.ListLevelNumber
            If level < 1 Then level = 1

            cleanText = Trim(Replace(Replace(Replace(Replace( _
                para.Range.Text, Chr(13), ""), Chr(7), ""), Chr(11), ""), Chr(12), ""))

            If Not inWord Then
                inWord = True
                currentWord.BlockType = listType
                currentWord.Source = "Word"
                currentWord.RangeStart = i
                currentWord.ItemCount = 0
            ElseIf listType <> currentWord.BlockType Then
                ' Тип сменился — сохраняем и начинаем новый
                blocks(blockCount) = currentWord
                blockCount = blockCount + 1
                currentWord.BlockType = listType
                currentWord.Source = "Word"
                currentWord.RangeStart = i
                currentWord.ItemCount = 0
            End If

            If Len(cleanText) > 0 Then
                currentWord.Items(currentWord.ItemCount).Text = cleanText
                currentWord.Items(currentWord.ItemCount).level = level
                currentWord.ItemCount = currentWord.ItemCount + 1
            End If
            currentWord.RangeEnd = i

        ' ========== Не Word list ==========
        Else

            ' Закрываем Word блок если был открыт
            If inWord Then
                blocks(blockCount) = currentWord
                blockCount = blockCount + 1
                inWord = False
                currentWord.ItemCount = 0
            End If

            ' --- Проверяем на PlainText ---
            rawTxt = Replace(Replace(Replace(para.Range.Text, Chr(13), ""), Chr(7), ""), Chr(11), "")
            txt = Trim(rawTxt)
            ltrimmed = LTrim(rawTxt)
            ptType = ""

            hasLeadingSpace = (Len(rawTxt) > 0) And _
                              (Left(rawTxt, 1) = " " Or Left(rawTxt, 1) = Chr(9))

            If Len(txt) >= 2 Then

                ' Numbered: цифры + ) или . + пробел
                pos = 1
                Do While pos <= Len(txt)
                    If Mid(txt, pos, 1) >= "0" And Mid(txt, pos, 1) <= "9" Then
                        pos = pos + 1
                    Else
                        Exit Do
                    End If
                Loop
                If pos > 1 And pos <= Len(txt) Then
                    If Mid(txt, pos, 1) = ")" Or Mid(txt, pos, 1) = "." Then
                        If pos + 1 <= Len(txt) Then
                            If Mid(txt, pos + 1, 1) = " " Then
                                ptType = "Numbered"
                            End If
                        End If
                    End If
                End If

                ' Буква + ) + пробел: a) b) c)
                If ptType = "" Then
                    firstChar = Left(txt, 1)
                    If (firstChar >= "a" And firstChar <= "z") Or _
                       (firstChar >= "A" And firstChar <= "Z") Then
                        If Len(txt) >= 3 Then
                            If Mid(txt, 2, 1) = ")" And Mid(txt, 3, 1) = " " Then
                                ptType = "Numbered"
                            End If
                        End If
                    End If
                End If

                ' Bullet-dash / Bullet-star
                If ptType = "" Then
                    If Left(ltrimmed, 2) = "- " Then
                        ptType = "Bullet-dash"
                        txt = Trim(ltrimmed)
                    ElseIf Left(ltrimmed, 2) = "* " Then
                        ptType = "Bullet-star"
                        txt = Trim(ltrimmed)
                    End If
                End If

                ' Защита от длинных строк
                If ptType <> "" Then
                    If Len(txt) > 80 Then ptType = ""
                End If

                ' Continuation
                If ptType = "" And inPT And hasLeadingSpace Then
                    If Left(ltrimmed, 2) <> "- " And Left(ltrimmed, 2) <> "* " Then
                        If Len(txt) <= 80 Then
                            ptType = "CONTINUATION"
                        End If
                    End If
                End If

            End If

            ' Закрываем PT блок если тип сменился
            If ptType <> "" And ptType <> "CONTINUATION" And inPT Then
                If ptType <> prevPTType Then
                    blocks(blockCount) = currentPT
                    blockCount = blockCount + 1
                    inPT = False
                    currentPT.ItemCount = 0
                    prevPTType = ""
                End If
            End If

            If ptType <> "" And ptType <> "CONTINUATION" Then
                If Not inPT Then
                    inPT = True
                    currentPT.BlockType = ptType
                    currentPT.Source = "PlainText"
                    currentPT.RangeStart = i
                    currentPT.ItemCount = 0
                End If
                currentPT.Items(currentPT.ItemCount).Text = txt
                currentPT.Items(currentPT.ItemCount).level = 1
                currentPT.ItemCount = currentPT.ItemCount + 1
                currentPT.RangeEnd = i
                prevPTType = ptType

            ElseIf ptType = "CONTINUATION" Then
                If currentPT.ItemCount > 0 Then
                    Dim lastIdx As Integer
                    lastIdx = currentPT.ItemCount - 1
                    currentPT.Items(lastIdx).Text = currentPT.Items(lastIdx).Text & " " & txt
                    currentPT.RangeEnd = i
                End If

            Else
                If inPT Then
                    blocks(blockCount) = currentPT
                    blockCount = blockCount + 1
                    inPT = False
                    currentPT.ItemCount = 0
                    prevPTType = ""
                End If
            End If

        End If

        i = i + 1
    Next para

    ' Закрываем последний блок
    If inWord Then
        blocks(blockCount) = currentWord
        blockCount = blockCount + 1
    End If
    If inPT Then
        blocks(blockCount) = currentPT
        blockCount = blockCount + 1
    End If

    ' --- Output ---
    report = "=== ListBlock Structure ===" & vbNewLine & vbNewLine
    For b = 0 To blockCount - 1
        report = report & "=== ListBlock " & (b + 1) & " ===" & vbNewLine
        report = report & "Type:   " & blocks(b).BlockType & vbNewLine
        report = report & "Source: " & blocks(b).Source & vbNewLine
        report = report & "Range:  " & blocks(b).RangeStart & " - " & blocks(b).RangeEnd & vbNewLine
        report = report & "Items:  " & blocks(b).ItemCount & vbNewLine
        For it = 0 To blocks(b).ItemCount - 1
            report = report & "  [Lvl." & blocks(b).Items(it).level & "] " & blocks(b).Items(it).Text & vbNewLine
        Next it
        report = report & vbNewLine
    Next b
    report = report & "Total ListBlocks: " & blockCount

    ResultText = report
    UserForm1.Show
End Sub

