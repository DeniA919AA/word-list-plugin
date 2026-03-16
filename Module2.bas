Attribute VB_Name = "Module2"
Sub DetectPlainText()
    Dim doc As Document
    Dim para As Paragraph
    Dim i As Long
    Dim report As String
    Dim txt As String
    Dim rawTxt As String
    Dim ltrimmed As String
    Dim listType As String
    Dim foundCount As Integer
    Dim prevType As String
    Dim blockCount As Integer
    Dim ItemCount As Integer
    Dim inBlock As Boolean
    Dim pos As Long
    Dim firstChar As String
    Dim k As Long
    Dim lastNL As Long
    Dim hasLeadingSpace As Boolean

    Set doc = ActiveDocument
    report = "=== DetectPlainText: Plain-text lists ===" & vbNewLine & vbNewLine
    foundCount = 0
    blockCount = 0
    prevType = ""
    ItemCount = 0
    inBlock = False
    i = 1

    For Each para In doc.Paragraphs
        rawTxt = Replace(Replace(Replace(para.Range.Text, Chr(13), ""), Chr(7), ""), Chr(11), "")
        txt = Trim(rawTxt)
        ltrimmed = LTrim(rawTxt)
        listType = ""

        If i >= 1 And i <= 80 Then
            Debug.Print "para " & i & " | raw=[" & rawTxt & "] | ltrimmed=[" & ltrimmed & "] | hasLead=" & hasLeadingSpace
        End If

        hasLeadingSpace = (Len(rawTxt) > 0) And _
                          (Left(rawTxt, 1) = " " Or Left(rawTxt, 1) = Chr(9))

        If Len(txt) >= 2 Then

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
                            listType = "Numbered"
                        End If
                    End If
                End If
            End If

            If listType = "" Then
                firstChar = Left(txt, 1)
                If (firstChar >= "a" And firstChar <= "z") Or _
                   (firstChar >= "A" And firstChar <= "Z") Then
                    If Len(txt) >= 3 Then
                        If Mid(txt, 2, 1) = ")" And Mid(txt, 3, 1) = " " Then
                            listType = "Numbered"
                        End If
                    End If
                End If
            End If

            If listType = "" Then
                If Left(ltrimmed, 2) = "- " Then
                    listType = "Bullet-dash"
                    txt = Trim(ltrimmed)
                ElseIf Left(ltrimmed, 2) = "* " Then
                    listType = "Bullet-star"
                    txt = Trim(ltrimmed)
                End If
            End If

            If listType <> "" Then
                If Len(txt) > 80 Then listType = ""
            End If

            If listType = "" And inBlock And hasLeadingSpace Then
                If Left(ltrimmed, 2) <> "- " And Left(ltrimmed, 2) <> "* " Then
                    If Len(txt) <= 80 Then
                        listType = "CONTINUATION"
                    End If
                End If
            End If

        End If

        If listType <> "" And listType <> "CONTINUATION" And inBlock Then
            If listType <> prevType Then
                report = report & "End: para " & (i - 1) & vbNewLine
                report = report & "Items: " & ItemCount & vbNewLine & vbNewLine
                inBlock = False
                prevType = ""
            End If
        End If

        If listType <> "" And listType <> "CONTINUATION" Then
            If Not inBlock Then
                blockCount = blockCount + 1
                ItemCount = 0
                inBlock = True
                report = report & "--- Block " & blockCount & " ---" & vbNewLine
                report = report & "Start: para " & i & vbNewLine
            End If
            foundCount = foundCount + 1
            ItemCount = ItemCount + 1
            report = report & "  [" & listType & "] " & txt & vbNewLine
            prevType = listType

        ElseIf listType = "CONTINUATION" Then
            lastNL = 0
            For k = Len(report) - 1 To 1 Step -1
                If Mid(report, k, Len(vbNewLine)) = vbNewLine Then
                    lastNL = k
                    Exit For
                End If
            Next k
            If lastNL > 0 Then
                report = Left(report, Len(report) - Len(vbNewLine)) & " " & txt & vbNewLine
            End If

        Else
            If inBlock Then
                report = report & "End: para " & (i - 1) & vbNewLine
                report = report & "Items: " & ItemCount & vbNewLine & vbNewLine
                inBlock = False
                prevType = ""
            End If
        End If

        i = i + 1
    Next para

    If inBlock Then
        report = report & "End: para " & (i - 1) & vbNewLine
        report = report & "Items: " & ItemCount & vbNewLine
    End If

    report = report & vbNewLine & "Total items: " & foundCount
    ResultText = report
    UserForm1.Show
End Sub

