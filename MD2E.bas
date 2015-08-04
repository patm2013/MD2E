Attribute VB_Name = "MD2E"
'ExcelMarkdown macro
'Created by Patrick Moore
'Make sure to add references to Microsoft Scripting Runtime
'and Microsoft VB Regular Expressions 5.5
Sub mkdown()
Attribute mkdown.VB_ProcData.VB_Invoke_Func = "Q\n14"

    'set range
    Dim rngSelection As Range
    Set rngSelection = Selection
    
    For Each c In rngSelection
    
    Dim rngNew As Range
    Set rngNew = c.Offset(0, 1)
    rngNew.Clear
    rngNew.ClearFormats
    
    'cache original string
    Dim strOriginal As String
    Dim strMod As String
    strOriginal = c.Value
    'look for bullets
    rngNew.Value = bullets(strOriginal)
    strMod = rngNew.Value
    'set new cell value
    rngNew.Value = replace(rngNew.Value, "*", "")
    rngNew.Value = replace(rngNew.Value, "_", "")
    rngNew.Value = replace(rngNew.Value, "|", "")
    rngNew.Value = replace(rngNew.Value, "%", "")
    
    'run regular expression
    Set matches = regx("([\*][^\*]*[\*])|([\_][^\_]*[\_])|([\|][^\|]*[\|])|([\%][^\%]*[\%])", strMod)
    '([\*][^\W]{1,1}[^\*]*[^\W]{1,1}[\*])|([\*][^\W]{1,1}[^\*]*[^\W]{1,1}[\_])
    For x = 0 To matches.Count - 1
        Debug.Print matches(x)
    Next
    
    
    
    Dim dctMatches As Dictionary
    Set dctMatches = indexVals(strMod, matches)
    
    'declare variables
    Dim tempDic As New Dictionary
    Dim tempCol As New Collection
    
    
    Set tempCol = dctMatches("bold")
    'for each match bold characters
    For x = 1 To tempCol.Count
        Set tempDic = tempCol(x)
        Debug.Print "Bold " & "----" & tempDic("intStart") & " ---- " & tempDic("intLength")
        Call boldCells(rngNew, tempDic("intStart"), tempDic("intLength"))
    Next
    
    Set tempCol = dctMatches("underline")
    'for each match underline characters
    For x = 1 To tempCol.Count
        Set tempDic = tempCol(x)
        Debug.Print "Underline " & "----" & tempDic("intStart") & " ---- " & tempDic("intLength")
        Call underlineCells(rngNew, tempDic("intStart"), tempDic("intLength"))
    Next
    
    Set tempCol = dctMatches("titles")
    'for each match red characters
    For x = 1 To tempCol.Count
        Set tempDic = tempCol(x)
        Debug.Print "Title " & "----" & tempDic("intStart") & " ---- " & tempDic("intLength")
        Call titleCells(rngNew, tempDic("intStart"), tempDic("intLength"))
    Next
    
    
    Set tempCol = dctMatches("red")
    'for each match red characters
    For x = 1 To tempCol.Count
        Set tempDic = tempCol(x)
        Debug.Print "Color " & "----" & tempDic("intStart") & " ---- " & tempDic("intLength")
        Call redCells(rngNew, tempDic("intStart"), tempDic("intLength"))
    Next
    
    
    Next
End Sub

Function regx(strPattern, strOriginal) As Variant

    Dim regEx As New RegExp
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = strPattern
    End With
    
    Set regx = regEx.Execute(strOriginal)

End Function

Function regxR(strPattern, strNew, strOriginal) As String

    Dim regEx As New RegExp
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = strPattern
    End With
    
    regxR = regEx.replace(strOriginal, strNew)

End Function

Sub boldCells(rng As Range, intStart As Integer, intLength As Integer)

    rng.Characters(intStart, intLength).Font.bold = True

End Sub

Sub underlineCells(rng As Range, intStart As Integer, intLength As Integer)

    rng.Characters(intStart, intLength).Font.Underline = True

End Sub

Sub redCells(rng As Range, intStart As Integer, intLength As Integer)

    rng.Characters(intStart, intLength).Font.Color = RGB(255, 0, 0)

End Sub


Sub titleCells(rng As Range, intStart As Integer, intLength As Integer)

    With rng.Characters(intStart, intLength).Font
        .Underline = True
        .bold = True
    End With

End Sub
Function indexVals(strOriginal As String, matches As Variant) As Dictionary

    Dim rootDic As Dictionary
    Set rootDic = New Dictionary

    'set variables
    Dim intStart As Integer
    Dim intLength As Integer
    Dim intEnd As Integer
    Dim strNew As String
    
    
    Dim colBold As Collection
    Set colBold = New Collection
    
    Dim colRed As Collection
    Set colRed = New Collection
    
    Dim colUnderline As Collection
    Set colUnderline = New Collection
    
    Dim colTitle As Collection
    Set colTitle = New Collection
    
    Dim dctIndex As Dictionary
    
    strNew = strOriginal
    
    'for each match add indexes to dictionary
    For x = 0 To matches.Count - 1
        'clear out tmpDictionary
        Set dctIndex = New Dictionary
        'add integer indexes to dictionary
        dctIndex.Add "intStart", InStr(strNew, matches(x))
        dctIndex.Add "intLength", Len(matches(x)) - (2)
        dctIndex.Add "intEnd", intStart + intLength
        'add dictionary to collection
        'remove the stars from current occurrence
        Select Case Left(matches(x), 1)
            Case "*"
                colBold.Add dctIndex
                strNew = replace(strNew, "*", "", 1, 2)
            Case "_"
                colUnderline.Add dctIndex
                strNew = replace(strNew, "_", "", 1, 2)
            Case "|"
                colRed.Add dctIndex
                strNew = replace(strNew, "|", "", 1, 2)
            Case "%"
                colTitle.Add dctIndex
                strNew = replace(strNew, "%", "", 1, 2)
        End Select
    Next
    
    rootDic.Add "bold", colBold
    rootDic.Add "underline", colUnderline
    rootDic.Add "red", colRed
    rootDic.Add "titles", colTitle
    
    Set indexVals = rootDic
    
End Function



Function bullets(strSplit As String) As String

    Dim strArr() As String
    Dim strTest As String
    strTest = Chr(10)
    strArr = Split(strSplit, strTest)
    Dim strCurrent As String
    Dim strTrimmed As String
    For x = LBound(strArr) To UBound(strArr)
        strCurrent = strArr(x)
        If Len(strCurrent) > 1 Then
            strTrimmed = Right(strCurrent, Len(strCurrent) - 1)
            'Debug.Print strTrimmed
            Select Case Left(strCurrent, 1)
                Case "+"
                    strArr(x) = Chr(149) & Space(1) & strTrimmed
                Case "-"
                    strArr(x) = Space(3) & Chr(149) & Space(1) & strTrimmed
            End Select
        End If
       ' Debug.Print CStr(x) & " ---- " & strArr(x)
    Next

    Dim strCombined As String

    For x = LBound(strArr) To UBound(strArr)
        If x <> UBound(strArr) Then
            strCombined = strCombined & strArr(x) & Chr(10)
        Else
            strCombined = strCombined & strArr(x)
        End If
    Next
    
    bullets = strCombined

End Function


