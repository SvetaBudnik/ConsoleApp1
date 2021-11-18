Private Sub CommandButton1_Click()
    Dim docEnd As Range
    Dim tblNew As Table
    Dim tblName As String
    Dim Columns As Integer
    Dim Rows As Integer

    If IsNumeric(TextBoxColumn.Value) Then
        Columns = TextBoxColumn.Value
    Else
        MsgBox "Вы не ввели количество стобцов, по умолчанию выбрано 1"
  Columns = 1
    End If

    If IsNumeric(TextBoxRow.Value) Then
        Rows = TextBoxRow.Value
    Else
        MsgBox "Вы не ввели количество строк, по умолчанию выбрано 1"
  Rows = 1
    End If

    tblName = "Таблица" & CStr(ActiveDocument.Tables.Count) & ". " & TextBoxName.Value
 Set docEnd = ActiveDocument.Content
 docEnd.Collapse Direction:=wdCollapseEnd
 docEnd.InsertAfter Text:=tblName
 docEnd.Collapse Direction:=wdCollapseEnd
 
 Set tblNew = ActiveDocument.Tables.Add( _
 Range:=docEnd, _
 NumRows:=Rows, _
 NumColumns:=Columns)
 
 With tblNew
        .Borders.InsideLineStyle = wdLineStyleSingle
        .Borders.OutsideLineStyle = wdLineStyleSingle
        For intX = 1 To Rows
            .Cell(intX, 1).Range.InsertAfter CStr(intX) & "."
    Next intX
        .Rows.Alignment = wdAlignRowCenter
    End With

    Unload Me
End Sub

Private Sub TextBoxColumn_Change()
    If IsNumeric(TextBoxColumn.Value) Then
        If (TextBoxColumn.Value > 63) Or (TextBoxColumn.Value < 1) Then
            MsgBox "Значение должно быть в диапазоне от 1 до 63"
  TextBoxColumn.Value = "1"
        End If
    End If
End Sub

Private Sub TextBoxRow_Change()
    If IsNumeric(TextBoxRow.Value) Then
        If (TextBoxRow.Value > 128) Or (TextBoxRow.Value < 1) Then
            MsgBox "Значение должно быть в диапазоне от 1 до 128"
  TextBoxRow.Value = "1"
        End If
    End If
End Sub

