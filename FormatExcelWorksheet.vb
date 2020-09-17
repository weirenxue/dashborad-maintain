Sub main20200730()
    Call CopyPasteAllToValue
    Call SeparateTableExperimentalEducation
    Call AddFailingFields
    Call AddSystemErrorField
    Call CheckNFields
End Sub


Sub CopyPasteAllToValue()
    Dim worksheetsArray As Variant
    worksheetsArray = Array("公版", "北科大", "臺北市", "高雄市", "新北市", "其他", "實驗教育")
    For Each currentWorksheet In worksheetsArray
        Sheets(currentWorksheet).Select
        Cells.Select
        Selection.EntireColumn.Hidden = False
        Application.CutCopyMode = False
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Next currentWorksheet
End Sub




Sub AddFailingFields()
    Dim addKAndLField As Variant
    addKAndLField = Array("公版", "北科大", "臺北市", "高雄市", "新北市", "其他", "實驗教育", "海外平台")
    For Each currentWorksheet In addKAndLField
        Sheets(currentWorksheet).Select
        If currentWorksheet <> "公版" Then
            If currentWorksheet = "北科大" Then
                Columns("J:K").Select
                Selection.Delete Shift:=xlToLeft
            End If
            If currentWorksheet = "實驗教育" Or currentWorksheet = "海外平台" Then
                Columns("J:J").Select
                Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                Range("J1").Select
                ActiveCell.FormulaR1C1 = "課程諮詢記錄件數"
            End If
            Columns("K:K").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("K1").Select
            ActiveCell.FormulaR1C1 = "課程諮詢記錄學生數"
            If currentWorksheet <> "其他" Then
                Columns("L:L").Select
                Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                Range("L1").Select
                ActiveCell.FormulaR1C1 = "備註"
            End If
        End If
        ActiveWindow.View = xlPageBreakPreview
    Next currentWorksheet
End Sub

Sub SeparateTableExperimentalEducation()
    Sheets.Add(After:=Worksheets("實驗教育")).Name = "海外平台"
    Sheets("實驗教育").Select
    Rows("14:18").Select
    Selection.Cut
    Sheets("海外平台").Select
    ActiveSheet.Paste
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Sheets("實驗教育").Select
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
End Sub

Sub AddSystemErrorField()
    Dim lastColumnNumber As Long
    
    Dim worksheetsArray As Variant
    worksheetsArray = Array("公版", "北科大", "臺北市", "高雄市", "新北市", "其他", "實驗教育", "海外平台")
    Dim wrongWorksheets As String
    wrongWorksheets = ""
    For Each currentWorksheet In worksheetsArray
        Sheets(currentWorksheet).Select
	lastColumnNumber = ActiveSheet.UsedRange.Columns.Count
        If Cells(1, lastColumnNumber).Value <> "系統故障" Then
        	Cells(1, lastColumnNumber + 1).Value = "系統故障"
   	End If
    Next currentWorksheet
End Sub

Sub CheckNFields()
    Dim worksheetsArray As Variant
    worksheetsArray = Array("公版", "北科大", "臺北市", "高雄市", "新北市", "其他", "實驗教育", "海外平台")
    Dim wrongWorksheets As String
    wrongWorksheets = ""
    For Each currentWorksheet In worksheetsArray
        Sheets(currentWorksheet).Select
        If ActiveSheet.UsedRange.Columns.Count <> 15 Then
            wrongWorksheets = wrongWorksheets + " '" + currentWorksheet + "', "
        End If
    Next currentWorksheet
    If wrongWorksheets <> "" Then
        MsgBox wrongWorksheets + "資料不齊全"
    Else
        MsgBox "格式化成功。"
    End If
End Sub
