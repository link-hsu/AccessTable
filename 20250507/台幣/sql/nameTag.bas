Option Explicit

Sub NameMultipleSheetCellsByValue_NoRestrictions()
    Dim config As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim area As Range, cell As Range
    Dim rngAll As Range
    Dim nm As String
    
    '=== Step 1: 定義要處理的工作表與範圍 ===
    config = Array( _
        Array("Sheet1", "A1:A5,C1:C3,E2"), _
        Array("Sheet2", "B2:B6,D1"), _
        Array("DataSheet", "F1:F10,H2:H4") _
    )
    
    '=== Step 2: 逐一處理每個工作表與其對應範圍 ===
    For i = LBound(config) To UBound(config)
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(config(i)(0))
        If ws Is Nothing Then
            MsgBox "找不到工作表：" & config(i)(0), vbExclamation
            Err.Clear
            GoTo NextConfig
        End If
        
        Set rngAll = ws.Range(config(i)(1))
        If rngAll Is Nothing Then
            MsgBox "無效範圍：" & config(i)(1) & "（工作表：" & ws.Name & "）", vbExclamation
            Err.Clear
            GoTo NextConfig
        End If
        On Error GoTo 0
        
        '=== Step 3: 命名每個儲存格（無名稱驗證）===
        For Each area In rngAll.Areas
            For Each cell In area.Cells
                nm = Trim(cell.Value)
                
                If Len(nm) = 0 Then GoTo SkipCell
                
                On Error Resume Next
                ThisWorkbook.Names.Add Name:=nm, RefersTo:="=" & "'" & ws.Name & "'!" & cell.Address
                If Err.Number <> 0 Then
                    MsgBox "命名失敗：""" & nm & """，位置：" & ws.Name & "!" & cell.Address & _
                           vbCrLf & "錯誤：" & Err.Description, vbCritical
                    Err.Clear
                End If
                On Error GoTo 0
                
SkipCell:
            Next cell
        Next area
        
NextConfig:
        Set ws = Nothing
        Set rngAll = Nothing
    Next i
    
    MsgBox "全部處理完成！", vbInformation
End Sub
