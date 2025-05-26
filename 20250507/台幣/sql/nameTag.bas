Option Explicit
' 透過row col 組合中間的value
Public Sub FillTableLabels( _
    Optional ByVal headerRow As Long = 1, _
    Optional ByVal headerCol As Long = 1)

    Dim ws        As Worksheet
    Dim lastRow   As Long, lastCol As Long
    Dim r         As Long, c As Long
    Dim rowHdr    As String, colHdr As String
    Dim target    As Range
    
    Set ws = ThisWorkbook.Worksheets("ai605")
    
    ' 找出範圍邊界
    With ws
        lastRow = .Cells(.Rows.Count, headerCol).End(xlUp).Row
        lastCol = .Cells(headerRow, .Columns.Count).End(xlToLeft).Column
    End With
    
    Application.ScreenUpdating = False
    
    ' 逐一把 "rowHeader_colHeader" 填入對應儲存格，並靠左上對齊
    For r = headerRow + 1 To lastRow
        rowHdr = Trim$(ws.Cells(r, headerCol).Value)
        If rowHdr = "" Then GoTo ContinueRows
        
        ' 清理：只留字母、數字和底線
        rowHdr = Replace(Application.WorksheetFunction.Clean(rowHdr), " ", "_")
        
        For c = headerCol + 1 To lastCol
            colHdr = Trim$(ws.Cells(headerRow, c).Value)
            If colHdr = "" Then GoTo ContinueCols
            
            colHdr = Replace(Application.WorksheetFunction.Clean(colHdr), " ", "_")
            
            ' 填入組合字串
            Set target = ws.Cells(r, c)
            target.Value = rowHdr & "_" & colHdr
            With target
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlTop
                .WrapText = False
            End With
            
ContinueCols:
        Next c
ContinueRows:
    Next r
    
    Application.ScreenUpdating = True
    
    MsgBox "填值完成！共處理 " & _
           (lastRow - headerRow) * (lastCol - headerCol) & " 個儲存格。", vbInformation

End Sub


Public Sub testSub()
    Call FillTableLabels(headerRow:=1, headerCol:=1)
End Sub




















' 將value加入nameTag中

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
