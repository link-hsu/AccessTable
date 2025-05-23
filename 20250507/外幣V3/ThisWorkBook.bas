Option Explicit

Private Sub Workbook_Open()
    Dim ws As Worksheet
    ' 清除分頁顏色
    For Each ws In Me.Worksheets
        ws.Tab.ColorIndex = xlColorIndexNone
    Next ws

    Me.Worksheets("TABLE41").Tab.ColorIndex = 4
    Me.Worksheets("AI822").Tab.ColorIndex = 4
    
    ' 開啟切換到ControlPanel分頁
    On Error Resume Next
    Me.Worksheets("ControlPanel").Activate
    On Error GoTo 0
End Sub
