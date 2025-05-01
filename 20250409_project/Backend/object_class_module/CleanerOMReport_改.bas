Option Compare Database

Implements ICleaner

Private clsColsToHandle As Variant

Private clsHasFile As Boolean
Private clsHasData As Boolean

Public Function ICleaner_HasFile() As Boolean
    ICleaner_HasFile = clsHasFile
End Function

Public Function ICleaner_HasData() As Boolean
    ICleaner_HasData = clsHasData
End Function

Public Sub ICleaner_Initialize(Optional ByVal sheetName As Variant = 1, _
                               Optional ByVal loopColumn As Integer = 1, _
                               Optional ByVal leftToDelete As Integer = 2, _
                               Optional ByVal rightToDelete As Integer = 3, _
                               Optional ByVal rowsToDelete As Variant, _
                               Optional ByVal colsToDelete As Variant, _
                               Optional ByVal colsToHandle As Variant)        
    clsColsToHandle = colsToHandle
End Sub

Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String, _
                                ByVal xlApp As Excel.Application)

    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    Dim i As Long, lastRow As Long
    Dim colsDict As Object
    Dim colIndex As Variant

    If Dir(fullFilePath) = "" Then
        clsHasFile = False
        MsgBox "File not found: " & fullFilePath
        Exit Sub
    Else
        clsHasFile = True
    End If

    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Sheets(1)

    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    clsHasData = (lastRow > 1)

    If Not (IsObject(clsColsToHandle) And TypeName(clsColsToHandle) = "Dictionary") Then
        GoTo CleanUp
    End If

    Set colsDict = clsColsToHandle

    For i = lastRow To 2 Step -1
        If colsDict.Exists("FormatCols") Then
            For Each colIndex In colsDict("FormatCols")
                xlsht.Cells(i, colIndex).Value = Format(xlsht.Cells(i, colIndex).Value, "00000000")
            Next
        End If

        If colsDict.Exists("RemovePercentCols") Then
            For Each colIndex In colsDict("RemovePercentCols")
                xlsht.Cells(i, colIndex).Value = Replace(xlsht.Cells(i, colIndex).Value, "%", "")
            Next
        End If
    Next

    xlbk.Save
CleanUp:
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing

    If clsHasFile And clsHasData Then
        MsgBox "完成清理：" & cleaningType & vbCrLf & fullFilePath
    End If
End Sub

' 空的 additionalClean（可留）
Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String, _
                                    ByVal xlApp As Excel.Application)
    ' No additional actions
End Sub



Option Explicit

Sub RunNoteTransactionClean()
    Dim xlApp As Excel.Application
    Set xlApp = New Excel.Application

    Dim cleaner As CleanRowsColsDelete
    Set cleaner = New CleanRowsColsDelete

    ' —— 建立單一報表的 config Dictionary —— 
    Dim colsDict As Object
    Set colsDict = CreateObject("Scripting.Dictionary")

    ' 統一編號要補零（F 欄＝第 6 欄）
    colsDict("FormatCols") = Array("F", "O", "P")
    ' 票載利率 T 欄（第 20 欄） 要移除百分比
    colsDict("RemovePercentCols") = Array("T")

    ' 初始化：把 colsDict 丟進去
    cleaner.ICleaner_Initialize sheetName:=1, loopColumn:=1, leftToDelete:=2, rightToDelete:=3, rowsToDelete:=Empty, colsToDelete:=Empty, colsToHandle:=colsDict

    ' 執行清理
    cleaner.ICleaner_CleanReport "C:\data\note_tx.xlsx", "票券交易明細表", xlApp

    ' 善後
    xlApp.Quit
    Set xlApp = Nothing
End Sub

Public Function GetCleaner(ByVal cleaningType As String) As ICleaner
    Dim colsDict As Object
    Set colsDict = CreateObject("Scripting.Dictionary")

    Select Case cleaningType
        'CleanerRowsDelete
        Case "DBU_MM4901B"
            Set GetCleaner = New CleanerRowsDelete
            GetCleaner.Initialize 1, 1, 2, 3, Array("USD", "AUD", "CAD", "CHF", "CNY", "EUR", "GBP", "HKD", "JPY", "NZD", "SGD", "THB", "ZAR", "TWD", "總計:", "主管")
        Case "OBU_MM4901B"
            Set GetCleaner = New CleanerRowsDelete
            GetCleaner.Initialize 1, 1, 2, 3, Array("USD", "AUD", "CAD", "CHF", "CNY", "EUR", "GBP", "HKD", "JPY", "NZD", "SGD", "THB", "ZAR", "TWD", "總計:", "主管")
        
        'CleanerUnitACCurr
        Case "DBU_AC5602"
            Set GetCleaner = New CleanerUnitCurr
            GetCleaner.Initialize 1, 3, 2, 3, Array("合計", "總資產"), Array("K", "G", "E", "D", "C", "A")
        Case "OBU_AC5602"
            Set GetCleaner = New CleanerUnitCurr
            GetCleaner.Initialize 1, 3, 2, 3, Array("合計", "總資產"), Array("K", "G", "E", "D", "C", "A")
        Case "OBU_AC4603"
            Set GetCleaner = New CleanerUnitCurr
            GetCleaner.Initialize 1, 3, 2, 4, Array("合計", "總資產:"), Array("K", "H", "F", "D", "C", "A")

        'CleanerRowsColsDelete
        Case "OBU_AC5411B"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, , , Array("小計", "總收入", "總支出", "純益"), Array("J", "H", "F", "E", "D", "A")
        Case "DBU_CM2810"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, 2, 3, Array("主管", "TWD", "總計"), Array("Q", "O")
        Case "DBU_DL9360"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 2, 2, 3, Array("交易"), Array("K")
            '*******要測試為什麼isempty和其他方法刪除不掉空格row，因為lastRow沒有那一行
        Case "OBU_DL6320"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, 2, 2, Array("總計", "襄理"), Array("R", "Q", "O", "N", "M")
        Case "DBU_DL6850"
            Set GetCleaner = New CleanerRowsColsDelete
            GetCleaner.Initialize 1, 1, 2, 2, Array("小計", "總計", "主管"), Array("L")
        
        'CleanerAC_PluralCurr
        Case "DBU_AC5601"
            Set GetCleaner = New CleanerPluralCurr
        Case "OBU_AC5601"
            Set GetCleaner = New CleanerPluralCurr
        Case "OBU_AC4620B"
            Set GetCleaner = New CleanerPluralCurr

        'CleanerIsEmpty
        Case "OBU_CF6320"
            Set GetCleaner = New CleanerIsEmpty
        Case "DBU_CF6850"
            Set GetCleaner = New CleanerIsEmpty
        Case "OBU_FC7700B"
            Set GetCleaner = New CleanerIsEmpty
        Case "OBU_FC9450B"
            Set GetCleaner = New CleanerIsEmpty
        
        'Special
        Case "FXDebtEvaluation"
            Set GetCleaner = New CleanerFXDebtEvaluation
        Case "AssetsImpairment"
            Set GetCleaner = New CleanerCleanAssetsImpairment

        Case "BondTransactionDetails"
            ' 統一編號要補零（F 欄＝第 6 欄）
            colsDict("FormatCols") = Array("F")
            Set GetCleaner = New CleanerOMReport
            GetCleaner.Initialize , , , , , , colsDict
        Case "BillTransactionDetails"
            ' 統一編號要補零（F 欄＝第 6 欄）
            colsDict("FormatCols") = Array("F", "O", "P")
            ' 票載利率 T 欄（第 20 欄） 要移除百分比
            colsDict("RemovePercentCols") = Array("T")
            Set GetCleaner = New CleanerOMReport
            GetCleaner.Initialize , , , , , , colsDict
            
        Case Else
            MsgBox "未知的清理類型: " & cleaningType
            Set GetCleaner = Nothing
    End Select
End Function
