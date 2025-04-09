Option Compare Database

Implements ICleaner

Public Sub ICleaner_Initialize(Optional ByVal sheetName As Variant = 1, _
                               Optional ByVal loopColumn As Integer = 1, _
                               Optional ByVal leftToDelete As Integer = 2, _
                               Optional ByVal rightToDelete As Integer = 3, _
                               Optional ByVal rowsToDelete As Variant, _
                               Optional ByVal colsToDelete As Variant)
End Sub

Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String) 
    Dim xlApp As Excel.Application
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet
    
    ' Dim xlbk As Workbook
    ' Dim xlsht As Worksheet

    Dim i As Integer, j As Integer
    Dim lastRow As Integer, startRow As Integer, endRow As Integer, newLastRow As Integer

    Dim assetRows As Collection

    Dim currArray() As String
    Dim tableColumns As Variant


    Redim currArray(0)
    ' Dim fullFilePath As String
    ' fullFilePath = "D:\DavidHsu\testFile\vba\test\OBU_AC5601_33_AC_E_20241231_r.xls"
    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
    End If

    If Dir(Replace(fullFilePath, "xls", "txt")) = "" Then
        MsgBox "File does not exist in path: " & Replace(fullFilePath, "xls", "txt")
    End If

    Set xlApp = Excel.Application
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False

    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    ' Set xlbk = Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Worksheets(1)
    
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    xlsht.Columns("B").Insert shift:=xlToRight

    For i = lastRow To 2 Step -1
        If IsEmpty(xlsht.Cells(i, 1).Value) Or _
           IsNumeric(xlsht.Cells(i, 1).Value) Or _
           xlsht.Cells(i, 1).Value = "放款類" Or _
           xlsht.Cells(i, 1).Value = "存款類" Or _
           xlsht.Cells(i, 1).Value = "負債類" Or _
           xlsht.Cells(i, 1).Value = "損益類 - 收入" Or _
           xlsht.Cells(i, 1).Value = "損益類 - 費用" Or _
           xlsht.Cells(i, 1).Value = "業主權益類" Or _
           Left(xlsht.Cells(i, 1).Value, 2) = "或有" Or _
           Left(xlsht.Cells(i, 1).Value, 2) = "主管" Then
            xlsht.Rows(i).Delete
        End If
        xlsht.Cells(i, "B") = xlsht.Cells(i, "C").Value & xlsht.Cells(i, "E").Value
    Next i

    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row

    xlsht.Columns("K").Delete
    xlsht.Columns("G").Delete
    xlsht.Columns("C:E").Delete

    currArray = ReadCurrencyFromTxt(Replace(fullFilePath, "xls", "txt"))
    'tableColumns = GetTableColumns(cleaningType)
    
    Set assetRows = New Collection

    ' 找出所有 "資產類" 出現的位置
    For i = 1 To lastRow
        If xlsht.Cells(i, 1).Value Like "*資產類*" Then
            assetRows.Add i
        End If
    Next i

    ' 依照奇數次到偶數次作為區間分割
    For i = 1 To assetRows.Count Step 2
        If i + 1 < assetRows.Count Then
            startRow = assetRows(i) + 1
            endRow = assetRows(i + 2) - 1
        Else
            startRow = assetRows(i) + 1
            endRow = lastRow
        End If
        
        ' 建立新的工作表
        xlbk.Sheets.Add After:=xlbk.Sheets(xlbk.Sheets.count)
        With ActiveSheet
            .Name = currArray(i \ 2)
            ' 複製資料 A:AA 欄
            xlsht.Range(xlsht.Cells(startRow, 1), xlsht.Cells(endRow, 27)).Copy
            .Range("A2").PasteSpecial Paste:=xlPasteValues
 
            newLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            For j = newLastRow To 2 Step -1
                If .Cells(j, 1).Value = "資產類" Then .Rows(j).Delete
                .Cells(j, "F").Value = currArray(i \ 2)
            Next j
            .Columns("A").Delete
            'Set Table Columns
            ' .Range("A1").Resize(1, UBound(tableColumns) - LBound(tableColumns) + 1).Value = tableColumns
        End With
    Next i
    
    xlsht.Delete
    
    ' 清除剪貼簿
    xlApp.CutCopyMode = False
    xlApp.ScreenUpdating = True
    xlApp.DisplayAlerts = True

    Set assetRows = Nothing

    xlbk.Save
    xlbk.Close False
    Set xlsht = Nothing
    Set xlbk = Nothing
    xlApp.Quit
    Set xlApp = Nothing

    ' On Error Resume Next
    ' Set xlsht = Nothing
    ' Set xlbk = Nothing
    ' Set assetRows = Nothing
    ' On Error GoTo 0
    
    MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
    Debug.Print "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
End Sub

'---------------------------------------------------------------
'function call and execute work for dictionary

Public Function ReadCurrencyFromTxt(ByVal filePath As String) As String()
    Dim fileNum As Integer
    Dim lineContent As String
    Dim arrCurr() As String
    Dim curr As String
    Dim i As Integer
    Dim dict As Object ' 使用 Dictionary 來存放不重複的貨幣

    ' 建立 Dictionary 來去除重複值
    Set dict = CreateObject("Scripting.Dictionary")

    ' 取得可用的檔案編號
    fileNum = FreeFile()

    ' 開啟檔案進行讀取
    Open filePath For Input As #fileNum
    
    ' 逐行讀取檔案內容
    Do Until EOF(fileNum)
        Line Input #fileNum, lineContent ' 讀取一行內容
    
        ' 如果該行包含 "幣? ??別"，則提取後面的值
        If Left(lineContent, 6) = "幣    別" Then
    
            ' 取出 "幣? ??別" 後的貨幣代碼
            curr = Trim(Mid(lineContent, 12, 3))
            
            ' 確保 curr 不是空值，且不重複加入
            If curr <> "" And Not dict.Exists(curr) Then
                dict.Add curr, Nothing
            End If
        End If
    Loop
    
    ' 關閉檔案
    Close #fileNum
    
    ' 將 Dictionary 轉換為陣列
    If dict.count > 0 Then
        ReDim arrCurr(dict.count - 1)
        i = 0
        Dim key As Variant
        For Each key In dict.Keys
            arrCurr(i) = key
            i = i + 1
        Next key
    End If  
    ' 回傳結果
    ReadCurrencyFromTxt = arrCurr  
End Function
    
Public Sub ICleaner_additionalClean(ByVal fullFilePath As String, _
                                    ByVal cleaningType As String, _
                                    ByVal dataDate As Date, _
                                    ByVal dataMonth As Date, _
                                    ByVal dataMonthString As String)
    'implement operations here
End Sub
