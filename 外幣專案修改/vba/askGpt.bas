' ========

1.製作完申報Excel表時，產生一個pdf檔案
2.提供一個參數為儲存pdf的Root Path

' ========

Control Panel

1.將搜尋篩選過後的pdf貼入PDFCreator資料夾中
2.將資料貼入人工製作報表中，call相關ButtonSub執行程序
3.產生輔助報表pdf檔案

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\AI
D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\AI\AI230
D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\AI\AI230\申報報表
D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\AI\AI230\申報報表\202506

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\AI\AI240

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\AI\AI821

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\AI\AI822申報檔

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\CNY

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\F1F2

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\F1F2\用表

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\F1F2\申報檔

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\FB

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\FM

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\FM\FM2

D:\DavidHsu\ReportCreator\外幣月報\底稿及申報檔\FM\FM5

###寫一個函數 input 為A檔案路徑、B檔案路徑、比對分頁、比對儲存格 (輸入類似這樣的 "C11:J36,C37:E47,G37:J47,F40,F42:F47,C48:C50,J48:J50,C52:C56,J52:J56,C57:J57")
我要將傳入的兩個檔案中同樣分頁同樣儲存格中的數值去逐一比對數值，
如果數值完全一致，回傳0或NULL
如果有數值不一致，回傳所有不一致的 (檔案名稱，分頁名稱，儲存格位置)

兩個檔案分頁和儲存格位置完全一致，但是其中的數值不同，我要寫出一個這樣的Function或Sub

' =====================

AI240

Option Explicit

Public Sub CopyDataToAI240_ButtonClick()
    CopyDataToAI240 ThisWorkbook
End Sub

Sub CopyDataToAI240(Optional ByVal wb As Workbook, _
                    Optional ByVal calledByOtherExcel As Boolean = False, _
                    Optional ByVal baseDatePassed As Variant = Empty)
    Dim targetBook As Workbook
    Dim wsDL6850 As Worksheet
    Dim wsAI240 As Worksheet
    Dim inputDate As Date
    Dim baseDate As Date
    Dim rowCount As Long
    Dim copyCount0To10 As Long
    Dim copyCount11To30 As Long
    Dim copyCount31To90 As Long
    Dim copyCount91To180 As Long
    Dim copyCount181To365 As Long
    Dim copyCount366To As Long
    Dim destRow0TO10 As Long
    Dim destRow11TO30 As Long
    Dim destRow31TO90 As Long
    Dim destRow91TO180 As Long
    Dim destRow181TO365 As Long
    Dim destRow366TO As Long
    Dim i As Long

    If wb Is Nothing Then
        Set targetBook = ThisWorkbook
    Else
        Set targetBook = wb
    End If

    '設定工作表
    Set wsDL6850 = targetBook.Worksheets("DL6850原始資料")
    Set wsAI240 = targetBook.Worksheets("AI240")

    If calledByOtherExcel Then
        inputDate = baseDatePassed
    Else
        '讀取基準日期
        inputDate = InputBox("請輸入基準日(日期格式yyyy/mm/dd)：")
    End If                

    baseDate = inputDate

    '填入基準日期至DL6850原始資料工作表和AI240工作表
    wsDL6850.Range("P1").Value = baseDate
    wsAI240.Range("A2").Value = baseDate
    
    
    '清空AI240工作表數據
    ' 清空範圍 A9:I58
    wsAI240.Range("A9:I58").ClearContents
    ' 清空範圍 L9:T58
    wsAI240.Range("L9:T58").ClearContents
    ' 清空範圍 A90:I139
    wsAI240.Range("A90:I139").ClearContents
    ' 清空範圍 L90:T139
    wsAI240.Range("L90:T139").ClearContents
    ' 清空範圍 A153:I162
    wsAI240.Range("A153:I162").ClearContents
    ' 清空範圍 L153:T162
    wsAI240.Range("L153:T162").ClearContents
    ' 清空範圍 A170:I179
    wsAI240.Range("A170:I179").ClearContents
    ' 清空範圍 L170:T179
    wsAI240.Range("L170:T179").ClearContents
    

    If Not calledByOtherExcel Then
        Call ImportDL6850CSV
    End If

    '刪除符合條件的資料（DL6850原始資料工作表 B欄位以及 E、H、C、J 欄位的條件）
    
    rowCount = wsDL6850.Cells(wsDL6850.Rows.Count, "B").End(xlUp).Row
    For i = rowCount To 2 Step -1
        If Left(wsDL6850.Range("B" & i).Value, 2) <> "TR" Then
            wsDL6850.Rows(i).Delete
        End If
    Next i
    
    rowCount = wsDL6850.Cells(wsDL6850.Rows.Count, "B").End(xlUp).Row
    For i = rowCount To 2 Step -1
        If (wsDL6850.Range("E" & i).Value <> "TWD" And wsDL6850.Range("H" & i).Value <> "TWD") _
        Or wsDL6850.Range("C" & i).Value <= baseDate _
        Or wsDL6850.Range("J" & i).Value > baseDate Then
            wsDL6850.Rows(i).Delete
        End If
    Next i
    
 

    '將符合條件的資料複製貼入AI240工作表
    rowCount = wsDL6850.Cells(wsDL6850.Rows.Count, "B").End(xlUp).Row
    
    
    
    ' SWOP(SS or SF) and OutFlow TWD(colH)
    '起始貼入的目標列
    destRow0TO10 = 9
    destRow11TO30 = 19
    destRow31TO90 = 29
    destRow91TO180 = 39
    destRow181TO365 = 49

    '初始化計數變數
    copyCount0To10 = 0
    copyCount11To30 = 0
    copyCount31To90 = 0
    copyCount91To180 = 0
    copyCount181To365 = 0

    For i = 2 To rowCount
        If (wsDL6850.Range("A" & i).Value Like "SS*" Or wsDL6850.Range("A" & i).Value Like "SF*") And wsDL6850.Range("H" & i).Value = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case 0 To 10
                    copyCount0To10 = copyCount0To10 + 1
                    If copyCount0To10 > 10 Then
                        MsgBox "此期間流出之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow0TO10 & ":I" & destRow0TO10).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow0TO10 = destRow0TO10 + 1

                Case 11 To 30
                    copyCount11To30 = copyCount11To30 + 1
                    If copyCount11To30 > 10 Then
                        MsgBox "此期間流出之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow11TO30 & ":I" & destRow11TO30).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow11TO30 = destRow11TO30 + 1

                Case 31 To 90
                    copyCount31To90 = copyCount31To90 + 1
                    If copyCount31To90 > 10 Then
                        MsgBox "此期間流出之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow31TO90 & ":I" & destRow31TO90).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow31TO90 = destRow31TO90 + 1

                Case 91 To 180
                    copyCount91To180 = copyCount91To180 + 1
                    If copyCount91To180 > 10 Then
                        MsgBox "此期間流出之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow91TO180 & ":I" & destRow91TO180).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow91TO180 = destRow91TO180 + 1

                Case 181 To 365
                    copyCount181To365 = copyCount181To365 + 1
                    If copyCount181To365 > 10 Then
                        MsgBox "此期間流出之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow181TO365 & ":I" & destRow181TO365).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow181TO365 = destRow181TO365 + 1
            End Select
        End If
    Next i
    
    
    
    ' SWOP(SS or SF) and InFlow TWD(colE)
    '起始貼入的目標列
    destRow0TO10 = 9
    destRow11TO30 = 19
    destRow31TO90 = 29
    destRow91TO180 = 39
    destRow181TO365 = 49

    '初始化計數變數
    copyCount0To10 = 0
    copyCount11To30 = 0
    copyCount31To90 = 0
    copyCount91To180 = 0
    copyCount181To365 = 0

    For i = 2 To rowCount
        If (wsDL6850.Range("A" & i).Value Like "SS*" Or wsDL6850.Range("A" & i).Value Like "SF*") And wsDL6850.Range("E" & i).Value = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case 0 To 10
                    copyCount0To10 = copyCount0To10 + 1
                    If copyCount0To10 > 10 Then
                        MsgBox "此期間流入之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow0TO10 & ":T" & destRow0TO10).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow0TO10 = destRow0TO10 + 1

                Case 11 To 30
                    copyCount11To30 = copyCount11To30 + 1
                    If copyCount11To30 > 10 Then
                        MsgBox "此期間流入之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow11TO30 & ":T" & destRow11TO30).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow11TO30 = destRow11TO30 + 1

                Case 31 To 90
                    copyCount31To90 = copyCount31To90 + 1
                    If copyCount31To90 > 10 Then
                        MsgBox "此期間流入之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow31TO90 & ":T" & destRow31TO90).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow31TO90 = destRow31TO90 + 1

                Case 91 To 180
                    copyCount91To180 = copyCount91To180 + 1
                    If copyCount91To180 > 10 Then
                        MsgBox "此期間流入之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow91TO180 & ":T" & destRow91TO180).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow91TO180 = destRow91TO180 + 1

                Case 181 To 365
                    copyCount181To365 = copyCount181To365 + 1
                    If copyCount181To365 > 10 Then
                        MsgBox "此期間流入之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow181TO365 & ":T" & destRow181TO365).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow181TO365 = destRow181TO365 + 1
            End Select
        End If
    Next i
    
    
    
    ' SPOT(FS) and OutFlow TWD(colH)
    '起始貼入的目標列
    destRow0TO10 = 90
    destRow11TO30 = 100
    destRow31TO90 = 110
    destRow91TO180 = 120
    destRow181TO365 = 130

    '初始化計數變數
    copyCount0To10 = 0
    copyCount11To30 = 0
    copyCount31To90 = 0
    copyCount91To180 = 0
    copyCount181To365 = 0

    For i = 2 To rowCount
        If wsDL6850.Range("A" & i).Value Like "FS*" And wsDL6850.Range("H" & i).Value = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case 0 To 10
                    copyCount0To10 = copyCount0To10 + 1
                    If copyCount0To10 > 10 Then
                        MsgBox "此期間流出之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow0TO10 & ":I" & destRow0TO10).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow0TO10 = destRow0TO10 + 1

                Case 11 To 30
                    copyCount11To30 = copyCount11To30 + 1
                    If copyCount11To30 > 10 Then
                        MsgBox "此期間流出之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow11TO30 & ":I" & destRow11TO30).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow11TO30 = destRow11TO30 + 1

                Case 31 To 90
                    copyCount31To90 = copyCount31To90 + 1
                    If copyCount31To90 > 10 Then
                        MsgBox "此期間流出之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow31TO90 & ":I" & destRow31TO90).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow31TO90 = destRow31TO90 + 1

                Case 91 To 180
                    copyCount91To180 = copyCount91To180 + 1
                    If copyCount91To180 > 10 Then
                        MsgBox "此期間流出之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow91TO180 & ":I" & destRow91TO180).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow91TO180 = destRow91TO180 + 1

                Case 181 To 365
                    copyCount181To365 = copyCount181To365 + 1
                    If copyCount181To365 > 10 Then
                        MsgBox "此期間流出之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow181TO365 & ":I" & destRow181TO365).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow181TO365 = destRow181TO365 + 1
            End Select
        End If
    Next i
    
    
    
    ' SPOT(FS) and InFlow TWD(colE)
    '起始貼入的目標列
    destRow0TO10 = 90
    destRow11TO30 = 100
    destRow31TO90 = 110
    destRow91TO180 = 120
    destRow181TO365 = 130

    '初始化計數變數
    copyCount0To10 = 0
    copyCount11To30 = 0
    copyCount31To90 = 0
    copyCount91To180 = 0
    copyCount181To365 = 0

    For i = 2 To rowCount
        If wsDL6850.Range("A" & i).Value Like "FS*" And wsDL6850.Range("E" & i).Value = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case 0 To 10
                    copyCount0To10 = copyCount0To10 + 1
                    If copyCount0To10 > 10 Then
                        MsgBox "此期間流入之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow0TO10 & ":T" & destRow0TO10).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow0TO10 = destRow0TO10 + 1

                Case 11 To 30
                    copyCount11To30 = copyCount11To30 + 1
                    If copyCount11To30 > 10 Then
                        MsgBox "此期間流入之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow11TO30 & ":T" & destRow11TO30).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow11TO30 = destRow11TO30 + 1

                Case 31 To 90
                    copyCount31To90 = copyCount31To90 + 1
                    If copyCount31To90 > 10 Then
                        MsgBox "此期間流入之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow31TO90 & ":T" & destRow31TO90).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow31TO90 = destRow31TO90 + 1

                Case 91 To 180
                    copyCount91To180 = copyCount91To180 + 1
                    If copyCount91To180 > 10 Then
                        MsgBox "此期間流入之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow91TO180 & ":T" & destRow91TO180).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow91TO180 = destRow91TO180 + 1

                Case 181 To 365
                    copyCount181To365 = copyCount181To365 + 1
                    If copyCount181To365 > 10 Then
                        MsgBox "此期間流入之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow181TO365 & ":T" & destRow181TO365).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow181TO365 = destRow181TO365 + 1
            End Select
        End If
    Next i
    











    'Case for over one year
    ' SWOP(SS or SF) and OutFlow TWD(colH)
    '起始貼入的目標列
    '初始化計數變數
    destRow366TO = 153
    copyCount366To = 0


    For i = 2 To rowCount
        If (wsDL6850.Range("A" & i).Value Like "SS*" Or wsDL6850.Range("A" & i).Value Like "SF*") And wsDL6850.Range("H" & i).Value = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case Is >= 366
                    copyCount366To = copyCount366To + 1
                    If copyCount366To > 10 Then
                        MsgBox "此期間流出之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow366TO & ":I" & destRow366TO).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow366TO = destRow366TO + 1
            End Select
        End If
    Next i
    
    
    
    ' SWOP(SS or SF) and InFlow TWD(colE)
    '起始貼入的目標列
    '初始化計數變數
    destRow366TO = 153
    copyCount366To = 0


    For i = 2 To rowCount
        If (wsDL6850.Range("A" & i).Value Like "SS*" Or wsDL6850.Range("A" & i).Value Like "SF*") And wsDL6850.Range("E" & i).Value = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case Is >= 366
                    copyCount366To = copyCount366To + 1
                    If copyCount366To > 10 Then
                        MsgBox "此期間流入之SWOP筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow366TO & ":T" & destRow366TO).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow366TO = destRow366TO + 1
            End Select
        End If
    Next i
    
    
    
    ' SPOT(FS) and OutFlow TWD(colH)
    '起始貼入的目標列
    '初始化計數變數
    destRow366TO = 170
    copyCount366To = 0



    For i = 2 To rowCount
        If wsDL6850.Range("A" & i).Value Like "FS*" And wsDL6850.Range("H" & i).Value = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case Is >= 366
                    copyCount366To = copyCount366To + 1
                    If copyCount366To > 10 Then
                        MsgBox "此期間流出之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("A" & destRow366TO & ":I" & destRow366TO).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow366TO = destRow366TO + 1
            End Select
        End If
    Next i
    
    
    
    ' SPOT(FS) and InFlow TWD(colE)
    '起始貼入的目標列
    '初始化計數變數
    destRow366TO = 170
    copyCount366To = 0



    For i = 2 To rowCount
        If wsDL6850.Range("A" & i).Value Like "FS*" And wsDL6850.Range("E" & i).Value = "TWD" Then
            Select Case wsDL6850.Range("N" & i).Value
                Case Is >= 366
                    copyCount366To = copyCount366To + 1
                    If copyCount366To > 10 Then
                        MsgBox "此期間流入之SPOT筆數超過10筆"
                        Exit Sub
                    End If
                    wsAI240.Range("L" & destRow366TO & ":T" & destRow366TO).Value = wsDL6850.Range("B" & i & ":J" & i).Value
                    destRow366TO = destRow366TO + 1
            End Select
        End If
    Next i





    '完成
    MsgBox "完成"
End Sub


Sub ImportDL6850CSV()
    Dim wbImport As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim filePath As String
    Dim lastRow As Long

    ' 開啟檔案選擇視窗
    filePath = Application.GetOpenFilename("Excel 檔案 (*.xls), *.xls", , "請選擇 DL6850 Excel 檔")
    If filePath = "False" Then Exit Sub '使用者按取消

    ' 開啟選取的 CSV 檔（轉為 Excel 格式）
    Workbooks.Open Filename:=filePath
    Set wbImport = ActiveWorkbook
    Set wsSource = wbImport.Sheets(1)

    ' 指定貼上的目標工作表
    Set wsDest = ThisWorkbook.Sheets("DL6850原始資料")

    ' 找出來源的最後一列（避免多餘空白列）
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' 清除目標區原有資料（可視需求調整）
    wsDest.Range("B2:L" & wsDest.Rows.Count).ClearContents

    ' 貼上來源 A:L 到目標 B:L（從第2列開始貼）
    wsSource.Range("A1:L" & lastRow).Copy
    wsDest.Range("B1").PasteSpecial xlPasteValues

    ' 關閉 CSV 檔，不儲存
    Application.DisplayAlerts = False
    wbImport.Close SaveChanges:=False
    Application.DisplayAlerts = True

    MsgBox "DL6850 資料匯入完成！", vbInformation
End Sub




' ======================
F1_F2

Option Explicit


Public Sub MainSub_ButtonClick()
    MainSub ThisWorkbook
End Sub

Sub MainSub(Optional ByVal wb As Workbook, _
            Optional ByVal calledByOtherExcel As Boolean = False, _
            Optional ByVal baseDatePassed As Variant = Empty)

    Dim targetBook As Workbook
    Dim wsSrc_DL6850 As Worksheet, wsDst_DL6850 As Worksheet
    Dim srcRng_DL6850 As Range, dstRng_DL6850 As Range
    Dim wsSrc_CM2810 As Worksheet, wsDst_CM2810 As Worksheet
    Dim srcRng_CM2810 As Range, dstRng_CM2810 As Range
    Dim lastRow As Long
    Dim i As Long

    If wb Is Nothing Then
        Set targetBook = ThisWorkbook
    Else
        Set targetBook = wb
    End If

    Call selectionProcess_DL6850(targetBook, calledByOtherExcel, baseDatePassed)
    Call selectionProcess_CM2810(targetBook)

    Set wsSrc_DL6850 = targetBook.Worksheets("底稿_含NT_原始資料")
    Set wsDst_DL6850 = targetBook.Worksheets("底稿_含NT")

    Set wsSrc_CM2810 = targetBook.Worksheets("國內顧客_原始資料")
    Set wsDst_CM2810 = targetBook.Worksheets("國內顧客")

    ' Copy Data for DL6850
    lastRow = wsSrc_DL6850.Cells(wsSrc_DL6850.Rows.Count, "I").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "來源沒有資料 (I 欄最後一列 < 2)。", vbInformation
        Exit Sub
    End If    

    Set srcRng_DL6850 = wsSrc_DL6850.Range("A2", wsSrc_DL6850.Cells(lastRow, "I"))
    ' 目標範圍從 B2 開始，大小與來源相同（A:I 共 9 欄 -> B:J 也 9 欄）
    Set dstRng_DL6850 = wsDst_DL6850.Range("B2").Resize(srcRng_DL6850.Rows.Count, srcRng_DL6850.Columns.Count)

    lastRow = wsDst_DL6850.Cells(wsDst_DL6850.Rows.Count, "I").End(xlUp).Row
    wsDst_DL6850.Range("B2:D" & lastRow).ClearContents

    dstRng_DL6850.Value = srcRng_DL6850.Value
      
    ' ===================
    ' Copy for CM2810

    lastRow = wsSrc_CM2810.Cells(wsSrc_CM2810.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "來源沒有資料 (I 欄最後一列 < 2)。", vbInformation
        Exit Sub
    End If    

    Set srcRng_CM2810 = wsSrc_CM2810.Range("A2", wsSrc_CM2810.Cells(lastRow, "H"))
    ' 目標範圍從 B2 開始，大小與來源相同（A:I 共 9 欄 -> B:J 也 9 欄）
    Set dstRng_CM2810 = wsDst_CM2810.Range("A2").Resize(srcRng_CM2810.Rows.Count, srcRng_CM2810.Columns.Count)

    lastRow = wsDst_CM2810.Cells(wsDst_CM2810.Rows.Count, "A").End(xlUp).Row
    wsDst_CM2810.Range("A2:H" & lastRow).ClearContents

    dstRng_CM2810.Value = srcRng_CM2810.Value

    
    '清空 底稿_無NT、國外即期、國外換匯、國內即期、國內換匯資料
    ClearRange targetBook, "底稿_無NT"
    ClearRange targetBook, "國外即期"
    ClearRange targetBook, "國外換匯"
    ClearRange targetBook, "國內即期"
    ClearRange targetBook, "國內換匯"
    
    '底稿_無NT

    lastRow = targetBook.Worksheets("底稿_含NT").Cells(Rows.Count, 1).End(xlUp).Row
    Dim destinationRow As Long
    destinationRow = 2
    For i = 2 To lastRow
        If targetBook.Worksheets("底稿_含NT").Cells(i, 13).Value = False Then
            targetBook.Worksheets("底稿_含NT").Rows(i).Copy Destination:=targetBook.Worksheets("底稿_無NT").Rows(destinationRow)
            destinationRow = destinationRow + 1
        End If
    Next i

    '國外即期
    lastRow = targetBook.Worksheets("底稿_無NT").Cells(Rows.Count, 1).End(xlUp).Row
    destinationRow = 2
    For i = 2 To lastRow
        If targetBook.Worksheets("底稿_無NT").Cells(i, 1).Value = "FS" And targetBook.Worksheets("底稿_無NT").Cells(i, 11).Value = "國外" Then
            targetBook.Worksheets("底稿_無NT").Rows(i).Copy Destination:=targetBook.Worksheets("國外即期").Rows(destinationRow)
            destinationRow = destinationRow + 1
        End If
    Next i

    '國外換匯
    destinationRow = 2
    For i = 2 To lastRow
        If targetBook.Worksheets("底稿_無NT").Cells(i, 1).Value = "SS" And targetBook.Worksheets("底稿_無NT").Cells(i, 11).Value = "國外" Then
            targetBook.Worksheets("底稿_無NT").Rows(i).Copy Destination:=targetBook.Worksheets("國外換匯").Rows(destinationRow)
            destinationRow = destinationRow + 1
        End If
    Next i

    '國內即期
    destinationRow = 2
    For i = 2 To lastRow
        If targetBook.Worksheets("底稿_無NT").Cells(i, 1).Value = "FS" And targetBook.Worksheets("底稿_無NT").Cells(i, 11).Value = "國內" Then
            targetBook.Worksheets("底稿_無NT").Rows(i).Copy Destination:=targetBook.Worksheets("國內即期").Rows(destinationRow)
            destinationRow = destinationRow + 1
        End If
    Next i

    '國內換匯
    destinationRow = 2
    For i = 2 To lastRow
        If targetBook.Worksheets("底稿_無NT").Cells(i, 1).Value = "SS" And targetBook.Worksheets("底稿_無NT").Cells(i, 11).Value = "國內" Then
            targetBook.Worksheets("底稿_無NT").Rows(i).Copy Destination:=targetBook.Worksheets("國內換匯").Rows(destinationRow)
            destinationRow = destinationRow + 1
        End If
    Next i
    
    MsgBox "已完成"

End Sub


Sub ClearRange(wb As Workbook, sheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rangeToClear As Range

    ' 定義要清空的工作表名稱
    Set ws = wb.Sheets(sheetName)

    ' 取得最後一行和最後一列的位置
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' 定義要清空的範圍
    Set rangeToClear = ws.Range("A2:M" & lastRow)

    ' 清空範圍內的數值
    rangeToClear.ClearContents
End Sub



' -------------
' dl6850篩選清單_11407_F1F2.xlsm

' 修改後

Sub selectionProcess_DL6850(ByVal wb As Workbook, _
                            ByVal calledByOtherExcel As Boolean, _
                            ByVal baseDatePassed As Variant)
    ' 提示使用者輸入起始日和結束日
    Dim startDate As Date
    Dim endDate As Date

    ' === 修改處（1/3）: 改為要求使用者輸入 年/月（ROC 年，例如 114/09） ===
    Dim ym As String                    ' *** MODIFIED ***
    Dim parts() As String               ' *** MODIFIED ***
    Dim y As Integer, m As Integer      ' *** MODIFIED ***  
    
    If calledByOtherExcel Then
        ym = baseDatePassed
    Else
        ym = InputBox("請輸入報表年月份(格式：YYY/MM，例如 114/09）", "輸入年月 (ROC年/月)")  ' *** MODIFIED ***
    End If
    
    Debug.print "ym: " & ym

    If Trim(ym) = "" Then Exit Sub    

    parts = Split(ym, "/")        ' *** MODIFIED ***

    Debug.print "parts(0): " & parts(0)
    Debug.print "parts(1): " & parts(1)

    If UBound(parts) <> 1 Then
        MsgBox "輸入格式錯誤，請使用 YYY/MM（例如 114/09）", vbExclamation
        Exit Sub
    End If

    y = CInt(parts(0))                  ' *** MODIFIED ***
    m = CInt(parts(1))                  ' *** MODIFIED ***
    
    y = y + 1911                        ' *** MODIFIED ***

    Debug.print "DL6850_y: " & y
    Debug.print "DL6850_m: " & m

    ' startDate = InputBox("請輸入起始日期", "起始日期(日期格式yyyy/mm/dd)：")
    ' endDate = InputBox("請輸入結束日期", "結束日期(日期格式yyyy/mm/dd)：")

    ' 建立本月第一天與本月最後一天
    ' *** MODIFIED ***
    startDate = DateSerial(y, m, 1)
    ' *** MODIFIED ***
    endDate = DateSerial(y, m + 1, 1) - 1
    ' === 修改結束 ===    
    
    '清除底稿NT工作表資料
    wb.Sheets("底稿_含NT_原始資料").Range("A:I").ClearContents
        
    '清除全部交易工作表多餘資料
    Dim ws As Worksheet
    Set ws = wb.Sheets("DL6850全部交易")

    Dim lastRowOrigin As Long
    lastRowOrigin = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    For i = lastRowOrigin To 2 Step -1
        If Left(ws.Cells(i, "A").Value, 2) <> "TR" Then
            ws.Rows(i).Delete
        End If
    Next i
    
    ' 第一個迴圈：清除不在日期範圍內的資料
    Dim lastRow As Long
    Dim j As Long
    
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

    For j = lastRow To 2 Step -1
        If ws.Cells(j, "I").Value < startDate Or ws.Cells(j, "I").Value > endDate Then
            ws.Rows(j).ClearContents
        End If
    Next j

    ' 第二個迴圈：刪除包含空白儲存格的整行
    Dim deleteRows As Range

    For j = lastRow To 2 Step -1
        If IsEmpty(ws.Cells(j, "I")) Then
            If deleteRows Is Nothing Then
                Set deleteRows = ws.Rows(j)
            Else
                Set deleteRows = Union(deleteRows, ws.Rows(j))
            End If
        End If
    Next j

    ' 刪除包含空白儲存格的整行
    If Not deleteRows Is Nothing Then
        deleteRows.Delete
    End If
    
    '複製全部交易工作表至底稿_含NT_原始資料
    
    Dim lastRowSource As Long
    lastRowSource = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim targetSheet As Worksheet
    Set targetSheet = wb.Sheets("底稿_含NT_原始資料")

    ws.Range("A1:I" & lastRowSource).Copy Destination:=targetSheet.Range("A1")

    MsgBox "完成" 

End Sub


' -----------------------
' cm2810篩選清單_11407_F1F2.xlsm
' 修改後

Sub selectionProcess_CM2810(ByVal wb As Workbook)

    '檢查是否存在名稱為「樞紐表」的分頁，若存在則刪除
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets("樞紐表").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
        
    '清除國內顧客工作表資料
    Dim ws As Worksheet
    Set ws = wb.Sheets("CM2810全部交易")

    Dim wsClear As Worksheet
    Set wsClear = wb.Sheets("國內顧客_原始資料")
    
    wb.Sheets("國內顧客_原始資料").Range("A:H").ClearContents
    ws.Range("A1:H1").Copy Destination:=wsClear.Range("A1")

    '處理全部交易資料
    ' 1.清除多餘資料
    Dim lastRowOrigin As Long
    lastRowOrigin = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    For i = lastRowOrigin To 2 Step -1
        If Left(ws.Cells(i, "A").Value, 2) <> "MB" Then
            ws.Rows(i).Delete
        End If
    Next i

    ws.Range("G:V").ClearContents
    ws.Range("G1").Value = "筆數"
    ws.Range("H1").Value = "配對"

    ws.Range("A1").Value = "交易編號"
    ws.Range("B1").Value = "交易/交割日"
    ws.Range("C1").Value = "幣別"
    ws.Range("D1").Value = "交易金額1"
    ws.Range("E1").Value = "本位幣"
    ws.Range("F1").Value = "交易金額2"

    ' 2.排序資料並移動整列
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("A2:A" & lastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange ws.Range("A1:H" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' 3.產生樞紐分析表
    Dim pivotSheet As Worksheet
    Set pivotSheet = wb.Sheets.Add(After:=ws)
    pivotSheet.Name = "樞紐表"

    Dim pivotRange As Range
    Set pivotRange = ws.Range("A1:F" & lastRow)

    Dim pivotTable As pivotTable
    Set pivotTable = pivotSheet.PivotTableWizard(SourceType:=xlDatabase, _
        SourceData:=pivotRange, TableDestination:=pivotSheet.Cells(1, 1), _
        TableName:="樞紐分析表")

    pivotTable.PivotFields("交易編號").Orientation = xlRowField
    
    With pivotTable.PivotFields("幣別")
        .Orientation = xlDataField
        .Function = xlCount
    End With
    
    
    '全部交易工作表中插入公式
    ws.Range("G2").Formula = "=VLOOKUP(A2, 樞紐表!$A:$B, 2, FALSE)"
    ws.Range("H2").Formula = "=CONCATENATE(C2, E2)"
    ws.Range("G2:H2").AutoFill Destination:=ws.Range("G2:H" & lastRow)


    ' 4.複製符合條件的資料到"國內顧客"工作表
    Dim custSheet As Worksheet
    Set custSheet = wb.Sheets("國內顧客_原始資料")

    Dim custRow As Long
    custRow = 2

    Dim fCell As Range
    For Each fCell In ws.Range("G2:G" & lastRow)
        If fCell.Value = 2 Then
            ws.Rows(fCell.Row).Copy Destination:=custSheet.Range("A" & custRow)
            custRow = custRow + 1
        End If
    Next fCell

    MsgBox "完成"
    
End Sub



' ====================
FM2

Question
我要寫一段VBA代碼，要自動化處理程序
1.將資料貼入 OBU_MM4901B分頁
2.OBU_MM4901B中的C欄位為交易對手，我要建立一個資料結構，第一層是所有交易對手，要每行row去跑，新的交易時，新增該筆交易對手及交易的相關紀錄(所有紀錄資訊為A:K欄位)，若已經有儲存該交易對手資料，則將該筆交易紀錄資訊新增到已經儲存的交易對手項下新增一筆紀錄，
' 當跑到第一個交易對象時，將該交易對象名稱及row儲存起來，如果有同樣交易對象出現，就在已經紀錄的交易對象新增那個row，如果是沒有的交易對象，則新增該交易對象及row儲存起來
3.將2資料結構中的交易對象在"金融機構代號對照表"分頁中A欄位逐一檢查，如果是在A欄位中的其中一個，
在原來資料結構中該交易對手紀錄中新增一個歸類分類為DBU的紀錄，並且儲存C欄位的銀行代號，如果是在B欄位中的其中一個，在原來資料結構中該交易對手紀錄中心新增一個歸類分類為OBU，並且儲存C欄位的銀行代號，如果都沒找到，刪除該筆資料，將DBU及OBU分類儲存在資料結構中

4.
DBU 
INDEX1: ROW 3 TO 10
INDEX2: ROW 12 TO 19
INDEX3: ROW 21 TO 28
INDEX4: ROW 30 TO 37
INDEX5: ROW 39 TO 46

OBU
INDEX1: ROW 50 TO 57
INDEX2: ROW 59 TO 66
INDEX3: ROW 68 TO 75
INDEX4: ROW 77 TO 84
INDEX5: ROW 86 TO 93

上述3.中已經分配好的資料結構中，要逐一將結構中儲存的DBU及OBU，
依照所記載在資料結構中的資料，將分類屬於DBU的在 OBU_MM4901B 中紀錄的相關ROWS中的a:k欄位資料按照INDEX順序，一個INDEX代表一個交易對象，貼入DBU所列的相關ROW中
OBU亦同

5.另外按照上述方法，將記錄起來的銀行代號，逐一貼入FM2分頁C10往下的欄位中


Option Explicit

Public Sub ProcessFM2_ButtonClick()
    ProcessFM2 ThisWorkbook
End Sub

Sub ProcessFM2(Optional ByVal wb As Workbook)
    Dim targetBook As Workbook    
    
    Dim wsData As Worksheet
    Dim wsMap As Worksheet
    Dim wsCompute As Worksheet
    Dim wsFM2 As Worksheet

    If wb Is Nothing Then
        Set targetBook = ThisWorkbook
    Else
        Set targetBook = wb
    End If 

    On Error Resume Next
    Set wsData = targetBook.Worksheets("OBU_MM4901B")
    Set wsMap = targetBook.Worksheets("金融機構代號對照表")
    Set wsCompute = targetBook.Worksheets("計算表")
    Set wsFM2 = targetBook.Worksheets("FM2")
    On Error GoTo 0
    If wsData Is Nothing Or wsMap Is Nothing Or wsFM2 Is Nothing Then
        MsgBox "找不到必要的工作表，請確認有 OBU_MM4901B / 金融機構代號對照表 / FM2 三個工作表。", vbExclamation
        Exit Sub
    End If
    
    ' =============================
    ' 【修改處 1】：刪除 K 欄沒有資料的列
    ' =============================
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    Dim r As Long
    For r = lastRow To 2 Step -1  ' 從下往上刪，避免跳行
        If IsEmpty(wsData.Cells(r, "K").Value) Or Trim(wsData.Cells(r, "K").Value) = "" Then
            wsData.Rows(r).Delete
        End If
    Next r
    ' =============================
    
    Dim dict As Object ' key = 交易對手名稱, value = dictionary with collections
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 1) 讀取原始資料（A:K）建立資料結構，key = C欄 (交易對手)
    lastRow = wsData.Cells(wsData.Rows.Count, "C").End(xlUp).Row
    Dim key As Variant
    
    For r = 2 To lastRow
        key = Trim(CStr(wsData.Cells(r, "C").Value))
        If key <> "" Then
            If Not dict.Exists(key) Then
                Dim inner As Object
                Set inner = CreateObject("Scripting.Dictionary")
                inner.Add "Rows", CreateObject("System.Collections.ArrayList") ' 存原始 row numbers (若需要)
                inner.Add "Records", CreateObject("System.Collections.ArrayList") ' 存 A:K 的陣列
                inner.Add "Class", "" ' DBU 或 OBU
                inner.Add "BankCodes", CreateObject("System.Collections.ArrayList") ' 可能多個銀行代號 (但通常一個)
                dict.Add key, inner
            End If
            ' 取得 A:K 的值陣列
            Dim arrAK As Variant
            arrAK = Application.Index(wsData.Range("A" & r & ":K" & r).Value, 1, 0) ' 1D array
            dict(key)("Rows").Add r
            dict(key)("Records").Add arrAK
        End If
    Next r
    
    ' 2) 以 "金融機構代號對照表" 去比對交易對手，決定 DBU/OBU 與抓取銀行代號 (C欄)
    Dim mapLastRow As Long
    mapLastRow = wsMap.Cells(wsMap.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    Dim toRemove As Collection
    Set toRemove = New Collection
    
    Dim nameA As String, nameB As String, bankCode As String
    Dim comp As Long
    For Each key In dict.Keys
        Dim found As Boolean
        found = False
        ' 搜尋 A 欄 (DBU)
        For i = 1 To mapLastRow
            nameA = Trim(CStr(wsMap.Cells(i, "A").Value))
            If nameA <> "" Then
                comp = StrComp(key, nameA, vbTextCompare)
                If comp = 0 Then
                    found = True
                    dict(key)("Class") = "DBU"
                    bankCode = Trim(CStr(wsMap.Cells(i, "C").Value))
                    If bankCode <> "" Then dict(key)("BankCodes").Add bankCode
                    Exit For
                End If
            End If
        Next i
        If Not found Then
            ' 搜尋 B 欄 (OBU)
            For i = 1 To mapLastRow
                nameB = Trim(CStr(wsMap.Cells(i, "B").Value))
                If nameB <> "" Then
                    comp = StrComp(key, nameB, vbTextCompare)
                    If comp = 0 Then
                        found = True
                        dict(key)("Class") = "OBU"
                        bankCode = Trim(CStr(wsMap.Cells(i, "C").Value))
                        If bankCode <> "" Then dict(key)("BankCodes").Add bankCode
                        Exit For
                    End If
                End If
            Next i
        End If
        
        If Not found Then
            ' 若兩欄都沒找到，根據你的要求：刪除該筆資料（即從資料結構中移除，不改原表）
            toRemove.Add key
        End If
    Next key
    
    ' 刪除沒分類（找不到）的交易對手
    Dim idx As Long
    For idx = 1 To toRemove.Count
        dict.Remove toRemove(idx)
    Next idx
    
    ' 3) 準備 DBU / OBU index 區塊 (startRow, endRow) —— 按你給的設定
    Dim dbuBlocks As Variant, obuBlocks As Variant
    dbuBlocks = Array(Array(3, 10), Array(12, 19), Array(21, 28), Array(30, 37), Array(39, 46))
    obuBlocks = Array(Array(50, 57), Array(59, 66), Array(68, 75), Array(77, 84), Array(86, 93))
    
    ' 建立兩個清單，保留原來 dict 中的順序（dictionary iteration 順序為插入順序）
    Dim dbuList As Collection, obuList As Collection
    Set dbuList = New Collection
    Set obuList = New Collection
    For Each key In dict.Keys
        If dict(key)("Class") = "DBU" Then
            dbuList.Add key
        ElseIf dict(key)("Class") = "OBU" Then
            obuList.Add key
        End If
    Next key
    
    ' 4) 將每個交易對象的 A:K 寫入對應的 INDEX 區塊
    ' 注意：一個 INDEX 代表一個交易對象；若該交易對象的紀錄數超過此 INDEX容量，會只貼到該 INDEX 上限，並記錄警告。
    Dim j As Long, recCount As Long
    Dim blk As Variant
    Dim tRow As Long
    Dim writtenCount As Long
    Dim warnings As Collection
    Set warnings = New Collection
    
    ' 先清除目標區塊上的原始 A:K（視需求可保留；這裡先清空以避免殘留）
    Dim b As Long
    For b = LBound(dbuBlocks) To UBound(dbuBlocks)
        wsCompute.Range(wsCompute.Cells(dbuBlocks(b)(0), "A"), wsCompute.Cells(dbuBlocks(b)(1), "K")).ClearContents
    Next b
    For b = LBound(obuBlocks) To UBound(obuBlocks)
        wsCompute.Range(wsCompute.Cells(obuBlocks(b)(0), "A"), wsCompute.Cells(obuBlocks(b)(1), "K")).ClearContents
    Next b
    
    ' DBU 貼入
    For b = LBound(dbuBlocks) To UBound(dbuBlocks)
        If b + 1 > dbuList.Count Then Exit For ' 沒更多交易對手
        key = dbuList(b + 1) ' Collection 是 1-based
        recCount = dict(key)("Records").Count
        blk = dbuBlocks(b)
        writtenCount = 0
        For j = 0 To recCount - 1
            tRow = blk(0) + j
            If tRow <= blk(1) Then
                ' arr is 1-based from Application.Index; write A:K
                Dim recArr As Variant
                recArr = dict(key)("Records")(j)
                Dim colIdx As Long
                For colIdx = 1 To 11 ' A:K = 11 列
                    wsCompute.Cells(tRow, colIdx).Value = recArr(colIdx)
                Next colIdx
                writtenCount = writtenCount + 1
            Else
                ' 超過該 index 能放的列：跳出並記警告
                warnings.Add "DBU '" & key & "' 的紀錄數 (" & recCount & ") 超過 index" & (b + 1) & " 容量（" & (blk(1) - blk(0) + 1) & "），僅貼入前 " & writtenCount & " 筆。"
                Exit For
            End If
        Next j
    Next b
    
    ' OBU 貼入
    For b = LBound(obuBlocks) To UBound(obuBlocks)
        If b + 1 > obuList.Count Then Exit For
        key = obuList(b + 1)
        recCount = dict(key)("Records").Count
        blk = obuBlocks(b)
        writtenCount = 0
        For j = 0 To recCount - 1
            tRow = blk(0) + j
            If tRow <= blk(1) Then
                recArr = dict(key)("Records")(j)
                For colIdx = 1 To 11
                    wsCompute.Cells(tRow, colIdx).Value = recArr(colIdx)
                Next colIdx
                writtenCount = writtenCount + 1
            Else
                warnings.Add "OBU '" & key & "' 的紀錄數 (" & recCount & ") 超過 index" & (b + 1) & " 容量（" & (blk(1) - blk(0) + 1) & "），僅貼入前 " & writtenCount & " 筆。"
                Exit For
            End If
        Next j
    Next b
    
    ' 5) 將已記錄的銀行代號（不重複）貼入 FM2 C10 開始往下
    Dim bankSet As Object
    Set bankSet = CreateObject("Scripting.Dictionary")
    ' 依照插入順序收集 bank codes（先 DBU 再 OBU；也可以 alter）
    For idx = 1 To dbuList.Count
        key = dbuList(idx)
        Dim bcArr As Object
        Set bcArr = dict(key)("BankCodes")
        For j = 0 To bcArr.Count - 1
            If Not bankSet.Exists(bcArr(j)) Then bankSet.Add bcArr(j), 1
        Next j
    Next idx
    
    For idx = 1 To obuList.Count
        key = obuList(idx)
        Set bcArr = dict(key)("BankCodes")
        For j = 0 To bcArr.Count - 1
            If Not bankSet.Exists(bcArr(j)) Then bankSet.Add bcArr(j), 1
        Next j
    Next idx
    
    ' 清空 FM2 C10 往下一段（可視需求調整）
    Dim startFMrow As Long
    startFMrow = 10
    wsFM2.Range(wsFM2.Cells(startFMrow, "C"), wsFM2.Cells(wsFM2.Rows.Count, "C")).ClearContents
    
    Dim outRow As Long
    outRow = startFMrow
    Dim kKey As Variant
    For Each kKey In bankSet.Keys
        wsFM2.Cells(outRow, "C").Value = kKey
        outRow = outRow + 1
    Next kKey
    
    ' 顯示處理結果與警告
    Dim msg As String
    msg = "處理完成。" & vbCrLf & "DBU 數量: " & dbuList.Count & "，OBU 數量: " & obuList.Count & "。" & vbCrLf & "已將銀行代號貼至 FM2 C" & startFMrow & " 開始的欄位。"
    If warnings.Count > 0 Then
        msg = msg & vbCrLf & vbCrLf & "注意：" & vbCrLf
        For idx = 1 To warnings.Count
            msg = msg & "- " & warnings(idx) & vbCrLf
        Next idx
    End If
    MsgBox msg, vbInformation, "OBU 處理結果"
    
End Sub




' ==============

FM10

Option Explicit

Public Sub CopyAndDeleteRows_ButtonClick()
    CopyAndDeleteRows ThisWorkbook
End Sub

Sub CopyAndDeleteRows(Optional ByVal wb As Workbook)
    Dim targetBook As Workbook    
    Dim wsAC4603 As Worksheet
    Dim wsFM10 As Worksheet
    Dim n As Long
    Dim count As Long

    If wb Is Nothing Then
        Set targetBook = ThisWorkbook
    Else
        Set targetBook = wb
    End If    
    
    '<若欄位異動需更改項目>
    'AC4603檢核總欄位數，若AC4603欄位異動，確認所需欄位個數「更改下列count數值」
    count = 26

    ' 設定工作表名稱
    Set wsAC4603 = targetBook.Sheets("OBU_AC4603")
    Set wsFM10 = targetBook.Sheets("FM10底稿")

    ' 找到第n行的位置
    n = Application.Match("強制FVPL金融資產-公債-地方政府(外國)", wsAC4603.Range("A:A"), 0)

    ' 檢查是否找到了第n行
    If Not IsError(n) Then
        ' 檢查條件是否成立
        
        '---------------------------------------------
        '<若欄位異動需更改項目>
        '若欄位數異動，更改以下需檢核之欄位，欄位名稱需與報表名稱完全一致
        If wsAC4603.Range("A" & n + 1).Value = "強制FVPL金融資產-普通公司債(民營)(外國)" And _
           wsAC4603.Range("A" & n + 2).Value = "12005" And _
           wsAC4603.Range("A" & n + 3).Value = "強制FVPL金融資產評價調整-公債-地方-外國" And _
           wsAC4603.Range("A" & n + 4).Value = "強制FVPL金融資產評價調整-普通公司債(民營)(外國)" And _
           wsAC4603.Range("A" & n + 5).Value = "12007" And _
           wsAC4603.Range("A" & n + 6).Value = "FVOCI債務工具-公債-中央政府(外國)" And _
           wsAC4603.Range("A" & n + 7).Value = "FVOCI債務工具-普通公司債(公營)(外國)" And _
           wsAC4603.Range("A" & n + 8).Value = "FVOCI債務工具-普通公司債(民營)(外國)" And _
           wsAC4603.Range("A" & n + 9).Value = "FVOCI債務工具-金融債券-海外" And _
           wsAC4603.Range("A" & n + 10).Value = "12111" And _
           wsAC4603.Range("A" & n + 11).Value = "FVOCI債務工具評價調整-公債-中央政府(外國)" And _
           wsAC4603.Range("A" & n + 12).Value = "FVOCI債務工具評價調整-普通公司債(公營)(外國)" And _
           wsAC4603.Range("A" & n + 13).Value = "FVOCI債務工具評價調整-普通公司債(民營)(外國)" And _
           wsAC4603.Range("A" & n + 14).Value = "FVOCI債務工具評價調整-金融債券-海外" And _
           wsAC4603.Range("A" & n + 15).Value = "12113" And _
           wsAC4603.Range("A" & n + 16).Value = "AC債務工具投資-公債-中央政府(外國)" And _
           wsAC4603.Range("A" & n + 17).Value = "AC債務工具投資-普通公司債(民營)(外國)" And _
           wsAC4603.Range("A" & n + 18).Value = "AC債務工具投資-金融債券-海外" And _
           wsAC4603.Range("A" & n + 19).Value = "12201" And _
           wsAC4603.Range("A" & n + 20).Value = "累積減損-AC債務工具投資-公債-中央政府(外國)" And _
           wsAC4603.Range("A" & n + 21).Value = "累積減損-AC債務工具投資-普通公司(民營)(外國)" And _
           wsAC4603.Range("A" & n + 22).Value = "累積減損-AC債務工具投資-金融債券-海外" And _
           wsAC4603.Range("A" & n + 23).Value = "12203" And _
           wsAC4603.Range("A" & n + 24).Value = "拆放證券公司-OSU" And _
           wsAC4603.Range("A" & n + 25).Value = "15551" Then

            ' 刪除第n+count行至最後一行
            wsAC4603.Rows(n + count & ":" & wsAC4603.Rows.count).Delete

            ' 刪除第一行至第n-1行
            wsAC4603.Rows("1:" & n - 1).Delete
            
            '清除FM10底稿checkBox資料
            wsFM10.Range("A4:J" & (4 + count - 1)).ClearContents
            Application.CutCopyMode = False

            ' 複製AC4603數值內容到FM10底稿checkBox
            wsAC4603.Range("A1:J" & count).Copy
            wsFM10.Range("A4").Resize(count, 10).PasteSpecial Paste:=xlPasteValues
            
            MsgBox "完成"
            
            
        Else
            MsgBox "欄位有誤"
        End If
    Else
        MsgBox "找不到目標欄位 'FVOCI債務工具-公債-中央政府(外國)'"
    End If
End Sub



' ==================

FM11

' 修改前

Sub 匯入並篩選OBUAC5411B資料()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim importWB As Workbook
    Dim importFile As String
    Dim lastRow As Long, destRow As Long
    Dim cell As Range
    Dim keywords As Variant
    Dim i As Long
    Dim sumRange As Range

    ' 選取檔案
    importFile = Application.GetOpenFilename("Excel Files (*.xls;*.xlsx), *.xls;*.xlsx", , "請選取 OBU-AC5411B 檔案")
    If importFile = "False" Then Exit Sub ' 使用者取消

    ' 開啟來源檔案
    Set importWB = Workbooks.Open(importFile)
    
    ' 檢查是否存在名為 OBU-AC5411B 的分頁
    On Error Resume Next
    Set wsSource = importWB.Sheets(1)
    On Error GoTo 0
    If wsSource Is Nothing Then
        MsgBox "來源檔案中找不到分頁『OBU-AC5411B』", vbExclamation
        importWB.Close False
        Exit Sub
    End If

    ' 清除目前工作簿的 OBU-AC5411B 分頁舊資料（從第2列開始）
    Set wsDest = ThisWorkbook.Sheets("OBU-AC5411B")
    wsDest.Range("A2:Z" & wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row).ClearContents

    ' 複製來源檔案中第2列起資料貼到目前檔案中（保留標題列）
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        wsSource.Range("A2:Z" & lastRow).Copy Destination:=wsDest.Range("A2")
    End If
    
    ' 將 B 欄強制轉換成數值格式（避免 VLOOKUP 比對不到）
    With wsDest
        With .Range("B2:B" & .Cells(.Rows.Count, "B").End(xlUp).Row)
            .NumberFormat = "0"  ' 設定格式為一般數字
            .Value = .Value      ' 將文字轉換為純數值
        End With
    End With
    
    ' 關閉來源檔案
    importWB.Close False

    ' ---------- 以下是篩選與統計程式 ----------
    
    keywords = Array("FVPL", "FVOCI", "AC", "拆放證券公司息-OSU")

    ' 若目標工作表存在就刪除重建
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("OBU-AC5411B會科整理").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("OBU-AC5411B"))
    wsDest.Name = "OBU-AC5411B會科整理"

    ' 複製標題列
    ThisWorkbook.Sheets("OBU-AC5411B").Rows(1).Copy Destination:=wsDest.Rows(1)
    destRow = 2

    ' 遍歷 A 欄，找出符合關鍵字的列
    lastRow = ThisWorkbook.Sheets("OBU-AC5411B").Cells(ThisWorkbook.Sheets("OBU-AC5411B").Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        For Each keyword In keywords
            If InStr(ThisWorkbook.Sheets("OBU-AC5411B").Cells(i, "A").Value, keyword) > 0 Then
                ThisWorkbook.Sheets("OBU-AC5411B").Rows(i).Copy Destination:=wsDest.Rows(destRow)
                destRow = destRow + 1
                Exit For
            End If
        Next keyword
    Next i

    ' 總和 C 欄
    If destRow > 2 Then
        wsDest.Cells(destRow, "B").Value = "本月金額總和"
        Set sumRange = wsDest.Range("C2:C" & destRow - 1)
        wsDest.Cells(destRow, "C").Formula = "=SUM(" & sumRange.Address(False, False) & ")"
    
            ' 複製結果值到 FM11 計算1 的 G4 欄位（只取值，不取公式）
        ThisWorkbook.Sheets("FM11 計算1").Range("G4").Value = wsDest.Cells(destRow, "C").Value
    
    End If

    ' 自動欄寬
    wsDest.Columns.AutoFit

    MsgBox "匯入成功並完成篩選與總和計算！", vbInformation
    
    ThisWorkbook.Sheets("FM11 計算1").Activate
End Sub

' 修改後

Public Sub 匯入並篩選OBUAC5411B資料_ButtonClick()
    匯入並篩選OBUAC5411B資料 ThisWorkbook
End Sub

Sub 匯入並篩選OBUAC5411B資料(Optional ByVal wb As Workbook, _
                            Optional ByVal calledByOtherExcel As Boolean = False)
    Dim targetBook As Workbook                            
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim importWB As Workbook
    Dim importFile As String
    Dim lastRow As Long, destRow As Long
    Dim keyword As Variant    
    Dim keywords As Variant
    Dim i As Long
    Dim sumRange As Range


    If wb Is Nothing Then
        Set targetBook = ThisWorkbook
    Else
        Set targetBook = wb
    End If


    ' ### 修改開始 ###
    If Not calledByOtherExcel Then
        ' 選取檔案
        importFile = Application.GetOpenFilename("Excel Files (*.xls;*.xlsx), *.xls;*.xlsx", , "請選取 OBU-AC5411B 檔案")
        If importFile = "False" Then Exit Sub ' 使用者取消

        ' 開啟來源檔案
        Set importWB = Workbooks.Open(importFile)
        
        ' 檢查是否存在名為 OBU-AC5411B 的分頁
        On Error Resume Next
        Set wsSource = importWB.Sheets(1)
        On Error GoTo 0
        If wsSource Is Nothing Then
            MsgBox "來源檔案中找不到分頁『OBU-AC5411B』", vbExclamation
            importWB.Close False
            Exit Sub
        End If

        ' 清除目前工作簿的 OBU-AC5411B 分頁舊資料（從第2列開始）
        Set wsDest = targetBook.Sheets("OBU-AC5411B")
        wsDest.Range("A2:Z" & wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row).ClearContents

        ' 複製來源檔案中第2列起資料貼到目前檔案中（保留標題列）
        lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
        If lastRow >= 2 Then
            wsSource.Range("A2:Z" & lastRow).Copy Destination:=wsDest.Range("A2")
        End If
        
        ' 將 B 欄強制轉換成數值格式（避免 VLOOKUP 比對不到）
        With wsDest
            With .Range("B2:B" & .Cells(.Rows.Count, "B").End(xlUp).Row)
                .NumberFormat = "0"  ' 設定格式為一般數字
                .Value = .Value      ' 將文字轉換為純數值
            End With
        End With
        
        ' 關閉來源檔案
        importWB.Close False
    Else
        ' 若是被其他 Excel 呼叫，記錄 debug 訊息（可選）
        Debug.Print "匯入並篩選OBUAC5411B資料：由其他 Excel 呼叫，已略過選檔/開檔/複製動作。"
        ' 注意：被呼叫時需確保 ThisWorkbook 的 "OBU-AC5411B" 工作表已經有資料（若無資料則後續篩選會找不到東西)
    End If
    ' ### 修改結束 ###

    ' ---------- 以下是篩選與統計程式 ----------
    
    keywords = Array("FVPL", "FVOCI", "AC", "拆放證券公司息-OSU")

    ' 若目標工作表存在就刪除重建
    On Error Resume Next
    Application.DisplayAlerts = False
    targetBook.Worksheets("OBU-AC5411B會科整理").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsDest = targetBook.Sheets.Add(After:=targetBook.Sheets("OBU-AC5411B"))
    wsDest.Name = "OBU-AC5411B會科整理"

    ' 複製標題列
    targetBook.Sheets("OBU-AC5411B").Rows(1).Copy Destination:=wsDest.Rows(1)
    destRow = 2

    ' 遍歷 A 欄，找出符合關鍵字的列
    lastRow = targetBook.Sheets("OBU-AC5411B").Cells(targetBook.Sheets("OBU-AC5411B").Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        For Each keyword In keywords
            If InStr(targetBook.Sheets("OBU-AC5411B").Cells(i, "A").Value, keyword) > 0 Then
                targetBook.Sheets("OBU-AC5411B").Rows(i).Copy Destination:=wsDest.Rows(destRow)
                destRow = destRow + 1
                Exit For
            End If
        Next keyword
    Next i

    ' 總和 C 欄
    If destRow > 2 Then
        wsDest.Cells(destRow, "B").Value = "本月金額總和"
        Set sumRange = wsDest.Range("C2:C" & destRow - 1)
        wsDest.Cells(destRow, "C").Formula = "=SUM(" & sumRange.Address(False, False) & ")"
    
            ' 複製結果值到 FM11 計算1 的 G4 欄位（只取值，不取公式）
        targetBook.Sheets("FM11 計算1").Range("G4").Value = wsDest.Cells(destRow, "C").Value
    
    End If

    ' 自動欄寬
    wsDest.Columns.AutoFit

    MsgBox "匯入成功並完成篩選與總和計算！", vbInformation    
    targetBook.Sheets("FM11 計算1").Activate
End Sub


' ===========

表41

' 修改前

Option Explicit

Sub SortAndCopyData()

    Dim wsDL9360 As Worksheet
    Dim wsTarget As Worksheet
    Dim baseDate As Date
    Dim exchangeRate As Double
    Dim lastRow As Long
    Dim n As Long
    Dim m As Long
    
    '設定工作表
    Set wsDL9360 = ThisWorkbook.Sheets("DL9360")
    Set wsTarget = ThisWorkbook.Sheets("底稿(扣掉TWD)")

    '彈出視窗，填寫基準日及美元兌換匯率
    baseDate = InputBox("請輸入基準日(日期格式：yyyy/mm/dd)", "基準日")
    exchangeRate = InputBox("請輸入美元兌換匯率", "美元兌換匯率")
    wsTarget.Range("C66").Value = baseDate
    wsTarget.Range("E66").Value = exchangeRate
    
    
    '刪除B欄位非日期格式及刪除國內交易對手(銀行國際代碼末4碼非TWTP)之交易
    lastRow = wsDL9360.Cells(wsDL9360.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = lastRow To 2 Step -1
        If Not IsDate(wsDL9360.Cells(i, "B").Value) Or Mid(wsDL9360.Cells(i, "E").Value, 5, 2) = "TW" Then
            wsDL9360.Rows(i).Delete
        End If
    Next i
    
    '清除 底稿(扣掉TWD) 工作表資料
    wsTarget.Range("A2:K31").ClearContents
    wsTarget.Range("A33:K62").ClearContents
    
    
    '重新取得欄位數
    lastRow = wsDL9360.Cells(wsDL9360.Rows.Count, "A").End(xlUp).Row

    ' 排序J欄位
    wsDL9360.Range("A1:K1").CurrentRegion.Sort Key1:=wsDL9360.Range("J2"), Order1:=xlAscending, Header:=xlYes


    ' 尋找小於零的最後一筆資料
    For n = 2 To lastRow
        If wsDL9360.Cells(n, "J").Value >= 0 Then
            m = lastRow - n + 1
            '當處分利益或損失任一交易筆數超過30筆時，中斷執行
            If n > 31 Or m > 30 Then
                MsgBox "筆數太多"
                Exit Sub
            Else
                Exit For
            End If
        End If
    Next n
    
    
    
    ' 複製資料至底稿工作表
    wsDL9360.Range("A2:K" & n - 1).Copy
    wsTarget.Range("A2").PasteSpecial Paste:=xlPasteValues

    wsDL9360.Range("A" & n & ":K" & lastRow).Copy
    wsTarget.Range("A33").PasteSpecial Paste:=xlPasteValues
  

    ' 清除剪貼板
    Application.CutCopyMode = False

    MsgBox "完成"
    
End Sub

' 修改後

Option Explicit

Public Sub SortAndCopyData_ButtonClick()
    SortAndCopyData ThisWorkbook
End Sub

Sub SortAndCopyData(Optional ByVal wb As Workbook, _
                    Optional ByVal calledByOtherExcel As Boolean = False, _
                    Optional ByVal baseDatePassed As Variant = Empty)
    Dim targetBook As Workbook    
    Dim wsDL9360 As Worksheet
    Dim wsTarget As Worksheet
    Dim baseDate As Date
    Dim exchangeRate As Double
    Dim lastRow As Long
    Dim n As Long
    Dim m As Long

    If wb Is Nothing Then
        Set targetBook = ThisWorkbook
    Else
        Set targetBook = wb
    End If    

    '設定工作表
    Set wsDL9360 = targetBook.Sheets("DL9360")
    Set wsTarget = targetBook.Sheets("底稿(扣掉TWD)")

    If calledByOtherExcel Then
        baseDate = baseDatePassed
    Else
        baseDate = InputBox("請輸入基準日(日期格式：yyyy/mm/dd)", "基準日")
        exchangeRate = InputBox("請輸入美元兌換匯率", "美元兌換匯率")
    End If    


    '彈出視窗，填寫基準日及美元兌換匯率

    wsTarget.Range("C66").Value = baseDate
    
    '刪除B欄位非日期格式及刪除國內交易對手(銀行國際代碼末4碼非TWTP)之交易
    lastRow = wsDL9360.Cells(wsDL9360.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = lastRow To 2 Step -1
        If Not IsDate(wsDL9360.Cells(i, "B").Value) Or Mid(wsDL9360.Cells(i, "E").Value, 5, 2) = "TW" Then
            wsDL9360.Rows(i).Delete
        End If
    Next i
    
    '清除 底稿(扣掉TWD) 工作表資料
    wsTarget.Range("A2:K31").ClearContents
    wsTarget.Range("A33:K62").ClearContents
    
    '重新取得欄位數
    lastRow = wsDL9360.Cells(wsDL9360.Rows.Count, "A").End(xlUp).Row

    ' 排序J欄位
    wsDL9360.Range("A1:K1").CurrentRegion.Sort Key1:=wsDL9360.Range("J2"), Order1:=xlAscending, Header:=xlYes

    ' 尋找小於零的最後一筆資料
    For n = 2 To lastRow
        If wsDL9360.Cells(n, "J").Value >= 0 Then
            m = lastRow - n + 1
            '當處分利益或損失任一交易筆數超過30筆時，中斷執行
            If n > 31 Or m > 30 Then
                MsgBox "筆數太多"
                Exit Sub
            Else
                Exit For
            End If
        End If
    Next n
        
    ' 複製資料至底稿工作表
    wsDL9360.Range("A2:K" & n - 1).Copy
    wsTarget.Range("A2").PasteSpecial Paste:=xlPasteValues

    wsDL9360.Range("A" & n & ":K" & lastRow).Copy
    wsTarget.Range("A33").PasteSpecial Paste:=xlPasteValues
  

    ' 清除剪貼板
    Application.CutCopyMode = False

    MsgBox "完成"
    
End Sub


' ===================
