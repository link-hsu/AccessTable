下面的「子腳本超出範圍（Subscript out of range）」其實不是 Excel 讀不到儲存格，而是你的 outputArr 根本還沒被正確 ReDim，就直接用 outputArr(j,k) 去寫值，所以 VBA 跳出「陣列還沒配好尺寸」的錯誤。

⸻

一、為什麼會跳過 ReDim，導致陣列沒配好？

For i = 1 To securityRows.count
    If i + 1 <= securityRows.count Then
        If securityRows(i) + 1 = securityRows(i + 1) Then
            GoTo ContinueLoop    ' ← 如果第一組 securityRows(1)+1 = securityRows(2)，就跳到這裡
        Else
            startRow = …
            endRow   = …
        End If
    Else
        startRow = …
        endRow   = …
    End If

    If i = 1 Then
        ReDim outputArr(1 To lastRow, 1 To 32)   ' ← 這行只在 i=1 而且沒被 GoTo 跳過時執行
    End If

    ’… 接著才開始 outputArr(j,k) = …
ContinueLoop:
Next i

	•	當 securityRows(1)+1 = securityRows(2) 時，程式一進到第一圈就執行 GoTo ContinueLoop，直接跳到 Next i，從來沒執行到 If i=1 Then ReDim …。
	•	接著 i=2,3… 時，If i=1 Then 條件不符，也不會 ReDim，結果 outputArr 一直是未配置狀態，第一個 outputArr(j,k)=… 就炸掉。

⸻

二、修正方法

方法一：把 ReDim 搬到迴圈外，一開始就配置好

' 在 For i = 1 To securityRows.Count 之前
ReDim outputArr(1 To lastRow, 1 To 32)

For i = 1 To securityRows.Count
    ’… 不用在這裡 ReDim 了
    ’… 其餘邏輯不變
Next i

方法二：將 ReDim 提前到任何 GoTo 之前

For i = 1 To securityRows.Count

    ’— 一進來就保證做一次 ReDim（只有 i=1 時才做）
    If i = 1 Then
        ReDim outputArr(1 To lastRow, 1 To 32)
        WriteLog "已 ReDim outputArr: rows=1–" & lastRow & ", cols=1–32"
    End If

    If i + 1 <= securityRows.Count Then
        If securityRows(i) + 1 = securityRows(i + 1) Then
            GoTo ContinueLoop
        Else
            startRow = …
            endRow   = …
        End If
    Else
        startRow = …
        endRow   = …
    End If

    ’… 接著才是 j, k 迴圈
ContinueLoop:
Next i



⸻

三、如何用更完整的 Log 追蹤邏輯
	1.	在 ReDim 之後立刻記錄

If i = 1 Then
    ReDim outputArr(1 To lastRow, 1 To 32)
    WriteLog "【ReDim 完成】outputArr LBound1=" & LBound(outputArr,1) & ", UBound1=" & UBound(outputArr,1) & _
             "; LBound2=" & LBound(outputArr,2) & ", UBound2=" & UBound(outputArr,2)
End If


	2.	在每次寫入前，記錄陣列上下界

WriteLog "準備寫入 outputArr，j=" & j & ", k=" & k & _
         "; RowBound=1–" & UBound(outputArr,1) & ", ColBound=1–" & UBound(outputArr,2)
outputArr(j, k) = …


	3.	加上錯誤處理段，捕捉第一個錯誤並印出詳細資訊

On Error Resume Next
    outputArr(j, k) = xlsht.Cells(j, k).Value
    If Err.Number <> 0 Then
        WriteLog "### 錯誤 #" & Err.Number & ": " & Err.Description & _
                 " at outputArr(" & j & "," & k & ")" & _
                 "; LBound1=" & LBound(outputArr,1) & ", UBound1=" & UBound(outputArr,1) & _
                 "; LBound2=" & LBound(outputArr,2) & ", UBound2=" & UBound(outputArr,2)
        Err.Clear
    End If
On Error GoTo 0


	4.	（選用）將 GoTo 改成 Continue For，結構更清晰

For i = 1 To securityRows.Count
    If i+1 <= securityRows.Count And securityRows(i)+1 = securityRows(i+1) Then
        WriteLog "第 " & i & " 組 securityRows 連續，跳過"
        Continue For    ’VBA 2010+ 支援
    End If
    ’… 其餘處理
Next i



⸻

重點結論
	•	核心問題：ReDim outputArr(...) 被放在可能被 GoTo ContinueLoop 跳過的區塊，導致陣列根本沒配置就使用。
	•	最簡修法：把那行 ReDim 移出任何會被跳過的 If/GoTo 區塊，一開始就做一次。
	•	強烈建議：加上「印出 LBound/UBound」的 Log，以後一旦再碰到「Subscript out of range」就能馬上看出是哪個維度、哪個邊界不符。

如此一來，陣列一定會先正確配置，再進行 outputArr(j,k)=…，錯誤就會消失；而且未來若有問題，Log 裡的上下界資訊也能立即告訴你，哪一維、哪個索引超出範圍，非常好 debug。















==========================





Public Sub FXCleanReport(ByVal fullFilePath As String, _
                         ByVal cleaningType As String)

    Dim xlApp As Excel.Application
    Dim xlbk As Excel.Workbook
    Dim ws As Excel.Worksheet
    Dim xlsht As Excel.Worksheet
    Dim copyRg As Range
    Dim Rngs As Range
    Dim oneRng As Range
    Dim outputArr() As Variant
    Dim fvArray As Variant
    Dim mapGroupMeasurement As Object
    Dim groupMeasurement As Variant
    Dim i As Long, j As Long, k As Long
    Dim lastRow As Long
    Dim securityRows As Collection
    Dim category As Variant
    Dim columnsArray() As Variant
    Dim tempSplit As Variant
    Dim tempSave() As Variant
    Dim count As Long
    Dim splitCount As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim numRows As Long
    Dim numCols As Long

    '— 建立 Excel 物件並開檔
    Set xlApp = Excel.Application
    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    Set ws = xlbk.Worksheets("評估表")

    '— 複製工作表，操作複本
    ws.Copy After:=xlbk.Sheets(xlbk.Sheets.Count)
    ActiveSheet.Name = "評估表cp"
    Set xlsht = ActiveSheet

    '— 解除公式，改成值
    Set copyRg = xlsht.UsedRange
    copyRg.Value = copyRg.Value

    '— 取得欄位標題 A5:T5，拆解換行符
    Set Rngs = copyRg.Range("A5:T5")
    count = 0: splitCount = 0
    For Each oneRng In Rngs
        tempSplit = Split(oneRng.Value, vbLf)
        ReDim Preserve columnsArray(count)
        columnsArray(count) = Trim(tempSplit(0))
        count = count + 1
        If UBound(tempSplit) >= 1 Then
            ReDim Preserve tempSave(splitCount)
            tempSave(splitCount) = Trim(tempSplit(1))
            splitCount = splitCount + 1
        End If
    Next oneRng
    ReDim Preserve tempSave(splitCount)
    tempSave(splitCount) = "評價資產類別"
    For i = LBound(tempSave) To UBound(tempSave)
        ReDim Preserve columnsArray(count)
        columnsArray(count) = tempSave(i)
        count = count + 1
    Next i

    '— fvArray & groupMeasurement 映射
    fvArray = Array("FVPL-公債", "FVPL-公司債(公營)", "FVPL-公司債(民營)", "FVPL-金融債", _
                    "FVOCI-公債", "FVOCI-公司債(公營)", "FVOCI-公司債(民營)", "FVOCI-金融債", _
                    "AC-公債", "AC-公司債(公營)", "AC-公司債(民營)", "AC-金融債")
    groupMeasurement = Array("FVPL_GovBond_Foreign", "FVPL_CompanyBond_Foreign", "FVPL_CompanyBond_Foreign", "FVPL_FinancialBond_Foreign", _
                             "FVOCI_GovBond_Foreign", "FVOCI_CompanyBond_Foreign", "FVOCI_CompanyBond_Foreign", "FVOCI_FinancialBond_Foreign", _
                             "AC_GovBond_Foreign", "AC_CompanyBond_Foreign", "AC_CompanyBond_Foreign", "AC_FinancialBond_Foreign")
    Set mapGroupMeasurement = CreateObject("Scripting.Dictionary")
    For i = LBound(fvArray) To UBound(fvArray)
        mapGroupMeasurement.Add fvArray(i), groupMeasurement(i)
    Next i

    '— 刪除空白列與標註區
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then
        For i = lastRow To 1 Step -1
            If IsEmpty(xlsht.Cells(i, 1).Value) _
               Or xlsht.Cells(i, 1).Value = "Security_Id" Then
                xlsht.Rows(i).Delete
            ElseIf Left(Trim(xlsht.Cells(i, 1).Value), 2) = "標註" Then
                xlsht.Rows(i & ":" & lastRow).Delete
                Exit For
            End If
        Next i
    End If

    '— 尋找每個 category 起始列
    lastRow = xlsht.Cells(xlsht.Rows.Count, 1).End(xlUp).Row
    Set securityRows = New Collection
    For i = 1 To lastRow
        For Each category In fvArray
            If xlsht.Cells(i, 1).Value = category Then
                securityRows.Add i
            End If
        Next category
    Next i

    '— 重要：先一次性 ReDim outputArr，避免後面 Subscript out of range
    ReDim outputArr(1 To lastRow, 1 To 32)
    WriteLog "【ReDim 完成】outputArr LBound1=1, UBound1=" & lastRow & _
             "; LBound2=1, UBound2=32"

    '— 逐組區段，讀兩列合併資訊到 outputArr
    For i = 1 To securityRows.Count
        ' 如果和下一 category 連續就跳過處理
        If i + 1 <= securityRows.Count _
           And securityRows(i) + 1 = securityRows(i + 1) Then
            WriteLog "第 " & i & " 組 securityRows 連續，跳過"
            Continue For
        End If

        ' 決定本組資料起訖列
        If i + 1 <= securityRows.Count Then
            startRow = securityRows(i) + 1
            endRow = securityRows(i + 1) - 1
        Else
            startRow = securityRows(i) + 1
            endRow = lastRow
        End If

        category = xlsht.Cells(startRow - 1, 1).Value

        For j = startRow To endRow Step 2
            For k = 1 To 32
                '— 事前 log 陣列上下界
                WriteLog "準備寫入 outputArr，j=" & j & ", k=" & k & _
                         "; RowBound=1–" & UBound(outputArr, 1) & _
                         ", ColBound=1–" & UBound(outputArr, 2)
                On Error Resume Next
                    Select Case k
                        Case 1 To 20
                            outputArr(j, k) = xlsht.Cells(j, k).Value
                            ' AC 系列，把第17欄取到第20欄
                            If Left(category, 2) = "AC" And k <= 20 Then
                                outputArr(j, 20) = xlsht.Cells(j, 17).Value
                            End If
                        Case 21
                            outputArr(j, 21) = xlsht.Cells(j + 1, k - 20).Value   ' Issuer
                        Case 22
                            outputArr(j, 22) = xlsht.Cells(j + 1, k - 20).Value   ' Avg_Txnt_Rate
                        Case 23
                            outputArr(j, 23) = xlsht.Cells(j + 1, k - 20).Value   ' Avg_Buy_Price
                        Case 24
                            outputArr(j, 24) = xlsht.Cells(j + 1, k - 20).Value   ' Tot_Nominal_Amt_USD
                        Case 25
                            outputArr(j, 25) = xlsht.Cells(j + 1, k - 20).Value   ' Book_Value
                        Case 26
                            outputArr(j, 26) = xlsht.Cells(j + 1, k - 20).Value   ' PL_Amt_USD
                        Case 27
                            outputArr(j, 27) = xlsht.Cells(j + 1, k - 20).Value   ' Amortize_Amt
                        Case 28
                            outputArr(j, 28) = xlsht.Cells(j + 1, k - 20).Value   ' DVO1_USD
                        Case 29
                            outputArr(j, 29) = xlsht.Cells(j + 1, k - 20).Value   ' Interest_receivable_USD
                        Case 30
                            outputArr(j, 30) = xlsht.Cells(j + 1, k - 20).Value   ' 當日評等
                        Case 31
                            outputArr(j, 31) = category                           ' 評價類別
                        Case 32
                            If mapGroupMeasurement.Exists(category) Then
                                outputArr(j, 32) = mapGroupMeasurement(category)  ' groupMeasurement
                            Else
                                outputArr(j, 32) = ""
                            End If
                    End Select
                    If Err.Number <> 0 Then
                        WriteLog "### 錯誤 #" & Err.Number & ": " & Err.Description & _
                                 " at outputArr(" & j & "," & k & ")"
                        Err.Clear
                    End If
                On Error GoTo 0
            Next k
        Next j
    Next i

    '— 建立 OutputData 工作表並貼上 outputArr
    xlbk.Sheets.Add After:=xlbk.Sheets(xlbk.Sheets.Count)
    ActiveSheet.Name = "OutputData"
    numRows = UBound(outputArr, 1)
    numCols = UBound(outputArr, 2)
    ActiveSheet.Range(ActiveSheet.Cells(2, 1), _
                      ActiveSheet.Cells(numRows + 1, numCols)).Value = outputArr

    '— 刪除空白列
    For i = numRows + 1 To 2 Step -1
        If ActiveSheet.Cells(i, 1).Value = "" Then ActiveSheet.Rows(i).Delete
    Next i

    '— 寫入欄位標題
    ActiveSheet.Range("A1").Resize(1, UBound(columnsArray) + 1).Value = columnsArray

    '— 移除其他所有工作表
    For Each ws In xlbk.Sheets
        If Not ws Is ActiveSheet Then ws.Delete
    Next ws

    xlbk.Save
    xlbk.Close False

    Set xlsht = Nothing
    Set xlbk = Nothing
    Set xlApp = Nothing

End Sub

說明重點
	1.	一次性 ReDim 移到所有迴圈之前，避免被 GoTo 跳過。
	2.	在 ReDim 後與每次寫入前，都用 WriteLog 印出 LBound/UBound，如果未來再遇到範圍錯誤，一眼就能看出是哪個維度不對。
	3.	加了 On Error … 段，當真的有超出或其他錯誤時，會把錯誤編號與描述都寫到 Log，並標明出錯的 j,k。

這樣就能完整避免「陣列索引超出範圍」問題，並且大幅提升日後追蹤與除錯的效率。
