Public Function GetFieldFromXlRanges(ByVal rptSheetName   As String, _
                                     ByVal namesRange     As String, _
                                     ByVal addressesRange As String, _
                                     Optional ByVal initValue As Variant = Null) As Variant
    Dim rptSheet As Worksheet
    Dim nameAreas As Range, addrAreas As Range
    Dim area As Range, cell As Range
    Dim nameList As Collection, addrList As Collection
    Dim fieldDefs() As Variant
    Dim i As Long

    ' 1. 取得工作表
    On Error Resume Next
    Set rptSheet = ThisWorkbook.Worksheets(rptSheetName)
    On Error GoTo 0
    If rptSheet Is Nothing Then
        Err.Raise vbObjectError, , "找不到工作表：" & rptSheetName
    End If

    ' 2. 取得多段範圍
    On Error Resume Next
    Set nameAreas = rptSheet.Range(namesRange)
    Set addrAreas = rptSheet.Range(addressesRange)
    On Error GoTo 0
    If nameAreas Is Nothing Then _
        Err.Raise vbObjectError, , "namesRange 無效：" & namesRange
        WriteLog "namesRange 無效：" & namesRange
    If addrAreas Is Nothing Then _
        Err.Raise vbObjectError, , "addressesRange 無效：" & addressesRange
        WriteLog "addressesRange 無效：" & addressesRange
    ' 3. 把所有 namesRange 的每一個 cell 依照遍歷順序收進 nameList
    Set nameList = New Collection
    For Each area In nameAreas.Areas
        For Each cell In area.Cells
            nameList.Add cell
        Next
    Next

    ' 4. 把所有 addressesRange 的每一個 cell.value 依照遍歷順序收進 addrList
    Set addrList = New Collection
    For Each area In addrAreas.Areas
        For Each cell In area.Cells
            addrList.Add CStr(cell.Value)
        Next
    Next

    ' 5. 確認「名稱清單」與「位址清單」筆數相同
    If nameList.Count <> addrList.Count Then
        Err.Raise vbObjectError, , _
          "名稱筆數 (" & nameList.Count & ") 與位址筆數 (" & addrList.Count & ") 不一致。"
        WriteLog "名稱筆數 (" & nameList.Count & ") 與位址筆數 (" & addrList.Count & ") 不一致。"
    End If

    ' 6. 建立回傳陣列
    ReDim fieldDefs(0 To nameList.Count - 1)
    For i = 1 To nameList.Count
        ' 先清除 Err
        Err.Clear

        On Error Resume Next

        Dim nm As String
        nm = nameList(i).Name.Name
        On Error GoTo 0
        
        ' 如果沒取到名稱，就主動拋錯並中斷
        If nm = "" Then
            Err.Raise vbObjectError + 513, "GetFieldFromXlRanges", _
                      "第 " & i & " 筆 nameList 中的儲存格沒有名稱。"
            WriteLog "第 " & i & " 筆 nameList 中的儲存格沒有名稱。"
        End If

        ' cell.Name.Name 取的是「該儲存格所屬的定義名稱」；錯誤就回空字串
        fieldDefs(i - 1) = Array( _
            nm, _
            addrList(i), _
            initValue _
        )
        Err.Clear
        On Error GoTo 0
    Next

    GetFieldFromXlRanges = fieldDefs
End Function







    ## 2. `Init` 裡呼叫並傳給 `AddWorksheetFields`

```vb
Public Sub Init(ByVal reportName As String, _
                ByVal dataMonthStringROC      As String, _
                ByVal dataMonthStringROC_NUM  As String, _
                ByVal dataMonthStringROC_F1F2 As String)

    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")

    Select Case reportName

        ' …… 你原有的 Case ……

        Case "Table50"
            Dim defs As Variant
            ' ← 取得設定表 FieldConfig!R2:R50 (命名儲存格清單)、S2:S50 (對應位址)
            defs = Me.GetFieldFromXlRanges( _
                        "FieldConfig", "R2:R50", "S2:S50", dataMonthStringROC _
                   )   ' ### Modified
            ' ← 把它傳給既有的 AddWorksheetFields，一次加入所有欄位
            AddWorksheetFields "Table50", defs    ' ### Modified

        ' …… 其他 Case ……

    End Select
End Sub
```

---

## 3. Process 階段取回欄位，逐一寫進工作表

假設你的 Process\_Table50 程序要把已取到或計算好的值，寫回到 Excel：

```vb
Public Sub Process_Table50()
    Dim rpt As clsReport
    Set rpt = gReports("Table50")

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Table50")

    ' 取得「欄位名稱 → 儲存格位址」字典
    Dim posDict As Object
    Set posDict = rpt.GetAllFieldPositions("Table50")

    ' 取得「欄位名稱 → 欄位值」字典（事先透過 SetField 填好的）
    Dim valDict As Object
    Set valDict = rpt.GetAllFieldValues("Table50")

    Dim fld As Variant, addr As String, val As Variant

    For Each fld In posDict.Keys
        addr = posDict(fld)
        If valDict.Exists(fld) Then
            val = valDict(fld)
            ws.Range(addr).Value = val
        End If
    Next fld

    ' … 其餘更新 DB、驗證流程照舊 …
End Sub
