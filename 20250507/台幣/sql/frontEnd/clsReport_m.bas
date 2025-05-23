Public Sub Init(ByVal reportName As String, _
                ByVal dataMonthStringROC As String, _
                ByVal dataMonthStringROC_NUM As String, _
                ByVal dataMonthStringROC_F1F2 As String)
    Dim rptFieldInfo As Object
    Dim rptToInit As Variant
    Dim rptArray As Variant
    Set rptFieldInfo = CreateObject("Scripting.Dictionary")

    With rptFieldInfo
        .Add "TABLE10", _
        Array(
            Array("FOA", "S2:S21,W2:W21,AA2:AA21,AE2:AE21,AI2:AI21,AM2:AM21,AQ2:AQ21", "U2:U21,Y2:Y21,AC2:AC21,AG2:AG21,AK2:AK21,AO2:AO21,AS2:AS21")            
        )

        .Add "TABLE15A", _
        Array(
            Array("FOA", "S2:S10,W2:W10,AA2:AA10,AE2:AE10,AI2:AI10", "U2:U10,Y2:Y10,AC2:AC10,AG2:AG10,AK2:AK10")
        )

        .Add "TABLE15B", _
        Array()

        .Add "TABLE16", _
        Array(
            Array("FOA", "", "B2")
        )

        .Add "TABLE20", _
        Array(
            Array("FOA", "S2:S5,W2:W5,AA2:AA5", "U2:U5,Y2:Y5,AC2:AC5")
        )

        .Add "TABLE22", _
        Array(
            Array("FOA", "S2:S5,W2:W5", "U2:U5,Y2:Y5")
        )

        .Add "TABLE23", _
        Array(
            Array("FOA", "S2:S6", "U2:U6")
        )

        .Add "TABLE24", _
        Array(
            Array("FOA", "S2:S14,W2:W14,AA2:AA14,AE2:AE14,AI2:AI14,AM2:AM14,AQ2:AQ14,AU2:AU14", "U2:U14,Y2:Y14,AC2:AC14,AG2:AG14,AK2:AK14,AO2:AO14,AS2:AS14,AW2:AW14")
        )

        .Add "TABLE27", _
        Array(
            Array("FOA", "S2:S7,W2:W7,AA2:AA7,AE2:AE7,AI2:AI7", "U2:U7,Y2:Y7,AC2:AC7,AG2:AG7,AK2:AK7")
        )

        .Add "TABLE36", _
        Array(
            Array("FOA", "S2:S4,W2:W4,AA2:AA4", "U2:U4,Y2:Y4,AC2:AC4")
        )

        .Add "AI233", _
        Array(
            Array("Table1", "S2:S5,W2:W5,AA2:AA5,AE2:AE9,AI2:AI5,AM2:AM5,AQ2:AQ5,AU2:AU9", "U2:U5,Y2:Y5,AC2:AC5,AG2:AG9,AK2:AK5,AO2:AO5,AS2:AS5,AW2:AW9"),
            Array("Table2", "S10:S11,W10:W11,AA10:AA11", "U10:U11,Y10:Y11,AC10:AC11"),
            Array("Table4", "S12:S15,W12:W15,AA12:AA15,AE12:AE17", "U12:U15,Y12:Y15,AC12:AC15,AG12:AG17")
        )

        .Add "AI345", _
        Array(
            Array("", "", "")
        )

        .Add "AI405", _
        Array(
            Array("Table1", "S2:S5,W2:W5", "U2:U5,Y2:Y5")
        )

        .Add "AI410", _
        Array(
            Array("Table1", "S2:S8,W2:W8", "U2:U8,Y2:Y8")
        )

        .Add "AI430", _
        Array(
            Array("Table1", "S2:S8", "U2:U8")
        )

        .Add "AI601", _
        Array(
            Array("Table1", "S2:S48,W2:W48,AA2:AA48,AE2:AE48,AI2:AI48", "U2:U48,Y2:Y48,AC2:AC48,AG2:AG48,AK2:AK48"),
            Array("Table2", "AM2:AM48,AQ2:AQ48,AU2:AU48,AY2:AY48,BC2:BC48,BG2:BG48,BK2:BK48", "AO2:AO48,AS2:AS48,AW2:AW48,BA2:BA48,BE2:BE48,BI2:BI48,BM2:BM48"),
            Array("Table3", "S49:S65", "U49:U65")
        )

        .Add "AI605", _
        Array(
            Array("Table1", "S2:S3,W2:W3,AA2:AA3,AE2:AE3,AI2:AI3,AM2:AM3,AQ2:AQ3,AU2:AU3", "U2:U3,Y2:Y3,AC2:AC3,AG2:AG3,AK2:AK3,AO2:AO3,AS2:AS3,AW2:AW3"),
            Array("Table3", "S5:S6,W5:W6,", "U5:U6,Y5:Y6")
        )
    End With

    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")

    If rptFieldInfo.Exists(reportName) Then
        rptToInit = rptFieldInfo(reportName)

        If (Not IsArray(rptToInit)) Or UBound(rptToInit) < LBound(rptToInit) Then
            WriteLog "Init 跳過 [" & reportName & "]：fieldDefs 陣列為空"
        Else
            For i = LBound(rptToInit) To UBound(rptToInit)
                Dim rptSheet As Variant
                rptSheet = rptToInit(i)
                
                Dim initSheetName As String
                Dim nameTagRng As String
                Dim addrRng As String
                Dim initValue As Variant
                
                initSheetName = rptSheet(0)
                nameTagRng = rptSheet(1)
                addrRng = rptSheet(2)
                initValue = Null
                
                ' 跳過空的 range 定義
                If Trim(nameTagRng) = "" Or Trim(addrRng) = "" Then
                    WriteLog "Init 跳過 [" & reportName & "] 的 [" & initSheetName & "]：range 定義為空"
                Else
                    rptArray = Me.GetFieldFromXlRanges(reportName, nameTagRng, addrRng, initValue)
                    ' 呼叫 AddWorksheetFields，第一參數用 initSheetName
                    AddWorksheetFields initSheetName, rptArray
                End If
            Next i
        End If
    Else
        WriteLog "Init未定義報表：" & reportName
    End If
    
    Select Case reportName
        Case "TABLE10"
            AddDynamicField reportName, "TABLE10_申報時間", "D2", dataMonthStringROC
        Case "TABLE15A"
            AddDynamicField reportName, "TABLE15A_申報時間", "D2", dataMonthStringROC
        Case "TABLE15B"
            AddDynamicField reportName, "TABLE15B_申報時間", "D2", dataMonthStringROC
        Case "TABLE16"
            AddDynamicField reportName, "TABLE16_申報時間", "B2", dataMonthStringROC
        Case "TABLE20"
            AddDynamicField reportName, "TABLE20_申報時間", "I3", dataMonthStringROC
        Case "TABLE22"
            AddDynamicField reportName, "TABLE22_申報時間", "E2", dataMonthStringROC
        Case "TABLE23"
            AddDynamicField reportName, "TABLE23_申報時間", "E2", dataMonthStringROC
        Case "TABLE24"
            AddDynamicField reportName, "TABLE24_申報時間", "G2", dataMonthStringROC
        Case "TABLE27"
            AddDynamicField reportName, "TABLE27_申報時間", "E3", dataMonthStringROC
        Case "TABLE36"
            AddDynamicField reportName, "TABLE36_申報時間", "E2", dataMonthStringROC
        Case "AI233"
            AddDynamicField reportName, "AI233_申報時間", "B3", dataMonthStringROC_NUM
        Case "AI345"
            AddDynamicField reportName, "AI345_申報時間", "A2", dataMonthStringROC_NUM
        Case "AI405"
            AddDynamicField reportName, "AI405_申報時間", "B3", dataMonthStringROC_NUM
        Case "AI410"
            AddDynamicField reportName, "AI410_申報時間", "B3", dataMonthStringROC_NUM
        Case "AI430"
            AddDynamicField reportName, "AI430_申報時間", "B3", dataMonthStringROC_NUM
        Case "AI601"
            AddDynamicField reportName, "AI601_申報時間", "B3", dataMonthStringROC_NUM
        Case "AI605"
            AddDynamicField reportName, "AI605_申報時間", "B3", dataMonthStringROC_NUM
        ' 如有其他報表，依需求加入不同工作表及欄位定義
    End Select
End Sub







Q:
請問class clsReport中的init method裡面的rptFieldInfo，
要怎麼共用到主執行程序中的Process也可以使用?

A:
要讓 `rptFieldInfo` 這份「報表→工作表範圍設定」的對應表，不只在 `clsReport.Init` 裡面用得到，主執行程序的各個 `Process_XXX` 也能共用，最乾淨的做法是把它從「類別私有」提出，放到一個共用的標準 Module 裡，讓雙方都可以呼叫。

---

### 步驟 1：建立一個新的標準 Module（例如取名 `modRptConfig`）

```vb
' Module: modRptConfig
Option Explicit

' 回傳整份 rptFieldInfo Dictionary
Public Function GetRptFieldInfo() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    With dict
        .Add "TABLE10", Array( _
            Array("FOA", "S2:S21,W2:W21,AA2:AA21,AE2:AE21,AI2:AI21,AM2:AM21,AQ2:AQ21", _
                        "U2:U21,Y2:Y21,AC2:AC21,AG2:AG21,AK2:AK21,AO2:AO21,AS2:AS21") )
        .Add "TABLE15A", Array( _
            Array("FOA", "S2:S10,W2:W10,AA2:AA10,AE2:AE10,AI2:AI10", _
                        "U2:U10,Y2:Y10,AC2:AC10,AG2:AG10,AK2:AK10") )
        .Add "TABLE15B", Array()  ' 空陣列
        ' …（把原本 Init 內 With rptFieldInfo 的所有 .Add 都搬過來）…
        .Add "AI605", Array( _
            Array("Table1", "S2:S3,W2:W3,AA2:AA3,AE2:AE3,AI2:AI3,AM2:AM3,AQ2:AQ3,AU2:AU3", _
                        "U2:U3,Y2:Y3,AC2:AC3,AG2:AG3,AK2:AK3,AO2:AO3,AS2:AS3,AW2:AW3"), _
            Array("Table3", "S5:S6,W5:W6", "U5:U6,Y5:Y6") )
    End With
    
    Set GetRptFieldInfo = dict
End Function
```

---

### 步驟 2：在 `clsReport.Init` 裡呼叫這個 Function

把原本自己 `CreateObject("Scripting.Dictionary")` 然後 `.Add …` 的區塊刪掉，改成：

```vb
Public Sub Init( _ 
    ByVal reportName As String, _
    ByVal dataMonthStringROC As String, _
    ByVal dataMonthStringROC_NUM As String, _
    ByVal dataMonthStringROC_F1F2 As String)

    Dim rptFieldInfo As Object
    Dim rptToInit As Variant
    ' ← 由共用 Module 拿設定
    Set rptFieldInfo = GetRptFieldInfo()

    clsReportName = reportName
    Set clsWorksheets = CreateObject("Scripting.Dictionary")

    If rptFieldInfo.Exists(reportName) Then
        rptToInit = rptFieldInfo(reportName)
        ' …（後面不用動）…
    End If
    ' …Select Case etc.…
End Sub
```

---

### 步驟 3：在主程式或 `Process_XXX` 也同樣呼叫

假設你在某個 `Process_TABLE10` 想看它的「範圍定義」：

```vb
Public Sub Process_TABLE10()
    Dim rptFieldInfo As Object
    Set rptFieldInfo = GetRptFieldInfo()

    Dim defs As Variant
    defs = rptFieldInfo("TABLE10")
    ' defs 就是一個二維陣列，你可以拿來做進一步驗證或其他……
    
    ' 例如印出第一組定義的第一欄名稱：
    Debug.Print defs(0)(0)   ' FOA
End Sub
```

---

#### 為什麼這樣做？

1. **單一來源**：所有報表的範圍定義只寫在一份 Module 裡，維護一次，各處都生效。
2. **類別專注職責**：`clsReport` 只負責「拿到定義、初始化自己」，不必同時管理這一大段靜態設定。
3. **Process 也能拿**：主程式的任何 Sub/Function，只要 `Set dict = GetRptFieldInfo()`，就能取到一模一樣的設定，避免重複寫、不同步的風險。

照著上面三個步驟，把原本散落在 `clsReport.Init` 裡的那整段 `.Add`，移到新的 `GetRptFieldInfo` Function，就能達到「共用」的效果。
