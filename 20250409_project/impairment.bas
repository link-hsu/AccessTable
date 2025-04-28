你可以在原來的程式裡面，加一個「從 valueType 到對應 code」的 Dictionary（或用平行陣列也可以），然後在貼 L 欄之後，再一行把對應的 code 整列貼到 M 欄。以下示範用 Dictionary 的作法——把這段插到你宣告完 `valueType` 之後，並在貼完 L 欄後，貼一整欄 M。

```vb
    ' 在這裡定義你的 valueType（你原本就有）
    valueType = Array( _
       "強制FVPL金融資產-公債-中央政府", "強制FVPL金融資產-公債-地方政府(我國)", _
       … _
       "AC債務工具投資-金融債券-海外" _
    )

    ' 新增：定義平行的 codeArray
    Dim codeArray As Variant
    codeArray = Array( _
       "FVPL_GovBond_Domestic", "FVPL_GovBond_Domestic", _
       "FVPL_CompanyBond_Domestic", "FVPL_CompanyBond_Domestic", _
       "FVPL_CP_Domestic", "FVOCI_NCD_Domestic", _
       "FVOCI_GovBond_Domestic", "FVOCI_GovBond_Domestic", _
       "FVOCI_CompanyBond_Domestic", "FVOCI_CompanyBond_Domestic", _
       "AC_NCD_Domestic", "AC_GovBond_Domestic", _
       "AC_GovBond_Domestic", "AC_CompanyBond_Domestic", _
       "AC_CompanyBond_Domestic", _
       "FVPL_GovBond_Foreign", "FVPL_CompanyBond_Foreign", _
       "FVPL_CompanyBond_Foreign", "FVPL_FinancialBond_Foreign", _
       "FVOCI_GovBond_Foreign", "FVOCI_CompanyBond_Foreign", _
       "FVOCI_CompanyBond_Foreign", "FVOCI_FinancialBond_Foreign", _
       "AC_GovBond_Foreign", "AC_CompanyBond_Foreign", _
       "AC_CompanyBond_Foreign", "AC_FinancialBond_Foreign" _
    )

    ' 建立 Dictionary 方便 lookup
    Dim mapCode As Object
    Set mapCode = CreateObject("Scripting.Dictionary")
    Dim j As Long
    For j = LBound(valueType) To UBound(valueType)
        mapCode.Add valueType(j), codeArray(j)
    Next j
```

然後在你的分頁建立迴圈中，貼完 L 欄之後（就是這一行之後）：

```vb
            .Range("L2:L" & innerLastRow).Value = securityName(i)
            ' 新增：將對應的 code 貼到 M 欄
            .Range("M2:M" & innerLastRow).Value = mapCode(securityName(i))
```

完整範例（只秀重點區段）：

```vb
    ' …… 前面宣告與開檔省略 ……

    valueType = Array( … )    ' 你的 valueType
    codeArray  = Array( … )    ' 如上所示的 codeList

    ' 建 Dictionary
    Set mapCode = CreateObject("Scripting.Dictionary")
    For j = LBound(valueType) To UBound(valueType)
        mapCode.Add valueType(j), codeArray(j)
    Next j

    ' …… 找出 securityIndex, securityName 省略 ……

    For i = 1 To securityIndex.Count
        ' …… 計算 startRow, endRow, sheetName 省略 ……

        With ActiveSheet
            .Name = sheetName
            xlsht.Range(xlsht.Cells(startRow, "C"), xlsht.Cells(endRow, "M")).Copy
            .Range("A2").PasteSpecial Paste:=xlPasteValues

            innerLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row

            ' 把分頁名稱貼到 L
            .Range("L2:L" & innerLastRow).Value = securityName(i)
            ' 把對應 code 貼到 M
            .Range("M2:M" & innerLastRow).Value = mapCode(securityName(i))
        End With

    ContinueLoop:
    Next i
```

這樣每個新分頁裡，L 欄會是你的原始分類名稱，M 欄就會是一對一對應好的英文 code。





-------------------------------------------------------------------------------





減損表處理








-------------------------------------------------------------------------------------------------------------
```vb
Public Sub ICleaner_CleanReport(ByVal fullFilePath As String, _
                                ByVal cleaningType As String)
    Dim xlApp As Excel.Application
    Dim xlbk As Excel.Workbook
    Dim xlsht As Excel.Worksheet

    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim colArray As Variant
    Dim valueType As Variant
    Dim codeArray As Variant
    Dim eachType As Variant

    Dim securityIndex As Collection
    Dim securityName As Collection

    Dim startRow As Long
    Dim endRow As Long
    Dim innerLastRow As Long
    Dim sheetName As String

    Dim toDelete As Boolean

    Dim mapCode As Object

    If Dir(fullFilePath) = "" Then
        MsgBox "File does not exist in path: " & fullFilePath
        Exit Sub
    End If

    Set xlApp = Excel.Application
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False
    xlApp.AskToUpdateLinks = False

    Set xlbk = xlApp.Workbooks.Open(fullFilePath)
    Set xlsht = xlbk.Sheets("減損")

    colArray = Array("Security_id", _
                     "issuer", _
                     "成本", _
                     "應收利息", _
                     "信評", _
                     "PD", _
                     "LGD", _
                     "上期減損數(成本)", _
                     "本期減損數(成本)", _
                     "上期減損數(利息)", _
                     "本期減損數(利息)")

    valueType = Array( _
        "強制FVPL金融資產-公債-中央政府", "強制FVPL金融資產-公債-地方政府(我國)", _
        "強制FVPL金融資產-普通公司債(公", "強制FVPL金融資產-普通公司債(民", _
        "強制FVPL金融資產-商業本票", "FVOCI債務工具-央行NCD", _
        "FVOCI債務工具-公債-中央政府(我", "FVOCI債務工具-公債-地方政府(我國)", _
        "FVOCI債務工具-普通公司債（公營", "FVOCI債務工具-普通公司債（民營", _
        "AC債務工具-央行NCD", "AC債務工具投資-公債-中央政府(?", _
        "AC債務工具投資-公債-地方政府(?", "AC債務工具投資-普通公司債(公營", _
        "AC債務工具投資-普通公司債(民營", _
        "強制FVPL金融資產-公債-中央政府(外國)", "強制FVPL金融資產-普通公司債(公營)-海外", _
        "強制FVPL金融資產-普通公司債(民營)-海外", "FVOCI債務工具-公債-中央政府(外國)", _
        "FVOCI債務工具-普通公司債(公營)-海外", "FVOCI債務工具-普通公司債(民營)-海外", _
        "FVOCI債務工具-金融債券-海外", _
        "AC債務工具投資-公債-中央政府(外國)", "AC債務工具投資-普通公司債(公營)-海外", _
        "AC債務工具投資-普通公司債(民營)-海外", "AC債務工具投資-金融債券-海外" _
    )

    codeArray = Array( _
        "FVPL_GovBond_Domestic", "FVPL_GovBond_Domestic", _
        "FVPL_CompanyBond_Domestic", "FVPL_CompanyBond_Domestic", _
        "FVPL_CP_Domestic", "FVOCI_NCD_Domestic", _
        "FVOCI_GovBond_Domestic", "FVOCI_GovBond_Domestic", _
        "FVOCI_CompanyBond_Domestic", "FVOCI_CompanyBond_Domestic", _
        "AC_NCD_Domestic", "AC_GovBond_Domestic", _
        "AC_GovBond_Domestic", "AC_CompanyBond_Domestic", _
        "AC_CompanyBond_Domestic", _
        "FVPL_GovBond_Foreign", "FVPL_CompanyBond_Foreign", _
        "FVPL_CompanyBond_Foreign", "FVPL_FinancialBond_Foreign", _
        "FVOCI_GovBond_Foreign", "FVOCI_CompanyBond_Foreign", _
        "FVOCI_CompanyBond_Foreign", "FVOCI_FinancialBond_Foreign", _
        "AC_GovBond_Foreign", "AC_CompanyBond_Foreign", _
        "AC_CompanyBond_Foreign", "AC_FinancialBond_Foreign" _
    )

    ' 建立 Dictionary 方便對照
    Set mapCode = CreateObject("Scripting.Dictionary")
    For j = LBound(valueType) To UBound(valueType)
        mapCode.Add valueType(j), codeArray(j)
    Next j

    lastRow = xlsht.Cells(xlsht.Rows.Count, 3).End(xlUp).Row

    ' 刪空白與利息備抵數之後的列
    For i = lastRow To 1 Step -1
        If IsEmpty(xlsht.Cells(i, 3).Value) Then
            xlsht.Rows(i).Delete
        ElseIf Left(xlsht.Cells(i, "I").Value, 5) = "利息備抵數" Then
            xlsht.Rows(i & ":" & lastRow).Delete
            Exit For
        End If
    Next i

    lastRow = xlsht.Cells(xlsht.Rows.Count, 3).End(xlUp).Row

    Set securityName = New Collection
    Set securityIndex = New Collection

    ' 找出每個分類在原表中的起始列
    For i = 1 To lastRow
        For Each eachType In valueType
            If Trim(xlsht.Cells(i, 3).Value) = eachType Then
                securityIndex.Add i
                securityName.Add eachType
            End If
        Next eachType
    Next i

    ' 依序建立新分頁並貼資料
    For i = 1 To securityIndex.Count
        If i < securityIndex.Count Then
            startRow = securityIndex(i) + 1
            endRow = securityIndex(i + 1) - 1
        Else
            startRow = securityIndex(i) + 1
            endRow = lastRow
        End If

        sheetName = securityName(i)
        If InStr(sheetName, "?") > 0 Then
            sheetName = Replace(sheetName, "?", "")
        End If

        xlbk.Sheets.Add After:=xlbk.Sheets(xlbk.Sheets.Count)
        With ActiveSheet
            .Name = sheetName
            xlsht.Range(xlsht.Cells(startRow, "C"), xlsht.Cells(endRow, "M")).Copy
            .Range("A2").PasteSpecial Paste:=xlPasteValues

            innerLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row

            ' L 欄：原分類名稱
            .Range("L2:L" & innerLastRow).Value = securityName(i)
            ' M 欄：對應的英文 code
            .Range("M2:M" & innerLastRow).Value = mapCode(securityName(i))
        End With
    Next i

    ' 刪除非 valueType 的工作表
    For i = xlbk.Sheets.Count To 1 Step -1
        toDelete = True
        For Each eachType In valueType
            If xlbk.Sheets(i).Name = eachType Then
                toDelete = False
                Exit For
            End If
        Next eachType
        If toDelete Then xlbk.Sheets(i).Delete
    Next i

    xlApp.CutCopyMode = False
    xlApp.ScreenUpdating = True
    xlApp.DisplayAlerts = True
    xlApp.AskToUpdateLinks = True

    xlbk.Save
    xlbk.Close False
    xlApp.Quit

    Set mapCode = Nothing
    Set securityIndex = Nothing
    Set securityName = Nothing
    Set xlsht = Nothing
    Set xlbk = Nothing
    Set xlApp = Nothing

    MsgBox "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
    Debug.Print "完成清理 " & cleaningType & " ，路徑為: " & fullFilePath
End Sub
```
