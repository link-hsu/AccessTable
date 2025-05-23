Q:

	淨額	總額	餘額
FB1_FOA_0001_資產	FB1_FOA_0001_資產_淨額	FB1_FOA_0001_資產_總額	FB1_FOA_0001_資產_餘額
FB1_FOA_0002_負債	FB1_FOA_0001_資產_淨額	FB1_FOA_0001_資產_總額	FB1_FOA_0001_資產_餘額


請問我有辦法使用迴圈將
FB1_FOA_0001_資產
FB1_FOA_0002_負債
和
淨額
總額
餘額
組合而成的

FB1_FOA_0001_資產_淨額	FB1_FOA_0001_資產_總額	FB1_FOA_0001_資產_餘額
FB1_FOA_0001_資產_淨額	FB1_FOA_0001_資產_總額	FB1_FOA_0001_資產_餘額
這是Range("FB1_FOA_0001_資產_淨額")中的名稱，我已經將數值填入這個value上面了，
我要寫迴圈去執行

我要逐一去執行這樣的method
rpt.SetField "FOA", "FB1_FOA_0001_資產_淨額", CStr(xlsht.Range("FB1_FOA_0001_資產_淨額").Value)
其中要從FB1_FOA_0001_資產_淨額中的字串，擷取_中的第二個字串填入
請問要怎麼寫?
請給我一個function，填入必要參數就可以處理


A:

下面是一個通用的 VBA Sub，接受「報表物件」、「工作表」、主名稱陣列（BaseNames）和後綴陣列（Suffixes）四個參數，就能自動組合出所有 `BaseName_Suffix` 並呼叫 `rpt.SetField`。裡面用 `Split(..., "_")(1)` 取出第二個底線之間的字串當作第一個參數傳進去：

```vb
'''----------------------------------------------------------
' 依照 baseNames 和 suffixes 自動呼叫 rpt.SetField
'
' rpt        : 你的報表物件（必須有 SetField 方法）
' xlsht      : Excel 工作表
' baseNames  : String 陣列，元素形如 "FB1_FOA_0001_資產"
' suffixes   : String 陣列，元素形如 "淨額", "總額", "餘額"
'''----------------------------------------------------------
Public Sub PopulateRPTFields( _
    ByVal rpt As Object, _
    ByVal xlsht As Worksheet, _
    ByVal baseNames As Variant, _
    ByVal suffixes As Variant)

    Dim i As Long, j As Long
    Dim fullName As String
    Dim parts As Variant
    Dim fieldKey As String

    On Error GoTo ErrHandle

    For i = LBound(baseNames) To UBound(baseNames)
        For j = LBound(suffixes) To UBound(suffixes)
            ' 組合出完整名稱
            fullName = baseNames(i) & "_" & suffixes(j)
            
            ' 拆解底線字串，取第二段作為 rpt.SetField 的 Key
            parts = Split(fullName, "_")
            If UBound(parts) >= 1 Then
                fieldKey = parts(1)
            Else
                fieldKey = ""
            End If
            
            ' 讀取 Named Range 的值並傳給報表
            rpt.SetField fieldKey, fullName, CStr(xlsht.Range(fullName).Value)
        Next j
    Next i

    Exit Sub

ErrHandle:
    Debug.Print "PopulateRPTFields 錯誤：" & Err.Number & " - " & Err.Description
End Sub
```

### 範例呼叫方式

```vb
Sub TestPopulate()
    Dim arrBases   As Variant
    Dim arrSuffix  As Variant
    ' 主名稱清單
    arrBases  = Array("FB1_FOA_0001_資產", "FB1_FOA_0002_負債")
    ' 欄位後綴
    arrSuffix = Array("淨額", "總額", "餘額")
    
    ' 假設 rpt 是你的報表物件，Sheet1 已經做好 Named Range
    Call PopulateRPTFields( _
         rpt, _
         ThisWorkbook.Worksheets("Sheet1"), _
         arrBases, _
         arrSuffix _
    )
End Sub
```

只要把你的報表物件（具有 `SetField(Key, Name, Value)` 方法）、Excel 工作表，以及兩個字串陣列丟進去，就能自動跑完所有組合、抓值、並填入報表裡。



Q:

延續上面問題，那如果使用Sub搭配function來達成要怎麼寫


A:


這邊把「拆字串取 Key」獨立成一個 Function，然後主流程用 Sub 來呼叫它和填報表：

```vb
'''----------------------------------------------------------
' 取得 Named Range 全名裡，第二段底線之間的字串
'
' fullName : 形如 "FB1_FOA_0001_資產_淨額"
' 回傳值  : 第二段字串，例如 "FOA"
'''----------------------------------------------------------
Public Function ExtractFieldKey(ByVal fullName As String) As String
    Dim parts As Variant
    parts = Split(fullName, "_")
    If UBound(parts) >= 1 Then
        ExtractFieldKey = parts(1)
    Else
        ExtractFieldKey = ""
    End If
End Function


'''----------------------------------------------------------
' 主流程：自動組合 baseNames & suffixes，並把值傳給 rpt.SetField
'
' rpt        : 你的報表物件（必須有 SetField 方法）
' xlsht      : Excel 工作表
' baseNames  : String 陣列，元素形如 "FB1_FOA_0001_資產"
' suffixes   : String 陣列，元素形如 "淨額", "總額", "餘額"
'''----------------------------------------------------------
Public Sub PopulateRPTFields_WithFunc( _
    ByVal rpt As Object, _
    ByVal xlsht As Worksheet, _
    ByVal baseNames As Variant, _
    ByVal suffixes As Variant)

    Dim i As Long, j As Long
    Dim fullName As String
    Dim fieldKey As String
    
    On Error GoTo ErrHandle

    For i = LBound(baseNames) To UBound(baseNames)
        For j = LBound(suffixes) To UBound(suffixes)
            ' 組合出完整 Named Range 名稱
            fullName = baseNames(i) & "_" & suffixes(j)
            ' 用 Function 拆字串取 Key
            fieldKey = ExtractFieldKey(fullName)
            ' 讀取 Named Range 的值並傳給報表
            rpt.SetField fieldKey, fullName, CStr(xlsht.Range(fullName).Value)
        Next j
    Next i

    Exit Sub

ErrHandle:
    Debug.Print "PopulateRPTFields_WithFunc 錯誤：" & Err.Number & " - " & Err.Description
End Sub
```

### 範例呼叫

```vb
Sub TestPopulateWithFunc()
    Dim arrBases   As Variant
    Dim arrSuffix  As Variant
    
    ' 主名稱清單
    arrBases  = Array("FB1_FOA_0001_資產", "FB1_FOA_0002_負債")
    ' 欄位後綴
    arrSuffix = Array("淨額", "總額", "餘額")
    
    ' 呼叫帶 Function 的 Sub
    Call PopulateRPTFields_WithFunc( _
         rpt, _
         ThisWorkbook.Worksheets("Sheet1"), _
         arrBases, _
         arrSuffix _
    )
End Sub
```

這樣你就把拆 Key 的邏輯抽到 `ExtractFieldKey`，主流程 `PopulateRPTFields_WithFunc` 乾淨易讀，也方便重複使用或單獨測試。
