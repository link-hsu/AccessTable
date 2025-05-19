票券交易明細表

## 💵 票券交易明細表 BillTransactionDetails
| 中文欄位 | 英文命名 |
|----------|-----------|
交易日期 | `TradeDate`  
交割日期 | `SettlementDate`  
成交單編號 | `TradeID`  
帳務目的 | `AccountingPurpose`  
交易對手 | `Counterparty`  
統一編號 | `TaxID`  
交易員 | `Trader`  
覆核層級 | `ReviewLevel`  
交易別 | `TransactionType`  
訊息類別 | `MessageType`  
票券批號-子號 | `BillBatchSubNumber`  
票類 | `BillType`  
發票人 | `Issuer`  
保證人 | `Guarantor`  
統一編號 | `IssuerTaxID`, `GuarantorTaxID`, `CounterpartyTaxID`  
發票日 | `IssueDate`  
到期日 | `MaturityDate`  
面額 | `FaceValue`  
萬元單價/票載利率 | `PricePerTenThousandOrCoupon`  
成交利率 | `TradeYield`  
天數 | `Days`  
成交金額 | `TradeAmount`  
帳上成本 | `BookCost`  
利息累計稅款 | `AccruedTax`  
利息金額 | `InterestAmount`  
約定到期金額 | `AgreedMaturityAmount`  
約定到期日 | `AgreedMaturityDate`  
損益 | `ProfitLoss`  
帳上稅款 | `BookTax` 

- 筆數
    - 新臺幣商業本票
        - 新臺幣融資性本票
=COUNTIF('票券交易明細表選首購、承銷發行)'!L:L,"CP2")

'票券交易明細表選首購、承銷發行)'

首購承銷發行，問題，有辦法從資料判別嗎?還是一定需要拉不同的表


L欄位 票類 
CP2
    - 央行可轉讓定期存單
=COUNTIF('票券交易明細表選首購、承銷發行)'!L:L,"央行NCD")

L欄位 票類 
央行NCD

- 金額
    - 新臺幣商業本票
        - 新臺幣融資性本票
=SUMIF('票券交易明細表選首購、承銷發行)'!L:L,"CP2",'票券交易明細表選首購、承銷發行)'!S:S)

S欄位 面額
    - 央行可轉讓定期存單
=SUMIF('票券交易明細表選首購、承銷發行)'!L:L,"央行NCD",'票券交易明細表選首購、承銷發行)'!S:S)


我會從上面的 票券交易明細表 資料表中抓資料到Excel中，
然後我要寫迴圈從該資料表中抓資料，需要抓取的資料如上面的Excel函數內容，
請幫我寫vba



下面這段 VBA 巨集示範如何在同一張「票券交易明細表」中，直接用迴圈計算各票類（例如 `"CP2"`、`"央行NCD"`）的筆數與面額總和，並把結果寫到你指定的儲存格上。你可以按需要改 `billTypes` 與 `outputRange` 來擴充或調整輸出位置。

```vb
Sub SummarizeByBillType()
    Dim ws         As Worksheet
    Dim lastRow    As Long
    Dim i As Long, j As Long
    
    ' 要統計的票類清單
    Dim billTypes  As Variant
    billTypes = Array("CP2", "央行NCD")
    
    ' 對應的輸出起始儲存格 (count, sum)：  
    '  這裡示範 CP2 的「筆數」放在 E2、「金額」放在 F2；  
    '           NCD 的筆數放 E3，金額放 F3
    Dim outputRange As Variant
    outputRange = Array(Array("E2", "F2"), Array("E3", "F3"))
    
    ' 暫存計算結果
    Dim cnt() As Long
    Dim tot() As Double
    ReDim cnt(LBound(billTypes) To UBound(billTypes))
    ReDim tot(LBound(billTypes) To UBound(billTypes))
    
    ' 指定工作表
    Set ws = ThisWorkbook.Worksheets("票券交易明細表")
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row  ' L 欄為票類
    
    ' 掃描所有資料列
    For i = 2 To lastRow
        Dim bt As String
        bt = ws.Cells(i, "L").Value
        Dim fv As Double
        fv = ws.Cells(i, "S").Value    ' S 欄為面額
        
        ' 比對清單，若相符就累加
        For j = LBound(billTypes) To UBound(billTypes)
            If bt = billTypes(j) Then
                cnt(j) = cnt(j) + 1
                tot(j) = tot(j) + fv
            End If
        Next j
    Next i
    
    ' 輸出到指定儲存格
    For j = LBound(billTypes) To UBound(billTypes)
        ws.Range(outputRange(j)(0)).Value = cnt(j)
        ws.Range(outputRange(j)(1)).Value = tot(j)
    Next j
    
    MsgBox "統計完成！", vbInformation
End Sub
```

---

### 為什麼這樣寫？

1. **二維陣列管理輸出位置**
   `outputRange = Array(Array("E2","F2"), Array("E3","F3"))`
   讓你只要在同一個結構內一併設定「筆數」與「金額」要放到哪裡，比起重複寫 `Range("E2")…Range("F2")` 更集中、好維護。

2. **一層迴圈＋比對清單**

   * 外圈掃所有列：`For i = 2 To lastRow`
   * 內圈對已知的票類陣列跑一次：`For j = …`
   * 一筆資料同時能命中多種票類也沒問題（但通常一列只有一種票類）。

3. **ReDim 動態處理統計陣列**

   * `cnt(j)` 存筆數、`tot(j)` 存面額加總，維度自動跟 `billTypes` 數量一致。

4. **易於擴充**
   未來要多加票類，只要在 `billTypes` 裡新增字串，再在 `outputRange` 加入對應位置即可，程式本身不需動。

---

### 關於「首購承銷發行」要不要拉不同表？

* 若你的資料表裡 **沒有** 類似 `TransactionType`、`AccountingPurpose` 或其他欄位明確標示「首購／承銷／發行」，就無從以程式自動判別，只能額外在原始資料匯出前於 SQL 層面做標記或分表。
* 如果你的 `票券交易明細表` 裡已有一個欄位（比如 `TransactionType`）能分辨「首購」或「承銷發行」，則在內圈再多一個 `If ws.Cells(i, "TransactionType") = "首購" Then…` 即可，毋須拉新表。

希望這支巨集正好符合你的需求！如果有其他票類或篩選條件，隨時再調整對應陣列即可。
