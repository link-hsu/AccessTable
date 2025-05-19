


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




        - 銀行
=ROUND(SUMIF(票券交易明細表!E:E,"*銀行",票券交易明細表!S:S)/1000000,0)
E欄位 交易對手
S欄位 面  額
        - 票券金融公司
=ROUND(SUMIF(票券交易明細表!E:E,"*票券",票券交易明細表!S:S)/1000000,0)
        - 民營企業
=ROUND(SUM(票券交易明細表!S:S)/1000000,0)-E6-E10



我會從上面的 票券交易明細表 資料表中抓資料到Excel中，
然後我要寫迴圈從該資料表中抓資料，需要抓取的資料如上面的Excel函數內容，
請幫我寫vba


Answer:


下面這支 VBA 巨集會自動掃描工作表「票券交易明細表」，針對「交易對手」欄位含有「銀行」及「票券」關鍵字的筆數，各別累加「面額」欄位，並計算「民營企業」的餘額，最後除以 1,000,000 四捨五入後，放到你指定的儲存格（範例放在 E6、E10、E14）。

```vb
Sub CalculateBillSummaries()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim faceVal As Double
    Dim sumBank As Double, sumTicketCo As Double, sumAll As Double, sumPrivate As Double
    
    ' 指定工作表
    Set ws = ThisWorkbook.Worksheets("票券交易明細表")
    
    ' 找到最後一列
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    
    ' 初始化
    sumBank = 0: sumTicketCo = 0: sumAll = 0
    
    ' 從第 2 列開始迴圈（假設第 1 列是表頭）
    For i = 2 To lastRow
        faceVal = ws.Cells(i, "S").Value        ' S 欄是「面額」
        sumAll = sumAll + faceVal               ' 全部累計
        
        Select Case True
            Case InStr(ws.Cells(i, "E").Value, "銀行") > 0
                sumBank = sumBank + faceVal     ' 交易對手含「銀行」
            Case InStr(ws.Cells(i, "E").Value, "票券") > 0
                sumTicketCo = sumTicketCo + faceVal  ' 交易對手含「票券」
        End Select
    Next i
    
    ' 計算民營企業（總和扣除銀行與票券金額）
    sumPrivate = sumAll - sumBank - sumTicketCo
    
    ' 除以 1,000,000 並四捨五入到整數，寫回工作表
    ws.Range("E6").Value = Application.Round(sumBank / 1000000, 0)
    ws.Range("E10").Value = Application.Round(sumTicketCo / 1000000, 0)
    ws.Range("E14").Value = Application.Round(sumPrivate / 1000000, 0)
    
    MsgBox "計算完成！", vbInformation
End Sub
```

**說明：**

1. **工作表**：請確定名稱「票券交易明細表」與實際工作表一致。
2. **欄位對應**：

   * `E` 欄為「交易對手」
   * `S` 欄為「面額」
3. **結果輸出**：

   * `E6`：銀行累計（單位：百萬元）
   * `E10`：票券金融公司累計（單位：百萬元）
   * `E14`：民營企業累計（單位：百萬元）
     如需放到其他儲存格，請自行調整 `ws.Range(...)` 位置。

將上述程式碼貼到 VBA 編輯器的模組中，執行即可自動計算。
