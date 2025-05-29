- 融資性商業本票
    - 買進
        - 銀行
=ROUND(SUMIF(票券交易明細表!E:E,"*銀行",票券交易明細表!S:S)/1000000,0)
E欄位 交易對手
S欄位 面  額
        - 票券金融公司
=ROUND(SUMIF(票券交易明細表!E:E,"*票券",票券交易明細表!S:S)/1000000,0)
        - 民營企業
=ROUND(SUM(票券交易明細表!S:S)/1000000,0)-E6-E10






PARAMETERS DataMonthParam TEXT;
SELECT 
    IIf([BillTransactionByTradeDate.Counterparty] LIKE "*銀行", "銀行",
        IIf([BillTransactionByTradeDate.Counterparty] LIKE "*票券", "票券", "其他")) AS 交易對手類別,
    SUM([BillTransactionByTradeDate.FaceValue]) AS 總面額
FROM 
    BillTransactionByTradeDate
WHERE
    BillTransactionByTradeDate.BillType NOT IN ('央行NCD', '一年以上央行NCD')
    AND BillTransactionByTradeDate.TransactionType NOT IN ('兌償/到期還本', '攤提', '附買回履約', '附買回解約', '附賣回履約', '附賣回解約')
    AND BillTransactionByTradeDate.DataMonthString = [DataMonthParam]
GROUP BY 
    IIf([BillTransactionByTradeDate.Counterparty] LIKE "*銀行", "銀行",
        IIf([BillTransactionByTradeDate.Counterparty] LIKE "*票券", "票券", "其他"));
