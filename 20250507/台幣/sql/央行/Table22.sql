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
    IIf([BillTransactionDetails.Counterparty] LIKE "*銀行", "銀行",
        IIf([BillTransactionDetails.Counterparty] LIKE "*票券", "票券", "其他")) AS 交易對手類別,
    SUM([BillTransactionDetails.FaceValue]) AS 總面額
FROM 
    BillTransactionDetails
WHERE
    BillTransactionDetails.DataMonthString = [DataMonthParam]
GROUP BY 
    IIf([BillTransactionDetails.Counterparty] LIKE "*銀行", "銀行",
        IIf([BillTransactionDetails.Counterparty] LIKE "*票券", "票券", "其他"));
