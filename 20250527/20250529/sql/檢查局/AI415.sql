票券交易明細表
- 筆數
    - 新臺幣商業本票
        - 新臺幣融資性本票
=COUNTIF('Sheet1(不含履約.兌償到期)'!L:L,"CP2")

'Sheet1(不含履約.兌償到期)'
L欄位 票類 
CP2

不含履約、兌償到期

    - 央行可轉讓定期存單
=COUNTIF('Sheet1(不含履約.兌償到期)'!L:L,"央行NCD")

L欄位 票類 
央行NCD

- 金額
    - 新臺幣商業本票
        - 新臺幣融資性本票
=SUMIF('Sheet1(不含履約.兌償到期)'!L:L,"CP2",'Sheet1(不含履約.兌償到期)'!S:S)

S欄位 面額
    - 央行可轉讓定期存單
=SUMIF('Sheet1(不含履約.兌償到期)'!L:L,"央行NCD",'Sheet1(不含履約.兌償到期)'!S:S)



BillTransactionDetails

BillType
FaceValue



AI410_BillTradeOutstanding


-- 要確認是用交易日票券明細表還是交割日

PARAMETERS DataMonthParam TEXT;
SELECT 
    BillTransactionBySettlementDate.BillType AS 類型, 
    COUNT(BillTransactionBySettlementDate.BillType) AS 筆數, 
    SUM(BillTransactionBySettlementDate.FaceValue) AS 總面額
FROM 
    BillTransactionBySettlementDate
WHERE
    BillTransactionBySettlementDate.TransactionType NOT IN ('兌償/到期還本', '附買回履約', '附買回解約', '附賣回履約', '附賣回解約')
    AND BillTransactionBySettlementDate.DataMonthString = [DataMonthParam]
GROUP BY 
    BillTransactionBySettlementDate.BillType;

    找 CP2 和 央行NCD
