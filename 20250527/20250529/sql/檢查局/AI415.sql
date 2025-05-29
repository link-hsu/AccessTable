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


PARAMETERS DataMonthParam TEXT;
SELECT 
    BillTransactionDetails.BillType AS 類型, 
    COUNT(BillTransactionDetails.BillType) AS 筆數, 
    SUM(BillTransactionDetails.FaceValue) AS 總面額
FROM 
    BillTransactionDetails
WHERE 
    BillTransactionDetails.DataMonthString = [DataMonthParam]
GROUP BY 
    BillTransactionDetails.BillType;

    找 CP2 和 央行NCD
