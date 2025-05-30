1.更新餘額C
4.留意是否有新商品交易

- 金額
    - 新臺幣商業本票
        - 新臺幣融資性本票
=VLOOKUP(120050903,'C'!A:C,3,0)+
VLOOKUP(1250103,'C'!A:C,3,0)

120050903
1250103

S欄位 面額
    - 央行可轉讓定期存單
=VLOOKUP(1211109,'C'!A:C,3,0)+
VLOOKUP(1220109,'C'!A:C,3,0)


1211109
1220109




ai430		
		
- 金額		
    - 新臺幣商業本票		
        - 新臺幣融資性本票		
120050903   強制FVPL金融資產-商業本票                         
1250103 附賣回票券及債券投資-票券                         
		
S欄位 面額		
    - 央行可轉讓定期存單		
1211109 FVOCI債務工具-票券                                
1220109 AC債務工具投資-票券                               
		


PARAMETERS DataMonthParam TEXT;
SELECT
    AccountCodeMap.GroupFlag,
    AccountCodeMap.AccountTitle,
    AccountCodeMap.AssetType,
    ab.Amount
FROM AccountCodeMap
INNER JOIN
    (
        SELECT AccountBalance.AccountCode, AccountBalance.Amount
        FROM AccountBalance
        WHERE AccountBalance.DataMonthString = [DataMonthParam]
        AND AccountBalance.BalanceType = '餘額C'
    ) AS ab
ON
    AccountCodeMap.AccountCode = ab.AccountCode
WHERE
    AccountCodeMap.GroupFlag IN ('Bill', 'CP', 'RPRS')
    AND AccountCodeMap.Category IN ('Cost')


    
