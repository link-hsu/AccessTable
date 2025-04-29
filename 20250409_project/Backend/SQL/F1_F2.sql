Query F1_CM2810_LIST

SELECT DataID, DataMonthString, DealID, ContractDate, CurrencyType, Amount, StandardMoney, StandardMoneyAmount
FROM DBU_CM2810 AS cm
WHERE cm.DealID IN
    (
        SELECT sub.DealID
        FROM DBU_CM2810 AS sub
        WHERE sub.DataMonthString = '2024/11'
        GROUP BY sub.DealID
        HAVING 
            COUNT(*) = 2
            AND SUM(IIf([sub].[CurrencyType] = 'USD', 1, 0)) = 1
    )
    AND cm.DataMonthString = '2024/11';




PARAMETERS DataMonthParam TEXT;
SELECT DataID, DataMonthString, DealID, ContractDate, CurrencyType, Amount, StandardMoney, StandardMoneyAmount
FROM DBU_CM2810 AS cm
WHERE cm.DealID IN
    (
        SELECT sub.DealID
        FROM DBU_CM2810 AS sub
        WHERE sub.DataMonthString = [DataMonthParam]
        GROUP BY sub.DealID
        HAVING 
            COUNT(*) = 2
            AND SUM(IIf([sub].[CurrencyType] = 'USD', 1, 0)) = 1
    )
    AND cm.DataMonthString = [DataMonthParam];





Query F1_CM2810_Subtotal
    
SELECT 
    NonUSD.CurrencyType AS PairCurrency,
    Sum( Abs(USD.Amount) )     AS SumOfUSD,
    Sum( Abs(NonUSD.Amount) )  AS SumOfPairCurrency
FROM 
    DBU_CM2810 AS USD
    INNER JOIN DBU_CM2810 AS NonUSD
        ON USD.DealID = NonUSD.DealID
WHERE
    USD.CurrencyType = 'USD'
    AND NonUSD.CurrencyType <> 'USD'
    AND USD.DataMonthString = '2024/11'
    AND NonUSD.DataMonthString = '2024/11'
    AND USD.DealID IN
    (
        SELECT sub.DealID
        FROM DBU_CM2810 AS sub
        WHERE sub.DataMonthString = '2024/11'
        GROUP BY sub.DealID
        HAVING 
            COUNT(*) = 2
            AND SUM(IIf([sub].[CurrencyType] = 'USD', 1, 0)) = 1
    )
GROUP BY
    NonUSD.CurrencyType;
    


PARAMETERS DataMonthParam TEXT;
SELECT 
    NonUSD.CurrencyType AS PairCurrency,
    Sum( Abs(USD.Amount) )     AS SumOfUSD,
    Sum( Abs(NonUSD.Amount) )  AS SumOfPairCurrency
FROM 
    DBU_CM2810 AS USD
    INNER JOIN DBU_CM2810 AS NonUSD
        ON USD.DealID = NonUSD.DealID
WHERE
    USD.CurrencyType = 'USD'
    AND NonUSD.CurrencyType <> 'USD'
    AND USD.DataMonthString = [DataMonthParam]
    AND NonUSD.DataMonthString = [DataMonthParam]
    AND USD.DealID IN
    (
        SELECT sub.DealID
        FROM DBU_CM2810 AS sub
        WHERE sub.DataMonthString = [DataMonthParam]
        GROUP BY sub.DealID
        HAVING 
            COUNT(*) = 2
            AND SUM(IIf([sub].[CurrencyType] = 'USD', 1, 0)) = 1
    )
GROUP BY
    NonUSD.CurrencyType; 





增加日期篩選



Query F1_Foreign_DL6850_FS


SELECT 
    IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'USD' OR DBU_DL6850.SellCurrency = 'USD')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = "2024/11"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
															  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.BuyCurrency, Null));
        



PARAMETERS DataMonthParam TEXT;
SELECT 
    IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'USD' OR DBU_DL6850.SellCurrency = 'USD')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = [DataMonthParam]
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                                AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.BuyCurrency, Null));     
        
        
Query F1_Foreign_DL6850_SS
        
SELECT 
    IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'USD' OR DBU_DL6850.SellCurrency = 'USD')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = "2024/11"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
													  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.BuyCurrency, Null));


        
PARAMETERS DataMonthParam TEXT;        
SELECT 
    IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'USD' OR DBU_DL6850.SellCurrency = 'USD')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = [DataMonthParam]
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                            AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.BuyCurrency, Null));
     
        

Query F1_Domestic_DL6850_FS


SELECT 
    IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"
    AND (DBU_DL6850.BuyCurrency = 'USD' OR DBU_DL6850.SellCurrency = 'USD')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = "2024/11"
    AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
													  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.BuyCurrency, Null));
        


                      
PARAMETERS DataMonthParam TEXT;
SELECT 
    IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"
    AND (DBU_DL6850.BuyCurrency = 'USD' OR DBU_DL6850.SellCurrency = 'USD')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = [DataMonthParam]
		AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                            AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.BuyCurrency, Null));


        
        
Query F1_Domestic_DL6850_SS
        
SELECT 
    IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"		
    AND (DBU_DL6850.BuyCurrency = 'USD' OR DBU_DL6850.SellCurrency = 'USD')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = "2024/11"
		AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
													  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.BuyCurrency, Null));


        
PARAMETERS DataMonthParam TEXT;
SELECT 
    IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"		
    AND (DBU_DL6850.BuyCurrency = 'USD' OR DBU_DL6850.SellCurrency = 'USD')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = [DataMonthParam]
		AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                            AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    IIf(DBU_DL6850.BuyCurrency = 'USD', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'USD', DBU_DL6850.BuyCurrency, Null));      
        



=========================================================
Query F1_Foreign_DL6850_FS


SELECT 
    "EUR_" & IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'EUR' OR DBU_DL6850.SellCurrency = 'EUR')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = "2024/11"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
															  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    "EUR_" & IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.BuyCurrency, Null))

UNION ALL

SELECT 
    "GBP_" & IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'GBP' OR DBU_DL6850.SellCurrency = 'GBP')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = "2024/11"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
															  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    "GBP_" & IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.BuyCurrency, Null))

UNION ALL

SELECT 
    "JPY_" & IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'JPY' OR DBU_DL6850.SellCurrency = 'JPY')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = "2024/11"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
															  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    "JPY_" & IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.BuyCurrency, Null))

UNION ALL

SELECT 
    "CNY_" & IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'CNY' OR DBU_DL6850.SellCurrency = 'CNY')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = "2024/11"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
															  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    "CNY_" & IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.BuyCurrency, Null))


PARAMETERS DataMonthParam TEXT;
SELECT 
    "EUR_" & IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'EUR' OR DBU_DL6850.SellCurrency = 'EUR')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = [DataMonthParam]
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                                AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    "EUR_" & IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.BuyCurrency, Null));     
        
UNION ALL

SELECT 
    "GBP_" & IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'GBP' OR DBU_DL6850.SellCurrency = 'GBP')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = [DataMonthParam]
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                                AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    "GBP_" & IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.BuyCurrency, Null));     
        
UNION ALL

SELECT 
    "JPY_" & IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'JPY' OR DBU_DL6850.SellCurrency = 'JPY')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = [DataMonthParam]
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                                AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    "JPY_" & IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.BuyCurrency, Null));     
        
UNION ALL

SELECT 
    "CNY_" & IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'CNY' OR DBU_DL6850.SellCurrency = 'CNY')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = [DataMonthParam]
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                                AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    "CNY_" & IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.BuyCurrency, Null));     

Query F1_Foreign_DL6850_SS
        
SELECT 
    "EUR_" & IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'EUR' OR DBU_DL6850.SellCurrency = 'EUR')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = "2024/11"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
													  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    "EUR_" & IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.BuyCurrency, Null));

UNION ALL

SELECT 
    "GBP_" & IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'GBP' OR DBU_DL6850.SellCurrency = 'GBP')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = "2024/11"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
													  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    "GBP_" & IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.BuyCurrency, Null));

UNION ALL

SELECT 
    "JPY_" & IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'JPY' OR DBU_DL6850.SellCurrency = 'JPY')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = "2024/11"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
													  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    "JPY_" & IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.BuyCurrency, Null));

UNION ALL

SELECT 
    "CNY_" & IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'CNY' OR DBU_DL6850.SellCurrency = 'CNY')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = "2024/11"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
													  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    "CNY_" & IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.BuyCurrency, Null));

        
PARAMETERS DataMonthParam TEXT;        
SELECT 
    "EUR_" & IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'EUR' OR DBU_DL6850.SellCurrency = 'EUR')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = [DataMonthParam]
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                            AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    "EUR_" & IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.BuyCurrency, Null));
     
UNION ALL       

SELECT 
    "GBP_" & IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'GBP' OR DBU_DL6850.SellCurrency = 'GBP')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = [DataMonthParam]
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                            AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    "GBP_" & IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.BuyCurrency, Null));
     
UNION ALL

SELECT 
    "JPY_" & IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'JPY' OR DBU_DL6850.SellCurrency = 'JPY')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = [DataMonthParam]
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                            AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    "JPY_" & IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.BuyCurrency, Null));
     
UNION ALL

SELECT 
    "CNY_" & IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Foreign_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty IN ('GSILGB2X', 'BBVAESMMFXD')
    AND (DBU_DL6850.BuyCurrency = 'CNY' OR DBU_DL6850.SellCurrency = 'CNY')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = [DataMonthParam]
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                            AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    "CNY_" & IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.BuyCurrency, Null));


























=====================================================================




Query F1_Domestic_DL6850_FS


SELECT 
    "EUR_" & IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"
    AND (DBU_DL6850.BuyCurrency = 'EUR' OR DBU_DL6850.SellCurrency = 'EUR')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = "2024/11"
    AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
													  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    "EUR_" & IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.BuyCurrency, Null));
        
UNION ALL

SELECT 
    "GBP_" & IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"
    AND (DBU_DL6850.BuyCurrency = 'GBP' OR DBU_DL6850.SellCurrency = 'GBP')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = "2024/11"
    AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
													  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    "GBP_" & IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.BuyCurrency, Null));
        
UNION ALL

SELECT 
    "JPY_" & IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"
    AND (DBU_DL6850.BuyCurrency = 'JPY' OR DBU_DL6850.SellCurrency = 'JPY')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = "2024/11"
    AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
													  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    "JPY_" & IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.BuyCurrency, Null));
        
UNION ALL

SELECT 
    "CNY_" & IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"
    AND (DBU_DL6850.BuyCurrency = 'CNY' OR DBU_DL6850.SellCurrency = 'CNY')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = "2024/11"
    AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
													  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    "CNY_" & IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.BuyCurrency, Null));
        


                      
PARAMETERS DataMonthParam TEXT;
SELECT 
    "EUR_" & IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"
    AND (DBU_DL6850.BuyCurrency = 'EUR' OR DBU_DL6850.SellCurrency = 'EUR')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = [DataMonthParam]
		AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                            AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    "EUR_" & IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.BuyCurrency, Null));

UNION ALL

SELECT 
    "GBP_" & IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"
    AND (DBU_DL6850.BuyCurrency = 'GBP' OR DBU_DL6850.SellCurrency = 'GBP')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = [DataMonthParam]
		AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                            AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    "GBP_" & IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.BuyCurrency, Null));

UNION ALL

SELECT 
    "JPY_" & IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"
    AND (DBU_DL6850.BuyCurrency = 'JPY' OR DBU_DL6850.SellCurrency = 'JPY')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = [DataMonthParam]
		AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                            AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    "JPY_" & IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.BuyCurrency, Null));

UNION ALL

SELECT 
    "CNY_" & IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_FS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"
    AND (DBU_DL6850.BuyCurrency = 'CNY' OR DBU_DL6850.SellCurrency = 'CNY')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'FS'
    AND DBU_DL6850.DataMonthString  = [DataMonthParam]
		AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                            AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    "CNY_" & IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.BuyCurrency, Null));


        
        
Query F1_Domestic_DL6850_SS
        
SELECT 
    "EUR_" & IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"		
    AND (DBU_DL6850.BuyCurrency = 'EUR' OR DBU_DL6850.SellCurrency = 'EUR')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = "2024/11"
		AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
													  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    "EUR_" & IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.BuyCurrency, Null));

UNION ALL

SELECT 
    "GBP_" & IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"		
    AND (DBU_DL6850.BuyCurrency = 'GBP' OR DBU_DL6850.SellCurrency = 'GBP')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = "2024/11"
		AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
													  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    "GBP_" & IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.BuyCurrency, Null));

UNION ALL

SELECT 
    "JPY_" & IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"		
    AND (DBU_DL6850.BuyCurrency = 'JPY' OR DBU_DL6850.SellCurrency = 'JPY')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = "2024/11"
		AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
													  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    "JPY_" & IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.BuyCurrency, Null));

UNION ALL

SELECT 
    "CNY_" & IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"		
    AND (DBU_DL6850.BuyCurrency = 'CNY' OR DBU_DL6850.SellCurrency = 'CNY')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = "2024/11"
		AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate("2024/11" & "/01")
													  AND DateSerial(Year(CDate("2024/11" & "/01")), Month(CDate("2024/11" & "/01")) + 1, 0)
GROUP BY 
    "CNY_" & IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.BuyCurrency, Null));


        
PARAMETERS DataMonthParam TEXT;
SELECT 
    "EUR_" & IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"		
    AND (DBU_DL6850.BuyCurrency = 'EUR' OR DBU_DL6850.SellCurrency = 'EUR')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = [DataMonthParam]
		AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                            AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    "EUR_" & IIf(DBU_DL6850.BuyCurrency = 'EUR', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'EUR', DBU_DL6850.BuyCurrency, Null));      
        
UNION ALL

SELECT 
    "GBP_" & IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"		
    AND (DBU_DL6850.BuyCurrency = 'GBP' OR DBU_DL6850.SellCurrency = 'GBP')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = [DataMonthParam]
		AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                            AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    "GBP_" & IIf(DBU_DL6850.BuyCurrency = 'GBP', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'GBP', DBU_DL6850.BuyCurrency, Null));      
        
UNION ALL

SELECT 
    "JPY_" & IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"		
    AND (DBU_DL6850.BuyCurrency = 'JPY' OR DBU_DL6850.SellCurrency = 'JPY')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = [DataMonthParam]
		AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                            AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    "JPY_" & IIf(DBU_DL6850.BuyCurrency = 'JPY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'JPY', DBU_DL6850.BuyCurrency, Null));      
        
UNION ALL

SELECT 
    "CNY_" & IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.BuyCurrency, Null)) AS Curr,
    SUM(IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.BuyAmount, 
            IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.SellAmount, 0))) AS TotalAmount_Domestic_SS
FROM DBU_DL6850
WHERE DBU_DL6850.Counterparty NOT IN ('GSILGB2X', 'BBVAESMMFXD')
		'HTBKTWTPOBU要不要納入這部分要再確認
		'AND DBU_DL6850.Counterparty <> "HTBKTWTPOBU"		
    AND (DBU_DL6850.BuyCurrency = 'CNY' OR DBU_DL6850.SellCurrency = 'CNY')
    AND MID(DBU_DL6850.DealID, 5, 2) = 'SS'
		AND DBU_DL6850.DataMonthString  = [DataMonthParam]
		AND DBU_DL6850.BuyCurrency  <> "TWD"
    AND DBU_DL6850.SellCurrency  <> "TWD"
    AND DBU_DL6850.ContractDate BETWEEN CDate([DataMonthParam] & "/01") 
                            AND DateSerial(Year(CDate([DataMonthParam] & "/01")), Month(CDate([DataMonthParam] & "/01")) + 1, 0)
GROUP BY 
    "CNY_" & IIf(DBU_DL6850.BuyCurrency = 'CNY', DBU_DL6850.SellCurrency, 
        IIf(DBU_DL6850.SellCurrency = 'CNY', DBU_DL6850.BuyCurrency, Null));      
