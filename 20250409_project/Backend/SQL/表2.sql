Query 表2_DBU_AC5602_TWD
SELECT
    d.DataID,
    d.DataMonthString,
    d.AccountCode, 
    a.AccountTitle, 
    d.CurrencyType,
    d.NetBalance
FROM 
    DBU_AC5602 AS d
INNER JOIN 
    ACCAccount AS a 
ON 
    d.AccountCode = a.AccountCode
WHERE 
    d.AccountCode IN ("196017703") 
    AND d.CurrencyType = "TWD"
    AND d.DataMonthString = "2024/11";


PARAMETERS DataMonthParam TEXT;
SELECT 
    DBU_AC5602.DataID,
    DBU_AC5602.DataMonthString,
    DBU_AC5602.AccountCode, 
    ACCAccount.AccountTitle, 
    DBU_AC5602.CurrencyType,
    DBU_AC5602.NetBalance
FROM 
    DBU_AC5602 
INNER JOIN 
    ACCAccount 
ON 
    DBU_AC5602.AccountCode = ACCAccount.AccountCode
WHERE 
    DBU_AC5602.AccountCode IN ("196017703")
		AND DBU_AC5602.CurrencyType = "TWD"
    AND DBU_AC5602.DataMonthString = [DataMonthParam];




new 需要篩出最後一個DataDate
Query 表2_CloseRate_USDTWD

SELECT
		CloseRate.BaseCurrency,
		CloseRate.QuoteCurrency,
		CloseRate.Rate
FROM 
    CloseRate
WHERE
		CloseRate.BaseCurrency = "USD"
		AND CloseRate.QuoteCurrency = "TWD"
		AND CloseRate.DataMonthString = "2025/03"
    AND CloseRate.DataDate = (
        SELECT MAX(clsLast.DataDate)
        FROM CloseRate AS clsLast
        WHERE 
            clsLast.DataMonthString = "2025/03" 
            AND clsLast.BaseCurrency = "USD" 
            AND clsLast.QuoteCurrency = "TWD"
    );



PARAMETERS DataMonthParam TEXT;
SELECT
		CloseRate.BaseCurrency,
		CloseRate.QuoteCurrency,
		CloseRate.Rate
FROM 
    CloseRate
WHERE
		CloseRate.BaseCurrency = "USD"
		AND CloseRate.QuoteCurrency = "TWD"
		AND CloseRate.DataMonthString = [DataMonthParam]
    AND CloseRate.DataDate = (
    SELECT MAX(clsLast.DataDate)
    FROM CloseRate AS clsLast
    WHERE 
        clsLast.DataMonthString = [DataMonthParam]
        AND clsLast.BaseCurrency = "USD" 
        AND clsLast.QuoteCurrency = "TWD"
);
