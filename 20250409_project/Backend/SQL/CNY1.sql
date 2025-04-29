Query CNY1_DBU_AC5601

SELECT
    d.DataID,
    d.DataMonthString,
    d.AccountCode, 
    a.AccountTitle, 
    d.NetBalance
FROM 
    DBU_AC5601 AS d
INNER JOIN 
    ACCAccount AS a 
ON 
    d.AccountCode = a.AccountCode
WHERE 
    d.AccountCode IN ("155930402", "255930402") 
    AND d.CurrencyType = "CNY" 
    AND d.DataMonthString = "2025/02";




PARAMETERS DataMonthParam TEXT;
SELECT
    DBU_AC5601.DataID,
    DBU_AC5601.DataMonthString,
    DBU_AC5601.AccountCode, 
    ACCAccount.AccountTitle, 
    DBU_AC5601.NetBalance
FROM 
    DBU_AC5601 
INNER JOIN 
    ACCAccount 
ON 
    DBU_AC5601.AccountCode = ACCAccount.AccountCode
WHERE 
    DBU_AC5601.AccountCode IN ("155930402", "255930402") 
    AND DBU_AC5601.CurrencyType = "CNY" 
    AND DBU_AC5601.DataMonthString = [DataMonthParam];


