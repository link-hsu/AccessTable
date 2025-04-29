
Query FB3_OBU_MM4901B_LIST


SELECT
    OBU_MM4901B.DataID,
    OBU_MM4901B.DataMonthString,
    OBU_MM4901B.DealDate, 
    OBU_MM4901B.DealID, 
    OBU_MM4901B.CounterParty, 
    OBU_MM4901B.MaturityDate,
    OBU_MM4901B.CurrencyType,
    OBU_MM4901B.Amount,
    IIF(MID(OBU_MM4901B.DealID, 5, 2) = "MP", "MP", "MT") AS Category
FROM
    OBU_MM4901B
WHERE
    OBU_MM4901B.CurrencyType = "CNY" 
    AND OBU_MM4901B.DataMonthString = "2024/11";


----


PARAMETERS DataMonthParam TEXT;
SELECT
    OBU_MM4901B.DataID,
    OBU_MM4901B.DataMonthString,
    OBU_MM4901B.DealDate, 
    OBU_MM4901B.DealID, 
    OBU_MM4901B.CounterParty, 
    OBU_MM4901B.MaturityDate,
    OBU_MM4901B.CurrencyType,
    OBU_MM4901B.Amount,
    IIF(MID(OBU_MM4901B.DealID, 5, 2) = "MP", "MP", "MT") AS Category
FROM
    OBU_MM4901B
WHERE
    OBU_MM4901B.CurrencyType = "CNY" 
    AND OBU_MM4901B.DataMonthString = [DataMonthParam];



Query FB3_OBU_MM4901B_SUM


SELECT 
    SUM(IIF(MID(OBU_MM4901B.DealID, 5, 2) = "MP", OBU_MM4901B.Amount, 0)) AS Sum_MP,
    SUM(IIF(MID(OBU_MM4901B.DealID, 5, 2) = "MT", OBU_MM4901B.Amount, 0)) AS Sum_MT
FROM OBU_MM4901B 
WHERE 
    OBU_MM4901B.CurrencyType = "CNY" 
    AND OBU_MM4901B.DataMonthString = "2024/11";





PARAMETERS DataMonthParam TEXT;
SELECT 
    SUM(IIF(MID(OBU_MM4901B.DealID, 5, 2) = "MP", OBU_MM4901B.Amount, 0)) AS Sum_MP,
    SUM(IIF(MID(OBU_MM4901B.DealID, 5, 2) = "MT", OBU_MM4901B.Amount, 0)) AS Sum_MT
FROM OBU_MM4901B 
WHERE 
    OBU_MM4901B.CurrencyType = "CNY" 
    AND OBU_MM4901B.DataMonthString = [DataMonthParam];









Query FB3A_OBU_MM4901B 


SELECT 
    OBU_MM4901B.DataID,
    OBU_MM4901B.DataMonthString, 
    OBU_MM4901B.DealDate, 
    OBU_MM4901B.DealID, 
    OBU_MM4901B.CounterParty, 
    OBU_MM4901B.MaturityDate, 
    OBU_MM4901B.CurrencyType, 
    OBU_MM4901B.Amount,
    IIF(RIGHT(OBU_MM4901B.CounterParty, 3) <> "OBU" AND MID(OBU_MM4901B.DealID, 5, 2) = "MP", "DBU_MP",
        IIF(RIGHT(OBU_MM4901B.CounterParty, 3) <> "OBU" AND MID(OBU_MM4901B.DealID, 5, 2) = "MT", "DBU_MT",
            IIF(RIGHT(OBU_MM4901B.CounterParty, 3) = "OBU" AND MID(OBU_MM4901B.DealID, 5, 2) = "MP", "OBU_MP",
                "OBU_MT"
            )
        )
    ) AS Category,
    bk.BankCode
FROM
    OBU_MM4901B
LEFT JOIN BankDirectory As bk
    ON OBU_MM4901B.CounterParty = bk.SWIFT
WHERE 
    OBU_MM4901B.CurrencyType = "CNY" 
    AND OBU_MM4901B.DataMonthString = "2024/11";
    



PARAMETERS DataMonthParam TEXT;   
SELECT 
    OBU_MM4901B.DataID,
    OBU_MM4901B.DataMonthString, 
    OBU_MM4901B.DealDate, 
    OBU_MM4901B.DealID, 
    OBU_MM4901B.CounterParty, 
    OBU_MM4901B.MaturityDate, 
    OBU_MM4901B.CurrencyType, 
    OBU_MM4901B.Amount,
    IIF(RIGHT(OBU_MM4901B.CounterParty, 3) <> "OBU" AND MID(OBU_MM4901B.DealID, 5, 2) = "MP", "DBU_MP",
        IIF(RIGHT(OBU_MM4901B.CounterParty, 3) <> "OBU" AND MID(OBU_MM4901B.DealID, 5, 2) = "MT", "DBU_MT",
            IIF(RIGHT(OBU_MM4901B.CounterParty, 3) = "OBU" AND MID(OBU_MM4901B.DealID, 5, 2) = "MP", "OBU_MP",
                "OBU_MT"
            )
        )
    ) AS Category,
    bk.BankCode
FROM
    OBU_MM4901B
LEFT JOIN BankDirectory As bk
    ON OBU_MM4901B.CounterParty = bk.SWIFT
WHERE 
    OBU_MM4901B.CurrencyType = "CNY" 
    AND OBU_MM4901B.DataMonthString = [DataMonthParam];
