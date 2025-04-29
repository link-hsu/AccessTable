
Query AI821_OBU_MM4901B_LIST
SELECT
    OBU_MM4901B.DataID,
    OBU_MM4901B.DataMonthString,
    OBU_MM4901B.DealID,
    OBU_MM4901B.CounterParty,
    BankType.BankTypeName,
    BankDirectory.BankTypeCode,
    OBU_MM4901B.CurrencyType,
    OBU_MM4901B.Amount
FROM
    (OBU_MM4901B
        INNER JOIN BankDirectory
        ON OBU_MM4901B.CounterParty = BankDirectory.SWIFT)
        INNER JOIN BankType
        ON BankDirectory.BankTypeCode = BankType.BankTypeCode
WHERE OBU_MM4901B.CurrencyType = "CNY"
  And OBU_MM4901B.DataMonthString = "2024/11";



PARAMETERS DataMonthParam TEXT;
SELECT
    OBU_MM4901B.DataID,
    OBU_MM4901B.DataMonthString,
    OBU_MM4901B.DealID,
    OBU_MM4901B.CounterParty,
    BankType.BankTypeName,
    BankDirectory.BankTypeCode,
    OBU_MM4901B.CurrencyType,
    OBU_MM4901B.Amount
FROM
    (OBU_MM4901B
        INNER JOIN BankDirectory
        ON OBU_MM4901B.CounterParty = BankDirectory.SWIFT)
        INNER JOIN BankType
        ON BankDirectory.BankTypeCode = BankType.BankTypeCode
WHERE OBU_MM4901B.CurrencyType = "CNY"
  And OBU_MM4901B.DataMonthString = [DataMonthParam];






Query AI821_OBU_MM4901B_SUM

SELECT
    bt.BankTypeName,
    bd.BankTypeCode,
    SUM(o.Amount) AS TotalAmount
FROM
    (OBU_MM4901B AS o
        INNER JOIN BankDirectory AS bd
        ON o.CounterParty = bd.SWIFT)
        INNER JOIN BankType AS bt
        ON bd.BankTypeCode = bt.BankTypeCode
WHERE o.CurrencyType = "CNY"
  AND o.DataMonthString = "2024/11"
GROUP BY bd.BankTypeCode, bt.BankTypeName;



PARAMETERS DataMonthParam TEXT;
SELECT
    BankType.BankTypeName,
    BankDirectory.BankTypeCode,
    SUM(OBU_MM4901B.Amount) AS TotalAmount
FROM
    (OBU_MM4901B
        INNER JOIN BankDirectory
        ON OBU_MM4901B.CounterParty = BankDirectory.SWIFT)
        INNER JOIN BankType
        ON BankDirectory.BankTypeCode = BankType.BankTypeCode
WHERE OBU_MM4901B.CurrencyType = "CNY"
  AND OBU_MM4901B.DataMonthString = [DataMonthParam]
GROUP BY BankDirectory.BankTypeCode, BankType.BankTypeName;


