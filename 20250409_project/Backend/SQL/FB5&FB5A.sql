Query FB5_DL6320

SELECT
    OBU_DL6320.DataID,
    OBU_DL6320.DataMonthString,
    OBU_DL6320.DealID,
    OBU_DL6320.SellCurrency,
    OBU_DL6320.SellPrice,
    OBU_DL6320.DealDate,
    OBU_DL6320.SettlementDate
FROM 
    OBU_DL6320 
WHERE 
    OBU_DL6320.SellCurrency= "CNY" 
    AND OBU_DL6320.DataMonthString = "2024/11";



PARAMETERS DataMonthParam TEXT;
SELECT 
    OBU_DL6320.DataID,
    OBU_DL6320.DataMonthString,
    OBU_DL6320.DealID,
    OBU_DL6320.SellCurrency,
    OBU_DL6320.SellPrice,
    OBU_DL6320.DealDate,
    OBU_DL6320.SettlementDate
FROM 
    OBU_DL6320 
WHERE 
    OBU_DL6320.SellCurrency= "CNY" 
    AND OBU_DL6320.DataMonthString = [DataMonthParam];
