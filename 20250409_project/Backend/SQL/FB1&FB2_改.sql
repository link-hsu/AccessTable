Query FB2_OBU_AC4620B

PARAMETERS DataMonthParam TEXT;
SELECT
    OBU_AC4620B.DataID,
    OBU_AC4620B.DataMonthString, 
    OBU_AC4620B.AccountCode, 
    ACCAccount.AccountTitle, 
    OBU_AC4620B.NetBalance,
    OBU_AC4620B.CurrencyType
FROM 
    OBU_AC4620B
INNER JOIN 
    ACCAccount
ON 
    OBU_AC4620B.AccountCode = ACCAccount.AccountCode
WHERE 
    OBU_AC4620B.AccountCode IN ("115037101", "115037105", "115037115", "130152771", "130152773", "130152777") 
    AND OBU_AC4620B.CurrencyType = "CNY" 
    AND OBU_AC4620B.DataMonthString = [DataMonthParam];
    


SELECT
    d.DataID,
    d.DataMonthString,
    d.AccountCode, 
    a.AccountTitle, 
    d.NetBalance,
    d.CurrencyType
FROM 
    OBU_AC4620B AS d
INNER JOIN 
    ACCAccount AS a 
ON 
    d.AccountCode = a.AccountCode
WHERE     
    d.AccountCode IN ("115037101", "115037105", "115037115", "130152771", "130152773", "130152777") 
    AND d.CurrencyType = "CNY" 
    AND d.DataMonthString = "2024/11";







New
Query FB1_OBU_AC4620B_Subtotal


SELECT
    AccountCodeMap.AssetMeasurementType,
    SUM(oa.NetBalance) As SubtotalBalance
FROM AccountCodeMap
INNER JOIN
    (
        SELECT OBU_AC4620B.AccountCode, OBU_AC4620B.NetBalance
        FROM OBU_AC4620B
        WHERE OBU_AC4620B.DataMonthString = "2024/11"
        AND OBU_AC4620B.CurrencyType = "CNY"
    ) AS oa
ON
    AccountCodeMap.AccountCode = oa.AccountCode
WHERE
    AccountCodeMap.Category IN ("Cost" , "ValuationAdjust")
GROUP BY
    AccountCodeMap.AssetMeasurementType;


SELECT
    AccountCodeMap.AssetMeasurementType,
    SUM(OBU_AC4620B.NetBalance) As SubtotalBalance
FROM AccountCodeMap
INNER JOIN
    OBU_AC4620B
ON
    AccountCodeMap.AccountCode = OBU_AC4620B.AccountCode
WHERE
    AccountCodeMap.Category IN ("Cost" , "ValuationAdjust")
    AND OBU_AC4620B.DataMonthString = "2024/11"
    AND OBU_AC4620B.CurrencyType = "CNY"
GROUP BY
    AccountCodeMap.AssetMeasurementType;



PARAMETERS DataMonthParam TEXT;
SELECT
    AccountCodeMap.AssetMeasurementType,
    SUM(oa.NetBalance) As SubtotalBalance
FROM AccountCodeMap
INNER JOIN
    (
        SELECT OBU_AC4620B.AccountCode, OBU_AC4620B.NetBalance
        FROM OBU_AC4620B
        WHERE OBU_AC4620B.DataMonthString = [DataMonthParam]
        AND OBU_AC4620B.CurrencyType = "CNY"
    ) AS oa
ON
    AccountCodeMap.AccountCode = oa.AccountCode
WHERE
    AccountCodeMap.Category IN ('Cost' , 'ValuationAdjust')
GROUP BY
    AccountCodeMap.AssetMeasurementType;


PARAMETERS DataMonthParam TEXT;
SELECT
    AccountCodeMap.AssetMeasurementType,
    SUM(OBU_AC4620B.NetBalance) As SubtotalBalance
FROM AccountCodeMap
INNER JOIN
    OBU_AC4620B
ON
    AccountCodeMap.AccountCode = OBU_AC4620B.AccountCode
WHERE
    AccountCodeMap.Category IN ('Cost' , 'ValuationAdjust')
    AND OBU_AC4620B.DataMonthString = [DataMonthParam]
    AND OBU_AC4620B.CurrencyType = "CNY"
GROUP BY
    AccountCodeMap.AssetMeasurementType;



Old
Query FB1_OBU_AC4620B_Subtotal

SELECT IIf(OBU_AC4620B.AccountCode IN ("121110105", "121110125", "121110127", "121110147"), 'FVOCI_VALUE', 
        IIf(OBU_AC4620B.AccountCode IN ("121130105", "121130125", "121130127", "121130147"), 'FVOCI_ADJUSTMENT', 
            IIf(OBU_AC4620B.AccountCode IN ("122010105", "122010125", "122010127", "122010147"), 'AC_VALUE', 
                IIf(OBU_AC4620B.AccountCode IN ("122030105", "122030125", "122030127", "122030147"), 'AC_ADJUSTMENT', 'Other')
            )
        )
    ) AS AccountGroup, Sum(OBU_AC4620B.NetBalance) AS SubtotalNetBalance
FROM OBU_AC4620B
WHERE OBU_AC4620B.AccountCode IN ("121110105", "121110125", "121110127", "121110147", 
                                "121130105", "121130125", "121130127", "121130147", 
                                "122010105", "122010125", "122010127", "122010147", 
                                "122030105", "122030125", "122030127", "122030147") 
    AND OBU_AC4620B.CurrencyType = "CNY" 
    AND OBU_AC4620B.DataMonthString = "2024/11"
GROUP BY IIf(OBU_AC4620B.AccountCode IN ("121110105", "121110125", "121110127", "121110147"), 'FVOCI_VALUE', 
        IIf(OBU_AC4620B.AccountCode IN ("121130105", "121130125", "121130127", "121130147"), 'FVOCI_ADJUSTMENT', 
            IIf(OBU_AC4620B.AccountCode IN ("122010105", "122010125", "122010127", "122010147"), 'AC_VALUE', 
                IIf(OBU_AC4620B.AccountCode IN ("122030105", "122030125", "122030127", "122030147"), 'AC_ADJUSTMENT', 'Other')
            )
        )
    );



PARAMETERS DataMonthParam TEXT;

SELECT IIf(OBU_AC4620B.AccountCode IN ("121110105", "121110125", "121110127", "121110147"), 'FVOCI_VALUE', 
        IIf(OBU_AC4620B.AccountCode IN ("121130105", "121130125", "121130127", "121130147"), 'FVOCI_ADJUSTMENT', 
            IIf(OBU_AC4620B.AccountCode IN ("122010105", "122010125", "122010127", "122010147"), 'AC_VALUE', 
                IIf(OBU_AC4620B.AccountCode IN ("122030105", "122030125", "122030127", "122030147"), 'AC_ADJUSTMENT', 'Other')
            )
        )
    ) AS AccountGroup, Sum(OBU_AC4620B.NetBalance) AS SubtotalNetBalance
FROM OBU_AC4620B
WHERE OBU_AC4620B.AccountCode IN ("121110105", "121110125", "121110127", "121110147", 
                                "121130105", "121130125", "121130127", "121130147", 
                                "122010105", "122010125", "122010127", "122010147", 
                                "122030105", "122030125", "122030127", "122030147") 
    AND OBU_AC4620B.CurrencyType = "CNY" 
    AND OBU_AC4620B.DataMonthString = [DataMonthParam]
GROUP BY IIf(OBU_AC4620B.AccountCode IN ("121110105", "121110125", "121110127", "121110147"), 'FVOCI_VALUE', 
        IIf(OBU_AC4620B.AccountCode IN ("121130105", "121130125", "121130127", "121130147"), 'FVOCI_ADJUSTMENT', 
            IIf(OBU_AC4620B.AccountCode IN ("122010105", "122010125", "122010127", "122010147"), 'AC_VALUE', 
                IIf(OBU_AC4620B.AccountCode IN ("122030105", "122030125", "122030127", "122030147"), 'AC_ADJUSTMENT', 'Other')
            )
        )
    );
