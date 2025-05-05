Old
Query FM10_OBU_AC4603_LIST

SELECT
    OBU_AC4603.DataID,
    OBU_AC4603.DataMonthString,
    OBU_AC4603.AccountCode, 
    ACCAccount.AccountTitle, 
    OBU_AC4603.CurrencyType,
    OBU_AC4603.NetBalance
FROM 
    OBU_AC4603
INNER JOIN 
    ACCAccount
ON 
    OBU_AC4603.AccountCode = ACCAccount.AccountCode
WHERE 
    OBU_AC4603.AccountCode IN
      ("120050105", "120050125", "120050127",
       "120070105", "120070125", "120070127",
       "121110105", "121110125", "121110127", "121110147",
       "121130105", "121130125", "121130127", "121130147",
       "122010105", "122010125", "122010127", "122010147",
       "122030105", "122030125", "122030127", "122030147",
       "155517201") 
    AND OBU_AC4603.CurrencyType = "USD" 
    AND OBU_AC4603.DataMonthString = "2024/11";



PARAMETERS DataMonthParam TEXT;
SELECT
    OBU_AC4603.DataID,
    OBU_AC4603.DataMonthString,
    OBU_AC4603.AccountCode, 
    ACCAccount.AccountTitle, 
    OBU_AC4603.CurrencyType,
    OBU_AC4603.NetBalance
FROM 
    OBU_AC4603
INNER JOIN 
    ACCAccount
ON 
    OBU_AC4603.AccountCode = ACCAccount.AccountCode
WHERE 
    OBU_AC4603.AccountCode IN
      ("120050105", "120050125", "120050127",
       "120070105", "120070125", "120070127",
       "121110105", "121110125", "121110127", "121110147",
       "121130105", "121130125", "121130127", "121130147",
       "122010105", "122010125", "122010127", "122010147",
       "122030105", "122030125", "122030127", "122030147",
       "155517201") 
    AND OBU_AC4603.CurrencyType = "USD" 
    AND OBU_AC4603.DataMonthString = [DataMonthParam];



New
Query FM10_OBU_AC4603_LIST

SELECT
    OBU_AC4603.DataID,
    OBU_AC4603.DataMonthString,
    AccountCodeMap.AccountCode, 
    AccountCodeMap.AccountTitle, 
    OBU_AC4603.CurrencyType,
    OBU_AC4603.NetBalance
FROM 
    AccountCodeMap
INNER JOIN 
    OBU_AC4603
ON
    AccountCodeMap.AccountCode = OBU_AC4603.AccountCode
WHERE
    AccountCodeMap.GroupFlag = '外幣債'
    AND AccountCodeMap.Category IN ('Cost' , 'ValuationAdjust', 'ImpairmentLoss', 'otherFinancialAssets')
    AND OBU_AC4603.CurrencyType = 'USD' 
    AND OBU_AC4603.DataMonthString = '2024/11';

SELECT
    oa.DataID,
    oa.DataMonthString,
    AccountCodeMap.AccountCode, 
    AccountCodeMap.AccountTitle, 
    oa.CurrencyType,
    oa.NetBalance
FROM 
    AccountCodeMap
INNER JOIN 
    (
        SELECT
            OBU_AC4603.DataID,
            OBU_AC4603.DataMonthString,
            OBU_AC4603.CurrencyType,
            OBU_AC4603.NetBalance,
            OBU_AC4603.AccountCode
        FROM 
            OBU_AC4603
        WHERE
            OBU_AC4603.CurrencyType = 'USD' 
            AND OBU_AC4603.DataMonthString = '2024/11'
    ) AS oa
ON
    AccountCodeMap.AccountCode = oa.AccountCode
WHERE
    AccountCodeMap.GroupFlag = '外幣債'
    AND AccountCodeMap.Category IN ('Cost' , 'ValuationAdjust', 'ImpairmentLoss', 'otherFinancialAssets');



PARAMETERS DataMonthParam TEXT;
SELECT
    OBU_AC4603.DataID,
    OBU_AC4603.DataMonthString,
    AccountCodeMap.AccountCode, 
    AccountCodeMap.AccountTitle, 
    OBU_AC4603.CurrencyType,
    OBU_AC4603.NetBalance
FROM 
    AccountCodeMap
INNER JOIN 
    OBU_AC4603
ON
    AccountCodeMap.AccountCode = OBU_AC4603.AccountCode
WHERE
    AccountCodeMap.GroupFlag = '外幣債'
    AND AccountCodeMap.Category IN ('Cost' , 'ValuationAdjust', 'ImpairmentLoss', 'otherFinancialAssets')
    AND OBU_AC4603.CurrencyType = 'USD' 
    AND OBU_AC4603.DataMonthString = [DataMonthParam];


PARAMETERS DataMonthParam TEXT;
SELECT
    oa.DataID,
    oa.DataMonthString,
    AccountCodeMap.AccountCode, 
    AccountCodeMap.AccountTitle, 
    oa.CurrencyType,
    oa.NetBalance
FROM 
    AccountCodeMap
INNER JOIN 
    (
        SELECT
            OBU_AC4603.DataID,
            OBU_AC4603.DataMonthString,
            OBU_AC4603.CurrencyType,
            OBU_AC4603.NetBalance,
            OBU_AC4603.AccountCode
        FROM 
            OBU_AC4603
        WHERE
            OBU_AC4603.CurrencyType = 'USD' 
            AND OBU_AC4603.DataMonthString = [DataMonthParam]
    ) AS oa
ON
    AccountCodeMap.AccountCode = oa.AccountCode
WHERE
    AccountCodeMap.GroupFlag = '外幣債'
    AND AccountCodeMap.Category IN ('Cost' , 'ValuationAdjust', 'ImpairmentLoss', 'otherFinancialAssets');

-- ============================================================
-- ============================================================
Old version add FVPL
Query FM10_OBU_AC4603_Subtotal

SELECT 
  IIf(OBU_AC4603.AccountCode IN ("120050105", "120050125", "120050127"), 'FVPL_VALUE',
    IIf(OBU_AC4603.AccountCode IN ("120070105", "120070125", "120070127"), 'FVPL_ADJUSTMENT',
      IIf(OBU_AC4603.AccountCode IN ("121110105", "121110125", "121110127", "121110147"), 'FVOCI_VALUE',
        IIf(OBU_AC4603.AccountCode IN ("121130105", "121130125", "121130127", "121130147"), 'FVOCI_ADJUSTMENT',
          IIf(OBU_AC4603.AccountCode IN ("122010105", "122010125", "122010127", "122010147"), 'AC_VALUE',
            IIf(OBU_AC4603.AccountCode IN ("122030105", "122030125", "122030127", "122030147"), 'AC_ADJUSTMENT', 'Other')
          )
        )
      )
    )
  ) AS AccountGroup,
  Sum(OBU_AC4603.NetBalance) AS SubtotalNetBalance
FROM OBU_AC4603
WHERE 
  OBU_AC4603.AccountCode IN (
    "120050105","120050125","120050127",
    "120070105","120070125","120070127",
    "121110105","121110125","121110127","121110147",
    "121130105","121130125","121130127","121130147",
    "122010105","122010125","122010127","122010147",
    "122030105","122030125","122030127","122030147"
  )
  AND OBU_AC4603.CurrencyType = "USD"
  AND OBU_AC4603.DataMonthString = "2024/11"
GROUP BY 
  IIf(OBU_AC4603.AccountCode IN ("120050105", "120050125", "120050127"), 'FVPL_VALUE',
    IIf(OBU_AC4603.AccountCode IN ("120070105", "120070125", "120070127"), 'FVPL_ADJUSTMENT',
      IIf(OBU_AC4603.AccountCode IN ("121110105", "121110125", "121110127", "121110147"), 'FVOCI_VALUE',
        IIf(OBU_AC4603.AccountCode IN ("121130105", "121130125", "121130127", "121130147"), 'FVOCI_ADJUSTMENT',
          IIf(OBU_AC4603.AccountCode IN ("122010105", "122010125", "122010127", "122010147"), 'AC_VALUE',
            IIf(OBU_AC4603.AccountCode IN ("122030105", "122030125", "122030127", "122030147"), 'AC_ADJUSTMENT', 'Other')
          )
        )
      )
    )
  );





PARAMETERS DataMonthParam TEXT;
SELECT 
  IIf(OBU_AC4603.AccountCode IN ("120050105", "120050125", "120050127"), 'FVPL_VALUE',
    IIf(OBU_AC4603.AccountCode IN ("120070105", "120070125", "120070127"), 'FVPL_ADJUSTMENT',
      IIf(OBU_AC4603.AccountCode IN ("121110105", "121110125", "121110127", "121110147"), 'FVOCI_VALUE',
        IIf(OBU_AC4603.AccountCode IN ("121130105", "121130125", "121130127", "121130147"), 'FVOCI_ADJUSTMENT',
          IIf(OBU_AC4603.AccountCode IN ("122010105", "122010125", "122010127", "122010147"), 'AC_VALUE',
            IIf(OBU_AC4603.AccountCode IN ("122030105", "122030125", "122030127", "122030147"), 'AC_ADJUSTMENT', 'Other')
          )
        )
      )
    )
  ) AS AccountGroup,
  Sum(OBU_AC4603.NetBalance) AS SubtotalNetBalance
FROM OBU_AC4603
WHERE 
  OBU_AC4603.AccountCode IN (
    "120050105","120050125","120050127",
    "120070105","120070125","120070127",
    "121110105","121110125","121110127","121110147",
    "121130105","121130125","121130127","121130147",
    "122010105","122010125","122010127","122010147",
    "122030105","122030125","122030127","122030147"
  )
  AND OBU_AC4603.CurrencyType = "USD"
  AND OBU_AC4603.DataMonthString = [DataMonthParam]
GROUP BY 
  IIf(OBU_AC4603.AccountCode IN ("120050105", "120050125", "120050127"), 'FVPL_VALUE',
    IIf(OBU_AC4603.AccountCode IN ("120070105", "120070125", "120070127"), 'FVPL_ADJUSTMENT',
      IIf(OBU_AC4603.AccountCode IN ("121110105", "121110125", "121110127", "121110147"), 'FVOCI_VALUE',
        IIf(OBU_AC4603.AccountCode IN ("121130105", "121130125", "121130127", "121130147"), 'FVOCI_ADJUSTMENT',
          IIf(OBU_AC4603.AccountCode IN ("122010105", "122010125", "122010127", "122010147"), 'AC_VALUE',
            IIf(OBU_AC4603.AccountCode IN ("122030105", "122030125", "122030127", "122030147"), 'AC_ADJUSTMENT', 'Other')
          )
        )
      )
    )
  );




New version add FVPL
Query FM10_OBU_AC4603_Subtotal

SELECT 
  AccountCodeMap.AssetMeasurementType & "_" & AccountCodeMap.Category AS MeasurementCategory,
  Sum(OBU_AC4603.NetBalance) AS SubtotalBalance
FROM
    AccountCodeMap
INNER JOIN
    OBU_AC4603
ON
    AccountCodeMap.AccountCode = OBU_AC4603.AccountCode
WHERE
    AccountCodeMap.Category IN ("Cost" , "ValuationAdjust", "ImpairmentLoss", 'OSU')
    AND OBU_AC4603.CurrencyType = "USD"
    AND OBU_AC4603.DataMonthString = "2024/11"
GROUP BY
    AccountCodeMap.AssetMeasurementType, AccountCodeMap.Category;

PARAMETERS DataMonthParam TEXT;
SELECT 
  AccountCodeMap.AssetMeasurementType & "_" & AccountCodeMap.Category AS MeasurementCategory,
  Sum(OBU_AC4603.NetBalance) AS SubtotalBalance
FROM
    AccountCodeMap
INNER JOIN
    OBU_AC4603
ON
    AccountCodeMap.AccountCode = OBU_AC4603.AccountCode
WHERE
    AccountCodeMap.Category IN ("Cost" , "ValuationAdjust", "ImpairmentLoss", 'OSU')
    AND OBU_AC4603.CurrencyType = "USD"
    AND OBU_AC4603.DataMonthString = [DataMonthParam]
GROUP BY
    AccountCodeMap.AssetMeasurementType, AccountCodeMap.Category;

=================================================

SELECT 
  AccountCodeMap.AssetMeasurementType & "_" & AccountCodeMap.Category AS MeasurementCategory,
  Sum(oa.NetBalance) AS SubtotalBalance
FROM
    AccountCodeMap
INNER JOIN
    (
        SELECT
            OBU_AC4603.AccountCode,
            OBU_AC4603.NetBalance
        FROM OBU_AC4603
        WHERE
            OBU_AC4603.CurrencyType = "USD"
            AND OBU_AC4603.DataMonthString = "2024/11"
    ) AS oa
ON
    AccountCodeMap.AccountCode = oa.AccountCode
WHERE
    AccountCodeMap.Category IN ("Cost" , "ValuationAdjust", "ImpairmentLoss", 'OSU')
GROUP BY
    AccountCodeMap.AssetMeasurementType, AccountCodeMap.Category;

PARAMETERS DataMonthParam TEXT;
SELECT 
  AccountCodeMap.AssetMeasurementType & "_" & AccountCodeMap.Category AS MeasurementCategory,
  Sum(oa.NetBalance) AS SubtotalBalance
FROM
    AccountCodeMap
INNER JOIN
    (
        SELECT
            OBU_AC4603.AccountCode,
            OBU_AC4603.NetBalance
        FROM OBU_AC4603
        WHERE
            OBU_AC4603.CurrencyType = 'USD'
            AND OBU_AC4603.DataMonthString = [DataMonthParam]
    ) AS oa
ON
    AccountCodeMap.AccountCode = oa.AccountCode
WHERE
    AccountCodeMap.Category IN ('Cost' , 'ValuationAdjust', 'ImpairmentLoss', 'OSU')
GROUP BY
    AccountCodeMap.AssetMeasurementType, AccountCodeMap.Category;
