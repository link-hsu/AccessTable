New Query AI602_GroupedAssetsImpairment

從 AssetsImpairment 做 GroupedMeasurement 的分類處理

PARAMETERS DataMonthParam TEXT;
SELECT 
    AssetsImpairment.BondCode,
    AssetsImpairment.Measurement,
    Switch(
        AssetsImpairment.Measurement IN ('強制FVPL金融資產-公債-中央政府(外'), 'FVPL_政府公債_外國_減損',
        AssetsImpairment.Measurement IN ('強制FVPL金融資產-普通公司債(公營)-海外', '強制FVPL金融資產-普通公司債(民營)-海外'), 'FVPL_公司債_外國_減損',
        AssetsImpairment.Measurement IN ('強制FVPL金融資產-金融債券-海外'), 'FVPL_金融債_外國_減損',
        AssetsImpairment.Measurement IN ('FVOCI債務工具-公債-中央政府(外國)'), 'FVOCI_政府公債_外國_減損',
        AssetsImpairment.Measurement IN ('FVOCI債務工具-普通公司債(公營)-海外', 'FVOCI債務工具-普通公司債(民營)-海外'), 'FVOCI_公司債_外國_減損',
        AssetsImpairment.Measurement IN ('FVOCI債務工具-金融債券-海外'), 'FVOCI_金融債_外國_減損',
        AssetsImpairment.Measurement IN ('AC債務工具投資-公債-中央政府(外國)'), 'AC_政府公債_外國_減損',
        AssetsImpairment.Measurement IN ('AC債務工具投資-普通公司債(公營)-海外', 'AC債務工具投資-普通公司債(民營)-海外'), 'AC_公司債_外國_減損',
        AssetsImpairment.Measurement IN ('AC債務工具投資-金融債券-海外'), 'AC_金融債_外國_減損',
        True, 'Others'
    ) AS GroupedMeasurement,
    AssetsImpairment.CurrImpairmentCost
FROM AssetsImpairment
WHERE AssetsImpairment.DataMonthString = [DataMonthParam];


New Query AI602_SumIPNotUSD

取得非 USD 債券對應的減損總額（每種 GroupedMeasurement）

PARAMETERS DataMonthParam TEXT;
SELECT 
    GAIP.GroupedMeasurement,
    SUM(GAIP.CurrImpairmentCost) AS Total_CurrImpairmentCost
FROM AI602_GroupedAssetsImpairment AS GAIP
INNER JOIN FXDebtEvaluation AS fx
    ON GAIP.BondCode = MID(fx.Security_id, LEN(fx.Security_id) - 11, 12)
WHERE fx.Ccy <> 'USD'
  AND fx.DataMonthString = [DataMonthParam]
GROUP BY GAIP.GroupedMeasurement;


New Query AI602_SumIpUSD

合併計算：總減損、非 USD 減損、USD 減損

PARAMETERS DataMonthParam TEXT;
SELECT 
    GAIP.GroupedMeasurement,
    SUM(GAIP.CurrImpairmentCost) AS Total_CurrImpairmentCost,
    IIf(NoUSD.Total_CurrImpairmentCost IS NULL, 0, NoUSD.Total_CurrImpairmentCost) AS Impairment_Not_USD,
    SUM(GAIP.CurrImpairmentCost) - IIf(NoUSD.Total_CurrImpairmentCost IS NULL, 0, NoUSD.Total_CurrImpairmentCost) AS Impairment_OnlyWith_USD
FROM AI602_GroupedAssetsImpairment AS GAIP
LEFT JOIN AI602_SumIPNotUSD AS NoUSD
    ON GAIP.GroupedMeasurement = NoUSD.GroupedMeasurement
GROUP BY GAIP.GroupedMeasurement, NoUSD.Total_CurrImpairmentCost;



建立新的對照Table AI602_ACCGroup，再使用以下Query Join進去

New Query AI602_GroupedAC5601 

PARAMETERS DataMonthParam TEXT;
SELECT
  o.DataID,
  o.DataMonthString,
  o.AccountCode,
  a.AccountTitle,
  IIF(g.GroupedAccount IS NULL, 'Others', g.GroupedAccount) AS GroupedAccount,
  o.CurrencyType,
  o.NetBalance
FROM
  (OBU_AC5601 AS o
    INNER JOIN ACCAccount AS a
      ON o.AccountCode = a.AccountCode)
  LEFT JOIN AI602_ACCGroup AS g
    ON a.AccountTitle = g.AccountTitle
WHERE
  o.CurrencyType = "USD"
  AND o.DataMonthString = [DataMonthParam];




New Query AI602_Subtotal

SELECT AI602_SumIpUSD.GroupedMeasurement AS MeasureType, 
       AI602_SumIpUSD.Impairment_OnlyWith_USD AS SubNetBalance 
FROM AI602_SumIpUSD
UNION ALL 
SELECT AI602_GroupedAC5601.GroupedAccount AS MeasureType, 
       SUM(AI602_GroupedAC5601.NetBalance) AS SubNetBalance 
FROM AI602_GroupedAC5601 
GROUP BY AI602_GroupedAC5601.GroupedAccount;
